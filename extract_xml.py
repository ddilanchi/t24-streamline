"""
extract_xml.py
Parses .NET BinaryFormatter (MS-NRBF) binary to extract ComplianceXML
from EnergyPro .bld files — no EnergyPro DLLs required.

Usage:
  python extract_xml.py              <- all .bld files in BLD EXAMPLES\
  python extract_xml.py path\to.bld  <- specific file
"""

import sys
import os
import io
import struct
import gzip
import zlib

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BLD_DIR    = os.path.join(SCRIPT_DIR, "BLD EXAMPLES")
OUT_DIR    = os.path.join(SCRIPT_DIR, "extracted_xml")

TARGET_FIELD = "<ComplianceXML>k__BackingField"

# ---------------------------------------------------------------------------
# Minimal MS-NRBF parser
# Spec: https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-nrbf
# ---------------------------------------------------------------------------

class NrbfReader:
    def __init__(self, data: bytes):
        self.s = io.BytesIO(data)
        self.classes = {}   # object_id -> (class_name, [field_names], [bin_types], [extras])
        self.objects = {}   # object_id -> value
        self.libraries = {} # lib_id -> name
        self.found_xml = None  # set when we find it

    # --- low-level reads ---
    def byte(self):
        b = self.s.read(1)
        if not b: raise EOFError
        return b[0]

    def int32(self):
        return struct.unpack('<i', self.s.read(4))[0]

    def varint(self):
        r, shift = 0, 0
        for _ in range(5):
            b = self.byte()
            r |= (b & 0x7F) << shift
            if not (b & 0x80): return r
            shift += 7
        return r

    def lps(self):
        """Length-prefixed string"""
        n = self.varint()
        return self.s.read(n).decode('utf-8', errors='replace')

    def primitive(self, tc):
        sizes = {1:1, 2:1, 3:2, 6:8, 7:2, 8:4, 9:8, 10:1, 11:4, 12:8, 13:8, 14:2, 15:4, 16:8}
        if tc == 5:   return self.lps()   # Decimal stored as string
        if tc == 18:  return self.lps()   # String
        if tc in sizes:
            raw = self.s.read(sizes[tc])
            if tc == 1:  return bool(raw[0])
            if tc == 2:  return raw[0]
            if tc == 3:  return raw[0] | (raw[1] << 8)
            if tc == 6:  return struct.unpack('<d', raw)[0]
            if tc == 7:  return struct.unpack('<h', raw)[0]
            if tc == 8:  return struct.unpack('<i', raw)[0]
            if tc == 9:  return struct.unpack('<q', raw)[0]
            if tc == 10: return struct.unpack('<b', raw)[0]
            if tc == 11: return struct.unpack('<f', raw)[0]
            if tc == 12: return struct.unpack('<q', raw)[0]  # TimeSpan
            if tc == 13: return struct.unpack('<q', raw)[0]  # DateTime
            if tc == 14: return struct.unpack('<H', raw)[0]
            if tc == 15: return struct.unpack('<I', raw)[0]
            if tc == 16: return struct.unpack('<Q', raw)[0]
        return None

    # --- member type extra info ---
    def read_extra(self, bt):
        if bt == 0:   return ('Prim', self.byte())
        if bt == 1:   return ('String', None)
        if bt == 2:   return ('Object', None)
        if bt == 3:   return ('SysClass', self.lps())
        if bt == 4:   return ('Class', self.lps(), self.int32())
        if bt == 5:   return ('ObjArr', None)
        if bt == 6:   return ('StrArr', None)
        if bt == 7:   return ('PrimArr', self.byte())
        return ('Unk', bt)

    # --- record reading ---
    def record(self):
        rt = self.byte()
        return self.record_of_type(rt)

    def record_of_type(self, rt):

        # 0: Header
        if rt == 0:
            root = self.int32(); self.s.read(12)
            return ('Header', root)

        # 12: BinaryLibrary — no data, read next
        if rt == 12:
            lid = self.int32(); name = self.lps()
            self.libraries[lid] = name
            return self.record()   # transparent

        # 11: MessageEnd
        if rt == 11:
            return ('END',)

        # 10: ObjectNull
        if rt == 10:
            return None

        # 13: ObjectNullMultiple256
        if rt == 13:
            return [None] * self.byte()

        # 14: ObjectNullMultiple
        if rt == 14:
            return [None] * self.int32()

        # 6: BinaryObjectString
        if rt == 6:
            oid = self.int32(); val = self.lps()
            self.objects[oid] = val
            return val

        # 8: MemberPrimitiveTyped
        if rt == 8:
            tc = self.byte(); return self.primitive(tc)

        # 9: MemberReference
        if rt == 9:
            ref_id = self.int32()
            return ('Ref', ref_id)

        # 5: ClassWithMembersAndTypes
        if rt == 5:
            return self.class_record(with_types=True, sys=False)

        # 4: SystemClassWithMembersAndTypes
        if rt == 4:
            return self.class_record(with_types=True, sys=True)

        # 3: ClassWithMembers
        if rt == 3:
            return self.class_record(with_types=False, sys=False)

        # 2: SystemClassWithMembers
        if rt == 2:
            return self.class_record(with_types=False, sys=True)

        # 1: ClassWithId (reuses prior class definition)
        if rt == 1:
            oid = self.int32(); meta_id = self.int32()
            cdef = self.classes.get(meta_id)
            if cdef is None:
                raise ValueError(f"ClassWithId references unknown meta_id {meta_id}")
            cname, fnames, btypes, extras = cdef
            return self.read_class_values(oid, cname, fnames, btypes, extras)

        # 7: BinaryArray
        if rt == 7:
            return self.binary_array()

        # 15/16/17: Single arrays
        if rt in (15, 16, 17):
            return self.single_array(rt)

        raise ValueError(f"Unknown record type {rt} at offset {self.s.tell()}")

    def class_record(self, with_types, sys):
        oid = self.int32()
        cname = self.lps()
        n = self.int32()
        fnames = [self.lps() for _ in range(n)]

        if with_types:
            btypes = [self.byte() for _ in range(n)]
            extras = [self.read_extra(bt) for bt in btypes]
        else:
            btypes = [2] * n   # Object type
            extras = [('Object', None)] * n

        if not sys:
            self.int32()  # library id

        self.classes[oid] = (cname, fnames, btypes, extras)
        return self.read_class_values(oid, cname, fnames, btypes, extras)

    def read_class_values(self, oid, cname, fnames, btypes, extras):
        vals = {}
        skip = 0
        for fname, bt, ex in zip(fnames, btypes, extras):
            if skip > 0:
                vals[fname] = None
                skip -= 1
                continue

            if bt == 0:  # inline primitive
                vals[fname] = self.primitive(ex[1])
            else:
                v = self.record()
                # ObjectNullMultiple returns a list of Nones
                if isinstance(v, list):
                    vals[fname] = None
                    skip = len(v) - 1
                else:
                    vals[fname] = v

            # Check right here if we found the XML
            if fname == TARGET_FIELD and isinstance(vals[fname], str) and len(vals[fname]) > 100:
                self.found_xml = vals[fname]

        obj = {'__class__': cname, **vals}
        self.objects[oid] = obj
        return obj

    def binary_array(self):
        oid = self.int32()
        atype = self.byte()   # 0=single, 1=jagged, 2=rectangular, 3=singleOffset, 4=jaggedOffset, 5=rectOffset
        rank = self.int32()
        lengths = [self.int32() for _ in range(rank)]
        if atype in (3, 4, 5):
            [self.int32() for _ in range(rank)]  # lower bounds
        bt = self.byte()
        ex = self.read_extra(bt)

        total = 1
        for l in lengths: total *= l

        if bt == 0:  # primitive array — read inline, return byte count
            tc = ex[1]
            sizes = {1:1,2:1,3:2,6:8,7:2,8:4,9:8,10:1,11:4,12:8,13:8,14:2,15:4,16:8}
            sz = sizes.get(tc, 4)
            data = self.s.read(total * sz)
            self.objects[oid] = data
            return data

        # object/class array
        elems = []
        skip = 0
        for _ in range(total):
            if skip > 0:
                elems.append(None); skip -= 1; continue
            v = self.record()
            if isinstance(v, list):
                elems.append(None); skip = len(v) - 1
            else:
                elems.append(v)
        self.objects[oid] = elems
        return elems

    def single_array(self, rt):
        oid = self.int32(); length = self.int32()
        if rt == 15:   # ArraySinglePrimitive
            tc = self.byte()
            sizes = {1:1,2:1,3:2,6:8,7:2,8:4,9:8,10:1,11:4,12:8,13:8,14:2,15:4,16:8}
            sz = sizes.get(tc, 4)
            data = self.s.read(length * sz)
            self.objects[oid] = data
            return data
        elems = []
        skip = 0
        for _ in range(length):
            if skip > 0:
                elems.append(None); skip -= 1; continue
            v = self.record()
            if isinstance(v, list):
                elems.append(None); skip = len(v) - 1
            else:
                elems.append(v)
        self.objects[oid] = elems
        return elems

    def resolve(self, val):
        """Follow Ref chains"""
        seen = set()
        while isinstance(val, tuple) and val[0] == 'Ref':
            ref_id = val[1]
            if ref_id in seen: break
            seen.add(ref_id)
            val = self.objects.get(ref_id)
        return val

    def run(self):
        try:
            while True:
                r = self.record()
                if r == ('END',): break
                if self.found_xml: break
        except EOFError:
            pass
        except Exception as e:
            if self.found_xml:
                return  # already got what we need
            # Try to resolve any Ref-typed value for our target
            for oid, obj in self.objects.items():
                if isinstance(obj, dict) and TARGET_FIELD in obj:
                    v = self.resolve(obj[TARGET_FIELD])
                    if isinstance(v, str) and len(v) > 100:
                        self.found_xml = v
                        return
            raise


# ---------------------------------------------------------------------------
# Compressed content scan (fallback)
# ---------------------------------------------------------------------------

def scan_compressed(data: bytes) -> str | None:
    # Try GZip
    for i in range(len(data) - 2):
        if data[i] == 0x1f and data[i+1] == 0x8b:
            try:
                dec = gzip.decompress(data[i:])
                if b"<?xml" in dec or b"<Title24" in dec:
                    return dec.decode('utf-8', errors='replace')
            except Exception:
                pass
    # Try zlib
    for magic in (b"\x78\x9c", b"\x78\x01", b"\x78\xda"):
        idx = 0
        while True:
            idx = data.find(magic, idx)
            if idx == -1: break
            try:
                dec = zlib.decompress(data[idx:])
                if b"<?xml" in dec or b"<Title24" in dec:
                    return dec.decode('utf-8', errors='replace')
            except Exception:
                pass
            idx += 1
    return None


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def process(bld_path: str):
    print(f"\n{'='*60}")
    print(f"File: {os.path.basename(bld_path)}")

    with open(bld_path, 'rb') as f:
        data = f.read()

    base = os.path.splitext(os.path.basename(bld_path))[0]

    # 1. NRBF parse
    xml = None
    try:
        reader = NrbfReader(data)
        reader.run()
        xml = reader.found_xml
        if xml is None:
            # Check all parsed objects for the target field
            for obj in reader.objects.values():
                if isinstance(obj, dict) and TARGET_FIELD in obj:
                    v = reader.resolve(obj[TARGET_FIELD])
                    if isinstance(v, str) and len(v) > 100:
                        xml = v
                        break
    except Exception as e:
        print(f"  NRBF parse error: {e}")

    if xml:
        out = os.path.join(OUT_DIR, f"{base}.xml")
        with open(out, 'w', encoding='utf-8') as f:
            f.write(xml)
        print(f"  [OK] ComplianceXML extracted ({len(xml):,} chars) -> {os.path.basename(out)}")
        return

    # 2. Compressed fallback
    print("  ComplianceXML not found via NRBF — trying compressed scan...")
    xml = scan_compressed(data)
    if xml:
        out = os.path.join(OUT_DIR, f"{base}_decompressed.xml")
        with open(out, 'w', encoding='utf-8') as f:
            f.write(xml)
        print(f"  [OK] Found in compressed block ({len(xml):,} chars) -> {os.path.basename(out)}")
        return

    print("  [NONE] No XML content found in this file.")
    print("  ComplianceXML may be null (project not yet run through compliance).")


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    paths = sys.argv[1:] if len(sys.argv) > 1 else [
        os.path.join(BLD_DIR, f)
        for f in os.listdir(BLD_DIR) if f.lower().endswith('.bld')
    ]
    for p in paths:
        process(p)
    print(f"\nDone. Output: {OUT_DIR}")


if __name__ == "__main__":
    main()
