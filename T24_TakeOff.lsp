;;; ============================================================
;;; T24 TakeOff Wizard  -  T24_TakeOff.lsp
;;; Workflow: TZ-ZONE -> TZ-WIN / TZ-DOOR / TZ-SKY
;;;
;;; Commands:
;;;   TZ-ZONE   - Click room name text; auto-traces boundary from that point,
;;;               ceiling height, floor number, condition type
;;;   TZ-WIN    - Click at a window location; prompts area, U-Factor, SHGC
;;;   TZ-DOOR   - Click at a door location; prompts area
;;;   TZ-SKY    - Click inside a zone for a skylight; prompts area, U, SHGC
;;;   TZ-RESET  - Clears all T24 XDATA and labels (use with caution)
;;; ============================================================

(vl-load-com)

;; ── Configuration ────────────────────────────────────────────────────────────
(setq *TZ-APP*         "T24TOW")
(setq *TZ-LYR-ZONE*    "T24-ZONE")
(setq *TZ-LYR-LABEL*   "T24-LABEL")
(setq *TZ-LYR-WIN*     "T24-WIN")
(setq *TZ-LYR-DOOR*    "T24-DOOR")
(setq *TZ-LYR-SKY*     "T24-SKY")
(setq *TZ-UNIT-FT*     12.0)   ; 1 foot = 12 drawing units (inches)
(setq *TZ-MARKER-RAD*  6.0)    ; Opening circle marker radius (6 inches)
(setq *TZ-LYR-WALL*    "T24-WALL")
(setq *TZ-WALL-RAD*    4.0)    ; Wall marker circle radius (4 inches)

;; ── Layer setup ───────────────────────────────────────────────────────────────
(defun tz-make-layer (name color ltype / )
  (if (null (tblsearch "LAYER" name))
    (entmake
      (list '(0 . "LAYER")
            '(100 . "AcDbSymbolTableRecord")
            '(100 . "AcDbLayerTableRecord")
            (cons 2  name)
            '(70 . 0)
            (cons 62 color)
            (cons 6  ltype)))))

(defun tz-setup ( / )
  (tz-make-layer *TZ-LYR-ZONE*  4  "Continuous")  ; cyan
  (tz-make-layer *TZ-LYR-LABEL* 7  "Continuous")  ; white
  (tz-make-layer *TZ-LYR-WIN*   2  "Continuous")  ; yellow
  (tz-make-layer *TZ-LYR-DOOR*  1  "Continuous")  ; red
  (tz-make-layer *TZ-LYR-SKY*   5  "Continuous")  ; blue
  (tz-make-layer *TZ-LYR-WALL*  3  "Continuous")) ; green

;; ── XDATA registration ────────────────────────────────────────────────────────
(defun tz-regapp ( / )
  (if (null (tblsearch "APPID" *TZ-APP*))
    (regapp *TZ-APP*)))

;; ── Geometry helpers ──────────────────────────────────────────────────────────

;; Returns list of (x y) pairs from an LWPOLYLINE entity
(defun tz-get-pts (ent / obj coords pts i)
  (setq obj    (vlax-ename->vla-object ent)
        coords (vlax-safearray->list
                 (vlax-variant-value (vla-get-Coordinates obj)))
        pts    '()
        i      0)
  (while (< i (length coords))
    (setq pts (append pts (list (list (nth i coords) (nth (1+ i) coords)))))
    (setq i (+ i 2)))
  pts)

;; Average-of-vertices centroid (good enough for room shapes)
(defun tz-centroid (pts / n sx sy)
  (setq n (length pts) sx 0.0 sy 0.0)
  (foreach p pts (setq sx (+ sx (car p)) sy (+ sy (cadr p))))
  (list (/ sx n) (/ sy n) 0.0))

;; Ray-casting point-in-polygon (reused from RoomTools)
(defun tz-pip (pt poly / x y n i j inside xi yi xj yj)
  (setq x (car pt) y (cadr pt) n (length poly) inside nil j (1- n) i 0)
  (while (< i n)
    (setq xi (car (nth i poly)) yi (cadr (nth i poly))
          xj (car (nth j poly)) yj (cadr (nth j poly)))
    (if (and (or (and (< yi y) (>= yj y))
                 (and (< yj y) (>= yi y)))
             (< x (+ xi (/ (* (- xj xi) (- y yi)) (- yj yi)))))
      (setq inside (not inside)))
    (setq j i i (1+ i)))
  inside)

;; ── Entity makers ─────────────────────────────────────────────────────────────

(defun tz-make-text (pt height str / )
  (entmake
    (list '(0 . "TEXT") '(100 . "AcDbEntity") '(100 . "AcDbText")
          (cons 10 (list (car pt) (cadr pt) 0.0))
          (cons 40 height) (cons 1 str)
          '(41 . 1.0) '(50 . 0.0) '(7 . "Standard")
          '(72 . 1) (cons 11 (list (car pt) (cadr pt) 0.0))
          '(100 . "AcDbText") '(73 . 2))))

(defun tz-make-circle (ctr r / )
  (entmake
    (list '(0 . "CIRCLE") '(100 . "AcDbEntity") '(100 . "AcDbCircle")
          (cons 10 (list (car ctr) (cadr ctr) 0.0))
          (cons 40 r))))

(defun tz-make-line (p1 p2 / )
  (entmake
    (list '(0 . "LINE") '(100 . "AcDbEntity") '(100 . "AcDbLine")
          (cons 10 (list (car p1) (cadr p1) 0.0))
          (cons 11 (list (car p2) (cadr p2) 0.0)))))

;; ── XDATA writers/readers ────────────────────────────────────────────────────

(defun tz-set-xdata (ent xlist / edata)
  (tz-regapp)
  (setq edata (entget ent))
  ;; Remove existing T24TOW xdata first
  (setq edata (vl-remove-if
    '(lambda (x) (and (= (car x) -3)
                      (assoc *TZ-APP* (cdr x))))
    edata))
  (entmod (append edata (list (cons -3 (list xlist))))))

;; Get raw xdata list for T24TOW appid from an entity
(defun tz-get-xdata (ent / raw grp)
  (setq raw (entget ent (list *TZ-APP*)))
  (setq grp (assoc -3 raw))
  (if grp (cdr (assoc *TZ-APP* (cdr grp))) nil))

;; Get the nth occurrence of a group-code value in an xdata list
(defun tz-xd-nth (xd grp idx / count result)
  (setq count -1  result nil)
  (foreach item xd
    (if (and (= (car item) grp) (null result))
      (progn
        (setq count (1+ count))
        (if (= count idx) (setq result (cdr item))))))
  ;; Safety guard: if result is boolean T (XDATA round-trip bug), treat as nil
  (if (= result T) nil result))

;; Safe numeric xdata reader — guarantees a number, never T or nil
(defun tz-xd-num (xd grp idx default / val)
  (setq val (tz-xd-nth xd grp idx))
  (if (numberp val) val default))

;; ── Auto ID generators ───────────────────────────────────────────────────────

(defun tz-pad (s n / )
  (while (< (strlen s) n) (setq s (strcat "0" s)))
  s)

(defun tz-next-zone-id ( / ss)
  (setq ss (ssget "X" (list (cons 8 *TZ-LYR-ZONE*) '(0 . "LWPOLYLINE"))))
  (strcat "Z-" (tz-pad (itoa (1+ (if ss (sslength ss) 0))) 3)))

(defun tz-next-open-id ( / ss1 ss2 ss3)
  (setq ss1 (ssget "X" (list (cons 8 *TZ-LYR-WIN*)  '(0 . "CIRCLE")))
        ss2 (ssget "X" (list (cons 8 *TZ-LYR-DOOR*) '(0 . "CIRCLE")))
        ss3 (ssget "X" (list (cons 8 *TZ-LYR-SKY*)  '(0 . "CIRCLE"))))
  (strcat "O-" (tz-pad (itoa (1+ (+ (if ss1 (sslength ss1) 0)
                                     (if ss2 (sslength ss2) 0)
                                     (if ss3 (sslength ss3) 0)))) 3)))

;; ── Visual labels ─────────────────────────────────────────────────────────────

(defun tz-zone-label (centroid zone-name area-ft ceil-ht floor / lsave th)
  (setq lsave (getvar "CLAYER"))
  (setvar "CLAYER" *TZ-LYR-LABEL*)
  (setq th (* 0.5 *TZ-UNIT-FT*))   ; 6 inches
  ;; Zone name – large
  (tz-make-text centroid (* 1.4 th) zone-name)
  ;; Area | ceiling | floor
  (tz-make-text
    (list (car centroid) (- (cadr centroid) (* 2.4 th)) 0.0)
    (* 0.85 th)
    (strcat (rtos area-ft 2 1) " sqft  |  "
            (rtos ceil-ht 2 1) "' clg  |  Fl " (itoa floor)))
  (setvar "CLAYER" lsave))

(defun tz-open-label (pt open-id open-type area-ft uf shgc / lsave th off lbl)
  (setq lsave (getvar "CLAYER"))
  (setvar "CLAYER" *TZ-LYR-LABEL*)
  (setq th  (* 0.08 *TZ-UNIT-FT*)
        off (+ *TZ-MARKER-RAD* (* 1.5 th)))
  (setq lbl (strcat open-id "  " open-type "  " (rtos area-ft 2 1) " sqft"))
  (if (> uf 0)
    (setq lbl (strcat lbl "  U=" (rtos uf 2 2))))
  (if (> shgc 0)
    (setq lbl (strcat lbl "  SHGC=" (rtos shgc 2 2))))
  (tz-make-text (list (car pt) (+ (cadr pt) off) 0.0) th lbl)
  (setvar "CLAYER" lsave))

;; ── Popup choice dialog (appears near cursor) ───────────────────────────────
(defun tz-boundary-popup ( / dcl-path dcl-id result f)
  "Shows a dialog with boundary failure options. Returns keyword string."
  (setq dcl-path (vl-filename-mktemp "tz_bnd" nil ".dcl"))
  (setq f (open dcl-path "w"))
  (write-line "tz_bnd : dialog {" f)
  (write-line "  label = \"BOUNDARY Failed\";" f)
  (write-line "  : column {" f)
  (write-line "    : button { key = \"retry\";  label = \"&Retry at new point\"; width = 28; fixed_width = true; }" f)
  (write-line "    : button { key = \"patch\";  label = \"&Patch gaps (draw lines)\"; width = 28; fixed_width = true; }" f)
  (write-line "    : button { key = \"manual\"; label = \"&Manual corners\"; width = 28; fixed_width = true; }" f)
  (write-line "    : button { key = \"select\"; label = \"&Select existing polyline\"; width = 28; fixed_width = true; }" f)
  (write-line "    : spacer { height = 0.3; }" f)
  (write-line "    : button { key = \"cancel\"; label = \"Cancel\"; is_cancel = true; width = 28; fixed_width = true; }" f)
  (write-line "  }" f)
  (write-line "}" f)
  (close f)

  (setq result "Retry")
  (setq dcl-id (load_dialog dcl-path))
  (if (>= dcl-id 0)
    (progn
      (new_dialog "tz_bnd" dcl-id)
      (action_tile "retry"  "(setq result \"Retry\")  (done_dialog 1)")
      (action_tile "patch"  "(setq result \"Patch\")  (done_dialog 1)")
      (action_tile "manual" "(setq result \"Manual\") (done_dialog 1)")
      (action_tile "select" "(setq result \"Select\") (done_dialog 1)")
      (action_tile "cancel" "(setq result nil) (done_dialog 0)")
      (start_dialog)
      (unload_dialog dcl-id)))
  (vl-file-delete dcl-path)
  result)

;; ── Manual corner-pick polyline builder ──────────────────────────────────────
(defun tz-pick-corners ( / pts pt first-pt ent elist p)
  (princ "\n[T24] Click room corners in order. Press Enter when done (min 3 points).")
  (setq pts '())
  (while
    (progn
      (setq pt (if pts
                 (getpoint (last pts) "\n[T24]   Next corner (Enter to finish): ")
                 (getpoint "\n[T24]   First corner: ")))
      (not (null pt)))
    (setq pts (append pts (list (list (car pt) (cadr pt))))))
  (if (< (length pts) 3)
    (progn (alert "[T24] Need at least 3 points. Cancelled.") nil)
    (progn
      ;; Build closed LWPOLYLINE
      (setq elist
        (list '(0 . "LWPOLYLINE")
              '(100 . "AcDbEntity")
              (cons 8 *TZ-LYR-ZONE*)
              '(100 . "AcDbPolyline")
              (cons 90 (length pts))
              '(70 . 1)))  ; closed
      (foreach p pts
        (setq elist (append elist (list (cons 10 (list (car p) (cadr p)))))))
      (entmake elist)
      (entlast))))

;; ── Door swing layer freeze/thaw helpers ─────────────────────────────────────
;; Common door swing layer names across different CAD standards
(setq *TZ-DOOR-LAYERS*
  '("A-DOOR" "A-DOOR-SWING" "A-FLOR-DOOR" "DOOR" "DOORS" "DOOR-SWING"
    "S-DOOR" "S-DOOR-SWING" "DOOR SWING" "A-DOOR SWING"))

;; Use _.LAYER command for freeze/thaw — works on xref-dependent layers too
(defun tz-freeze-layer (lname / ldata)
  (setq ldata (tblsearch "LAYER" lname))
  (if (and ldata (= 0 (logand (cdr (assoc 70 ldata)) 1)))  ; exists and not frozen
    (progn
      (command "_.LAYER" "_Freeze" lname "")
      T)
    nil))

(defun tz-freeze-door-layers ( / frozen)
  (setq frozen '())
  (foreach lname *TZ-DOOR-LAYERS*
    (if (tz-freeze-layer lname)
      (setq frozen (cons lname frozen))))
  frozen)

(defun tz-thaw-layers (layer-names / )
  (foreach lname layer-names
    (command "_.LAYER" "_Thaw" lname "")))


;; ── Boundary creation ────────────────────────────────────────────────────────
;; Uses _-BOUNDARY with progressive gap tolerance.
;; Tries 0 → 4 → 12 → 36 → 48 until a polyline is created.
;; Also handles REGION output (explode + PEDIT join → polyline).

;; Core: run _-BOUNDARY at pt with given gap-tol, return largest polyline or nil
(defun tz-try-boundary (pt gap-tol / old-gaptol last-ent ent etype best-area cur-area
                                     region-last ss-exp scan-ent)
  (setq old-gaptol (getvar "HPGAPTOL"))
  (setvar "HPGAPTOL" gap-tol)
  (setq last-ent (entlast)  ent nil)
  (command "_-BOUNDARY" pt "")
  (if (not (equal (entlast) last-ent))
    (progn
      ;; Convert any REGIONs to polylines first
      (setq scan-ent (entnext last-ent))
      (while scan-ent
        (setq etype (cdr (assoc 0 (entget scan-ent))))
        (if (= etype "REGION")
          (progn
            (princ "\n[T24] Converting region to polyline...")
            (setq region-last (entlast))
            (command "_.EXPLODE" scan-ent "")
            (setq ss-exp (ssadd)  scan-ent (entnext region-last))
            (while scan-ent
              (ssadd scan-ent ss-exp)
              (setq scan-ent (entnext scan-ent)))
            (if (> (sslength ss-exp) 0)
              (command "_.PEDIT" (ssname ss-exp 0) "_Yes" "_Join" ss-exp "" ""))))
        (setq scan-ent (entnext scan-ent)))
      ;; Find largest polyline
      (setq scan-ent (entnext last-ent)  best-area 0.0)
      (while scan-ent
        (if (= (cdr (assoc 0 (entget scan-ent))) "LWPOLYLINE")
          (progn
            (setq cur-area (vlax-curve-getarea (vlax-ename->vla-object scan-ent)))
            (if (> cur-area best-area)
              (setq best-area cur-area  ent scan-ent))))
        (setq scan-ent (entnext scan-ent)))
      ;; Delete everything else created
      (setq scan-ent (entnext last-ent))
      (while scan-ent
        (if (not (equal scan-ent ent))
          (entdel scan-ent))
        (setq scan-ent (entnext scan-ent)))))
  (setvar "HPGAPTOL" old-gaptol)
  ent)

;; Quick attempt: HPGAPTOL=0 only. Fast to succeed or fail.
;; If this fails the popup appears immediately — Retry uses tz-hatch-boundary-full.
(defun tz-hatch-boundary (pt / ent)
  (command "_.VIEW" "_Save" "TZ-BOUNDARY-VIEW")
  (command "_.ZOOM" "_Extents")
  (setq ent (tz-try-boundary pt 0.0))
  (command "_.VIEW" "_Restore" "TZ-BOUNDARY-VIEW")
  (command "_.VIEW" "_Delete" "TZ-BOUNDARY-VIEW" "")
  ent)

;; Full retry: tries progressive gap tolerances 4 → 12 → 36 → 48
(defun tz-hatch-boundary-full (pt / ent)
  (command "_.VIEW" "_Save" "TZ-BOUNDARY-VIEW")
  (command "_.ZOOM" "_Extents")
  (setq ent nil)
  (foreach gap-tol '(4.0 12.0 36.0 48.0)
    (if (null ent)
      (progn
        (setq ent (tz-try-boundary pt gap-tol))
        (if ent
          (princ (strcat "\n[T24] Boundary ok (gap tol " (rtos gap-tol 2 0) "\")"))))))
  (command "_.VIEW" "_Restore" "TZ-BOUNDARY-VIEW")
  (command "_.VIEW" "_Delete" "TZ-BOUNDARY-VIEW" "")
  ent)

;; ── Polyline cleaner: flatten arcs, remove stubs, collapse door triangles ─────

(defun tz-pline-verts (obj / coords nv i verts)
  (setq coords (vlax-safearray->list
                 (vlax-variant-value (vla-get-Coordinates obj)))
        nv     (/ (length coords) 2)
        verts  '()
        i      0)
  (repeat nv
    (setq verts (append verts
      (list (list (nth (* 2 i) coords)
                  (nth (+ 1 (* 2 i)) coords)
                  (vla-GetBulge obj i)))))
    (setq i (1+ i)))
  verts)

(defun tz-set-pline-verts (obj verts / n coords sa i)
  (setq n (length verts)  coords '())
  (foreach v verts
    (setq coords (append coords (list (car v) (cadr v)))))
  (setq sa (vlax-make-safearray vlax-vbDouble (cons 0 (1- (length coords)))))
  (vlax-safearray-fill sa coords)
  (vla-put-Coordinates obj sa)
  (setq i 0)
  (repeat n (vla-SetBulge obj i 0.0) (setq i (1+ i))))

(defun tz-cp-seg-len (v1 v2)
  (sqrt (+ (expt (- (car v2) (car v1)) 2)
           (expt (- (cadr v2) (cadr v1)) 2))))

(defun tz-clean-pline (ent / obj verts n i j bulge slen
                           arc-list door-list used-arcs
                           ai av dj dmid dist best-arc best-dist
                           lo hi to-remove new-verts)
  (setq obj   (vlax-ename->vla-object ent)
        verts (tz-pline-verts obj)
        n     (length verts))

  ;; Step 1: Collect arc vertex positions and door-width segments (before flattening)
  ;;   arc-list  — each entry: (index x y)
  ;;   door-list — each entry: (start-idx end-idx midx midy)
  ;;   Door widths: 30-44" covers 2'6" through 3'8"
  (setq arc-list '()  i 0)
  (repeat n
    (setq bulge (caddr (nth i verts)))
    (if (and bulge (/= bulge 0.0))
      (setq arc-list (cons (list i (car (nth i verts)) (cadr (nth i verts)))
                           arc-list)))
    (setq i (1+ i)))

  (setq door-list '()  i 0)
  (repeat n
    (setq j    (rem (1+ i) n)
          slen (tz-cp-seg-len
                 (list (car (nth i verts)) (cadr (nth i verts)))
                 (list (car (nth j verts)) (cadr (nth j verts)))))
    (if (and (>= slen 30.0) (<= slen 44.0))
      (setq door-list
        (cons (list i j
                (* 0.5 (+ (car (nth i verts)) (car (nth j verts))))
                (* 0.5 (+ (cadr (nth i verts)) (cadr (nth j verts)))))
              door-list)))
    (setq i (1+ i)))

  ;; Flatten all arc segments to straight lines
  (setq verts (mapcar '(lambda (v) (list (car v) (cadr v) 0.0)) verts))

  ;; Step 2: For each door segment find the nearest arc (geometric distance, max 72").
  ;;         Remove the shorter vertex path between them — this collapses the door notch.
  (setq to-remove '()  used-arcs '())
  (foreach door door-list
    (setq dj        (car door)
          dmid      (list (caddr door) (cadddr door))
          best-arc  nil
          best-dist 72.0)          ;; 72" (6') max pairing radius
    ;; Find nearest unused arc
    (foreach arc arc-list
      (if (not (member (car arc) used-arcs))
        (progn
          (setq dist (tz-cp-seg-len dmid (list (cadr arc) (caddr arc))))
          (if (< dist best-dist)
            (setq best-arc arc  best-dist dist)))))
    (if best-arc
      (progn
        (setq ai (car best-arc))
        (setq used-arcs (cons ai used-arcs))
        ;; Remove the shorter polygon path from ai to dj (both inclusive)
        (setq lo (min ai dj)  hi (max ai dj))
        (if (<= (- hi lo) (- (+ n lo) hi))
          ;; Direct span lo..hi is shorter
          (progn
            (setq i lo)
            (repeat (1+ (- hi lo))
              (setq to-remove (cons i to-remove))
              (setq i (1+ i))))
          ;; Wrap-around span is shorter: hi..n-1 then 0..lo
          (progn
            (setq i hi)
            (repeat (- n hi)
              (setq to-remove (cons i to-remove))
              (setq i (1+ i)))
            (setq i 0)
            (repeat (1+ lo)
              (setq to-remove (cons i to-remove))
              (setq i (1+ i))))))))

  (if to-remove
    (progn
      (setq new-verts '()  i 0)
      (repeat n
        (if (not (member i to-remove))
          (setq new-verts (append new-verts (list (nth i verts)))))
        (setq i (1+ i)))
      (setq verts new-verts)))

  (if (>= (length verts) 3)
    (tz-set-pline-verts obj verts)
    (princ "\n[T24] Warning: cleanup left < 3 vertices, shape unchanged."))
  ent)

;; ── TZ-ZONE ──────────────────────────────────────────────────────────────────
(defun c:TZ-ZONE ( / *error* sel sel2 txt-ent txt-edata txt-str txt-lyr txt-pt
                     ent last-ent pts area-ft centroid
                     zone-id zone-name ceil-ht floor condition occupancy
                     ce cd cl edata froze txt-lyrs-frozen old-gaptol
                     txt-layers choice ldata ed2
                     patch-lines pp1 pp2
                     ss-near j near-ent near-ed near-str near-pt near-dist
                     nearby-texts txt-ins blk-ref blk-name blk-def
                     sub-ent sub-ed)
  ;; Allow (command ...) calls inside *error* handler (AutoCAD 2015+ requirement)
  (*push-error-using-command*)

  (defun *error* (msg)
    (if ce (setvar "CMDECHO" ce))
    (if cd (setvar "CMDDIA"  cd))
    (if cl (setvar "CLAYER"  cl))
    (if froze (tz-thaw-layers froze))
    (if txt-lyrs-frozen (tz-thaw-layers txt-lyrs-frozen))
    (setvar "HPGAPTOL" (if (boundp 'old-gaptol) old-gaptol 0.0))
    (if (not (member msg '("Function cancelled" "quit / exit abort" "")))
      (princ (strcat "\n[T24] Error: " msg)))
    (princ))

  (tz-setup)
  (setq ce nil  cd nil  cl nil  froze nil  txt-lyrs-frozen nil)

  ;; ── Front-load session defaults (asked once, applied to every room) ─────────
  (setq ceil-ht (getreal "\n[T24] Default ceiling height in feet <9>: "))
  (if (null ceil-ht) (setq ceil-ht 9.0))

  (setq floor (getint "\n[T24] Floor / story number <1>: "))
  (if (null floor) (setq floor 1))

  ;; Condition and occupancy default to generic values — sort out in EnergyPro
  (setq condition "Conditioned"
        occupancy "Residential")

  (princ (strcat "\n[T24] Session: " (rtos ceil-ht 2 1) "' ceiling  |  Floor " (itoa floor)))
  (princ "\n[T24] Now click room name text for each zone. Press Enter when done.")

  ;; ── Main room loop ────────────────────────────────────────────────────────
  (while
    (progn
      (setq sel (nentsel "\n[T24] Click room name text (Enter to finish): "))
      (not (null sel)))

    ;; nentsel returns (ent pick-pt matrix parent-ent...)
    (setq txt-ent   (car sel)
          txt-pt    (cadr sel)      ; WCS click point — seed for BOUNDARY
          txt-edata (entget txt-ent)
          txt-lyr   nil             ; reset each iteration — prevents bleed from prior room
          txt-layers '())

    ;; ── Determine zone name: from text entity or typed manually ─────────
    (if (member (cdr (assoc 0 txt-edata)) '("TEXT" "MTEXT"))
      ;; Clicked on text — read it
      (progn
        (setq txt-str (vl-string-trim " " (cdr (assoc 1 txt-edata)))
              txt-lyr (cdr (assoc 8 txt-edata)))
        (if (and txt-lyr (wcmatch txt-lyr "T24-*"))
          ;; Clicked on one of our own labels/markers — warn and skip
          (progn
            (princ "\n[T24] Clicked a T24 entity — click the room name text instead.")
            (setq zone-name nil))
          ;; Valid drawing text
          (progn
            (if txt-lyr (setq txt-layers (list txt-lyr)))
            (if (= txt-str "") (setq txt-str nil))
            (setq zone-name txt-str))))
      ;; Missed text — prompt to type a name
      (progn
        (setq zone-name
          (getstring T "\n[T24] No text found. Type zone name: "))
        (if (= zone-name "") (setq zone-name nil))))

    (if (null zone-name)
      (princ "\n[T24] No name entered, skipping.")

      (progn
        ;; ── Auto-find nearby text to build full name ────────────────────
        ;; Get clicked text position in its coordinate space
        (setq txt-ins (cdr (assoc 10 txt-edata))
              nearby-texts '())
        ;; Prefer alignment point (group 11) if set
        (if (and (assoc 11 txt-edata)
                 (not (equal (cdr (assoc 11 txt-edata)) '(0.0 0.0 0.0) 0.01)))
          (setq txt-ins (cdr (assoc 11 txt-edata))))

        (if (>= (length sel) 4)
          ;; ── Nested entity (xref/block) — walk the block definition ──
          (progn
            (setq blk-ref  (car (cadddr sel))
                  blk-name (cdr (assoc 2 (entget blk-ref)))
                  blk-def  (tblobjname "BLOCK" blk-name))
            (if blk-def
              (progn
                (setq near-ent (entnext blk-def))
                (while near-ent
                  (setq near-ed (entget near-ent))
                  (cond
                    ;; Direct TEXT/MTEXT in block
                    ((and (member (cdr (assoc 0 near-ed)) '("TEXT" "MTEXT"))
                          (not (eq near-ent txt-ent)))
                     (setq near-str (vl-string-trim " " (cdr (assoc 1 near-ed))))
                     (if (/= near-str "")
                       (progn
                         (setq near-pt (cdr (assoc 10 near-ed)))
                         (if (and (assoc 11 near-ed)
                                  (not (equal (cdr (assoc 11 near-ed)) '(0.0 0.0 0.0) 0.01)))
                           (setq near-pt (cdr (assoc 11 near-ed))))
                         (setq near-dist (distance txt-ins near-pt))
                         (if (< near-dist 18.0)
                           (setq nearby-texts
                             (cons (list near-dist near-str) nearby-texts))))))
                    ;; INSERT (sub-block) — check its ATTRIBs for text
                    ((= (cdr (assoc 0 near-ed)) "INSERT")
                     ;; ATTRIBs follow the INSERT as ATTRIB entities
                     (if (= (cdr (assoc 66 near-ed)) 1)  ; has attribs
                       (progn
                         (setq sub-ent (entnext near-ent))
                         (while (and sub-ent
                                     (setq sub-ed (entget sub-ent))
                                     (= (cdr (assoc 0 sub-ed)) "ATTRIB"))
                           (setq near-str (vl-string-trim " " (cdr (assoc 1 sub-ed))))
                           (if (/= near-str "")
                             (progn
                               (setq near-pt (cdr (assoc 10 sub-ed)))
                               (if (and (assoc 11 sub-ed)
                                        (not (equal (cdr (assoc 11 sub-ed)) '(0.0 0.0 0.0) 0.01)))
                                 (setq near-pt (cdr (assoc 11 sub-ed))))
                               (setq near-dist (distance txt-ins near-pt))
                               (if (< near-dist 18.0)
                                 (setq nearby-texts
                                   (cons (list near-dist near-str) nearby-texts)))))
                           (setq sub-ent (entnext sub-ent)))))))
                  (setq near-ent (entnext near-ent)))))))

          ;; ── Not nested — spatial ssget within 18" box only ──
          (if txt-lyr
            (progn
              (setq ss-near (ssget "C"
                              (list (- (car txt-ins) 18.0) (- (cadr txt-ins) 18.0) 0.0)
                              (list (+ (car txt-ins) 18.0) (+ (cadr txt-ins) 18.0) 0.0)
                              (list (cons 8 txt-lyr)
                                    '(-4 . "<OR")
                                      '(0 . "TEXT")
                                      '(0 . "MTEXT")
                                    '(-4 . "OR>"))))
              (if ss-near
                (progn
                  (setq j 0)
                  (repeat (sslength ss-near)
                    (setq near-ent (ssname ss-near j)
                          near-ed  (entget near-ent))
                    (if (not (equal near-ent txt-ent))
                      (progn
                        (setq near-str (vl-string-trim " " (cdr (assoc 1 near-ed))))
                        (if (and (/= near-str "") (/= near-str zone-name))
                          (progn
                            (setq near-pt (cdr (assoc 10 near-ed)))
                            (if (and (assoc 11 near-ed)
                                     (not (equal (cdr (assoc 11 near-ed)) '(0.0 0.0 0.0) 0.01)))
                              (setq near-pt (cdr (assoc 11 near-ed))))
                            (setq near-dist (distance txt-ins near-pt))
                            (if (< near-dist 18.0)
                              (setq nearby-texts
                                (cons (list near-dist near-str) nearby-texts)))))))
                    (setq j (1+ j)))))))

        ;; Sort by distance, append only the closest one
        (if nearby-texts
          (progn
            (setq nearby-texts
              (vl-sort nearby-texts '(lambda (a b) (< (car a) (car b)))))
            (setq zone-name (strcat zone-name " " (cadr (car nearby-texts))))
            (princ (strcat "\n[T24]   Auto-appended: \"" (cadr (car nearby-texts)) "\""))))
        ;; Hard cap at 30 characters
        (if (> (strlen zone-name) 30)
          (setq zone-name (substr zone-name 1 30)))

        (princ (strcat "\n[T24] Zone name: \"" zone-name "\""))

        ;; ── Freeze text layers + door layers before BOUNDARY ──────────────
        (setq ce (getvar "CMDECHO")
              cd (getvar "CMDDIA")
              cl (getvar "CLAYER"))
        (setvar "CMDECHO" 0)
        (setvar "CMDDIA"  0)
        (setvar "CLAYER"  *TZ-LYR-ZONE*)

        ;; Text layers NOT frozen — A-AREA-IDEN and similar layers often contain
        ;; room boundary polylines alongside text; freezing them breaks boundary detection

        ;; ── Run boundary (tz-hatch-boundary manages HPGAPTOL internally) ──
        (setq ent (tz-hatch-boundary txt-pt))

        ;; If failed, let user retry, patch gaps, or fall back
        (setq patch-lines '())
        (while (and (null ent)
                    (progn
                      (princ "\n[T24] BOUNDARY failed at that point.")
                      (setq choice (tz-boundary-popup))
                      (and choice  ;; nil = Cancel → exit loop
                           (or (= choice "Retry") (= choice "Patch")))))

          (if (= choice "Patch")
            ;; ── Patch mode: draw lines to seal gaps, then auto-retry ──
            (progn
              (princ "\n[T24]   Draw lines to seal gaps. Click pairs of points, Enter when done.")
              (while
                (progn
                  (setq pp1 (getpoint "\n[T24]   Line start (Enter to finish patching): "))
                  (not (null pp1)))
                (setq pp2 (getpoint pp1 "\n[T24]   Line end: "))
                (if pp2
                  (progn
                    (tz-make-line pp1 pp2)
                    (setq patch-lines (cons (entlast) patch-lines))
                    (princ "\n[T24]   Patch line drawn."))
                  (princ "\n[T24]   No endpoint, skipped.")))
              (if patch-lines
                (progn
                  (princ (strcat "\n[T24]   " (itoa (length patch-lines))
                                 " patch line(s). Retrying..."))
                  (setq ent (tz-hatch-boundary-full txt-pt)))
                (princ "\n[T24]   No lines drawn.")))

            ;; ── Retry mode: pick a new point, try full gap tolerance range ──
            (progn
              (setq txt-pt (getpoint "\n[T24]   Click inside the room: "))
              (if txt-pt
                (setq ent (tz-hatch-boundary-full txt-pt))
                (princ "\n[T24]   No point picked."))))

          ) ;; end while

        ;; Clean up temporary lines after boundary succeeds or user moves on
        (foreach pent patch-lines (entdel pent))
        (if patch-lines
          (princ (strcat "\n[T24]   Cleaned up " (itoa (length patch-lines)) " patch line(s).")))

        ;; Frozen layers stay frozen across rooms — thawed at end of session below
        (setvar "CMDDIA"  cd)
        (setvar "CMDECHO" ce)
        (setvar "CLAYER"  cl)

        ;; ── Manual fallback if BOUNDARY never succeeded ──────────────────
        (if (null ent)
          (cond
            ((null choice)
             (princ "\n[T24] Cancelled, skipping this zone."))
            ((= choice "Select")
             (progn
               (princ "\n[T24] Select an existing closed polyline: ")
               (setq sel2 (entsel))
               (if (null sel2) (progn (princ "\n[T24] Cancelled.") (setq choice nil)))
               (if sel2
                 (progn
                   (setq ent (car sel2))
                   (if (/= (cdr (assoc 0 (entget ent))) "LWPOLYLINE")
                     (progn (princ "\n[T24] Not a polyline, skipping.") (setq ent nil)))))))
            (T  ;; "Manual" — pick corners
             (setq ent (tz-pick-corners)))))

        (if ent
          (progn
            ;; Move to T24-ZONE layer if BOUNDARY put it elsewhere
            (setq edata (entget ent))
            (if (/= (cdr (assoc 8 edata)) *TZ-LYR-ZONE*)
              (entmod (subst (cons 8 *TZ-LYR-ZONE*) (assoc 8 edata) edata)))

            ;; Clean: flatten arcs, collapse door notches
            (tz-clean-pline ent)

            ;; Compute area and centroid
            (setq pts      (tz-get-pts ent)
                  area-ft  (/ (vlax-curve-getarea (vlax-ename->vla-object ent))
                              (* *TZ-UNIT-FT* *TZ-UNIT-FT*))
                  centroid (tz-centroid pts)
                  zone-id  (tz-next-zone-id))

            ;; Store XDATA
            (tz-set-xdata ent
              (list *TZ-APP*
                (cons 1000 "ZONE")
                (cons 1000 zone-id)
                (cons 1000 zone-name)
                (cons 1040 ceil-ht)
                (cons 1070 floor)
                (cons 1000 condition)
                (cons 1000 occupancy)))

            ;; Visual label — use txt-pt (click point, inside the room)
            (tz-zone-label txt-pt zone-name area-ft ceil-ht floor)

            ;; Bring zone polyline + label to top of draw order
            (command "_.DRAWORDER" ent "" "_Front")
            (command "_.DRAWORDER" (entlast) "" "_Front")

            (princ (strcat "\n[T24] Tagged: \"" zone-name
                           "\"  " (rtos area-ft 2 1) " sqft  Floor " (itoa floor))))
          (princ "\n[T24] No boundary created, skipping this zone."))
            ))) ; end if ent, progn, if zone-name, while

  ;; Thaw any layers frozen during the session (e.g. txt-lyrs-frozen, froze)
  (if txt-lyrs-frozen (tz-thaw-layers txt-lyrs-frozen))
  (if froze (tz-thaw-layers froze))
  (princ "\n[T24] Done tagging zones.")
  (princ))

;; ── TZ-WIN ───────────────────────────────────────────────────────────────────
(defun c:TZ-WIN ( / pt ent open-id area-ft lsave)
  (tz-setup)
  (setq area-ft (getreal "\n[T24] Window area (sqft) <30>: "))
  (if (null area-ft) (setq area-ft 30.0))
  (princ (strcat "\n[T24] Placing " (rtos area-ft 2 1)
                 " sqft windows. Click locations, Enter when done."))
  (while
    (progn
      (setq pt (getpoint "\n[T24] Click window location: "))
      (not (null pt)))
    (setq open-id (tz-next-open-id)
          lsave   (getvar "CLAYER"))
    (setvar "CLAYER" *TZ-LYR-WIN*)
    (tz-make-circle pt *TZ-MARKER-RAD*)
    (setq ent (entlast))
    (setvar "CLAYER" lsave)
    (tz-set-xdata ent
      (list *TZ-APP*
        (cons 1000 "OPENING")
        (cons 1000 open-id)
        (cons 1000 "Window")
        (cons 1040 area-ft)
        (cons 1040 0.0)
        (cons 1040 0.0)))
    (tz-open-label pt open-id "Window" area-ft 0.0 0.0)
    (princ (strcat "\n[T24] " open-id " placed.")))
  (princ "\n[T24] Done placing windows.")
  (princ))

;; ── TZ-DOOR ──────────────────────────────────────────────────────────────────
(defun c:TZ-DOOR ( / pt ent open-id area-ft lsave)
  (tz-setup)
  (setq pt (getpoint "\n[T24] Click at door location: "))
  (if (null pt) (progn (princ "\n[T24] Cancelled.") (exit)))

  (setq open-id (tz-next-open-id))

  (setq area-ft (getreal "\n[T24] Door area (sqft) <21>: "))
  (if (null area-ft) (setq area-ft 21.0))

  (setq lsave (getvar "CLAYER"))
  (setvar "CLAYER" *TZ-LYR-DOOR*)
  (tz-make-circle pt *TZ-MARKER-RAD*)
  (setq ent (entlast))
  (setvar "CLAYER" lsave)

  (tz-set-xdata ent
    (list *TZ-APP*
      (cons 1000 "OPENING")
      (cons 1000 open-id)
      (cons 1000 "Door")
      (cons 1040 area-ft)
      (cons 1040 0.0)
      (cons 1040 0.0)))

  (tz-open-label pt open-id "Door" area-ft 0.0 0.0)
  (princ (strcat "\n[T24] Door " open-id " | " (rtos area-ft 2 1) " sqft"))
  (princ))

;; ── TZ-SKY ───────────────────────────────────────────────────────────────────
(defun c:TZ-SKY ( / pt ent open-id area-ft uf shgc lsave r)
  (tz-setup)
  (setq pt (getpoint "\n[T24] Click at skylight location (inside zone): "))
  (if (null pt) (progn (princ "\n[T24] Cancelled.") (exit)))

  (setq open-id (tz-next-open-id))

  (setq area-ft (getreal "\n[T24] Skylight area (sqft) <12>: "))
  (if (null area-ft) (setq area-ft 12.0))

  (setq uf (getreal "\n[T24] U-Factor <0.50>: "))
  (if (null uf) (setq uf 0.50))

  (setq shgc (getreal "\n[T24] SHGC <0.25>: "))
  (if (null shgc) (setq shgc 0.25))

  (setq lsave (getvar "CLAYER")
        r      *TZ-MARKER-RAD*)
  (setvar "CLAYER" *TZ-LYR-SKY*)
  ;; Circle + X cross to distinguish from windows
  (tz-make-circle pt r)
  (setq ent (entlast))   ; save circle reference before adding lines
  (tz-make-line (list (- (car pt) r) (- (cadr pt) r))
                (list (+ (car pt) r) (+ (cadr pt) r)))
  (tz-make-line (list (+ (car pt) r) (- (cadr pt) r))
                (list (- (car pt) r) (+ (cadr pt) r)))
  (setvar "CLAYER" lsave)

  (tz-set-xdata ent
    (list *TZ-APP*
      (cons 1000 "OPENING")
      (cons 1000 open-id)
      (cons 1000 "Skylight")
      (cons 1040 area-ft)
      (cons 1040 uf)
      (cons 1040 shgc)))

  (tz-open-label pt open-id "Skylight" area-ft uf shgc)
  (princ (strcat "\n[T24] Skylight " open-id " | " (rtos area-ft 2 1) " sqft"))
  (princ))

;; ── Wall helpers ──────────────────────────────────────────────────────────────

;; Count existing wall markers and return next "W-NNN" id
(defun tz-next-wall-id ( / ss)
  (setq ss (ssget "X" (list (cons 8 *TZ-LYR-WALL*) '(0 . "CIRCLE"))))
  (strcat "W-" (tz-pad (itoa (1+ (if ss (sslength ss) 0))) 3)))

;; Scan all T24-ZONE polylines, return (ent p1 p2 idx dist) for closest edge to pt
;; Returns nil if no zone polylines exist
(defun tz-nearest-edge (pt / ss i ent pts n j p1 p2
                            best-ent best-p1 best-p2 best-idx best-dist
                            dx dy seg-len t0 cx cy d)
  (setq ss (ssget "X" (list (cons 8 *TZ-LYR-ZONE*) '(0 . "LWPOLYLINE"))))
  (if (null ss) nil
    (progn
      (setq best-dist 1e20  best-ent nil  i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i)
              pts (tz-get-pts ent)
              n   (length pts)
              j   0)
        (repeat n
          (setq p1 (nth j pts)
                p2 (nth (rem (1+ j) n) pts)
                dx (- (car p2) (car p1))
                dy (- (cadr p2) (cadr p1))
                seg-len (sqrt (+ (* dx dx) (* dy dy))))
          (if (> seg-len 0.1)
            (progn
              ;; Project pt onto segment, clamp t to [0,1]
              (setq t0 (/ (+ (* (- (car pt) (car p1)) dx)
                             (* (- (cadr pt) (cadr p1)) dy))
                          (+ (* dx dx) (* dy dy))))
              (if (< t0 0.0) (setq t0 0.0))
              (if (> t0 1.0) (setq t0 1.0))
              (setq cx (+ (car p1) (* t0 dx))
                    cy (+ (cadr p1) (* t0 dy))
                    d  (sqrt (+ (* (- (car pt) cx) (- (car pt) cx))
                               (* (- (cadr pt) cy) (- (cadr pt) cy)))))
              (if (< d best-dist)
                (setq best-dist d
                      best-ent  ent
                      best-p1   p1
                      best-p2   p2
                      best-idx  j))))
          (setq j (1+ j)))
        (setq i (1+ i)))
      (if best-ent
        (list best-ent best-p1 best-p2 best-idx best-dist)
        nil))))

;; Compute wall azimuth: right-perpendicular of p1→p2, CW from North
(defun tz-wall-azimuth (p1 p2 / dx dy len nx ny az)
  (setq dx (- (car p2) (car p1))
        dy (- (cadr p2) (cadr p1))
        len (sqrt (+ (* dx dx) (* dy dy))))
  (if (< len 1e-6) 0.0
    (progn
      ;; Right perpendicular: (dy, -dx) normalized
      (setq nx (/ dy len)
            ny (/ (- dx) len)
            az (* (/ 180.0 pi) (atan2 nx ny)))
      (if (< az 0.0) (setq az (+ az 360.0)))
      az)))

;; Read zone-id from a polyline's XDATA
(defun tz-zone-of-ent (ent / xd)
  (setq xd (tz-get-xdata ent))
  (if (and xd (equal (tz-xd-nth xd 1000 0) "ZONE"))
    (tz-xd-nth xd 1000 1)
    nil))

;; Place label text at click point for a wall marker
(defun tz-wall-label (pt wall-id wall-type len-ft az / lsave th off)
  (setq lsave (getvar "CLAYER")
        th  (* 0.08 *TZ-UNIT-FT*)
        off (+ *TZ-WALL-RAD* (* 1.5 th)))
  (setvar "CLAYER" *TZ-LYR-LABEL*)
  (tz-make-text
    (list (car pt) (+ (cadr pt) off) 0.0) th
    (strcat wall-id "  " wall-type
            "  " (rtos len-ft 2 1) "'  " (rtos az 2 0) (chr 176)))
  (setvar "CLAYER" lsave))

;; ── TZ-WALL ──────────────────────────────────────────────────────────────────
;; Works like TZ-WIN: click locations, circle goes at click point.
;; Auto-detects nearest polyline edge for length/azimuth/zone (best effort).
(defun c:TZ-WALL ( / pt edge-info edge-ent edge-p1 edge-p2 edge-dist
                     wall-id wall-type wall-ht zone-id
                     len-ft area-sqft az lsave ent type-choice
                     dx dy len-in)
  (tz-setup)

  ;; ── Session defaults ────────────────────────────────────────────────────
  (setq wall-ht (getreal "\n[T24] Wall height in feet <9>: "))
  (if (null wall-ht) (setq wall-ht 9.0))

  (initget "Exterior Interior")
  (setq type-choice (getkword "\n[T24] Wall type [Exterior/Interior] <Exterior>: "))
  (if (or (null type-choice) (= type-choice "Exterior"))
    (setq wall-type "Exterior Wall")
    (setq wall-type "Interior Wall"))

  (princ (strcat "\n[T24] Placing " wall-type " walls, "
                 (rtos wall-ht 2 1) "' height."))
  (princ "\n[T24] Click on walls. Enter when done.")

  ;; ── Click loop ──────────────────────────────────────────────────────────
  (while
    (progn
      (setq pt (getpoint "\n[T24] Click wall location (Enter to finish): "))
      (not (null pt)))

    (setq wall-id (tz-next-wall-id))

    ;; Best-effort: find nearest polyline edge for auto-computed properties
    (setq edge-info (tz-nearest-edge pt)
          edge-p1   nil  edge-p2 nil
          len-ft    0.0  az 0.0  zone-id "Z-???")

    (if edge-info
      (progn
        (setq edge-ent  (nth 0 edge-info)
              edge-p1   (nth 1 edge-info)
              edge-p2   (nth 2 edge-info)
              edge-dist (nth 4 edge-info)
              dx        (- (car edge-p2) (car edge-p1))
              dy        (- (cadr edge-p2) (cadr edge-p1))
              len-in    (sqrt (+ (* dx dx) (* dy dy)))
              len-ft    (/ len-in *TZ-UNIT-FT*)
              az        (tz-wall-azimuth edge-p1 edge-p2)
              zone-id   (or (tz-zone-of-ent edge-ent) "Z-???"))))

    (setq area-sqft (* len-ft wall-ht))

    ;; Place marker circle at CLICK POINT on T24-WALL layer
    (setq lsave (getvar "CLAYER"))
    (setvar "CLAYER" *TZ-LYR-WALL*)
    (tz-make-circle pt *TZ-WALL-RAD*)
    (setq ent (entlast))
    (setvar "CLAYER" lsave)

    ;; Store XDATA on wall marker
    (tz-set-xdata ent
      (list *TZ-APP*
        (cons 1000 "WALL")
        (cons 1000 wall-id)
        (cons 1000 wall-type)
        (cons 1000 zone-id)
        (cons 1040 wall-ht)
        (cons 1040 len-ft)
        (cons 1040 area-sqft)
        (cons 1040 az)
        (cons 1040 (if edge-p1 (car  edge-p1) (car  pt)))
        (cons 1040 (if edge-p1 (cadr edge-p1) (cadr pt)))
        (cons 1040 (if edge-p2 (car  edge-p2) (car  pt)))
        (cons 1040 (if edge-p2 (cadr edge-p2) (cadr pt)))))

    ;; Visual label
    (tz-wall-label pt wall-id wall-type len-ft az)

    (princ (strcat "\n[T24] " wall-id ": " (rtos len-ft 2 1) "' "
                   wall-type " in " zone-id
                   " (" (rtos az 2 0) (chr 176) ")")))

  (princ "\n[T24] Done placing walls.")
  (princ))

;; ── TZ-LISTDATA ──────────────────────────────────────────────────────────────
(defun c:TZ-LISTDATA ( / ss i ent xd)
  (setq ss (ssget "X" (list (cons 8 *TZ-LYR-ZONE*) '(0 . "LWPOLYLINE"))))
  (if (null ss) (progn (princ "\n[T24] No zones found.") (exit)))
  (princ "\n")
  (princ "\n--- T24 Zone Data ---")
  (setq i 0)
  (repeat (sslength ss)
    (setq ent (ssname ss i)
          xd  (tz-get-xdata ent))
    (if (and xd (equal (tz-xd-nth xd 1000 0) "ZONE"))
      (princ (strcat "\n  " (tz-xd-nth xd 1000 1)
                     "  |  " (tz-xd-nth xd 1000 2)
                     "  |  " (rtos (/ (vlax-curve-getarea (vlax-ename->vla-object ent))
                                      144.0) 2 1) " sqft"
                     "  |  Fl " (itoa (tz-xd-num xd 1070 0 1))
                     "  |  " (tz-xd-nth xd 1000 3))))
    (setq i (1+ i)))
  (princ "\n---------------------")
  (princ))

;; ── TZ-SHOWVERTS ─────────────────────────────────────────────────────────────
;; Select a polyline, labels each vertex with its interior angle for debugging
(defun c:TZ-SHOWVERTS ( / ent obj verts n i v0 v1 v2
                          dx1 dy1 dx2 dy2 denom cross-val ang lsave th)
  (setq ent (car (entsel "\n[T24] Select polyline to inspect: ")))
  (if (null ent) (progn (princ "\n[T24] Cancelled.") (exit)))
  (setq obj   (vlax-ename->vla-object ent)
        verts (tz-pline-verts obj)
        n     (length verts)
        lsave (getvar "CLAYER")
        th    (* 0.08 *TZ-UNIT-FT*))
  (setvar "CLAYER" *TZ-LYR-LABEL*)
  (setq i 0)
  (repeat n
    (setq v0   (nth (rem (+ n i -1) n) verts)
          v1   (nth i verts)
          v2   (nth (rem (1+ i) n) verts)
          dx1  (- (car v1)  (car v0))  dy1 (- (cadr v1) (cadr v0))
          dx2  (- (car v2)  (car v1))  dy2 (- (cadr v2) (cadr v1))
          denom (* (sqrt (+ (* dx1 dx1) (* dy1 dy1)))
                   (sqrt (+ (* dx2 dx2) (* dy2 dy2)))))
    (setq ang (if (> denom 0.0)
               (* (/ 180.0 pi)
                  (atan (sqrt (max 0.0 (- 1.0 (* (/ (+ (* dx1 (- dx2)) (* dy1 (- dy2))) denom)
                                                  (/ (+ (* dx1 (- dx2)) (* dy1 (- dy2))) denom)))))
                        (/ (+ (* dx1 (- dx2)) (* dy1 (- dy2))) denom)))
               0.0))
    ;; simpler: use cross product to get sin, dot product to get cos
    (if (> denom 0.0)
      (progn
        (setq cross-val (/ (- (* dx1 dy2) (* dy1 dx2)) denom)
              ang       (* (/ 180.0 pi) (atan (abs cross-val)
                              (/ (+ (* dx1 dx2) (* dy1 dy2)) denom))))))
    (tz-make-text
      (list (+ (car v1) 3) (+ (cadr v1) 3) 0.0)
      th
      (strcat (itoa i) ":" (rtos ang 2 1) (chr 176)))
    (setq i (1+ i)))
  (setvar "CLAYER" lsave)
  (princ (strcat "\n[T24] Labeled " (itoa n) " vertices."))
  (princ))

;; ── JSON helpers for TZ-EXPORT ───────────────────────────────────────────────

(defun tz-jstr (s / )
  ;; JSON-encode a string — escape backslashes and double-quotes
  (setq s (vl-string-subst "\\\\" "\\" s)
        s (vl-string-subst "\\\"" "\"" s))
  (strcat "\"" s "\""))

(defun tz-jnum (v / )
  (rtos (float v) 2 4))

(defun tz-jpt (x y / )
  (strcat "[" (tz-jnum x) "," (tz-jnum y) "]"))

(defun tz-jpts (pts / result first-p)
  ;; Build [[x1,y1],[x2,y2],...] from list of (x y) pairs
  (setq result "[" first-p T)
  (foreach p pts
    (if first-p (setq first-p nil) (setq result (strcat result ",")))
    (setq result (strcat result "[" (tz-jnum (car p)) "," (tz-jnum (cadr p)) "]")))
  (strcat result "]"))

;; ── TZ-EXPORT ─────────────────────────────────────────────────────────────────
;; Scans all T24-tagged entities in the drawing and writes <dwg>_t24.json
;; Then run: python tz_to_excel.py "<dwg>_t24.json"  to get Excel + geometry
(defun c:TZ-EXPORT ( / path dwg-dir dwg-name fp
                       ss i ent xd obj pts area-sqft cent
                       zid zname ht fl cond occ
                       wid wtype wzid wht wlen warea waz wp1x wp1y wp2x wp2y
                       opid optype oparea opuf opshgc opctr
                       zone-started wall-started open-started
                       n-zones n-walls n-opens)

  (setq dwg-dir  (getvar "DWGPREFIX")
        dwg-name (vl-filename-base (getvar "DWGNAME"))
        path     (strcat dwg-dir dwg-name "_t24.json"))

  (setq fp (open path "w"))
  (if (null fp)
    (progn (alert (strcat "[T24] Cannot write to:\n" path)) (exit)))

  (princ (strcat "\n[T24] Writing: " path))

  (write-line "{" fp)
  (write-line "  \"zones\": [" fp)

  ;; ── Zones ────────────────────────────────────────────────────────────────
  (setq ss           (ssget "X" (list (cons 8 *TZ-LYR-ZONE*) '(0 . "LWPOLYLINE")))
        zone-started nil
        n-zones      0)

  (if ss
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i)
              xd  (tz-get-xdata ent))
        (if (and xd (equal (tz-xd-nth xd 1000 0) "ZONE"))
          (progn
            (setq obj       (vlax-ename->vla-object ent)
                  pts       (tz-get-pts ent)
                  area-sqft (/ (vlax-curve-getarea obj)
                               (* *TZ-UNIT-FT* *TZ-UNIT-FT*))
                  cent      (tz-centroid pts)
                  zid       (or (tz-xd-nth xd 1000 1) "Z-???")
                  zname     (or (tz-xd-nth xd 1000 2) "")
                  ht        (tz-xd-num xd 1040 0 9.0)
                  fl        (tz-xd-num xd 1070 0 1)
                  cond      (or (tz-xd-nth xd 1000 3) "Conditioned")
                  occ       (or (tz-xd-nth xd 1000 4) "Residential"))
            (if zone-started (write-line "    }," fp))
            (write-line "    {" fp)
            (write-line (strcat "      \"id\": "            (tz-jstr zid)                       ",") fp)
            (write-line (strcat "      \"name\": "          (tz-jstr zname)                     ",") fp)
            (write-line (strcat "      \"vertices\": "      (tz-jpts pts)                       ",") fp)
            (write-line (strcat "      \"centroid\": "      (tz-jpt (car cent) (cadr cent))      ",") fp)
            (write-line (strcat "      \"area_sqft\": "     (tz-jnum area-sqft)                  ",") fp)
            (write-line (strcat "      \"ceiling_ht_ft\": " (tz-jnum ht)                        ",") fp)
            (write-line (strcat "      \"floor\": "         (itoa (fix fl))                     ",") fp)
            (write-line (strcat "      \"condition\": "     (tz-jstr cond)                      ",") fp)
            (write-line (strcat "      \"occupancy\": "     (tz-jstr occ))                          fp)
            (setq zone-started T  n-zones (1+ n-zones))))
        (setq i (1+ i)))))

  (if zone-started (write-line "    }" fp))
  (write-line "  ]," fp)
  (write-line "  \"walls\": [" fp)

  ;; ── Walls ─────────────────────────────────────────────────────────────────
  (setq ss           (ssget "X" (list (cons 8 *TZ-LYR-WALL*) '(0 . "CIRCLE")))
        wall-started nil
        n-walls      0)

  (if ss
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i)
              xd  (tz-get-xdata ent))
        (if (and xd (equal (tz-xd-nth xd 1000 0) "WALL"))
          (progn
            (setq wid   (or (tz-xd-nth xd 1000 1) "W-???")
                  wtype (or (tz-xd-nth xd 1000 2) "Exterior Wall")
                  wzid  (or (tz-xd-nth xd 1000 3) "Z-???")
                  wht   (tz-xd-num xd 1040 0 9.0)
                  wlen  (tz-xd-num xd 1040 1 0.0)
                  warea (tz-xd-num xd 1040 2 0.0)
                  waz   (tz-xd-num xd 1040 3 0.0)
                  wp1x  (tz-xd-num xd 1040 4 0.0)
                  wp1y  (tz-xd-num xd 1040 5 0.0)
                  wp2x  (tz-xd-num xd 1040 6 0.0)
                  wp2y  (tz-xd-num xd 1040 7 0.0))
            (if wall-started (write-line "    }," fp))
            (write-line "    {" fp)
            (write-line (strcat "      \"id\": "         (tz-jstr wid)                    ",") fp)
            (write-line (strcat "      \"type\": "       (tz-jstr wtype)                  ",") fp)
            (write-line (strcat "      \"zone_id\": "    (tz-jstr wzid)                   ",") fp)
            (write-line (strcat "      \"height_ft\": "  (tz-jnum wht)                    ",") fp)
            (write-line (strcat "      \"length_ft\": "  (tz-jnum wlen)                   ",") fp)
            (write-line (strcat "      \"area_sqft\": "  (tz-jnum warea)                  ",") fp)
            (write-line (strcat "      \"azimuth\": "    (tz-jnum waz)                    ",") fp)
            (write-line (strcat "      \"p1\": "         (tz-jpt wp1x wp1y)               ",") fp)
            (write-line (strcat "      \"p2\": "         (tz-jpt wp2x wp2y))                  fp)
            (setq wall-started T  n-walls (1+ n-walls))))
        (setq i (1+ i)))))

  (if wall-started (write-line "    }" fp))
  (write-line "  ]," fp)
  (write-line "  \"openings\": [" fp)

  ;; ── Openings ─────────────────────────────────────────────────────────────
  (setq open-started nil  n-opens 0)

  (foreach lyr (list *TZ-LYR-WIN* *TZ-LYR-DOOR* *TZ-LYR-SKY*)
    (setq ss (ssget "X" (list (cons 8 lyr) '(0 . "CIRCLE"))))
    (if ss
      (progn
        (setq i 0)
        (repeat (sslength ss)
          (setq ent (ssname ss i)
                xd  (tz-get-xdata ent))
          (if (and xd (equal (tz-xd-nth xd 1000 0) "OPENING"))
            (progn
              (setq opctr  (cdr (assoc 10 (entget ent)))
                    opid   (or (tz-xd-nth xd 1000 1) "O-???")
                    optype (or (tz-xd-nth xd 1000 2) "Window")
                    oparea (tz-xd-num xd 1040 0 0.0)
                    opuf   (tz-xd-num xd 1040 1 0.0)
                    opshgc (tz-xd-num xd 1040 2 0.0))
              (if open-started (write-line "    }," fp))
              (write-line "    {" fp)
              (write-line (strcat "      \"id\": "       (tz-jstr opid)                       ",") fp)
              (write-line (strcat "      \"type\": "     (tz-jstr optype)                     ",") fp)
              (write-line (strcat "      \"position\": " (tz-jpt (car opctr) (cadr opctr))    ",") fp)
              (write-line (strcat "      \"area_sqft\": " (tz-jnum oparea)                    ",") fp)
              (write-line (strcat "      \"u_factor\": "  (tz-jnum opuf)                      ",") fp)
              (write-line (strcat "      \"shgc\": "      (tz-jnum opshgc))                       fp)
              (setq open-started T  n-opens (1+ n-opens))))
          (setq i (1+ i))))))

  (if open-started (write-line "    }" fp))
  (write-line "  ]" fp)
  (write-line "}" fp)
  (close fp)

  (princ (strcat "\n[T24] Exported " (itoa n-zones) " zones, "
                 (itoa n-walls) " walls, " (itoa n-opens) " openings"))
  (princ (strcat "\n[T24] File: " path))
  (princ (strcat "\n[T24] Next: python tz_to_excel.py \"" path "\""))
  (princ))

;; ── TZ-RESET ─────────────────────────────────────────────────────────────────
(defun c:TZ-RESET ( / ans ss i ent)
  (setq ans (getstring "\n[T24] Delete ALL T24 labels and markers? Type YES to confirm: "))
  (if (/= (strcase ans) "YES")
    (progn (princ "\n[T24] Cancelled.") (exit)))
  (foreach lyr (list *TZ-LYR-LABEL* *TZ-LYR-WIN* *TZ-LYR-DOOR* *TZ-LYR-SKY* *TZ-LYR-WALL*)
    (setq ss (ssget "X" (list (cons 8 lyr))))
    (if ss
      (progn
        (setq i 0)
        (repeat (sslength ss)
          (entdel (ssname ss i))
          (setq i (1+ i))))))
  (princ "\n[T24] Labels and markers cleared. Zone polylines and XDATA retained.")
  (princ))

;; ── Load message ─────────────────────────────────────────────────────────────
(princ "\n")
(princ "\n+--------------------------------------------+")
(princ "\n|    T24 TakeOff Wizard  v2 Loaded           |")
(princ "\n+--------------------------------------------+")
(princ "\n|  TZ-ZONE      - Click room name text       |")
(princ "\n|  TZ-WALL      - Place wall markers         |")
(princ "\n|  TZ-WIN       - Place window markers       |")
(princ "\n|  TZ-DOOR      - Place door marker          |")
(princ "\n|  TZ-SKY       - Place skylight marker      |")
(princ "\n|  TZ-EXPORT    - Write _t24.json for Excel  |")
(princ "\n|  TZ-LISTDATA  - List zone data             |")
(princ "\n|  TZ-SHOWVERTS - Inspect polyline vertices  |")
(princ "\n|  TZ-RESET     - Clear labels/markers       |")
(princ "\n+--------------------------------------------+")
(princ "\n")
(princ)
