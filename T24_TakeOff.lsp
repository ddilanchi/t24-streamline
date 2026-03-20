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
;;;   TZ-WALL   - Place wall markers on individual edge segments
;;;   TZ-WSIDE  - Select zone polyline, click outside to indicate side;
;;;               uses bounding-box projection for full wall length
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
  (tz-make-layer *TZ-LYR-ZONE*  6  "Continuous")  ; magenta — zone boundaries
  (tz-make-layer *TZ-LYR-LABEL* 7  "Continuous")  ; white — labels
  (tz-make-layer *TZ-LYR-WIN*   4  "Continuous")  ; cyan/light blue — windows
  (tz-make-layer *TZ-LYR-DOOR*  3  "Continuous")  ; green — doors
  (tz-make-layer *TZ-LYR-SKY*   2  "Continuous")  ; yellow — skylights
  (tz-make-layer *TZ-LYR-WALL*  1  "Continuous")  ; red — walls
  ;; Ensure DASHED linetype is available for dimension lines
  (if (null (tblsearch "LTYPE" "DASHED"))
    (command "_.LINETYPE" "_Load" "DASHED" "" "")))

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

(defun tz-make-circle (ctr r / prev)
  (setq prev (entlast))
  (entmake
    (list '(0 . "CIRCLE") '(100 . "AcDbEntity") '(100 . "AcDbCircle")
          (cons 10 (list (car ctr) (cadr ctr) 0.0))
          (cons 40 r)
          '(370 . 50)))   ; lineweight 0.50mm — thick/visible marker
  (if (equal (entlast) prev)
    (progn (princ "\n[T24] WARNING: Failed to create marker circle.") nil)
    (entlast)))

(defun tz-make-line (p1 p2 / )
  (entmake
    (list '(0 . "LINE") '(100 . "AcDbEntity") '(100 . "AcDbLine")
          (cons 10 (list (car p1) (cadr p1) 0.0))
          (cons 11 (list (car p2) (cadr p2) 0.0)))))

;; Bring all T24 entities to front of draw order
(defun tz-bring-to-front ( / ss)
  (foreach lyr (list *TZ-LYR-ZONE* *TZ-LYR-LABEL* *TZ-LYR-WALL*
                     *TZ-LYR-WIN* *TZ-LYR-DOOR* *TZ-LYR-SKY*)
    (setq ss (ssget "X" (list (cons 8 lyr))))
    (if ss (command "_.DRAWORDER" ss "" "_Front"))))

;; ── Visibility check ────────────────────────────────────────────────────────
;; Returns T if entity's layer is thawed and on (i.e. visually present).
(defun tz-ent-visible-p (ent / lyr-name lyr-ent flags)
  (setq lyr-name (cdr (assoc 8 (entget ent)))
        lyr-ent  (tblobjname "LAYER" lyr-name))
  (if (null lyr-ent) nil
    (progn
      (setq flags (cdr (assoc 70 (entget lyr-ent))))
      ;; Bit 1 = frozen.  Negative color (group 62) = layer off.
      (and (= (logand flags 1) 0)
           (> (cdr (assoc 62 (entget lyr-ent))) 0)))))

;; Returns T if a layer name is thawed and on.
;; For block-def entities on layer "0", pass the parent INSERT's layer instead.
(defun tz-layer-visible-p (lyr-name / lyr-ent flags)
  (setq lyr-ent (tblobjname "LAYER" lyr-name))
  (if (null lyr-ent) nil
    (progn
      (setq flags (cdr (assoc 70 (entget lyr-ent))))
      (and (= (logand flags 1) 0)
           (> (cdr (assoc 62 (entget lyr-ent))) 0)))))

;; ── XDATA writers/readers ────────────────────────────────────────────────────

(defun tz-set-xdata (ent xlist / edata grp3 others)
  (tz-regapp)
  (setq edata (entget ent (list *TZ-APP*)))
  ;; Remove existing T24TOW xdata, preserving other apps' xdata
  (setq grp3 (assoc -3 edata))
  (if grp3
    (progn
      ;; Keep all app entries except T24TOW
      (setq others (vl-remove-if
        '(lambda (x) (= (car x) *TZ-APP*))
        (cdr grp3)))
      ;; Strip the old -3 group
      (setq edata (vl-remove-if '(lambda (x) (= (car x) -3)) edata))
      ;; Re-add other apps' xdata if any
      (if others
        (setq edata (append edata (list (cons -3 others)))))
    )
  )
  (entmod (append edata (list (cons -3 (list xlist))))))

;; Get raw xdata list for T24TOW appid from an entity
(defun tz-get-xdata (ent / raw grp)
  (setq raw (entget ent (list *TZ-APP*)))
  (setq grp (assoc -3 raw))
  (if grp (cdr (assoc *TZ-APP* (cdr grp))) nil))

;; Get the nth occurrence of a group-code value in an xdata list
(defun tz-xd-nth (xd grp idx / matches val)
  (setq matches '())
  (foreach item xd
    (if (= (car item) grp)
      (setq matches (append matches (list (cdr item))))))
  (if (< idx (length matches))
    (progn
      (setq val (nth idx matches))
      ;; Safety guard: if result is boolean T, treat as nil
      (if (= val T) nil val))
    nil))

;; Safe numeric xdata reader — guarantees a number, never T or nil
(defun tz-xd-num (xd grp idx default / val)
  (setq val (tz-xd-nth xd grp idx))
  (if (numberp val) val default))

;; ── Auto ID generators ───────────────────────────────────────────────────────

(defun tz-pad (s n / )
  (while (< (strlen s) n) (setq s (strcat "0" s)))
  s)

;; Session counters — initialized from drawing on first call, then only increment
(defun tz-next-zone-id ( / ss)
  (if (null *TZ-ZONE-COUNT*)
    (progn
      (setq ss (ssget "X" (list (cons 8 *TZ-LYR-ZONE*) '(0 . "LWPOLYLINE"))))
      (setq *TZ-ZONE-COUNT* (if ss (sslength ss) 0))))
  (setq *TZ-ZONE-COUNT* (1+ *TZ-ZONE-COUNT*))
  (strcat "Z-" (tz-pad (itoa *TZ-ZONE-COUNT*) 3)))

(defun tz-next-open-id ( / ss1 ss2 ss3)
  (if (null *TZ-OPEN-COUNT*)
    (progn
      (setq ss1 (ssget "X" (list (cons 8 *TZ-LYR-WIN*)  '(0 . "CIRCLE")))
            ss2 (ssget "X" (list (cons 8 *TZ-LYR-DOOR*) '(0 . "CIRCLE")))
            ss3 (ssget "X" (list (cons 8 *TZ-LYR-SKY*)  '(0 . "CIRCLE"))))
      (setq *TZ-OPEN-COUNT* (+ (if ss1 (sslength ss1) 0)
                                (if ss2 (sslength ss2) 0)
                                (if ss3 (sslength ss3) 0)))))
  (setq *TZ-OPEN-COUNT* (1+ *TZ-OPEN-COUNT*))
  (strcat "O-" (tz-pad (itoa *TZ-OPEN-COUNT*) 3)))

;; ── Visual labels ─────────────────────────────────────────────────────────────

(defun tz-zone-label (centroid zone-name area-ft ceil-ht floor zone-id / lsave th lbl-ent)
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
  ;; Tag the detail label with zone-id XDATA so the reactor can find it
  (if zone-id
    (progn
      (setq lbl-ent (entlast))
      (if lbl-ent
        (tz-set-xdata lbl-ent
          (list *TZ-APP*
            (cons 1000 "ZONE-LABEL")
            (cons 1000 zone-id))))))
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

;; ── Session setup dialog ─────────────────────────────────────────────────────
;; Popup dialog to collect session defaults before TZ-ZONE starts.
;; Sets outer-scope variables: ceil-ht, floor, gap-tol, *TZ-CLIMATE-ZONE*, *TZ-NORTH-ANGLE*
;; Drawing-level settings (climate zone, north arrow) persist as globals.
;; Per-session settings (ceiling ht, floor, gap tol) also remember last-used
;; values for convenience but are expected to change between runs.
;; Front orientation is always 0 (EnergyPro default).
(defun tz-session-dialog ( / dcl-path f dcl-id status)
  ;; Initialize on first run only — remembers last-used values for pre-fill
  (if (null *TZ-LAST-CEIL*)     (setq *TZ-LAST-CEIL*     9.0))
  (if (null *TZ-LAST-FLOOR*)    (setq *TZ-LAST-FLOOR*    1))
  (if (null *TZ-LAST-GAP*)      (setq *TZ-LAST-GAP*      1.0))
  (if (null *TZ-CLIMATE-ZONE*)  (setq *TZ-CLIMATE-ZONE*  "3"))
  (if (null *TZ-NORTH-ANGLE*)   (setq *TZ-NORTH-ANGLE*   0.0))

  (setq dcl-path (vl-filename-mktemp "tz_sess" nil ".dcl"))
  (setq f (open dcl-path "w"))
  (write-line "tz_sess : dialog {" f)
  (write-line "  label = \"T24 Zone Setup\";" f)
  (write-line "  : column {" f)
  (write-line "    : row {" f)
  (write-line "      : edit_box { key = \"ceil\";   label = \"Ceiling Height (ft)\"; edit_width = 8; }" f)
  (write-line "      : edit_box { key = \"floor\";  label = \"Floor Number\";        edit_width = 8; }" f)
  (write-line "    }" f)
  (write-line "    : row {" f)
  (write-line "      : edit_box { key = \"gap\";    label = \"Gap Tolerance (in)\";  edit_width = 8; }" f)
  (write-line "      : edit_box { key = \"czone\";  label = \"Climate Zone\";        edit_width = 8; }" f)
  (write-line "    }" f)
  (write-line "    : edit_box { key = \"north\"; label = \"North Arrow Direction (deg in drawing)\"; edit_width = 8; }" f)
  (write-line "    : spacer { height = 0.3; }" f)
  (write-line "    : row {" f)
  (write-line "      : button { key = \"accept\"; label = \"OK\"; is_default = true; width = 12; fixed_width = true; }" f)
  (write-line "      : button { key = \"cancel\"; label = \"Cancel\"; is_cancel = true; width = 12; fixed_width = true; }" f)
  (write-line "    }" f)
  (write-line "  }" f)
  (write-line "}" f)
  (close f)

  (setq dcl-id (load_dialog dcl-path))
  (if (< dcl-id 0)
    (progn (princ "\n[T24] Dialog load failed, using defaults.") nil)
    (progn
      (new_dialog "tz_sess" dcl-id)
      ;; Pre-fill with last-used / persistent values
      (set_tile "ceil"  (rtos *TZ-LAST-CEIL* 2 1))
      (set_tile "floor" (itoa *TZ-LAST-FLOOR*))
      (set_tile "gap"   (rtos *TZ-LAST-GAP* 2 1))
      (set_tile "czone" *TZ-CLIMATE-ZONE*)
      (set_tile "north" (rtos *TZ-NORTH-ANGLE* 2 1))

      ;; Callbacks — save all values; ceiling/floor/gap remembered for next pre-fill
      (action_tile "accept" "(setq *TZ-LAST-CEIL*    (atof (get_tile \"ceil\"))
                                   *TZ-LAST-FLOOR*   (atoi (get_tile \"floor\"))
                                   *TZ-LAST-GAP*     (atof (get_tile \"gap\"))
                                   *TZ-CLIMATE-ZONE*  (get_tile \"czone\")
                                   *TZ-NORTH-ANGLE*   (atof (get_tile \"north\")))
                              (done_dialog 1)")
      (action_tile "cancel" "(done_dialog 0)")

      (setq status (start_dialog))
      (unload_dialog dcl-id)
      (vl-file-delete dcl-path)

      ;; Validate
      (if (= status 0) (progn (princ "\n[T24] Cancelled.") (exit)))
      (if (<= *TZ-LAST-CEIL* 0)        (setq *TZ-LAST-CEIL* 9.0))
      (if (<= *TZ-LAST-FLOOR* 0)       (setq *TZ-LAST-FLOOR* 1))
      (if (<= *TZ-LAST-GAP* 0)         (setq *TZ-LAST-GAP* 1.0))
      (if (= *TZ-CLIMATE-ZONE* "")     (setq *TZ-CLIMATE-ZONE* "3"))

      ;; Copy to local vars used by the rest of tz-zone
      (setq ceil-ht *TZ-LAST-CEIL*
            floor   *TZ-LAST-FLOOR*
            gap-tol *TZ-LAST-GAP*)
    )))

;; ── Popup choice dialog (appears near cursor) ───────────────────────────────
(defun tz-boundary-popup ( / dcl-path dcl-id result f)
  "Shows a dialog with boundary failure options. Returns keyword string."
  (setq dcl-path (vl-filename-mktemp "tz_bnd" nil ".dcl"))
  (setq f (open dcl-path "w"))
  (write-line "tz_bnd : dialog {" f)
  (write-line "  label = \"BOUNDARY Failed\";" f)
  (write-line "  : column {" f)
  (write-line "    : button { key = \"retry\";  label = \"&Retry at new point\"; width = 32; fixed_width = true; }" f)
  (write-line "    : button { key = \"patch\";  label = \"&Patch gaps (draw lines)\"; width = 32; fixed_width = true; }" f)
  (write-line "    : spacer { height = 0.2; }" f)
  (write-line "    : button { key = \"manual\"; label = \"&Draw polyline (click corners)\"; width = 32; fixed_width = true; }" f)
  (write-line "    : button { key = \"rect\";   label = \"Draw &rectangle (2 corners)\"; width = 32; fixed_width = true; }" f)
  (write-line "    : button { key = \"select\"; label = \"&Select existing polyline\"; width = 32; fixed_width = true; }" f)
  (write-line "    : spacer { height = 0.3; }" f)
  (write-line "    : button { key = \"cancel\"; label = \"Cancel\"; is_cancel = true; width = 32; fixed_width = true; }" f)
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
      (action_tile "rect"   "(setq result \"Rect\")   (done_dialog 1)")
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

;; ── Rectangle (two-corner) polyline builder ──────────────────────────────────
(defun tz-pick-rectangle ( / p1 p2 x1 y1 x2 y2 elist)
  (princ "\n[T24] Click two opposite corners of the rectangle.")
  (setq p1 (getpoint "\n[T24]   First corner: "))
  (if (null p1) (progn (princ "\n[T24] Cancelled.") (setq p1 nil)))
  (if p1
    (progn
      (setq p2 (getcorner p1 "\n[T24]   Opposite corner: "))
      (if (null p2) (progn (princ "\n[T24] Cancelled.") (setq p2 nil)))))
  (if (and p1 p2)
    (progn
      (setq x1 (car p1) y1 (cadr p1)
            x2 (car p2) y2 (cadr p2))
      (setq elist
        (list '(0 . "LWPOLYLINE")
              '(100 . "AcDbEntity")
              (cons 8 *TZ-LYR-ZONE*)
              '(100 . "AcDbPolyline")
              '(90 . 4)
              '(70 . 1)  ; closed
              (cons 10 (list x1 y1))
              (cons 10 (list x2 y1))
              (cons 10 (list x2 y2))
              (cons 10 (list x1 y2))))
      (entmake elist)
      (entlast))))

;; ── Layer freeze/thaw helpers ────────────────────────────────────────────────
;; Use _.-LAYER command (command-line version, NOT dialog _.LAYER).
;; entmod and VLA both silently fail on XREF-dependent layers.

(defun tz-freeze-layer (lname)
  (command "_.-LAYER" "_Freeze" lname "")
  (princ))

(defun tz-thaw-layer (lname)
  (command "_.-LAYER" "_Thaw" lname "")
  (princ))

(defun tz-thaw-layers (layer-names)
  (foreach lname layer-names (tz-thaw-layer lname)))


;; ── Boundary creation ────────────────────────────────────────────────────────
;; Uses _-BOUNDARY at a fixed gap tolerance set once per session (default 1").
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

;; Run BOUNDARY at pt using the session gap tolerance. Returns polyline or nil.
;; Tries at current zoom first, then zooms to a local window around the point,
;; then falls back to zoom extents. Runs REGENALL once before starting.
(defun tz-hatch-boundary (pt gap-tol / ent local-rad)
  (command "_.REGENALL")
  ;; Try 1: current zoom (fastest)
  (setq ent (tz-try-boundary pt gap-tol))
  ;; Try 2: zoom to local area around the pick point (~100' radius)
  (if (null ent)
    (progn
      (setq local-rad 1200.0)  ; 100 feet in inches
      (command "_.ZOOM" "_Window"
        (list (- (car pt) local-rad) (- (cadr pt) local-rad))
        (list (+ (car pt) local-rad) (+ (cadr pt) local-rad)))
      (setq ent (tz-try-boundary pt gap-tol))
      (command "_.ZOOM" "_Previous")))
  ;; Try 3: zoom extents (slowest, last resort)
  (if (null ent)
    (progn
      (command "_.ZOOM" "_Extents")
      (setq ent (tz-try-boundary pt gap-tol))
      (command "_.ZOOM" "_Previous")))
  ent)

;; ── Hatch-based boundary detection ──────────────────────────────────────────
;; Creates a temporary HATCH, extracts its associative boundary polyline,
;; then deletes the hatch. Works where _-BOUNDARY fails (XREF geometry).
(defun tz-try-hatch-boundary (pt gap-tol / old-gaptol old-hpbound old-hpname
                                           old-hpassoc last-ent
                                           hatch-ent scan-ent ent etype
                                           best-area cur-area)
  ;; Save and set sysvars
  (setq old-gaptol  (getvar "HPGAPTOL")
        old-hpbound (getvar "HPBOUNDRETAIN")
        old-hpname  (getvar "HPNAME")
        old-hpassoc (getvar "HPASSOC"))
  (setvar "HPGAPTOL" gap-tol)
  (setvar "HPBOUNDRETAIN" 1)   ;; retain boundary polyline
  (setvar "HPNAME" "SOLID")    ;; solid fill
  (setvar "HPASSOC" 0)         ;; non-associative (simpler cleanup)
  (setq last-ent (entlast))

  ;; Simple hatch: pick internal point, press Enter to finish
  (setq *tz-t1* (getvar "MILLISECS"))
  (command "_-HATCH" pt "")
  (princ (strcat "\n[TIME]     _-HATCH cmd: " (itoa (- (getvar "MILLISECS") *tz-t1*)) "ms"))

  ;; Collect all new entities, find largest polyline, delete rest
  (setq *tz-t1* (getvar "MILLISECS"))
  (if (not (equal (entlast) last-ent))
    (progn
      (setq scan-ent (entnext last-ent)  ent nil  best-area 0.0  all-new '())
      (while scan-ent
        (setq all-new (cons scan-ent all-new))
        (setq scan-ent (entnext scan-ent)))
      ;; Find largest polyline
      (foreach e all-new
        (if (and (entget e) (= (cdr (assoc 0 (entget e))) "LWPOLYLINE"))
          (progn
            (setq cur-area (vlax-curve-getarea (vlax-ename->vla-object e)))
            (if (> cur-area best-area)
              (setq best-area cur-area  ent e)))))
      ;; Delete everything except the keeper
      (foreach e all-new
        (if (not (equal e ent)) (entdel e)))))

  (princ (strcat "\n[TIME]     entity cleanup: " (itoa (- (getvar "MILLISECS") *tz-t1*)) "ms"))

  ;; Restore sysvars
  (setvar "HPGAPTOL" old-gaptol)
  (setvar "HPBOUNDRETAIN" old-hpbound)
  (setvar "HPNAME" old-hpname)
  (setvar "HPASSOC" old-hpassoc)
  ent)

;; Hatch-based boundary: zoom extents, try once, zoom back
(defun tz-hatch-boundary-v2 (pt gap-tol / ent t1)
  (setq t1 (getvar "MILLISECS"))
  (command "_.ZOOM" "_Extents")
  (princ (strcat "\n[TIME]   zoom extents: " (itoa (- (getvar "MILLISECS") t1)) "ms"))
  (setq t1 (getvar "MILLISECS"))
  (setq ent (tz-try-hatch-boundary pt gap-tol))
  (princ (strcat "\n[TIME]   hatch+cleanup: " (itoa (- (getvar "MILLISECS") t1)) "ms"))
  (setq t1 (getvar "MILLISECS"))
  (command "_.ZOOM" "_Previous")
  (princ (strcat "\n[TIME]   zoom previous: " (itoa (- (getvar "MILLISECS") t1)) "ms"))
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


;; Scan for nearby text around a world-coordinate point using nentselp.
;; Probes a grid of points within radius, reads text/attrib values,
;; and adds unique results to outer-scope nearby-texts list.
;; Works correctly with XREFs (nentselp penetrates them).
(defun tz-scan-nearby-text (center-pt skip-ent skip-str radius
                            / step gx gy probe-pt sel2 ent2 ed2 str2 lyr2
                              pt2 dist2 seen)
  (setq step (/ radius 3.0)  ;; probe every ~12" in a grid
        seen '())
  (setq gx (- radius))
  (while (<= gx radius)
    (setq gy (- radius))
    (while (<= gy radius)
      (setq probe-pt (list (+ (car center-pt) gx)
                           (+ (cadr center-pt) gy)
                           0.0))
      (setq sel2 (nentselp probe-pt))
      (if (and sel2 (listp sel2))
        (progn
          (setq ent2 (car sel2)
                ed2  (entget ent2))
          (if (and (member (cdr (assoc 0 ed2)) '("TEXT" "MTEXT" "ATTRIB"))
                   (not (eq ent2 skip-ent))
                   (not (member (vla-get-handle (vlax-ename->vla-object ent2)) seen)))
            (progn
              (setq str2 (vl-string-trim " " (cdr (assoc 1 ed2)))
                    lyr2 (cdr (assoc 8 ed2)))
              ;; Track by handle so we don't add duplicates
              (setq seen (cons (vla-get-handle (vlax-ename->vla-object ent2)) seen))
              (if (and (/= str2 "")
                       (/= str2 skip-str)
                       (not (tz-skip-text-p str2))
                       (not (wcmatch (strcase lyr2) "T24-*")))
                (progn
                  ;; Use world-coordinate distance from click point
                  (setq pt2 (cadr sel2)  ;; nentselp returns pick point
                        dist2 (distance (list (car center-pt) (cadr center-pt))
                                        (list (car pt2) (cadr pt2))))
                  (setq nearby-texts
                    (cons (list dist2 str2) nearby-texts))))))))
      (setq gy (+ gy step)))
    (setq gx (+ gx step))))

;; Returns T if string looks like a label we should skip (not a room number).
;; Matches: sqft labels, occupancy codes, condition labels, etc.
(defun tz-skip-text-p (s / up)
  (setq up (strcase s))
  (or (wcmatch up "OCC")
      (wcmatch up "OCC.*")
      (wcmatch up "OCCUPANCY")
      (wcmatch up "COND")
      (wcmatch up "CONDITIONED")
      (wcmatch up "UNCONDITIONED")
      (wcmatch up "UNCOND")
      (wcmatch up "PLENUM")
      (wcmatch up "GARAGE")
      (wcmatch up "EXTERIOR")
      (wcmatch up "ATTIC")
      (wcmatch up "EXIT")
      (wcmatch up "EXIT*")
      (wcmatch up "STAIR*")
      (wcmatch up "ELEV*")
      (wcmatch up "CORRIDOR")
      (wcmatch up "HALL*")
      (wcmatch up "MECH*")
      (wcmatch up "ELECTRICAL")
      (wcmatch up "TRASH")
      (wcmatch up "JANITOR")
      (wcmatch up "STORAGE")
      (tz-sqft-text-p s)))

;; Returns T if string looks like a square-footage label (e.g. "250 SQ. FT.")
(defun tz-sqft-text-p (s / up)
  (setq up (strcase s))
  (or (wcmatch up "*SQFT*")
      (wcmatch up "*SQ. FT*")
      (wcmatch up "*SQ FT*")
      (wcmatch up "*S.F.*")
      ;; Number followed by SF/FT unit -- but not bare numbers (those are room numbers)
      (wcmatch up "*#` SF")
      (wcmatch up "*#` FT")))

;; Euclidean distance between two 2D points
(defun tz-cp-seg-len (p1 p2 / dx dy)
  (setq dx (- (car p2) (car p1))
        dy (- (cadr p2) (cadr p1)))
  (sqrt (+ (* dx dx) (* dy dy))))

;; Door-collapse: delete door-notch vertices from a BOUNDARY polyline.
;; Walks both directions from each arc (door-swing) through tiny segments
;; to find the 30-44" door panel, then removes the shorter polygon path
;; between arc and door (the notch).
(defun tz-door-collapse (ent / obj verts n n-orig i k slen slen2
                             new-verts found-door arc-idx arc-pt arc-key
                             best-j best-k mid-pt d
                             inside-end fwd-dist bwd-dist to-remove
                             skipped-arcs k2 j p dbg-num vpt
                             fwd-door-j fwd-door-k bwd-door-j bwd-door-k)
  (setq obj   (vlax-ename->vla-object ent)
        verts (tz-pline-verts obj)
        n     (length verts)
        n-orig n)

  (setq found-door T  skipped-arcs '())
  (while found-door
    (setq found-door nil  n (length verts))

    ;; Find first arc vertex not in skip list
    (setq arc-idx nil  i 0)
    (while (and (null arc-idx) (< i n))
      (if (and (caddr (nth i verts)) (/= (caddr (nth i verts)) 0.0))
        (progn
          (setq arc-key (strcat (rtos (car (nth i verts)) 2 1) ","
                                (rtos (cadr (nth i verts)) 2 1)))
          (if (not (member arc-key skipped-arcs))
            (setq arc-idx i))))
      (setq i (1+ i)))

    (if arc-idx
      (progn
        (setq arc-pt (list (car (nth arc-idx verts))
                           (cadr (nth arc-idx verts))))
        (setq arc-key (strcat (rtos (car arc-pt) 2 1) ","
                              (rtos (cadr arc-pt) 2 1)))

        ;; ── Walk FORWARD from arc through tiny segs (< 10") ──
        (setq fwd-door-j nil  fwd-door-k nil  i 1)
        (while (and (null fwd-door-j) (<= i 8))
          (setq j (rem (+ arc-idx i) n)
                k (rem (+ arc-idx i 1) n)
                slen (tz-cp-seg-len
                       (list (car (nth j verts)) (cadr (nth j verts)))
                       (list (car (nth k verts)) (cadr (nth k verts)))))
          (cond
            ((and (>= slen 30.0) (<= slen 44.0))
             (setq fwd-door-j j  fwd-door-k k))
            ((and (>= slen 10.0) (<= slen 25.0))
             (progn
               (setq k2 (rem (+ arc-idx i 2) n)
                     slen2 (tz-cp-seg-len
                             (list (car (nth k verts)) (cadr (nth k verts)))
                             (list (car (nth k2 verts)) (cadr (nth k2 verts)))))
               (if (and (>= (+ slen slen2) 30.0) (<= (+ slen slen2) 44.0))
                 (setq fwd-door-j j  fwd-door-k k2))
               (setq i 99)))
            ((< slen 10.0)
             (setq i (1+ i)))
            (T (setq i 99))))

        ;; ── Walk BACKWARD from arc through tiny segs (< 10") ──
        ;; Start at i=0 to include the segment ending at the arc vertex
        (setq bwd-door-j nil  bwd-door-k nil  i 0)
        (while (and (null bwd-door-j) (<= i 8))
          (setq j (rem (+ (- arc-idx i 1) n n) n)
                k (rem (+ (- arc-idx i) n n) n)
                slen (tz-cp-seg-len
                       (list (car (nth j verts)) (cadr (nth j verts)))
                       (list (car (nth k verts)) (cadr (nth k verts)))))
          (cond
            ((and (>= slen 30.0) (<= slen 44.0))
             (setq bwd-door-j j  bwd-door-k k))
            ((and (>= slen 10.0) (<= slen 25.0))
             (progn
               (setq k2 (rem (+ (- arc-idx i 2) n n) n)
                     slen2 (tz-cp-seg-len
                             (list (car (nth k2 verts)) (cadr (nth k2 verts)))
                             (list (car (nth j verts)) (cadr (nth j verts)))))
               (if (and (>= (+ slen slen2) 30.0) (<= (+ slen slen2) 44.0))
                 (setq bwd-door-j k2  bwd-door-k k))
               (setq i 99)))
            ((< slen 10.0)
             (setq i (1+ i)))
            (T (setq i 99))))

        (cond
          ;; Only one direction found a door
          ((and fwd-door-j (null bwd-door-j))
           (setq best-j fwd-door-j  best-k fwd-door-k))
          ((and bwd-door-j (null fwd-door-j))
           (setq best-j bwd-door-j  best-k bwd-door-k))

          ;; Both directions found doors — use forward
          ((and fwd-door-j bwd-door-j)
           (setq best-j fwd-door-j  best-k fwd-door-k))

          ;; Neither direction
          (T (setq best-j nil)))

        (if best-j
          (progn
            ;; Which end of door is closer to arc (polygon distance)?
            (setq fwd-dist (rem (+ (- best-j arc-idx) n) n)
                  bwd-dist (rem (+ (- arc-idx best-j) n) n))
            (if (> fwd-dist bwd-dist)
              (setq fwd-dist bwd-dist))
            (setq d fwd-dist)
            (setq fwd-dist (rem (+ (- best-k arc-idx) n) n)
                  bwd-dist (rem (+ (- arc-idx best-k) n) n))
            (if (> fwd-dist bwd-dist)
              (setq fwd-dist bwd-dist))
            (if (<= d fwd-dist)
              (setq inside-end best-j)
              (setq inside-end best-k))

            ;; If inside-end == arc-idx, use the other endpoint
            (if (= inside-end arc-idx)
              (if (= inside-end best-j)
                (setq inside-end best-k)
                (setq inside-end best-j)))

            ;; Shorter polygon path from arc-idx to inside-end
            (setq fwd-dist (rem (+ (- inside-end arc-idx) n) n)
                  bwd-dist (rem (+ (- arc-idx inside-end) n) n))

            ;; Build removal list
            (setq to-remove '())
            (if (<= fwd-dist bwd-dist)
              (progn
                (setq p arc-idx)
                (repeat (1+ fwd-dist)
                  (setq to-remove (cons p to-remove))
                  (setq p (rem (1+ p) n))))
              (progn
                (setq p arc-idx)
                (repeat (1+ bwd-dist)
                  (setq to-remove (cons p to-remove))
                  (setq p (rem (+ (1- p) n) n)))))

            ;; Draw debug circles at each removed vertex
            (setq dbg-num 1)
            (foreach ri to-remove
              (setq vpt (nth ri verts))
              (entmake (list '(0 . "CIRCLE") '(8 . "T24-DBG") '(62 . 1)
                             (cons 10 (list (car vpt) (cadr vpt) 0.0))
                             '(40 . 4.0)))
              (entmake (list '(0 . "TEXT") '(8 . "T24-DBG") '(62 . 1)
                             (cons 10 (list (car vpt) (cadr vpt) 0.0))
                             (cons 1 (itoa dbg-num))
                             '(40 . 3.0) '(72 . 1) '(73 . 2)
                             (cons 11 (list (car vpt) (cadr vpt) 0.0))))
              (setq dbg-num (1+ dbg-num)))

            ;; Remove vertices
            (setq new-verts '()  k 0)
            (repeat n
              (if (not (member k to-remove))
                (setq new-verts (append new-verts (list (nth k verts)))))
              (setq k (1+ k)))
            (setq verts new-verts  found-door T))
          ;; No door -- skip this arc
          (progn
            (setq skipped-arcs (cons arc-key skipped-arcs))
            (setq found-door T))))))

  ;; Write modified vertices back to polyline
  (if (< (length verts) n-orig)
    (tz-set-pline-verts obj verts)))

;; ── Collinear segment merge ──────────────────────────────────────────────────
;; Remove redundant midpoints where consecutive segments are nearly collinear.
;; Angle tolerance: ~1 degree.  Repeats until stable.

(defun tz-seg-angle (p1 p2 / dx dy)
  "Return angle in degrees (0-360) from p1 to p2."
  (setq dx (- (car p2) (car p1))
        dy (- (cadr p2) (cadr p1)))
  (if (and (< (abs dx) 0.001) (< (abs dy) 0.001))
    nil
    (rem (+ (* (/ 180.0 pi) (atan dy dx)) 360.0) 360.0)))

(defun tz-merge-collinear (ent / obj verts n changed i n-orig
                                p-prev p-cur p-next a1 a2 diff new-verts)
  (setq obj   (vlax-ename->vla-object ent)
        verts (tz-pline-verts obj)
        n     (length verts)
        n-orig n
        changed T)
  (while changed
    (setq changed nil  n (length verts)  new-verts '()  i 0)
    (repeat n
      (setq p-prev (nth (rem (+ (1- i) n) n) verts)
            p-cur  (nth i verts)
            p-next (nth (rem (1+ i) n) verts)
            a1 (tz-seg-angle (list (car p-prev) (cadr p-prev))
                             (list (car p-cur) (cadr p-cur)))
            a2 (tz-seg-angle (list (car p-cur) (cadr p-cur))
                             (list (car p-next) (cadr p-next))))
      (if (and a1 a2
               ;; No bulge on this vertex or the previous one
               (or (null (caddr p-cur))  (= (caddr p-cur) 0.0))
               (or (null (caddr p-prev)) (= (caddr p-prev) 0.0))
               ;; Angle difference < 1 degree
               (progn
                 (setq diff (abs (- a1 a2)))
                 (if (> diff 180.0) (setq diff (- 360.0 diff)))
                 (< diff 1.0)))
        (setq changed T)  ;; skip this vertex (collinear midpoint)
        (setq new-verts (append new-verts (list p-cur))))
      (setq i (1+ i)))
    (if changed (setq verts new-verts)))
  (if (< (length verts) n-orig)
    (progn
      (princ (strcat "\n[T24] Merged collinear: " (itoa n-orig) " -> "
                     (itoa (length verts)) " verts"))
      (tz-set-pline-verts obj verts))))

;; ── TZ-ZONE ──────────────────────────────────────────────────────────────────
(defun c:TZ-ZONE ( / *error* sel sel2 txt-ent txt-edata txt-str txt-lyr txt-pt
                     ent last-ent pts area-ft centroid
                     zone-id zone-name ceil-ht floor condition occupancy
                     ce cd cl edata froze txt-lyrs-frozen gap-tol hpg-save
                     txt-layers choice ldata ed2
                     patch-lines pp1 pp2
                     ss-near j near-ent near-ed near-str near-pt near-dist
                     nearby-texts txt-ins blk-ref blk-name blk-def blk-ref-lyr
                     sub-ent sub-ed near-lyr)
  ;; Allow (command ...) calls inside *error* handler (AutoCAD 2015+ requirement)
  (*push-error-using-command*)
  (setq *TZ-BUSY* T)

  (defun *error* (msg)
    (setq *TZ-BUSY* nil)
    (if ce (setvar "CMDECHO" ce))
    (if cd (setvar "CMDDIA"  cd))
    (if cl (setvar "CLAYER"  cl))
    (if froze (tz-thaw-layers froze))
    (if txt-lyrs-frozen (tz-thaw-layers txt-lyrs-frozen))
    (if hpg-save (setvar "HPGAPTOL" hpg-save) (setvar "HPGAPTOL" 0.0))
    (*pop-error-using-command*)
    (if (not (member msg '("Function cancelled" "quit / exit abort" "")))
      (princ (strcat "\n[T24] Error: " msg)))
    (princ))

  (tz-setup)
  (setq ce nil  cd nil  cl nil  froze nil  txt-lyrs-frozen nil
        hpg-save (getvar "HPGAPTOL"))


  ;; ── Session setup dialog ────────────────────────────────────────────────────
  ;; Globals persist between runs; dialog pre-fills with previous values
  (setq ceil-ht 9.0  floor 1  gap-tol 1.0)
  (tz-session-dialog)  ;; populates ceil-ht, floor, gap-tol from persistent globals

  ;; Condition and occupancy default to generic values — sort out in EnergyPro
  (setq condition "Conditioned"
        occupancy "Nonresidential")

  (princ (strcat "\n[T24] Session: " (rtos ceil-ht 2 1) "' ceiling  |  Floor " (itoa floor)
                 "  |  Gap tol " (rtos gap-tol 2 1) "\""
                 "  |  CZ " *TZ-CLIMATE-ZONE*
                 "  |  North " (rtos *TZ-NORTH-ANGLE* 2 1) (chr 176)))
  (princ "\n[T24] Now click room name text for each zone. Press Enter when done.")

  ;; ── Main room loop ────────────────────────────────────────────────────────
  (while
    (progn
      (setq zone-name nil  txt-lyr nil  txt-layers '()  txt-edata nil)
      ;; Retry loop — keep clicking until text is found or Enter pressed
      (setq sel T)
      (while (and sel (null zone-name))
        (initget "Undo")
        (setq sel (nentsel "\n[T24] Click room name text [Undo] (Enter = manual entry): "))
        (if (= sel "Undo")
          (progn
            (command "_.UNDO" "")
            (princ "\n[T24] Last zone undone.")
            (setq sel T))
        (if sel
          (progn
            (setq txt-ent   (car sel)
                  txt-pt    (cadr sel)
                  txt-edata (entget txt-ent))
            (if (member (cdr (assoc 0 txt-edata)) '("TEXT" "MTEXT"))
              (progn
                (setq txt-str (vl-string-trim " " (cdr (assoc 1 txt-edata)))
                      txt-lyr (cdr (assoc 8 txt-edata)))
                (cond
                  ((and txt-lyr (wcmatch txt-lyr "T24-*"))
                   (princ "\n[T24] Clicked a T24 entity -- try again."))
                  ((and txt-lyr (not (tz-layer-visible-p txt-lyr)))
                   (princ (strcat "\n[T24] Text on frozen/off layer \"" txt-lyr "\" -- try again.")))
                  (T
                   (if txt-lyr (setq txt-layers (list txt-lyr)))
                   (if (= txt-str "") (setq txt-str nil))
                   (setq zone-name txt-str))))
              (princ "\n[T24] Not text -- try again, or press Enter for manual entry."))))))
      ;; If Enter was pressed (sel = nil), offer manual point + typed name
      (cond
        (zone-name T)
        (T
         (setq txt-pt (getpoint "\n[T24] Click inside the room (Enter to finish all): "))
         (if txt-pt
           (progn
             (setq zone-name
               (getstring T "\n[T24] Type zone name: "))
             (if (= zone-name "") (setq zone-name nil))
             T)
           nil))))

    (if (null zone-name)
      (princ "\n[T24] No name entered, skipping.")

      (progn
        ;; Start undo group so "Undo" keyword can revert this entire zone
        (command "_.UNDO" "_Begin")
        ;; ── Auto-find nearby text to build full name ────────────────────
        (setq *tz-t0* (getvar "MILLISECS"))
        (if txt-pt
          (progn
            (setq nearby-texts '())
            (tz-scan-nearby-text txt-pt txt-ent zone-name 36.0)
            ;; Sort by distance, append only the closest one
            (if nearby-texts
              (progn
                (setq nearby-texts
                  (vl-sort nearby-texts '(lambda (a b) (< (car a) (car b)))))
                (princ (strcat "\n[T24]   Found " (itoa (length nearby-texts)) " nearby text(s):"))
                (foreach nt nearby-texts
                  (princ (strcat "\n[T24]     " (rtos (car nt) 2 1) "\" \"" (cadr nt) "\"")))
                (setq zone-name (strcat zone-name " " (cadr (car nearby-texts))))
                (princ (strcat "\n[T24]   Auto-appended: \"" (cadr (car nearby-texts)) "\""))))))
        ;; Hard cap at 30 characters
        (if (> (strlen zone-name) 30)
          (setq zone-name (substr zone-name 1 30)))

        (princ (strcat "\n[T24] Zone name: \"" zone-name "\""))
        (princ (strcat "\n[TIME] Text search: " (itoa (- (getvar "MILLISECS") *tz-t0*)) "ms"))
        (princ (strcat "\n[T24] Click point: " (rtos (car txt-pt) 2 2) ", " (rtos (cadr txt-pt) 2 2)))

        ;; ── Freeze text layers + door layers before BOUNDARY ──────────────
        (setq *tz-t0* (getvar "MILLISECS"))
        (setq ce (getvar "CMDECHO")
              cd (getvar "CMDDIA")
              cl (getvar "CLAYER"))
        (setvar "CMDECHO" 0)
        (setvar "CMDDIA"  0)
        (setvar "CLAYER"  *TZ-LYR-ZONE*)

        (if txt-lyr
          (progn
            (tz-freeze-layer txt-lyr)
            (if (not (member txt-lyr txt-lyrs-frozen))
              (setq txt-lyrs-frozen (cons txt-lyr txt-lyrs-frozen)))))
        (princ (strcat "\n[TIME] Layer setup: " (itoa (- (getvar "MILLISECS") *tz-t0*)) "ms"))

        ;; ── Run boundary ──
        (princ "\n[T24] Running HATCH boundary...")
        (setq *tz-t0* (getvar "MILLISECS"))
        (setq ent (tz-hatch-boundary-v2 txt-pt gap-tol))
        (princ (strcat "\n[TIME] BOUNDARY: " (itoa (- (getvar "MILLISECS") *tz-t0*)) "ms"))
        (if ent
          (princ "\n[T24] Boundary created.")
          (princ "\n[T24] Boundary returned nil."))

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
                  (setq ent (tz-hatch-boundary-v2 txt-pt gap-tol)))
                (princ "\n[T24]   No lines drawn.")))

            ;; ── Retry mode: pick a new point, try full gap tolerance range ──
            (progn
              (setq txt-pt (getpoint "\n[T24]   Click inside the room: "))
              (if txt-pt
                (setq ent (tz-hatch-boundary-v2 txt-pt gap-tol))
                (princ "\n[T24]   No point picked."))))

          ) ;; end while

        ;; Clean up temporary lines after boundary succeeds or user moves on
        (foreach pent patch-lines (entdel pent))
        (if patch-lines
          (princ (strcat "\n[T24]   Cleaned up " (itoa (length patch-lines)) " patch line(s).")))

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
            ((= choice "Rect")
             (setq ent (tz-pick-rectangle)))
            (T  ;; "Manual" — pick corners
             (setq ent (tz-pick-corners)))))

        (if ent
          (progn
            ;; Collapse door notches (arc + tiny segs + door panel)
            (setq *tz-t0* (getvar "MILLISECS"))
            (tz-door-collapse ent)
            (princ (strcat "\n[TIME] Door collapse: " (itoa (- (getvar "MILLISECS") *tz-t0*)) "ms"))

            ;; Merge collinear segments (BOUNDARY splits straight walls)
            (setq *tz-t0* (getvar "MILLISECS"))
            (tz-merge-collinear ent)
            (princ (strcat "\n[TIME] Merge collinear: " (itoa (- (getvar "MILLISECS") *tz-t0*)) "ms"))

            ;; Shape simplification (set by TZ-ZONE1/2/3, nil for normal TZ-ZONE)
            (if *TZ-SHAPE-MODE*
              (progn
                (setq *tz-t0* (getvar "MILLISECS"))
                (tz-shape-simplify ent *TZ-SHAPE-MODE*)
                (princ (strcat "\n[TIME] Shape simplify: " (itoa (- (getvar "MILLISECS") *tz-t0*)) "ms"))))

            ;; Finalize: layer, area, xdata, label, reactor
            (setq *tz-t0* (getvar "MILLISECS"))
            ;; Move to T24-ZONE layer if BOUNDARY put it elsewhere
            (setq edata (entget ent))
            (if (/= (cdr (assoc 8 edata)) *TZ-LYR-ZONE*)
              (entmod (subst (cons 8 *TZ-LYR-ZONE*) (assoc 8 edata) edata)))

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
            (tz-zone-label txt-pt zone-name area-ft ceil-ht floor zone-id)

            ;; Attach object reactor for dynamic area updates
            (setq *TZ-REACTORS*
              (cons (vlr-object-reactor
                      (list (vlax-ename->vla-object ent))
                      "T24-Zone-Update"
                      '((:vlr-modified . tz-zone-modified-callback)))
                    (if *TZ-REACTORS* *TZ-REACTORS* '())))

            (princ (strcat "\n[TIME] Finalize: " (itoa (- (getvar "MILLISECS") *tz-t0*)) "ms"))
            (princ (strcat "\n[T24] Tagged: \"" zone-name
                           "\"  " (rtos area-ft 2 1) " sqft  Floor " (itoa floor))))
          (princ "\n[T24] No boundary created, skipping this zone."))
        ;; Thaw the text layer per-room (whether boundary succeeded or not)
        (setq *tz-t0* (getvar "MILLISECS"))
        (if txt-lyr (tz-thaw-layer txt-lyr))
        (command "_.UNDO" "_End")
        (princ (strcat "\n[TIME] Thaw+undo: " (itoa (- (getvar "MILLISECS") *tz-t0*)) "ms"))
            ))) ; end progn(zone-name), if zone-name, while

  ;; Thaw any layers frozen during the session (e.g. txt-lyrs-frozen, froze)
  (if txt-lyrs-frozen (tz-thaw-layers txt-lyrs-frozen))
  (if froze (tz-thaw-layers froze))
  (setq *TZ-BUSY* nil)
  (*pop-error-using-command*)
  (c:TZ-WATCH)
  (tz-bring-to-front)
  (princ "\n[T24] Done tagging zones.")
  (princ))

;; ── TZ-WIN ───────────────────────────────────────────────────────────────────
;; Two modes: Area (enter sqft + click) or Points (click two points × height).
(defun c:TZ-WIN ( / pt pt1 pt2 ent open-id area-ft lsave mode
                     edge-info zone-id wall-id win-ht
                     dx dy len-in len-ft mid-pt)
  (tz-setup)
  (setq lsave (getvar "CLAYER"))
  (defun *error* (msg)
    (if lsave (setvar "CLAYER" lsave))
    (if (not (member msg '("Function cancelled" "quit / exit abort" "")))
      (princ (strcat "\n[T24] Error: " msg)))
    (princ))

  (initget "Area Points")
  (setq mode (getkword "\n[T24] Enter A to type area, P to pick two points <A>: "))
  (if (null mode) (setq mode "Area"))

  (if (= mode "Points")
    ;; ── Points mode: two clicks × height ──────────────────────────────────
    (progn
      (setq win-ht (getreal "\n[T24] Window height (ft) <5>: "))
      (if (null win-ht) (setq win-ht 5.0))
      (princ (strcat "\n[T24] Click two points per window, "
                     (rtos win-ht 2 1) "' height. Enter when done."))
      (while
        (progn
          (setq pt1 (getpoint "\n[T24] Window start point (Enter to finish): "))
          (not (null pt1)))
        (setq pt2 (getpoint pt1 "\n[T24] Window end point: "))
        (if pt2
          (progn
            (setq dx     (- (car pt2) (car pt1))
                  dy     (- (cadr pt2) (cadr pt1))
                  len-in (sqrt (+ (* dx dx) (* dy dy)))
                  len-ft (/ len-in *TZ-UNIT-FT*)
                  area-ft (* len-ft win-ht)
                  mid-pt (list (/ (+ (car pt1) (car pt2)) 2.0)
                               (/ (+ (cadr pt1) (cadr pt2)) 2.0) 0.0))
            (setq open-id (tz-next-open-id))
            ;; Auto-detect zone and wall
            (setq edge-info (tz-nearest-edge mid-pt)
                  zone-id nil  wall-id nil)
            (if edge-info
              (setq zone-id (tz-zone-of-ent (nth 0 edge-info))))
            (if (null zone-id) (setq zone-id "Z-???"))
            (setq wall-id (tz-nearest-wall-id mid-pt zone-id))
            (if (null wall-id) (setq wall-id "W-???"))
            ;; Place marker at midpoint
            (setvar "CLAYER" *TZ-LYR-WIN*)
            (setq ent (tz-make-circle mid-pt *TZ-MARKER-RAD*))
            (setvar "CLAYER" lsave)
            (if ent
              (progn
                (tz-set-xdata ent
                  (list *TZ-APP*
                    (cons 1000 "OPENING")
                    (cons 1000 open-id)
                    (cons 1000 "Window")
                    (cons 1000 zone-id)
                    (cons 1000 wall-id)
                    (cons 1040 area-ft)
                    (cons 1040 0.0)
                    (cons 1040 0.0)))
                (tz-open-label mid-pt open-id "Window" area-ft 0.0 0.0)
                (princ (strcat "\n[T24] " open-id ": " (rtos len-ft 2 1)
                               "' x " (rtos win-ht 2 1) "'h = "
                               (rtos area-ft 2 1) " sqft → " wall-id)))
              (princ "\n[T24] Skipping — marker failed.")))
          (princ "\n[T24] No end point, skipped."))))

    ;; ── Area mode: enter sqft + click location ────────────────────────────
    (progn
      (setq area-ft (getreal "\n[T24] Window area (sqft) <30>: "))

      (if (null area-ft) (setq area-ft 30.0))
      (princ (strcat "\n[T24] Placing " (rtos area-ft 2 1)
                     " sqft windows. Click locations, Enter when done."))
      (while
        (progn
          (setq pt (getpoint "\n[T24] Click on zone boundary for window: "))
          (not (null pt)))
        (setq pt (tz-snap-to-zone pt))
        (setq open-id (tz-next-open-id))
        ;; Auto-detect zone and wall
        (setq edge-info (tz-nearest-edge pt)
              zone-id nil  wall-id nil)
        (if edge-info
          (setq zone-id (tz-zone-of-ent (nth 0 edge-info))))
        (if (null zone-id) (setq zone-id "Z-???"))
        (setq wall-id (tz-nearest-wall-id pt zone-id))
        (if (null wall-id) (setq wall-id "W-???"))
        (setvar "CLAYER" *TZ-LYR-WIN*)
        (setq ent (tz-make-circle pt *TZ-MARKER-RAD*))
        (setvar "CLAYER" lsave)
        (if ent
          (progn
            (tz-set-xdata ent
              (list *TZ-APP*
                (cons 1000 "OPENING")
                (cons 1000 open-id)
                (cons 1000 "Window")
                (cons 1000 zone-id)
                (cons 1000 wall-id)
                (cons 1040 area-ft)
                (cons 1040 0.0)
                (cons 1040 0.0)))
            (tz-open-label pt open-id "Window" area-ft 0.0 0.0)
            (princ (strcat "\n[T24] " open-id " placed → " wall-id)))
          (princ "\n[T24] Skipping — marker failed.")))))
  (tz-bring-to-front)
  (princ "\n[T24] Done placing windows.")
  (princ))

;; ── TZ-DOOR ──────────────────────────────────────────────────────────────────
(defun c:TZ-DOOR ( / pt ent open-id area-ft lsave edge-info zone-id wall-id)
  (tz-setup)
  (setq lsave (getvar "CLAYER"))
  (defun *error* (msg)
    (if lsave (setvar "CLAYER" lsave))
    (if (not (member msg '("Function cancelled" "quit / exit abort" "")))
      (princ (strcat "\n[T24] Error: " msg)))
    (princ))
  (setq area-ft (getreal "\n[T24] Door area (sqft) <21>: "))
  (if (null area-ft) (setq area-ft 21.0))
  (princ (strcat "\n[T24] Placing " (rtos area-ft 2 1)
                 " sqft doors. Click locations, Enter when done."))
  (while
    (progn
      (setq pt (getpoint "\n[T24] Click on zone boundary for door: "))
      (not (null pt)))
    (setq pt (tz-snap-to-zone pt))
    (setq open-id (tz-next-open-id))
    ;; Auto-detect zone and wall
    (setq edge-info (tz-nearest-edge pt)
          zone-id nil  wall-id nil)
    (if edge-info
      (setq zone-id (tz-zone-of-ent (nth 0 edge-info))))
    (if (null zone-id) (setq zone-id "Z-???"))
    (setq wall-id (tz-nearest-wall-id pt zone-id))
    (if (null wall-id) (setq wall-id "W-???"))
    (setvar "CLAYER" *TZ-LYR-DOOR*)
    (setq ent (tz-make-circle pt *TZ-MARKER-RAD*))
    (setvar "CLAYER" lsave)
    (if ent
      (progn
        (tz-set-xdata ent
          (list *TZ-APP*
            (cons 1000 "OPENING")
            (cons 1000 open-id)
            (cons 1000 "Door")
            (cons 1000 zone-id)
            (cons 1000 wall-id)
            (cons 1040 area-ft)
            (cons 1040 0.0)
            (cons 1040 0.0)))
        (tz-open-label pt open-id "Door" area-ft 0.0 0.0)
        (princ (strcat "\n[T24] " open-id " placed → " wall-id)))
      (princ "\n[T24] Skipping — marker failed.")))
  (tz-bring-to-front)
  (princ "\n[T24] Done placing doors.")
  (princ))

;; ── TZ-SKY ───────────────────────────────────────────────────────────────────
(defun c:TZ-SKY ( / pt ent open-id area-ft uf shgc lsave r edge-info zone-id)
  (tz-setup)
  (setq lsave (getvar "CLAYER"))
  (defun *error* (msg)
    (if lsave (setvar "CLAYER" lsave))
    (if (not (member msg '("Function cancelled" "quit / exit abort" "")))
      (princ (strcat "\n[T24] Error: " msg)))
    (princ))
  (setq area-ft (getreal "\n[T24] Skylight area (sqft) <12>: "))
  (if (null area-ft) (setq area-ft 12.0))
  (setq uf 0.50 shgc 0.25)
  (princ (strcat "\n[T24] Placing " (rtos area-ft 2 1)
                 " sqft skylights. Click locations, Enter when done."))
  (setq r *TZ-MARKER-RAD*)
  (while
    (progn
      (setq pt (getpoint "\n[T24] Click inside zone for skylight: "))
      (not (null pt)))
    (setq open-id (tz-next-open-id))
    ;; Auto-detect zone (skylights go to roof, no wall-id needed)
    (setq edge-info (tz-nearest-edge pt)
          zone-id nil)
    (if edge-info
      (setq zone-id (tz-zone-of-ent (nth 0 edge-info))))
    (if (null zone-id) (setq zone-id "Z-???"))
    (setvar "CLAYER" *TZ-LYR-SKY*)
    ;; Circle + X cross to distinguish from windows
    (setq ent (tz-make-circle pt r))
    (if ent
      (progn
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
            (cons 1000 zone-id)
            (cons 1000 (strcat zone-id "-ROOF"))
            (cons 1040 area-ft)
            (cons 1040 uf)
            (cons 1040 shgc)))
        (tz-open-label pt open-id "Skylight" area-ft uf shgc)
        (princ (strcat "\n[T24] " open-id " placed → " zone-id "-ROOF")))
      (progn
        (setvar "CLAYER" lsave)
        (princ "\n[T24] Skipping — marker failed."))))
  (tz-bring-to-front)
  (princ "\n[T24] Done placing skylights.")
  (princ))

;; ── Wall helpers ──────────────────────────────────────────────────────────────

;; Session counter for wall IDs
(defun tz-next-wall-id ( / ss)
  (if (null *TZ-WALL-COUNT*)
    (progn
      (setq ss (ssget "X" (list (cons 8 *TZ-LYR-WALL*) '(0 . "CIRCLE"))))
      (setq *TZ-WALL-COUNT* (if ss (sslength ss) 0))))
  (setq *TZ-WALL-COUNT* (1+ *TZ-WALL-COUNT*))
  (strcat "W-" (tz-pad (itoa *TZ-WALL-COUNT*) 3)))

;; Snap a point onto the nearest T24-ZONE polyline edge.
;; Returns the projected point, or the original pt if no zones exist.
(defun tz-snap-to-zone (pt / info p1 p2 dx dy t0 cx cy)
  (setq info (tz-nearest-edge pt))
  (if (null info) pt
    (progn
      (setq p1 (nth 1 info)  p2 (nth 2 info)
            dx (- (car p2) (car p1))
            dy (- (cadr p2) (cadr p1))
            t0 (/ (+ (* (- (car pt) (car p1)) dx)
                     (* (- (cadr pt) (cadr p1)) dy))
                  (+ (* dx dx) (* dy dy))))
      (if (< t0 0.0) (setq t0 0.0))
      (if (> t0 1.0) (setq t0 1.0))
      (list (+ (car p1) (* t0 dx))
            (+ (cadr p1) (* t0 dy))))))

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

;; Find nearest wall marker to a point.  Returns wall-id string or nil.
;; Find nearest wall marker to a point, optionally filtered to a specific zone.
;; If host-zone is provided, only walls belonging to that zone are considered.
(defun tz-nearest-wall-id (pt host-zone / ss i went wpt wxd wzid d best-id best-dist)
  (setq ss (ssget "X" (list (cons 8 *TZ-LYR-WALL*) '(0 . "CIRCLE"))))
  (if ss
    (progn
      (setq i 0  best-dist 1e20  best-id nil)
      (repeat (sslength ss)
        (setq went (ssname ss i)
              wpt  (cdr (assoc 10 (entget went)))
              wxd  (tz-get-xdata went))
        (if (and wxd (= (tz-xd-nth wxd 1000 0) "WALL"))
          (progn
            (setq wzid (tz-xd-nth wxd 1000 3))
            ;; Only match walls in the same zone (or all if no filter)
            (if (or (null host-zone) (= wzid host-zone))
              (progn
                (setq d (distance (list (car pt) (cadr pt))
                                  (list (car wpt) (cadr wpt))))
                (if (< d best-dist)
                  (setq best-dist d
                        best-id (tz-xd-nth wxd 1000 1)))))))
        (setq i (1+ i)))
      best-id)))

;; Compute wall azimuth: right-perpendicular of p1→p2, CW from true North.
;; Applies *TZ-NORTH-ANGLE* correction: if the north arrow in the drawing
;; is rotated, raw AutoCAD angles are offset from true north.
;; true_azimuth = (raw_azimuth - north_arrow_angle) mod 360
(defun tz-wall-azimuth (p1 p2 / dx dy len nx ny az north-adj)
  (setq dx (- (car p2) (car p1))
        dy (- (cadr p2) (cadr p1))
        len (sqrt (+ (* dx dx) (* dy dy))))
  (if (< len 1e-6) 0.0
    (progn
      ;; Right perpendicular: (dy, -dx) normalized
      (setq nx (/ dy len)
            ny (/ (- dx) len)
            az (* (/ 180.0 pi) (atan nx ny)))
      (if (< az 0.0) (setq az (+ az 360.0)))
      ;; Correct for north arrow rotation
      (setq north-adj (if *TZ-NORTH-ANGLE* *TZ-NORTH-ANGLE* 0.0))
      (setq az (rem (+ (- az north-adj) 360.0) 360.0))
      az)))

;; Read zone-id from a polyline's XDATA
(defun tz-zone-of-ent (ent / xd)
  (setq xd (tz-get-xdata ent))
  (if (and xd (equal (tz-xd-nth xd 1000 0) "ZONE"))
    (tz-xd-nth xd 1000 1)
    nil))

;; Place label text at click point for a wall marker
(defun tz-wall-label (pt wall-id wall-type len-ft wall-ht area-sqft az label-ht
                      / lsave th off line2-off)
  (setq lsave    (getvar "CLAYER")
        th       label-ht
        off      (+ *TZ-WALL-RAD* (* 1.5 th))
        line2-off (+ off (* 1.4 th)))
  (setvar "CLAYER" *TZ-LYR-LABEL*)
  ;; Line 1: ID and type
  (tz-make-text
    (list (car pt) (+ (cadr pt) line2-off) 0.0) th
    (strcat wall-id "  " wall-type))
  ;; Line 2: dimensions and area
  (tz-make-text
    (list (car pt) (+ (cadr pt) off) 0.0) th
    (strcat (rtos len-ft 2 1) "' x " (rtos wall-ht 2 1) "'h  "
            (rtos area-sqft 2 1) " sqft  "
            (rtos az 2 0) (chr 176)))
  (setvar "CLAYER" lsave))

;; ── TZ-WALL ──────────────────────────────────────────────────────────────────
;; Two-point wall: click two points to define a wall span on/near a zone.
;; Length = distance between the two points.  Azimuth from the wall line.
;; Marker placed at midpoint.
(defun c:TZ-WALL ( / pt1 pt2 edge-info
                     wall-id wall-type wall-ht zone-id label-ht
                     len-ft area-sqft az lsave ent type-choice
                     dx dy len-in mid-pt count)
  (tz-setup)
  (setq lsave (getvar "CLAYER"))
  (defun *error* (msg)
    (if lsave (setvar "CLAYER" lsave))
    (redraw)
    (if (not (member msg '("Function cancelled" "quit / exit abort" "")))
      (princ (strcat "\n[T24] Error: " msg)))
    (princ))

  ;; ── Session defaults ────────────────────────────────────────────────────
  (setq wall-ht (getreal "\n[T24] Wall height (ft) <9>: "))
  (if (null wall-ht) (setq wall-ht 9.0))

  (setq label-ht 6.0)

  (initget "Exterior Interior")
  (setq type-choice (getkword "\n[T24] E = Exterior, I = Interior <E>: "))
  (if (or (null type-choice) (= type-choice "Exterior"))
    (setq wall-type "Exterior Wall")
    (setq wall-type "Interior Wall"))

  (princ (strcat "\n[T24] " wall-type ", " (rtos wall-ht 2 1) "' height."))
  (princ "\n[T24] Click two points per wall. U to undo. Enter when done.")

  ;; ── Click loop — two points per wall ──────────────────────────────────
  (setq count 0)
  (while
    (progn
      (initget "Undo")
      (setq pt1 (getpoint "\n[T24] Wall start point [Undo] (Enter to finish): "))
      (cond
        ((= pt1 "Undo")
         (if (> count 0)
           (progn
             (command "_.UNDO" "")
             (setq count (1- count))
             (princ "\n[T24] Last wall undone."))
           (princ "\n[T24] Nothing to undo."))
         T)
        ((null pt1) nil)
        (T T)))

    (if (and pt1 (listp pt1))
      (progn
        ;; Rubber-band line from pt1 to second click
        (setq pt2 (getpoint pt1 "\n[T24] Wall end point: "))
        (if pt2
          (progn
            ;; Compute wall properties from the two points
            (setq dx     (- (car pt2)  (car pt1))
                  dy     (- (cadr pt2) (cadr pt1))
                  len-in (sqrt (+ (* dx dx) (* dy dy)))
                  len-ft (/ len-in *TZ-UNIT-FT*)
                  az     (tz-wall-azimuth pt1 pt2)
                  mid-pt (list (/ (+ (car pt1) (car pt2)) 2.0)
                               (/ (+ (cadr pt1) (cadr pt2)) 2.0) 0.0))

            ;; Find nearest zone for zone-id
            (setq edge-info (tz-nearest-edge mid-pt)
                  zone-id "Z-???")
            (if edge-info
              (setq zone-id (tz-zone-of-ent (nth 0 edge-info))))
            (if (null zone-id) (setq zone-id "Z-???"))

            (setq wall-id   (tz-next-wall-id)
                  area-sqft  (* len-ft wall-ht))

            ;; ── Undo group ──
            (command "_.UNDO" "_Begin")

            ;; Place marker circle at midpoint
            (setvar "CLAYER" *TZ-LYR-WALL*)
            (setq ent (tz-make-circle mid-pt *TZ-WALL-RAD*))
            (setvar "CLAYER" lsave)

            (if ent
              (progn
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
                    (cons 1040 (car  pt1))
                    (cons 1040 (cadr pt1))
                    (cons 1040 (car  pt2))
                    (cons 1040 (cadr pt2))))
                ;; Visual line on T24-LABEL layer
                (setvar "CLAYER" *TZ-LYR-LABEL*)
                (entmake (list '(0 . "LINE") '(100 . "AcDbEntity") '(100 . "AcDbLine")
                           (cons 10 (list (car pt1) (cadr pt1) 0.0))
                           (cons 11 (list (car pt2) (cadr pt2) 0.0))))
                (setvar "CLAYER" lsave)
                ;; Label
                (tz-wall-label mid-pt wall-id wall-type len-ft wall-ht area-sqft az label-ht)
                (setq count (1+ count))
                (princ (strcat "\n[T24] " wall-id ": " (rtos len-ft 2 1) "' "
                               wall-type " in " zone-id
                               " (" (rtos az 2 0) (chr 176) ")")))
              (princ "\n[T24] Failed to create wall marker."))

            (command "_.UNDO" "_End"))
          (princ "\n[T24] No end point, skipped.")))))

  (tz-bring-to-front)
  (princ (strcat "\n[T24] Done — placed " (itoa count) " walls."))
  (princ))

;; ── TZ-WSIDE ─────────────────────────────────────────────────────────────────
;; Bounding-box wall side: single click outside any zone to auto-detect the
;; nearest zone and which side (N/S/E/W).  Uses full projected dimension.
;; One continuous loop across all zones — click-click-click, Enter when done.
;; Shows preview line before placing; each wall is individually undoable.
(defun c:TZ-WSIDE ( / edge-info ent pts zone-id centroid
                      wall-ht label-ht type-choice wall-type
                      min-x max-x min-y max-y
                      click-pt dx dy side
                      len-ft az p1 p2 mid-pt
                      wall-id area-sqft marker lsave
                      dim-off lbl-off count
                      pv1 pv2 gr gr-type gr-data
                      loop-on prev-side prev-ent)
  (tz-setup)
  (setq lsave (getvar "CLAYER"))
  (defun *error* (msg)
    (if lsave (setvar "CLAYER" lsave))
    (redraw)  ; clear any grdraw previews
    (if (not (member msg '("Function cancelled" "quit / exit abort" "")))
      (princ (strcat "\n[T24] Error: " msg)))
    (princ))

  ;; ── Session defaults ───────────────────────────────────────────────────────
  (setq wall-ht (getreal "\n[T24] Wall height (ft) <9>: "))
  (if (null wall-ht) (setq wall-ht 9.0))

  (setq label-ht 6.0)

  (initget "Exterior Interior")
  (setq type-choice (getkword "\n[T24] E = Exterior, I = Interior <E>: "))
  (if (or (null type-choice) (= type-choice "Exterior"))
    (setq wall-type "Exterior Wall")
    (setq wall-type "Interior Wall"))

  (princ (strcat "\n[T24] " wall-type ", " (rtos wall-ht 2 1) "' height."))
  (princ "\n[T24] Move mouse near zone sides — preview shows which wall.")
  (princ "\n[T24] Click to place, U to undo last, Enter/Esc to finish.")

  ;; ── grread loop — live preview as mouse moves, place on click ─────────────
  (setq count 0  loop-on T  prev-side nil  prev-ent nil
        dim-off (* 3.0 *TZ-UNIT-FT*))  ; 3 feet offset

  (while loop-on
    (setq gr (grread T 4 0))  ; track mouse, no cursor change
    (setq gr-type (car gr)
          gr-data (cadr gr))

    (cond
      ;; ── Mouse move (type 5) — update preview ──
      ((= gr-type 5)
       (setq edge-info (tz-nearest-edge gr-data))
       (if edge-info
         (progn
           (setq ent (nth 0 edge-info)
                 pts (tz-get-pts ent)
                 centroid (tz-centroid pts))
           (setq min-x (apply 'min (mapcar 'car pts))
                 max-x (apply 'max (mapcar 'car pts))
                 min-y (apply 'min (mapcar 'cadr pts))
                 max-y (apply 'max (mapcar 'cadr pts)))
           (setq dx (- (car gr-data)  (car centroid))
                 dy (- (cadr gr-data) (cadr centroid)))
           (if (> (abs dx) (abs dy))
             (if (> dx 0) (setq side "E") (setq side "W"))
             (if (> dy 0) (setq side "N") (setq side "S")))
           ;; Only redraw preview if side or zone changed
           (if (or (/= side prev-side) (not (equal ent prev-ent)))
             (progn
               (redraw)  ; clear old preview
               (cond
                 ((= side "N")
                  (setq pv1 (list min-x (+ max-y dim-off) 0.0)
                        pv2 (list max-x (+ max-y dim-off) 0.0)))
                 ((= side "S")
                  (setq pv1 (list min-x (- min-y dim-off) 0.0)
                        pv2 (list max-x (- min-y dim-off) 0.0)))
                 ((= side "E")
                  (setq pv1 (list (+ max-x dim-off) min-y 0.0)
                        pv2 (list (+ max-x dim-off) max-y 0.0)))
                 ((= side "W")
                  (setq pv1 (list (- min-x dim-off) min-y 0.0)
                        pv2 (list (- min-x dim-off) max-y 0.0))))
               (grdraw pv1 pv2 2 1)  ; yellow preview line
               (setq prev-side side  prev-ent ent))))))

      ;; ── Left click (type 3) — place the wall ──
      ((= gr-type 3)
       (setq click-pt gr-data)
       (redraw)  ; clear preview
       (setq edge-info (tz-nearest-edge click-pt))
       (if (null edge-info)
         (princ "\n[T24] No T24-ZONE polylines found.")
         (progn
           (setq ent (nth 0 edge-info)
                 pts (tz-get-pts ent)
                 zone-id (tz-zone-of-ent ent)
                 centroid (tz-centroid pts))
           (if (null zone-id) (setq zone-id "Z-???"))
           (setq min-x (apply 'min (mapcar 'car pts))
                 max-x (apply 'max (mapcar 'car pts))
                 min-y (apply 'min (mapcar 'cadr pts))
                 max-y (apply 'max (mapcar 'cadr pts)))
           (setq dx (- (car click-pt) (car centroid))
                 dy (- (cadr click-pt) (cadr centroid)))
           (if (> (abs dx) (abs dy))
             (if (> dx 0) (setq side "E") (setq side "W"))
             (if (> dy 0) (setq side "N") (setq side "S")))
           (cond
             ((= side "N")
              (setq len-ft (/ (- max-x min-x) *TZ-UNIT-FT*)  az 0.0
                    p1 (list min-x max-y)  p2 (list max-x max-y)
                    mid-pt (list (/ (+ min-x max-x) 2.0) max-y 0.0)
                    pv1 (list min-x (+ max-y dim-off) 0.0)
                    pv2 (list max-x (+ max-y dim-off) 0.0)))
             ((= side "S")
              (setq len-ft (/ (- max-x min-x) *TZ-UNIT-FT*)  az 180.0
                    p1 (list min-x min-y)  p2 (list max-x min-y)
                    mid-pt (list (/ (+ min-x max-x) 2.0) min-y 0.0)
                    pv1 (list min-x (- min-y dim-off) 0.0)
                    pv2 (list max-x (- min-y dim-off) 0.0)))
             ((= side "E")
              (setq len-ft (/ (- max-y min-y) *TZ-UNIT-FT*)  az 90.0
                    p1 (list max-x min-y)  p2 (list max-x max-y)
                    mid-pt (list max-x (/ (+ min-y max-y) 2.0) 0.0)
                    pv1 (list (+ max-x dim-off) min-y 0.0)
                    pv2 (list (+ max-x dim-off) max-y 0.0)))
             ((= side "W")
              (setq len-ft (/ (- max-y min-y) *TZ-UNIT-FT*)  az 270.0
                    p1 (list min-x min-y)  p2 (list min-x max-y)
                    mid-pt (list min-x (/ (+ min-y max-y) 2.0) 0.0)
                    pv1 (list (- min-x dim-off) min-y 0.0)
                    pv2 (list (- min-x dim-off) max-y 0.0))))

           (setq wall-id  (tz-next-wall-id)
                 area-sqft (* len-ft wall-ht))

           ;; ── Undo group for this wall ──
           (command "_.UNDO" "_Begin")

           ;; Place marker circle at bounding-edge midpoint
           (setvar "CLAYER" *TZ-LYR-WALL*)
           (setq marker (tz-make-circle mid-pt *TZ-WALL-RAD*))
           (setvar "CLAYER" lsave)

           (if marker
             (progn
               (tz-set-xdata marker
                 (list *TZ-APP*
                   (cons 1000 "WALL")
                   (cons 1000 wall-id)
                   (cons 1000 wall-type)
                   (cons 1000 zone-id)
                   (cons 1040 wall-ht)
                   (cons 1040 len-ft)
                   (cons 1040 area-sqft)
                   (cons 1040 az)
                   (cons 1040 (car  p1))
                   (cons 1040 (cadr p1))
                   (cons 1040 (car  p2))
                   (cons 1040 (cadr p2))))
               ;; Visual dimension line (dashed)
               (setvar "CLAYER" *TZ-LYR-LABEL*)
               (entmake (list '(0 . "LINE") '(100 . "AcDbEntity")
                          '(6 . "DASHED") '(48 . 12.0)
                          '(100 . "AcDbLine")
                          (cons 10 pv1) (cons 11 pv2)))
               (setvar "CLAYER" lsave)
               ;; Visual label — 6' further out
               (setq lbl-off (+ dim-off (* 6.0 *TZ-UNIT-FT*)))
               (cond
                 ((= side "N") (setq mid-pt (list (car mid-pt) (+ (cadr mid-pt) lbl-off) 0.0)))
                 ((= side "S") (setq mid-pt (list (car mid-pt) (- (cadr mid-pt) lbl-off) 0.0)))
                 ((= side "E") (setq mid-pt (list (+ (car mid-pt) lbl-off) (cadr mid-pt) 0.0)))
                 ((= side "W") (setq mid-pt (list (- (car mid-pt) lbl-off) (cadr mid-pt) 0.0))))
               (tz-wall-label mid-pt wall-id wall-type len-ft wall-ht area-sqft az label-ht)
               (setq count (1+ count))
               (princ (strcat "\n[T24] " wall-id ": " (rtos len-ft 2 1) "' "
                              wall-type " in " zone-id
                              " (" (rtos az 2 0) (chr 176) ") [" side " side]")))
             (princ "\n[T24] Failed to create wall marker."))

           (command "_.UNDO" "_End")))
       (setq prev-side nil  prev-ent nil))  ; reset preview tracking

      ;; ── Keyboard input (type 2) ──
      ((= gr-type 2)
       (cond
         ;; Enter (13), Space (32), Esc (0) → finish
         ((member gr-data '(13 32 0))
          (redraw)
          (setq loop-on nil))
         ;; U/u (85/117) → undo last wall
         ((member gr-data '(85 117))
          (if (> count 0)
            (progn
              (redraw)
              (command "_.UNDO" "")
              (setq count (1- count)  prev-side nil  prev-ent nil)
              (princ "\n[T24] Last wall undone."))
            (princ "\n[T24] Nothing to undo.")))))))

  (tz-bring-to-front)
  (princ (strcat "\n[T24] Done — placed " (itoa count) " wall sides."))
  (princ))

;; ── TZ-LISTDATA ──────────────────────────────────────────────────────────────
;; Shows full hierarchy: Zones → Walls → Openings nested under their wall.
(defun c:TZ-LISTDATA ( / ss ss-w ss-o i j k ent xd zid
                          ent2 xd2 wid
                          ent3 xd3 opzid opwid)
  (setq ss   (ssget "X" (list (cons 8 *TZ-LYR-ZONE*) '(0 . "LWPOLYLINE")))
        ss-w (ssget "X" (list (cons 8 *TZ-LYR-WALL*) '(0 . "CIRCLE"))))
  (princ "\n")
  (princ "\n========== T24 DATA ==========")
  (if (null ss)
    (princ "\n  (no zones)")
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i)
              xd  (tz-get-xdata ent))
        (if (and xd (equal (tz-xd-nth xd 1000 0) "ZONE"))
          (progn
            (setq zid (tz-xd-nth xd 1000 1))
            (princ (strcat "\n ZONE " zid
                           "  \"" (tz-xd-nth xd 1000 2) "\""
                           "  " (rtos (/ (vlax-curve-getarea (vlax-ename->vla-object ent))
                                         (* *TZ-UNIT-FT* *TZ-UNIT-FT*)) 2 1) " sqft"
                           "  " (rtos (tz-xd-num xd 1040 0 9.0) 2 1) "' clg"
                           "  Fl " (itoa (fix (tz-xd-num xd 1070 0 1)))))
            ;; ── Walls under this zone, with openings nested ──
            (if ss-w
              (progn
                (setq j 0)
                (repeat (sslength ss-w)
                  (setq ent2 (ssname ss-w j)
                        xd2  (tz-get-xdata ent2))
                  (if (and xd2 (= (tz-xd-nth xd2 1000 0) "WALL")
                           (= (tz-xd-nth xd2 1000 3) zid))
                    (progn
                      (setq wid (tz-xd-nth xd2 1000 1))
                      (princ (strcat "\n   +-- WALL " wid
                                     "  " (tz-xd-nth xd2 1000 2)
                                     "  " (rtos (tz-xd-num xd2 1040 1 0.0) 2 1) "'"
                                     " x " (rtos (tz-xd-num xd2 1040 0 9.0) 2 1) "'h"
                                     "  " (rtos (tz-xd-num xd2 1040 2 0.0) 2 1) " sqft"
                                     "  " (rtos (tz-xd-num xd2 1040 3 0.0) 2 0) "\260"))
                      ;; ── Openings under THIS wall (match by wall-id) ──
                      (foreach lyr (list *TZ-LYR-WIN* *TZ-LYR-DOOR* *TZ-LYR-SKY*)
                        (setq ss-o (ssget "X" (list (cons 8 lyr) '(0 . "CIRCLE"))))
                        (if ss-o
                          (progn
                            (setq k 0)
                            (repeat (sslength ss-o)
                              (setq ent3 (ssname ss-o k)
                                    xd3  (tz-get-xdata ent3))
                              (if (and xd3 (= (tz-xd-nth xd3 1000 0) "OPENING")
                                       (= (tz-xd-nth xd3 1000 4) wid))
                                (princ (strcat "\n   |  +-- " (tz-xd-nth xd3 1000 2)
                                               " " (tz-xd-nth xd3 1000 1)
                                               "  " (rtos (tz-xd-num xd3 1040 0 0.0) 2 1) " sqft")))
                              (setq k (1+ k))))))))
                  (setq j (1+ j)))))
            ;; ── Openings in this zone with no wall or unassigned ──
            (foreach lyr (list *TZ-LYR-WIN* *TZ-LYR-DOOR* *TZ-LYR-SKY*)
              (setq ss-o (ssget "X" (list (cons 8 lyr) '(0 . "CIRCLE"))))
              (if ss-o
                (progn
                  (setq k 0)
                  (repeat (sslength ss-o)
                    (setq ent3 (ssname ss-o k)
                          xd3  (tz-get-xdata ent3))
                    (if (and xd3 (= (tz-xd-nth xd3 1000 0) "OPENING")
                             (= (tz-xd-nth xd3 1000 3) zid)
                             (or (null (tz-xd-nth xd3 1000 4))
                                 (= (tz-xd-nth xd3 1000 4) "W-???")))
                      (princ (strcat "\n   +-- " (tz-xd-nth xd3 1000 2)
                                     " " (tz-xd-nth xd3 1000 1)
                                     "  " (rtos (tz-xd-num xd3 1040 0 0.0) 2 1) " sqft"
                                     "  (no wall)")))
                    (setq k (1+ k))))))))
        (setq i (1+ i)))))

  ;; ── Orphan openings (no zone-id at all) ──
  (princ "\n")
  (foreach lyr (list *TZ-LYR-WIN* *TZ-LYR-DOOR* *TZ-LYR-SKY*)
    (setq ss-o (ssget "X" (list (cons 8 lyr) '(0 . "CIRCLE"))))
    (if ss-o
      (progn
        (setq k 0)
        (repeat (sslength ss-o)
          (setq ent3 (ssname ss-o k)
                xd3  (tz-get-xdata ent3))
          (if (and xd3 (= (tz-xd-nth xd3 1000 0) "OPENING")
                   (null (tz-xd-nth xd3 1000 3)))
            (princ (strcat "\n +-- UNASSIGNED " (tz-xd-nth xd3 1000 2)
                           " " (tz-xd-nth xd3 1000 1)
                           "  " (rtos (tz-xd-num xd3 1040 0 0.0) 2 1) " sqft")))
          (setq k (1+ k))))))

  (princ "\n==============================")
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
    ;; Interior angle via cross product (sin) and dot product (cos)
    (setq ang (if (> denom 0.0)
               (progn
                 (setq cross-val (/ (- (* dx1 dy2) (* dy1 dx2)) denom))
                 (* (/ 180.0 pi) (atan (abs cross-val)
                                       (/ (+ (* dx1 dx2) (* dy1 dy2)) denom))))
               0.0))
    (tz-make-text
      (list (+ (car v1) 3) (+ (cadr v1) 3) 0.0)
      th
      (strcat (itoa i) ":" (rtos ang 2 1) (chr 176)))
    (setq i (1+ i)))
  (setvar "CLAYER" lsave)
  (princ (strcat "\n[T24] Labeled " (itoa n) " vertices."))
  (princ))

;; ── TZ-VALIDATE ────────────────────────────────────────────────────────────
;; Pre-export check: verifies zones have walls, openings are inside zones, no dup IDs
(defun c:TZ-VALIDATE ( / ss-z ss-w ss-o i ent xd zid zone-ids wall-zones
                         open-pos open-id open-type all-zone-pts
                         problems z-pts inside-any j z-ent wid oid
                         dup-check id-list)
  (setq problems 0)
  (princ "\n")
  (princ "\n=== T24 Validation ===")

  ;; Gather zone IDs and vertices
  (setq ss-z (ssget "X" (list (cons 8 *TZ-LYR-ZONE*) '(0 . "LWPOLYLINE"))))
  (if (null ss-z)
    (progn (princ "\n  ERROR: No zones found.") (setq problems (1+ problems)))
    (progn
      (setq zone-ids '()  all-zone-pts '()  i 0)
      (repeat (sslength ss-z)
        (setq ent (ssname ss-z i)
              xd  (tz-get-xdata ent))
        (if (and xd (equal (tz-xd-nth xd 1000 0) "ZONE"))
          (progn
            (setq zid (tz-xd-nth xd 1000 1))
            (setq zone-ids (cons zid zone-ids))
            (setq all-zone-pts (cons (cons zid (tz-get-pts ent)) all-zone-pts))))
        (setq i (1+ i)))
      (princ (strcat "\n  Zones: " (itoa (length zone-ids))))

      ;; Check for duplicate zone IDs
      (setq dup-check '())
      (foreach id zone-ids
        (if (member id dup-check)
          (progn
            (princ (strcat "\n  ERROR: Duplicate zone ID: " id))
            (setq problems (1+ problems)))
          (setq dup-check (cons id dup-check))))))

  ;; Check walls
  (setq ss-w (ssget "X" (list (cons 8 *TZ-LYR-WALL*) '(0 . "CIRCLE"))))
  (if (null ss-w)
    (progn (princ "\n  WARNING: No wall markers found.") (setq problems (1+ problems)))
    (progn
      (setq wall-zones '()  id-list '()  i 0)
      (repeat (sslength ss-w)
        (setq ent (ssname ss-w i)
              xd  (tz-get-xdata ent))
        (if (and xd (equal (tz-xd-nth xd 1000 0) "WALL"))
          (progn
            (setq wid (tz-xd-nth xd 1000 1)
                  zid (tz-xd-nth xd 1000 3))
            ;; Check for duplicate wall IDs
            (if (member wid id-list)
              (progn
                (princ (strcat "\n  ERROR: Duplicate wall ID: " wid))
                (setq problems (1+ problems)))
              (setq id-list (cons wid id-list)))
            (if (and zid (not (member zid wall-zones)))
              (setq wall-zones (cons zid wall-zones)))))
        (setq i (1+ i)))
      (princ (strcat "\n  Walls: " (itoa (sslength ss-w))))
      ;; Check that each zone has at least one wall
      (if zone-ids
        (foreach zid zone-ids
          (if (not (member zid wall-zones))
            (progn
              (princ (strcat "\n  WARNING: Zone " zid " has no wall markers."))
              (setq problems (1+ problems))))))))

  ;; Check openings
  (setq id-list '())
  (foreach lyr (list *TZ-LYR-WIN* *TZ-LYR-DOOR* *TZ-LYR-SKY*)
    (setq ss-o (ssget "X" (list (cons 8 lyr) '(0 . "CIRCLE"))))
    (if ss-o
      (progn
        (setq i 0)
        (repeat (sslength ss-o)
          (setq ent (ssname ss-o i)
                xd  (tz-get-xdata ent))
          (if (and xd (equal (tz-xd-nth xd 1000 0) "OPENING"))
            (progn
              (setq oid      (tz-xd-nth xd 1000 1)
                    open-type (tz-xd-nth xd 1000 2)
                    open-pos (cdr (assoc 10 (entget ent))))
              ;; Duplicate ID check
              (if (member oid id-list)
                (progn
                  (princ (strcat "\n  ERROR: Duplicate opening ID: " oid))
                  (setq problems (1+ problems)))
                (setq id-list (cons oid id-list)))
              ;; Check if opening is inside any zone
              (if all-zone-pts
                (progn
                  (setq inside-any nil)
                  (foreach zp all-zone-pts
                    (if (tz-pip (list (car open-pos) (cadr open-pos)) (cdr zp))
                      (setq inside-any T)))
                  (if (not inside-any)
                    (progn
                      (princ (strcat "\n  WARNING: " (if open-type open-type "Opening")
                                     " " oid " is outside all zone boundaries."))
                      (setq problems (1+ problems))))))))
          (setq i (1+ i))))))

  ;; Summary
  (princ "\n")
  (if (= problems 0)
    (princ "\n  All checks passed. Ready to export.")
    (princ (strcat "\n  Found " (itoa problems) " issue(s). Review before exporting.")))
  (princ "\n======================")
  (princ))

;; ── TZ-EDIT ────────────────────────────────────────────────────────────────
;; Click a zone polyline to view and edit its properties
(defun c:TZ-EDIT ( / sel ent xd zid zname cht fl cond occ area-ft
                     new-val pts centroid ss-lbl i lbl-ent lbl-pt lbl-dist
                     lsave)
  (tz-setup)
  (setq lsave (getvar "CLAYER"))
  (defun *error* (msg)
    (if lsave (setvar "CLAYER" lsave))
    (if (not (member msg '("Function cancelled" "quit / exit abort" "")))
      (princ (strcat "\n[T24] Error: " msg)))
    (princ))

  (princ "\n[T24] Click a zone polyline to edit: ")
  (setq sel (entsel))
  (if (null sel) (progn (princ "\n[T24] Cancelled.") (exit)))
  (setq ent (car sel))
  (if (/= (cdr (assoc 0 (entget ent))) "LWPOLYLINE")
    (progn (princ "\n[T24] Not a polyline.") (exit)))

  (setq xd (tz-get-xdata ent))
  (if (or (null xd) (not (equal (tz-xd-nth xd 1000 0) "ZONE")))
    (progn (princ "\n[T24] Not a T24 zone polyline.") (exit)))

  ;; Read current values
  (setq zid   (tz-xd-nth xd 1000 1)
        zname (tz-xd-nth xd 1000 2)
        cht   (tz-xd-num xd 1040 0 9.0)
        fl    (fix (tz-xd-num xd 1070 0 1))
        cond  (tz-xd-nth xd 1000 3)
        occ   (tz-xd-nth xd 1000 4))
  (if (null cond) (setq cond "Conditioned"))
  (if (null occ)  (setq occ "Residential"))
  (setq area-ft (/ (vlax-curve-getarea (vlax-ename->vla-object ent))
                    (* *TZ-UNIT-FT* *TZ-UNIT-FT*)))

  ;; Display current data
  (princ "\n")
  (princ "\n--- Zone Data ---")
  (princ (strcat "\n  ID:          " zid))
  (princ (strcat "\n  Name:        " zname))
  (princ (strcat "\n  Area:        " (rtos area-ft 2 1) " sqft"))
  (princ (strcat "\n  Ceiling Ht:  " (rtos cht 2 1) "'"))
  (princ (strcat "\n  Floor:       " (itoa fl)))
  (princ (strcat "\n  Condition:   " cond))
  (princ (strcat "\n  Occupancy:   " occ))
  (princ "\n-----------------")

  ;; Edit prompts (Enter keeps current value)
  (setq new-val (getstring T (strcat "\n[T24] Zone name <" zname ">: ")))
  (if (/= new-val "") (setq zname new-val))

  (setq new-val (getreal (strcat "\n[T24] Ceiling height ft <" (rtos cht 2 1) ">: ")))
  (if new-val (setq cht new-val))

  (setq new-val (getint (strcat "\n[T24] Floor number <" (itoa fl) ">: ")))
  (if new-val (setq fl new-val))

  (setq new-val (getstring (strcat "\n[T24] Condition <" cond ">: ")))
  (if (/= new-val "") (setq cond new-val))

  (setq new-val (getstring (strcat "\n[T24] Occupancy <" occ ">: ")))
  (if (/= new-val "") (setq occ new-val))

  ;; Update XDATA
  (tz-set-xdata ent
    (list *TZ-APP*
      (cons 1000 "ZONE")
      (cons 1000 zid)
      (cons 1000 zname)
      (cons 1040 cht)
      (cons 1070 fl)
      (cons 1000 cond)
      (cons 1000 occ)))

  ;; Delete old label text near this zone's centroid
  (setq pts      (tz-get-pts ent)
        centroid (tz-centroid pts))
  (setq ss-lbl (ssget "X" (list (cons 8 *TZ-LYR-LABEL*) '(0 . "TEXT"))))
  (if ss-lbl
    (progn
      (setq i 0)
      (repeat (sslength ss-lbl)
        (setq lbl-ent  (ssname ss-lbl i)
              lbl-pt   (cdr (assoc 10 (entget lbl-ent)))
              lbl-dist (distance (list (car centroid) (cadr centroid))
                                 (list (car lbl-pt) (cadr lbl-pt))))
        ;; Delete labels within 3' of centroid (same zone's labels)
        (if (< lbl-dist (* 3.0 *TZ-UNIT-FT*))
          (entdel lbl-ent))
        (setq i (1+ i)))))

  ;; Re-create label with updated data
  (tz-zone-label centroid zname area-ft cht fl zid)

  (princ (strcat "\n[T24] Zone " zid " updated."))
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
                       opid optype oparea opuf opshgc opctr opzid opwid
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
  (write-line (strcat "  \"climate_zone\": " (tz-jstr (if *TZ-CLIMATE-ZONE* *TZ-CLIMATE-ZONE* "3")) ",") fp)
  (write-line "  \"front_orientation\": 0," fp)
  (write-line (strcat "  \"north_arrow_angle\": " (tz-jnum (if *TZ-NORTH-ANGLE* *TZ-NORTH-ANGLE* 0.0)) ",") fp)
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
        (if (and xd (equal (tz-xd-nth xd 1000 0) "ZONE")
                 (tz-ent-visible-p ent))
          (progn
            (setq obj       (vlax-ename->vla-object ent)
                  pts       (tz-get-pts ent)
                  area-sqft (/ (vlax-curve-getarea obj)
                               (* *TZ-UNIT-FT* *TZ-UNIT-FT*))
                  cent      (tz-centroid pts)
                  zid       (tz-xd-nth xd 1000 1)
                  zname     (tz-xd-nth xd 1000 2)
                  ht        (tz-xd-num xd 1040 0 9.0)
                  fl        (tz-xd-num xd 1070 0 1)
                  cond      (tz-xd-nth xd 1000 3)
                  occ       (tz-xd-nth xd 1000 4))
            (if (null zid)   (setq zid "Z-???"))
            (if (null zname) (setq zname ""))
            (if (null cond)  (setq cond "Conditioned"))
            (if (null occ)   (setq occ "Residential"))
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
        (if (and xd (equal (tz-xd-nth xd 1000 0) "WALL")
                 (tz-ent-visible-p ent))
          (progn
            (setq wid   (tz-xd-nth xd 1000 1)
                  wtype (tz-xd-nth xd 1000 2)
                  wzid  (tz-xd-nth xd 1000 3))
            (if (null wid)   (setq wid "W-???"))
            (if (null wtype) (setq wtype "Exterior Wall"))
            (if (null wzid)  (setq wzid "Z-???"))
            (setq wht   (tz-xd-num xd 1040 0 9.0)
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
          (if (and xd (equal (tz-xd-nth xd 1000 0) "OPENING")
                   (tz-ent-visible-p ent))
            (progn
              (setq opctr  (cdr (assoc 10 (entget ent)))
                    opid   (tz-xd-nth xd 1000 1)
                    optype (tz-xd-nth xd 1000 2)
                    oparea (tz-xd-num xd 1040 0 0.0))
              (if (null opid)   (setq opid "O-???"))
              (if (null optype) (setq optype "Window"))
              ;; zone_id and wall_id at 1000 indices 3 and 4 (new format)
              ;; Falls back gracefully for old openings without these fields
              (setq opzid  (tz-xd-nth xd 1000 3)
                    opwid  (tz-xd-nth xd 1000 4)
                    opuf   (tz-xd-num xd 1040 1 0.0)
                    opshgc (tz-xd-num xd 1040 2 0.0))
              (if open-started (write-line "    }," fp))
              (write-line "    {" fp)
              (write-line (strcat "      \"id\": "       (tz-jstr opid)                       ",") fp)
              (write-line (strcat "      \"type\": "     (tz-jstr optype)                     ",") fp)
              (if opzid
                (write-line (strcat "      \"zone_id\": "  (tz-jstr opzid)                    ",") fp))
              (if opwid
                (write-line (strcat "      \"wall_id\": "  (tz-jstr opwid)                    ",") fp))
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

;; ── TZ-EXPORT-ALL ────────────────────────────────────────────────────────────
;; Exports current drawing, then shows how to merge all JSONs in the folder.
;; AutoCAD LISP can't switch documents, so run TZ-EXPORT in each drawing
;; first, then use the Python command to merge.
(defun c:TZ-EXPORT-ALL ( / dwg-dir)
  ;; Export current drawing
  (c:TZ-EXPORT)
  (setq dwg-dir (getvar "DWGPREFIX"))
  (princ "\n")
  (princ "\n[T24] To merge all drawings into one Excel:")
  (princ "\n[T24]   1. Switch to each open drawing and run TZ-EXPORT")
  (princ (strcat "\n[T24]   2. Then run:  python tz_to_excel.py \"" dwg-dir "\""))
  (princ "\n[T24]   This will find all *_t24.json files in the folder and merge them.")
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
  ;; Reset session counters so IDs start fresh
  (setq *TZ-ZONE-COUNT* nil  *TZ-OPEN-COUNT* nil  *TZ-WALL-COUNT* nil)
  (princ "\n[T24] Labels and markers cleared. Zone polylines and XDATA retained.")
  (princ))

;; ── TZ-RESET-ALL ────────────────────────────────────────────────────────────
;; Full reset: clears everything including zone polylines
(defun c:TZ-RESET-ALL ( / ans ss i)
  (setq ans (getstring "\n[T24] Delete ALL T24 data including zone boundaries? Type YES to confirm: "))
  (if (/= (strcase ans) "YES")
    (progn (princ "\n[T24] Cancelled.") (exit)))
  (foreach lyr (list *TZ-LYR-LABEL* *TZ-LYR-WIN* *TZ-LYR-DOOR* *TZ-LYR-SKY* *TZ-LYR-WALL* *TZ-LYR-ZONE*)
    (setq ss (ssget "X" (list (cons 8 lyr))))
    (if ss
      (progn
        (setq i 0)
        (repeat (sslength ss)
          (entdel (ssname ss i))
          (setq i (1+ i))))))
  ;; Reset session counters
  (setq *TZ-ZONE-COUNT* nil  *TZ-OPEN-COUNT* nil  *TZ-WALL-COUNT* nil)
  (princ "\n[T24] Full reset complete. All T24 data removed.")
  (princ))


(setq *TZ-REACTORS* nil)  ;; list of active reactors

(defun tz-zone-modified-callback (owner reactor params / ent xd zid zname cht fl
                                   area-ft centroid pts obj
                                   ss i tent txd ted tstr)
  "Object reactor callback: fires after a zone polyline is modified."
  (setq ent (vlax-vla-object->ename owner))
  (if (null ent) (exit))
  ;; Only process LWPOLYLINE on T24-ZONE layer
  (if (and (= (cdr (assoc 0 (entget ent))) "LWPOLYLINE")
           (= (cdr (assoc 8 (entget ent))) *TZ-LYR-ZONE*))
    (progn
      (setq xd (tz-get-xdata ent))
      (if (and xd (= (tz-xd-nth xd 1000 0) "ZONE"))
        (progn
          (setq zid   (tz-xd-nth xd 1000 1)
                zname (tz-xd-nth xd 1000 2)
                cht   (tz-xd-num xd 1040 0 9.0)
                fl    (fix (tz-xd-num xd 1070 0 1)))
          ;; Recalculate area from geometry
          (setq obj     (vlax-ename->vla-object ent)
                area-ft (/ (vlax-curve-getarea obj)
                           (* *TZ-UNIT-FT* *TZ-UNIT-FT*))
                pts     (tz-get-pts ent)
                centroid (tz-centroid pts))
          ;; Find THIS zone's detail label by matching zone-id in XDATA
          (setq ss (ssget "X" (list (cons 8 *TZ-LYR-LABEL*) '(0 . "TEXT"))))
          (if ss
            (progn
              (setq i 0)
              (repeat (sslength ss)
                (setq tent (ssname ss i)
                      txd  (tz-get-xdata tent))
                (if (and txd
                         (= (tz-xd-nth txd 1000 0) "ZONE-LABEL")
                         (= (tz-xd-nth txd 1000 1) zid))
                  ;; Found our label — only update if text actually changed
                  (progn
                    (setq ted (entget tent))
                    (setq tstr (strcat (rtos area-ft 2 1) " sqft  |  "
                                       (rtos cht 2 1) "' clg  |  Fl " (itoa fl)))
                    (if (/= (cdr (assoc 1 ted)) tstr)
                      (progn
                        (entmod (subst (cons 1 tstr) (assoc 1 ted) ted))
                        (entupd tent)
                        (princ (strcat "\n[T24] Updated: " zname " -> " (rtos area-ft 2 1) " sqft"))))))
                (setq i (1+ i))))))))))
(defun c:TZ-WATCH ( / ss i ent obj reactor-list)
  "Install reactors on all zone polylines for dynamic area updates."
  (tz-setup)
  ;; Remove existing reactors
  (if *TZ-REACTORS*
    (progn
      (foreach r *TZ-REACTORS*
        (if (not (vlr-removed-p r))
          (vlr-remove r)))
      (setq *TZ-REACTORS* nil)))
  ;; Find all zone polylines
  (setq ss (ssget "X" (list (cons 8 *TZ-LYR-ZONE*) '(0 . "LWPOLYLINE"))))
  (if (null ss)
    (progn (princ "\n[T24] No zone polylines found.") (princ))
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i)
              obj (vlax-ename->vla-object ent))
        ;; Only attach to entities with ZONE xdata
        (if (and (tz-get-xdata ent)
                 (= (tz-xd-nth (tz-get-xdata ent) 1000 0) "ZONE"))
          (progn
            (setq reactor-list
              (vlr-object-reactor
                (list obj)
                "T24-Zone-Update"
                '((:vlr-modified . tz-zone-modified-callback))))
            (setq *TZ-REACTORS* (cons reactor-list *TZ-REACTORS*))))
        (setq i (1+ i)))
      (princ (strcat "\n[T24] Watching " (itoa (length *TZ-REACTORS*))
                     " zone(s) for changes. Edit polylines to auto-update area."))
      (princ))))

;; ── Auto-rewatch: command reactor to reattach zone reactors after edits ──────
;; When polylines are heavily edited (PEDIT, STRETCH, grip edits, UNDO),
;; AutoCAD may recreate the entity internally, which orphans object reactors.
;; This command reactor re-runs the watch logic after those commands finish.

(setq *TZ-CMD-REACTOR* nil)

(defun tz-cmd-ended-callback (reactor params / cmd)
  (setq cmd (strcase (car params)))
  (if (member cmd '("GRIP_STRETCH" "STRETCH" "PEDIT" "MOVE" "COPY"
                     "MIRROR" "ROTATE" "SCALE" "UNDO" "U" "MREDO"
                     "PROPERTIES" "MATCHPROP"))
    ;; Re-attach object reactors on all zone polylines
    (tz-rewatch-zones)))


;; ── Auto-rewatch: command reactor to reattach zone reactors after edits ──────
(setq *TZ-CMD-REACTOR* nil)

(defun tz-cmd-ended-callback (reactor params / cmd)
  (if (null *TZ-BUSY*)
    (progn
      (setq cmd (strcase (car params)))
      (if (member cmd '("GRIP_STRETCH" "STRETCH" "PEDIT" "MOVE" "COPY"
                         "MIRROR" "ROTATE" "SCALE" "UNDO" "U" "MREDO"))
        (c:TZ-WATCH)))))

;; Install command reactor at load time
(if *TZ-CMD-REACTOR*
  (if (not (vlr-removed-p *TZ-CMD-REACTOR*))
    (vlr-remove *TZ-CMD-REACTOR*)))
(setq *TZ-CMD-REACTOR*
  (vlr-command-reactor nil
    '((:vlr-commandEnded . tz-cmd-ended-callback))))

;; Auto-attach reactors to existing zones on load
(c:TZ-WATCH)


;; ── Shape simplification helpers ─────────────────────────────────────────────

;; 2D cross product of vectors OA and OB
(defun tz-cross-2d (o a b)
  (- (* (- (car a) (car o)) (- (cadr b) (cadr o)))
     (* (- (cadr a) (cadr o)) (- (car b) (car o)))))

;; Graham scan convex hull. Returns list of (x y) in CCW order.
(defun tz-convex-hull (pts / sorted n lower upper p)
  (if (<= (length pts) 2)
    pts
    (progn
      (setq sorted
        (vl-sort pts
          '(lambda (a b)
             (if (= (car a) (car b))
               (< (cadr a) (cadr b))
               (< (car a) (car b))))))
      (setq lower '())
      (foreach p sorted
        (while (and (>= (length lower) 2)
                    (<= (tz-cross-2d (nth (- (length lower) 2) lower)
                                     (nth (- (length lower) 1) lower) p)
                        0.0))
          (setq lower (reverse (cdr (reverse lower)))))
        (setq lower (append lower (list p))))
      (setq upper '())
      (foreach p (reverse sorted)
        (while (and (>= (length upper) 2)
                    (<= (tz-cross-2d (nth (- (length upper) 2) upper)
                                     (nth (- (length upper) 1) upper) p)
                        0.0))
          (setq upper (reverse (cdr (reverse upper)))))
        (setq upper (append upper (list p))))
      (append (reverse (cdr (reverse lower)))
              (reverse (cdr (reverse upper)))))))

;; Concavity filter: compute convex hull, then walk the original polygon.
;; Only fill concavities smaller than threshold (sqft). Keep large ones.
(defun tz-filter-concavities (pts threshold-sqft / hull hull-area orig-area
                                    concavity-area)
  (setq hull (tz-convex-hull pts))
  (if (null hull)
    pts
    (progn
      ;; If the concavity (hull area - original area) is small, use hull
      (setq hull-area (tz-poly-area hull)
            orig-area (tz-poly-area pts)
            concavity-area (/ (- hull-area orig-area) (* *TZ-UNIT-FT* *TZ-UNIT-FT*)))
      (princ (strcat "\n[T24] Concavity: " (rtos concavity-area 2 1)
                     " sqft (threshold " (rtos threshold-sqft 2 0) ")"))
      (if (<= concavity-area threshold-sqft)
        (progn (princ " -> using convex hull") hull)
        (progn (princ " -> keeping original shape") pts)))))

;; Signed polygon area (shoelace formula), returns absolute area in drawing units^2
(defun tz-poly-area (pts / n i sum p1 p2)
  (setq n (length pts) sum 0.0 i 0)
  (repeat n
    (setq p1 (nth i pts)
          p2 (nth (rem (1+ i) n) pts)
          sum (+ sum (- (* (car p1) (cadr p2)) (* (car p2) (cadr p1))))
          i (1+ i)))
  (abs (/ sum 2.0)))

;; Axis-aligned bounding box. Returns 4 corner points.
(defun tz-bounding-box (pts / min-x min-y max-x max-y)
  (setq min-x (apply 'min (mapcar 'car pts))
        max-x (apply 'max (mapcar 'car pts))
        min-y (apply 'min (mapcar 'cadr pts))
        max-y (apply 'max (mapcar 'cadr pts)))
  (list (list min-x min-y) (list max-x min-y)
        (list max-x max-y) (list min-x max-y)))

;; Dispatcher: apply shape simplification mode to a polyline entity
(defun tz-shape-simplify (ent mode / pts new-pts obj n-before n-after label)
  (setq pts (tz-get-pts ent)
        n-before (length pts))
  (cond
    ((= mode 1)  ; Convex hull
     (setq new-pts (tz-convex-hull pts)
           label "Convex hull"))
    ((= mode 2)  ; Concavity threshold (50 sqft default)
     (setq new-pts (tz-filter-concavities pts 50.0)
           label "Concavity filter"))
    ((= mode 3)  ; Bounding box
     (setq new-pts (tz-bounding-box pts)
           label "Bounding box"))
    (T (setq new-pts pts label "None")))
  (if (and new-pts (>= (length new-pts) 3) (not (equal new-pts pts)))
    (progn
      (setq obj (vlax-ename->vla-object ent)
            n-after (length new-pts))
      (tz-set-pline-verts obj
        (mapcar '(lambda (p) (list (car p) (cadr p) 0.0)) new-pts))
      (princ (strcat "\n[T24] " label ": " (itoa n-before)
                     " -> " (itoa n-after) " vertices")))
    (princ (strcat "\n[T24] " label ": no change"))))

;; ── TZ-ZONE1 / TZ-ZONE2 / TZ-ZONE3 ────────────────────────────────────────
;; Thin wrappers: set shape mode, call TZ-ZONE, clear mode.

(defun c:TZ-ZONECONVEX ()
  (princ "\n[T24] Mode: CONVEX HULL")
  (setq *TZ-SHAPE-MODE* 1)
  (c:TZ-ZONE)
  (setq *TZ-SHAPE-MODE* nil)
  (princ))

(setq *TZ-SHAPE-MODE* nil)

;; ── TZ-ZTEST — Diagnostic boundary tool ──────────────────────────────────────
;; Evaluates what BOUNDARY sees at a pick point. Reports nearby entity types,
;; Z-elevations, blocks/proxies, and tries BOUNDARY at multiple gap tolerances.
;; Leaves result polylines visible for inspection.
(defun c:TZ-ZTEST ( / pt radius ss i ent ed etype lyr elev
                       type-counts z-vals z-val block-names proxy-count
                       aec-count xref-count total
                       gap-list gap last-ent scan-ent new-ent area-sqft
                       results ce saved-gaptol)
  (tz-setup)
  (setq ce (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)

  ;; ── Pick point ──
  (setq pt (getpoint "\n[ZTEST] Click inside room to diagnose: "))
  (if (null pt) (progn (setvar "CMDECHO" ce) (princ "\n[ZTEST] Cancelled.") (exit)))

  (princ (strcat "\n[ZTEST] Point: " (rtos (car pt) 2 4) ", " (rtos (cadr pt) 2 4)
                 ", Z=" (rtos (if (caddr pt) (caddr pt) 0.0) 2 4)))

  ;; ── Scan nearby geometry ──
  (setq radius 600.0)  ; 50 feet in inches
  (princ (strcat "\n[ZTEST] Scanning entities within " (rtos radius 2 0) "\" radius..."))
  (setq ss (ssget "_C"
             (list (- (car pt) radius) (- (cadr pt) radius))
             (list (+ (car pt) radius) (+ (cadr pt) radius))))

  (if (null ss)
    (princ "\n[ZTEST] WARNING: No entities found near pick point!")
    (progn
      (setq total (sslength ss)
            type-counts '()
            z-vals '()
            block-names '()
            proxy-count 0
            aec-count 0
            xref-count 0
            i 0)
      (repeat total
        (setq ent (ssname ss i)
              ed  (entget ent)
              etype (cdr (assoc 0 ed))
              lyr   (cdr (assoc 8 ed)))

        ;; Count entity types
        (setq type-counts
          (if (assoc etype type-counts)
            (mapcar '(lambda (x) (if (equal (car x) etype) (cons etype (1+ (cdr x))) x))
                    type-counts)
            (cons (cons etype 1) type-counts)))

        ;; Check Z-elevations
        (cond
          ((assoc 38 ed)  ; LWPOLYLINE elevation
           (setq z-val (cdr (assoc 38 ed)))
           (if (not (member z-val z-vals)) (setq z-vals (cons z-val z-vals))))
          ((assoc 10 ed)  ; Insertion/start point Z
           (setq z-val (caddr (cdr (assoc 10 ed))))
           (if (and z-val (not (member z-val z-vals)))
             (setq z-vals (cons z-val z-vals)))))

        ;; Check for problematic entity types
        (cond
          ((= etype "INSERT")
           (setq block-names (cons (cdr (assoc 2 ed)) block-names)))
          ((or (= etype "ACAD_PROXY_ENTITY") (wcmatch etype "Aec*,aec*"))
           (setq proxy-count (1+ proxy-count)))
          ((wcmatch etype "*XREF*,*Xref*")
           (setq xref-count (1+ xref-count))))

        (setq i (1+ i)))

      ;; ── Report ──
      (princ (strcat "\n[ZTEST] " (itoa total) " entities found:"))
      (foreach tc (vl-sort type-counts '(lambda (a b) (> (cdr a) (cdr b))))
        (princ (strcat "\n[ZTEST]   " (car tc) ": " (itoa (cdr tc)))))

      ;; Z-elevations
      (princ (strcat "\n[ZTEST] Z-elevations found: "
        (if z-vals
          (apply 'strcat (mapcar '(lambda (z) (strcat (rtos z 2 4) " ")) z-vals))
          "none detected")))
      (if (> (length z-vals) 1)
        (princ "\n[ZTEST] *** WARNING: Multiple Z-elevations! BOUNDARY may not see all geometry. ***"))

      ;; Blocks
      (if block-names
        (progn
          (princ (strcat "\n[ZTEST] Block references found (" (itoa (length block-names)) "):"))
          ;; Show unique block names
          (foreach bn (vl-remove-if
                        '(lambda (x) (member x (cdr (member x block-names))))
                        block-names)
            (princ (strcat "\n[ZTEST]   INSERT: \"" bn "\"")))
          (princ "\n[ZTEST] *** WARNING: Blocks are invisible to BOUNDARY. Explode them first. ***")))

      ;; Proxies / AEC
      (if (> proxy-count 0)
        (princ (strcat "\n[ZTEST] *** WARNING: " (itoa proxy-count) " proxy/AEC entities — BOUNDARY ignores these! ***")))

      ;; Layer summary
      (princ "\n[ZTEST] Current ELEVATION sysvar: ")
      (princ (rtos (getvar "ELEVATION") 2 4))
    ))

  ;; ── Try BOUNDARY at multiple zoom levels and gap tolerances ──
  (princ "\n\n[ZTEST] Running REGENALL...")
  (command "_.REGENALL")
  (setq gap-list '(0.0 1.0 4.0 12.0 36.0 48.0)
        saved-gaptol (getvar "HPGAPTOL")
        results '())

  ;; Test at 3 zoom levels × 6 gap tolerances
  (foreach zoom-info (list
    (cons "Current zoom" nil)
    (cons "Local zoom (100')" 1200.0)
    (cons "Zoom Extents" T))

    (princ (strcat "\n\n[ZTEST] === " (car zoom-info) " ==="))

    ;; Set zoom level
    (cond
      ((null (cdr zoom-info)) nil)  ; current zoom, do nothing
      ((= (cdr zoom-info) T)        ; zoom extents
       (command "_.ZOOM" "_Extents"))
      (T                             ; local window
       (command "_.ZOOM" "_Window"
         (list (- (car pt) (cdr zoom-info)) (- (cadr pt) (cdr zoom-info)))
         (list (+ (car pt) (cdr zoom-info)) (+ (cadr pt) (cdr zoom-info))))))

    (foreach gap gap-list
      (setvar "HPGAPTOL" gap)
      (setq last-ent (entlast))
      (command "_-BOUNDARY" pt "")

      (if (equal (entlast) last-ent)
        (princ (strcat "\n[ZTEST]   Gap " (rtos gap 2 1) "\": FAILED"))
        (progn
          ;; Collect all created entities
          (setq scan-ent (entnext last-ent)  new-ent nil  area-sqft 0.0)
          (while scan-ent
            (if (= (cdr (assoc 0 (entget scan-ent))) "LWPOLYLINE")
              (progn
                (setq area-sqft (/ (vlax-curve-getarea (vlax-ename->vla-object scan-ent))
                                   (* *TZ-UNIT-FT* *TZ-UNIT-FT*)))
                (if (null new-ent)
                  (setq new-ent scan-ent)
                  ;; Keep the smallest polyline (likely the room, not an escape)
                  (if (< area-sqft (/ (vlax-curve-getarea (vlax-ename->vla-object new-ent))
                                      (* *TZ-UNIT-FT* *TZ-UNIT-FT*)))
                    (progn (entdel new-ent) (setq new-ent scan-ent))
                    (entdel scan-ent)))))
            (setq scan-ent (entnext scan-ent)))

          (if new-ent
            (progn
              (setq area-sqft (/ (vlax-curve-getarea (vlax-ename->vla-object new-ent))
                                 (* *TZ-UNIT-FT* *TZ-UNIT-FT*)))
              (princ (strcat "\n[ZTEST]   Gap " (rtos gap 2 1) "\": "
                             (rtos area-sqft 2 1) " sqft"
                             (if (> area-sqft 5000.0) "  *** ESCAPED ***" "")))
              ;; Keep first reasonable result for visual inspection
              (if (and (null results) (> area-sqft 10.0))
                (progn
                  (entmod (subst '(62 . 3) (assoc 62 (entget new-ent))
                                 (if (assoc 62 (entget new-ent))
                                   (entget new-ent)
                                   (append (entget new-ent) '((62 . 3))))))
                  (entmod (subst (cons 8 *TZ-LYR-LABEL*) (assoc 8 (entget new-ent)) (entget new-ent)))
                  (setq results (cons (list (car zoom-info) gap area-sqft new-ent) results))
                  (princ "  [KEPT - green]"))
                (entdel new-ent)))
            (princ (strcat "\n[ZTEST]   Gap " (rtos gap 2 1) "\": entity created but no polyline"))))))

    ;; Restore zoom if we changed it
    (if (cdr zoom-info) (command "_.ZOOM" "_Previous")))

  (setvar "HPGAPTOL" saved-gaptol)

  ;; ── Summary ──
  (princ "\n\n[ZTEST] === Summary ===")
  (if results
    (progn
      (princ (strcat "\n[ZTEST] Best result: " (rtos (caddr (car results)) 2 1) " sqft"
                     " at \"" (car (car results)) "\" gap=" (rtos (cadr (car results)) 2 1) "\""))
      (princ "\n[ZTEST] Kept on T24-LABEL layer (green)."))
    (princ "\n[ZTEST] All 18 attempts failed. BOUNDARY cannot see room geometry."))
  (princ "\n[ZTEST] Suggestions if BOUNDARY escapes:")
  (princ "\n[ZTEST]   1. EXPLODE all blocks near the room (check INSERT count above)")
  (princ "\n[ZTEST]   2. Check Z-elevations: FLATTEN or set ELEVATION sysvar")
  (princ "\n[ZTEST]   3. Look for proxy/AEC entities (need BURST or EXPLODEPROXY)")
  (princ "\n[ZTEST]   4. Check for overlapping duplicate lines on same edge")
  (princ "\n[ZTEST]   5. Use TZ-ZONE fallback: Draw rectangle or Draw polyline")
  (setvar "CMDECHO" ce)
  (princ))

;; ── Load message ─────────────────────────────────────────────────────────────
(princ "\n")
(princ "\n+--------------------------------------------+")
(princ "\n|    T24 TakeOff Wizard  v2 Loaded           |")
(princ "\n+--------------------------------------------+")
(princ "\n|  TZ-ZONE      - Click room name text       |")
(princ "\n|  TZ-WALL      - Place wall markers         |")
(princ "\n|  TZ-WSIDE     - Wall side (bbox projected)  |")
(princ "\n|  TZ-WIN       - Place window markers       |")
(princ "\n|  TZ-DOOR      - Place door marker          |")
(princ "\n|  TZ-SKY       - Place skylight marker      |")
(princ "\n|  TZ-EDIT      - Edit a zone's properties   |")
(princ "\n|  TZ-VALIDATE  - Pre-export data check      |")
(princ "\n|  TZ-EXPORT    - Write _t24.json for Excel  |")
(princ "\n|  TZ-EXPORT-ALL- Export all open drawings   |")
(princ "\n|  TZ-LISTDATA  - List zone data             |")
(princ "\n|  TZ-SHOWVERTS - Inspect polyline vertices  |")
(princ "\n|  TZ-ZONECONVEX- Zone + convex hull          |")
(princ "\n|  TZ-WATCH     - Auto-update on pline edit  |")
(princ "\n|  TZ-RESET     - Clear labels/markers       |")
(princ "\n|  TZ-RESET-ALL - Full reset (incl. zones)   |")
(princ "\n+--------------------------------------------+")
(princ "\n")
(princ)
