;;; ============================================================================
;;; METRACAD | Automated CAD Length Measurement Utility
;;; Version: 1.0.0
;;; Author:  Jean Carlos
;;; License: MIT
;;; ============================================================================

(vl-load-com)

;;; GLOBAL CONFIGURATION
(setq *MC-VERSION* "1.0.0")
(setq *MC-LAST* nil)   
(setq *MC-HISTORY* nil)   
(setq *MC-DECIMALS* 2)    
(setq *MC-UNIT* "m")   

;;; UTILITY FUNCTIONS
(defun MC:fmt (n) (rtos n 2 *MC-DECIMALS*))

(defun MC:pad-r (str width)
  (if (> width (strlen str))
    (strcat str (substr "                                " 1 (- width (strlen str))))
    (substr str 1 width)))

(defun MC:pad-l (str width)
  (if (> width (strlen str))
    (strcat (substr "                                " 1 (- width (strlen str))) str)
    str))

(defun MC:timestamp () (menucmd "m=$(edtime,$(getvar,date),DD/MM/YYYY HH:MM:SS)"))

(defun MC:dwg-name () (vla-get-Name (vla-get-ActiveDocument (vlax-get-acad-object))))

(defun MC:clipboard (text / tmp f)
  (setq tmp (strcat (getvar "TEMPPREFIX") "mc_clip.tmp"))
  (setq f (open tmp "w"))
  (write-line text f) (close f)
  (vl-cmdf "_.SHELL" (strcat "cmd /c type \"" tmp "\" | clip"))
  (vl-file-delete tmp))

;;; CORE MEASUREMENT LOGIC
(defun c:metracad (/ *error* ss i ent obj type len sum-lin sum-pol sum-arc 
                     sum-cir sum-spl sum-ell cnt-lin cnt-pol cnt-arc 
                     cnt-cir cnt-spl cnt-ell total-sum layer-data mode)

  (defun *error* (msg)
    (if (not (member msg '("Function cancelled" "quit / exit abort")))
      (princ (strcat "\nError: " msg)))
    (princ))

  ;; Configuration Input
  (initget "2 3 4")
  (setq dec-opt (getkword (strcat "\nPrecision [2/3/4] <" (itoa *MC-DECIMALS*) ">: ")))
  (if dec-opt (setq *MC-DECIMALS* (atoi dec-opt)))

  (initget "m cm mm ft")
  (setq unit-opt (getkword (strcat "\nUnit [m/cm/mm/ft] <" *MC-UNIT* ">: ")))
  (if unit-opt (setq *MC-UNIT* unit-opt))

  ;; Selection Logic
  (initget "Selection Layer Color")
  (setq mode (getkword "\nMode [Selection/Layer/Color] <Selection>: "))
  (setq mode (if mode mode "Selection"))

  (setq filter '((0 . "LINE,LWPOLYLINE,POLYLINE,ARC,CIRCLE,SPLINE,ELLIPSE")))

  (cond
    ((= mode "Selection") (setq ss (ssget filter)))
    ((= mode "Layer") 
     (if (setq ent (car (entsel "\nSelect object for layer filter: ")))
       (setq ss (ssget "_X" (list (assoc 8 (entget ent)) (car filter))))))
    ((= mode "Color")
     (if (setq ent (car (entsel "\nSelect object for color filter: ")))
       (setq ss (ssget "_X" (list (cons 62 (cond ((cdr (assoc 62 (entget ent)))) (256))) (car filter)))))))

  (if (null ss) (progn (princ "\nEmpty selection set.") (exit)))

  ;; Processing
  (setq cnt-lin 0 cnt-pol 0 cnt-arc 0 cnt-cir 0 cnt-spl 0 cnt-ell 0
        sum-lin 0.0 sum-pol 0.0 sum-arc 0.0 sum-cir 0.0 sum-spl 0.0 sum-ell 0.0
        layer-data nil)

  (repeat (setq i (sslength ss))
    (setq ent (ssname ss (setq i (1- i)))
          obj (vlax-ename->vla-object ent)
          type (cdr (assoc 0 (entget ent)))
          lay (cdr (assoc 8 (entget ent)))
          len (vlax-curve-getDistAtParam obj (vlax-curve-getEndParam obj)))
    
    (cond
      ((= type "LINE") (setq cnt-lin (1+ cnt-lin) sum-lin (+ sum-lin len)))
      ((wcmatch type "*POLYLINE") (setq cnt-pol (1+ cnt-pol) sum-pol (+ sum-pol len)))
      ((= type "ARC") (setq cnt-arc (1+ cnt-arc) sum-arc (+ sum-arc len)))
      ((= type "CIRCLE") (setq cnt-cir (1+ cnt-cir) sum-cir (+ sum-cir len)))
      ((= type "SPLINE") (setq cnt-spl (1+ cnt-spl) sum-spl (+ sum-spl len)))
      ((= type "ELLIPSE") (setq cnt-ell (1+ cnt-ell) sum-ell (+ sum-ell len))))
    
    (setq entry (assoc lay layer-data))
    (if entry
      (setq layer-data (subst (list lay (+ (cadr entry) len) (1+ (caddr entry))) entry layer-data))
      (setq layer-data (append layer-data (list (list lay len 1))))))

  (setq total-sum (+ sum-lin sum-pol sum-arc sum-cir sum-spl sum-ell))

  ;; Report Generation
  (setq report (strcat "\nMETRACAD v" *MC-VERSION* "\n"
                       "File: " (MC:dwg-name) " | " (MC:timestamp) "\n"
                       "Mode: " mode " | Unit: " *MC-UNIT* "\n"
                       "------------------------------------------------\n"
                       "TYPE          COUNT       LENGTH\n"
                       "------------------------------------------------\n"))

  (defun add-row (lbl c s)
    (if (> c 0) (setq report (strcat report (MC:pad-r lbl 14) (MC:pad-l (itoa c) 5) (MC:pad-l (MC:fmt s) 15) " " *MC-UNIT* "\n"))))

  (add-row "Lines" cnt-lin sum-lin)
  (add-row "Polylines" cnt-pol sum-pol)
  (add-row "Arcs" cnt-arc sum-arc)
  (add-row "Circles" cnt-cir sum-cir)
  (add-row "Splines" cnt-spl sum-spl)
  (add-row "Ellipses" cnt-ell sum-ell)

  (setq report (strcat report "------------------------------------------------\n"
                       (MC:pad-r "TOTAL" 14) (MC:pad-l (itoa (sslength ss)) 5) (MC:pad-l (MC:fmt total-sum) 15) " " *MC-UNIT* "\n"))

  (setq *MC-LAST* report)
  (setq *MC-HISTORY* (append *MC-HISTORY* (list (list (MC:timestamp) (MC:dwg-name) mode (MC:fmt total-sum) *MC-UNIT*))))
  
  (princ report) (alert report)
  (princ))

;;; EXPORT COMMAND
(defun c:metracadexport (/ path f)
  (if (null *MC-LAST*) (progn (princ "\nNo data to export.") (exit)))
  (setq path (getfiled "Export CSV Result" (strcat (getvar "DWGPREFIX") "quantities") "csv" 1))
  (if path
    (progn
      (setq f (open path "w"))
      (write-line "Date,File,Mode,Total,Unit" f)
      (foreach e *MC-HISTORY*
        (write-line (strcat (nth 0 e) "," (nth 1 e) "," (nth 2 e) "," (nth 3 e) "," (nth 4 e)) f))
      (close f) (princ (strcat "\nExported: " path))))
  (princ))

(princ (strcat "\nMetraCAD v" *MC-VERSION* " loaded. Commands: METRACAD, METRACADEXPORT"))
(princ)