;;;; omega.lisp — the family's first executed trace.
;;;; Run with: sbcl --script omega.lisp

(defvar *fuel* 0
  "Expansion budget. Decremented once per macroexpand-1.
   When it reaches zero, expansion did not terminate on its own.")

(defvar *steps* 0 "How many expansions actually happened.")

;;; ---------------------------------------------------------------
;;; 1. The stairwell, as METAMETALISP4.md §-1.5 defined it:
;;;    index set = the rationals; the interval (-2, -1) is dense;
;;;    each expansion halves the remaining distance to the reader.
;;; ---------------------------------------------------------------

(defmacro bisect (reader section)
  "Splice a new section at the midpoint, then keep going."
  (let ((mid (/ (+ reader section) 2)))
    `(splice ,mid (bisect ,reader ,mid))))

;;; ---------------------------------------------------------------
;;; 2. The same macro, indexed by an ordinal instead of a rational.
;;;    From ω you may step to any natural number — that step is a
;;;    choice — but after it, every step is forced, and 0 has no
;;;    predecessor in ω. Termination is a theorem, not a fiat.
;;; ---------------------------------------------------------------

(defmacro countdown (n)
  (cond ((eq n 'ω) '(countdown 5))            ; the leap: unforced, marked
        ((zerop n) '(reader))                  ; the least element: occupied
        (t `(splice ,n (countdown ,(1- n)))))) ; the descent: forced

;;; ---------------------------------------------------------------
;;; A full expander with a budget. Walks the whole tree, expanding
;;; every macro call it finds, until either no macro calls remain
;;; (termination) or the budget is spent (divergence).
;;; ---------------------------------------------------------------

(defun expand-all (form)
  (cond ((not (consp form)) form)
        ((macro-function (car form))
         (when (zerop *fuel*)
           (throw 'out-of-fuel form))
         (decf *fuel*)
         (incf *steps*)
         (expand-all (macroexpand-1 form)))
        (t (cons (car form) (mapcar #'expand-all (cdr form))))))

(defun run (form fuel)
  (setf *fuel* fuel *steps* 0)
  (format t "~&> (macroexpand '~S)  ; budget ~D~%" form fuel)
  (let ((result (catch 'out-of-fuel (expand-all form))))
    (if (zerop *fuel*)
        (format t "⇒ DIVERGED after ~D expansions; last call site still pending: ~S~%~%"
                *steps* result)
        (format t "⇒ HALTED after ~D expansions:~%   ~S~%~%" *steps* result))))

(run '(bisect -2 -1) 8)
(run '(bisect -2 -1) 40)
(run '(countdown ω) 40)

;;; And the single-step expander the essay called its termination discipline:
(format t "> (macroexpand-1 '(bisect -2 -1))~%⇒ ~S~%" (macroexpand-1 '(bisect -2 -1)))
(format t "~&;; macroexpand-1 did not terminate the stairwell. It declined to begin it.~%")
