;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom4)


#|

COMPARISON:


VBA:
Workbooks = Global.Workbooks
Workbook  = Workbooks.Add(template := xlWorksheet)
Worksheet = Workbook.Worksheets(1)
Print(Worksheet.Range("A1"))
Worksheet.Name = "New Name"


LispWorks:
(com:with-temp-interface (workbook)
    (com:with-temp-interface (workbooks)
        (com:invoke-dispatch-get-property global "Workbooks")
      (com:invoke-dispatch-method workbooks "Add" +xl-worksheet+))
  (com:with-temp-interface (worksheet)
      (com:with-temp-interface (sheets)
          (com:invoke-dispatch-get-property workbook "Worksheets")
        (com:invoke-dispatch-get-property sheets "Item" 1))
    (let ((address "A1"))
      (get-value worksheet address))
    (setf (com:invoke-dispatch-get-property worksheet "Name") "New name")))


CCOM4:
#.(enable-ccom-syntax)
(cclet* ((workbooks (?'workbooks global))
         (workbook  (!'add workbooks +xl-worksheet+))
         (worksheet (?'item (?worksheets workbook) 1)))
  (print (?'value2 (?'range worksheet "A1")))
  (setf (?'name worksheet) "New name"))
#.(disable-ccom-syntax)








Accessors can be passed as arguments:

#.(enable-ccom-syntax)
(defun action (accessor ws)
  (setf (?accessor (?'range ws "C3")) "Juhhéééé!")
  (print (?accessor (?'range ws "B2"))))

(defun main (file)
  (with-ccom-context
    (cclet* ((excel  (?'application (com:create-object :progid "Excel.Application")))
             (wbook  (!'open (?'workbooks excel) file nil nil))
             (wsheet (?'item (?'worksheets wbook) 1)))
      (action 'value2 wsheet)
      (!'save wbook)
      (!'close wbook)
      (!'quit excel))))
#.(disable-ccom-syntax)


The ! and ? macro characters only work as prefixes to the first element of a complex form:

(enable-ccom-syntax)
(print (read-from-string "!'method"))
(disable-ccom-syntax))
=> !

|#





;;; ----------------------------------------------------------------------
;;; COM property and method calls


(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun intern-by-name (name package)
    "Create symbol named NAME in PACKAGE and return it."
    (let ((uppercase-name (string-upcase name)))
      (shadow uppercase-name package)
      (find-symbol uppercase-name package)))


  (defun dispatch-function (char)
    "Return the name of the dispatch function designated by CHAR."
    (case char
      (#\! 'com:invoke-dispatch-method)
      (#\? 'com:invoke-dispatch-get-property)
      (otherwise nil)))


  (defun dispatch-expr (design-char object operator args)
    "Build a dispatch expression from rearranged parts of the input expression."
    (append (list (dispatch-function design-char)
                  object `(symbol-name ,operator))
            args))


  (defun com-fns (stream char)
    "Transform expressions containing ! or ? into dispatch expressions."
    (declare (ignore char))
    (let ((expr (read-delimited-list #\) stream t)))
      (if (and (listp expr)
               (>= (length expr) 2)
               (symbolp (first expr)))
        ;; Complex expression with at least two elements,
        ;;   the first element being a symbol:
        ;;   might be a method call / prop access.
        (destructuring-bind (a b &optional c &rest args) expr
          (let ((op-name (symbol-name a)))
            (cond (;; !SYMBOL
                   (and (> (length op-name) 1)
                        (member (char op-name 0) '(#\! #\?) :test #'char=))
                   (dispatch-expr (char op-name 0) b
                                  (intern-by-name (subseq op-name 1) *package*)
;                                  (cons c args)))
                                  (if c (cons c args) args)))
                  (;; !'SYMBOL
                   (and (member op-name '("!" "?") :test #'string=)
                        (listp b)
                        (eq (first b) 'quote))
                   (dispatch-expr (char op-name 0) c b args))
                  (;; Otherwise, return the expression.
                   t expr))))
        ;; Simple expression or a list of <=1 element,
        ;;   not a method call / prop access.
        expr)))


  (defparameter *previous-readtable* nil)

  
  (defun enable-ccom-syntax ()
    "To enable ? & ! syntax, insert this code: #.(ENABLE-?!)"
    (setf *previous-readtable* *readtable*
          *readtable* (copy-readtable))
    (set-macro-character #\( #'com-fns))

  
  (defun disable-ccom-syntax ()
    "To disable ? & ! syntax, insert this code: #.(DISABLE-?!)"
    (setf *readtable* *previous-readtable*)))
  

(defparameter *com-init-count* 0
  "Is COM initialized?")


(defmacro with-ccom-context (&body body)
  "Binding context for COM operations."
  `(let ((orig *com-init-count*))
     (unwind-protect 
         (progn
           (when (zerop *com-init-count*)
             (com::co-initialize))
           (incf *com-init-count*)
           ,@body)
       (progn
         (setf *com-init-count* orig)
         (when (zerop *com-init-count*)
           (com::co-uninitialize))))))


;;; ----------------------------------------------------------------------
;;; Binding interface pointers


(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun live-iptr-p (thing)
    "Is THING an interface pointer?"
    (and thing
         (not (eq thing :nothing))
         (typep thing 'com::com-interface)))


  (defun ensure-refcount (iptr &optional (safe 10))
    "Ensure ref counter for IPTR is set to SAFE."
    (let* ((current (com:add-ref iptr))
           (difference (abs (- safe current)))
           (stepper (if (< current safe) #'com:add-ref #'com:release)))
      (unless (zerop difference)
        (dotimes (i (1- difference))
          (funcall stepper iptr))
        (funcall stepper iptr))))


  (defun cclet-addref (vars vals body)
    "Building binding clauses for CCLET*."
    (let ((var (first vars))
           (val (first vals)))
      (print (get-macro-character #\())
      `((let ((,var ,val))
          (when (live-iptr-p ,var)
            (ensure-refcount ,var))
          ,@(if (second vars)
              (cclet-addref (rest vars) (rest vals) body)
              body))))))


(defmacro cclet* (bindings &body body)
  "Local binding context for both interface pointers and plain Lisp values."
  (if bindings
    (let* ((vars (mapcar #'first bindings))
           (vals (mapcar #'second bindings)))
      `(progn
         ,@(cclet-addref
            vars vals
            `((unwind-protect
                  (progn
                    ,@body)
                (mapc #'(lambda (var)
                          (when (live-iptr-p var)
                            (ensure-refcount var 0)))
                      ',(reverse vars)))))))
    `(progn ,@body)))
