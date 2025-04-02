;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom3)


#|

VBA:
Workbooks = Global.Workbooks
Workbook  = Workbooks.Add(template := xlWorksheet)
Worksheet = Workbook.Worksheets(1)
Print(Worksheet.Range("A1"))
Worksheet.Name = "New Name"


Lisp:
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


Lisp & CCOM:
(cclet* ((workbooks (?workbooks global))
         (workbook  (!add workbooks +xl-worksheet+))
         (worksheet (?item (?worksheets workbook) 1))
         (address   "A1"))
  (print (?value2 (?range worksheet address)))
  (setf (?name worksheet) "New name"))

|#


;;; ----------------------------------------------------------------------
;;; Utilities


(defparameter *com-init-count* 0
  "Is COM initialized?")


(defmacro with-com-initialized (&body body)
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
;;; COM property and method calls


;;; COM 'get property' read macro.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defparameter *ccom-accessor-prefix* "CCOM-ACCESSOR-"
    "Prefix to use for property accessor function names.")

  
  (defun intern-accessor (property-name package)
    "Create & return prefixed accessor fn symbol for property in PACKAGE."
    (let ((symbol-name
           (string-upcase
            (concatenate 'string *ccom-accessor-prefix* property-name))))
      (shadow symbol-name package)
      (find-symbol symbol-name package)))

  
  (defun decompose-lw-setf-symbol (string)
    "Extract package & setf-symbol from string provided by CELL-ERROR-NAME."
    (let* ((start 0)
           (quotes (loop for i from 0 below 4
                         for new = (position #\" string
                                    :test #'char= :start start)
                         when new do (setf start (1+ new))
                         when new collect new)))
      (when (= (length quotes) 4)
        (list (subseq string (1+ (first quotes)) (second quotes))
              (subseq string (1+ (third quotes)) (fourth quotes))))))

  
  (defun error-symbol (error)
    "List info on symbol in ERROR: (package symbol-name 'setf|nil)."
    (let ((datum (cell-error-name error)))
      (cond ((symbolp datum)
             (let ((symbol-name (symbol-name datum)))
               (if (find #\" symbol-name :test #'char=)
                 ;; 'SETF::\"CCOM\"\ \"VALUE2\"
                 (append (decompose-lw-setf-symbol symbol-name)
                         (list 'setf))
                 ;; 'VALUE2
                 (list (package-name (symbol-package datum))
                       (symbol-name datum)
                       nil))))
            ;; '(SETF VALUE2)
            ((and (listp datum)
                  (= (length datum) 2)
                  (eql (first datum) 'setf))
             (list (package-name (symbol-package (second datum)))
                   (symbol-name (second datum))
                   'setf))
            (t (list nil nil)))))

  
  ;; Read macro to add property accessor prefix to the following symbol.
  (set-macro-character 
   #\?
   (lambda (stream char)
     (declare (ignore char))
     (let* ((prop      (read stream t nil t))
            (prop-name (symbol-name prop)))
       (intern-accessor prop-name *package*)))))


(defparameter *property-accessors-on* nil
  "Should undefined COM prop accessor be created on the fly?")


(defun property-accessors-on ()
  "Get value of *PROPERTY-ACCESSORS-ON*."
  *property-accessors-on*)


(defun (setf property-accessors-on) (value)
  "Set value of *PROPERTY-ACCESSORS-ON*."
  (setf *property-accessors-on* (not (not value))))


(defmacro with-property-accessors (&body body)
  "Define undefined prop accessors in BODY if *PROPERTY-ACCESSORS-ON* = T."
  `(handler-bind
       ((undefined-function
         #'(lambda (error)
             (when *property-accessors-on*
               (destructuring-bind (package-name symbol-name type)
                   (error-symbol error)
                 (when (and symbol-name
                            (> (length symbol-name) 15)
                            (string= (subseq symbol-name 0 14)
                                     *ccom-accessor-prefix*))
                   (let* ((propname (subseq symbol-name 14))
                          (propset  (concatenate 'string "SET-" propname))
                          (getsym   (intern-accessor propname package-name))
                          (setsym   (intern-accessor propset package-name))
                          (getfn    #'(lambda (obj &rest args)
                                        (apply
                                         #'com:invoke-dispatch-get-property
                                         obj propname args)))
                          (setfn    #'(lambda (value obj &rest args)
                                        (apply
                                         #'com:invoke-dispatch-put-property
                                         obj propname
                                         (nconc args (list value))))))
                     (setf (fdefinition getsym) getfn
                           (fdefinition setsym) setfn
                           (fdefinition `(setf ,getsym))
                           (fdefinition setsym))
                     (invoke-restart (find-restart 'use-value error)
                                     (if (eq type 'setf)
                                       setfn
                                       getfn)))))))))
     ,@body))


;; COM 'method call' read macro.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (set-macro-character
   #\!
   (lambda (stream char)
     (declare (ignore char))
     (let ((method (read stream t nil t)))
       `(lambda (obj &rest rest)
          (apply #'com:invoke-dispatch-method
                 obj (symbol-name ',method) rest))))))


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
      `((let ((,var ,val))
          (when (live-iptr-p ,var)
            (ensure-refcount ,var))
          ,@(if (second vars)
              (cclet-addref (rest vars) (rest vals) body)
              body))))))


#|(defmacro cclet* (bindings &body body)
  "Local binding context for both interface pointers and plain Lisp values."
  (if bindings
    (let ((vars (mapcar #'first bindings)))
      `(with-com-initialized
         (let* ,bindings
           (flet ((do-iptrs (fn &optional (order-fn #'identity))
                    (dolist (var (funcall order-fn ',vars))
                      (when (live-iptr-p var)
                        (funcall fn var)))))
             (unwind-protect
                 (progn
                   ,@body)
               (do-iptrs #'com:release #'reverse))))))
    `(progn
       ,@body)))|#
#|(defmacro cclet* (bindings &body body)
  "Local binding context for both interface pointers and plain Lisp values."
  (if bindings
    (let ((vars (mapcar #'first bindings)))
      `(with-com-initialized
         (let* ,bindings
           (flet ((do-iptrs (&key (refcount 10) (order-fn #'identity))
                    (dolist (var (funcall order-fn ',vars))
                      (when (live-iptr-p var)
                        (ensure-refcount var refcount)))))
             (unwind-protect
                 (progn
                   (do-iptrs)
                   ,@body)
               (do-iptrs :refcount 0 :order-fn #'reverse))))))
    `(progn
       ,@body)))|#
#|(defmacro cclet* (bindings &body body)
  "Local binding context for both interface pointers and plain Lisp values."
  (if bindings
    (let* ((vars    (mapcar #'first bindings))
           (vals    (mapcar #'second bindings))
           (empties (mapcar #'(lambda (var)
                                (list var nil))
                            vars)))
      `(with-com-initialized 
         (let* ,empties
           (unwind-protect
               (progn
                 (mapc #'(lambda (var val)
                           (setf (symbol-value var) val)
                           (when (live-iptr-p (symbol-value var))
                             (ensure-refcount (symbol-value var))))
                       ',vars (list ,@vals))
                 ,@body)
             (mapc #'(lambda (var)
                       (when (live-iptr-p (symbol-value var))
                         (ensure-refcount (symbol-value var) 0)))
                   ',(reverse vars))))))
    `(progn ,@body)))|#
(defmacro cclet* (bindings &body body)
  "Local binding context for both interface pointers and plain Lisp values."
  (if bindings
    (let* ((vars (mapcar #'first bindings))
           (vals (mapcar #'second bindings)))
      `(with-com-initialized 
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
