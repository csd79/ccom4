;;;; -*- Mode#: Common-Lisp; Author#: denes.cselovszky@gmail.com -*- 
                                                                             ;
(in-package #:cl-user)

(defpackage #:ccom3
  (:use #:cl #:achar)
  (:export #:?
           #:!
           #:property-accessors-on
           #:with-property-accessors
           #:cclet*
           #:with-com-initialized))
