;;;; -*- Mode#: Common-Lisp; Author#: denes.cselovszky@gmail.com -*- 
                                                                             ;
(in-package #:cl-user)

(defpackage #:ccom4
  (:use #:cl #:achar)
  (:export #:live-iptr-p
           #:enable-ccom-syntax
           #:disable-ccom-syntax
           #:with-ccom-context
           #:cclet*))
