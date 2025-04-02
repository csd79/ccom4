(require "automation")

(defsystem "ccom3"
  :description "LispWorks 7.1 COM, but more lispy? v3"
  :author      "Denes Cselovszki <denes.cselovszki@gmail.com>"
  :version     "0.24"
  :depends-on  ("achar" )
  :serial      t
  :components  ((:file "package")
                (:file "ccom3")))
