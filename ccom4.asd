(require "automation")

(defsystem "ccom4"
  :description "Extra syntax for the LispWorks COM interface, v4"
  :author      "Denes Cselovszki <denes.cselovszki@gmail.com>"
  :version     "0.28"
  :depends-on  ("achar" )
  :serial      t
  :components  ((:file "package")
                (:file "ccom4")))
