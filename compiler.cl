;Compiler for Kit, a toy lisp-like language.  Will eventually compile to VBA.  

;Key insight for the compiler:  a compiler is like an interpreter, but instead of actually
;doing the computation it simply outputs code which does instead

(defparameter *infix-operators* (list '+ '* '- '/ '=))
(setf *infix-operators* (list '+ '* '- '/ '=))

(defparameter *builtins* '())
(defparameter *macros* '())

(defun comp (exp env)
  (if (atom exp)
      (atom-comp exp env)
      (cond
	((equal (car exp) 'quote) (cadr exp))
	((equal (car exp) 'if) (if-comp exp env))
	((equal (car exp) 'progn) (progn-comp (cdr exp) env)) 
	;((equal (car exp) 'setf) (concatenate 'string "call MySet(" (comp (cadr exp) env) ", " (comp (caddr exp) env) ")"))
	((assoc (car exp) *builtins*) (funcall (cdr (assoc (car exp) *builtins*)) exp env))
	((assoc (car exp) *macros*) (comp (macro-expand (cdr (assoc (car exp) *macros*)) exp) env))
	(t (function-comp exp env)))))

(defun prefix-to-infix (lst token)
  ;inserts token between each pair of elements in lst, then returns the result as a string
  (concatenate 'string "("
	       (reduce #'(lambda (x y) (concatenate 'string (string x) " " token " " (string y))) lst)
	       ")"))

(defun args-to-string (lst)
  (concatenate 'string "("
	       (reduce #'(lambda (x y) (concatenate 'string (string x) ", " (string y))) lst)
	       ")"))

(defun atom-comp (exp env)
  (cond ((symbolp exp)
	 (if (assoc exp env :test #'equal)
	     (cdr (assoc exp env :test #'equal))
	     (format nil "~a" exp)))
	((stringp exp)
	 (concatenate 'string (string #\") exp (string #\")))
	(t (format nil "~a" exp))))

;replace aliases with a thing like variablename_number
(defun new-symbol-comp (symb env)
  (let* ((old-var (comp symb env))
	 (pos (position #\_ old-var :test #'equal)))
    (if pos
	(let* ((new-num (+ (parse-integer (subseq old-var (+ pos 1))) 1))
	       (new-var (concatenate 'string (subseq old-var 0 (+ pos 1)) (format nil "~a" new-num)))
	       (new-env (acons symb new-var env)))
	  (values new-var new-env))
	(let* ((new-var (concatenate 'string old-var (string #\_) (format nil "~a" 1)))
	       (new-env (acons symb new-var env)))
	  (values new-var new-env)))))

(defun set-comp (var val)
  (concatenate 'string "myset(" var ", " val ")"))

(defun if-comp (exp env)
	(multiple-value-let (((ret-val-name new-env) (new-symbol-comp 'retval env))
			     ((test-val test-body) (comp (cadr exp) new-env))
			     ((if-val if-body) (comp (caddr exp) new-env))
			     ((else-val else-body) (comp (cadddr exp) new-env)))
	  (values ret-val-name
		  (concatenate 'string test-body (string #\newline)
			       "if " test-val " then:" (string #\newline)
			       if-body (string #\newline)
			       (set-comp ret-val-name if-val)
			       "else " (string #\newline)
			       else-body (string #\newline)
			       (set-comp ret-val-name else-val) (string #\newline)
			       "end if" (string #\newline)))))

(defun function-comp (exp env)
  (let ((func (car exp))
	(args (mapcar #'(lambda (x) (comp x env)) (cdr exp))))
    (if (find func *infix-operators*)
	(prefix-to-infix args (format nil "~a" func))
	(concatenate 'string (symbol-name func) (args-to-string args)))))

(defun list-comp (exp env)
  (let ((exps (mapcar #'(lambda (x) (comp x env)) exp)))
    (reduce #'(lambda (x y) (concatenate 'string x (string #\newline) y)) exps)))

(defun progn-comp (exp env)
  (if (cdr exp)
      (values (comp (car (last exp)) env)
	      (list-comp (butlast exp) env))
      (comp (car exp) env)))

(defun defun-comp (exp env)
  (let ((name (comp (cadr exp) env))
	(args (mapcar #'(lambda (x) (comp x env)) (caddr exp))))
    (multiple-value-bind (ret-val body) (progn-comp (cdddr exp) env)
      (concatenate 'string
		   "public function " name (args-to-string args) (string #\newline)
		   body (string #\newline)
		   name " = " ret-val (string #\newline)
		   "end function"))))

(defun defmacro-comp (exp env)
  (let ((name (cadr exp))
	(args (caddr exp))
	(body (cdddr exp)))
    (eval `(setf *macros* (acons (quote ,name) #'(lambda ,args .,body) *macros*)))))

(defun macro-expand (macro-fun exp)
  (loop ;with lst = (list macro-fun)
     for sub-exp in (cdr exp)
     collecting `(quote ,sub-exp) into sub-exps
     finally (return (eval (concatenate 'list `(funcall ,macro-fun) sub-exps)))))

(defun while-comp (exp env)
  (let ((test (comp (cadr exp) env))
	(body (body-comp (cddr exp) env)))
    (concatenate 'string "while " test (string #\newline)
		 body (string #\newline)
		 "end while")))

(defun for-comp (exp env)
  (multiple-value-bind (var new-env) (new-symbol-comp (cadr exp) env)
    (let ((start (comp (caddr exp) new-env))
	  (end (comp (cadddr exp) new-env))
	  (body (body-comp (cddddr exp) new-env)))
      (concatenate 'string "for " var " = " start " to " end (string #\newline)
		   body (string #\newline)
		   "next " var))))

(defun let-comp (exp env)
  (multiple-value-bind (bindings new-env) (let-bindings (cadr exp) env)
    (multiple-value-bind (ret-val body) (progn-comp (cddr exp) new-env)
      (values ret-val
	      (concatenate 'string (prefix-to-infix bindings (string #\newline)) (string #\newline)
		   body)))))

(defun let-bindings (bindings env)
  (print "let-bindings")
  (print bindings)
  (let ((new-env env))
    (values
     (mapcar #'(lambda (x) (multiple-value-bind (var temp-env) (new-symbol-comp (car x) new-env)
			     (print var)
			     (setf new-env temp-env)
			     (concatenate 'string "myset(" var ", " (comp (cadr x) new-env) ")")))
	     bindings)
     new-env)))

(defmacro multiple-value-let (bindings &body body)
  (multiple-value-let-helper bindings body))

(defun multiple-value-let-helper (bindings body)
  (let ((var (caar bindings))
	(value-form (cadar bindings)))
    (if (cdr bindings)
	`(multiple-value-bind ,var ,value-form ,(multiple-value-let-helper (cdr bindings) body))
	(append `(multiple-value-bind ,var ,value-form)
		body))))

(defun outer-comp (exp)
  (multiple-value-bind (ret-val body) (comp exp '())
    (concatenate 'string body (string #\newline)
		 ret-val)))

(eval-when (:execute)
  (setf *builtins* '())
  (setf *builtins* (acons 'defun (symbol-function 'defun-comp) *builtins*))
  (setf *builtins* (acons 'defmac (symbol-function 'defmacro-comp) *builtins*))
  (setf *builtins* (acons 'for (symbol-function 'for-comp) *builtins*))
  (setf *builtins* (acons 'while (symbol-function 'while-comp) *builtins*))
  (setf *builtins* (acons 'let (symbol-function 'let-comp) *builtins*))
  (print (outer-comp '(* 1 2 3 4)))
  (print (outer-comp '(+ 1 2 (* 3 4))))
  (print (outer-comp '(plus 1 2 3 4 )))
  (print (outer-comp '(if (= a b) (+ 1 2) (* 3 4))))
  (print (outer-comp '(progn (+ 1 2 3) (times 3 4))))
  (print (outer-comp '(setf a 3)))
  (print (outer-comp '(defun plus (x y)
		 (+ x y))))
  (print (outer-comp '(for i 0 10 
		       (for i 0 10
			(msgbox i))
		       (msgbox "outer loop")
		       (msgbox i))))
  (outer-comp '(defmac when (condition &rest body)
		`(if ,condition (progn ,@body))))
  (print *macros*)
  (print (outer-comp '(when (> x 10) (print "big"))))
  (print (outer-comp '(defun DSA-gen-keys (p q g)
		       (let ((priv-key (random q (make-random-state t))))
			 (let ((pub-key (modexp g priv-key p)))
			   (values priv-key pub-key)))))))
