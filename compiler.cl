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
	((equal (car exp) 'if) (concatenate 'string "if " (comp (cadr exp) env) " then:" (string #\newline)
			   "    " (comp (caddr exp) env) (string #\newline)
			   "else:" (string #\newline)
			   "    " (comp (cadddr exp) env) (string #\newline)
			   "end if"))
	((equal (car exp) 'progn) (list-comp (cdr exp) env)) 
	((equal (car exp) 'setf) (concatenate 'string "call MySet(" (comp (cadr exp) env) ", " (comp (caddr exp) env) ")"))
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
	  (list new-var new-env))
	(let* ((new-var (concatenate 'string old-var (string #\_) (format nil "~a" 1)))
	       (new-env (acons symb new-var env)))
	  (list new-var new-env)))))

(defun function-comp (exp env)
  (let ((func (car exp))
	(args (mapcar #'(lambda (x) (comp x env)) (cdr exp))))
    (if (find func *infix-operators*)
	(prefix-to-infix args (format nil "~a" func))
	(concatenate 'string (symbol-name func) (args-to-string args)))))

(defun list-comp (exp env)
  (let ((exps (mapcar #'(lambda (x) (comp x env)) exp)))
    (reduce #'(lambda (x y) (concatenate 'string x (string #\newline) y)) exps)))

(defun body-comp (exp env)
  (if (cdr exp)
      (list-comp exp env)
      (comp (car exp) env)))

(defun defun-comp (exp env)
  (let ((name (comp (cadr exp) env))
	(args (mapcar #'(lambda (x) (comp x env)) (caddr exp)))
	(body (body-comp (butlast (cdddr exp)) env))
	(ret-val (comp (car (last exp)) env)))
    (concatenate 'string
		 "public function " name (args-to-string args) (string #\newline)
		 body (string #\newline)
		 name " = " ret-val (string #\newline)
		 "end function")))

(defmacro defmacro-comp (exp env)
  (let ((name (cadr exp))
	(args (caddr exp))
	(body (cdddr exp)))
    `(setf *macros* (acons (quote ,name) #'(lambda ,args .,body) *macros*))))

(defun macro-expand (macro-fun exp)
  (loop with lst = (list macro-fun)
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
  (let* ((new-var-lst (new-symbol-comp (cadr exp) env))
	 (var (car new-var-lst))
	 (new-env (cadr new-var-lst))
	 (start (comp (caddr exp) new-env))
	 (end (comp (cadddr exp) new-env))
	 (body (body-comp (cddddr exp) new-env)))
    (concatenate 'string "for " var " = " start " to " end (string #\newline)
		 body (string #\newline)
		 "next " var)))

(defun outer-comp (exp)
  (comp exp '()))

(eval-when (:execute)
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
  (print (outer-comp '(when (> x 10) (print "big")))))

(eval-when (:execute)
  (setf *builtins* '())
  (setf *builtins* (acons 'defun (symbol-function 'defun-comp) *builtins*))
  (setf *builtins* (acons 'for (symbol-function 'for-comp) *builtins*))
  (setf *builtins* (acons 'while (symbol-function 'while-comp) *builtins*))
  (setf *builtins* (acons 'defmac (symbol-function 'defmacro-comp) *builtins*)))
