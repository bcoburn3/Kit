;Compiler for Kit, a toy lisp-like language.  Will eventually compile to VBA.  

;Key insight for the compiler:  a compiler is like an interpreter, but instead of actually
;doing the computation it simply outputs code which does instead

(defparameter *infix-operators* (list '+ '* '- '/ '=))
(setf *infix-operators* (list '+ '* '- '/ '=))

(defparameter *builtins* '())
(defparameter *macros* '())
(defparameter *lambda-bodies* '())
(defparameter *lambda-counter* 0)

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
  (concatenate 'string
	       (reduce #'(lambda (x y) (concatenate 'string (string x) " " token " " (string y))) lst)))

(defun add-newlines (lst)
  (prefix-to-infix lst #\newline))

(defun args-to-string (lst)
  (concatenate 'string "("
	       (reduce #'(lambda (x y) (concatenate 'string (string x) ", " (string y))) lst)
	       ")"))

(defun int-string (int)
  (format nil "~a" int))

(defun multiple-value-let-helper (bindings body)
  (let ((var (caar bindings))
	(value-form (cadar bindings)))
    (if (cdr bindings)
	`(multiple-value-bind ,var ,value-form ,(multiple-value-let-helper (cdr bindings) body))
	(append `(multiple-value-bind ,var ,value-form)
		body))))

(defmacro multiple-value-let (bindings &body body)
  (multiple-value-let-helper bindings body))

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
  ;dumb hack around the fact that VBA requires a "set a = b" if be is an object
  ;and "a = b" if b is not an object
  (concatenate 'string "if isobject(" val ")" (string #\newline)
	       "set " var " = " val (string #\newline)
	       "else" (string #\newline)
	       var " = " val (string #\newline)
	       "end if"))

(defun if-comp (exp env)
	(multiple-value-let (((ret-val-name new-env) (new-symbol-comp 'retval env))
			     ((test-val test-body) (comp (cadr exp) new-env))
			     ((if-val if-body) (comp (caddr exp) new-env))
			     ((else-val else-body) (comp (cadddr exp) new-env)))
	  (values ret-val-name
		  (concatenate 'string test-body (string #\newline)
			       "if " test-val " then:" (string #\newline)
			       if-body (string #\newline)
			       (set-comp ret-val-name if-val) (string #\newline)
			       "else " (string #\newline)
			       else-body (string #\newline)
			       (set-comp ret-val-name else-val) (string #\newline)
			       "end if" ))))

(defun function-comp (id formal-args captured-args body)
  (let ((prelude (concatenate 'string "instance = args.item(1)" #\newline
			      "Dim LR as New Scripting.Dictionary" #\newline
			      "Set LR = Lambda_Dict.Item(instance)" #\newline
			      "Dim env as New Scripting.Dictionary" #\newline
			      "set env = args.Item(2)"))
	(formal-args-list (loop for i = 3 then (+ i 1)
			       for arg in formal-args
			       collect (concatenate 'string "env.Item(\"" (string arg) "\")"
						       " = args.item(" (int-string i) "\")")))
	(captured-args-list (loop for arg in captured-args
				 collect (concatenate 'string "env.Item(\"" (string arg) "\")"
							 " = LR.Item(\"" (string arg) "\")")))
	(writeback-list (loop for arg in captured-args
			     collect (concatenate 'string "LR.Item(\"" (string arg) "\")"
						     " = env.Item(\"" (string arg) "\")"))))
    (multiple-value-bind (ret-val body-text) (progn-comp body '())
      (concatenate 'string "Public Function " id "(args)" #\newline
		   prelude #\newline
		   (add-newlines formal-args-list) #\newline
		   (add-newlines captured-args-list) #\newline
		   body-text #\newline
		   (add-newlines writeback-list) #\newline
		   ret-val #\newline
		   "End Function"))))

(defun lambda-comp (exp env)
  ;General plan:  At the location where the lambda is called, add some code
  ;to save the variables that will be captured, and return a unique ID for the 
  ;anonymous function instance.  Then add the code that will actually be executed
  ;when that function is called to a list of functions to be added to the end of 
  ;the compiled code later
  (let ((lambda-args (cadr exp))
	(id (concatenate 'string "lambda_" (format nil "~a" (incf *lambda-counter*)))))
    (let ((captured-vars (find-captured-vars (cddr exp) lambda-args)))
      (let ((lambda-body (function-comp id lambda-args captured-vars (cddr exp)))
	    (captured-saves 
	     (map 'list #'(lambda (x) (concatenate 'string 
						   "call " id ".add(\"" (string x)
						   "env.item(\"" (string x) "\"))"))
		  lambda-args))
	    (ret-val (concatenate 'string "(\"" id "\" & Lambda_Counter)")))
	(push lambda-body *lambda-bodies*)
	(values 
	 ret-val
	 (concatenate 'string "Dim " id " As New Scripting.Dictionary" #\newline
			     "Call " id ".Add(\"FuncID\", \"" ID "\")" #\newline
			     (prefix-to-infix captured-saves #\newline)
			     "Lambda_Counter = Lambda_Counter + 1" #\newline
			     "Call LambdaDict.Add((\"" id "\" & Lambda_Counter), " id ")" #\newline))))))

(defun find-captured-vars (exp args)
  ;simply returns a list of every symbol in exp that isn't in args.  Simply accepting
  ;the inefficiency caused by capturing every function name in every lambda, even builtins etc
  (if (atom exp)
      (if (and (symbolp exp)
	       (not (find exp args)))
	  (list exp))
      (remove-duplicates
       (apply #'append (map 'list #'(lambda (x) (find-captured-vars x args)) exp)))))

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
    (eval `(progn (setf *macros* (acons (quote ,name) #'(lambda ,args .,body) *macros*))
		  (values "name" "")))))

(defun macro-expand (macro-fun exp)
  (loop ;with lst = (list macro-fun)
     for sub-exp in (cdr exp)
     collecting `(quote ,sub-exp) into sub-exps
     finally (return (eval (concatenate 'list `(funcall ,macro-fun) sub-exps)))))

(defun while-comp (exp env)
  (let ((test (comp (cadr exp) env))
	(body (progn-comp (cddr exp) env)))
    (concatenate 'string "while " test (string #\newline)
		 body (string #\newline)
		 "end while")))

(defun for-comp (exp env)
  (multiple-value-bind (var new-env) (new-symbol-comp (cadr exp) env)
    (let ((start (comp (caddr exp) new-env))
	  (end (comp (cadddr exp) new-env))
	  (body (progn-comp (cddddr exp) new-env)))
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
  (let ((new-env env))
    (values
     (mapcar #'(lambda (x) (multiple-value-bind (var temp-env) (new-symbol-comp (car x) new-env)
			     (setf new-env temp-env)
			     (concatenate 'string "myset(" var ", " (comp (cadr x) new-env) ")")))
	     bindings)
     new-env)))

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
