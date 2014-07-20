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
	((equal (car exp) 'def) (def-comp exp env))
	((equal (car exp) 'set!) (set-comp exp env))
	((equal (car exp) 'fn) (lambda-comp exp env))
	;((equal (car exp) 'let) 
	;((assoc (car exp) *builtins*) (funcall (cdr (assoc (car exp) *builtins*)) exp env))
	((equal (car exp) 'defmac) (defmacro-comp exp env))
	((assoc (car exp) *macros*) (comp (macro-expand (cdr (assoc (car exp) *macros*)) exp) env))
	(t (funcall-comp exp env)))))

(defun atom-comp (exp env)
  (cond ((symbolp exp)
	 ;(string exp))
	 (concatenate 'string "Lookup(\"" (string exp) "\", Env)"))
	((stringp exp)
	 (concatenate 'string "Create(\"" exp ", \"string\")"))
	((numberp exp)
	 (concatenate 'string "Create(" (format nil "~a" exp) ", \"number\")"))
	(t (format nil "~a" exp))))

(defun if-comp (exp env)
	(multiple-value-let (((ret-val-name new-env) (new-symbol-comp 'retval env))
			     ((test-val test-body) (comp (cadr exp) new-env))
			     ((if-val if-body) (comp (caddr exp) new-env))
			     ((else-val else-body) (comp (cadddr exp) new-env)))
	  (values ret-val-name
		  (concatenate 'string test-body (string #\newline)
			       "if " test-val " then:" (string #\newline)
			       if-body (string #\newline)
			       "Lookup(\"" ret-val-name "\", Env) = (" if-val ")" (string #\newline)
			       "else " (string #\newline)
			       else-body (string #\newline)
			       "Lookup(\"" ret-val-name "\", Env) = (" else-val ")" (string #\newline)
			       "end if" ))))

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

(defun list-comp (exp env)
  (let ((exps (mapcar #'(lambda (x) (nth-value 1 (comp x env))) exp)))
    (reduce #'(lambda (x y) (concatenate 'string x (string #\newline) y)) exps)))

(defun progn-comp (exp env)
  (if (cdr exp)
      (values (comp (car (last exp)) env)
	      (list-comp (butlast exp) env))
      (comp (car exp) env)))

(defun def-comp (exp env)
  (multiple-value-let (((ret-val val-body) (comp (caddr exp) env)))
		       ;((symb symb-body) (comp (cadr exp) env)))
      (let ((body (concatenate 'string 
			       val-body (string #\newline)
			       ;symb-body (string #\newline)
			       "set GlobalEnv.Item(\"" (string (cadr exp)) "\") = (" ret-val ")")))
	(values ret-val body))))

(defun set-comp (exp env)
  (multiple-value-let (((ret-val val-body) (comp (caddr exp) env)))
		       ;((symb symb-body) (comp (cadr exp) env)))
      (let ((body (concatenate 'string 
			       val-body (string #\newline)
			       ;symb-body (string #\newline)
			       "set Env.Item(\"" (string (cadr exp)) "\") = (" ret-val ")")))
	(values ret-val body))))

(defun lambda-comp (exp env)
  ;General plan:  At the location where the lambda is called, add some code
  ;to save the variables that will be captured, and return a unique ID for the 
  ;anonymous function instance.  Then add the code that will actually be executed
  ;when that function is called to a list of functions to be added to the end of 
  ;the compiled code later
  (let ((lambda-args (cadr exp))
	(id (concatenate 'string "lambda_" (format nil "~a" (incf *lambda-counter*)))))
    (let ((lambda-body (function-comp id lambda-args (cddr exp)))
	  (ret-val (concatenate 'string "Create(\"" id "_\" & Lambda_Counter, \"func\")")))
      ;(print "lambda-comp")
      ;(print exp)
      ;(print lambda-args)
      (push lambda-body *lambda-bodies*)
      (values 
       ret-val
       (concatenate 'string 
		    "Dim " id " As New Collection" (string #\newline)
		    "Call " id ".Add(\"" ID "\")" (string #\newline)
		    "Dim Env_" id " as New Scripting.Dictionary" (string #\newline)
		    "set Env_" id " = Env" (string #\newline)
		    "Call " id ".Add(Env_" id ")" (string #\newline)
		    "Lambda_Counter = Lambda_Counter + 1" (string #\newline)
		    "Set GlobalEnv.item(\"" id "_\" & Lambda_Counter) =  " id (string #\newline))))))

(defun function-comp (id formal-args body)
  (let ((prelude (concatenate 'string 
			      "instance = args.item(1).getval" (string #\newline)
			      "Dim Env as New Scripting.Dictionary" (string #\newline)
			      "Dim LR as New Collection" (string #\newline)
			      "Set LR = Lookup(instance, Env)" (string #\newline)
			      "set Env = LR.item(2)"))
	(formal-args-list (loop for i = 2 then (+ i 1)
			       for arg in formal-args
			       collect (concatenate 'string "set Env.Item(\"" (string arg) "\")"
						       " = args.item(" (int-string i) ")")))
	(writeback (concatenate 'string 
				"Dim LR_writeback as New Collection" (string #\newline)
				"Call LR_writeback.add(LR.item(1))" (string #\newline)
				"Call LR_writeback.add(Env)" (string #\newline)
				"set GlobalEnv.item(instance) = LR_writeback")))
    (multiple-value-bind (ret-val body-text) (progn-comp body '())
      (concatenate 'string "Public Function " id "(args)" (string #\newline)
		   prelude (string #\newline)
		   (add-newlines formal-args-list) (string #\newline)
		   body-text (string #\newline)
		   writeback (string #\newline)
		   "set current_res = " ret-val (string #\newline)
		   "End Function"))))

(defun find-captured-vars (exp args)
  ;simply returns a list of every symbol in exp that isn't in args.  Simply accepting
  ;the inefficiency caused by capturing every function name in every lambda, even builtins etc
  (if (atom exp)
      (if (and (symbolp exp)
	       (not (find exp args)))
	  (list exp))
      (remove-duplicates
       (apply #'append (map 'list #'(lambda (x) (find-captured-vars x args)) exp)))))

(defun defmacro-comp (exp env)
  (let ((name (cadr exp))
	(args (caddr exp))
	(body (cdddr exp)))
    (eval `(progn (setf *macros* (acons (quote ,name) #'(lambda ,args .,body) *macros*))
		  (values "name" "")))))

(defun macro-expand (macro-fun exp)
  ;(print "macro expand")
  ;(print exp)   
  (let ((res (loop ;with lst = (list macro-fun)
		for sub-exp in (cdr exp)
		collecting `(quote ,sub-exp) into sub-exps
		finally (return (eval (concatenate 'list `(funcall ,macro-fun) sub-exps))))))
    ;(print res)
    res))

(defun funcall-comp (exp env)
  (loop for sub-exp in exp
     for lst = (multiple-value-list (comp sub-exp env))
     collecting (car lst) into ret-vals
     collecting (cadr lst) into bodies
     finally (return (values (concatenate 'string "funcall" (args-to-string ret-vals))
			     (add-newlines bodies)))))


;possibly dead code
(defun defun-comp (exp env)
  (let ((name (comp (cadr exp) env))
	(args (mapcar #'(lambda (x) (comp x env)) (caddr exp))))
    (multiple-value-bind (ret-val body) (progn-comp (cdddr exp) env)
      (concatenate 'string
		   "public function " name (args-to-string args) (string #\newline)
		   body (string #\newline)
		   name " = " ret-val (string #\newline)
		   "end function"))))

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

;utility functions
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
(defun outer-comp (exp)
  (setf *lambda-bodies* '())
  (multiple-value-bind (ret-val body) (comp exp '())
    (concatenate 'string 
		 (string #\newline)
		 "public sub main()" (string #\newline)
		 "call BuiltinInit()" (string #\newline)
		 "dim Env as new Scripting.Dictionary" (string #\newline)
		 body (string #\newline)
		 ret-val (string #\newline)
		 "end sub" (string #\newline)
		 (string #\newline)
		 "'lambda functions" (string #\newline)
		 (string #\newline)
		 (add-newlines *lambda-bodies*))))

(defun prefix-to-infix (lst token)
  ;inserts token between each pair of elements in lst, then returns the result as a string
  (concatenate 'string
	       (reduce #'(lambda (x y) (concatenate 'string (string x) " " token " " (string y))) lst)))

(defun add-newlines (lst)
  (if lst
      (prefix-to-infix lst (string #\newline))
      " "))

(defun args-to-string (lst)
  ;(let ((eval-lst (map 'list 
;		       #'(lambda (x) (concatenate 'string "Env.item(\"" (string x) "\")"))
;		       lst)))
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

(defmacro let-test (bindings &rest body)
  (loop for binding in bindings
     collecting (car binding) into vars
     collecting (cadr binding) into vals
     finally (return (concatenate 'list '(fn) (list vars) body vals))))

(eval-when (:execute)
  (setf *macros* '())
  (defmacro-comp '(defmac let (bindings &rest body)
			    (loop for binding in bindings
			       collecting (car binding) into vars
			       collecting (cadr binding) into vals
			       finally (return `((fn ,vars ,@body) ,@vals))))
				 ;(return (concatenate 'list '(fn) (list vars) body vals))))
      '())
  (defmacro-comp '(defmac defn (name args &rest body)
		   `(def ,name (fn ,args ,@body)))
      '())
  (format t "~A" (outer-comp '(defn accum (a)
			       (fn (x) (set! a (+ x a))
				a))))
  (format t "~A" (outer-comp '(let ((fun (accum 4)))
			       (fun 1)))))

(eval-when (:execute)
  (setf *builtins* '())
  (setf *builtins* (acons 'defun (symbol-function 'defun-comp) *builtins*))
  (setf *builtins* (acons 'defmac (symbol-function 'defmacro-comp) *builtins*))
  (setf *builtins* (acons 'for (symbol-function 'for-comp) *builtins*))
  (setf *builtins* (acons 'while (symbol-function 'while-comp) *builtins*))
  (setf *builtins* (acons 'let (symbol-function 'let-comp) *builtins*))
  (setf *macros* '())
  (defmacro-comp '(defmac let (bindings &rest body)
			    (loop for binding in bindings
			       collecting (car binding) into vars
			       collecting (cadr binding) into vals
			       finally (return `((fn ,vars ,body) ,vals))))
				 ;(return (concatenate 'list '(fn) (list vars) body vals))))
      '())
  (defmacro-comp '(defmac defn (name args &rest body)
		   `(def ,name (fn ,args ,@body)))
      '())
  (print (outer-comp '(* 1 2 3 4)))
  (print (outer-comp '(+ 1 2 (* 3 4))))
  (print (outer-comp '(plus 1 2 3 4 )))
  (print (outer-comp '(if (= a b) (+ 1 2) (* 3 4))))
  (print (outer-comp '(progn (+ 1 2 3) (times 3 4))))
  (print (outer-comp '(setf a 3)))
  (print (outer-comp '(defn plus (x y)
		 (+ x y))))
  (print (outer-comp '(for i 0 10 
		       (for i 0 10
			(msgbox i))
		       (msgbox "outer loop")
		       (msgbox i))))
  (outer-comp '(defmac when (condition &rest body)
		`(if ,condition (progn ,@body))))
  (print (outer-comp '(when (> x 10) (print "big"))))
  (print "last function")
  (format t "~A" (outer-comp '(defn DSA-gen-keys (p q g)
		       (let ((priv-key (random q (make-random-state t))))
			 (let ((pub-key (modexp g priv-key p)))
			   (values priv-key pub-key)))))))
