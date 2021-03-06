;Compiler for Kit, a toy lisp-like language.  Will eventually compile to VBA.  

;Key insight for the compiler:  a compiler is like an interpreter, but instead of actually
;doing the computation it simply outputs code which does instead

(defun multiple-value-let-helper (bindings body)
  (let ((var (caar bindings))
	(value-form (cadar bindings)))
    (if (cdr bindings)
	`(multiple-value-bind ,var ,value-form ,(multiple-value-let-helper (cdr bindings) body))
	(append `(multiple-value-bind ,var ,value-form)
		body))))

(defmacro multiple-value-let (bindings &body body)
  (multiple-value-let-helper bindings body))

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
	((equal (car exp) 'defmac) (defmacro-comp exp env))
	((assoc (car exp) *macros*) (comp (macro-expand (cdr (assoc (car exp) *macros*)) exp) env))
	((equal (car exp) 'quote) (quote-comp exp env))
	(t (funcall-comp exp env)))))

(defun atom-comp (exp env)
  (cond ((symbolp exp)
	 (cat "Lookup(\"" (string exp) "\", Env)"))
	((stringp exp)
	 (cat "Create(\"" exp ", \"string\")"))
	((numberp exp)
	 (cat "Create(" (format nil "~a" exp) ", \"number\")"))
	(t (format nil "~a" exp))))

(defun if-comp (exp env)
	(multiple-value-let (((ret-val-name new-env) (values "CurrentRes" env))
			     ((test-val test-body) (comp (cadr exp) new-env))
			     ((if-val if-body) (comp (caddr exp) new-env))
			     ((else-val else-body) (comp (cadddr exp) new-env)))
	  (values ret-val-name
		  (cat test-body (string #\newline)
		       "if " test-val " then:" (string #\newline)
		       if-body (string #\newline)
		       "Lookup(\"" ret-val-name "\", Env) = (" if-val ")" (string #\newline)
		       "else " (string #\newline)
		       else-body (string #\newline)
		       "Lookup(\"" ret-val-name "\", Env) = (" else-val ")" (string #\newline)
		       "end if" (string #\newline)))))

(defun list-comp (exp env)
  (let ((exps (mapcar #'(lambda (x) (nth-value 1 (comp x env))) exp)))
    (reduce #'(lambda (x y) (cat x (string #\newline) y)) exps)))

(defun progn-comp (exp env)
  (if (cdr exp)
      (values (comp (car (last exp)) env)
	      (list-comp (butlast exp) env))
      (comp (car exp) env)))

(defun def-comp (exp env)
  (multiple-value-let (((ret-val val-body) (comp (caddr exp) env)))
      (let ((body (cat val-body 
			"set GlobalEnv.Item(\"" (string (cadr exp)) "\") = (" ret-val ")")))
	(values ret-val body))))

(defun set-comp (exp env)
  (multiple-value-let (((ret-val val-body) (comp (caddr exp) env)))
      (let ((body (cat val-body (string #\newline)
			"set Env.Item(\"" (string (cadr exp)) "\") = (" ret-val ")")))
	(values ret-val body))))

(defun lambda-comp (exp env)
  ;General plan:  At the location where the lambda is called, add some code
  ;to save the variables that will be captured, and return a unique ID for the 
  ;anonymous function instance.  Then add the code that will actually be executed
  ;when that function is called to a list of functions to be added to the end of 
  ;the compiled code later
  (let ((lambda-args (cadr exp))
	(id (cat "lambda_" (format nil "~a" (incf *lambda-counter*)))))
    (let ((lambda-body (function-comp id lambda-args (cddr exp)))
	  (ret-val (cat "Create(\"" id "_\" & Lambda_Counter, \"func\")")))
      ;(print "lambda-comp")
      ;(print exp)
      ;(print lambda-args)
      (push lambda-body *lambda-bodies*)
      (values 
       ret-val
       (cat "Dim " id " As New Collection" (string #\newline)
	    "Call " id ".Add(\"" ID "\")" (string #\newline)
	    "Dim Env_" id " as New Scripting.Dictionary" (string #\newline)
	    "set Env_" id " = Env" (string #\newline)
	    "Call " id ".Add(Env_" id ")" (string #\newline)
	    "Lambda_Counter = Lambda_Counter + 1" (string #\newline)
	     "Set GlobalEnv.item(\"" id "_\" & Lambda_Counter) =  " id (string #\newline))))))

(defun function-comp (id formal-args body)
  (let ((prelude (cat-newlines "instance = args.item(1).getval"
			       "Dim Env as New Scripting.Dictionary"
			       "Dim LR as New Collection"
			       "Set LR = Lookup(instance, Env)"
			       "set Env = LR.item(2)"))
	(formal-args-list (loop for i = 2 then (+ i 1)
			       for arg in formal-args
			       collect (cat "set Env.Item(\"" (string arg) "\")"
					     " = args.item(" (int-string i) ")")))
	(writeback (cat-newlines "Dim LR_writeback as New Collection"
				 "Call LR_writeback.add(LR.item(1))"
				 "Call LR_writeback.add(Env)"
				 "set GlobalEnv.item(instance) = LR_writeback")))
    (multiple-value-bind (ret-val body-text) (progn-comp body '())
      (cat "Public Function " id "(args)" (string #\newline)
	   prelude (string #\newline)
	   (add-newlines formal-args-list) (string #\newline)
	   body-text (string #\newline)
	   writeback (string #\newline)
	   "set current_res = " ret-val (string #\newline)
	   "End Function" (string #\newline)))))

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
     finally (return (values (cat "funcall" (args-to-string ret-vals))
			     (add-newlines bodies)))))

(defun quote-comp (exp env)
  (cond ((atom exp) (quote-atom exp env))
	((listp exp) (quote-list exp env))))
;TODO:  add vectors, maps, etc

(defun quote-atom (exp env)
  (cond ((symbolp exp) (cat "Create(\"" (string exp) "\", \"symb\")"))
	((stringp exp) (cat "Create(\"" exp "\", \"symb\")"))
	((numberp exp) (cat "Create(\"" (format nil "~a" exp) "\", \"symb\")"))))

(defun quote-list (exp env)
  (let ((quoted-list (map 'list (lambda (x) (quote-comp x env)) exp)))
    (cat-newlines "Funcall(Lookup(\"List\", Env), "
		 (args-to-string quoted-list))))
	

;utility functions
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
(defun outer-comp (exp)
  (setf *lambda-bodies* '())
  (multiple-value-bind (ret-val body) (comp exp '())
    (cat-newlines (string #\newline)
	  "public sub main()"
	  "call BuiltinInit()"
	  "dim Env as new Scripting.Dictionary"
	  body
	  ret-val
	  "end sub"
	  (string #\newline)
	  "'lambda functions"
	  (string #\newline)
	  (add-newlines *lambda-bodies*))))

(defun cat-newlines (&rest strs)
  (if (> (length strs) 1)
      (add-newlines strs)
      (car strs)))

(defmacro cat (&rest strs)
  `(concatenate 'string ,@strs))

(defun prefix-to-infix (lst token)
  ;inserts token between each pair of elements in lst, then returns the result as a string
  (concatenate 'string
	       (reduce #'(lambda (x y) (concatenate 'string (string x) " " token " " (string y))) lst)))

(defun add-newlines (lst)
  (if lst
      (prefix-to-infix lst (string #\newline))
      " "))

(defun args-to-string (lst)
  (concatenate 'string "("
	       (reduce #'(lambda (x y) (concatenate 'string (string x) ", " (string y))) lst)
	       ")"))

(defun int-string (int)
  (format nil "~a" int))

(defmacro let-test (bindings &rest body)
  (loop for binding in bindings
     collecting (car binding) into vars
     collecting (cadr binding) into vals
     finally (return (concatenate 'list '(fn) (list vars) body vals))))

(defun inline-fun (name body)
  (cat "Public Function " name "(args as collection)" (string #\newline)
       "'function prelude" (string #\newline)
       "    instance = args.item(1).GetVal" (string #\newline)
       "    Dim LR as New Collection" (string #\newline)
       "    Set LR = GlobalEnv.item(instance)" (string #\newline)
       "    Dim Env as new Scripting.Dictionary" (string #\newline)
       "    Set Env = LR.item(2)" (string #\newline)
       "'actual function body" (string #\newline)
       body (string #\newline)
       "'local environment writeback" (string #\newline)
       "    Dim LR_Writeback as New Collection" (string #\newline)
       "    Call LR_Writeback.Add(LR.item(1))" (string #\newline)
       "    Call LR_Writeback.Add(Env)" (string #\newline)
       "    Set GlobalEnv.item(instance) = LR_Writeback" (string #\newline)
       "'return value" (string #\newline)
       "    set Module1.current_res = res" (string #\newline)
       "End Function" (string #\newline)))

(defun inline-builtin-fun (name symbol body)
  (list (inline-fun name body)
	(fun-env-add name symbol)))

(defun fun-env-add (name symbol)
  (cat "Dim " name "LR as New Collection" (string #\newline)
       name "LR.Add \"" name "\"" (string #\newline)
       name "LR.Add NullEnv" (string #\newline)
       "Set GlobalEnv.item(\"" name "LR\") = " name "LR" (string #\newline)
       "Set GlobalEnv.item(\"" symbol "\") = Create(\"" name "LR\", \"func\")" (string #\newline)))

(defun add-inline-builtins (&rest funcs)
  (let ((all-list (map 'list (lambda (lst) (apply #'inline-builtin-fun lst)) funcs)))
    (let ((body-list (map 'list #'car all-list))
	  (env-list (map 'list #'cadr all-list)))
      (cat-newlines (add-newlines env-list) 
	    (string #\newline)
	    (add-newlines body-list)))))

;tests
(eval-when (:execute)
  (format t "~A" (add-inline-builtins 
		  (list "Cons" "cons"
			(cat-newlines "Dim res as New List"
				      "Set res.Car = args.Item(2)"
				      "Set res.Cdr = args.Item(3)"))
		  (list "Car" "car"
			(cat-newlines "Set res = args.Item(2).Car"))
		  (list "Cdr" "cdr"
			(cat-newlines "Set res = args.Item(2).Cdr"))
		  (list "CoreWhile" "core_while"
					;args(2) is a test/body function, which must return
					;a list with true or false as the first element and
					;the return value of the body function as the second
					;loops until the first value is false
			(cat-newlines "Dim LoopRes as new list"
				      "Dim LoopTest as boolean"
				      "Do"
				      "Set LoopRes = Funcall(args.item(2))"
				      "Set LoopTest = CoreTest(LoopRes.Car)"
				      "Loop While LoopTest = True"
				      "Set res = LoopRes.cdr")))))
			
(eval-when (:execute)
  (setf *macros* '())
  (defmacro-comp '(defmac let (bindings &rest body)
			    (loop for binding in bindings
			       collecting (car binding) into vars
			       collecting (cadr binding) into vals34
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
