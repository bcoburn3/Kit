;Compiler for Kit, a toy lisp-like language.  Will eventually compile to VBA.  

;Key insight for the compiler:  a compiler is like an interpreter, but instead of actually
;doing the computation it simply outputs code which does instead


#|(define (evaluate e env)
  (if (atom? e) 
      (cond ((symbol? e) (lookup e env))
            ((or (number? e)(string? e)(char? e)(boolean? e)(vector? e))
             e )
            (else (wrong "Cannot evaluate" e)) )
      (case (car e)
        ((quote)  (cadr e))
        ((if)     (if (not (eq? (evaluate (cadr e) env) the-false-value))
                      (evaluate (caddr e) env)
                      (evaluate (cadddr e) env) ))
        ((begin)  (eprogn (cdr e) env))
        ((set!)   (update! (cadr e) env (evaluate (caddr e) env)))
        ((lambda) (make-function (cadr e) (cddr e) env))
        (else     (invoke (evaluate (car e) env)
                          (evlis (cdr e) env) )) ) ) )|#

(defvar *infix-operators* (list '+ '* '- '/ '=))
(setf *infix-operators* (list '+ '* '- '/ '=))

(defun comp (exp)
  (if (atom exp)
      (format nil "~a" exp)
      (cond ; (car exp)
	((equal (car exp) 'quote) (cadr exp))
	((equal (car exp) 'if) (concatenate 'string "if " (comp (cadr exp)) " then:" (string #\newline)
			   "    " (comp (caddr exp)) (string #\newline)
			   "else:" (string #\newline)
			   "    " (comp (cadddr exp)) (string #\newline)
			   "end if"))
	((equal (car exp) 'progn) (progn-comp exp)) ;probably just insert a newline between every expression
	((equal (car exp) 'setf) (concatenate 'string "call MySet(" (comp (cadr exp)) ", " (comp (caddr exp)) ")"))
	((equal (car exp) 'defun) (defun-comp exp))
	(t (function-comp exp)))))

(defun prefix-to-infix (lst token)
  ;inserts token between each pair of elements in lst, then returns the result as a string
  (concatenate 'string "("
	       (reduce #'(lambda (x y) (concatenate 'string (string x) " " token " " (string y))) lst)
	       ")"))

(defun args-to-string (lst)
  (concatenate 'string "("
	       (reduce #'(lambda (x y) (concatenate 'string (string x) ", " (string y))) lst)
	       ")"))

(defun function-comp (exp)
  (let ((func (car exp))
	(args (mapcar #'comp (cdr exp))))
    (if (find func *infix-operators*)
	(prefix-to-infix args (format nil "~a" func))
	(concatenate 'string (symbol-name func) (args-to-string args)))))

(defun progn-comp (exp)
  (let ((sub-exps (mapcar #'comp (cdr exp))))
    (reduce #'(lambda (x y) (concatenate 'string x (string #\newline) y)) sub-exps)))

(defun defun-comp (exp)
  (let ((name (comp (cadr exp)))
	 (args (comp (caddr exp)))
	 (body (mapcar #'comp (cdddr exp))))
    (concatenate 'string
		 "public sub " name (args-to-string args) (string #\newline)
		 (progn-comp body) (string #\newline)
		 "end sub")))

(print *infix-operators*)	

(eval-when (:execute)
  (print (comp '(* 1 2 3 4)))
  (print (comp '(+ 1 2 (* 3 4))))
  (print (comp '(plus 1 2 3 4)))
  (print (comp '(if (= a b) (+ 1 2) (* 3 4))))
  (print (comp '(progn (+ 1 2 3) (times 3 4))))
  (print (comp '(setf a 3)))
  (print (comp '(defun plus (x y)
		 (+ x y)))))
