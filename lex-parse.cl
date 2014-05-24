;a toy lisp lexer/parser

;assumes that there are no errors in the input, and requires all symbols to not start with a 
;number.

(defpackage :kit)
(in-package :kit)

(defstruct lexer
  start
  pos
  code
  tokens
  len)

(defun tokenize (code)
  ;returns a list of tokens, 
  ;tokens have the form (identifer string)
  (let ((l (make-lexer :start 0 :pos 0 :code code :tokens '() :len (length code))))
    (loop for state = #'lex-open-paren then (funcall state l)
	 while state)
    (reverse (lexer-tokens l))))

(defun next (l)
  ;returns the next input character
  (incf (lexer-pos l))
  (if (>= (lexer-pos l) (lexer-len l))
      'eof
      (let ((res (elt (lexer-code l) (lexer-pos l))))
	res)))

(defun emit (l token)
  (if (car token)
      (push token (lexer-tokens l))))

(defun is-number (char)
  ;returns t if char is the start of a number, nil otherwise
  (if (find char "1234567890.")
      t
      nil))

(defun next-state (char)
  ;generic next state function, because the transitions after numbers, symbols, strings
  ;and close parens are all the same
  (cond ((equal char #\() #'lex-open-paren)
	((equal char #\)) #'lex-close-paren)
	((is-number char) #'lex-number)
	((equal char #\") #'lex-string)
	((or (equal char #\space) (equal char #\newline)) #'lex-whitespace)
	(t #'lex-symbol)))

(defun lex-open-paren (l)
  ;lexer state for an open paren
  ;emit an open paren token, transition to the next state
  (emit l (list 'l-paren (elt (lexer-code l) (lexer-start l))))
  (incf (lexer-start l))
  (next-state (next l)))

(defun lex-close-paren (l)
  ;lexer state for a close paren
  ;emit a close paren token, transition to the next state
  (emit l (list 'r-paren (elt (lexer-code l) (lexer-start l))))
  (incf (lexer-start l))
  (next-state (next l)))

(defmacro lex-state (name until-cond emit-symb)
  `(defun ,name (l)
     (loop for char = (next l)
	until ,until-cond
	until (equal char 'eof)
	finally (if (equal char 'eof)
		    (return nil)
		    (progn
		      (emit l (list ,emit-symb (subseq (lexer-code l) (lexer-start l) (lexer-pos l))))
		      (setf (lexer-start l) (lexer-pos l))
		      (return (next-state char)))))))

(lex-state lex-whitespace (not (or (equal char #\space) (equal char #\newline))) nil)
(lex-state lex-string (or (equal char #\() (equal char #\)) (equal char #\space)) 'string)
(lex-state lex-number (not (is-number char)) 'number)
(lex-state lex-symbol (or (equal char #\space) (equal char #\newline)) 'symbol)

(eval-when (:execute)
  (print (run-lexer "(+ 1 23.4 \"abcd\" (* 3 4))")))

(defun to-tree (tokens)
  (loop ;with pos = 0
     for token = (pop tokens)
     with buffer-stack = '()
     with cur-buffer = '()
     ;with temp-buffer = '()
     while tokens
     do (cond 
	  ((equal (car token) 'l-paren) (progn (push cur-buffer buffer-stack)
					       (setf cur-buffer '())))
	  ((equal (car token) 'r-paren) (setf cur-buffer (append (pop buffer-stack) (list cur-buffer))))
	  (t (setf cur-buffer (append cur-buffer (list token)))))
       finally (return cur-buffer)))

(eval-when (:execute)
  (print (tokenize "(+ 1 23.4 \"abcd\" (* 3 4) (* 1 2))"))
  (print (to-tree (tokenize "(+ 1 23.4 \"abcd\" (* 3 4) (+ 1 2))"))))
