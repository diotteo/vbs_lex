from enum import Enum, auto

class LexemeType(Enum):
	SPACE = auto()
	KEYWORD = auto()
	PROCEDURE = auto()
	IDENTIFIER = auto()
	VARIABLE = auto()
	OBJECT = auto()
	CLASS = auto()
	SUB = auto()
	FUNCTION = auto()
	PROPERTY_GET = auto()
	PROPERTY_LET = auto()
	PROPERTY_SET = auto()
	PAREN_BEGIN = auto()
	PAREN_END = auto()
	COMMA = auto()
	STRING = auto()
	DATE = auto()
	INTEGER = auto()
	REAL = auto()
	SPECIAL_VALUE = auto()
	SPECIAL_OBJECT = auto()
	DOT = auto()
	OPERATOR = auto()
	COMMENT = auto()
	LINE_CONT = auto()
	STATEMENT_CONCAT = auto()
	NEWLINE = auto()

class TokenType(Enum):
	INIT = auto()
	SPACE = auto()
	LEXEME = auto()
	PAREN_BEGIN = auto()
	PAREN_END = auto()
	COMMA = auto()
	LITERAL_STRING = auto()
	LITERAL_INTEGER = auto()
	LITERAL_DATE = auto()
	DOT = auto()
	OPERATOR = auto()
	COMMENT = auto()
	LINE_CONT = auto()
	STATEMENT_CONCAT = auto()
	NEWLINE = auto()

class Lexeme:
	def __init__(self, s, type_, token_type, fpath, line, col):
		self.s = s
		self.type = type_
		self.token_type = token_type
		self.fpath = fpath
		self.line = line
		self.col = col
		self.prev = None
		self.next = None

	def __repr__(self):
		return '{}:{}:{}:{}'.format(self.fpath, self.line, self.col, str(self))

	def __str__(self):
		s = self.s
		if s == '\n':
			s = '\\n'

		type_s = '?' if self.type is None else self.type.name
		token_s = '?' if self.token_type is None else self.token_type.name
		return '{} ({}) {}'.format(type_s, token_s, s)


	@staticmethod
	def str_from_lexemes(lxms):
		return ''.join((l.s for l in lxms))


	@staticmethod
	def from_Lexeme(lxm, lex_type):
		if lex_type is None:
			lex_type = lxm.type
		new_lxm = Lexeme(lxm.s, lex_type, lxm.token_type, lxm.fpath, lxm.line, lxm.col)
		new_lxm.prev = lxm.prev
		new_lxm.next = lxm.next
		return new_lxm


	@staticmethod
	def from_LexemeList(lxms, lex_type):
		lxm = lxms[0]
		last_lxm = lxms[-1]
		s = Lexeme.str_from_lexemes(lxms)
		new_lxm = Lexeme(s, lex_type, None, lxm.fpath, lxm.line, lxm.col)
		new_lxm.prev = lxm.prev
		new_lxm.next = last_lxm.next

		if new_lxm.prev is not None:
			new_lxm.prev.next = new_lxm
		if new_lxm.next is not None:
			new_lxm.next.prev = new_lxm
		return new_lxm
