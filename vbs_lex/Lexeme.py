from enum import Enum, auto
from .LexemeBase import LexemeBase

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
	PROPERTY = auto()
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

class Lexeme(LexemeBase):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)

	@staticmethod
	def str_from_lexemes(lxms):
		return ''.join((l.s for l in lxms))


	@staticmethod
	def from_LexemeBase(lxm, lex_type=None):
		if lex_type is None:
			lex_type = lxm.type
		new_lxm = Lexeme(lxm.s, lex_type, lxm.fpath, lxm.line, lxm.col)
		new_lxm.prev = lxm.prev
		new_lxm.next = lxm.next
		return new_lxm


	@staticmethod
	def from_LexemeBaseList(lxms, lex_type=None):
		lxm = lxms[0]
		s = Lexeme.str_from_lexemes(lxms)
		if lex_type is None:
			lex_type = lxm.type
		new_lxm = Lexeme(s, lex_type, lxm.fpath, lxm.line, lxm.col)
		new_lxm.prev = lxm.prev
		new_lxm.next = lxm.next
		return new_lxm
