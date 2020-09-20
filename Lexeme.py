from enum import Enum, auto
from LexemeBase import LexemeBase

class LexemeType(Enum):
	SPACE = auto()
	KEYWORD = auto()
	IDENTIFIER = auto()
	PAREN_BEGIN = auto()
	PAREN_END = auto()
	COMMA = auto()
	STRING = auto()
	DATE = auto()
	INTEGER = auto()
	REAL = auto()
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
	def from_LexemeBase(lxm, lex_type=None):
		if lex_type is None:
			lex_type = lxm.type
		return Lexeme(lxm.s, lex_type, lxm.line, lxm.col)

	@staticmethod
	def from_LexemeBaseList(lxms, lex_type=None):
		lxm = lxms[0]
		s = str_from_lexemes(lxms)
		if lex_type is None:
			lex_type = lxm.type
		return Lexeme(s, lex_type, lxm.line, lxm.col)
