from enum import Enum, auto
from .LexemeBase import LexemeBase

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

class Token(LexemeBase):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
