from enum import Enum, auto
import pdb

SPECIAL_VALUES = (
	'EMPTY',
	'NOTHING',
	'NULL',
	'TRUE',
	'FALSE')

SPECIAL_OBJECTS = (
	'WSCRIPT',
)

KEYWORDS = (
	'BYREF',
	'BYVAL',
	'CLASS',
	'DIM',
	'EACH',
	'END',
	'EXIT',
	'FOR',
	'FUNCTION',
	'GET',
	'LET',
	'NEW',
	'PRIVATE',
	'PROPERTY',
	'PUBLIC',
	'REDIM',
	'SET',
	'SUB')

class LexemeBase:
	def __init__(self, s, type_, line, col):
		self.s = s
		if type_ is None:
			pdb.set_trace()
		self.type = type_
		self.line = line
		self.col = col

	def __str__(self):
		s = self.s
		if s == '\n':
			s = '\\n'
		return '{}:{}:{} {}'.format(self.line, self.col, self.type.name, s)

class Token(LexemeBase):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)

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


def get_state_for_start_char(c):
	state = None
	if c == '(':
		state = TokenType.PAREN_BEGIN
	elif c == ')':
		state = TokenType.PAREN_END
	elif c == ',':
		state = TokenType.COMMA
	elif c == '.':
		state = TokenType.DOT
	elif c == '"':
		state = TokenType.LITERAL_STRING
	elif c == '#':
		state = TokenType.LITERAL_DATE
	elif c == '\'':
		state = TokenType.COMMENT
	elif c == '_':
		state = TokenType.LINE_CONT
	elif c == '\n':
		state = TokenType.NEWLINE
	elif c == ':':
		state = TokenType.STATEMENT_CONCAT
	elif c.isspace():
		state = TokenType.SPACE
	elif c.isalpha():
		state = TokenType.LEXEME
	elif c.isdigit():
		state = TokenType.LITERAL_INTEGER
	elif c in '<>+-/*=&':
		state = TokenType.OPERATOR

	return state


def tokenize_str(s):
	sm = TokenType.INIT
	
	tokens = []
	token_str = ''
	lineno = 0
	for line in s.splitlines(keepends=True):
		lineno += 1
		colno = 0
		for c in line:
			colno += 1
			next_state = None

			if sm in (TokenType.INIT, TokenType.SPACE):
				if c.isspace():
					sm = TokenType.SPACE
					token_str += c
					continue
				next_state = get_state_for_start_char(c)
			elif sm == TokenType.LEXEME:
				if c.isalnum() or c == '_':
					token_str += c
					continue
				next_state = get_state_for_start_char(c)
			elif sm == TokenType.LITERAL_STRING:
				#pdb.set_trace()
				token_str += c
				if c == '"':
					yield Token(token_str, sm, lineno, colno)
					token_str = ''
					sm = TokenType.INIT
				continue
			elif sm == TokenType.LITERAL_INTEGER:
				if c.isdigit():
					token_str += c
					continue
				next_state = get_state_for_start_char(c)
			elif sm == TokenType.LITERAL_DATE:
				token_str += c
				if c != '#':
					continue
				yield Token(token_str, sm, lineno, colno)
				token_str = ''
				sm = TokenType.INIT
				continue
			elif sm == TokenType.COMMENT:
				if c != '\n':
					token_str += c
					continue
				yield Token(token_str, sm, lineno, colno)
				token_str = ''
				sm = TokenType.INIT
				next_state = TokenType.NEWLINE
			elif sm == TokenType.OPERATOR:
				if len(token_str) > 0:
					start_c = token_str[0]
					if start_c == '<' and c in '>=':
							token_str += c
							continue
					elif start_c == '>' and c == '=':
						token_str += c
						continue
				next_state = get_state_for_start_char(c)
			else:
				next_state = get_state_for_start_char(c)

			if next_state is None:
				raise Exception('Unhandled {} character at {}:{}: {}'.format(sm, lineno, colno, c))
			if sm == TokenType.INIT and len(token_str) == 0:
				pass
			else:
				yield Token(token_str, sm, lineno, colno)
			sm = next_state
			token_str = c
	if sm != TokenType.INIT:
		yield Token(token_str, sm, lineno, colno)


def LexemeType_from_Token(token):
	lex_type = None

	if token.type == TokenType.LITERAL_DATE:
		return LexemeType.DATE
	elif token.type == TokenType.LITERAL_INTEGER:
		return LexemeType.INTEGER
	elif token.type == TokenType.LITERAL_STRING:
		return LexemeType.STRING
	elif token.type == TokenType.LEXEME:
		if token.s.upper() in KEYWORDS:
			return LexemeType.KEYWORD
		return LexemeType.IDENTIFIER

	return lex_type


def token_to_lex_str(s):
	tokens = tokenize_str(s)

	lxms = []
	for token in tokens:
		try:
			lex_type = LexemeType[token.type.name]
		except KeyError:
			lex_type = None

		if lex_type is None:
			lex_type = LexemeType_from_Token(token)
		lxms.append(Lexeme.from_LexemeBase(token, lex_type=lex_type))

	return lxms


#State machine
class PotLexemeSm(Enum):
	NUMERIC_SIGN = auto()
	NUMERIC_RADIX = auto()
	NUMERIC_DECIMAL_SEP = auto()
	NUMERIC_DECIMAL = auto()
	NUMERIC_EXP_SEP = auto()
	NUMERIC_EXPONENT = auto()


def str_from_lexemes(lxms):
	return ''.join((l.s for l in lxms))


def lex_compress(input_lxms):
	lxms = []
	prev_lxm = None
	pot_lexeme_sm = None
	pot_sub_lexemes = []
	for lxm in input_lxms:
		b_reprocess = True
		while b_reprocess:
			b_reprocess = False
			if pot_lexeme_sm is None:
				if lxm.s.upper() == 'MOD':
					new_lxm = Lexeme.from_LexemeBase(lxm, lex_type=LexemeType.OPERATOR)
					lxms.append(new_lxm)
				elif lxm.type == LexemeType.OPERATOR:
					if lxm.s in ('+', '-'):
						pot_sub_lexemes.append(lxm)
						pot_lexeme_sm = PotLexemeSm.NUMERIC_SIGN
					else:
						lxms.append(lxm)
				elif lxm.type == LexemeType.INTEGER:
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_RADIX
				elif lxm.type == LexemeType.DOT:
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_DECIMAL_SEP
				else:
					lxms.append(lxm)
			elif pot_lexeme_sm == PotLexemeSm.NUMERIC_SIGN:
				if lxm.type == LexemeType.INTEGER:
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_RADIX
				elif lxm.type == LexemeType.DOT:
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_DECIMAL_SEP
				else:
					lxms.extend(pot_sub_lexemes)
					pot_lexeme_sm = None
					pot_sub_lexemes = []
			elif pot_lexeme_sm == PotLexemeSm.NUMERIC_RADIX:
				if lxm.type == LexemeType.DOT:
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_DECIMAL_SEP
				elif lxm.s.upper() == 'E':
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_EXP_SEP
				else:
					new_lxm = Lexeme.from_LexemeBaseList(pot_sub_lexemes, lex_type=LexemeType.INTEGER)
					lxms.append(new_lxm)
					pot_lexeme_sm = None
					pot_sub_lexemes = []
					b_reprocess = True
			elif pot_lexeme_sm == PotLexemeSm.NUMERIC_DECIMAL_SEP:
				if lxm.type == LexemeType.INTEGER:
					pot_lexeme_sm = PotLexemeSm.NUMERIC_DECIMAL
					pot_sub_lexemes.append(lxm)
				elif lxm.s.upper() == 'E':
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_EXP_SEP
				elif len(pot_sub_lexemes) == 1:
					lxms.extend(pot_sub_lexemes)
					pot_lexeme_sm = None
					pot_sub_lexemes = []
					b_reprocess = True
				else:
					new_lxm = Lexeme.from_LexemeBaseList(pot_sub_lexemes, lex_type=LexemeType.REAL)
					lxms.append(new_lxm)
					pot_lexeme_sm = None
					pot_sub_lexemes = []
					b_reprocess = True
			elif pot_lexeme_sm == PotLexemeSm.NUMERIC_DECIMAL:
				if lxm.s.upper() == 'E':
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_EXP_SEP
				else:
					new_lxm = Lexeme.from_LexemeBaseList(pot_sub_lexemes, lex_type=LexemeType.REAL)
					lxms.append(new_lxm)
					pot_lexeme_sm = None
					pot_sub_lexemes = []
					b_reprocess = True
			elif pot_lexeme_sm == PotLexemeSm.NUMERIC_EXP_SEP:
				if lxm.type == LexemeType.INTEGER:
					pot_sub_lexemes.append(lxm)
					new_lxm = Lexeme.from_LexemeBaseList(pot_sub_lexemes, lex_type=LexemeType.REAL)
					lxms.append(new_lxm)
					pot_lexeme_sm = None
					pot_sub_lexemes = []

	return lxms


def lex_str(s):
	return lex_compress(token_to_lex_str(s))
