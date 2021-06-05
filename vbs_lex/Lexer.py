from enum import Enum, auto

from .CoreDataLists import *
from .Lexeme import *
from .LexemeException import LexemeException


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
	elif c in '<>+-/*=&\\':
		state = TokenType.OPERATOR

	return state


def process_lxm_triplet(prev, cur, next_, lxm):
	if prev is None:
		prev = lxm

		return (None, prev, cur, next_)
	elif cur is None:
		cur = lxm
		prev.next = cur
		cur.prev = prev

		return (prev, prev, cur, next_)
	elif next_ is None:
		next_ = lxm
		cur.next = next_
		next_.prev = cur
	else:
		prev = cur
		cur = next_
		next_ = lxm

		cur.next = next_
		next_.prev = cur

	return (cur, prev, cur, next_)


def tokenize_str(s, fpath=None):
	sm = TokenType.INIT

	prev_lxm = None
	cur_lxm = None
	next_lxm = None
	lxm_str = ''
	lineno = 0
	for line in s.splitlines(keepends=True):
		lineno += 1
		colno = 0
		for c in line:
			colno += 1
			next_state = None

			if sm in (TokenType.INIT, TokenType.SPACE):
				if c != '\n' and c.isspace():
					sm = TokenType.SPACE
					lxm_str += c
					continue
				next_state = get_state_for_start_char(c)
			elif sm == TokenType.LEXEME:
				if c.isalnum() or c == '_':
					lxm_str += c
					continue
				next_state = get_state_for_start_char(c)
			elif sm == TokenType.LITERAL_STRING:
				lxm_str += c
				if c == '"':
					new_lxm = Lexeme(lxm_str, None, sm, fpath, lineno, colno - len(lxm_str))
					triplet = process_lxm_triplet(prev_lxm, cur_lxm, next_lxm, new_lxm)
					yield_lxm, prev_lxm, cur_lxm, next_lxm = triplet
					if yield_lxm is not None:
						yield yield_lxm
					lxm_str = ''
					sm = TokenType.INIT
				continue
			elif sm == TokenType.LITERAL_INTEGER:
				if c.isdigit():
					lxm_str += c
					continue
				next_state = get_state_for_start_char(c)
			elif sm == TokenType.LITERAL_DATE:
				lxm_str += c
				if c != '#':
					continue
				new_lxm = Lexeme(lxm_str, None, sm, fpath, lineno, colno - len(lxm_str))
				triplet = process_lxm_triplet(prev_lxm, cur_lxm, next_lxm, new_lxm)
				yield_lxm, prev_lxm, cur_lxm, next_lxm = triplet
				if yield_lxm is not None:
					yield yield_lxm
				lxm_str = ''
				sm = TokenType.INIT
				continue
			elif sm == TokenType.COMMENT:
				if c != '\n':
					lxm_str += c
					continue
				new_lxm = Lexeme(lxm_str, None, sm, fpath, lineno, colno - len(lxm_str))
				triplet = process_lxm_triplet(prev_lxm, cur_lxm, next_lxm, new_lxm)
				yield_lxm, prev_lxm, cur_lxm, next_lxm = triplet
				if yield_lxm is not None:
					yield yield_lxm
				lxm_str = ''
				sm = TokenType.INIT
				next_state = TokenType.NEWLINE
			elif sm == TokenType.OPERATOR:
				if len(lxm_str) > 0:
					start_c = lxm_str[0]
					if start_c == '<' and c in '>=':
							lxm_str += c
							continue
					elif start_c == '>' and c == '=':
						lxm_str += c
						continue
				next_state = get_state_for_start_char(c)
			else:
				next_state = get_state_for_start_char(c)

			if next_state is None:
				raise Exception('Unhandled {} character at {}:{}:{}: {}'.format(sm, fpath, lineno, colno, c))
			if sm == TokenType.INIT and len(lxm_str) == 0:
				pass
			else:
				new_lxm = Lexeme(lxm_str, None, sm, fpath, lineno, colno - len(lxm_str))
				triplet = process_lxm_triplet(prev_lxm, cur_lxm, next_lxm, new_lxm)
				yield_lxm, prev_lxm, cur_lxm, next_lxm = triplet
				if yield_lxm is not None:
					yield yield_lxm
			sm = next_state
			lxm_str = c
	if sm != TokenType.INIT:
		new_lxm = Lexeme(lxm_str, None, sm, fpath, lineno, colno - len(lxm_str))
		triplet = process_lxm_triplet(prev_lxm, cur_lxm, next_lxm, new_lxm)
		yield_lxm, prev_lxm, cur_lxm, next_lxm = triplet
		if yield_lxm is not None:
			yield yield_lxm
	yield next_lxm


def LexemeType_from_TokenType(lxm):
	lex_type = None

	if lxm.token_type == TokenType.LITERAL_DATE:
		return LexemeType.DATE
	elif lxm.token_type == TokenType.LITERAL_INTEGER:
		return LexemeType.INTEGER
	elif lxm.token_type == TokenType.LITERAL_STRING:
		return LexemeType.STRING
	elif lxm.token_type == TokenType.LEXEME:
		return LexemeType.IDENTIFIER

	return lex_type


def lxms_from_str(s, fpath=None):
	prev_lxm = None
	for lxm in tokenize_str(s, fpath=fpath):
		try:
			lex_type = LexemeType[lxm.token_type.name]
		except KeyError:
			lex_type = LexemeType_from_TokenType(lxm)
		lxm.type = lex_type
		if prev_lxm is not None:
			prev_lxm.next = lxm
			lxm.prev = prev_lxm
		prev_lxm = lxm
		yield lxm


#State machine
class PotLexemeSm(Enum):
	NUMERIC_SIGN = auto()
	NUMERIC_RADIX = auto()
	NUMERIC_DECIMAL_SEP = auto()
	NUMERIC_DECIMAL = auto()
	NUMERIC_EXP_SEP = auto()
	NUMERIC_EXP_SIGN = auto()
	NUMERIC_EXPONENT = auto()


def lex_compress(input_lxms):
	prev_lxm = None
	pot_lexeme_sm = None
	pot_sub_lexemes = []
	for lxm in input_lxms:
		b_reprocess = True
		while b_reprocess:
			b_reprocess = False
			if pot_lexeme_sm is None:
				if lxm.s.upper() == 'MOD':
					lxm.type = LexemeType.OPERATOR
					yield lxm
				elif lxm.type == LexemeType.OPERATOR:
					if lxm.s in ('+', '-') and (prev_lxm is None or prev_lxm.type == LexemeType.OPERATOR):
						pot_sub_lexemes.append(lxm)
						pot_lexeme_sm = PotLexemeSm.NUMERIC_SIGN
					else:
						yield lxm
				elif lxm.type == LexemeType.INTEGER:
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_RADIX
				elif lxm.type == LexemeType.DOT:
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_DECIMAL_SEP
				else:
					yield lxm
			elif pot_lexeme_sm == PotLexemeSm.NUMERIC_SIGN:
				if lxm.type == LexemeType.INTEGER:
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_RADIX
				elif lxm.type == LexemeType.DOT:
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_DECIMAL_SEP
				else:
					for pot_lxm in pot_sub_lexemes:
						yield pot_lxm
					pot_lexeme_sm = None
					pot_sub_lexemes = []
					b_reprocess = True
			elif pot_lexeme_sm == PotLexemeSm.NUMERIC_RADIX:
				if lxm.type == LexemeType.DOT:
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_DECIMAL_SEP
				elif lxm.s.upper() == 'E':
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_EXP_SEP
				else:
					new_lxm = Lexeme.from_LexemeList(pot_sub_lexemes, LexemeType.INTEGER)
					yield new_lxm
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
					for pot_lxm in pot_sub_lexemes:
						yield pot_lxm
					pot_lexeme_sm = None
					pot_sub_lexemes = []
					b_reprocess = True
				else:
					new_lxm = Lexeme.from_LexemeList(pot_sub_lexemes, LexemeType.REAL)
					yield new_lxm
					pot_lexeme_sm = None
					pot_sub_lexemes = []
					b_reprocess = True
			elif pot_lexeme_sm == PotLexemeSm.NUMERIC_DECIMAL:
				if lxm.s.upper() == 'E':
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_EXP_SEP
				else:
					new_lxm = Lexeme.from_LexemeList(pot_sub_lexemes, LexemeType.REAL)
					yield new_lxm
					pot_lexeme_sm = None
					pot_sub_lexemes = []
					b_reprocess = True
			elif pot_lexeme_sm == PotLexemeSm.NUMERIC_EXP_SEP:
				if lxm.s in ('+', '-'):
					pot_sub_lexemes.append(lxm)
					pot_lexeme_sm = PotLexemeSm.NUMERIC_EXP_SIGN
				elif lxm.type == LexemeType.INTEGER:
					pot_sub_lexemes.append(lxm)
					new_lxm = Lexeme.from_LexemeList(pot_sub_lexemes, LexemeType.REAL)
					yield new_lxm
					pot_lexeme_sm = None
					pot_sub_lexemes = []
				else:
					raise LexemeException(lxm, 'Error on lexeme:{}'.format(repr(lxm)))
			elif pot_lexeme_sm == PotLexemeSm.NUMERIC_EXP_SIGN:
				if lxm.type == LexemeType.INTEGER:
					pot_sub_lexemes.append(lxm)
					new_lxm = Lexeme.from_LexemeList(pot_sub_lexemes, LexemeType.REAL)
					yield new_lxm
					pot_lexeme_sm = None
					pot_sub_lexemes = []
				else:
					raise LexemeException(lxm, 'Error on lexeme:{}'.format(repr(lxm)))
		prev_lxm = lxm

	if len(pot_sub_lexemes) > 0:
		lxm = pot_sub_lexemes[0]
		raise LexemeException(lxm, 'pot_sub_lexemes is not empty!:{}'.format(repr(lxm)))


def identifier_to_specific_type(lxms):
	prev_lxm = None
	prev_s = None
	prev_type = None

	for lxm in lxms:
		if lxm.type in (LexemeType.SPACE, LexemeType.NEWLINE, LexemeType.LINE_CONT):
			yield lxm
			continue

		up_s = lxm.s.upper()
		if lxm.type == LexemeType.IDENTIFIER:
			if prev_type != LexemeType.DOT:
				new_type = lxm.type
				if up_s in OPERATORS:
					new_type = LexemeType.OPERATOR
				elif up_s in KEYWORDS:
					new_type = LexemeType.KEYWORD
				elif up_s in PROCEDURES:
					new_type = LexemeType.PROCEDURE
				elif up_s in SPECIAL_VALUES:
					new_type = LexemeType.SPECIAL_VALUE
				elif up_s in SPECIAL_OBJECTS:
					new_type = LexemeType.SPECIAL_OBJECT
				lxm.type = new_type
		yield lxm

		prev_lxm = lxm
		prev_type = prev_lxm.type
		prev_s = prev_lxm.s.upper()


def lex_str(s, fpath=None):
	return identifier_to_specific_type(lex_compress(lxms_from_str(s, fpath=fpath)))
