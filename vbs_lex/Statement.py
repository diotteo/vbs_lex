import pdb

from enum import Enum, auto
from .Lexeme import LexemeType
from .LexemeException import LexemeException

class StatementSm(Enum):
	INIT = auto()
	REGULAR_STMT = auto()
	END_STMT = auto()
	EXIT_STMT = auto()
	IDENTIFIER_STMT = auto()
	VISIBILITY_STMT = auto()
	PROPERTY_DECL_STMT = auto()


class StatementType(Enum):
	UNKNOWN = auto()
	BLANK_LINE = auto()
	CLASS_BEGIN = auto()
	CLASS_END = auto()
	FUNCTION_BEGIN = auto()
	FUNCTION_END = auto()
	PROPERTY_GET_BEGIN = auto()
	PROPERTY_LET_BEGIN = auto()
	PROPERTY_SET_BEGIN = auto()
	PROPERTY_END = auto()
	SUB_BEGIN = auto()
	SUB_END = auto()

	CONST_DECLARE = auto()
	DIM = auto()
	REDIM = auto()
	FIELD_DECLARE = auto()

	VAR_ASSIGNMENT = auto()
	OBJECT_ASSIGNMENT = auto()
	PROC_CALL = auto()
	IMPLICIT_PROC_CALL = auto()

	LOOP_EXIT = auto()
	PROC_EXIT = auto()

	DO_LOOP_BEGIN = auto()
	DO_LOOP_END = auto()
	FOR_LOOP_BEGIN = auto()
	FOR_LOOP_END = auto()
	WHILE_LOOP_BEGIN = auto()
	WHILE_LOOP_END = auto()
	IF_BEGIN = auto()
	IF_ELSE = auto()
	IF_ELSE_IF = auto()
	IF_END = auto()
	SELECT_BEGIN = auto()
	SELECT_CASE = auto()
	SELECT_END = auto()
	WITH_BEGIN = auto()
	WITH_END = auto()

	RANDOMIZE = auto()
	EXECUTE = auto()
	EXECUTEGLOBAL = auto()
	OPTION = auto()
	ON_ERROR = auto()


class Statement:
	def __init__(self, type_, lxms):
		self.type = type_
		self.lxms = lxms

	def __str__(self):
		return self.type.name + ': ' + ' '.join((str(x) for x in self.lxms))

	@staticmethod
	def _init_state(stmt_lxms, stmt_type, sm):
		#If statement is straightforward, set the type immediately, set sm to REGULAR_STMT and just accumulate lexemes until the end of the statement
		#For more complicated statement, leave stmt_type UNKNOWN and set SM per what we know so far

		keyword_str2type_d = {
				'CLASS': StatementType.CLASS_BEGIN,
				'SUB': StatementType.SUB_BEGIN,
				'FUNCTION': StatementType.FUNCTION_BEGIN,
				'DIM': StatementType.DIM,
				'REDIM': StatementType.REDIM,
				'CONST': StatementType.CONST_DECLARE,
				'CALL': StatementType.PROC_CALL,
				'SET': StatementType.OBJECT_ASSIGNMENT,
				'SELECT': StatementType.SELECT_BEGIN,
				'CASE': StatementType.SELECT_CASE,
				'IF': StatementType.IF_BEGIN,
				'ELSE': StatementType.IF_ELSE,
				'ELSEIF': StatementType.IF_ELSE_IF,
				'DO': StatementType.DO_LOOP_BEGIN,
				'LOOP': StatementType.DO_LOOP_END,
				'WHILE': StatementType.WHILE_LOOP_BEGIN,
				'WEND': StatementType.WHILE_LOOP_END,
				'FOR': StatementType.FOR_LOOP_BEGIN,
				'NEXT': StatementType.FOR_LOOP_END,
				'EXECUTE': StatementType.EXECUTE,
				'EXECUTEGLOBAL': StatementType.EXECUTEGLOBAL,
				'OPTION': StatementType.OPTION,
				'ON': StatementType.ON_ERROR,
				'RANDOMIZE': StatementType.RANDOMIZE,
				}

		lxm = stmt_lxms[-1]
		cur_s = lxm.s.upper()
		sm = StatementSm.REGULAR_STMT
		if lxm.type == LexemeType.KEYWORD:
			if cur_s == 'END':
				sm = StatementSm.END_STMT
			elif cur_s == 'EXIT':
				sm = StatementSm.EXIT_STMT
			elif cur_s in ('PUBLIC', 'PRIVATE'):
				sm = StatementSm.VISIBILITY_STMT
			elif cur_s == 'PROPERTY':
				sm = StatementType.PROPERTY_DECL_STMT,
			elif cur_s in keyword_str2type_d:
				stmt_type = keyword_str2type_d[cur_s]
			else:
				raise LexemeException(lxm, 'Unhandled statement-start keyword: {}'.format(repr(lxm)))
		elif lxm.type == LexemeType.IDENTIFIER:
			sm = StatementSm.IDENTIFIER_STMT
		elif lxm.type == LexemeType.SPECIAL_OBJECT:
			if cur_s in ('WSCRIPT', 'ERR'):
				stmt_type = StatementType.IMPLICIT_PROC_CALL
			else:
				raise LexemeException(lxm, 'Unhandled statement-start special object: {}'.format(repr(lxm)))
		elif lxm.type == LexemeType.PROCEDURE:
			stmt_type = StatementType.IMPLICIT_PROC_CALL
		else:
			raise LexemeException(lxm, 'Unhandled statement-start lexeme: {}'.format(repr(lxm)))

		return stmt_type, sm


	@staticmethod
	def _visibility_stmt_state(stmt_lxms, stmt_type, sm):
		lxm = stmt_lxms[-1]
		cur_s = lxm.s.upper()

		if sm == StatementSm.VISIBILITY_STMT:
			if lxm.type == LexemeType.IDENTIFIER:
				stmt_type = StatementType.FIELD_DECLARE
				sm = StatementSm.REGULAR_STMT
			elif lxm.type == LexemeType.KEYWORD:
				if cur_s == 'FUNCTION':
					stmt_type = StatementType.FUNCTION_BEGIN
					sm = StatementSm.REGULAR_STMT
				elif cur_s == 'SUB':
					stmt_type = StatementType.SUB_BEGIN
					sm = StatementSm.REGULAR_STMT
				elif cur_s == 'PROPERTY':
					sm = StatementSm.PROPERTY_DECL_STMT
				else:
					raise LexemeException(lxm, 'Unhandled visibility keyword: {}'.format(repr(lxm)))
			else:
				raise LexemeException(lxm, 'Unhandled visibility lexeme type: {}'.format(repr(lxm)))
		else:
			raise LexemeException(lxm, 'Unhandled visibility state: {}'.format(repr(sm)))
		return stmt_type, sm

	@staticmethod
	def _property_decl_state(stmt_lxms, stmt_type, sm):
		lxm = stmt_lxms[-1]
		cur_s = lxm.s.upper()

		if sm == StatementSm.PROPERTY_DECL_STMT:
			if lxm.type == LexemeType.KEYWORD:
				if cur_s in ('GET', 'LET', 'SET'):
					stmt_type = StatementType['PROPERTY_' + cur_s + '_BEGIN']
					sm = StatementSm.REGULAR_STMT
				else:
					raise LexemeException(lxm, 'Unhandled property-decl keyword: {}'.format(repr(lxm)))
			else:
				raise LexemeException(lxm, 'Unhandled property-decl lexeme type: {}'.format(repr(lxm)))
		else:
			raise LexemeException(lxm, 'Unhandled property-decl state: {}'.format(repr(sm)))

		return stmt_type, sm

	@staticmethod
	def _identifier_stmt_state(stmt_lxms, stmt_type, sm):
		lxm = stmt_lxms[-1]
		cur_s = lxm.s.upper()

		if sm == StatementSm.IDENTIFIER_STMT:
			if lxm.type == LexemeType.OPERATOR:
				if cur_s == '=':
					stmt_type = StatementType.VAR_ASSIGNMENT
					sm = StatementSm.REGULAR_STMT
				else:
					raise LexemeException(lxm, 'Unhandled identifier-statement operator: {}'.format(repr(lxm)))

			#foo new Bar, 3, "baz" 'Should be a valid proc call
			elif lxm.type == LexemeType.KEYWORD:
				if cur_s == 'NEW':
					stmt_type = StatementType.IMPLICIT_PROC_CALL
					sm = StatementSm.REGULAR_STMT
				else:
					raise LexemeException(lxm, 'Unhandled identifier-statement keyword: {}'.format(repr(lxm)))

			elif lxm.type in (
					LexemeType.IDENTIFIER,
					LexemeType.OBJECT,
					LexemeType.PAREN_BEGIN,
					LexemeType.STRING,
					LexemeType.DATE,
					LexemeType.INTEGER,
					LexemeType.REAL,
					LexemeType.SPECIAL_VALUE,
					LexemeType.SPECIAL_OBJECT,
					LexemeType.PROCEDURE
					):
				stmt_type = StatementType.IMPLICIT_PROC_CALL
				sm = StatementSm.REGULAR_STMT
			elif lxm.type == LexemeType.DOT:
				#Keep processing, as if next lexeme was the first one
				pass
			else:
				raise LexemeException(lxm, 'Unhandled identifier-statement lexeme type: {}'.format(repr(lxm)))
		else:
			raise LexemeException(lxm, 'Unhandled identifier-statement state: {}'.format(repr(sm)))

		return stmt_type, sm

	@staticmethod
	def _end_stmt_state(stmt_lxms, stmt_type, sm):
		lxm = stmt_lxms[-1]
		cur_s = lxm.s.upper()

		if sm == StatementSm.END_STMT:
			if lxm.type == LexemeType.KEYWORD:
				if cur_s in ('CLASS', 'FUNCTION', 'PROPERTY', 'SUB', 'SELECT', 'FOR', 'IF'):
					stmt_type = StatementType[cur_s + '_END']
					sm = StatementSm.REGULAR_STMT
				else:
					raise LexemeException(lxm, 'Unhandled end-statement keyword: {}'.format(repr(lxm)))
			else:
				raise LexemeException(lxm, 'Unhandled end-statement lexeme type: {}'.format(repr(lxm)))
		else:
			raise LexemeException(lxm, 'Unhandled end-statement state: {}'.format(repr(sm)))

		return stmt_type, sm

	@staticmethod
	def _exit_stmt_state(stmt_lxms, stmt_type, sm):
		lxm = stmt_lxms[-1]
		cur_s = lxm.s.upper()

		if sm == StatementSm.EXIT_STMT:
			if lxm.type == LexemeType.KEYWORD:
				if cur_s in ('FUNCTION', 'SUB', 'PROPERTY'):
					stmt_type = StatementType.PROC_EXIT
					sm = StatementSm.REGULAR_STMT
				elif cur_s in ('DO', 'FOR'):
					stmt_type = StatementType.LOOP_EXIT
					sm = StatementSm.REGULAR_STMT
				else:
					raise LexemeException(lxm, 'Unhandled exit-statement keyword: {}'.format(repr(lxm)))
			else:
				raise LexemeException(lxm, 'Unhandled exit-statement lexeme type: {}'.format(repr(lxm)))
		else:
			raise LexemeException(lxm, 'Unhandled exit-statement state: {}'.format(repr(sm)))

		return stmt_type, sm


	@staticmethod
	def statement_list_from_lexemes(lxms):
		stmt_type = StatementType.UNKNOWN
		sm = StatementSm.INIT
		stmts = []
		cur_stmt_lxms = []
		b_is_line_cont = False
		b_complete_stmt = False
		for lxm in lxms:
			if lxm.type in (LexemeType.SPACE, LexemeType.COMMENT):
				continue
			elif lxm.type == LexemeType.LINE_CONT:
				b_is_line_cont = True
				continue
			elif lxm.type == LexemeType.NEWLINE:
				if b_is_line_cont:
					b_is_line_cont = False
					continue
				else:
					b_complete_stmt = True
			elif lxm.type == LexemeType.STATEMENT_CONCAT:
				b_complete_stmt = True

			if b_complete_stmt:
				b_complete_stmt = False

				if len(cur_stmt_lxms) == 0 and stmt_type == StatementType.UNKNOWN:
					stmt_type = StatementType.BLANK_LINE
				stmts.append(Statement(stmt_type, cur_stmt_lxms))
				cur_stmt_lxms = []
				stmt_type = StatementType.UNKNOWN
				sm = StatementSm.INIT
			else:
				b_is_line_cont = False
				cur_stmt_lxms.append(lxm)

				process_state_func = None
				if sm == StatementSm.INIT:
					process_state_func = Statement._init_state
				elif sm == StatementSm.REGULAR_STMT:
					pass
				elif sm == StatementSm.END_STMT:
					process_state_func = Statement._end_stmt_state
				elif sm == StatementSm.VISIBILITY_STMT:
					process_state_func = Statement._visibility_stmt_state
				elif sm == StatementSm.IDENTIFIER_STMT:
					process_state_func = Statement._identifier_stmt_state
				elif sm == StatementSm.EXIT_STMT:
					process_state_func = Statement._exit_stmt_state
				elif sm == StatementSm.PROPERTY_DECL_STMT:
					process_state_func = Statement._property_decl_state
				else:
					raise LexemeException(lxm, 'Unhandled state: {}'.format(repr(sm)))

				if process_state_func is not None:
					stmt_type, sm = process_state_func(cur_stmt_lxms, stmt_type, sm)

		return stmts
