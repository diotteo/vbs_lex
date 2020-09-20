from enum import Enum, auto
import re

REAL_RE = re.compile('^(-|\+)?(\d*\.\d+)|(\d+\.\d*)$')
INT_RE = re.compile('^(-|\+)?\d+$')
DATE_RE = re.compile('^#.*#$')

class VarType(Enum):
	EMPTY = auto()
	OBJECT = auto()
	NULL = auto()
	ARRAY = auto()
	INTEGER = auto()
	REAL = auto()
	BOOLEAN = auto()
	DATE = auto()

	SPECIAL_VALUES = ('empty', 'nothing', 'null', 'true', 'false')

	def varTypeFromStr(s):
		up_s = s.upper()
		if up_s == 'EMPTY':
			return VarType.EMPTY
		elif up_s == 'NOTHING':
			return VarType.OBJECT
		elif up_s == 'NULL':
			return VarType.NULL
		elif isBoolean(up_s):
			return VarType.BOOLEAN
		elif isInt(up_s):
			return VarType.INTEGER
		elif isReal(up_s):
			return VarType.REAL
		elif isDate(up_s):
			return VarType.DATE


	def isBoolean(s):
		return up_s in ('TRUE', 'FALSE')

	def isInt(s):
		return INT_RE.match(s) is not None

	def isReal(s):
		return REAL_RE.match(s) is not None

	def isDate(s):
		return DATE_RE.match(s) is not None


KEYWORDS = (
'byRef',
'byVal',
'class',
'dim',
'each',
'end',
'exit',
'for',
'function',
'get',
'let',
'new',
'private',
'property',
'public',
'set',
'sub')

FUNCTIONS = (
('abs', 1, 1), #bool->int, int->int, real->real
('array', 0, None), #whatever->array
('asc', 1, 1),
('atn', 1, 1),
('call', 1, 1),
('cbool', 1, 1),
('cbyte', 1, 1),
('ccur', 1, 1),
('cdate', 1, 1),
('cdbl', 1, 1),
('chr', 1, 1),
('cint', 1, 1),
('clng', 1, 1),
('cos', 1, 1),
('createObject', 1, 2),
('csng', 1, 1),
('cstr', 1, 1),
('date', 0, 0),
('dateAdd', 3, 3),
('dateDiff', 3, 5),
('datePart', 2, 4),
('dateSerial', 3, 3),
('dateValue', 1, 1),
('day', 1, 1),
('eval', 1, 1),
('exp', 1, 1),
('filter', 2, 4),
('fix', 1, 1),
('formatCurrency', 1, 5),
('formatDateTime', 2, 2),
('formatNumber', 1, 5),
('formatPercent', 1, 5),
('getref', 1, 1),
('hex', 1, 1),
('hour', 1, 1),
('inStr', 2, 4),
('inStrRev', 2, 4),
('int', 1, 1),
('isArray', 1, 1),
('isDate', 1, 1),
('isEmpty', 1, 1),
('isNull', 1, 1),
('isNumeric', 1, 1),
('isObject', 1, 1),
('join', 1, 2),
('lbound', 1, 2),
('lcase', 1, 1),
('left', 2, 2),
('len', 1, 1),
('log', 1, 1),
('ltrim', 1, 1),
('mid', 2, 3),
('minute', 1, 1),
('month', 1, 1),
('monthName', 1, 2),
('now', 0, 0),
('oct', 1, 1),
('replace', 3, 6),
('rgb', 3, 3),
('right', 2, 2),
('rnd', 0, 1),
('round', 1, 2),
('rtrim', 1, 1),
('scriptEngine', 0, 0),
('scriptEngineBuildVersion', 0, 0),
('scriptEngineMajorVersion', 0, 0),
('scriptEngineMinorVersion', 0, 0),
('second', 1, 1),
('sgn', 1, 1),
('sin', 1, 1),
('space', 1, 1),
('split', 1, 4),
('sqr', 1, 1),
('strComp', 2, 3),
('string', 2, 2),
('strReverse', 1, 1),
('tan', 1, 1),
('time', 0, 0),
('timeSerial', 3, 3),
('timeValue', 1, 1),
('timer', 0, 0),
('trim', 1, 1),
('typeName', 1, 1),
('ubound', 1, 2),
('ucase', 1, 1),
('varType', 1, 1),
('weekday', 1, 2),
('weekdayName', 1, 3),
('year', 1, 1))

OPERATORS = ('+', '-', '/', '*', 'mod', '&')

class LexemeType(Enum):
	INIT = auto()
	SPACE = auto()
	KEYWORD = auto()
	IDENTIFIER = auto()
	PAREN_BEGIN = auto()
	PAREN_END = auto()
	COMMA = auto()
	LITERAL = auto()
	OPERATOR = auto()
	LINE_CONT = auto()
