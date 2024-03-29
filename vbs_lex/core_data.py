"""Reserved or otherwise special identifiers"""

SPECIAL_VALUES = (
    'EMPTY',
    'NULL',
    'TRUE',
    'FALSE',

    'VBCR',
    'VBCRLF',
    'VBLF',
    'VBFORMFEED',
    'VBNULLCHAR',
    'VBNULLSTRING',
    'VBTAB',
    'VBVERTICALTAB',

    'VBBINARYCOMPARE',
    'VBTEXTCOMPARE',
    'VBDATABASECOMPARE',

    'VBBLACK',
    'VBRED',
    'VBGREEN',
    'VBYELLOW',
    'VBBLUE',
    'VBMAGENTA',
    'VBCYAN',
    'VBWHITE',

    'VBSUNDAY',
    'VBMONDAY',
    'VBTUESDAY',
    'VBWEDNESDAY',
    'VBTHURSDAY',
    'VBFRIDAY',
    'VBSATURDAY',
    'VBFIRSTJAN1',
    'VBFIRSTFOURDAYS',
    'VBFIRSTFULLWEEK',
    'VBUSESYSTEM',
    'VBUSESYSTEMDAYOFWEEK',
    'VBGENERALDATE',
    'VBLONGDATE',
    'VBSHORTDATE',
    'VBLONGTIME',
    'VBSHORTTIME',

    'VBOKONLY',
    'VBOKCANCEL',
    'VBABORTRETRYIGNORE',
    'VBYESNOCANCEL',
    'VBYESNO',
    'VBRETRYCANCEL',
    'VBCRITICAL',
    'VBQUESTION',
    'VBEXCLAMATION',
    'VBINFORMATION',
    'VBDEFAULTBUTTON1',
    'VBDEFAULTBUTTON2',
    'VBDEFAULTBUTTON3',
    'VBDEFAULTBUTTON4',
    'VBAPPLICATIONMODAL',
    'VBSYSTEMMODAL',
    'VBOK',
    'VBCANCEL',
    'VBABORT',
    'VBRETRY',
    'VBIGNORE',
    'VBYES',
    'VBNO',

    'VBEMPTY',
    'VBNULL',
    'VBINTEGER',
    'VBLONG',
    'VBSINGLE',
    'VBDOUBLE',
    'VBCURRENCY',
    'VBDATE',
    'VBSTRING',
    'VBOBJECT',
    'VBERROR',
    'VBBOOLEAN',
    'VBVARIANT',
    'VBDATAOBJECT',
    'VBDECIMAL',
    'VBBYTE',
    'VBARRAY',

    'VBOBJECTERROR',
    )

SPECIAL_OBJECTS = (
    'ERR',
    'NOTHING',
    'WSCRIPT',
    'ME',
    'DEBUG',
    )

# &H12AB = 0x12AB
OPERATORS = (
    '&',
    '*',
    '+',
    '-',
    '/',
    '\\',
    '<=',
    '<>',
    '=',
    '>=',
    '^',
    'AND',
    'EQV',
    'IMP',
    'IS',
    'MOD',
    'NOT',
    'OR',
    'XOR',
    )

KEYWORDS = (
    'BYREF',
    'BYVAL',
    'CALL',
    'CASE',
    'CLASS',
    'CONST',
    'DIM',
    'DO',
    'EACH',
    'ELSE',
    'ELSEIF',
    'END',
    'ERASE',
    'ERROR',
    'EXECUTE',
    'EXECUTEGLOBAL',
    'EXIT',
    'EXPLICIT',
    'IF',
    'IN',
    'FOR',
    'FUNCTION',
    'GET',
    'GOTO',
    'LET',
    'LOOP',
    'NEW',
    'NEXT',
    'ON',
    'OPTION',
    'PRESERVE',
    'PRIVATE',
    'PROPERTY',
    'PUBLIC',
    'RANDOMIZE',
    'REDIM',
    'REM',
    'RESUME',
    'SELECT',
    'SET',
    'STEP',
    'SUB',
    'THEN',
    'TO',
    'UNTIL',
    'WEND',
    'WHILE',
    'WITH',
    )

#Partial list: http://www.csidata.com/custserv/onlinehelp/VBSdocs/vbsfun4.htm
#Partial list: https://ss64.com/vb/
_PROCEDURES_raw = (
    ('ABS', 1, 1), #bool->int, int->int, real->real
    ('ARRAY', 0, None), #whatever->array
    ('ASC', 1, 1),
    ('ASCB', 1, 1),
    ('ASCW', 1, 1),
    ('ATN', 1, 1),
    ('CBOOL', 1, 1),
    ('CBYTE', 1, 1),
    ('CCUR', 1, 1),
    ('CDATE', 1, 1),
    ('CDBL', 1, 1),
    ('CHR', 1, 1),
    ('CHRB', 1, 1),
    ('CHRW', 1, 1),
    ('CINT', 1, 1),
    ('CLNG', 1, 1),
    ('COS', 1, 1),
    ('CREATEOBJECT', 1, 2),
    ('CSNG', 1, 1),
    ('CSTR', 1, 1),
    ('DATE', 0, 0),
    ('DATEADD', 3, 3),
    ('DATEDIFF', 3, 5),
    ('DATEPART', 2, 4),
    ('DATESERIAL', 3, 3),
    ('DATEVALUE', 1, 1),
    ('DAY', 1, 1),
    ('ESCAPE', 1, 1),
    ('EVAL', 1, 1),
    ('EXP', 1, 1),
    ('FILTER', 2, 4),
    ('FIX', 1, 1),
    ('FORMATCURRENCY', 1, 5),
    ('FORMATDATETIME', 2, 2),
    ('FORMATNUMBER', 1, 5),
    ('FORMATPERCENT', 1, 5),
    ('GETLOCALE', 0, 0),
    ('GETOBJECT', 1, 2),
    ('GETREF', 1, 1),
    ('HEX', 1, 1),
    ('HOUR', 1, 1),
    ('INPUTBOX', 1, 7),
    ('INSTR', 2, 4),
    ('INSTRREV', 2, 4),
    ('INT', 1, 1),
    ('ISARRAY', 1, 1),
    ('ISDATE', 1, 1),
    ('ISEMPTY', 1, 1),
    ('ISNULL', 1, 1),
    ('ISNUMERIC', 1, 1),
    ('ISOBJECT', 1, 1),
    ('JOIN', 1, 2),
    ('LBOUND', 1, 2),
    ('LCASE', 1, 1),
    ('LEFT', 2, 2),
    ('LEFTB', 2, 2),
    ('LEN', 1, 1),
    ('LENB', 1, 1),
    ('LOADPICTURE', 1, 1),
    ('LOG', 1, 1),
    ('LTRIM', 1, 1),
    ('MID', 2, 3),
    ('MIDB', 2, 3),
    ('MINUTE', 1, 1),
    ('MONTH', 1, 1),
    ('MONTHNAME', 1, 2),
    ('MSGBOX', 1, 5), # MsgBox(prompt[, buttons][, title][, helpfile, context])
    ('NOW', 0, 0),
    ('OCT', 1, 1),
    ('REPLACE', 3, 6),
    ('RGB', 3, 3),
    ('RIGHT', 2, 2),
    ('RIGHTB', 2, 2),
    ('RND', 0, 1),
    ('ROUND', 1, 2),
    ('RTRIM', 1, 1),
    ('SETLOCALE', 1, 1),
    ('SCRIPTENGINE', 0, 0),
    ('SCRIPTENGINEBUILDVERSION', 0, 0),
    ('SCRIPTENGINEMAJORVERSION', 0, 0),
    ('SCRIPTENGINEMINORVERSION', 0, 0),
    ('SECOND', 1, 1),
    ('SGN', 1, 1),
    ('SIN', 1, 1),
    ('SPACE', 1, 1),
    ('SPLIT', 1, 4),
    ('SQR', 1, 1),
    ('STRCOMP', 2, 3),
    ('STRING', 2, 2),
    ('STRREVERSE', 1, 1),
    ('TAN', 1, 1),
    ('TIME', 0, 0),
    ('TIMESERIAL', 3, 3),
    ('TIMEVALUE', 1, 1),
    ('TIMER', 0, 0),
    ('TRIM', 1, 1),
    ('TYPENAME', 1, 1),
    ('UBOUND', 1, 2),
    ('UCASE', 1, 1),
    ('UNESCAPE', 1, 1),
    ('VARTYPE', 1, 1),
    ('WEEKDAY', 1, 2),
    ('WEEKDAYNAME', 1, 3),
    ('YEAR', 1, 1),
    )

PROCEDURES = {x[0]: x for x in _PROCEDURES_raw}
