"""VBScript syntax parsing: turn a series of Lexemes into a Namespace tree

Classes:

    Statement
    Namespace
    Variable
    ExternalVariable
"""

from .namespace import Namespace
from .statement import Statement
from .variable import Variable
from .external_variable import ExternalVariable
