from .Variable import Variable, VariableReference
from .Lexeme import LexemeType
from .ExternalLexeme import ExternalLexeme

class ExternalVariable(Variable):
	def __init__(self, name, ns=None):
		def_lxm = ExternalLexeme(name, LexemeType.VARIABLE)
		super().__init__(ns, def_lxm, b_is_def=True)
