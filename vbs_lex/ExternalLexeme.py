from .Lexeme import Lexeme

class ExternalLexeme(Lexeme):
	def __init__(self, name, type_):
		super().__init__(name, type_, None, None, None)
