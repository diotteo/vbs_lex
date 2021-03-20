from .LexemeBase import LexemeBase

class ExternalLexeme(LexemeBase):
	def __init__(self, name, type_):
		super().__init__(name, type_, None, None, None)
