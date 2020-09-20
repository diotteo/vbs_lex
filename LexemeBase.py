class LexemeBase:
	def __init__(self, s, type_, line, col):
		self.s = s
		self.type = type_
		self.line = line
		self.col = col

	def __str__(self):
		s = self.s
		if s == '\n':
			s = '\\n'
		return '{}:{}:{} {}'.format(self.line, self.col, self.type.name, s)

