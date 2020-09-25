class LexemeBase:
	def __init__(self, s, type_, fpath, line, col):
		self.s = s
		self.type = type_
		self.fpath = fpath
		self.line = line
		self.col = col
		self.prev = None
		self.next = None

	def __repr__(self):
		return '{}:{}:{}:{}'.format(self.fpath, self.line, self.col, str(self))

	def __str__(self):
		s = self.s
		if s == '\n':
			s = '\\n'
		return '{} {}'.format(self.type.name, s)

