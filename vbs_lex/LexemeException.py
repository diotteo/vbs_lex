class LexemeException(Exception):
	def __init__(self, lxm, message):
		self.message = '{}: {}'.format(lxm, message)
