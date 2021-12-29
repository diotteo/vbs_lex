from .Statement import StatementType

class StatementGroup:
	def __init__(self, type_):
		self.stmts = []
		self.type = type_
		self.parent = None
		self.children = []
		self.ns = None

	def new_child(self, type_):
		child = StatementGroup(type_)
		child.parent = self
		self.children.append(child)
		return child

	def add_stmt(self, stmt):
		self.stmts.append(stmt)


	@staticmethod
	def from_statements(stmts):
		#holds the statements
		stmt_grp = StatementGroup(None)
		top_stmt_grp = stmt_grp

		for stmt in stmts:
			if stmt.type == StatementType.CLASS_BEGIN:
				stmt_grp = stmt_grp.new_child(stmt.type)
				stmt_grp.add_stmt(stmt)

			elif stmt.type == StatementType.SUB_BEGIN:
				stmt_grp = stmt_grp.new_child(stmt.type)
				stmt_grp.add_stmt(stmt)

			elif stmt.type == StatementType.FUNCTION_BEGIN:
				stmt_grp = stmt_grp.new_child(stmt.type)
				stmt_grp.add_stmt(stmt)

			elif stmt.type in (
					StatementType.PROPERTY_GET_BEGIN,
					StatementType.PROPERTY_LET_BEGIN,
					StatementType.PROPERTY_SET_BEGIN,
					):
				stmt_grp = stmt_grp.new_child(stmt.type)
				stmt_grp.add_stmt(stmt)

			elif stmt.type in (
					StatementType.CLASS_END,
					StatementType.SUB_END,
					StatementType.FUNCTION_END,
					StatementType.PROPERTY_END,
					):

				stmt_grp.add_stmt(stmt)
				stmt_grp = stmt_grp.parent

			else:
				stmt_grp.add_stmt(stmt)

		return top_stmt_grp


	@property
	def groups(self):
		cur_lvl = None
		next_lvl = [self]

		while len(next_lvl) > 0:
			cur_lvl = next_lvl
			next_lvl = []
			for grp in cur_lvl:
				yield grp
				next_lvl.extend(grp.children)


	def print_groups(self):
		for grp in self.groups:
			name_s = '?'
			if grp.ns is not None:
				name_s = grp.ns.name
			print('{}:'.format(name_s))
			for stmt in grp.stmts:
				print('  {}'.format(str(stmt)))
			print()
