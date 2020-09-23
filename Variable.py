class Variable:
	def __init__(self, def_ns, def_lxm, b_is_def=True):
		ref = VariableReference(self, def_ns, def_lxm)
		if b_is_def:
			self.definition = ref
		else:
			self.definition = None
		self.name = def_lxm.s
		self.m_refs = [ref]

	def add_ref(self, ref_ns, ref_lxm):
		self.m_refs.append(VariableReference(self, ref_ns, ref_lxm))

	@staticmethod
	def new_implicit_def(self, ns, lxm):
		return Variable(ns, lxm, b_is_def=False)

class VariableReference:
	def __init__(self, variable, ns, lxm):
		self.variable = variable
		self.ns = ns
		self.lxm = lxm
