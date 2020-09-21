from enum import Enum, auto

class NamespaceSm(Enum):
	INIT = auto()
	CLASS_BEGIN = auto()
	CLASS_END = auto()
	FUNCTION_BEGIN = auto()
	FUNCTION_END = auto()
	SUB_BEGIN = auto()
	SUB_END = auto()
	PROPERTY_BEGIN = auto()
	PROPERTY_GET = auto()
	PROPERTY_LET = auto()
	PROPERTY_SET = auto()
	PROPERTY_END = auto()


class Namespace:
	global_ = None

	def __init__(self, parent, lxm):
		self.parent = parent
		self.lxm = lxm
		self.m_classes = {}
		self.m_objects = {}
		self.m_vars = {}
		self.m_functions = {}
		self.m_subs = {}
		self.m_properties = {}

	@property
	def name(self):
		return self.lxm.s

	@staticmethod
	def get_global_ns():
		if Namespace.global_ is None:
			Namespace.global_ = Namespace(None, None)
		return Namespace.global_

	def add_class(self, lxm):
		self.m_classes[lxm.s] = lxm
		return Namespace(self, lxm)
	def add_object(self, lxm):
		self.m_objects[lxm.s] = lxm
		return Namespace(self, lxm)
	def add_var(self, lxm):
		self.m_vars[lxm.s] = lxm
		return Namespace(self, lxm)
	def add_function(self, lxm):
		self.m_functions[lxm.s] = lxm
		return Namespace(self, lxm)
	def add_sub(self, lxm):
		self.m_subs[lxm.s] = lxm
		return Namespace(self, lxm)
	def add_property(self, lxm):
		self.m_properties[lxm.s] = lxm
		return Namespace(self, lxm)
