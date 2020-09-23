from enum import Enum, auto
from Variable import *
from Lexeme import LexemeType

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
		self.m_vars = {}
		self.m_classes = {}
		self.m_functions = {}
		self.m_subs = {}
		self.m_properties = {}
		self.ns_list = []

	@property
	def name(self):
		if self.lxm is None:
			return '[global]'
		return self.lxm.s

	@property
	def full_name(self):
		s = ''
		cur = self
		while cur.parent is not None:
			if len(s) == 0:
				s = cur.name
			else:
				s = cur.name + '.' + s
			cur = cur.parent
		return s

	def __str__(self):
		if self.lxm is None:
			return self.name
		type_ = self.lxm.type.name
		return '{} {}'.format(type_, self.full_name)

	@property
	def vars(self):
		return self.m_vars

	@staticmethod
	def get_global_ns():
		if Namespace.global_ is None:
			Namespace.global_ = Namespace.new_top_ns()
		return Namespace.global_

	@staticmethod
	def new_top_ns():
		return Namespace(None, None)

	@staticmethod
	def reset_global_ns():
		Namespace.global_ = Namespace.new_top_ns()
		return Namespace.global_

	def add_var(self, lxm):
		var = Variable(self, lxm)
		self.m_vars[lxm.s.upper()] = var
		return var

	def add_implicit_var(self, lxm):
		var = Variable.new_implicit_def(self, self, lxm)
		self.m_vars[lxm.s.upper()] = var
		return var

	def get_var(self, s):
		up_s = s.upper()
		if up_s in self.m_vars:
			return self.m_vars[up_s]
		elif self.parent is None:
			return None
		else:
			return self.parent.get_var(up_s)

	def add_class(self, lxm):
		self.m_classes[lxm.s] = lxm
		sub_ns = Namespace(self, lxm)
		self.ns_list.append(sub_ns)
		return sub_ns
	def add_function(self, lxm):
		self.m_functions[lxm.s] = lxm
		sub_ns = Namespace(self, lxm)
		self.ns_list.append(sub_ns)
		return sub_ns
	def add_sub(self, lxm):
		self.m_subs[lxm.s] = lxm
		sub_ns = Namespace(self, lxm)
		self.ns_list.append(sub_ns)
		return sub_ns
	def add_property(self, lxm):
		self.m_properties[lxm.s] = lxm
		sub_ns = Namespace(self, lxm)
		self.ns_list.append(sub_ns)
		return sub_ns


	@staticmethod
	def process_lexemes(lxms, top_ns=None):
		if top_ns is None:
			top_ns = Namespace.new_top_ns()
		ns = top_ns
		prev_lxm = None
		prev_s = None
		sm = NamespaceSm.INIT
		for lxm in lxms:
			if lxm.type in (LexemeType.SPACE, LexemeType.NEWLINE, LexemeType.LINE_CONT, LexemeType.STATEMENT_CONCAT):
				pass
			else:
				cur_s = lxm.s.upper()

				if sm == NamespaceSm.INIT:
					if prev_s in ('DIM', 'CONST') and prev_lxm.type == LexemeType.KEYWORD:
						if lxm.type == LexemeType.IDENTIFIER:
							ns.add_var(lxm)
						else:
							raise Exception('Unexpected followup to DIM keyword: {}'.format(repr(lxm)))
					elif prev_s == 'END' and prev_lxm.type == LexemeType.KEYWORD:
						b_ns_end = True
						if lxm.type != LexemeType.KEYWORD:
							raise Exception('Unexpected followup to END keyword: {}'.format(repr(lxm)))
						elif cur_s == 'CLASS':
							sm = NamespaceSm.CLASS_END
						elif cur_s == 'SUB':
							sm = NamespaceSm.SUB_END
						elif cur_s == 'FUNCTION':
							sm = NamespaceSm.FUNCTION_END
						elif cur_s == 'PROPERTY':
							sm = NamespaceSm.PROPERTY_END
						else:
							#end if, end for, etc
							b_ns_end = False

						if b_ns_end:
							ns = ns.parent
							sm = NamespaceSm.INIT
					elif (prev_s == 'EXIT' and prev_lxm.type == LexemeType.KEYWORD) or prev_s == '.':
						pass
					elif lxm.type == LexemeType.KEYWORD:
						if cur_s == 'CLASS':
							sm = NamespaceSm.CLASS_BEGIN
						elif cur_s == 'SUB':
							sm = NamespaceSm.SUB_BEGIN
						elif cur_s == 'FUNCTION':
							sm = NamespaceSm.FUNCTION_BEGIN
						elif cur_s == 'PROPERTY':
							sm = NamespaceSm.PROPERTY_BEGIN
					elif lxm.type == LexemeType.IDENTIFIER:
						var = ns.get_var(lxm.s)
						if var is None:
							var = Namespace.get_global_ns().add_implicit_var(lxm)
						else:
							var.add_ref(ns, lxm)
				elif sm == NamespaceSm.CLASS_BEGIN:
					if lxm.type != LexemeType.IDENTIFIER:
						raise Exception('Not a class identifier: {}'.format(repr(lxm)))
					lxm.type = LexemeType.CLASS
					ns = ns.add_class(lxm)
					sm = NamespaceSm.INIT
				elif sm == NamespaceSm.SUB_BEGIN:
					if lxm.type != LexemeType.IDENTIFIER:
						raise Exception('not a sub identifier: {}'.format(repr(lxm)))
					lxm.type = LexemeType.SUB
					ns = ns.add_sub(lxm)
					sm = NamespaceSm.INIT
				elif sm == NamespaceSm.FUNCTION_BEGIN:
					if lxm.type != LexemeType.IDENTIFIER:
						raise Exception('Not a function identifier: {}'.format(repr(lxm)))
					lxm.type = LexemeType.FUNCTION
					ns = ns.add_function(lxm)
					sm = NamespaceSm.INIT
				elif sm == NamespaceSm.PROPERTY_BEGIN:
					if cur_s == 'GET':
						sm == NamespaceSm.PROPERTY_GET
					elif cur_s == 'LET':
						sm == NamespaceSm.PROPERTY_LET
					elif cur_s == 'SET':
						sm == NamespaceSm.PROPERTY_SET
					else:
						raise Exception('Not a property type: {}'.format(repr(lxm)))
				elif sm in (NamespaceSm.PROPERTY_GET, NamespaceSm.PROPERTY_LET, NamespaceSm.PROPERTY_SET):
					if lxm.type != LexemeType.IDENTIFIER:
						raise Exception('Not a property identifier: {}'.format(repr(lxm)))
					lxm.type = LexemeType.PROPERTY
					ns = ns.add_property(lxm)
					sm = NamespaceSm.INIT
				else:
					raise Exception('Unhandled state: {}'.format(sm.name))

				prev_lxm = lxm
				prev_s = prev_lxm.s.upper()
		return top_ns
