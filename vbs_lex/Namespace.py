from enum import Enum, auto
from .Variable import *
from .Lexeme import LexemeType

import pdb

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
	ARGUMENT_LIST_EXPECT = auto()
	ARGUMENT_LIST_BEGIN = auto()
	ARGUMENT_MODIFIER = auto()
	ARGUMENT_IDENTIFIER = auto()
	ARGUMENT_LIST_COMMA = auto()
	ARGUMENT_LIST_END = auto()


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
		self.m_use_refs = []
		self.m_subobjects = []

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

	@property
	def classes(self):
		return self.m_classes

	@property
	def functions(self):
		return self.m_functions

	@property
	def subs(self):
		return self.m_subs

	@property
	def properties(self):
		return self.m_properties

	@property
	def use_refs(self):
		return self.m_use_refs

	def add_use_ref(self, lxm):
		self.m_use_refs.append(lxm)

	@property
	def subobjects(self):
		return self.m_subobjects

	def add_subobject(self, lxm):
		self.m_subobjects.append(lxm)


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

	def add_implicit_var(self, lxm, top_ns):
		var = Variable.new_implicit_def(self, lxm)
		top_ns.m_vars[lxm.s.upper()] = var
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
		sub_ns = Namespace(self, lxm)
		self.m_classes[lxm.s.upper()] = sub_ns
		return sub_ns
	def add_function(self, lxm):
		sub_ns = Namespace(self, lxm)
		self.m_functions[lxm.s.upper()] = sub_ns
		return sub_ns
	def add_sub(self, lxm):
		sub_ns = Namespace(self, lxm)
		self.m_subs[lxm.s.upper()] = sub_ns
		return sub_ns
	def add_property(self, lxm):
		sub_ns = Namespace(self, lxm)
		self.m_properties[lxm.s.upper()] = sub_ns
		return sub_ns


	@staticmethod
	def is_ignored_lxm(lxm):
		return lxm.type in (LexemeType.SPACE, LexemeType.NEWLINE, LexemeType.LINE_CONT, LexemeType.STATEMENT_CONCAT)

	@staticmethod
	def prev_key_lxm(lxm):
		p = lxm.prev
		while p is not None and Namespace.is_ignored_lxm(p):
			p = p.prev
		return p


	@staticmethod
	def process_lexemes(lxms, top_ns=None):
		if top_ns is None:
			top_ns = Namespace.new_top_ns()
		ns = top_ns
		prev_lxm = None
		prev_s = None
		sm = NamespaceSm.INIT
		for lxm in lxms:
			if Namespace.is_ignored_lxm(lxm):
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
					elif (prev_s == 'EXIT' and prev_lxm.type == LexemeType.KEYWORD) or lxm.type == LexemeType.DOT:
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
						if cur_s == ns.name.upper():
							ns.add_use_ref(lxm)
						else:
							if Namespace.is_lxm_match(Namespace.prev_key_lxm(lxm), LexemeType.DOT, '.'):
								top_ns.add_subobject(lxm)
							else:
								var = ns.get_var(lxm.s)
								if var is None:
									var = ns.add_implicit_var(lxm, top_ns)
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
					sm = NamespaceSm.ARGUMENT_LIST_EXPECT
				elif sm == NamespaceSm.FUNCTION_BEGIN:
					if lxm.type != LexemeType.IDENTIFIER:
						raise Exception('Not a function identifier: {}'.format(repr(lxm)))
					lxm.type = LexemeType.FUNCTION
					ns = ns.add_function(lxm)
					sm = NamespaceSm.ARGUMENT_LIST_EXPECT
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
					sm = NamespaceSm.ARGUMENT_LIST_EXPECT
				elif sm == NamespaceSm.ARGUMENT_LIST_EXPECT:
					if lxm.type != LexemeType.PAREN_BEGIN:
						sm = NamespaceSm.INIT
					else:
						sm = NamespaceSm.ARGUMENT_LIST_BEGIN
				elif sm in (NamespaceSm.ARGUMENT_LIST_BEGIN, NamespaceSm.ARGUMENT_LIST_COMMA):
					if lxm.type == LexemeType.KEYWORD:
						if cur_s not in ('BYREF', 'BYVAL'):
							raise Exception('Unhandled argument modifier {}'.format(cur_s))
						else:
							sm = NamespaceSm.ARGUMENT_MODIFIER
					elif lxm.type == LexemeType.IDENTIFIER:
						ns.add_var(lxm)
						sm = NamespaceSm.ARGUMENT_IDENTIFIER
					elif lxm.type == LexemeType.PAREN_END:
						sm = NamespaceSm.INIT
					else:
						raise Exception('Unhandled lexeme type {}'.format(lxm.type.name))
				elif sm == NamespaceSm.ARGUMENT_MODIFIER:
					if lxm.type == LexemeType.IDENTIFIER:
						ns.add_var(lxm)
						sm = NamespaceSm.ARGUMENT_IDENTIFIER
					else:
						raise Exception('Unhandled lexeme type {}'.format(lxm.type.name))
				elif sm == NamespaceSm.ARGUMENT_IDENTIFIER:
					if lxm.type == LexemeType.COMMA:
						sm = NamespaceSm.ARGUMENT_LIST_COMMA
					elif lxm.type == LexemeType.PAREN_END:
						sm = NamespaceSm.INIT
				else:
					raise Exception('Unhandled state: {}'.format(sm.name))

				prev_lxm = lxm
				prev_s = prev_lxm.s.upper()

		delvar_list = []
		for up_varname, var in top_ns.vars.items():
			if var.definition is not None:
				continue

			b_delvar = True
			if up_varname in top_ns.functions:
				#print('{} is a function'.format(var.name))
				for ref in var.refs:
					top_ns.functions[up_varname].add_use_ref(ref)
			elif up_varname in top_ns.subs:
				#print('{} is a sub'.format(var.name))
				for ref in var.refs:
					top_ns.subs[up_varname].add_use_ref(ref)
			elif up_varname in top_ns.classes:
				#print('{} is a class'.format(var.name))
				for ref in var.refs:
					top_ns.classes[up_varname].add_use_ref(ref)
			elif up_varname in top_ns.properties:
				#print('{} is a property'.format(var.name))
				for ref in var.refs:
					top_ns.properties[up_varname].add_use_ref(ref)
			else:
				b_delvar = False
				#print(var.name)

			if b_delvar:
				delvar_list.append(up_varname)

		for varname in delvar_list:
			del top_ns.vars[varname]

		return top_ns


	@staticmethod
	def is_lxm_match(lxm, lxm_type, lxm_s):
		if lxm is None:
			return False
		return lxm.type == lxm_type and lxm.s.upper() == lxm_s
