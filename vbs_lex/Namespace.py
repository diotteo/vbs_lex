from enum import Enum, auto
from .Variable import *
from .Lexeme import LexemeType
from .Statement import StatementType

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
	def add_property(self, type_, lxm):
		sub_ns = Namespace(self, lxm)
		self.m_properties.setdefault(lxm.s.upper(), {})[type_] = sub_ns
		return sub_ns


	@staticmethod
	def get_type_lexeme_idx(needle_lxm_type, needle_str, stmt, start=0, end=None):
		needle_up_s = needle_str.upper()
		for idx, lxm in enumerate(stmt.lxms[start:end], start):
			if lxm.type == needle_lxm_type and lxm.s.upper() == needle_up_s:
				return idx
		raise ValueError('{} not found in statement'.format(needle_str))


	@staticmethod
	def set_identifier_lexeme_type(lxm, type_):
		if lxm.type != LexemeType.IDENTIFIER:
			raise TypeError('Not an identifier {}'.format(lxm))
		lxm.type = type_


	def parse_arglist(self, stmt, start):
		sm = NamespaceSm.ARGUMENT_LIST_EXPECT

		for idx, lxm in enumerate(stmt.lxms[start:], start):
			next_state = None
			if sm == NamespaceSm.ARGUMENT_LIST_EXPECT:
				if lxm.type == LexemeType.PAREN_BEGIN:
					next_state = NamespaceSm.ARGUMENT_LIST_BEGIN
			elif sm in (NamespaceSm.ARGUMENT_LIST_BEGIN, NamespaceSm.ARGUMENT_LIST_COMMA):
				if lxm.type == LexemeType.KEYWORD:
					if lxm.s.upper() in ('BYREF', 'BYVAL'):
						next_state = NamespaceSm.ARGUMENT_MODIFIER
				elif lxm.type == LexemeType.IDENTIFIER:
					self.add_var(lxm)
					next_state = NamespaceSm.ARGUMENT_IDENTIFIER
				elif lxm.type == LexemeType.PAREN_END:
					assert idx+1 == len(stmt.lxms)
					return
			elif sm == NamespaceSm.ARGUMENT_MODIFIER:
				if lxm.type == LexemeType.IDENTIFIER:
					self.add_var(lxm)
					next_state = NamespaceSm.ARGUMENT_IDENTIFIER
			elif sm == NamespaceSm.ARGUMENT_IDENTIFIER:
				if lxm.type == LexemeType.COMMA:
					next_state = NamespaceSm.ARGUMENT_LIST_COMMA
				elif lxm.type == LexemeType.PAREN_END:
					if idx+1 != len(stmt.lxms):
						pdb.set_trace()
						raise Exception('paren end should be end of statement: {}'.format(repr(stmt.lxms[0])))
					#assert idx+1 == len(stmt.lxms)
					return
			if next_state is None:
				raise Exception('Expected transtion from {}, got {}'.format(sm, lxm))
			sm = next_state


	@staticmethod
	def from_statements(stmts, top_ns=None):
		if top_ns is None:
			top_ns = Namespace.new_top_ns()
		ns = top_ns

		for stmt in stmts:
			if stmt.type == StatementType.CLASS_BEGIN:
				idx = Namespace.get_type_lexeme_idx(LexemeType.KEYWORD, 'CLASS', stmt)
				lxm = stmt.lxms[idx+1]
				Namespace.set_identifier_lexeme_type(lxm, LexemeType.CLASS)
				ns = ns.add_class(lxm)
			elif stmt.type == StatementType.SUB_BEGIN:
				idx = Namespace.get_type_lexeme_idx(LexemeType.KEYWORD, 'SUB', stmt)
				lxm = stmt.lxms[idx+1]
				Namespace.set_identifier_lexeme_type(lxm, LexemeType.SUB)
				ns = ns.add_sub(lxm)
				ns.parse_arglist(stmt, idx+2)
			elif stmt.type == StatementType.FUNCTION_BEGIN:
				idx = Namespace.get_type_lexeme_idx(LexemeType.KEYWORD, 'FUNCTION', stmt)
				lxm = stmt.lxms[idx+1]
				Namespace.set_identifier_lexeme_type(lxm, LexemeType.FUNCTION)
				ns = ns.add_function(lxm)
				ns.parse_arglist(stmt, idx+2)
			elif stmt.type in (
					StatementType.PROPERTY_GET_BEGIN,
					StatementType.PROPERTY_LET_BEGIN,
					StatementType.PROPERTY_SET_BEGIN,
					):
				idx = Namespace.get_type_lexeme_idx(LexemeType.KEYWORD, 'PROPERTY', stmt)
				lxm = stmt.lxms[idx+2]

				stmt_type_name_list = stmt.type.name.split('_')
				lxm_type_str = '_'.join(stmt_type_name_list[:2])
				prop_type = stmt_type_name_list[1]
				Namespace.set_identifier_lexeme_type(lxm, LexemeType[lxm_type_str])
				ns = ns.add_property(prop_type, lxm)
				ns.parse_arglist(stmt, idx+3)
			elif stmt.type in (
					StatementType.CLASS_END,
					StatementType.SUB_END,
					StatementType.FUNCTION_END,
					StatementType.PROPERTY_END,
					):
				ns = ns.parent

			elif stmt.type == StatementType.CONST_DECLARE:
				lxm = stmt.lxms[1]
				ns.add_var(lxm)
			elif stmt.type in (
					StatementType.DIM,
					StatementType.REDIM,
					StatementType.FIELD_DECLARE,
					):
				for lxm in stmt.lxms[1:]:
					if lxm.type == LexemeType.IDENTIFIER:
						ns.add_var(lxm)
			elif stmt.type in (
					StatementType.VAR_ASSIGNMENT,
					StatementType.OBJECT_ASSIGNMENT,
					):
				start_idx = 0 if stmt.type == StatementType.VAR_ASSIGNMENT else 1
				end_idx = Namespace.get_type_lexeme_idx(LexemeType.OPERATOR, '=', stmt)
				lxm = stmt.lxms[start_idx]
				if lxm.s.upper() == ns.name.upper():
					ns.add_use_ref(lxm)
				elif stmt.lxms[start_idx+1].type == LexemeType.DOT:
					var = ns.get_var(lxm.s)
					var.add_ref(ns, lxm)
				else:
					var = ns.get_var(lxm.s)
					if var is None:
						var = ns.add_implicit_var(lxm, top_ns)
					else:
						var.add_ref(ns, lxm)
			else:
				#Ignored statements
				pass

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
