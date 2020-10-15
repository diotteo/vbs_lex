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

	@property
	def top_ns(self):
		top = self
		while top.parent is not None:
			top = top.parent
		return top


	@staticmethod
	def new_top_ns():
		return GlobalNamespace(None, None)

	@staticmethod
	def reset_global_ns():
		Namespace.global_ = Namespace.new_top_ns()
		return Namespace.global_

	def add_var(self, lxm):
		var = Variable(self, lxm)
		self.m_vars[lxm.s.upper()] = var
		return var

	def add_implicit_var(self, lxm):
		var = Variable.new_implicit_def(self, lxm)
		self.top_ns.m_vars[lxm.s.upper()] = var
		return var

	def get_var(self, s):
		up_s = s.upper()
		if up_s in self.m_vars:
			return self.m_vars[up_s]
		elif self.parent is None:
			return None
		else:
			return self.parent.get_var(up_s)

	def add_var_ref_or_implicit(self, lxm):
		var = self.get_var(lxm.s)
		if var is None:
			var = self.add_implicit_var(lxm)
		else:
			var.add_ref(self, lxm)
		return var

	def get_proc(self, s):
		up_s = s.upper()
		if up_s in self.m_functions:
			return self.m_functions[up_s]
		elif up_s in self.m_subs:
			return self.m_subs[up_s]
		elif up_s in self.m_properties:
			return self.m_properties[up_s]
		elif self.parent is None:
			return None
		else:
			return self.parent.get_proc(up_s)

	def add_class(self, lxm):
		sub_ns = Class(self, lxm)
		self.m_classes[lxm.s.upper()] = sub_ns
		return sub_ns
	def add_function(self, lxm):
		sub_ns = Function(self, lxm)
		self.m_functions[lxm.s.upper()] = sub_ns
		return sub_ns
	def add_sub(self, lxm):
		sub_ns = Sub(self, lxm)
		self.m_subs[lxm.s.upper()] = sub_ns
		return sub_ns
	def add_property(self, type_, lxm):
		sub_ns = Property(self, lxm)
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


	def parse_def_arglist(self, stmt, start):
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

		#procedure calls can happen in:
		# * assignments (right side)
		# * procedure call arguments
		# * select case statements
		# * if, else if statements
		# * do while, do until statements
		# * for statements
		# * new statements
		# * with statements
		# * redim statements (inside parentheses)
		potential_identifier_uses = []

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
				ns.parse_def_arglist(stmt, idx+2)
			elif stmt.type == StatementType.FUNCTION_BEGIN:
				idx = Namespace.get_type_lexeme_idx(LexemeType.KEYWORD, 'FUNCTION', stmt)
				lxm = stmt.lxms[idx+1]
				Namespace.set_identifier_lexeme_type(lxm, LexemeType.FUNCTION)
				ns = ns.add_function(lxm)
				ns.parse_def_arglist(stmt, idx+2)
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
				ns.parse_def_arglist(stmt, idx+3)
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
				if stmt.type == StatementType.REDIM:
					potential_identifier_uses.append((stmt, ns))

				for lxm in stmt.lxms[1:]:
					if lxm.type == LexemeType.IDENTIFIER:
						ns.add_var(lxm)
			elif stmt.type in (
					StatementType.VAR_ASSIGNMENT,
					StatementType.OBJECT_ASSIGNMENT,
					):
				potential_identifier_uses.append((stmt, ns))

				start_idx = 0 if stmt.type == StatementType.VAR_ASSIGNMENT else 1
				end_idx = Namespace.get_type_lexeme_idx(LexemeType.OPERATOR, '=', stmt)
				lxm = stmt.lxms[start_idx]
				if lxm.s.upper() == ns.name.upper():
					ns.add_use_ref(lxm)
				elif stmt.lxms[start_idx+1].type == LexemeType.DOT:
					var = ns.get_var(lxm.s)
					var.add_ref(ns, lxm)
				else:
					ns.add_var_ref_or_implicit(lxm)
			elif stmt.type in (
					StatementType.PROC_CALL,
					StatementType.IMPLICIT_PROC_CALL,
					StatementType.DO_LOOP_BEGIN,
					StatementType.FOR_LOOP_BEGIN,
					StatementType.WHILE_LOOP_BEGIN,
					StatementType.IF_BEGIN,
					StatementType.IF_ELSE_IF,
					StatementType.SELECT_BEGIN,
					StatementType.WITH_BEGIN,
					#StatementType.UNASSIGNED_ARITHMETIC,
					#StatementType.UNASSIGNED_NEW,
					):
				potential_identifier_uses.append((stmt, ns))
			else:
				#Ignored statements
				pass

		Namespace.process_potential_uses(potential_identifier_uses)
		return top_ns


	@staticmethod
	def identifiers_from_rvalue_list(lxms, start=0, end=None):
		identifiers = []
		cur_identifier_last_type = None
		paren_lvl = 0
		i = start
		for i, lxm in enumerate(lxms[start:end], start):
			if lxm.type == LexemeType.PAREN_BEGIN:
				cur_identifier_last_type = None
				paren_lvl += 1
			elif lxm.type == LexemeType.PAREN_END:
				cur_identifier_last_type = None
				paren_lvl -= 1
				if paren_lvl < 0:
					break
			elif cur_identifier_last_type is not None:
				if lxm.type == LexemeType.IDENTIFIER:
					#implicit proc calls only separate the first argument from
					#the procedure name by a space
					if cur_identifier_last_type == LexemeType.IDENTIFIER:
						identifiers.append(lxm)
					cur_identifier_last_type = lxm.type
				elif lxm.type == LexemeType.DOT:
					cur_identifier_last_type = lxm.type
				else:
					cur_identifier_last_type = None
			elif lxm.type == LexemeType.IDENTIFIER:
				identifiers.append(lxm)
				cur_identifier_last_type = lxm.type

		return identifiers, i


	@staticmethod
	def process_potential_uses(potential_identifier_uses):
		#Go through potential_identifier_uses,
		#
		#Parsing each stmt for identifiers and trying to match them to
		#  their respective definition.
		#For vars/objects, we should be able to ns.getvar() them:
		#  If they don't exist, they are implicit global variables
		#For classes: always look at the global namespace, that's where they're allowed
		#For functions and subs, they need to be defined in our namespace chain
		#For properties, they need to be defined in our direct parent
		#  (properties are only valid in classes, classes are only valid in
		#  global namespace and nested functions don't exist in VB)

		for stmt, ns in potential_identifier_uses:
			identifiers = []
			if stmt.type == StatementType.REDIM:
				i = 0
				while i < len(stmt.lxms):
					lxm = stmt.lxms[i]
					if lxm.type == LexemeType.IDENTIFIER and stmt.lxms[i+1].type == LexemeType.PAREN_BEGIN:
						cur_identifiers, i = Namespace.identifiers_from_rvalue_list(stmt.lxms, i+2)
						identifiers.extend(cur_identifiers)
						assert stmt.lxms[i].type == LexemeType.PAREN_END
					i += 1
				pass
			elif stmt.type in (
					StatementType.VAR_ASSIGNMENT,
					StatementType.OBJECT_ASSIGNMENT,
					):
				eq_idx = Namespace.get_type_lexeme_idx(LexemeType.OPERATOR, '=', stmt)
				identifiers, i = Namespace.identifiers_from_rvalue_list(stmt.lxms, eq_idx+1)
			elif stmt.type in (
					StatementType.PROC_CALL,
					StatementType.IMPLICIT_PROC_CALL,
					):
				#Skip 'call' keyword
				start = int(stmt.type == StatementType.PROC_CALL)

				if len(stmt.lxms) > start+1 and stmt.lxms[start+1].type == LexemeType.DOT:
					lxm = stmt.lxms[start]
					var = ns.get_var(lxm.s)

					#It shouldn't be possible for var to be None since it is
					#not possible to dot-access a field on an empty var)
					b_is_special_object = False
					if var is None:
						if lxm.type == LexemeType.SPECIAL_OBJECT:
							#FIXME: add those as variables to top_ns?
							b_is_special_object = True
						else:
							#Could be an object defined in an included file
							ns.top_ns.add_implicit_var(lxm)

					if var is not None:
						var.add_ref(ns, lxm)

					expected_type = LexemeType.IDENTIFIER
					last = start+2
					for i, lxm in enumerate(stmt.lxms[start+2:], start+2):
						if lxm.type != expected_type:
							#TODO: add proc call to unknown call list?
							break
						expected_type = LexemeType.IDENTIFIER if expected_type == LexemeType.DOT else LexemeType.DOT
						last = i
				else:
					lxm = stmt.lxms[start]
					proc = ns.get_proc(lxm.s)
					if proc is None:
						#Implicitly-defined procedure
						#This can happen whenever we have eval'ed statements (such as including another file through ExecuteGlobal)
						proc = ns.top_ns.add_implicit_proc(lxm)
						identifiers, i = Namespace.identifiers_from_rvalue_list(stmt.lxms, start+1)

					#We have to figure out the right property
					elif isinstance(proc, dict):
						let_prop = proc.get('LET')
						set_prop = proc.get('SET')
						if (let_prop is None) == (set_prop is None):
							raise Exception('Wtf? property has both set and let? {}'.format(lxm))
						elif let_prop is not None:
							proc = let_prop
						else:
							proc = set_prop
					proc.add_use_ref(lxm)
					last = start

				identifiers, i = Namespace.identifiers_from_rvalue_list(stmt.lxms, last+1)
			elif stmt.type == StatementType.FOR_LOOP_BEGIN:
				b_is_foreach = stmt.lxms[1].s.upper() == 'EACH'
				if b_is_foreach:
					start = 2
				else:
					start = 1

				#lxm must be a simple var (and is assignment)
				lxm = stmt.lxms[start]
				ns.add_var_ref_or_implicit(lxm)

				# * for each [var] in [expr]
				# * for [var] = [expr] to [expr]
				#for our purposes, we don't need to care about the distinction,
				#the 'to' keyword will be ignored correctly
				identifiers, i = Namespace.identifiers_from_rvalue_list(stmt.lxms, start+2)
			elif stmt.type in (
					StatementType.DO_LOOP_BEGIN,
					StatementType.WHILE_LOOP_BEGIN,
					):
				start = 1 + int(stmt.type == StatementType.DO_LOOP_BEGIN)
				identifiers, i = Namespace.identifiers_from_rvalue_list(stmt.lxms, start)
			elif stmt.type in (
					StatementType.IF_BEGIN,
					StatementType.IF_ELSE_IF,
					StatementType.SELECT_BEGIN,
					StatementType.WITH_BEGIN,
					StatementType.RANDOMIZE,
					):

					if stmt.type in (
							StatementType.IF_ELSE_IF,
							StatementType.SELECT_BEGIN,
							):
						start = 2
					else:
						start = 1

					identifiers, i = Namespace.identifiers_from_rvalue_list(stmt.lxms, start)

			#TODO: erase stmt, execute, executeGlobal
			else:
				raise Exception('Unhandled potential-identifier-use-statement type:{}'.format(stmt.type))

			if len(identifiers) > 0:
				Namespace.add_identifiers_use_refs(ns, identifiers)


	@staticmethod
	def add_identifiers_use_refs(ns, identifiers):
		for lxm in identifiers:
			#FIXME: If an undefined identifier is only used in the rvalue of an assignment, we cannot know if it is a procedure
			#In fact, since it's possible to use getref() to assign a procedure reference to a variable (and make a procedure call with that),
			#the distinction becomes blurry
			proc = ns.get_proc(lxm.s)
			if proc is None:
				ns.add_var_ref_or_implicit(lxm)
			else:
				if isinstance(proc, dict):
					prop_get = proc.get('GET')
					if prop_get is None:
						raise Exception('Property use but not get?! {}'.format(lxm))
					proc = prop_get
				proc.add_use_ref(lxm)


	def print_ns(self, indent=0):
		pad_str = ' '*indent
		print('{}{}'.format(pad_str, str(self)))

		local_vars = []
		foreign_vars = []
		for var_name, var in {k: self.vars[k] for k in sorted(self.vars.keys())}.items():
			if var.definition is None or var.definition.ns != self:
				foreign_vars.append(var)
			else:
				local_vars.append(var)

		local_realm = 'local'
		if self.parent is None:
			foreign_realm = 'implicit'
		else:
			foreign_realm = 'foreign'

		for var_locality, var_list in ((local_realm, local_vars), (foreign_realm, foreign_vars)):
			if len(var_list) < 1:
				continue
			print('{}  {}:'.format(pad_str, var_locality))
			for var in var_list:
				print('{}   * {}'.format(pad_str, var.name))

		if len(self.vars) > 0:
			print()

		for ns in self.classes.values():
			ns.print_ns(indent+2)
		for ns in self.functions.values():
			ns.print_ns(indent+2)
		for ns in self.subs.values():
			ns.print_ns(indent+2)
		for prop_dict in self.properties.values():
			for prop_type, ns in prop_dict.items():
				ns.print_ns(indent+2)


class GlobalNamespace(Namespace):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self.m_implicit_procs = {}
		self.m_vb_procs = {}
		self.m_vb_objects = {}
		self.m_vb_values = {}

	def add_implicit_proc(self, lxm):
		sub_ns = Proc(self, lxm)
		self.m_implicit_procs[lxm.s.upper()] = sub_ns
		return sub_ns


class ScopedNamespaceBase(Namespace):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)

class Class(ScopedNamespaceBase):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)

class CallableNamespace(ScopedNamespaceBase):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)

class Property(CallableNamespace):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)

class Proc(CallableNamespace):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)

class Function(Proc):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)

class Sub(Proc):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
