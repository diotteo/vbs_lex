#! /usr/bin/env python3

import sys, os
import argparse
import pdb

from vbs_lex.Lexer import lex_str
from vbs_lex.Namespace import Namespace
from vbs_lex.Lexeme import LexemeType
from vbs_lex.Statement import Statement

parser = argparse.ArgumentParser(description='my lexer')
parser.add_argument('files', nargs='+', type=argparse.FileType('r', encoding='utf-8-sig'))
parser.add_argument('--statements', '-s', action='store_true')
parser.add_argument('--namespaces', '-n', action='store_true')
parser.add_argument('--lexemes', '-l', action='store_true')
parser.add_argument('--variables', '-v', action='store_true')
parser.add_argument('--implicit-decls', '-i', action='store_true')

args = parser.parse_args()


def get_var_ref_lines(ns):
	lines = []
	for varname, var in ns.vars.items():
		lines.append(' * {}:'.format(var.name))
		for ref in var.refs:
			lines.append('  in {} at {}'.format(ref.ns, repr(ref.lxm)))
	return lines


def get_global_decl_lines(ns, decl_type='all'):
	"""decl_type can be 'all', 'implicit' or 'explicit'"""
	b_do_implicit = decl_type in ('all', 'implicit')
	b_do_explicit = decl_type in ('all', 'explicit')

	lines = []
	top_ns = ns.top_ns
	for var in top_ns.vars.values():
		s = ' * {}'.format(var.name)

		b_is_implicit = var.definition is None
		if b_is_implicit and b_do_implicit:
			lines.append(s + ' (implicit) [{}:{}:{}]'.format(var.refs[0].lxm.fpath, var.refs[0].lxm.line, var.refs[0].lxm.col))
		elif not b_is_implicit and b_do_explicit:
			lines.append(s)
	return lines


#pdb.set_trace()
for f in args.files:
	out_lines = []
	lxms = list(lex_str(f.read(), fpath=f.name))
	if args.lexemes:
		for lxm in lxms:
			#	if lxm.type == LexemeType.IDENTIFIER:
			#		print(repr(lxm))
			out_lines.append(str(lxm))

	stmts = Statement.statement_list_from_lexemes(lxms)
	if args.statements:
		for stmt in stmts:
			out_lines.append(str(stmt))

	file_ns = Namespace.from_statements(stmts)
	if args.namespaces:
		out_lines.extend(file_ns.get_ns_lines())

	if args.variables:
		out_lines.extend(get_var_ref_lines(file_ns))

	if args.variables or args.implicit_decls:
		out_lines.extend(get_global_decl_lines(file_ns))

	if len(out_lines) > 0:
		print('{}:\n{}'.format(f.name, '\n'.join(out_lines)))
