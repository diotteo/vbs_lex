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

def print_var_refs(ns):
	for varname, var in ns.vars.items():
		print('* {}:'.format(var.name))
		for ref in var.refs:
			print('  in {} at {}'.format(ref.ns, repr(ref.lxm)))

def print_globals(ns):
	top_ns = ns.top_ns
	for var in top_ns.vars.values():
		print('* {}{}'.format(var.name, ' (implicit)' if var.definition is None else ''))


#pdb.set_trace()
for f in args.files:
	print('{}:'.format(f.name))

	lxms = lex_str(f.read(), fpath=f.name)
	if args.lexemes:
		for lxm in lxms:
			#	if lxm.type == LexemeType.IDENTIFIER:
			#		print(repr(lxm))
			print(lxm)

	stmts = Statement.statement_list_from_lexemes(lxms)
	if args.statements:
		for stmt in stmts:
			print(stmt)

	file_ns = Namespace.from_statements(stmts)
	if args.namespaces:
		file_ns.print_ns()

	if args.variables:
		print_var_refs(file_ns)

	if args.variables or args.implicit_decls:
		print_globals(file_ns)
