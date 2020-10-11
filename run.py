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

args = parser.parse_args()


def print_ns(base_ns, indent=0):
	pad_str = ' '*indent
	print('{}{}'.format(pad_str, str(base_ns)))

	local_vars = []
	foreign_vars = []
	for var_name, var in {k: base_ns.vars[k] for k in sorted(base_ns.vars.keys())}.items():
		if var.definition is None or var.definition.ns != base_ns:
			foreign_vars.append(var)
		else:
			local_vars.append(var)
		#print('{}* {}'.format(' '*(indent+2), var_name))

	if base_ns.parent is None:
		local_realm = 'local'
		foreign_realm = 'implicit'
	else:
		local_realm = 'local'
		foreign_realm = 'foreign'

	for var_locality, var_list in ((local_realm, local_vars), (foreign_realm, foreign_vars)):
		if len(var_list) < 1:
			continue
		print('{}  {}:'.format(pad_str, var_locality))
		for var in var_list:
			print('{}   * {}'.format(pad_str, var.name))

	if len(base_ns.vars) > 0:
		print()

	for ns in base_ns.classes.values():
		print_ns(ns, indent+2)
	for ns in base_ns.functions.values():
		print_ns(ns, indent+2)
	for ns in base_ns.subs.values():
		print_ns(ns, indent+2)
	for prop_dict in base_ns.properties.values():
		for prop_type, ns in prop_dict.items():
			print_ns(ns, indent+2)


#pdb.set_trace()
for f in args.files:
	lxms = lex_str(f.read(), fpath=f.name)

	stmts = Statement.statement_list_from_lexemes(lxms)
	#for stmt in stmts:
	#	print(stmt)

	file_ns = Namespace.from_statements(stmts)
	print_ns(file_ns)

	#for lxm in lxms:
	#	if lxm.type == LexemeType.IDENTIFIER:
	#		print(repr(lxm))
