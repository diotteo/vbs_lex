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

def print_var_refs(ns):
	for varname, var in ns.vars.items():
		print('* {}:'.format(var.name))
		for ref in var.refs:
			print('  in {} at {}'.format(ref.ns, repr(ref.lxm)))


#pdb.set_trace()
for f in args.files:
	lxms = lex_str(f.read(), fpath=f.name)

	stmts = Statement.statement_list_from_lexemes(lxms)
	#for stmt in stmts:
	#	print(stmt)

	file_ns = Namespace.from_statements(stmts)
	#file_ns.print_ns()
	print_var_refs(file_ns)

	#for lxm in lxms:
	#	if lxm.type == LexemeType.IDENTIFIER:
	#		print(repr(lxm))
