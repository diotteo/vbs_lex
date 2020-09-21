#! /usr/bin/env python3

import sys, os
import argparse
import pdb

from Lexer import lex_str
from Lexeme import LexemeType
from Namespace import *

parser = argparse.ArgumentParser(description='my lexer')
parser.add_argument('files', nargs='+', type=argparse.FileType('r', encoding='utf-8-sig'))

args = parser.parse_args()

for f in args.files:
	lxms = lex_str(f.read(), fpath=f.name)

	ns = Namespace.get_global_ns()
	ns_list = [ns,]
	prev_lxm = None
	sm = NamespaceSm.INIT
	for lxm in lxms:
		if lxm.type in (LexemeType.SPACE, LexemeType.NEWLINE, LexemeType.LINE_CONT, LexemeType.STATEMENT_CONCAT):
			pass
		else:
			prev_s = None
			if prev_lxm is not None:
				prev_s = prev_lxm.s.upper()
			cur_s = lxm.s.upper()

			if sm == NamespaceSm.INIT:
				if prev_s == 'END':
					b_ns_end = True
					if cur_s == 'CLASS':
						sm = NamespaceSm.CLASS_END
					elif cur_s == 'SUB':
						sm = NamespaceSm.SUB_END
					elif cur_s == 'FUNCTION':
						sm = NamespaceSm.FUNCTION_END
					elif cur_s == 'PROPERTY':
						sm = NamespaceSm.PROPERTY_END

					else: #end if, end for, etc
						b_ns_end = False

					if b_ns_end:
						ns = ns.parent
						sm = NamespaceSm.INIT
				elif prev_s in ('EXIT', '.'):
					pass
				elif cur_s == 'CLASS':
					sm = NamespaceSm.CLASS_BEGIN
				elif cur_s == 'SUB':
					sm = NamespaceSm.SUB_BEGIN
				elif cur_s == 'FUNCTION':
					sm = NamespaceSm.FUNCTION_BEGIN
				elif cur_s == 'PROPERTY':
					sm = NamespaceSm.PROPERTY_BEGIN
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

	for lxm in lxms:
		if lxm.type == LexemeType.IDENTIFIER:
			print(repr(lxm))
