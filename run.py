#! /usr/bin/env python3

import sys, os
import argparse
import pdb

from Lexer import lex_str
from Lexeme import LexemeType

parser = argparse.ArgumentParser(description='my lexer')
parser.add_argument('files', nargs='+', type=argparse.FileType('r', encoding='utf-8-sig'))

args = parser.parse_args()

for f in args.files:
	lxms = lex_str(f.read(), fpath=f.name)

	prev_lxm = None
	for lxm in lxms:
		if lxm.type in (LexemeType.SPACE, LexemeType.NEWLINE):
			pass
		elif lxm.type == LexemeType.IDENTIFIER:
			if prev_lxm.type == LexemeType.DOT:
				pass
			else:
				print(repr(lxm))
		prev_lxm = lxm
