#! /usr/bin/env python3

import sys, os
import argparse
import pdb

from Lexer import lex_str

parser = argparse.ArgumentParser(description='my lexer')
parser.add_argument('files', nargs='+', type=argparse.FileType('r', encoding='utf-8-sig'))

args = parser.parse_args()

for f in args.files:
	lxms = lex_str(f.read())
	print('\n'.join((str(x) for x in lxms)))
