#! /usr/bin/env python3

import sys, os
import argparse
import pdb

parser = argparse.ArgumentParser(description='my lexer')
parser.addArgument('fpaths', nargs='+', type=argparse.FileType('r'))

args = parser.parse_args()

for fpath in args.fpaths:
	with open(fpath) as f:
		for line in f:
			print(line)
