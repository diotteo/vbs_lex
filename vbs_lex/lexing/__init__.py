"""VBScript lexical analysis: turn a string into a series of Lexemes

Classes:

    LexemeType
    Lexeme
    ExternalLexeme

Functions:
    lex_file(fpath)
    lex_str(s, fpath=None)
"""

from .classes import ExternalLexeme, Lexeme
from .types import LexemeType
from .lexer import lex_str, lex_file
