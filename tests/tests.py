import unittest

class TestPass(unittest.TestCase):
    def test_lex_for_step(self):
        from src.vbs_lex.lexing import lex_str, LexemeType, Lexeme
        with open('example_code/for_step.vbs') as f:
            s = f.read()
            lxms = lex_str(s, f.name)
        self.assertTrue(Lexeme.str_from_lexemes(lxms), s)


if __name__ == '__main__':
    unittest.main()
