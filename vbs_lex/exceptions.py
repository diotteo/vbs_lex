"""vbs_lex exceptions

Classes:

    LexemeException
"""

class LexemeException(Exception):
    def __init__(self, lxm, message):
        self.message = '{}: {}'.format(lxm, message)
