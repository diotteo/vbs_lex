# vbs_lex

Python 3 library to statically analyze VBScript code.

## Running

To try out the code, use the runner:

```bash
$ ./run.py -v example_code/for_step.vbs
example_code/for_step.vbs:
 * a:
  in [global] at example_code/for_step.vbs:1:5:IDENTIFIER (LEXEME) a
  in [global] at example_code/for_step.vbs:1:9:IDENTIFIER (LEXEME) a
 * i (implicit):
  in [global] at example_code/for_step.vbs:3:5:IDENTIFIER (LEXEME) i
  in [global] at example_code/for_step.vbs:5:20:IDENTIFIER (LEXEME) i
```

## Tests

```bash
$ python -m unittest tests/tests.py
```
