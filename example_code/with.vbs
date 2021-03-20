class TestObject
	public a
	public b
	public c
	public d
end class

set o = new TestObject
'with new TestObject
with o
	.a = 1
	.b = 2
	.c = 3
	.d = 4
end with

wscript.echo o.a
