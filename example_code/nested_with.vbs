class TestObject
	public a
	public b
	public c
	public d
end class

set o = new TestObject
set oo = new TestObject
'with new TestObject
with o
	set .a = oo
	with .a
		.a = 5
		.b = 6
	end with
	.b = 2
	.c = 3
	.d = 4
end with

wscript.echo o.b
