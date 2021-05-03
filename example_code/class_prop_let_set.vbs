class Foo
	public sub Class_Initialize()
	end sub

	public sub Class_Terminate()
	end sub

	public function myfunc(a, b)
		dim foo
		some_letter = 3
		foo = some_getter

		myfunc = a + b
	end function

	private m_priv_var

	public property get some_getter()
		some_getter = 3
	end property

	private m_a
	public property let some_letter(a)
		m_a = a
	end property

	private m_c
	public property set some_setter(a)
		set m_c = a
	end property

	private m_b
	public property get b()
		b = m_b
	end property

	public property let b(bb)
		m_b = bb
	end property

	private property get aa()
		aa = "foo"
	end property

	property get bb()
		bb = "bar"
	end property
end class

set a = new Foo
wscript.echo a.myfunc(1, 2)
