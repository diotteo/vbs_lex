class Foo
	public sub Class_Initialize()
	end sub

	public sub Class_Terminate()
	end sub

	public function myfunc(a, b)
		myfunc = a + b
	end function
end class

set a = new Foo
wscript.echo a.myfunc(1, 2)
