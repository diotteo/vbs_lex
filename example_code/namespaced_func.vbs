class Foo
	public function bar()
		wscript.echo "Foo.bar"
	end function
end class

function bar()
	wscript.echo "Top bar"
end function

set a = new Foo
a.bar()
wscript.echo a.bar()
