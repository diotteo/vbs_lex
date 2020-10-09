sub foo(a, ByVal b, ByRef c)
	dim d

	e = a+b+c + d*2

end sub

function bar(a, b)
	bar = a+b
end function

call foo(1, 2, 3)

bar a, b
