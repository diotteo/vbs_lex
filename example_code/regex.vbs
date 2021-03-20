' https://www.vbsedit.com/html/05f9ee2e-982f-4727-839e-b1b8ed696d0a.asp
' https://www.vbsedit.com/html/ab0766e1-7037-45ed-aa23-706f58358c0e.asp

dim a
set a = new RegExp
a.Pattern = "foo"
set matches = a.Execute("foobar")
for each m in matches
	wscript.echo "m = " & m.value
next
