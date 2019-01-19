<%
CountFile=Server.MapPath("txtcounter.txt")
SetFileObject=Server.CreateObject("Scripting.FileSystemObject")
SetOut=FileObject.OpenTextFile(CountFile,1,FALSE,FALSE)
counter=Out.ReadLine
Out.Close
SETFileObject=Server.CreateObject("Scripting.FileSystemObject")
SetOut=FileObject.CreateTextFile(CountFile,TRUE,FALSE)
Application.lock
counter= counter + 1
Out.WriteLine(counter)
Application.unlock
Response.Write"document.write("&counter&")"
'为了在页面正确显示计数器的值，调用VBScript函数Document.write
Out.Close
%>
