<%@ language=vbscript %>
<% option explicit %>
<%
dim Option_TLS12 : Option_TLS12 = 2048
dim Option_TLS1 : Option_TLS1 = 512
dim Option_TLS : Option_TLS = 128

dim objHttp

Set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
'Set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
'objHttp.Option(9) = Option_TLS12
objHttp.open "POST", "https://howsmyssl.com/a/check", False
objHttp.Send
Response.Write objHttp.responseText 
Response.Write "<br>--------------------------<br>"
Set objHttp = Nothing 

Set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
'objHttp.Option(9) = Option_TLS12
objHttp.open "POST", "https://howsmyssl.com/a/check", False
objHttp.Send
Response.Write objHttp.responseText 
Response.Write "<br>--------------------------<br>"
Set objHttp = Nothing 

On Error Resume Next
Set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
objHttp.Option(9) = Option_TLS12
objHttp.open "POST", "https://howsmyssl.com/a/check", False
objHttp.Send
Response.Write objHttp.responseText 
Set objHttp = Nothing 
If ERR THEN response.write "ERR:"&Option_TLS12
On Error Goto 0

On Error Resume Next
Set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
objHttp.Option(9) = Option_TLS1
objHttp.open "POST", "https://howsmyssl.com/a/check", False
objHttp.Send
Response.Write objHttp.responseText 
Set objHttp = Nothing 
If ERR THEN response.write "ERR:"&Option_TLS1
On Error Goto 0

On Error Resume Next
Set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
objHttp.Option(9) = Option_TLS
objHttp.open "POST", "https://howsmyssl.com/a/check", False
objHttp.Send
Response.Write objHttp.responseText 
Set objHttp = Nothing 
If ERR THEN response.write "ERR:"&Option_TLS
On Error Goto 0
%>