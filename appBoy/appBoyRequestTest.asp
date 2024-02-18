<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% session.CodePage = "65001" %>
<% Server.ScriptTimeOut = 1200 %>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim userid, username

	userid = request("userid")
	username = request("username")

	response.write "userid="&userid&"<br>username="&username&"<br>"&Request.ServerVariables("REMOTE_ADDR")

%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->