<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% session.CodePage = "65001" %>
<% Server.ScriptTimeOut = 1200 %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim userid, username
	Dim rpatype, rpatitle, rpacontents, vquery, rpaissuccess

	rpatype = request("type")
	rpatitle = request("title")
	rpacontents = request("contents")
	rpaissuccess = request("issuccess")


'	userid = request("userid")
'	username = request("username")

'	response.write "userid="&userid&"<br>username="&username&"<br>"&Request.ServerVariables("REMOTE_ADDR")

	If trim(rpatype) = "" or trim(rpatitle) = "" or trim(rpacontents) = "" Then
		response.write "error"
		response.end
	end if


	vquery = " INSERT INTO db_sitemaster.dbo.tbl_RpaSuccessMessageReceive "
	vquery = vquery &" (rpatype, rpatitle, rpacontents, rpaissuccess, regdate) VALUES "
	vquery = vquery &" ('"&rpatype&"', '"&rpatitle&"', '"&rpacontents&"', '"&rpaissuccess&"', getdate())"
	dbget.Execute vquery	

	response.write "ok"
	response.end
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->