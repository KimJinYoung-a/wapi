<%@ language=vbscript %>
<% option explicit %>
<%
Dim isTest : isTest = False
If isTest <> True Then
	Response.ContentType = "application/json"
	Response.AddHeader "Accept", "application/json"
	Response.Charset = "UTF-8"
Else
	Response.Charset = "UTF-8"
End If
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/aspJSON1.17.asp"-->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim IS_TEST_MODE : IS_TEST_MODE = False
Dim refIP : refIP = Request.ServerVariables ("REMOTE_ADDR")
Dim sqlStr
Dim obj, sData

sData = BinaryToText(Request.BinaryRead(request.TotalBytes), "UTF-8")
sqlStr = ""
sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_tmp_GSOrderData "
sqlStr = sqlStr & " (regdate, refip, dataStr) VALUES "
sqlStr = sqlStr & " (getdate(),'"&refIP&"','" & html2db(sData) & "')"
dbget.Execute sqlStr

Call GetOrderFromJson_gseshop_Recv(sData)
If (IS_TEST_MODE = False) Then
	Set obj = jsObject()
		obj("resultCd") = "S"
		obj("resultMsg") = ""
		response.write obj.jsString
	Set obj = nothing
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->