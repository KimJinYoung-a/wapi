<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, mallid, orderState, paramData, action, ErrCode, isSuccessYn, idx, ErrMsg
Dim sqlStr, procStr
itemid			= request("itemid")
mallid			= request("mallid")
orderState		= request("orderState")
idx				= request("idx")
ErrCode			= request("ErrCode")
action			= request("action")
ErrMsg			= request("ErrMsg")

If ErrCode = "OK" Then
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE [db_etcmall].[dbo].[tbl_outmall_diffOrder] SET " & vbcrlf
	sqlStr = sqlStr & " isOk = 'Y' "	 & vbcrlf
	sqlStr = sqlStr & " WHERE idx = '"&idx&"'  "
	dbget.Execute sqlStr
	Call SugiQueLogInsert(mallid, action, itemid, "OK", "S_OK", "system")
End If
response.write "S_OK"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->