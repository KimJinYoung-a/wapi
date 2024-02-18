<%@ language=vbscript %><% option explicit %><?xml version="1.0"  encoding="euc-kr"?>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<% Response.contentType = "text/xml; charset=euc-kr" %>
<!-- #include virtual="/lib/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<response>
<%
dim mode
dim param1, param2, param3

mode = request("mode")
param1 = request("param1")
param2 = request("param2")
param3 = request("param3")

dim sqlStr

if mode="cdl" then
	sqlStr = "select top 100 * from [db_item].[dbo].tbl_Cate_large"
	sqlStr = sqlStr + " where code_large<'999'"
	sqlStr = sqlStr + " and display_yn='Y'"
	sqlStr = sqlStr + " order by orderNO, code_large"

	rsget.Open sqlStr, dbget, 1
	do until rsget.Eof
		response.write "<item>" + VbCrlf
		response.write "<value1>" + rsget("code_large") + "</value1>" + VbCrlf
		response.write "<value2><![CDATA[" + db2html(rsget("code_nm")) + "]]></value2>" + VbCrlf
		response.write "</item>" + VbCrlf
		rsget.moveNext
	loop
	rsget.close

elseif mode="cdm" then
	sqlStr = "select top 100 * from [db_item].[dbo].tbl_Cate_mid"
	sqlStr = sqlStr + " where code_large='" + param1 + "'"
	sqlStr = sqlStr + " and display_yn='Y'"
	sqlStr = sqlStr + " order by orderNO, code_mid"

	rsget.Open sqlStr, dbget, 1
	do until rsget.Eof
		response.write "<item>" + VbCrlf
		response.write "<value1>" + rsget("code_mid") + "</value1>" + VbCrlf
		response.write "<value2><![CDATA[" + db2html(rsget("code_nm")) + "]]></value2>" + VbCrlf
		response.write "</item>" + VbCrlf

		rsget.moveNext
	loop
	rsget.close

elseif mode="cds" then
	sqlStr = "select top 100 * from [db_item].[dbo].tbl_Cate_small"
	sqlStr = sqlStr + " where code_large='" + param1 + "'"
	sqlStr = sqlStr + " and code_mid='" + param2 + "'"
	sqlStr = sqlStr + " and display_yn='Y'"
	sqlStr = sqlStr + " order by orderNO, code_small"

	rsget.Open sqlStr, dbget, 1
	do until rsget.Eof
		response.write "<item>" + VbCrlf
		response.write "<value1>" + rsget("code_small") + "</value1>" + VbCrlf
		response.write "<value2><![CDATA[" + db2html(rsget("code_nm")) + "]]></value2>" + VbCrlf
		response.write "</item>" + VbCrlf

		rsget.moveNext
	loop
	rsget.close

elseif mode="cdselect" then


end if
%>
</response>
<!-- #include virtual="/lib/db/dbclose.asp" -->