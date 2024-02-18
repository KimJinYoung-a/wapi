<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
''UTF-8 로 해야 한글이 안깨지게 받음.

dim idx     : Idx=requestCheckVar(request("idx"),10)
dim itemid  : itemid=requestCheckVar(request("itemid"),10)
dim ErrCode : ErrCode=requestCheckVar(request("ErrCode"),32)
dim ErrMsg  : ErrMsg=requestCheckVar(request("ErrMsg"),500)

dim sqlStr
sqlStr = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&ErrCode&"','"&html2DB(ErrMsg)&"'"
dbget.Execute sqlStr

response.write "S_OK"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->