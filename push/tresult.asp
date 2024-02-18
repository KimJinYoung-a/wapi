<%@ language=vbscript %>
<% option Explicit %>
<%
'#######################################################
'	Description	: 푸시 결과 클릭 업데이트
'	History : 서동석 생성
'             2019.06.19 한용민 수정
'#######################################################
%>
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
''푸시 결과클릭  업데이트
Dim ikey : ikey=request("ikey")
Dim pkey : pkey=request("pkey")
Dim pid : pid=request("pid")
Dim targetkey : targetkey=request("targetkey")
Dim refIP : refIP=request.ServerVariables("REMOTE_ADDR")

Dim sqlStr
sqlStr = "exec db_AppNoti.dbo.sp_ten_AppPushMsg_Click_SAVE_WithPKey '"&ikey&"','"&pkey&"','"&pid&"','"&refIP&"','"& targetkey &"'"
dbAppNotiget.execute sqlStr

'response.write "pkey:"&pkey
'response.write "<BR>"
'response.write "ikey:"&ikey
'response.write "<BR>"
'response.write "pid:"&pid
%>
<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->