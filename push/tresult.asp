<%@ language=vbscript %>
<% option Explicit %>
<%
'#######################################################
'	Description	: Ǫ�� ��� Ŭ�� ������Ʈ
'	History : ������ ����
'             2019.06.19 �ѿ�� ����
'#######################################################
%>
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
''Ǫ�� ���Ŭ��  ������Ʈ
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