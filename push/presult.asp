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
Dim psKey : psKey=request("ikey")
Dim deviceid : deviceid=request("pid")
Dim targetkey : targetkey=request("targetkey")
Dim refIP : refIP=request.ServerVariables("REMOTE_ADDR")

Dim sqlStr
sqlStr = "exec db_AppNoti.dbo.sp_ten_AppPushMsg_Click_SAVE "&psKey&",'"&deviceid&"','"&refIP&"','"& targetkey &"'"
dbAppNotiget.execute sqlStr

'response.write psKey
'response.write "<BR>"
'response.write deviceid
%>
<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->