<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 1200  ''�ʴ���
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim IsChangedEP : IsChangedEP = (request("epType")="chg")
Dim sqlStr
''�ۼ��ð� üũ
If (IsChangedEP) Then
    sqlStr = "exec [db_outmall].[dbo].[sp_Ten_Naver_EPDataChgInsert] "
    dbCTget.CommandTimeout = 120 ''2021/04/14 �߰�
    dbCTget.execute sqlStr
	response.write "��� EP ��� ���� �Ϸ�"
Else
    sqlStr = "exec [db_outmall].[dbo].[sp_Ten_Naver_EPDataInsert] "
    dbCTget.CommandTimeout = 300 ''2021/04/14 �߰�
    dbCTget.execute sqlStr
	response.write "��ü EP ��� ���� �Ϸ�"
End If
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->