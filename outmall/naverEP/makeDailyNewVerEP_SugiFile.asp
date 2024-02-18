<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 1200  ''초단위
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim IsChangedEP : IsChangedEP = (request("epType")="chg")
Dim sqlStr
''작성시간 체크
If (IsChangedEP) Then
    sqlStr = "exec [db_outmall].[dbo].[sp_Ten_Naver_EPDataChgInsert] "
    dbCTget.CommandTimeout = 120 ''2021/04/14 추가
    dbCTget.execute sqlStr
	response.write "요약 EP 모수 생성 완료"
Else
    sqlStr = "exec [db_outmall].[dbo].[sp_Ten_Naver_EPDataInsert] "
    dbCTget.CommandTimeout = 300 ''2021/04/14 추가
    dbCTget.execute sqlStr
	response.write "전체 EP 모수 생성 완료"
End If
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->