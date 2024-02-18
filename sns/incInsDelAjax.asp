<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<%
'####################################################
' Description :  sns 데이터 추가삭제
' History : 2015.07.21 원승현 생성
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
dim eCode
IF application("Svr_Info") = "Dev" THEN
	eCode   =  66295		''[소품전] 반짝반짝 스티커
Else
	eCode   =  77064		''[소품전] 반짝반짝 스티커
End If

dim userid, mode, sqlstr, refer, vSnsId, vChk, instaSns, vLink, vImg_low, vImg_thum, vImg_stand, vText
dim vSnsUserid, vSnsUserName, vQuery, i, vJsonSnsUrl

refer = request.ServerVariables("HTTP_REFERER")
'userid = GetLoginUserID
vSnsId = request("snsid")
vJsonSnsUrl = request("JsonSnsUrl")

'// 바로 접속시엔 오류 표시
if InStr(refer,"10x10.co.kr")<1 then
'	Response.Write "Err|잘못된 접속입니다."
'	Response.End
end If


'// 데이터 내역 검색
sqlstr = "select count(*) "
sqlstr = sqlstr & " from [db_Appwish].[dbo].[tbl_snsSelectData]"
sqlstr = sqlstr & " where snsid='"& vSnsId &"' and evt_code="& eCode &""

'response.write sqlstr & "<br>"
rsCTget.Open sqlstr, dbCTget, adOpenForwardOnly, adLockReadOnly
If rsCTget(0) >= 1 Then
	'// 값이 있을경우엔 Delete
	vChk = "Delete"
Else
	'// 값이 없을경우엔 insert
	vChk = "insert"
End IF
rsCTget.close

If vChk = "insert" Then
	Set instaSns = JSON.parse(getJsonAsp(""&vJsonSnsUrl&"",""))

	If IsNull(instaSns) Or instaSns<>"" Then
		For i=0 To 30
			If instaSns.data.Get(i).caption.id = vSnsId Then
				vLink = instaSns.data.Get(i).link
				vImg_low = instaSns.data.Get(i).images.low_resolution.url
				vImg_thum = instaSns.data.Get(i).images.thumbnail.url
				vImg_stand = instaSns.data.Get(i).images.standard_resolution.url
				vText = nl2br(html2db(instaSns.data.Get(i).caption.text))
				vSnsUserid = instaSns.data.Get(i).user.id
				vSnsUserName = instaSns.data.Get(i).user.username
				Exit For
			End If
		Next
	End If
	Set instaSns = Nothing

	vQuery = "INSERT INTO [db_Appwish].[dbo].[tbl_snsSelectData] (snsid, link, img_low, img_thum, img_stand, text, snsuserid, snsusername, regdate, evt_code)" + vbcrlf
	vQuery = vQuery & " VALUES('"& vSnsId &"', '"& vLink &"', '"& vImg_low &"', '"& vImg_thum &"', '"& vImg_stand &"'" + vbcrlf
	vQuery = vQuery & " , '"& vText &"', '"& vSnsUserid &"', '"& vSnsUserName &"', getdate(), "& eCode &")"

	'response.write vQuery & "<br>"
	dbCTget.execute vQuery
	Response.Write "OK|1|"&vSnsId
	Response.End
Else
	vQuery = "Delete From [db_Appwish].[dbo].[tbl_snsSelectData] Where snsid='"&vSnsId&"' and evt_code="& eCode &""
	
	dbCTget.execute vQuery
	Response.Write "OK|2|"&vSnsId
	Response.End
End If

'// asp json 거시기
Function getJsonAsp(url, param)
	Dim objHttp
	Dim strJsonText
	Set objHttp = server.CreateObject("Microsoft.XMLHTTP")
	If IsNull(objHttp) Then
		response.write "서버 연결 오류"
		response.End
	End If
	objHttp.Open "Get", url, False
	objHttp.SetRequestHeader "Content-Type","text/plain"
	objHttp.Send param
	strJsonText = objHttp.responseText
	Set objHttp = Nothing

	getJsonAsp = strJsonText

End Function

Set instaSns = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->