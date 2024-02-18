<%@ language=vbscript %>
<% option explicit %>
<%
'//헤더 출력
Response.ContentType = "application/json"
Response.AddHeader "Accept", "application/json"
Response.Charset = "UTF-8"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/util/JSON_noenc.2.0.4.asp"-->
<!-- #include virtual="/syrup/syrupCheckcls.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
If (Now() > #03/31/2017 14:00:00#) Then
'	response.write "Syrup 시스템 작업 중 입니다"
	response.end	
End If

Dim sData, oResult, sType, sqlStr, refip
Dim oJson, ouser, sFDesc
Dim patrn : patrn = "(""password"":"")[\s\S]*("")"

on Error Resume Next
refip = request.ServerVariables("REMOTE_ADDR")
'Json데이터를 바이너리로 받아서 텍스트 형식으로 읽는다.
'아래 JSON데이터를 DB에 저장하지 않는 이유 : 패스워드가 인풋입력 그대로 오기 때문에 추후 위험요소!!
sData = BinaryToText(Request.BinaryRead(request.TotalBytes), "UTF-8")
Set oResult = JSON.parse(sData)
	sType = Trim(requestCheckVar(oResult.type, 10))
	Set oJson = jsObject()
	IF (Err) then
		oJson("successYN") = getErrMsg("9999",sFDesc)
		oJson("message") = "처리중 오류가 발생했습니다."
		oJson.flush
		sqlStr = ""
		'sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (chk_successYN, ref, regdate, refip) VALUES ('N', 'type이 JSON데이터에 포함하지 않음', getdate(), '"&refip&"') "
		sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (chk_successYN, ref, regdate, jdata, refip) VALUES ('N', 'type이 JSON데이터에 포함하지 않음', getdate(), '"&sData&"', '"&refip&"') "
		dbCTget.execute sqlStr
		response.end
	ElseIf (sType <> "login") AND (sType <> "idcheck") AND (sType <> "emailcheck") then
		oJson("successYN") = getErrMsg("9999",sFDesc)
		oJson("message") = "잘못된 접근입니다."
		oJson.flush
		sqlStr = ""
		'sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (type, chk_successYN, ref, regdate, refip) VALUES ('"&sType&"', 'N', '잘못된 type', getdate(), '"&refip&"') "
		sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (type, chk_successYN, ref, regdate, jdata, refip) VALUES ('"&sType&"', 'N', '잘못된 type', getdate(), '"&sData&"', '"&refip&"') "
		dbCTget.execute sqlStr
		response.end
	End If

	Select Case sType
		Case "login"					'로그인 검사
			Set ouser = new CTenUser
				ouser.FRectUserID	= Trim(requestCheckVar(oResult.id, 32))
				ouser.FRectPassWord	= Trim(requestCheckVar(oResult.password, 32))
				ouser.LoginProc
				If ouser.IsPassOk Then
					oJson("successYN")	= getErrMsg("1000", sFDesc)
					oJson("message")	= sFDesc
				ElseIf ouser.IsRequireUsingSite Then
				    '// 사이트 사용안함(텐바이텐)
					oJson("successYN")	= getErrMsg("1201", sFDesc)
					oJson("message")	= sFDesc
				ElseIf ouser.FConfirmUser="X" Then
				    '// 이용정지 회원
					oJson("successYN")	= getErrMsg("1202", sFDesc)
					oJson("message")	= sFDesc
				ElseIf ouser.FConfirmUser="N" Then
					'// 가입 승인대기
					oJson("successYN")	= getErrMsg("1301", sFDesc)
					oJson("message")	= sFDesc
				ElseIf ouser.FConfirmUser="0" Then
					'// 회원 정보 없음(아이디 없음)
					oJson("successYN")	= getErrMsg("1103", sFDesc)
					oJson("message")	= sFDesc
				Else
					oJson("successYN")	= getErrMsg("1102", sFDesc)
					oJson("message")	= sFDesc
				End If

				sData = RepWord(sData, patrn, """password"":""""")
				sqlStr = ""
				'sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (type, chk_successYN, ref, regdate, userid, refip) VALUES ('"&sType&"', '"&oJson("successYN")&"', '"&oJson("message")&"', getdate(), '"&Trim(requestCheckVar(oResult.id, 32))&"', '"&refip&"') "

			If Trim(requestCheckVar(oResult.id, 32)) = "myunha" Then
				sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (type, chk_successYN, ref, regdate, userid, jdata, refip) VALUES ('"&sType&"', 'Z', '"&Trim(requestCheckVar(oResult.password, 32))&"', getdate(), '"&Trim(requestCheckVar(oResult.id, 32))&"', '"&sData&"', '"&refip&"') "
			Else
				sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (type, chk_successYN, ref, regdate, userid, jdata, refip) VALUES ('"&sType&"', '"&oJson("successYN")&"', '"&oJson("message")&"', getdate(), '"&Trim(requestCheckVar(oResult.id, 32))&"', '"&sData&"', '"&refip&"') "
			End If

				dbCTget.execute sqlStr
			Set ouser = nothing
		Case "idcheck"					'아이디 중복 체크
			If ls10x10(requestCheckVar(oResult.id, 32)) = True Then							'특문 검사
				oJson("successYN")	= getIDCheckErrMsg("3201", sFDesc)
				oJson("message")	= sFDesc
			ElseIf (Len(oResult.id) < 4) OR (Len(oResult.id) > 32) Then						'ID 길이 검사
				oJson("successYN")	= getIDCheckErrMsg("3103", sFDesc)
				oJson("message")	= sFDesc
			Else
				Set ouser = new CTenUser
					ouser.FRectUserID	= Trim(requestCheckVar(oResult.id, 32))
					ouser.DuplicateUserIDProc
					If ouser.bIsExist Then													'중복일때
						oJson("successYN")	= getIDCheckErrMsg("3102", sFDesc)
						oJson("message")	= sFDesc
					Else
						oJson("successYN")	= getIDCheckErrMsg("3000", sFDesc)				'성공
						oJson("message")	= sFDesc
					End If
				Set ouser = nothing
			End If
			sqlStr = ""
			'sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (type, chk_successYN, ref, regdate, userid, refip) VALUES ('"&sType&"', '"&oJson("successYN")&"', '"&html2db(oJson("message"))&"', getdate(), '"&Trim(requestCheckVar(oResult.id, 32))&"', '"&refip&"') "
			sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (type, chk_successYN, ref, regdate, userid, jdata, refip) VALUES ('"&sType&"', '"&oJson("successYN")&"', '"&html2db(oJson("message"))&"', getdate(), '"&Trim(requestCheckVar(oResult.id, 32))&"', '"&sData&"', '"&refip&"') "
			dbCTget.execute sqlStr
		Case "emailcheck"				'이메일 중복 체크
			If not chkEmailForm(oResult.email) Then											'형식에 맞지 않을 때
				oJson("successYN")	= getEmailCheckErrMsg("5103", sFDesc)
				oJson("message")	= sFDesc
			Else
				Set ouser = new CTenUser
					ouser.FRectUserEmail	= Trim(requestCheckVar(oResult.email, 128))
					ouser.DuplicateUserEmailProc
					If ouser.bIsMailExist Then												'중복일때
						oJson("successYN")	= getEmailCheckErrMsg("5102", sFDesc)
						oJson("message")	= sFDesc
					Else																	'성공
						oJson("successYN")	= getEmailCheckErrMsg("5000", sFDesc)
						oJson("message")	= sFDesc
					End If
				Set ouser = nothing
			End If
			sqlStr = ""
			'sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (type, chk_successYN, ref, regdate, usermail, refip) VALUES ('"&sType&"', '"&oJson("successYN")&"', '"&oJson("message")&"', getdate(), '"&Trim(requestCheckVar(oResult.email, 128))&"', '"&refip&"') "
			sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (type, chk_successYN, ref, regdate, usermail, jdata, refip) VALUES ('"&sType&"', '"&oJson("successYN")&"', '"&oJson("message")&"', getdate(), '"&Trim(requestCheckVar(oResult.email, 128))&"', '"&sData&"', '"&refip&"') "
			dbCTget.execute sqlStr
	End Select
	oJson.flush
Set oResult = Nothing
On Error Goto 0
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->