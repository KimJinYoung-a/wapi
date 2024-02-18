<%@ language="VBScript" %>
<%


  Option Explicit

  Const lngMaxFormBytes = 800

  Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
  Dim strMethod, lngPos, datNow, strQueryString, strURL

	If Response.Buffer Then
		Response.Clear
		Response.Status = "500 Internal Server Error"
		Response.ContentType = "text/html"
		''Response.ChaRset = "euc-kr"
		Response.Expires = 0
	End If

	'// 오류객체 선언
	Set objASPError = Server.GetLastError

	'### 오류 메시지 작성 ###
	Dim bakCodepage, strMsg

	'// 오류 유형 정보
	strMsg = "<li>오류 유형:<br>"

	on error resume next
		bakCodepage = Session.Codepage
		Session.Codepage = 1252
		on error goto 0

		strMsg = strMsg & Server.HTMLEncode(objASPError.Category)

		If objASPError.ASPCode > "" Then strMsg = strMsg & Server.HTMLEncode(", " & objASPError.ASPCode)
			strMsg = strMsg &  Server.HTMLEncode(" (0x" & Hex(objASPError.Number) & ")" ) & "<br>"
		If objASPError.ASPDescription > "" Then
			strMsg = strMsg & Server.HTMLEncode(objASPError.ASPDescription) & "<br>"
		elseIf (objASPError.Description > "") Then
			strMsg = strMsg & Server.HTMLEncode(objASPError.Description) & "<br>"
		end if

		blnErrorWritten = False

		'IIS에서 발생한 오류 코드를 출력합니다.
		If objASPError.Source > "" Then
			strServername = LCase(Request.ServerVariables("SERVER_NAME"))
			strServerIP = Request.ServerVariables("LOCAL_ADDR")
			strRemoteIP =  Request.ServerVariables("REMOTE_ADDR")

			If (strServerIP = strRemoteIP) And objASPError.File <> "?" Then
				strMsg = strMsg & Server.HTMLEncode(objASPError.File)

				If objASPError.Line > 0 Then strMsg = strMsg & ", line " & objASPError.Line
				If objASPError.Column > 0 Then strMsg = strMsg & ", column " & objASPError.Column
				strMsg = strMsg & "<br>"
				strMsg = strMsg & "<font style=""COLOR:000000; FONT: 8pt/11pt courier new""><b>"
				strMsg = strMsg & Server.HTMLEncode(objASPError.Source) & "<br>"
				If objASPError.Column > 0 Then strMsg = strMsg & String((objASPError.Column - 1), "-") & "^<br>"
				strMsg = strMsg & "</b></font>"
				blnErrorWritten = True
			End If
		End If

		If Not blnErrorWritten And objASPError.File <> "?" Then
			strMsg = strMsg & "<b>" & Server.HTMLEncode(  objASPError.File)
			If objASPError.Line > 0 Then strMsg = strMsg & Server.HTMLEncode(", line " & objASPError.Line)
			If objASPError.Column > 0 Then strMsg = strMsg & ", column " & objASPError.Column
			strMsg = strMsg & "</b><br>"
		End If
		strMsg = strMsg & "</li>"

	'// 접속자 브라우저 정보
	strMsg = strMsg & "<li>브라우저 종류:<br>"
	strMsg = strMsg & Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT"))
	strMsg = strMsg & "<br><br></li>"

    strMsg = strMsg & "<li>접속자 IP:<br>"
	strMsg = strMsg & Server.HTMLEncode(Request.ServerVariables("REMOTE_ADDR"))
	strMsg = strMsg & "<br><br></li>"

	strMsg = strMsg & "<li>경유페이지:<br>"
	strMsg = strMsg & request.ServerVariables("HTTP_REFERER")
	strMsg = strMsg & "<br><br></li>"

	'// 오류 페이지 정보
	strMsg = strMsg & "<li>페이지:<br>"
	strMethod = Request.ServerVariables("REQUEST_METHOD")
	strMsg = strMsg & "HOST : " & Request.ServerVariables("HTTP_HOST") & "<BR>"
	strMsg = strMsg & strMethod & " : "

	If strMethod = "POST" Then
		strMsg = strMsg & Request.TotalBytes & " bytes to "
	End If

	strMsg = strMsg & Request.ServerVariables("SCRIPT_NAME")
	strMsg = strMsg & "</li>"

	If strMethod = "POST" Then
		strMsg = strMsg & "<br><li>POST Data:<br>"

		'실행에 관련된 에러를 출력합니다.
		On Error Resume Next
		If Request.TotalBytes > lngMaxFormBytes Then
			strMsg = strMsg & Server.HTMLEncode(Left(Request.Form, lngMaxFormBytes)) & " . . ."'
		Else
			strMsg = strMsg & Server.HTMLEncode(Request.Form)
		End If
		On Error Goto 0
		strMsg = strMsg & "</li>"
	elseif strMethod = "GET" then
		strMsg = strMsg & "<br><li>GET Data:<br>"
		strMsg = strMsg & Request.QueryString
	End If
	strMsg = strMsg & "<br><br></li>"

	'// 오류 발생시간 정보
	strMsg = strMsg & "<li>시간:<br>"
	datNow = Now()
	strMsg = strMsg & Server.HTMLEncode(FormatDateTime(datNow, 1) & ", " & FormatDateTime(datNow, 3))

	'// 사용자 정보
	strMsg = strMsg & "<br><br><li>접속User:<br>"
	strMsg = strMsg & session("ssBctID")

	on error resume next
		Session.Codepage = bakCodepage
	on error goto 0
	strMsg = strMsg & "<br><br></li>"

	'### 시스템팀 구성원에게 오류 발생 내용 발송 ###
	dim cdoMessage,cdoConfig

	Set cdoConfig = CreateObject("CDO.Configuration")

	'-> 서버 접근방법을 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)

	'-> 서버 주소를 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")= "110.93.128.95"

	'-> 접근할 포트번호를 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

	'-> 접속시도할 제한시간을 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 15

	'-> SMTP 접속 인증방법을 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

	'-> SMTP 서버에 인증할 ID를 입력합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"

	'-> SMTP 서버에 인증할 암호를 입력합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"

	cdoConfig.Fields.Update

	Set cdoMessage = CreateObject("CDO.Message")
	Set cdoMessage.Configuration = cdoConfig

'	cdoMessage.To 		= "kobula@10x10.co.kr;tozzinet@10x10.co.kr;kjy8517@10x10.co.kr;errmail@10x10.co.kr;thensi7@10x10.co.kr;corpse2@10x10.co.kr;"
	cdoMessage.To 		= "errmail@10x10.co.kr"
	cdoMessage.From 	= "webserver@10x10.co.kr"
	cdoMessage.SubJect 	= "["&date()&"] WAPI 페이지 오류 발생"
	cdoMessage.HTMLBody	= strMsg
	cdoMessage.Send

	on error resume next
	'// Sentry로 에러 전송
	Dim oJsonSentry, sCurrUrl, appPageCheck, sentryClientId, sentrySendBody, sentrySendServer, sentryErrorTypeMsg, sentryBlnErrorWritten
	Dim sentryStrServerName, sentryStrServerIP, sentryStrRemoteIP
	Dim sentryErrorTagsFile, sentryErrorTagsline, sentryMethod, sentryMethodData

	select case Trim(application("Svr_Info"))
		case "Dev"
			sentrySendServer = "http://aspsentrydev.10x10.co.kr/api/Sentry/CaptureError"
		case else
			sentrySendServer = "http://aspsentry.10x10.co.kr/api/Sentry/CaptureError"
	End Select

	'// 앱과 모바일을 구분한다.
	sCurrUrl = Request.ServerVariables("SCRIPT_NAME")
	sCurrUrl = Lcase(sCurrUrl)	

	sentryClientId = "10x10-asp-wapi"

	'// Sentry로 보낼 오류 메시지
	sentryErrorTypeMsg = "" '// 오류 메시지
	sentryErrorTagsFile = "" '// 오류 파일
	sentryErrorTagsline = "" '// 오류 라인, 컬럼
	sentryMethodData = "" '// Method별 Data
	sentryStrServerName = LCase(Request.ServerVariables("SERVER_NAME")) '// 서버명
	sentryStrServerIP = Request.ServerVariables("LOCAL_ADDR") '// 서버 ip
	sentryStrRemoteIP =  Request.ServerVariables("REMOTE_ADDR") '// 접속자 ip

	'// 메시지 정의
	sentryErrorTypeMsg = sentryErrorTypeMsg & objASPError.Category
	If objASPError.ASPCode > "" Then sentryErrorTypeMsg = sentryErrorTypeMsg & ", " & objASPError.ASPCode
		sentryErrorTypeMsg = sentryErrorTypeMsg &  " (0x" & Hex(objASPError.Number) & ")"
	If objASPError.ASPDescription > "" Then 
		sentryErrorTypeMsg = sentryErrorTypeMsg & objASPError.ASPDescription
	elseIf (objASPError.Description > "") Then 
		sentryErrorTypeMsg = sentryErrorTypeMsg & objASPError.Description
	end if

	sentryBlnErrorWritten = False
	'파일, line, column
	If objASPError.Source > "" Then
		If (sentryStrServerIP = sentryStrRemoteIP) And objASPError.File <> "?" Then
			sentryErrorTagsFile = sentryErrorTagsFile & objASPError.File

			If objASPError.Line > 0 Then sentryErrorTagsline = sentryErrorTagsline & "line " & objASPError.Line
			If objASPError.Column > 0 Then sentryErrorTagsline = sentryErrorTagsline & ", column " & objASPError.Column
			sentryErrorTagsline = sentryErrorTagsline & objASPError.Source
			If objASPError.Column > 0 Then sentryErrorTagsline = sentryErrorTagsline & String((objASPError.Column - 1), "-")
			sentryBlnErrorWritten = True
		End If
	End If

	If Not sentryBlnErrorWritten And objASPError.File <> "?" Then
		sentryErrorTagsFile = sentryErrorTagsFile &   objASPError.File
		If objASPError.Line > 0 Then sentryErrorTagsline = sentryErrorTagsline & ", line " & objASPError.Line
		If objASPError.Column > 0 Then sentryErrorTagsline = sentryErrorTagsline & ", column " & objASPError.Column
	End If

	'// method 구분
	sentryMethod = Request.ServerVariables("REQUEST_METHOD")

	If sentryMethod = "POST" Then
		'실행에 관련된 에러를 출력합니다.
		On Error Resume Next

		sentryMethodData = Request.TotalBytes & " bytes to "

		If Request.TotalBytes > lngMaxFormBytes Then
			sentryMethodData = sentryMethodData & Server.HTMLEncode(Left(Request.Form, lngMaxFormBytes)) & " . . ."
		Else
			sentryMethodData = sentryMethodData & Server.HTMLEncode(Request.Form)
		End If
		On Error Goto 0
		sentryMethodData = Request.TotalBytes & " bytes to "&Request.Form
	ElseIf sentryMethod = "GET" Then
		sentryMethodData = Request.QueryString
	End If

	sentryErrorTypeMsg = Replace(sentryErrorTypeMsg, """", "")

	sentrySendBody = ""
	sentrySendBody = sentrySendBody & " { "
	sentrySendBody = sentrySendBody & " 	""clientName"" : """&sentryClientId&""","
	sentrySendBody = sentrySendBody & " 	""message"" : """&sentryErrorTypeMsg&""","
	sentrySendBody = sentrySendBody & " 	""tags"" : { "
	sentrySendBody = sentrySendBody & " 		""file"" : """&sentryErrorTagsFile&""","
	' sentrySendBody = sentrySendBody & " 		""line"" : """&sentryErrorTagsline&""","
	sentrySendBody = sentrySendBody & " 		""line"" : """&objASPError.Line&""","
	sentrySendBody = sentrySendBody & " 		""remoteIp"" : """&sentryStrRemoteIP&""","
	sentrySendBody = sentrySendBody & " 		""server"" : """&application("Svr_Info")&""""
	sentrySendBody = sentrySendBody & " 	}, "
	sentrySendBody = sentrySendBody & " 	""headers"" : { "
	sentrySendBody = sentrySendBody & " 		""user-agent"" : """&Request.ServerVariables("HTTP_USER_AGENT")&""","
	sentrySendBody = sentrySendBody & " 		""referer"" : """&request.ServerVariables("HTTP_REFERER")&""","
	sentrySendBody = sentrySendBody & " 		""host"" : """&Request.ServerVariables("HTTP_HOST")&""""
	sentrySendBody = sentrySendBody & " 	}, "
	sentrySendBody = sentrySendBody & " 	""request"" : { "
	sentrySendBody = sentrySendBody & " 		""url"" : """&Request.ServerVariables("SCRIPT_NAME")&""","
	sentrySendBody = sentrySendBody & " 		""method"" : """&sentryMethod&""","
	sentrySendBody = sentrySendBody & " 		""data"" : """&sentryMethodData&""""
	sentrySendBody = sentrySendBody & " 	}, "
	sentrySendBody = sentrySendBody & " 	""user"" : { "
	sentrySendBody = sentrySendBody & " 		""name"" : ""system"","
	sentrySendBody = sentrySendBody & " 		""ip"" : """&sentryStrRemoteIP&""""
	sentrySendBody = sentrySendBody & "     }, "
    sentrySendBody = sentrySendBody & "     ""tmeta"" : { "
    sentrySendBody = sentrySendBody & "         ""service_name"" : ""asperror"""
	sentrySendBody = sentrySendBody & " 	} "
	sentrySendBody = sentrySendBody & " } "

	' logone 로그 by JaeSeok
	set oJsonSentry = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	' oJsonSentry.open "POST", sentrySendServer, False
	oJsonSentry.open "POST", "http://172.16.0.218/", False
	oJsonSentry.setRequestHeader "Content-Type", "application/json; charset=utf-8"
	oJsonSentry.setRequestHeader "key","lkzxljk-fqwo@i3J875qlkzLjdv"
	oJsonSentry.setRequestHeader "CharSet", "utf-8" '있어도 되고 없어도 되고
	oJsonSentry.setRequestHeader "Accept","application/json"
	oJsonSentry.setRequestHeader "api-key-v1","bd05f7a763aa2978aeea5e8f2a8a3242abc0cbffeb3c28e0b056cef4e282eee9"
	oJsonSentry.setRequestHeader "host_lo", "logoneapi.10x10.co.kr" 
	oJsonSentry.send sentrySendBody

	'If InStr(oJsonSentry.responseText, "success") > 0 Then
		'response.write oJsonSentry.responseText
	'End If

	on error goto 0

	Set cdoMessage = nothing
	Set cdoConfig = nothing
	Set oJsonSentry = Nothing

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<title>텐바이텐 10X10 = 감성채널 감성에너지..</title>
</head>
<body>
<table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
	<td><img src="http://fiximage.10x10.co.kr/web2008/cscenter/sorry_top2.gif" width="500" height="278" /></td>
</tr>
<tr>
	<td height="66" bgcolor="#f7f7f7">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td align="right"><a href="javascript:history.back()" onFocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2008/cscenter/btn_back.gif" width="94" height="35" border="0" /></a></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</body>
</html>