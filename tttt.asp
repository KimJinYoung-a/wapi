<%=request.servervariables("remote_addr")%><br>
<%= now() %><br>
1111112222

<%
Dim oJsonSentry, sCurrUrl, sentryClientId, sentrySendBody, sentrySendServer, sentryErrorTypeMsg, sentryBlnErrorWritten
Dim sentryStrServerName, sentryStrServerIP, sentryStrRemoteIP
Dim sentryErrorTagsFile, sentryErrorTagsline, sentryMethod, sentryMethodData
Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
Dim strMethod, lngPos, datNow, strQueryString, strURL

Set objASPError = Server.GetLastError
Dim bakCodepage, strMsg
strMsg = "<li>오류 유형:<br>"
bakCodepage = Session.Codepage

select case Trim(application("Svr_Info"))
		case "Dev"
			sentrySendServer = "http://aspsentrydev.10x10.co.kr/api/Sentry/CaptureError"
		case else
			sentrySendServer = "http://aspsentry.10x10.co.kr/api/Sentry/CaptureError"
	End Select

	'// pc는 clientid가 하나
	sCurrUrl = Request.ServerVariables("SCRIPT_NAME")
	sCurrUrl = Lcase(sCurrUrl)
	sentryClientId = "10x10-asp-pc"

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
		sentryMethodData = Request.TotalBytes & " bytes to "&Request.Form
	ElseIf sentryMethod = "GET" Then
		sentryMethodData = Request.QueryString
	End If

	Dim message2
	message2 = Replace("Microsoft VBScript 런타임 오류 (0x800A000D) 형식이 일치하지 않습니다.: '[string: ""]'", """", "")

    sentrySendBody = ""
	sentrySendBody = sentrySendBody & " { "
	sentrySendBody = sentrySendBody & " 	""clientName"" : """&sentryClientId&""","
	sentrySendBody = sentrySendBody & " 	""message"" : """&sentryErrorTypeMsg&""","
	sentrySendBody = sentrySendBody & " 	""tags"" : { "
	sentrySendBody = sentrySendBody & " 		""file"" : """&sentryErrorTagsFile&""","
	sentrySendBody = sentrySendBody & " 		""line"" : """&sentryErrorTagsline&""","
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
	sentrySendBody = sentrySendBody & " 		""name"" : ""system ttt"","
	sentrySendBody = sentrySendBody & " 		""name2"" : """&message2&""","
	sentrySendBody = sentrySendBody & " 		""ip"" : """&sentryStrRemoteIP&""""
	sentrySendBody = sentrySendBody & "     }, "
    sentrySendBody = sentrySendBody & "     ""tmeta"" : { "
    sentrySendBody = sentrySendBody & "         ""service_name"" : ""restapi"""
	sentrySendBody = sentrySendBody & " 	} "
	sentrySendBody = sentrySendBody & " } "

	' set oJsonSentry = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	' oJsonSentry.open "POST", sentrySendServer, False
	' oJsonSentry.setRequestHeader "Content-Type", "application/json; charset=utf-8"
	' oJsonSentry.setRequestHeader "key","lkzxljk-fqwo@i3J875qlkzLjdv"
	' oJsonSentry.setRequestHeader "CharSet", "utf-8" '있어도 되고 없어도 되고
	' oJsonSentry.setRequestHeader "Accept","application/json"
	' oJsonSentry.send sentrySendBody

	set oJsonSentry2 = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oJsonSentry2.open "POST", "http://172.16.0.218/", False
	oJsonSentry2.setRequestHeader "Content-Type", "application/json; charset=utf-8"
	oJsonSentry2.setRequestHeader "CharSet", "utf-8" '있어도 되고 없어도 되고
	oJsonSentry2.setRequestHeader "Accept","application/json"
	oJsonSentry2.setRequestHeader "api-key-v1","bd05f7a763aa2978aeea5e8f2a8a3242abc0cbffeb3c28e0b056cef4e282eee9"
	oJsonSentry2.setRequestHeader "host_lo", "logoneapi.10x10.co.kr" 
	oJsonSentry2.send sentrySendBody


    response.write oJsonSentry2.responseText
    ' response.write1 oJsonSentry2.responseText

    Set oJsonSentry = Nothing

%>