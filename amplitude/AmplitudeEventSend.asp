<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<% Server.ScriptTimeOut = 60*60 %>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbDatamartopen.asp" -->

<!-- #include virtual="/lib/function.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'#######################################################
'	History	: 2018.08.06 원승현 생성
'	Description : Amplitude에 이벤트 전송
'#######################################################
	Dim mode, userId, oXML, sendJsonData, query1, appBoySendEnable, myJson, query2, platForm, itemids, limitUseq, totalcnt, i
	Dim ampurl, platformApiKey

	mode = request("mode")
	userId = request("userid")
	platForm = request("platForm")
	itemids = request("itemids")

	appBoySendEnable = False


	If Trim(mode)="" Then
		response.End
	End If

	Select Case Trim(platform)
		Case "mobile"
			IF (application("Svr_Info")	= "Dev") Then
				'// Test용
				platformApiKey = "fd52aea3710592a696baae5447a6630d"
			Else
				'// 실서비스용
				platformApiKey = "3e5d96e41fc92b60c3a28f9fb4ae7620"
			End If
		Case "pcweb"
			IF (application("Svr_Info")	= "Dev") Then		
				'// Test용
				platformApiKey = "31e6741da66c20e94f5807bb844e129f"
			Else
				'// 실서비스용
				platformApiKey = "91316725130cbc5b997cef756ce9388a"
			End If
		Case "app"
			IF (application("Svr_Info")	= "Dev") Then				
				'// Test용
				platformApiKey = "accf99428106843efdd88df080edd82e"
			Else
				'// 실서비스용
				platformApiKey = "3de77f281d7d09a7903c1d1fa2e4fa2d"
			End If
		Case Else
			platformApiKey = ""
	End Select

	'// api_key가 없으면 아예 발송을 하지 않는다.
	If Trim(platformApiKey)="" Then
		response.write "필요한 APIKEY가 제공되지 않았습니다."
		response.End
	End If

	Select Case Trim(mode)

		'// 카카오 알림톡으로 상품구매 2일 후 전송되는 유저들을 amplitude에 보내준다.
		Case "amplitudeUserPropertiesSend"
			query1 = ""
			query1 = query1 + "	Select Idx, FixedDate, UserSeq*3 AS UserSeq, UserId, UserCell, OrderSerial, ItemName, ItemCount, IsSend, RegDate From db_log.dbo.tbl_SendReviewUsers WITH (NOLOCK) "
			query1 = query1 + "	Where UserSeq IS NOT NULL And FixedDate = CONVERT(VARCHAR(10), GETDATE(), 120)  "
			rsget.Open query1,dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.bof Or rsget.eof) Then
				Do Until rsget.eof

					'// IsSend값을 기준으로 1은 발송 0은 미발송
					If rsget("IsSend") Then
						ampurl = "api_key="&platformApiKey&"&event=[{""user_id"":"""&rsget("UserSeq")&""", ""event_type"":""$identify"", ""user_properties"":{""$set"":{""kakaopush"":""test""},""$prepend"":{""kakaopushhistory"":""test""}}}]"
					Else
						ampurl = "api_key="&platformApiKey&"&event=[{""user_id"":"""&rsget("UserSeq")&""", ""event_type"":""$identify"", ""user_properties"":{""$set"":{""kakaopush"":""control""},""$prepend"":{""kakaopushhistory"":""control""}}}]"
					End If

					'response.write ampurl&"<br>"

					set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP")	'xmlHTTP컨퍼넌트 선언
					oXML.open "POST", "https://api.amplitude.com/httpapi", False
					oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"
'					oXML.setRequestHeader "CharSet", "utf-8" '있어도 되고 없어도 되고
'					oXML.setRequestHeader "Accept","application/x-www/form-urlencoded"
					oXML.send ampurl

					If InStr(oXML.responseText, "success") > 0 Then
						response.write oXML.responseText
					Else
'						response.ContentType = "text/xml"
						response.write oXML.responseText
					End If
					Set oXML = Nothing
					response.write ampurl&"<br>"
				rsget.movenext
				Loop
			Else
				response.write "발송할 데이터가 없습니다."
				response.End
			End If
			rsget.close
			response.End
	End Select



	Function Ceil(ByVal intParam)
	 Ceil = -(Int(-(intParam)))
	End Function
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbDatamartclose.asp" -->