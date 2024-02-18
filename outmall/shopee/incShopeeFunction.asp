<%
Dim apiUrl, partnerid, partnerkey, redurl, shopid, shopAuthCode
If application("Svr_Info")="Dev" Then
	redurl			= "http://wapi.10x10.co.kr/outmall/shopee/returnCode.asp"
	apiUrl			= "https://partner.test-stable.shopeemobile.com"
	partnerid		= 1001368
	partnerkey		= "0e4c1f9dd251156386afae7e97f2c46aaa2c7272e27ba2dba11d9ef2a8a9aedc"
	shopid			= 11138
	shopAuthCode	= "6e6c7061634974637568676e61694164"
Else
	redurl			= "http://wapi.10x10.co.kr/outmall/shopee/returnCode.asp"
'	apiUrl			= ""
'	partnerid		= ""
'	partnerkey		= ""
'	shopid			= ""
'	shopAuthCode	= ""

	'추후 아래 내용은 바꿔야함!!!!!!
	apiUrl			= "https://partner.test-stable.shopeemobile.com"
	partnerid		= 1001368
	partnerkey		= "0e4c1f9dd251156386afae7e97f2c46aaa2c7272e27ba2dba11d9ef2a8a9aedc"
	shopid			= 11139
	shopAuthCode	= "6e57714b5678764c79594c744f6d4357"
End If

'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
Public Function fnShopAuth()
	Dim timest, apiPath, basestr, sign, strParam
	timest = getTimestamp()
	apiPath = "/api/v2/shop/auth_partner"
	basestr = partnerid & apiPath & timest
	sign = SHA256SignAndEncode(basestr, partnerkey)

	strParam = "?timestamp="&timest&"&partner_id="&partnerid&"&redirect="&redurl&"&sign="&sign

	response.redirect(apiURL & apiPath & strParam)
	response.end
End Function

Public Function fnAccessToken()
	Dim timest, apiPath, basestr, sign, strParam, strBody, strSql
	Dim access_token, refresh_token
	timest = getTimestamp()
	apiPath = "/api/v2/auth/token/get"
	basestr = partnerid & apiPath & timest
	sign = SHA256SignAndEncode(basestr, partnerkey)
	strParam = "?timestamp="&timest&"&partner_id="&partnerid&"&sign="&sign

	Dim obj, objXML, iRbody, strObj, returnCode, datalist, i
	Set obj = jsObject()
		obj("code") = shopAuthCode
		obj("shop_id") = shopid
		obj("partner_id") = partnerid
		strBody = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & apiPath & strParam, false
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strBody)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				access_token		= strObj.access_token
				refresh_token		= strObj.refresh_token

				strSql = ""
				strSql = strSql & "	DELETE FROM db_etcmall.dbo.tbl_shopee_token "
				dbget.Execute(strSql)

				strSql = ""
				strSql = strSql & "	INSERT INTO db_etcmall.dbo.tbl_shopee_token (access_token, refresh_token, regdate) "
				strSql = strSql & "	VALUES ('"& access_token &"', '"& refresh_token &"', GETDATE()) "
				dbget.Execute(strSql)

				rw "access_token : " & access_token
				rw "refresh_token : " & refresh_token
			Set strObj = nothing
		Else
			rw BinaryToText(objXML.ResponseBody,"utf-8")
		End If
	Set objXML= nothing
End Function

Public Function fnRefreshToken()
	Dim timest, apiPath, basestr, sign, strParam, strBody
	Dim access_token, refresh_token
	timest = getTimestamp()
	apiPath = "/api/v2/auth/access_token/get"
	basestr = partnerid & apiPath & timest
	sign = SHA256SignAndEncode(basestr, partnerkey)
	strParam = "?timestamp="&timest&"&partner_id="&partnerid&"&sign="&sign

	Dim obj, objXML, iRbody, strObj, returnCode, datalist, i
	Set obj = jsObject()
		obj("refresh_token") = getRefreshToken()
		obj("shop_id") = shopid
		obj("partner_id") = partnerid
		strBody = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & apiPath & strParam, false
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strBody)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				access_token		= strObj.access_token
				refresh_token		= strObj.refresh_token

				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_shopee_token "
				strSql = strSql & " SET access_token = '"& access_token &"' "
				strSql = strSql & " , refresh_token = '"& refresh_token &"' "
				strSql = strSql & " , regdate = GETDATE() "
				dbget.Execute(strSql)

				rw "access_token : " & access_token
				rw "refresh_token : " & refresh_token
			Set strObj = nothing
		Else
			rw BinaryToText(objXML.ResponseBody,"utf-8")
		End If
	Set objXML= nothing
End Function

Public Function fnShopeeImageUpload()
	Dim timest, apiPath, basestr, sign, strParam, strBody
	Dim access_token, refresh_token, FPost

	timest = getTimestamp()
	apiPath = "/api/v2/media_space/upload_image"
	basestr = partnerid & apiPath & timest
	sign = SHA256SignAndEncode(basestr, partnerkey)
	strParam = "?timestamp="&timest&"&partner_id="&partnerid&"&sign="&sign

	Set FPost = New FilePost
		' Add Some File
		FPost.AddFile("/outmall/B003896771-1.jpg")
		FPost.AddText "image", "/outmall/B003896771-1.jpg"
		FPost.PostURL(APIURL & apiPath & strParam)
		FPost.FormType("POST")
		If FPost.PostFiles() = True Then 
			rw "========"
			rw "========"
			rw FPost.HTTPAnswer()
			rw "========"
			rw APIURL & apiPath & strParam
		Else
			rw "FAIL"
			rw "STATUS CODE: "&FPost.HTTPStatus()
			rw "HTTP ANSWER: "&FPost.HTTPAnswer()
		End If
	Set FPost = Nothing
	response.end
End Function

function readBinaryFile(sPath)
    Dim objStream
    Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Type = 1 ' adTypeBinary
		objStream.Open
		objStream.LoadFromFile sPath
		readBinaryFile = objStream.Read
		objStream.Close
    Set objStream = Nothing
end function

Public Function fnShopeeCategory()
	Dim timest, apiPath, basestr, sign, strParam, strBody
	Dim access_token
	access_token = getAccessToken()
	timest = getTimestamp()
	apiPath = "/api/v2/product/get_category"
	basestr = partnerid & apiPath & timest & access_token & shopid
	sign = SHA256SignAndEncode(basestr, partnerkey)
	strParam = "?timestamp="&timest&"&partner_id="&partnerid&"&sign="&sign&"&shop_id="&shopid&"&access_token="&access_token


	Dim obj, objXML, iRbody, strObj, returnCode, datalist, i
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", APIURL & apiPath & strParam & "&language=en", false
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
response.write iRbody

			Set strObj = JSON.parse(iRbody)

			Set strObj = nothing
		Else
			rw BinaryToText(objXML.ResponseBody,"utf-8")
		End If
	Set objXML= nothing
response.end
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################

'################################################# 각 기능 별 파라메터 정리시작 ###############################################
Function getTimestamp()
	Dim tmptime
	tmptime = FormatDate(dateadd("h", -9, now()), "0000-00-00 00:00:00")

	If LCase(typename(tmptime)) = "string" Then
		tmptime = CDate(tmptime)
	End If
	getTimestamp = DateDiff("s", "1970-01-01 00:00:00", tmptime)
End Function

Function SHA256SignAndEncode(sIn, sKey)
	 Dim sha
	 Set sha = GetObject("script:"&Server.MapPath("/outmall/shopee/sha256ByPlantext.wsc"))
	 sha.hexcase = 0
	 SHA256SignAndEncode = sha.b64_hmac_sha256(sKey, sIn)
End Function

Public Function getAccessToken()
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT access_token "
	strSql = strSql & " FROM db_etcmall.dbo.tbl_shopee_token "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) then
		getAccessToken = rsget("access_token")
	End If
	rsget.Close
End Function

Public Function getRefreshToken()
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT refresh_token "
	strSql = strSql & " FROM db_etcmall.dbo.tbl_shopee_token "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) then
		getRefreshToken = rsget("refresh_token")
	End If
	rsget.Close
End Function

'################################################# 각 기능 별 파라메터 정리 끝 ###############################################
%>