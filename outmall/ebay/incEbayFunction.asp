<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'1.1 Site카테고리조회 API
Public Function fnEbaytGetSiteCate(vToken, vDepth, vCateCode, vGubun, iErrStr)
	Dim objXML, xmlDOM, strRst, iMessage
	Dim buf, iRbody, strObj, i, strSql
	Dim siteCatCode, subCats
	Dim catCode, catName, isLeaf, parentCatCode
	parentCatCode = vCateCode
	Select Case vDepth
		Case "1"		vCateCode = ""
		Case Else		vCateCode = "/" & vCateCode
	End Select
'	On Error Resume Next
	fnEbaytGetSiteCate = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", APISSLURL & "/item/v1/categories/site-cats" & vCateCode, false
		objXML.setRequestHeader "Authorization", "Bearer " & vToken
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				Set subCats = strObj.subCats
					For i = 0 to subCats.length - 1
						catCode = subCats.get(i).catCode
						catName = subCats.get(i).catName
						'IsLeafCategory = subCats.get(i).isLeaf
						If subCats.get(i).isLeaf = "True" Then
							isLeaf = "Y"
						Else
							isLeaf = "N"
						End If
						strSql = ""
						strSql = strSql & " INSERT INTO db_temp.dbo.tbl_ebay_siteCategory "
						strSql = strSql & " (gubun, depth, parentCatCode, catCode, catName, isLeaf) "
						strSql = strSql & " VALUES ('"& vGubun &"', '"& vDepth &"', '"& Chkiif(vDepth="1","0", parentCatCode) &"', '" & catCode & "', '" & catName & "', '" & isLeaf & "') "
						dbget.execute(strSql)
					Next
					iErrStr = "ok"
				Set subCats = nothing
			Set strObj = nothing
			fnEbaytGetSiteCate = true
		Else
			iErrStr = "no 통신오류"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'1.2 ESM카테고리조회 API
Public Function fnEbaytGetCate(vToken, vCode, iErrstr)
	Dim objXML, xmlDOM, strRst, iMessage
	Dim buf, iRbody, strObj, i, strSql
	Dim sdCategoryTree, SDCategoryCode, SDCategoryName, IsLeafCategory, vv
	On Error Resume Next
	fnEbaytGetCate = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", APISSLURL & "/item/v1/categories/sd-cats/" & vCode, false
		objXML.setRequestHeader "Authorization", "Bearer " & vToken
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")

			Set strObj = JSON.parse(iRbody)
				Set sdCategoryTree = strObj.sdCategoryTree
					strSql = ""
					strSql = strSql & " DELETE FROM db_temp.dbo.tbl_ebay_esmCategory "
					dbget.execute(strSql)
					For i = 0 to sdCategoryTree.length - 1
						SDCategoryCode = sdCategoryTree.get(i).SDCategoryCode
						SDCategoryName = sdCategoryTree.get(i).SDCategoryName
						'IsLeafCategory = sdCategoryTree.get(i).IsLeafCategory
						If sdCategoryTree.get(i).IsLeafCategory = "True" Then
							IsLeafCategory = "Y"
						Else
							IsLeafCategory = "N"
						End If
						strSql = ""
						strSql = strSql & " INSERT INTO db_temp.dbo.tbl_ebay_esmCategory "
						strSql = strSql & " (SDCategoryCode, SDCategoryName, IsLeafCategory) "
						strSql = strSql & " VALUES ('" & SDCategoryCode & "', '" & SDCategoryName & "', '" & IsLeafCategory & "') "
						dbget.execute(strSql)
					Next

					strSql = ""
					strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_ebay_esmCategory "
					dbget.execute(strSql)

					strSql = ""
					strSql = strSql & " SELECT SDCategoryCode "
					strSql = strSql & " , CASE WHEN (right(SDCategoryCode, 16) = '0000000000000000') THEN '0' "
					strSql = strSql & " 	   WHEN (right(SDCategoryCode, 12) = '000000000000') and (right(SDCategoryCode, 16) <> '0000000000000000') THEN left(SDCategoryCode, 4) + '0000000000000000' "
					strSql = strSql & " 	   WHEN (right(SDCategoryCode,  8) = '00000000')     and (right(SDCategoryCode, 12) <> '000000000000')     THEN left(SDCategoryCode, 8) + '000000000000' "
					strSql = strSql & " 	   WHEN (right(SDCategoryCode,  4) = '0000')         and (right(SDCategoryCode, 8) <> '00000000')		   THEN left(SDCategoryCode,12) + '00000000' "
					strSql = strSql & " else left(SDCategoryCode, 16) + '0000' "
					strSql = strSql & " end as parentSDCategoryCode "
					strSql = strSql & " , SDCategoryName, IsLeafCategory "
					strSql = strSql & " INTO #CleanTable "
					strSql = strSql & " FROM db_temp.dbo.tbl_ebay_esmCategory  "
					dbget.execute(strSql)

					strSql = ""
					strSql = strSql & " ;WITH CTETABLE(SDCategoryCode, parentSDCategoryCode, SDCategoryName, SDCategoryName2, LV, IsLeafCategory) as ( "
					strSql = strSql & " 	SELECT A.SDCategoryCode, A.parentSDCategoryCode "
					strSql = strSql & " 	, convert(varchar(300), A.SDCategoryName) as SDCategoryName "
					strSql = strSql & " 	, SDCategoryName as SDCategoryName2 "
					strSql = strSql & " 	, 1 "
					strSql = strSql & " 	, A.IsLeafCategory "
					strSql = strSql & " 	FROM #CleanTable A "
					strSql = strSql & " 	WHERE A.parentSDCategoryCode = '0' "
					strSql = strSql & " 	UNION ALL "
					strSql = strSql & " 	SELECT B.SDCategoryCode, B.parentSDCategoryCode "
					strSql = strSql & " 	, convert(varchar(300), C.SDCategoryName + ' > ' + B.SDCategoryName) as SDCategoryName "
					strSql = strSql & " 	, B.SDCategoryName as SDCategoryName2 "
					strSql = strSql & " 	, (C.LV + 1) LV "
					strSql = strSql & " 	, B.IsLeafCategory "
					strSql = strSql & " 	FROM #CleanTable B, "
					strSql = strSql & " 	CTETABLE C "
					strSql = strSql & " 	WHERE B.parentSDCategoryCode = C.SDCategoryCode "
					strSql = strSql & " ) "
					strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_ebay_esmCategory ( SDCategoryCode, parentSDCategoryCode, SDCategoryName, SDCategoryName2, LV, regdate) "
					strSql = strSql & " SELECT SDCategoryCode, parentSDCategoryCode, SDCategoryName, SDCategoryName2, LV, getdate() "
					strSql = strSql & " FROM CTETABLE "
					strSql = strSql & " where IsLeafCategory = 'Y' "
					strSql = strSql & " ORDER BY SDCategoryName, LV "
					dbget.execute(strSql)
					iErrStr = "ok"
				Set sdCategoryTree = nothing
			Set strObj = nothing
			fnEbaytGetCate = true
		Else
			iErrStr = "no 통신오류"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'1.3 Site-ESM카테고리조회 API
Public Function fnEbaytGetMatchCate(vToken, vCode, iErrstr)
	Dim objXML, xmlDOM, strRst, iMessage
	Dim buf, iRbody, strObj, i, strSql
	Dim Gmkt, Iac, gmKtCateCode, gmKtCateName, IacCateCode, IacCateName
	'On Error Resume Next
	fnEbaytGetMatchCate = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", APISSLURL & "/item/v1/categories/sd-cats/" & vCode & "/site-cats" , false
		objXML.setRequestHeader "Authorization", "Bearer " & vToken
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				Set Gmkt = strObj.MatchedCategory.Gmkt
					strSql = ""
					strSql = strSql &  " DELETE FROM [db_etcmall].[dbo].[tbl_ebay_matchCategory] WHERE SDCategoryCode = '"& vCode &"' and gubun = 'G' "
					dbget.execute(strSql)
					For i = 0 to Gmkt.length - 1
						gmKtCateCode = Gmkt.get(i).catCode
						'gmKtCateName = Gmkt.get(i).catName

						strSql = ""
						strSql = strSql &  " INSERT INTO [db_etcmall].[dbo].[tbl_ebay_matchCategory] (SDCategoryCode, siteCateCode, gubun, regdate) VALUES "
						strSql = strSql &  " ('"& vCode &"', '"& gmKtCateCode &"', 'G', getdate()) "
						dbget.execute(strSql)
					Next
				Set Gmkt = nothing

				Set Iac = strObj.MatchedCategory.Iac
					strSql = ""
					strSql = strSql &  " DELETE FROM [db_etcmall].[dbo].[tbl_ebay_matchCategory] WHERE SDCategoryCode = '"& vCode &"' and gubun = 'A' "
					dbget.execute(strSql)
					For i = 0 to Iac.length - 1
						IacCateCode = Iac.get(i).catCode
						'IacCateName = Iac.get(i).catName

						strSql = ""
						strSql = strSql &  " INSERT INTO [db_etcmall].[dbo].[tbl_ebay_matchCategory] (SDCategoryCode, siteCateCode, gubun, regdate) VALUES "
						strSql = strSql &  " ('"& vCode &"', '"& IacCateCode &"', 'A', getdate()) "
						dbget.execute(strSql)
					Next
				Set Iac = nothing
			Set strObj = nothing
			iErrstr = "OK(matched)"
			fnEbaytGetMatchCate = true
		Else
			iErrStr = "no 통신오류"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'3.1 상품등록
Public Function fnEbayItemReg(vToken, iitemid, strParam, byRef iErrStr, iSellCash, iAuctionSellYn, ilimityn, ilimitno, ilimiysold, iitemname, iimageNm)
	Dim objXML, xmlDOM, strRst, iMessage
	Dim buf, iRbody, strObj, i, strSql
	Dim Gmkt, Iac, gmKtCateCode, gmKtCateName, IacCateCode, IacCateName
	'On Error Resume Next
	fnEbayItemReg = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", APISSLURL & "/item/v1/goods?isSync=true" , false
		objXML.setRequestHeader "Authorization", "Bearer " & vToken
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)
response.write BinaryToText(objXML.ResponseBody, "utf-8")
'response.write strParam
response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
rw iRbody
'			iErrstr = "OK(matched)"
			fnEbayItemReg = true
		Else
rw iRbody
'			iErrStr = "no 통신오류"
		End If
	Set objXML = Nothing
response.end
	On Error Goto 0
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################
Function fnEbaytMakeSiteCate(vGubun)
	Dim strSql
	strSql = ""
	strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_ebay_siteCategory WHERE gubun = '"& vGubun &"' "
	dbget.execute(strSql)

	strSql = ""
	strSql = strSql & " ;WITH CTETABLE(catcode, parentCatcode, catname, catname2, LV, isLeaf) as ( "
	strSql = strSql & " 	SELECT A.catcode, A.parentCatcode "
	strSql = strSql & " 	, convert(varchar(300), A.catname) as catname "
	strSql = strSql & " 	, catname as catname2 "
	strSql = strSql & " 	, 1 "
	strSql = strSql & " 	, A.isLeaf "
	strSql = strSql & " 	FROM db_temp.dbo.tbl_ebay_siteCategory  A with(nolock)"
	strSql = strSql & " 	WHERE A.parentCatcode = '0' "
	strSql = strSql & " 	AND A.gubun = '"& vGubun &"' "
	strSql = strSql & " 	UNION ALL "
	strSql = strSql & " 	SELECT B.catcode, B.parentCatcode "
	strSql = strSql & " 	, convert(varchar(300), C.catname + ' > ' + B.catName) as catname "
	strSql = strSql & " 	, B.catName as catname2 "
	strSql = strSql & " 	, (C.LV + 1) LV "
	strSql = strSql & " 	, B.isLeaf "
	strSql = strSql & " 	FROM db_temp.dbo.tbl_ebay_siteCategory  B with(nolock), "
	strSql = strSql & " 	CTETABLE C "
	strSql = strSql & " 	WHERE B.parentCatcode = C.catcode "
	strSql = strSql & " 	AND B.gubun = '"& vGubun &"' "
	strSql = strSql & " ) "

	strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_ebay_siteCategory (cateCode, parentCateCode, cateName, cateName2, LV, regdate, gubun) "
	strSql = strSql & " SELECT catcode, parentCatcode, catname, catname2, LV, getdate(), '"& vGubun &"' "
	strSql = strSql & " FROM CTETABLE "
	strSql = strSql & " WHERE isLeaf = 'Y' "
	strSql = strSql & " GROUP BY catcode, parentCatcode, catname, catname2, LV "
	strSql = strSql & " ORDER BY catname, LV "
	dbget.execute(strSql)
	rw "ok"
End Function

Function dummyDataReg(vGubun, iitemid, iItemName)
	Dim strSql
	Select Case vGubun
		Case "A"
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_auction1010_regitem where itemid="&iitemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_auction1010_regitem "
			strSql = strSql & " 	(itemid, regdate, reguserid, auctionstatCD, regitemname, auctionSellYn)"
			strSql = strSql & " 	VALUES ("&iitemid&", getdate(), '"&session("SSBctID")&"', '1', '"&iItemName&"', 'N')"
			strSql = strSql & " END "
			dbget.Execute strSql
		Case "G"
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_gmarket1010_regitem where itemid="&iitemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_gmarket1010_regitem "
			strSql = strSql & " 	(itemid, regdate, reguserid, gmarketstatCD, regitemname, gmarketSellYn)"
			strSql = strSql & " 	VALUES ("&iitemid&", getdate(), '"&session("SSBctID")&"', '1', '"&iItemName&"', 'N')"
			strSql = strSql & " END "
			dbget.Execute strSql
	End Select
End Function

Function getToken(vGubun)
	Dim sKey, dAttributes, sToken
	sKey = "IbP+5cXPDUGiHz3r62s9Cw=="
	Set dAttributes = Server.CreateObject("Scripting.Dictionary")
		dAttributes.Add "iss", "10x10.co.kr"
		dAttributes.Add "sub", "sell"
		dAttributes.Add "aud", "sa.esmplus.com"
		dAttributes.Add "iat", SecsSinceEpoch
		If vGubun = "A" Then
			dAttributes.Add "ssi", "A:ten10x10"
		ElseIf vGubun = "G" Then
			dAttributes.Add "ssi", "G:ten10x10"
		End If
		getToken = JWTEncode(dAttributes, sKey)
	Set dAttributes = nothing
End Function

Function JWTEncode(dPayload, sSecret)
	Dim sPayload, sHeader, sBase64Payload, sBase64Header
	Dim sSignature, sToken

	sPayload = DictionaryToJSONString(dPayload)
	sHeader  = JWTHeaderDictionary()

	sBase64Payload = SafeBase64Encode(sPayload)
	sBase64Header  = SafeBase64Encode(sHeader)

	sPayload       = sBase64Header & "." & sBase64Payload
	sSignature     = SHA256SignAndEncode(sPayload, sSecret)
	sToken         = sPayload & "." & sSignature

	JWTEncode = sToken
End Function

' SHA256 HMAC
Function SHA256SignAndEncode(sIn, sKey)
	Dim sSignature, sha256

	'Open WSC object to access the encryption function
	Set sha256 = GetObject("script:"&Server.MapPath("/lib/util/sha256.wsc"))

	'SHA256 sign data
	sSignature = sha256.b64_hmac_sha256(sKey, sIn)
	sSignature = Base64ToSafeBase64(sSignature)

	SHA256SignAndEncode = sSignature
End Function

' Returns a static JWT header dictionary
Function JWTHeaderDictionary()
	Dim dOut
	Set dOut = Server.CreateObject("Scripting.Dictionary")
	dOut.Add "typ", "JWT"
	dOut.Add "alg", "HS256"
	dOut.Add "kid", "tenbyten0"
	JWTHeaderDictionary = DictionaryToJSONString(dOut)
End Function
%>
