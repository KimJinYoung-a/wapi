<!-- #include virtual="/outmall/jungsan/include/dvim_brix_crypto-js-master_VB.asp"-->
<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'출고지 등록
Public Function fnCoupangDeliveryReg(iMakerid, iMaeipdiv, iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, isRegYn, strObj, phoneNumner, phoneNumner2
    isRegYn = "N"
	If iMaeipdiv = "U" Then
	    istrParam = "makerID="&iMakerid
		'/////// 우리DB에 일단 저장.. 누락 분이 있다면 통신하지 말고 에러처리 ///////
		strSql = "EXEC [db_etcmall].[dbo].[usp_API_Coupang_deliveryInfo_Add] '"&iMakerid&"' "
		dbget.Execute strSql

		'////// 우편번호에 하이픈(-)이 있으면 수정 처리
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_branddelivery_mapping "
		strSql = strSql & " SET returnZipCode =  "
		strSql = strSql & " 	Case WHEN charindex('-',returnZipCode) > 0 THEN replace(returnZipCode, '-', '')  "
		strSql = strSql & " 	ELSE returnZipCode END "
		strSql = strSql & " WHERE makerid = '"& iMakerid &"' "
		dbget.Execute strSql

		strSql = ""
		strSql = strSql & " SELECT top 1 len(companyContactNumber) as phoneNumner, len(phoneNumber2) as phoneNumner2 "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_branddelivery_mapping "
		strSql = strSql & " WHERE makerid = '"&iMakerid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			phoneNumner = rsget("phoneNumner")
			phoneNumner2 = rsget("phoneNumner2")
		End If
		rsget.Close

		If (phoneNumner < 11) or (phoneNumner > 14) Then
			iErrStr = "ERR||"&iMakerid&"||실패[출고지] 전화번호 길이 오류 11~14자리 요망"
			Exit Function
		End If

		If (phoneNumner2 < 11) or (phoneNumner2 > 14) Then
			iErrStr = "ERR||"&iMakerid&"||실패[출고지] 전화번호 길이 오류2 11~14자리 요망"
			Exit Function
		End If

		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as cnt "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_branddelivery_mapping "
		strSql = strSql & " WHERE makerid = '"&iMakerid&"' "
		strSql = strSql & " and isNull(companyContactNumber, '') <> '' "
		strSql = strSql & " and isNull(phoneNumber2, '') <> '' "
		strSql = strSql & " and isNull(returnZipCode, '') <> '' "
		strSql = strSql & " and isNull(returnAddress, '') <> '' "
		strSql = strSql & " and isNull(returnAddressDetail, '') <> '' "
		strSql = strSql & " and isNull(deliveryCode, '') <> '' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If rsget("cnt") > 0 Then
			isRegYn = "Y"
		End If
		rsget.Close
		'//////////////////////////////////////////////////////////////////////
		If isRegYn = "Y" Then
			On Error Resume Next
			Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
				objXML.open "POST", "http://xapi.10x10.co.kr:8080/Deliveries/Coupang/origin", false
				objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				objXML.Send(istrParam)

				If Err.number <> 0 Then
					iErrStr = "ERR||"&iMakerid&"||실패[출고지] " & Err.Description
					Exit Function
				End If
'rw BinaryToText(objXML.ResponseBody,"utf-8")
				If objXML.Status = "200" OR objXML.Status = "201" Then
					iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
					Set strObj = JSON.parse(iRbody)
						'rw strObj.outboundShippingPlaceCode 이걸로 DB업데이트 하려했는 데, 이미 API서버에서 구현한듯..
						iErrStr = "OK||"&iMakerid&"||성공[출고지]"
					Set strObj = nothing
				Else
					iErrStr = "ERR||"&iMakerid&"||실패[출고지] " & iRbody
				End If
			Set objXML = nothing
		Else
			iErrStr = "ERR||"&iMakerid&"||실패[출고지] 정보누락"
		End If
	Else		'매입 or 특정이라면 출고지는 도봉물류로 통일
		strSql = ""
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_coupang_branddelivery_mapping WHERE makerid='"&iMakerid&"' )"
		strSql = strSql & " BEGIN "
		strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_coupang_branddelivery_mapping "
		strSql = strSql & " (makerid, vendorId, deliveryCode, companyContactNumber, notJeju, outboundShippingPlaceCode, regdate ) VALUES "
		strSql = strSql & " ('"&iMakerid&"', '', 'HANJIN', '1644-6035', '3000', '122412', getdate()) END "
		dbget.Execute strSql
		iErrStr = "OK||"&iMakerid&"||성공[출고지]"
	End If
End Function

'상품 등록
Public Function fnCoupangItemReg(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj
	istrParam = "itemid="&iitemid
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "http://xapi.10x10.co.kr:8080/Products/Coupang", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상품등록] " & Err.Description
			Exit Function
		End If
		'rw objXML.Status
		'rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			'response.write iRbody
			Set strObj = JSON.parse(iRbody)
				strSql = " EXEC db_etcmall.[dbo].[usp_API_Coupang_RegItemInfo_Upd] '"&iitemid&"', 'I' "
				dbget.execute strSql

				iErrStr = "OK||"&iitemid&"||성공[상품등록]"
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[상품등록] "&iMessage
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품등록] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

'상품 조회
Public Function fnCoupangStatChk(iitemid, icoupangGoodno, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, retCode
    Dim productId, regedItemname, statusName, coupangStatcd, strObjitems
    Dim vItemoption, vendorItemId, sellerProductItemId, vOptionName, vItemsu
    Dim firstVendorItemId, regedProductId
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://xapi.10x10.co.kr:8080/Products/Coupang/singleproduct/"&icoupangGoodno, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[조회] " & Err.Description
			Exit Function
		End If
		'rw objXML.Status
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		'response.write iRbody
		'response.end
			Set strObj = JSON.parse(iRbody)
				retCode			= strObj.code
				statusName		= strObj.data.statusName
				productId		= strObj.data.productId
				regedItemname	= strObj.data.sellerProductName

				If retCode = "SUCCESS" Then
					Select Case statusName
						Case "승인완료"					coupangStatcd = 7
						Case "승인반려"					coupangStatcd = 2
						Case "승인대기중","승인요청"	coupangStatcd = 3
						Case "부분승인완료"				coupangStatcd = 4
						Case Else						coupangStatcd = 3
					End Select

					'productId가 존재하면 coupangStatcd를 업데이트 안 함 / 2018-11-08 15:55 김진영 수정
					strSql = ""
					strSql = strSql & " SELECT TOP 1 ISNULL(productId, '') as regedProductId "
					strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_regitem "
					strSql = strSql & " WHERE itemid = '"& iitemid &"' "
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					if (Not rsget.EOF) then
						regedProductId = rsget("regedProductId")
					end if
					rsget.Close

					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_regitem " & VbCRLF
					strSql = strSql & " SET lastConfirmdate = getdate() "& VbCRLF
					If regedProductId = "" or coupangStatcd = 7 Then
						strSql = strSql & "	,coupangStatcd='"&coupangStatcd&"' "
					End If
					strSql = strSql & " ,productId='" & productId & "' "
					strSql = strSql & " WHERE itemid='" & iitemid & "'"& VbCRLF
					dbget.Execute(strSql)

					' If coupangStatcd = 7 Then
						set strObjitems = strObj.data.items
						strSql = ""
						strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_coupang_regedoption WHERE itemid = '"&iitemid&"' "
						dbget.Execute(strSql)
							For i=0 to strObjitems.length-1
								If i = 0 Then
									firstVendorItemId = strObjitems.get(i).vendorItemId
								End If

								vendorItemId = strObjitems.get(i).vendorItemId
								vItemoption	= Split(strObjitems.get(i).externalVendorSku, "_")(1)
								sellerProductItemId = strObjitems.get(i).sellerProductItemId
								vItemsu		= strObjitems.get(i).maximumBuyCount

								If vItemoption <> "0000" Then
									vOptionName = Trim(html2db(replace(strObjitems.get(i).itemName, regedItemname, "")))
									strSql = ""
									strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Coupang_RegItemOption_Add] '"&iitemid&"', '"&vItemoption&"', '"&vendorItemId&"', '"&sellerProductItemId&"', '"&vItemsu&"', '"&vOptionName&"', 'Y' "
									dbget.Execute(strSql)
								Else
									vOptionName = "단일상품"
									strSql = ""
									strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Coupang_RegItemOption_Add] '"&iitemid&"', '"&vItemoption&"', '"&vendorItemId&"', '"&sellerProductItemId&"', '"&vItemsu&"', '"&vOptionName&"', 'N' "
									dbget.Execute(strSql)
								End If
							Next
							strSql = ""
							strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Coupang_RegItemOptionCnt_Upd] '"&iitemid&"', '"&firstVendorItemId&"' "
							dbget.Execute(strSql)
						set strObjitems = nothing
					' End If
					iErrStr = "OK||"&iitemid&"||성공[조회("&statusName&")]"
				Else
					iErrStr = "ERR||"&iitemid&"||실패[조회]NOT SUCCESS"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||실패[조회]통신오류"
		End If
	Set objXML = nothing
End Function

'상품 상태 변경
Public Function fnCoupangSellyn(iitemid, ichgSellyn, ivendorItemId, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, retCode
	istrParam = "vendorItemId="&ivendorItemId
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If ichgSellyn = "Y" Then
			objXML.open "POST", "http://xapi.10x10.co.kr:8080/Products/Coupang/productagainsell", false
		ElseIf ichgSellyn = "N" Then
			objXML.open "POST", "http://xapi.10x10.co.kr:8080/Products/Coupang/stop", false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = ivendorItemId
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				retCode			= strObj.code
				If retCode <> "SUCCESS" Then
					iErrStr = ivendorItemId
				End If
			Set strObj = nothing
		Else
			iErrStr = ivendorItemId
		End If
	Set objXML = nothing
End Function

'상품 삭제
Public Function fnCoupangDelete(iitemid, icoupangGoodno, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, retCode
	On Error Resume Next
	istrParam = "sellerProductId="&icoupangGoodno
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "DELETE", "http://xapi.10x10.co.kr:8080/Products/Coupang", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상품삭제] " & Err.Description
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			'response.write iRbody
			Set strObj = JSON.parse(iRbody)
				retCode			= strObj.code
				If retCode <> "SUCCESS" Then
					iErrStr = "ERR||"&iitemid&"||실패[상품삭제]NOT SUCCESS"
				Else
					strSql = ""
					strSql = strSql &" INSERT INTO [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] " & VBCRLF
					strSql = strSql &" SELECT TOP 1 'coupang', i.itemid, r.CoupangGoodNo, r.CoupangRegdate, getdate(), r.lastErrStr" & VBCRLF
					strSql = strSql &" FROM db_item.dbo.tbl_item as i " & VBCRLF
					strSql = strSql &" JOIN db_etcmall.dbo.tbl_coupang_regItem as r on i.itemid = r.itemid " & VBCRLF
					strSql = strSql &" WHERE i.itemid = "&iitemid & VBCRLF
					dbget.Execute(strSql)

					strSql = ""
					strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_coupang_regitem " & vbcrlf
					strSql = strSql & " WHERE itemid = '"&iitemid&"' "
					dbget.Execute(strSql)

					strSql = ""
					strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_coupang_regedoption " & vbcrlf
					strSql = strSql & " WHERE itemid = '"&iitemid&"' "
					dbget.Execute(strSql)

					strSql = ""
					strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_outmall_API_Que " & vbcrlf
					strSql = strSql & " WHERE itemid = '"&iitemid&"' " & vbcrlf
					strSql = strSql & " and mallid = '"&CMALLNAME&"' " & vbcrlf
					dbget.Execute(strSql)
				End If
				iErrStr = "OK||"&iitemid&"||성공[상품삭제]"
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||실패[상품삭제] 통신오류"
		End If
	Set objXML = nothing
End Function

'상품 가격 수정
Public Function fnCoupangPrice(iitemid, ivendorItemId, imustprice, imustOptionprice, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, retCode
	istrParam = "vendorItemId="&ivendorItemId&"&Price="&imustOptionprice
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "http://xapi.10x10.co.kr:8080/Products/Coupang/price", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = ivendorItemId
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				retCode			= strObj.code
				If retCode <> "SUCCESS" Then
					iErrStr = ivendorItemId
				End If
			Set strObj = nothing
		Else
			iErrStr = ivendorItemId
		End If
	Set objXML = nothing
End Function

'상품 재고 수정
Public Function fnCoupangQuantity(iitemid, ivendorItemId, iquantity, isNameDiff, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, retCode
    If iquantity < 0 OR isNameDiff = 1 Then
    	iquantity = 0
    End If

	istrParam = "vendorItemId="&ivendorItemId&"&quantity="&iquantity
'rw istrParam
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "PUT", "http://xapi.10x10.co.kr:8080/Products/Coupang/productqtychange", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = ivendorItemId
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				retCode			= strObj.code
				If retCode <> "SUCCESS" Then
					iErrStr = ivendorItemId
				Else
					If iquantity = 0 Then
						strSql = ""
						strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_regedoption "
						strSql = strSql & " SET outmallsellyn = 'N' "
						strSql = strSql & " , outmalllimitno = 0 "
						strSql = strSql & " WHERE vendorItemId = '"&ivendorItemId&"' "
						dbget.Execute(strSql)
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = ivendorItemId
		End If
	Set objXML = nothing
End Function

'상품 수정
Public Function fnCoupangItemEdit(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj
	istrParam = "itemid="&iitemid
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "PUT", "http://xapi.10x10.co.kr:8080/Products/Coupang", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상품수정] " & Err.Description
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			strSql = ""
			strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Coupang_RegItemInfo_Upd] '"&iitemid&"', 'R' "
			dbget.Execute(strSql)
			iErrStr = "OK||"&iitemid&"||성공[상품수정]"
		Else
			'iErrStr = "ERR||"&iitemid&"||실패[상품수정] 통신오류"
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[상품수정] "&Replace(iMessage, "'", "")
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품수정] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

'즉시할인쿠폰 등록
Public Function fnCoupangCouponReg(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, iCode
	istrParam = "?idx="&iitemid
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		IF (application("Svr_Info") = "Dev") then
			objXML.open "GET", "http://localhost:62569/coupangnew/coupon/add" & istrParam, false
		Else
			objXML.open "GET", "http://xapi.10x10.co.kr:8080/coupangnew/coupon/add" & istrParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[쿠폰등록] " & Err.Description
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				iCode				= strObj.code
				'rw iRbody
				If iCode <> "SUCCESS" Then
					iErrStr = "ERR||"&iitemid&"||실패[쿠폰등록] "&Replace(iMessage, "'", "")
				Else
					iErrStr = "OK||"&iitemid&"||성공[쿠폰등록]"
				End If
			Set strObj = nothing
		Else
			'iErrStr = "ERR||"&iitemid&"||실패[쿠폰등록] 통신오류"
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[쿠폰등록] "&Replace(iMessage, "'", "")
				Else
					iErrStr = "ERR||"&iitemid&"||실패[쿠폰등록] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

'즉시할인쿠폰 조회
Public Function fnCoupangCouponStat(iitemid, irequestedId, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, iCode
	istrParam = "?requestedId="&irequestedId
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		IF (application("Svr_Info") = "Dev") then
			objXML.open "GET", "http://localhost:62569/coupangnew/coupon/stat" & istrParam, false
		Else
			objXML.open "GET", "http://xapi.10x10.co.kr:8080/coupangnew/coupon/stat" & istrParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[쿠폰조회] " & Err.Description
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				iCode				= strObj.code
				If iCode <> "SUCCESS" Then
					iErrStr = "ERR||"&iitemid&"||실패[쿠폰조회] "&Replace(iMessage, "'", "")
				Else
					iErrStr = "OK||"&iitemid&"||성공[쿠폰조회]"
				End If
			Set strObj = nothing
		Else
			'iErrStr = "ERR||"&iitemid&"||실패[쿠폰조회] 통신오류"
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[쿠폰조회] "&Replace(iMessage, "'", "")
				Else
					iErrStr = "ERR||"&iitemid&"||실패[쿠폰조회] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

'즉시할인 상품 생성
Public Function fnCoupangCouponItemReg(iitemid, icouponId, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, iCode
	istrParam = "?midx="&iitemid&"&couponid="&icouponId
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		IF (application("Svr_Info") = "Dev") then
			objXML.open "GET", "http://localhost:62569/coupangnew/coupon/additem" & istrParam, false
		Else
			objXML.open "GET", "http://xapi.10x10.co.kr:8080/coupangnew/coupon/additem" & istrParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[쿠폰상품등록] " & Err.Description
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				iCode				= strObj.code
				If iCode <> "SUCCESS" Then
					iErrStr = "ERR||"&iitemid&"||실패[쿠폰상품등록] "&Replace(iMessage, "'", "")
				Else
					iErrStr = "OK||"&iitemid&"||성공[쿠폰상품등록]"
				End If
			Set strObj = nothing
		Else
			'iErrStr = "ERR||"&iitemid&"||실패[쿠폰상품등록] 통신오류"
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[쿠폰상품등록] "&Replace(iMessage, "'", "")
				Else
					iErrStr = "ERR||"&iitemid&"||실패[쿠폰상품등록] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

'즉시할인 상품 삭제
Public Function fnCoupangCouponItemDel(iitemid, icouponId, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, iCode
	istrParam = "?midx="&iitemid&"&couponid="&icouponId
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		IF (application("Svr_Info") = "Dev") then
			objXML.open "GET", "http://localhost:62569/coupangnew/coupon/delitem" & istrParam, false
		Else
			objXML.open "GET", "http://xapi.10x10.co.kr:8080/coupangnew/coupon/delitem" & istrParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[쿠폰상품삭제] " & Err.Description
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				iCode				= strObj.code
				If iCode <> "SUCCESS" Then
					iErrStr = "ERR||"&iitemid&"||실패[쿠폰상품삭제] "&Replace(iMessage, "'", "")
				Else
					strSql = ""
					strSql = strSql & " DELETE FROM [db_etcmall].[dbo].[tbl_coupang_CouponItem_detail] "
					strSql = strSql & " WHERE itemid in ( "
					strSql = strSql & " 	SELECT itemid FROM [db_etcmall].[dbo].[tbl_coupang_CouponItem_detail] WHERE itemType = 'D' and midx = "&iitemid&" "
					strSql = strSql & " ) "
					strSql = strSql & " and midx = "&iitemid&" "
					dbget.Execute(strSql)

					strSql = ""
					strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_coupang_CouponItem_detail] "
					strSql = strSql & " WHERE midx="&iitemid&" and itemType = 'D' "
					dbget.Execute(strSql)
					iErrStr = "OK||"&iitemid&"||성공[쿠폰상품삭제]"
				End If
			Set strObj = nothing
		Else
			'iErrStr = "ERR||"&iitemid&"||실패[쿠폰상품삭제] 통신오류"
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[쿠폰상품삭제] "&Replace(iMessage, "'", "")
				Else
					iErrStr = "ERR||"&iitemid&"||실패[쿠폰상품삭제] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

Function fnBrandmaeipdiv(iMakerid)
	Dim strSql
	strSql = strSql & " SELECT TOP 1 maeipdiv "
	strSql = strSql & " FROM db_user.dbo.tbl_user_c "
	strSql = strSql & " WHERE userid = '"& iMakerid &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.EOF) then
	    fnBrandmaeipdiv = rsget("maeipdiv")
	end if
    rsget.Close
End Function

Function getCoupangCouponNotRegCount(iidx)
	Dim strSql, addSql, i
	strSql = ""
	strSql = strSql & " SELECT COUNT(*) as cnt "
	strSql = strSql & " FROM [db_etcmall].[dbo].[tbl_coupang_Coupon_master] "
	strSql = strSql & " WHERE idx in (" & iidx & ")"
	strSql = strSql & " and isNull(requestedId, '') = '' "
	strSql = strSql & " and isNull(couponId, '') = '' "
	strSql = strSql & addSql											'카테고리 매칭 상품만
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getCoupangCouponNotRegCount = rsget("cnt")
	End If
	rsget.Close
End Function

Function getCoupangCouponItemNotRegCount(iidx)
	Dim strSql, addSql, i
	strSql = ""
	strSql = strSql & " SELECT itemid "
	strSql = strSql & " INTO #TMPTBL "
	strSql = strSql & " FROM [db_etcmall].[dbo].[tbl_coupang_CouponItem_detail] "
	strSql = strSql & " WHERE midx = '"& iidx &"' "
	dbget.Execute(strSql)

	strSql = ""
	strSql = strSql & " INSERT INTO #TMPTBL (itemid) "
	strSql = strSql & " SELECT r.itemid "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_etcmall.dbo.tbl_coupang_regItem as r on i.itemid = r.itemid "
	strSql = strSql & " JOIN [db_etcmall].[dbo].[tbl_coupang_CouponCate_detail] as d on i.cate_large = d.cdl and i.cate_mid = d.cdm "
	strSql = strSql & " WHERE d.midx = '"& iidx &"' "
	strSql = strSql & " and i.isusing = 'Y' "
	dbget.Execute(strSql)

	strSql = ""
	strSql = strSql & " SELECT COUNT(*) as CNT "
	strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_regedoption "
	strSql = strSql & " WHERE itemid in (SELECT itemid FROM #TMPTBL GROUP BY itemid) "
	strSql = strSql & " and outmallSellyn = 'Y' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getCoupangCouponItemNotRegCount = rsget("cnt")
	End If
	rsget.Close
End Function

Function getCoupangCouponRequestedId(iidx)
	Dim strSql, addSql, i
	strSql = ""
	strSql = strSql & " SELECT TOP 1 requestedId"
	strSql = strSql & " FROM [db_etcmall].[dbo].[tbl_coupang_Coupon_master] "
	strSql = strSql & " WHERE idx in (" & iidx & ")"
	strSql = strSql & " and isNull(requestedId, '') <> '' "
	strSql = strSql & addSql											'카테고리 매칭 상품만
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getCoupangCouponRequestedId = rsget("requestedId")
	End If
	rsget.Close
End Function

Function getCoupangGoodno(iitemid)
	Dim strSql
	strSql = strSql & " SELECT TOP 1 isnull(coupangGoodno, '') as coupangGoodno "
	strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_regitem "
	strSql = strSql & " WHERE itemid = '"& iitemid &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.EOF) then
	    getCoupangGoodno = rsget("coupangGoodno")
	end if
    rsget.Close
End Function

Function getCoupangVendorItemidList(iitemid)
	Dim strSql
	strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Coupang_VendorItemIdList_Get] '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.EOF) then
	    getCoupangVendorItemidList = rsget.getRows
	end if
    rsget.Close
End Function

Function getCoupangVendorItemidSellNList(iitemid)
	Dim strSql
	strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Coupang_VendorItemIdListSellN_Get] '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.EOF) then
	    getCoupangVendorItemidSellNList = rsget.getRows
	end if
    rsget.Close
End Function

Function getCoupangVendorItemidChkStatList(iitemid)
	Dim strSql
	strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Coupang_VendorItemIdChkStatList_Get] '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.EOF) then
	    getCoupangVendorItemidChkStatList = rsget.getRows
	end if
    rsget.Close
End Function

Function ArrErrStrInfo(iaction, iValue, iitemid, ierrVendorItemId)
	Dim ErrStrComma, strSql
	If iaction = "EditSellYn" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||실패[상태변경] " & ErrStrComma
		Else
			If iValue = "N" Then
				strSql = ""
				strSql = strSql & " UPDATE R"
				strSql = strSql & "	Set coupangSellYn = 'N'"
				strSql = strSql & "	,accFailCnt = 0"
				strSql = strSql & "	,coupangLastUpdate = getdate()"
				strSql = strSql & "	From db_etcmall.dbo.tbl_coupang_regitem  R"
				strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
				dbget.Execute(strSql)
				ArrErrStrInfo = "OK||"&iitemid&"||품절처리[상태변경]"
			Else
				strSql = ""
				strSql = strSql & " UPDATE R"
				strSql = strSql & "	Set coupangSellYn = 'Y'"
				strSql = strSql & "	,coupangLastUpdate = getdate()"
				strSql = strSql & "	From db_etcmall.dbo.tbl_coupang_regitem  R"
				strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
				dbget.Execute(strSql)
				ArrErrStrInfo = "OK||"&iitemid&"||판매[상태변경]"
			End If
		End If
	ElseIf iaction = "PRICE" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||실패[가격수정] " & ErrStrComma
		Else
		    strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_regitem " & VbCRLF
			strSql = strSql & "	SET coupangLastUpdate = getdate() " & VbCRLF
			strSql = strSql & "	, coupangPrice = " & iValue & VbCRLF
			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
			strSql = strSql & " WHERE itemid='" & iitemid & "'"& VbCRLF
			dbget.Execute(strSql)
			ArrErrStrInfo =  "OK||"&iitemid&"||성공[가격수정]"
		End If
	ElseIf iaction = "QTY" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||실패[재고수정] " & ErrStrComma
		Else
		    strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_regedoption "
			strSql = strSql & " SET outmalllimitno = CASE WHEN (A.outmallOptName <> '단일상품') AND (A.outmalloptName <> isNULL(B.optionname,''))  THEN 0 "
			strSql = strSql & " WHEN isnull(B.itemoption, '') = '' AND i.limityn = 'Y' THEN i.limitno - i.limitsold - 5 "
			strSql = strSql & " WHEN isnull(B.itemoption, '') = '' AND i.limityn = 'N' THEN 9999 "
			strSql = strSql & " WHEN isnull(B.itemoption, '') <> '' AND i.limityn = 'Y' THEN B.optlimitno - B.optlimitsold - 5 "
			strSql = strSql & " WHEN isnull(B.itemoption, '') <> '' AND i.limityn = 'N' THEN 9999 END "
			strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_regedoption as A "
			strSql = strSql & " JOIN db_item.dbo.tbl_item as i on A.itemid = i.itemid "
			strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_option as B on A.itemid = B.itemid and A.itemoption = B.itemoption "
			strSql = strSql & " WHERE A.itemid = '"&iitemid&"' "
			dbget.Execute(strSql)
			ArrErrStrInfo =  "OK||"&iitemid&"||성공[재고수정]"
		End If
	End If

End Function

Public Function fnCoupangCateMeta(icatekey)
	Dim access_key, secret_key, vendorId
	Dim authorization

	Dim objXML, xmlDOM, sqlStr, iRbody, strObj
	Dim retCode, iMessage, path, params, url, method
	Dim datalist, i, j, itemlist


	access_key = "0af06fb7-3deb-4ac3-9a84-6d409a26d831"
	secret_key = "5474f1108ac5631e5977d4a6b7a6387426533582"
	vendorId = "A00039305"
	path = "/v2/providers/seller_api/apis/api/v1/marketplace/meta/display-categories/" & icatekey
	params = ""
	url = "https://api-gateway.coupang.com" & path
	method = "GET"
	authorization = generateHmac(path, method, params, access_key, secret_key)

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open method, url & "?" & params, false
		objXML.setRequestHeader "Authorization", authorization
		objXML.setRequestHeader "X-Requested-By", vendorId
		objXML.send()
rw BinaryToText(objXML.ResponseBody,"utf-8")
 		If objXML.Status = "200" OR objXML.Status = "201" Then
 			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
rw "---------------"
rw iRbody
response.end
		end if
	set objXML = nothing

End Function


Function generateHmac(url, method, querystring, apikey, seckey)
    dim signedDate, signature
    dim encString
    dim requestData

    '[timestamp, httpMethod, requestPath, queryString]
    signedDate = generateSignedDate()
    requestData = signedDate & method & url & querystring
    set encString = mac256(requestData, seckey)
    signature = ", signature=" & cstr(encString)

    generateHmac = "CEA algorithm=HmacSHA256, access-key=" & apikey & ", signed-date=" & signedDate & signature
End Function

'/******  MAC Function ******/
' CEA algorithm=HmacSHA256, access-key=0a3a0f34-7852-4ad8-9368-766290b8b1ab, signed-date=190201T042152Z, signature=737374df2de01ad31cc85c14c42faa0711e7d156d8350e57c911b20916d4ba77
'Input String|WordArray , Returns WordArray
Function mac256(ent, seckey)
    Dim encWA
    Set encWA = ConvertUtf8StrToWordArray(ent)
    Dim keyWA
    Set keyWA = ConvertUtf8StrToWordArray(seckey)
    Dim resWA
    Set resWA = CryptoJS.HmacSHA256(encWA, keyWA)
    Set mac256 = resWA
End Function

'Input (Utf8)String|WordArray Returns WordArray
Function ConvertUtf8StrToWordArray(data)
    If (typename(data) = "String") Then
        Set ConvertUtf8StrToWordArray = CryptoJS.enc.Utf8.parse(data)
    Elseif (typename(data) = "JScriptTypeInfo") Then
        On error resume next
        'Set ConvertUtf8StrToWordArray = CryptoJS.enc.Utf8.parse(data.toString(CryptoJS.enc.Utf8))
        Set ConvertUtf8StrToWordArray = CryptoJS.lib.WordArray.create().concat(data) 'Just assert that data is WordArray
        If Err.number>0 Then
            Set ConvertUtf8StrToWordArray = Nothing
        End if
        On error goto 0
    Else
        Set ConvertUtf8StrToWordArray = Nothing
    End if
End Function

Function generateSignedDate()
    Dim nowDateTime

    'ISO TIMEZONE
    nowDateTime = DateAdd("H", -9, now())

    generateSignedDate = ToIsoDate(nowDateTime) & "T" & ToIsoTime(nowDateTime) & "Z"

End Function

Function ToIsoDate(datetime)
    ToIsoDate = CStr(Mid(Year(datetime), 3, 2)) & StrN2(Month(datetime)) & StrN2(Day(datetime))
End Function

Function ToIsoTime(datetime)
    ToIsoTime = StrN2(Hour(datetime)) & StrN2(Minute(datetime)) & StrN2(Second(datetime))
End Function

Function StrN2(n)
    If Len(CStr(n)) < 2 Then StrN2 = "0" & n Else StrN2 = n
End Function
%>