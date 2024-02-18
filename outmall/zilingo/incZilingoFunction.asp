<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'상품 등록
Public Function fnZilingoItemReg(iitemid, iitemoption, istrParam, iOrgprice, irateprice, iMultiplerate, iExchangeRate, iquantity, byRef iErrStr)
	Dim objJSON, jsonDOM, strSql, resultCode, productNo, iRbody
	Dim iMessage, AssignedRow
	Dim strObj, isSuccess, productClientId, newItemid
	newItemid = iitemid & "_" & iitemoption
	On Error Resume Next
	Set objJSON= CreateObject("Microsoft.XMLHTTP")
	    objJSON.Open "POST", zilingoAPIURL & "/api/v1/products/upload" , False
		objJSON.setRequestHeader "Content-Type", "application/json"
		objJSON.SetRequestHeader "sellerId", zilingoSELLERID
		objJSON.SetRequestHeader "apiKey", zilingoAPIKEY
		objJSON.SetRequestHeader "locale", zilingoLOCALE
		objJSON.Send(istrParam)
' BinaryToText(objJSON.ResponseBody,"euc-kr")
'response.write objJSON.ResponseTEXT
		If objJSON.Status = "200" Then
			iRbody = BinaryToText(objJSON.ResponseBody,"euc-kr")
			'response.write iRbody
			Set strObj = JSON.parse(iRbody)
				isSuccess = strObj.STATUS
				If isSuccess = "SUCCESS" Then
					productClientId = strObj.productClientId

					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET zilingoRegdate = getdate()" & VbCrlf
					If (productClientId <> "") Then
					    strSql = strSql & "	, zilingoStatCd = '3'"& VbCRLF					'승인대기
					End If
					strSql = strSql & " ,zilingoTmpGoodno = '" & productClientId & "'" & VbCrlf
					strSql = strSql & " ,zilingolastupdate = getdate()"
					strSql = strSql & " ,zilingoPrice = '"&irateprice&"' " & VbCrlf
					strSql = strSql & "	,regOrgprice = " & iOrgprice & VbCRLF
					strSql = strSql & " ,zilingosellYn = 'N' "& VbCrlf
					strSql = strSql & " ,accFailCNT = 0" & VbCrlf                 ''실패회수 초기화
					strSql = strSql & " ,regimageName = '"&iimageNm&"'"& VbCrlf
					strSql = strSql & " ,multiplerate = '"&iMultiplerate&"' " & vbcrlf
					strSql = strSql & " ,exchangeRate = '"&iExchangeRate&"' " & vbcrlf
					strSql = strSql & " ,quantity = '"&iquantity&"' " & vbcrlf
					strSql = strSql & " FROM db_etcmall.dbo.tbl_zilingo_regitem R" & VbCrlf
					strSql = strSql & " where R.itemid = " & iitemid
					strSql = strSql & " and R.itemoption = '"&iitemoption&"' "
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||"&iitemoption&"||[UPLOAD]성공"
				Else
					iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||[UPLOAD] "& strObj.REASON
				End If
			Set strObj = Nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||ZILINGO 결과 분석 중에 오류가 발생했습니다.[ERR-REG]"
		End If
	Set objJSON= nothing
End Function

'상품 조회
Public Function fnZilingoTmpGoodNo(iitemid, iitemoption, istrTmpGoodNo, byRef iErrStr)
	Dim objJSON, iRbody, strObj, zilingoGoodNo, strSql, strStatus, i, zilingoSKUId
	On Error Resume Next
	Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objJSON.Open "GET", zilingoAPIURL & "/api/v1/products/byClientId/"&istrTmpGoodNo&"" , False
		objJSON.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objJSON.SetRequestHeader "sellerId", zilingoSELLERID
		objJSON.SetRequestHeader "apiKey", zilingoAPIKEY
		objJSON.SetRequestHeader "locale", zilingoLOCALE
		objJSON.Send()
		If objJSON.Status = "200" Then
			iRbody = BinaryToText(objJSON.ResponseBody,"euc-kr")
			'response.write iRbody
			Set strObj = JSON.parse(iRbody)
				zilingoGoodNo = strObj.productId
				strStatus	  = strObj.status
				For i=0 to strObj.SKUs.length-1
					zilingoSKUId = strObj.SKUs.get(i).zilingoSKUId
				Next
				strSql = ""
				strSql = strSql & " UPDATE R" & VbCrlf
				strSql = strSql & " SET lastStatCheckDate = getdate()" & VBCRLF
				If strStatus = "NEEDS_ACTION" Then					'반려인듯
					strSql = strSql & " ,zilingoStatCd = '40'"& VbCRLF
					strSql = strSql & " ,zilingoSellyn = 'N'"& VbCRLF
				ElseIf strStatus = "APPROVED" Then					'승인인듯
					strSql = strSql & " ,zilingoStatCd = '7'"& VbCRLF
					strSql = strSql & " ,zilingoSellyn = 'Y'"& VbCRLF
				Else												'그외는 승인대기
					strSql = strSql & " ,zilingoStatCd = '3'"& VbCRLF
					strSql = strSql & " ,zilingoSellyn = 'N'"& VbCRLF
				End If
				strSql = strSql & " ,zilingoSkuGoodNo = '" & zilingoSKUId & "'" & VbCrlf
				strSql = strSql & " ,zilingoGoodno = '" & zilingoGoodNo & "'" & VbCrlf
				strSql = strSql & " FROM db_etcmall.dbo.tbl_zilingo_regitem R" & VbCrlf
				strSql = strSql & " WHERE R.itemid = " & iitemid
				strSql = strSql & " and R.itemoption = '"&iitemoption&"' "
				dbget.execute strSql
				iErrStr =  "OK||"&iitemid&"||"&iitemoption&"||[CHKSTAT]성공"
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||ZILINGO 결과 분석 중에 오류가 발생했습니다.[ERR-CHKSTAT]"
		End If
	Set objJSON= nothing
	On Error Goto 0
End Function

'재고 조회
Public Function fnZilingoSKUGoodNo(iitemid, iitemoption, istrParam, byRef iErrStr)
	Dim objJSON, iRbody, strObj, strSql, i, quantity
	On Error Resume Next
	Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objJSON.Open "POST", zilingoAPIURL & "/api/v1/products/quantities/forSKUs" , False
		objJSON.setRequestHeader "Content-Type", "application/json"
		objJSON.SetRequestHeader "sellerId", zilingoSELLERID
		objJSON.SetRequestHeader "apiKey", zilingoAPIKEY
		objJSON.SetRequestHeader "locale", zilingoLOCALE
		objJSON.Send(istrParam)
		If objJSON.Status = "200" Then
			iRbody = BinaryToText(objJSON.ResponseBody,"euc-kr")
			Set strObj = JSON.parse(iRbody)
				For i=0 to strObj.zilingoSKUQuantities.length-1
					quantity = strObj.zilingoSKUQuantities.get(i).quantity
				Next
				strSql = ""
				strSql = strSql & " UPDATE R" & VbCrlf
				strSql = strSql & " SET quantity = '" & quantity & "'" & VbCrlf
				strSql = strSql & " FROM db_etcmall.dbo.tbl_zilingo_regitem R" & VbCrlf
				strSql = strSql & " WHERE R.itemid = " & iitemid
				strSql = strSql & " and R.itemoption = '"&iitemoption&"' "
				dbget.execute strSql
				iErrStr =  "OK||"&iitemid&"||"&iitemoption&"||[CHKQTY]성공"
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||ZILINGO 결과 분석 중에 오류가 발생했습니다.[ERR-CHKQTY]"
		End If
	Set objJSON= nothing
	On Error Goto 0
End Function

'재고 수정
Public Function fnZilingoEditQuantity(iitemid, iitemoption, imaylimitEa, istrParam, byRef iErrStr)
	Dim objJSON, iRbody, strObj, strSql, i, quantity
	On Error Resume Next
	Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objJSON.Open "POST", zilingoAPIURL & "/api/v1/products/updateQuantities" , False
		objJSON.setRequestHeader "Content-Type", "application/json"
		objJSON.SetRequestHeader "sellerId", zilingoSELLERID
		objJSON.SetRequestHeader "apiKey", zilingoAPIKEY
		objJSON.SetRequestHeader "locale", zilingoLOCALE
		objJSON.Send(istrParam)
		If objJSON.Status = "200" Then
			iRbody = BinaryToText(objJSON.ResponseBody,"euc-kr")
			'response.write iRbody
			Set strObj = JSON.parse(iRbody)
				isSuccess = strObj.STATUS
				If isSuccess = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " UPDATE R SET " & VbCrlf
					If imaylimitEa <= 0 Then
						strSql = strSql & "  quantity = quantity - " & imaylimitEa & VbCrlf
					Else
						strSql = strSql & "  quantity = quantity + " & imaylimitEa & VbCrlf
					End If
					strSql = strSql & " ,zilingolastupdate = getdate() "
					strSql = strSql & " FROM db_etcmall.dbo.tbl_zilingo_regitem R" & VbCrlf
					strSql = strSql & " WHERE R.itemid = " & iitemid
					strSql = strSql & " and R.itemoption = '"&iitemoption&"' "
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||"&iitemoption&"||[EDITQTY]성공"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||ZILINGO 결과 분석 중에 오류가 발생했습니다.[ERR-EDITQTY]"
		End If
	Set objJSON= nothing
	On Error Goto 0
End Function

'재고 수정 0으로
Public Function fnZilingoEditQuantityZero(iitemid, iitemoption, imaylimitEa, istrParam, byRef iErrStr)
	Dim objJSON, iRbody, strObj, strSql, i, quantity
	On Error Resume Next
	Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objJSON.Open "POST", zilingoAPIURL & "/api/v1/products/updateQuantities" , False
		objJSON.setRequestHeader "Content-Type", "application/json"
		objJSON.SetRequestHeader "sellerId", zilingoSELLERID
		objJSON.SetRequestHeader "apiKey", zilingoAPIKEY
		objJSON.SetRequestHeader "locale", zilingoLOCALE
		objJSON.Send(istrParam)
		If objJSON.Status = "200" Then
			iRbody = BinaryToText(objJSON.ResponseBody,"euc-kr")
			'response.write iRbody
			Set strObj = JSON.parse(iRbody)
				isSuccess = strObj.STATUS
				If isSuccess = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " UPDATE R SET " & VbCrlf
					strSql = strSql & " quantity = 0 "
					strSql = strSql & " ,zilingoSellyn = 'N' "
					strSql = strSql & " ,accFailCnt = 0 "
					strSql = strSql & " ,zilingolastupdate = getdate() "
					strSql = strSql & " FROM db_etcmall.dbo.tbl_zilingo_regitem R" & VbCrlf
					strSql = strSql & " WHERE R.itemid = " & iitemid
					strSql = strSql & " and R.itemoption = '"&iitemoption&"' "
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||"&iitemoption&"||품절처리"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||ZILINGO 결과 분석 중에 오류가 발생했습니다.[ERR-EDITSELLYN]"
		End If
	Set objJSON= nothing
	On Error Goto 0
End Function

'상품 가격 수정
Public Function fnZilingoItemPrice(iitemid, iitemoption, istrParam, iOrgprice, irateprice, iMultiplerate, iExchangeRate, byRef iErrStr)
	Dim objJSON, jsonDOM, strSql, resultCode, productNo, iRbody
	Dim iMessage, AssignedRow
	Dim strObj, isSuccess
	On Error Resume Next
	Set objJSON= CreateObject("Microsoft.XMLHTTP")
	    objJSON.Open "POST", zilingoAPIURL & "/api/v1/products/updatePrice/forProduct" , False
		objJSON.setRequestHeader "Content-Type", "application/json"
		objJSON.SetRequestHeader "sellerId", zilingoSELLERID
		objJSON.SetRequestHeader "apiKey", zilingoAPIKEY
		objJSON.SetRequestHeader "locale", zilingoLOCALE
		objJSON.Send(istrParam)
		If objJSON.Status = "200" Then
			iRbody = BinaryToText(objJSON.ResponseBody,"euc-kr")
			Set strObj = JSON.parse(iRbody)
				isSuccess	= strObj.STATUS
				If isSuccess = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET zilingolastupdate = getdate()"
					strSql = strSql & " ,zilingoPrice = '"&irateprice&"' " & VbCrlf
					strSql = strSql & "	,regOrgprice = " & iOrgprice & VbCRLF
					strSql = strSql & " ,accFailCNT = 0" & VbCrlf                 ''실패회수 초기화
					strSql = strSql & " ,multiplerate = '"&iMultiplerate&"' " & vbcrlf
					strSql = strSql & " ,exchangeRate = '"&iExchangeRate&"' " & vbcrlf
					strSql = strSql & " FROM db_etcmall.dbo.tbl_zilingo_regitem R" & VbCrlf
					strSql = strSql & " where R.itemid = " & iitemid
					strSql = strSql & " and R.itemoption = '"&iitemoption&"' "
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||"&iitemoption&"||[PRICE]성공"
				Else
					'iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||[PRICE] "& db2html(strObj.MESSAGE)
					iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||[PRICE]실패"
				End If
			Set strObj = Nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||ZILINGO 결과 분석 중에 오류가 발생했습니다.[ERR-PRICE]"
		End If
	Set objJSON= nothing
End Function

'상품 가격 수정
Public Function fnZilingoItemPriceBySkuNo(iitemid, iitemoption, istrParam, iOrgprice, irateprice, iMultiplerate, iExchangeRate, byRef iErrStr)
	Dim objJSON, jsonDOM, strSql, resultCode, productNo, iRbody
	Dim iMessage, AssignedRow
	Dim strObj, isSuccess
	On Error Resume Next
	Set objJSON= CreateObject("Microsoft.XMLHTTP")
	    objJSON.Open "POST", zilingoAPIURL & "/api/v1/products/updatePrice/forSKUs" , False
		objJSON.setRequestHeader "Content-Type", "application/json"
		objJSON.SetRequestHeader "sellerId", zilingoSELLERID
		objJSON.SetRequestHeader "apiKey", zilingoAPIKEY
		objJSON.SetRequestHeader "locale", zilingoLOCALE
		objJSON.Send(istrParam)
		If objJSON.Status = "200" Then
			iRbody = BinaryToText(objJSON.ResponseBody,"euc-kr")
			'response.write iRbody
			'response.end
			Set strObj = JSON.parse(iRbody)
				isSuccess	= strObj.status
				If isSuccess = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET zilingolastupdate = getdate()"
					strSql = strSql & " ,zilingoPrice = '"&irateprice&"' " & VbCrlf
					strSql = strSql & "	,regOrgprice = " & iOrgprice & VbCRLF
					strSql = strSql & " ,accFailCNT = 0" & VbCrlf                 ''실패회수 초기화
					strSql = strSql & " ,multiplerate = '"&iMultiplerate&"' " & vbcrlf
					strSql = strSql & " ,exchangeRate = '"&iExchangeRate&"' " & vbcrlf
					strSql = strSql & " FROM db_etcmall.dbo.tbl_zilingo_regitem R" & VbCrlf
					strSql = strSql & " where R.itemid = " & iitemid
					strSql = strSql & " and R.itemoption = '"&iitemoption&"' "
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||"&iitemoption&"||[PRICE]성공"
				Else
					'iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||[PRICE] "& db2html(strObj.MESSAGE)
					iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||[PRICE]실패"
				End If
			Set strObj = Nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||ZILINGO 결과 분석 중에 오류가 발생했습니다.[ERR-PRICE]"
		End If
	Set objJSON= nothing
End Function

'카테고리 정보 얻기
Public Function fnZilingoSubCategory()
	Dim objXML, iRbody, jsResult, strParam, i, j, lp
	Dim attributes, attributeChoices
	Dim colors, sizes, capacities
	Dim depth1Id, depth1Name, depth2Id, depth2Name, depth3Id, depth3Name, isOptional, isMultiSelectable, optFlag, multiFlag
	Dim colorId, colorName
	Dim sizeId, sizeName
	Dim capacitiesId, capacitiesName
	Dim topTemparr

	strSql = ""
	strSql = strSql & " SELECT ROW_NUMBER() OVER (ORDER BY depth3Code ASC) AS RowNum, depth3Code "
	strSql = strSql & " INTO #TBL1 "
	strSql = strSql & " FROM db_etcmall.[dbo].[tbl_zilingo_category] "
	strSql = strSql & " GROUP BY depth3Code "
	strSql = strSql & " ORDER BY depth3Code asc "
	dbget.execute strSql

	strSql = ""
	'strSql = strSql & " SELECT depth3Code FROM #TBL1 WHERE RowNum <= 100 "
	'strSql = strSql & " SELECT depth3Code FROM #TBL1 WHERE RowNum > 100 and RowNum <= 200 "
	strSql = strSql & " SELECT depth3Code FROM #TBL1 WHERE RowNum > 200 "
	rsget.Open strSql,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		topTemparr = rsget.getRows
	End If
	rsget.Close

	For lp = 0 To Ubound(topTemparr, 2)
	'	On Error Resume Next
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		    objXML.Open "GET", zilingoAPIURL & "/api/v1/subCategories/byId/"&topTemparr(0, lp)&"" , False
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.SetRequestHeader "sellerId", zilingoSELLERID
			objXML.SetRequestHeader "apiKey", zilingoAPIKEY
			objXML.SetRequestHeader "locale", zilingoLOCALE
			objXML.Send()
			iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
			If objXML.Status = "200" Then
				SET jsResult = JSON.parse(iRbody)
					depth1Id	= html2db(jsResult.id)
					depth1Name	= html2db(jsResult.name)

					SET attributes = jsResult.attributes
						For i=0 to attributes.length-1
							depth2Id			= html2db(attributes.get(i).id)
							depth2Name			= html2db(attributes.get(i).name)
							isOptional			= html2db(attributes.get(i).isOptional)
							isMultiSelectable	= html2db(attributes.get(i).isMultiSelectable)
							If isOptional = "True" Then
								optFlag = "Y"
							Else
								optFlag = "N"
							End If

							If isMultiSelectable = "True" Then
								multiFlag = "Y"
							Else
								multiFlag = "N"
							End If

							SET attributeChoices = attributes.get(i).attributeChoices
								For j=0 to attributeChoices.length-1
									depth3Id	= html2db(attributeChoices.get(j).id)
									depth3Name	= html2db(attributeChoices.get(j).name)
									strSql = ""
									strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_zilingo_subCategory (gubun, depth1Id, depth1Name, depth2Id, depth2Name, depth3Id, depth3Name, isOptional, isMultiSelectable) VALUES "
									strSql = strSql & " ('attributeChoices', '"&depth1Id&"', '"&depth1Name&"', '"&depth2Id&"', '"&depth2Name&"', '"&depth3Id&"', '"&depth3Name&"', '"&optFlag&"', '"&multiFlag&"') "
									dbget.execute strSql
								Next
	'						rw "----------------------------------------------------------------------"
	'						rw ""
							SET attributeChoices = nothing
						Next
					SET attributes = Nothing

					SET colors = jsResult.colors
						For i=0 to colors.length-1
							colorId		= html2db(colors.get(i).id)
							colorName	= html2db(colors.get(i).name)
							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_zilingo_subCategory (gubun, depth1Id, depth1Name, depth2Id, depth2Name) VALUES "
							strSql = strSql & " ('colors', '"&depth1Id&"', '"&depth1Name&"', '"&colorId&"', '"&colorName&"') "
							dbget.execute strSql
	'						rw "----------------------------------------------------------------------"
	'						rw ""
						Next
					SET colors = Nothing

					SET sizes = jsResult.sizes
						For i=0 to sizes.length-1
							sizeId		= html2db(sizes.get(i).id)
							sizeName	= html2db(sizes.get(i).name)
							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_zilingo_subCategory (gubun, depth1Id, depth1Name, depth2Id, depth2Name) VALUES "
							strSql = strSql & " ('sizes', '"&depth1Id&"', '"&depth1Name&"', '"&sizeId&"', '"&sizeName&"') "
							dbget.execute strSql
	'						rw "----------------------------------------------------------------------"
	'						rw ""
						Next
					SET sizes = Nothing

					SET capacities = jsResult.capacities
						For i=0 to capacities.length-1
							capacitiesId	= html2db(capacities.get(i).id)
							capacitiesName	= html2db(capacities.get(i).name)
							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_zilingo_subCategory (gubun, depth1Id, depth1Name, depth2Id, depth2Name) VALUES "
							strSql = strSql & " ('capacities', '"&depth1Id&"', '"&depth1Name&"', '"&capacitiesId&"', '"&capacitiesName&"') "
							dbget.execute strSql
	'						rw "----------------------------------------------------------------------"
	'						rw ""
						Next
					SET capacities = Nothing

				SET jsResult = nothing
				response.write "OK||subCategory||성공"
			End If
		Set objXML = Nothing
	Next
'	On Error Goto 0
End Function

'############################################## 실제 수행하는 API 함수 모음 끝 ############################################

'################################################# 각 기능 별 파라메터 정리시작 ###############################################
'재고 조회 JSON
Public Function fnZilingoQuantitySearchJSON(istrSKUGoodNo)
	Dim strRst
	strRst = ""
	strRst = strRst & "{"
	strRst = strRst & "	""zilingoSKUIds"": ["""&istrSKUGoodNo&"""]"
	strRst = strRst & "}"
	fnZilingoQuantitySearchJSON = strRst
End Function

'재고 수정 JSON
Public Function fnZilingoQuantityEditJSON(iitemid, iitemoption, iquantity, imaylimitEa, istrSKUGoodNo)
	Dim strRst, strSql
	Dim vLimityn, vLimitNo, vLimitSold, vIsUsing, vSellYn, limitEA, DEFAULTQTY, maySellAvailQty
	Dim oIsusing, oOptsellyn, oOptlimitno, oOptlimitsold

	DEFAULTQTY = 999
	strSql = ""
	strSql = strSql & " SELECT TOP 1 limityn, limitno, limitsold, isusing, sellyn "
	strSql = strSql & " FROM db_item.dbo.tbl_item "
	strSql = strSql & " WHERE itemid = '"&iitemid&"' "
	rsget.Open strSql,dbget,1
	If not rsget.EOF Then
		vLimityn	= rsget("limityn")
		vLimitNo	= rsget("limitno")
		vLimitSold	= rsget("limitsold")
		vIsUsing	= rsget("isusing")
		vSellYn		= rsget("sellyn")
	End If
	rsget.Close

	'iquantity : 현재 10x10 질링고 SCM에 등록된 수량
	If vIsUsing <> "Y" OR vSellYn <> "Y" Then
		limitEA = -1 * iquantity
	Else
		If iitemoption = "0000" Then
			If vLimityn = "N" Then
				limitEA = DEFAULTQTY - iquantity
			Else
				maySellAvailQty = vLimitNo - vLimitSold - 5
				If maySellAvailQty < 1 Then
					limitEA = -1 * iquantity
				Else
					limitEA = maySellAvailQty - iquantity
				End If
			End If
		Else
			If vLimityn = "N" Then
				limitEA = DEFAULTQTY - iquantity
			Else
				strSql = ""
				strSql = strSql & " SELECT TOP 1 isusing, optsellyn, optlimitno, optlimitsold "
				strSql = strSql & " FROM db_item.dbo.tbl_item_option "
				strSql = strSql & " WHERE itemid = '"&iitemid&"' "
				strSql = strSql & " and itemoption = '"&iitemoption&"' "
				rsget.Open strSql,dbget,1
				If not rsget.EOF Then
					oIsusing		= rsget("isusing")
					oOptsellyn		= rsget("optsellyn")
					oOptlimitno		= rsget("optlimitno")
					oOptlimitsold	= rsget("optlimitsold")
				End If
				rsget.Close
				maySellAvailQty = oOptlimitno - oOptlimitsold - 5
				If (maySellAvailQty < 1) OR (oIsusing <> "Y") OR (oOptsellyn <> "Y") Then
					limitEA = -1 * iquantity
				Else
					limitEA = maySellAvailQty - iquantity
				End If
			End If
		End If
	End If
'	limitEA = 999

	imaylimitEa = limitEA
	strRst = ""
	strRst = strRst & "{"
	strRst = strRst & "	""skuDeltaQuantities"": [{"
	strRst = strRst & "		""zilingoSKUId"": """&istrSKUGoodNo&""","
	strRst = strRst & "		""deltaQuantity"": "&limitEA&""
	strRst = strRst & "	}]"
	strRst = strRst & "}"
	fnZilingoQuantityEditJSON = strRst
End Function

'재고 수정 0으로 JSON
Public Function fnZilingoQuantitySoldOutJSON(iitemid, iitemoption, iquantity, imaylimitEa, istrSKUGoodNo)
	Dim strRst
	strRst = ""
	strRst = strRst & "{"
	strRst = strRst & "	""skuDeltaQuantities"": [{"
	strRst = strRst & "		""zilingoSKUId"": """&istrSKUGoodNo&""","
	strRst = strRst & "		""deltaQuantity"": -10000"
	strRst = strRst & "	}]"
	strRst = strRst & "}"
	fnZilingoQuantitySoldOutJSON = strRst
End Function
'################################################# 각 기능 별 파라메터 정리 끝 ###############################################
%>