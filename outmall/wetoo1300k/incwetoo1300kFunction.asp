<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'상품등록
Public Function fnWetoo1300kItemReg(iitemid, strParam, byRef iErrStr, imustprice, ioptcnt, ilimityn, ilimitno, ilimitsold, iimageNm)
    Dim objXML, iRbody, strObj, i, strSql
	Dim returnCode, iMessage, product_code, limitsu
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", wetoo1300kAPIURL & "/enterstore/api/product_info.html", false
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)
'response.write strParam
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[등록] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "00" Then
					product_code = strObj.result.product.product_code

					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET wetoo1300kregdate = getdate()" & VbCrlf
				    strSql = strSql & "	,wetoo1300kStatCd = '7'"& VbCRLF
					strSql = strSql & " ,wetoo1300kGoodNo = '" & product_code & "'" & VbCrlf
					strSql = strSql & " ,wetoo1300klastupdate = getdate()"
					strSql = strSql & " ,wetoo1300kPrice = '"&imustprice&"' " & VbCrlf
					strSql = strSql & " ,wetoo1300ksellYn = 'Y' "& VbCrlf
					strSql = strSql & " ,accFailCNT = 0" & VbCrlf                 ''실패회수 초기화
					strSql = strSql & " ,saleregdate = getdate()"
					strSql = strSql & " ,regimageName = '"&iimageNm&"'"& VbCrlf
					strSql = strSql & " ,regedOptCnt = " & ioptcnt & VbCrlf
					strSql = strSql & " FROM db_etcmall.dbo.tbl_wetoo1300k_regitem R" & VbCrlf
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " where R.itemid = " & iitemid
					dbget.execute strSql

					If ioptcnt = 0 Then
						If (ilimityn="Y") then
							If (ilimitno - ilimitsold - 5) < 1 Then
								limitsu = 0
							Else
								limitsu = ilimitno - ilimitsold - 5
							End If
						Else
							limitsu = CDEFALUT_STOCK
						End If
						strSql = ""
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
						strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						strSql = strSql & " VALUES " & VBCRLF
						strSql = strSql & " ('"&iitemid&"',  '0000', 'wetoo1300k', '', '단일상품', 'Y', '"&ilimityn&"', '"&limitsu&"', '0', getdate()) "
						dbget.Execute strSql
					Else
						strSql = ""
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
						strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						strSql = strSql & " SELECT itemid, itemoption, 'wetoo1300k', '', optionname "
						strSql = strSql & " ,Case WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold <= 5) THEN 'N' " & VBCRLF
						strSql = strSql & " 	 WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold > 5) THEN optsellyn " & VBCRLF
						strSql = strSql & "	Else optsellyn End, optlimityn, " & VBCRLF
						strSql = strSql & " Case WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold <= 5) THEN '0' " & VBCRLF
						strSql = strSql & " 	 WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold > 5) THEN optlimitno - optlimitsold - 5 " & VBCRLF
						strSql = strSql & " 	 WHEN (optlimityn = 'N') THEN '"& CDEFALUT_STOCK &"' End " & VBCRLF
						strSql = strSql & " , optaddprice, getdate() " & VBCRLF
						strSql = strSql & " FROM db_item.dbo.tbl_item_option " & VBCRLF
						strSql = strSql & " WHERE isUsing='Y' and optsellyn='Y' and itemid= '"&iitemid&"' "
						dbget.Execute strSql
					End If
					iErrStr =  "OK||"&iitemid&"||등록성공(상품등록)"
				Else
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상품등록)"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||1300k 결과 분석 중에 오류가 발생했습니다.[ERR-REG-002]"
		End If
	Set objXML= nothing
End Function

'상품부분수정
Public Function fnWetoo1300kPriceSellyn(iitemid, ichgSellYn, strParam, imustPrice, iErrStr)
    Dim objXML, iRbody, strObj, i, strSql
	Dim returnCode, iMessage
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", wetoo1300kAPIURL & "/enterstore/api/product_brief.html", false
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[부분수정] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "00" Then
					'// 상품가격정보 수정
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_wetoo1300k_regitem  " & VbCRLF
					strSql = strSql & "	SET wetoo1300kLastUpdate=getdate() " & VbCRLF
					strSql = strSql & "	, wetoo1300kPrice = " & imustprice & VbCRLF
					If ichgSellYn = "Y" Then
						strSql = strSql & "	, wetoo1300kSellyn = 'Y' " & VbCRLF
					ElseIf ichgSellYn = "N" Then
						strSql = strSql & "	, wetoo1300kSellyn = 'N' " & VbCRLF
					ElseIf ichgSellYn = "X" Then
						strSql = strSql & "	, wetoo1300kSellyn = 'X' " & VbCRLF
					End IF
					strSql = strSql & "	,accFailCnt = 0"& VbCRLF
					strSql = strSql & " Where itemid='" & iitemid & "'"& VbCRLF
					dbget.Execute(strSql)
					If ichgSellYn = "Y" Then
						iErrStr =  "OK||"&iitemid&"||판매(부분수정)"
					ElseIf ichgSellYn = "N" Then
						iErrStr =  "OK||"&iitemid&"||품절처리(부분수정)"
					Else
						iErrStr =  "OK||"&iitemid&"||판매종료(부분수정)"
					End If
				Else
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(부분수정)"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||1300k 결과 분석 중에 오류가 발생했습니다.[ERR-EditSellyn-002]"
		End If
	Set objXML= nothing
End Function

'상품수정
Public Function fnWetoo1300kItemEdit(iitemid, strParam, byRef iErrStr, imustprice, ioptcnt, ilimityn, ilimitno, ilimiysold, iimageNm)
    Dim objXML, iRbody, strObj, i, strSql
	Dim returnCode, iMessage, limitsu
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", wetoo1300kAPIURL & "/enterstore/api/product_update.html", false
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)
'response.write strParam
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[수정] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "00" Then
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET wetoo1300klastupdate = getdate()"
					strSql = strSql & " ,wetoo1300kPrice = '"&imustprice&"' " & VbCrlf
					strSql = strSql & " ,accFailCNT = 0" & VbCrlf                 ''실패회수 초기화
					strSql = strSql & " ,regimageName = '"&iimageNm&"'"& VbCrlf
					strSql = strSql & " ,regedOptCnt = " & ioptcnt & VbCrlf
					strSql = strSql & " ,wetoo1300kSellyn = 'Y' " & VbCrlf
					strSql = strSql & " FROM db_etcmall.dbo.tbl_wetoo1300k_regitem R" & VbCrlf
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " where R.itemid = " & iitemid
					dbget.execute strSql

					strSql = ""
					strSql = strSql & " DELETE FROM db_item.dbo.tbl_OutMall_regedoption WHERE itemid = '"&iitemid&"' and mallid = 'wetoo1300k' "
					dbget.Execute strSql

					If ioptcnt = 0 Then
						If (ilimityn="Y") then
							If (ilimitno - ilimiysold - 5) < 1 Then
								limitsu = 0
							Else
								limitsu = ilimitno - ilimiysold - 5
							End If
						Else
							limitsu = CDEFALUT_STOCK
						End If
						strSql = ""
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
						strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						strSql = strSql & " VALUES " & VBCRLF
						strSql = strSql & " ('"&iitemid&"',  '0000', 'wetoo1300k', '', '단일상품', 'Y', '"&ilimityn&"', '"&limitsu&"', '0', getdate()) "
						dbget.Execute strSql
					Else
						strSql = ""
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
						strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						strSql = strSql & " SELECT itemid, itemoption, 'wetoo1300k', '', optionname "
						strSql = strSql & " ,Case WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold <= 5) THEN 'N' " & VBCRLF
						strSql = strSql & " 	 WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold > 5) THEN optsellyn " & VBCRLF
						strSql = strSql & "	Else optsellyn End, optlimityn, " & VBCRLF
						strSql = strSql & " Case WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold <= 5) THEN '0' " & VBCRLF
						strSql = strSql & " 	 WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold > 5) THEN optlimitno - optlimitsold - 5 " & VBCRLF
						strSql = strSql & " 	 WHEN (optlimityn = 'N') THEN '"& CDEFALUT_STOCK &"' End " & VBCRLF
						strSql = strSql & " , optaddprice, getdate() " & VBCRLF
						strSql = strSql & " FROM db_item.dbo.tbl_item_option " & VBCRLF
						strSql = strSql & " WHERE isUsing='Y' and optsellyn='Y' and itemid= '"&iitemid&"' "
						dbget.Execute strSql
					End If
					iErrStr =  "OK||"&iitemid&"||성공(수정)"
				Else
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(수정)"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||1300k 결과 분석 중에 오류가 발생했습니다.[ERR-EDIT-002]"
		End If
	Set objXML= nothing
End Function

'카테고리 조회
Public Function fnGetCateList()
    Dim objXML, obj, strParam, iRbody, strObj, categoryList, i, strSql
	Dim large_category, middle_category, small_category, detail_category, category_name
	Set obj = jsObject()
		Set obj("header") = jsObject()
			obj("header")("company_code") = company_code
			obj("header")("company_auth") = company_auth
			strParam = obj.jsString
	Set obj = nothing

	strSql = ""
	strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_wetoo1300k_category] "
	dbget.Execute(strSql)

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", wetoo1300kAPIURL & "/enterstore/api/category.html", false
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[카테고리] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
' response.write iRbody
' response.end
			Set strObj = JSON.parse(iRbody)
				Set categoryList = strObj.result.category
				If categoryList.length > 0 Then
					For i=0 to categoryList.length-1
						large_category	= categoryList.get(i).large_category
						middle_category	= categoryList.get(i).middle_category
						small_category	= categoryList.get(i).small_category
						detail_category	= categoryList.get(i).detail_category
						category_name	= categoryList.get(i).category_name

						strSql = ""
						strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_wetoo1300k_category] ([large_category], [middle_category], [small_category], [detail_category], [category_name]) "
						strSql = strSql & " VALUES ('"& large_category &"', '"& middle_category &"', '"& small_category &"', '"& detail_category &"', '"& html2db(category_name) &"') "
						dbget.Execute(strSql)
					Next
					rw "카테고리 완료 건 수 : " & categoryList.length
				End IF
				Set categoryList = nothing
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'############################################## 실제 수행하는 API 함수 모음 끝 ############################################
%>
