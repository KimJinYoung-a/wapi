<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'상품 등록
Public Function fnlfmallItemReg(iitemid, istrParam, iErrStr, igetMustprice, ilfmallSellYn, iLimityn, iLimitNo, iLimitSold, iItemName, ibasicimageNm)
    Dim objXML, strSql, i, iRbody, iMessage, ProductCode
	Dim xmlDOM, retCode
	Dim REQUEST_XML
	REQUEST_XML = "REQUEST_XML=" & Server.URLEncode(istrParam)

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/interface.do?cmd=saveProductNotiOpt", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(REQUEST_XML)

		If session("ssBctID")="kjy8517" Then
			response.write "<textarea cols=100 rows=30>"&istrParam&"</textarea>"
		End If

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
				xmlDOM.LoadXML iRbody

				If session("ssBctID")="kjy8517" Then
					response.write "<textarea cols=100 rows=30>"&iRbody&"</textarea>"
				End If

				retCode = xmlDOM.getElementsByTagName("ProductInfo/Body/Product/ResultCode").item(0).text
				If retCode = "SUCCESS" Then
					ProductCode = xmlDOM.getElementsByTagName("ProductCode").item(0).text
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET R.lfmallRegdate = getdate()" & VbCrlf
					If (ProductCode <> "") Then
					    strSql = strSql & "	, R.lfmallStatCd = '3'"& VbCRLF					'승인대기
					Else
						strSql = strSql & "	, R.lfmallStatCd = '1'"& VbCRLF					'전송시도
					End If
					strSql = strSql & " ,R.lfmallGoodNo = '" & ProductCode & "'" & VbCrlf
					strSql = strSql & " ,R.lfmallLastUpdate = getdate()"
					strSql = strSql & " ,R.lfmallPrice = '"&igetMustprice&"' " & VbCrlf
					strSql = strSql & " ,R.lfmallSellYn = 'N' "& VbCrlf
					strSql = strSql & " ,R.accFailCNT = 0" & VbCrlf                 ''실패회수 초기화
					strSql = strSql & " ,R.regitemname = i.itemname " & VbCRLF
					strSql = strSql & " ,R.regimageName = '"&ibasicimageNm&"'"& VbCrlf
					strSql = strSql & " FROM db_etcmall.dbo.tbl_lfmall_regitem R" & VbCrlf
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " where R.itemid = " & iitemid
					'rw strSql
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||성공(등록)"
				Else
					iMessage = xmlDOM.getElementsByTagName("ProductInfo/Body/Product/ErrorMessage").item(0).text
					If Len(iMessage) > 0 Then
						iErrStr =  "ERR||"&iitemid&"||실패[등록] "& html2db(iMessage)
					Else
						iErrStr = "ERR||"&iitemid&"||실패[등록] 결과없음"
					End If

					If Len(xmlDOM.getElementsByTagName("ProductInfo/Body/Product/ProductCode").item(0).text) > 0 Then
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCrlf
						strSql = strSql & " SET R.lfmallRegdate = getdate()" & VbCrlf
						If Len(xmlDOM.getElementsByTagName("ProductInfo/Body/Product/ProductCode").item(0).text) > 0 Then
							strSql = strSql & "	, R.lfmallStatCd = '3'"& VbCRLF					'승인대기
						Else
							strSql = strSql & "	, R.lfmallStatCd = '1'"& VbCRLF					'전송시도
						End If
						strSql = strSql & " ,R.lfmallGoodNo = '" & xmlDOM.getElementsByTagName("ProductInfo/Body/Product/ProductCode").item(0).text & "'" & VbCrlf
						strSql = strSql & " ,R.lfmallLastUpdate = getdate()"
						strSql = strSql & " ,R.lfmallPrice = '"&igetMustprice&"' " & VbCrlf
						strSql = strSql & " ,R.lfmallSellYn = 'N' "& VbCrlf
						strSql = strSql & " ,R.accFailCNT = 0" & VbCrlf                 ''실패회수 초기화
						strSql = strSql & " ,R.regitemname = i.itemname " & VbCRLF
						strSql = strSql & " ,R.regimageName = '"&ibasicimageNm&"'"& VbCrlf
						strSql = strSql & " FROM db_etcmall.dbo.tbl_lfmall_regitem R" & VbCrlf
						strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
						strSql = strSql & " where R.itemid = " & iitemid
						'rw strSql
						dbget.execute strSql
					End If
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||LFmall 결과 분석 중에 오류가 발생했습니다.[ERR-REG]"
		End If
	Set objXML = nothing
End Function

Public Function fnLfmallItemEdit(iitemid, istrParam, iErrStr, ibasicimageNm, igetMustprice, iLfmallGoodNo, iitemname)
    Dim objXML, strSql, i, iRbody, iMessage
	Dim xmlDOM, retCode
	Dim REQUEST_XML
	REQUEST_XML = "REQUEST_XML=" & Server.URLEncode(istrParam)

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/interface.do?cmd=saveProductNotiOpt", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(REQUEST_XML)

		If session("ssBctID")="kjy8517" Then
			response.write "<textarea cols=100 rows=30>"&istrParam&"</textarea>"
		End If

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
				xmlDOM.LoadXML iRbody

				If session("ssBctID")="kjy8517" Then
					response.write "<textarea cols=100 rows=30>"&iRbody&"</textarea>"
				End If

				If xmlDOM.getElementsByTagName("ProductInfo/Body/Product/ResultCode").length > 0 Then
					retCode = xmlDOM.getElementsByTagName("ProductInfo/Body/Product/ResultCode").item(0).text
					If retCode = "SUCCESS" Then
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCrlf
						strSql = strSql & " SET lfmallLastUpdate = getdate()"
						strSql = strSql & " ,lfmallprice = '"&igetMustprice&"' " & VbCrlf
						strSql = strSql & " ,accFailCNT = 0" & VbCrlf                 ''실패회수 초기화
						strSql = strSql & " ,regimageName = '"&ibasicimageNm&"'"& VbCrlf
						strSql = strSql & " ,regitemname = i.itemname " & VbCRLF
						strSql = strSql & " FROM db_etcmall.dbo.tbl_lfmall_regItem R" & VbCrlf
						strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
						strSql = strSql & " WHERE R.itemid = " & iitemid
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||성공(수정)"
					Else
						iMessage = xmlDOM.getElementsByTagName("ProductInfo/Header/ErrorMessage").item(0).text
						If Len(iMessage) > 0 Then
							iErrStr =  "ERR||"&iitemid&"||실패[수정] "& html2db(iMessage)
						Else
							iMessage = xmlDOM.getElementsByTagName("ProductInfo/Body/Product/ErrorMessage").item(0).text
							If Len(iMessage) > 0 Then
								iErrStr =  "ERR||"&iitemid&"||실패[수정] "& html2db(iMessage)
							Else
								iErrStr = "ERR||"&iitemid&"||실패[수정] 결과없음."
							End If
						End If
					End If
				Else
					iMessage = xmlDOM.getElementsByTagName("ProductInfo/Header/ErrorMessage").item(0).text
					If Len(iMessage) > 0 Then
						iErrStr =  "ERR||"&iitemid&"||실패[수정] "& html2db(iMessage)
					Else
						iErrStr = "ERR||"&iitemid&"||실패[수정] 결과없음"
					End If
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||LFmall 결과 분석 중에 오류가 발생했습니다.[ERR-EDIT]"
		End If
	Set objXML = nothing
End Function

'상품 상태 수정
Public Function fnLfmallSellYN(iitemid, istrParam, ichgSellYn, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage
	Dim xmlDOM, retCode
	Dim REQUEST_XML
	REQUEST_XML = "REQUEST_XML=" & Server.URLEncode(istrParam)

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/interface.do?cmd=manageProductStatus", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(REQUEST_XML)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
				xmlDOM.LoadXML iRbody
				retCode = xmlDOM.getElementsByTagName("ProductStatus/ResultCode").item(0).text

				If retCode = "SUCCESS" Then
					If ichgSellyn = "Y" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	Set lfmallSellyn = 'Y'"
						strSql = strSql & "	,lfmallLastUpdate = getdate()"
						strSql = strSql & "	From db_etcmall.dbo.tbl_lfmall_regitem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||판매(상태변경)"
					ElseIf ichgSellyn = "N" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	Set lfmallSellyn = 'N'"
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,lfmallLastUpdate = getdate()"
						strSql = strSql & "	From db_etcmall.dbo.tbl_lfmall_regitem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||품절처리(상태변경)"
					End If
				Else
					iMessage = xmlDOM.getElementsByTagName("ProductStatus/ErrorMessage").item(0).text
					If Instr(iMessage, "상품 코드의 판매 상태 및 변경 상태 코드 확인 바랍니다.") > 0 Then
						If ichgSellyn = "Y" Then
							strSql = ""
							strSql = strSql & " UPDATE R"
							strSql = strSql & "	Set lfmallSellyn = 'Y'"
							strSql = strSql & "	,lfmallLastUpdate = getdate()"
							strSql = strSql & "	From db_etcmall.dbo.tbl_lfmall_regitem  R"
							strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
							dbget.Execute(strSql)
							iErrStr =  "OK||"&iitemid&"||판매(상태변경)"
						ElseIf ichgSellyn = "N" Then
							strSql = ""
							strSql = strSql & " UPDATE R"
							strSql = strSql & "	Set lfmallSellyn = 'N'"
							strSql = strSql & "	,accFailCnt = 0"
							strSql = strSql & "	,lfmallLastUpdate = getdate()"
							strSql = strSql & "	From db_etcmall.dbo.tbl_lfmall_regitem  R"
							strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
							dbget.Execute(strSql)
							iErrStr =  "OK||"&iitemid&"||품절처리(상태변경)"
						End If
					Else
						iErrStr = "ERR||"&iitemid&"||실패[상태수정] "& html2db(iMessage)
					End If
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||LFmall 결과 분석 중에 오류가 발생했습니다.[ERR-SOLDOUT]"
		End If
	Set objXML = nothing
End Function

'상품 재고 수정
Public Function fnLfmallQuantity(iitemid, istrParam, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage
	Dim xmlDOM, retCode
	Dim REQUEST_XML
	REQUEST_XML = "REQUEST_XML=" & Server.URLEncode(istrParam)

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/interface.do?cmd=updateProductStock", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(REQUEST_XML)

		If session("ssBctID")="kjy8517" Then
			response.write "<textarea cols=100 rows=30>"&istrParam&"</textarea>"
		End If

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
				xmlDOM.LoadXML iRbody

				If session("ssBctID")="kjy8517" Then
					response.write "<textarea cols=100 rows=30>"&iRbody&"</textarea>"
				End If

				If (xmlDOM.getElementsByTagName("ProductStockInfo/Body/Product/ResultCode").length > 0) then
					retCode = xmlDOM.getElementsByTagName("ProductStockInfo/Body/Product/ResultCode").item(0).text
				Else
					retCode = "FAIL"
				End If

				If retCode = "SUCCESS" Then
					iErrStr =  "OK||"&iitemid&"||성공(재고)"
				Else
					If (xmlDOM.getElementsByTagName("ProductStockInfo/Body/Product/ErrorMessage").length > 0) then
						iMessage = xmlDOM.getElementsByTagName("ProductStockInfo/Body/Product/ErrorMessage").item(0).text
					Else
						iMessage = "FAIL"
					End If
					iErrStr =  "ERR||"&iitemid&"||실패[재고] "& html2db(iMessage)
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||LFmall 결과 분석 중에 오류가 발생했습니다.[ERR-SOLDOUT]"
		End If
	Set objXML = nothing
End Function

'상품 상세 조회
Public Function fnLfmallItemView(iitemid, istrParam, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage
	Dim xmlDOM, retCode, lfmallSellYn, statName
	Dim REQUEST_XML, ProductName, ProductPrice, ProductStatusCode, statusType, ProductCode
	Dim SubNodes, Nodes, outmallOptName
	Dim OptionNm1, OptionValue1, OptionNm2, OptionValue2, CurrentStockQty, ExtraCharge, SoldoutYn, OptionCode, outmallSellyn
	REQUEST_XML = "REQUEST_XML=" & Server.URLEncode(istrParam)

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/interface.do?cmd=getProductInfoListNew", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(REQUEST_XML)
		If session("ssBctID")="kjy8517" Then
			response.write "<textarea cols=100 rows=30>"&istrParam&"</textarea>"
		End If

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
				xmlDOM.LoadXML iRbody
				If session("ssBctID")="kjy8517" Then
					response.write "<textarea cols=100 rows=30>"&iRbody&"</textarea>"
				End If
				retCode = xmlDOM.getElementsByTagName("ProductInfo/Header/ResultCode").item(0).text
				If retCode = "SUCCESS" Then
					ProductCode			= xmlDOM.getElementsByTagName("Body/Product/ProductCode").item(0).Text
					ProductName			= xmlDOM.getElementsByTagName("Body/Product/ProductName").item(0).Text
					ProductPrice		= xmlDOM.getElementsByTagName("Body/Product/ProductPrice").item(0).Text
					ProductStatusCode	= xmlDOM.getElementsByTagName("Body/Product/ProductStatusCode").item(0).Text	'10 : 정보오류 / 20 : 정보부족 / 40 : 승인대기 / 60 : 자동품절 / 70 : 일시중단 / 90 : 정상상품 / 99 : 영구중단
					statName = ""
					Select Case ProductStatusCode
						CASE "10"	statName = "정보오류"
						CASE "20"	statName = "정보부족"
						CASE "40"	statName = "승인대기"
						CASE "60"	statName = "자동품절"
						CASE "70"	statName = "일시중단"
						CASE "90"	statName = "정상상품"
						CASE "99"	statName = "영구중단"
					End Select

					Select Case ProductStatusCode
						Case "90"
							statusType = "7"
							lfmallSellYn = "Y"
						Case "40"
							statusType = "3"
							lfmallSellYn = "N"
						Case "60", "70", "99"
							statusType = "7"
							lfmallSellYn = "N"
					End Select

					'옵션 조회는 일단 보류..옵션 재고 및 상태를 위해서는 꼭 넣어야함..2021-04-27 김진영
					If xmlDOM.getElementsByTagName("Body/Product/Option").length > 0 Then
						Set Nodes = xmlDOM.getElementsByTagName("Body/Product/Option")
							strSql = ""
							strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_lfmall_new_regedoption] WHERE itemid = "&iitemid & VbCRLF
							dbget.Execute strSql

							i = 0
							For each SubNodes in Nodes
								OptionValue1 =""
								OptionValue2 = ""
								outmallSellyn = "Y"

								OptionNm1 =  SubNodes.getElementsByTagName("OptionNm1")(0).Text
								OptionValue1 = SubNodes.getElementsByTagName("OptionValue1")(0).Text
								outmallOptName = ""

								If (SubNodes.getElementsByTagName("OptionNm2").length > 0) Then
									OptionNm2 = SubNodes.getElementsByTagName("OptionNm2")(0).Text
								End If
								If (SubNodes.getElementsByTagName("OptionValue2").length > 0) Then
									OptionValue2 = SubNodes.getElementsByTagName("OptionValue2")(0).Text
								End If

								CurrentStockQty = SubNodes.getElementsByTagName("CurrentStockQty")(0).Text
								ExtraCharge = SubNodes.getElementsByTagName("ExtraCharge")(0).Text
								SoldoutYn = SubNodes.getElementsByTagName("SoldoutYn")(0).Text

								If SoldoutYn = "Y" OR CurrentStockQty < 1 Then
									outmallSellyn = "N"
								End If
								OptionCode = SubNodes.getElementsByTagName("OptionCode")(0).Text
								outmallOptName = Chkiif(OptionValue2="", OptionValue1, OptionValue2)

								' outmallOptName = Trim(OptionValue1) & "," & Trim(OptionValue2)
								' If Right(outmallOptName,1) = "," Then
								' 	outmallOptName = Left(outmallOptName, Len(outmallOptName) - 1)
								' End If

								strSql = ""
								strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_lfmall_new_regedoption] ([itemid], [itemoption], [optionNm1], [optionValue1], [optionNm2], [optionValue2], [outmallSellyn], [outmalllimitno], [lastupdate]) VALUES "
								strSql = strSql & " ('"& iitemid &"', '"& i+1 &"', '"& OptionNm1 &"', '"& optionValue1 &"', '"& optionNm2 &"', '"& optionValue2 &"', '"& outmallSellyn &"', '"& CurrentStockQty &"', GETDATE()) "
								dbget.Execute strSql
								i = i + 1
							Next

						Set Nodes = nothing
					End If
					strSql = ""
					strSql = strSql & "EXEC [db_etcmall].[dbo].[usp_API_LFmall_ItemOptionMapping_Upd] '"& iitemid &"' "
					dbget.Execute strSql

					strSql = ""
					strSql = strSql & " UPDATE R" & VbCRLF
					' strSql = strSql & " SET lfmallPrice = " & ProductPrice & VbCRLF	'가격이 잘 수정되지 않는 부분이 나옴..일단 조회시 가격 수정 주석 / 2021-11-30 김진영
					strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0)"   &VbCRLF
					If (ProductStatusCode = "90") OR (ProductStatusCode = "60") OR (ProductStatusCode = "70") OR (ProductStatusCode = "40") Then
						strSql = strSql & " ,lfmallSellyn='"&lfmallSellYn&"'" & VbCRLF
						strSql = strSql & " ,lfmallstatcd='"&statusType&"'" & VbCRLF
					End If
					strSql = strSql & " ,lastStatCheckDate = getdate()" & VbCRLF
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_lfmall_regItem] R" & VbCRLF
					strSql = strSql & " JOIN ( " & VbCRLF
					strSql = strSql & " 	SELECT R.itemid,count(*) as CNT "
					strSql = strSql & " 	, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
					strSql = strSql & "		FROM db_etcmall.[dbo].[tbl_lfmall_regItem] R " & VbCRLF
					strSql = strSql & " 	JOIN db_etcmall.[dbo].[tbl_lfmall_new_regedoption] Ro " & VbCRLF
					strSql = strSql & " 		on R.itemid = Ro.itemid"   & VbCRLF
					strSql = strSql & "         and Ro.itemid = "&iitemid & VbCRLF
					strSql = strSql & " 	GROUP BY R.itemid "   & VbCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid " & VbCRLF
					strSql = strSql & " WHERE R.itemid="&iitemid & VbCRLF
				    dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||성공_"&statName&"(조회)"
				Else
					iMessage = xmlDOM.getElementsByTagName("ProductInfo/Header/ErrorMessage").item(0).text
					If Len(iMessage) > 0 Then
						iErrStr =  "ERR||"&iitemid&"||실패[조회] "& html2db(iMessage)
					Else
						iErrStr = "ERR||"&iitemid&"||실패[조회] 결과없음"
					End If
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||LFmall 결과 분석 중에 오류가 발생했습니다.[ERR-CHKSTAT]"
		End If
	Set objXML = nothing

	If Err.number <> 0 Then
		iErrStr = "ERR||"&iitemid&"||실패[조회] 오류발생..kjy"
	End If

	On Error Goto 0
End Function

'브랜드 목록 조회
Public Function getLfmallBrandView(istrParam)
    Dim objXML, strSql, i, iRbody, iMessage
	Dim xmlDOM, retCode, lfmallSellYn
	Dim REQUEST_XML, ProductName, ProductPrice, ProductStatusCode, statusType, ProductCode
	Dim SubNodes, Nodes
	REQUEST_XML = "REQUEST_XML=" & Server.URLEncode(istrParam)

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/interface.do?cmd=getBrandList", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(REQUEST_XML)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
				xmlDOM.LoadXML iRbody
				rw iRbody
				rw "--------------------"
				response.end
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||LFmall 결과 분석 중에 오류가 발생했습니다.[ERR-CHKSTAT]"
		End If
	Set objXML = nothing
End Function

'색상 목록 조회
Public Function getLfmallColorView(istrParam)
    Dim objXML, strSql, i, iRbody, iMessage
	Dim xmlDOM, retCode, lfmallSellYn
	Dim REQUEST_XML, ProductName, ProductPrice, ProductStatusCode, statusType, ProductCode
	Dim SubNodes, Nodes
	REQUEST_XML = "REQUEST_XML=" & Server.URLEncode(istrParam)

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/interface.do?cmd=getColorList", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(REQUEST_XML)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
				xmlDOM.LoadXML iRbody
				rw iRbody
				rw "--------------------"
				response.end
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||LFmall 결과 분석 중에 오류가 발생했습니다.[ERR-CHKSTAT]"
		End If
	Set objXML = nothing
End Function

'############################################## 실제 수행하는 API 함수 모음 끝 ############################################
Public Function getLfmallBrandListParameter()
	Dim strRst
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"&VBCRLF
	strRst = strRst & "<BrandInfo>"&VBCRLF
	strRst = strRst & "	<Header>"&VBCRLF
	strRst = strRst & "		<AuthId><![CDATA["&AuthId&"]]></AuthId>"&VBCRLF
	strRst = strRst & "		<AuthKey><![CDATA["&AuthKey&"]]></AuthKey>"&VBCRLF
	strRst = strRst & "		<Format>XML</Format>"&VBCRLF
	strRst = strRst & "		<Charset>UTF-8</Charset>"&VBCRLF
	strRst = strRst & "	</Header>"&VBCRLF
	strRst = strRst & "	<Body>"&VBCRLF
	strRst = strRst & "		<BrandGroupCode>ALL</BrandGroupCode>"&VBCRLF
	strRst = strRst & "	</Body>"&VBCRLF
	strRst = strRst & "</BrandInfo>"&VBCRLF
	getLfmallBrandListParameter = strRst
End Function

Public Function getLfmallColorListParameter()
	Dim strRst
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"&VBCRLF
	strRst = strRst & "<ColorInfo>"&VBCRLF
	strRst = strRst & "	<Header>"&VBCRLF
	strRst = strRst & "		<AuthId><![CDATA["&AuthId&"]]></AuthId>"&VBCRLF
	strRst = strRst & "		<AuthKey><![CDATA["&AuthKey&"]]></AuthKey>"&VBCRLF
	strRst = strRst & "		<Format>XML</Format>"&VBCRLF
	strRst = strRst & "		<Charset>UTF-8</Charset>"&VBCRLF
	strRst = strRst & "	</Header>"&VBCRLF
	strRst = strRst & "	<Body>"&VBCRLF
	strRst = strRst & "		<SearchGb>ALL</SearchGb>"&VBCRLF
	strRst = strRst & "	</Body>"&VBCRLF
	strRst = strRst & "</ColorInfo>"&VBCRLF
	getLfmallColorListParameter = strRst
End Function
%>