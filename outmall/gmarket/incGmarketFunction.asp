<%

'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'Additem 상품 기본정보 등록
Public Function fnGmarketItemReg(iitemid, istrParam, byRef iErrStr, iimageNm)
	Dim objXML, xmlDOM, strSql, goodsCd, iResult, iComment
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketAPIURL&"/v1/ItemService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(istrParam)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/AddItem"
		objXML.send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[AddItem] " & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
'response.write Replace(objXML.responseText,"soap:","")
'response.end
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "AddItemResponse" Then
				iResult = xmlDOM.getElementsByTagName("AddItemResult ").item(0).getAttribute("Result")
				If iResult = "Success" Then
					goodsCd = xmlDOM.getElementsByTagName("AddItemResult ").item(0).getAttribute("GmktItemNo")
					strSql = ""
					strSql = strSql & " UPDATE R"
					strSql = strSql & "	Set GmarketLastUpdate = getdate() "
					strSql = strSql & "	, GmarketGoodNo = '" & goodsCd & "'"
					strSql = strSql & "	, regImageName = '"&iimageNm&"' "
					strSql = strSql & "	, accFailCnt = 0"
					strSql = strSql & "	, GmarketRegdate = isNULL(GmarketRegdate, getdate())"
				    strSql = strSql & "	, GmarketStatCd=(CASE WHEN isNULL(GmarketStatCd, -1) < 3 then 3 ELSE GmarketStatCd END)"
				    strSql = strSql & "	, APIadditem = 'Y'"
					strSql = strSql & "	From db_etcmall.dbo.tbl_gmarket_regItem  R"
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||[AddItem]성공"
				Else
					iComment = replace(xmlDOM.getElementsByTagName("AddItemResult ").item(0).getAttribute("Comment"), "'", "")
					iErrStr = "ERR||"&iitemid&"||[AddItem] "& iComment
				End If
			Else
				iErrStr = "ERR||"&iitemid&"||[AddItem] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'AddOfficialInfo 상품 정보고시항목 등록
Public Function fnGmarketItemInfoCd(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, iMessage, iResult, iComment
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketAPIURL&"/v1/ItemService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/AddOfficialInfo"
		objXML.send(strParam)
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[AddInfoCd] " & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "AddOfficialInfoResponse" Then
				iResult = xmlDOM.getElementsByTagName("AddOfficialInfoResult ").item(0).getAttribute("Result")
				If iResult = "Success" Then
					strSql = ""
					strSql = strSql & " UPDATE R"
					strSql = strSql & "	Set APIaddgosi = 'Y'"
					strSql = strSql & "	From db_etcmall.dbo.tbl_gmarket_regItem R"
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||[AddInfoCd]성공"
				Else
					iComment = replace(xmlDOM.getElementsByTagName("AddOfficialInfoResult ").item(0).getAttribute("Comment"), "'", "")
					iErrStr = "ERR||"&iitemid&"||[AddInfoCd] "& iComment
				End If
			Else
				iErrStr = "ERR||"&iitemid&"||[AddInfoCd] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'AddOfficialInfo 어린이인증항목 등록
Public Function fnGmarketItemChildren(iitemid, strParam, byRef iErrStr)
'response.write strParam
'response.end
	Dim objXML, xmlDOM, strSql, iMessage, iResult, iComment
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketAPIURL&"/v1/ItemService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/AddIntegrateSafeCert"
		objXML.send(strParam)
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[AddCert] " & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
'response.write Replace(objXML.responseText,"soap:","")
'response.end
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "AddIntegrateSafeCertResponse" Then
				iResult = xmlDOM.getElementsByTagName("AddIntegrateSafeCertResult ").item(0).getAttribute("Result")
				If iResult = "Success" Then
					iErrStr =  "OK||"&iitemid&"||[AddCert]성공"
				Else
					iComment = replace(xmlDOM.getElementsByTagName("AddIntegrateSafeCertResult ").item(0).getAttribute("Comment"), "'", "")
					iErrStr = "ERR||"&iitemid&"||[AddCert] "& iComment
				End If
			Else
				iErrStr = "ERR||"&iitemid&"||[AddCert] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'상품별 반품 배송비 등록/수정
Public Function fnGmarketReturnFee(iitemid, strParam, iReturnFee, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, iMessage, iResult, iComment
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketAPIURL&"/v1/ItemService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/AddItemReturnFee"
		objXML.send(strParam)
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[AddFee] " & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "AddItemReturnFeeResponse" Then
				iResult = xmlDOM.getElementsByTagName("AddItemReturnFeeResult ").item(0).getAttribute("Result")
				If iResult = "Success" Then
					strSql = ""
					strSql = strSql & " UPDATE R"
					strSql = strSql & "	Set returnShippingFee = '"&iReturnFee&"'"
					strSql = strSql & "	From db_etcmall.dbo.tbl_gmarket_regItem R"
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||[AddFee]성공"
				Else
					iComment = replace(xmlDOM.getElementsByTagName("AddItemReturnFeeResult ").item(0).getAttribute("Comment"), "'", "")
					iErrStr = "ERR||"&iitemid&"||[AddFee] "& iComment
				End If
			Else
				iErrStr = "ERR||"&iitemid&"||[AddFee] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'AddItemOption 상품 옵션정보 등록
Public Function fnGmarketOPTReg(iitemid, strParam, byRef iErrStr, ilimityn, ilimitno, ilimitsold)
	Dim objXML, xmlDOM, strSql, iMessage, iResult, iComment, ocount, Toptionname, Tlimitno, Tlimitsold, Tlimityn, Titemsu
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketAPIURL&"/v1/ItemService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/AddItemOption"
		objXML.send(strParam)
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[AddOPT] " & Err.Description
			Exit Function
		End If
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))

			If (session("ssBctID")="kjy8517") Then
				rw Replace(objXML.responseText,"soap:","")
			End If

			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "AddItemOptionResponse" Then
				iResult = xmlDOM.getElementsByTagName("AddItemOptionResult ").item(0).getAttribute("Result")
				If iResult = "Success" Then
					strSql = ""
					strSql = strSql & " DELETE FROM db_item.dbo.tbl_OutMall_regedoption WHERE itemid = '"&iitemid&"' and mallid = 'gmarket1010' "
					dbget.Execute strSql

					strSql = ""
					strSql = strSql &  "SELECT count(*) as cnt "
					strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
					strSql = strSql & " WHERE itemid=" & iitemid
					strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
						ocount = rsget("cnt")
					rsget.Close

					If ocount = 0 Then
						Toptionname		= "단일상품"
						Tlimitno		= ilimitno
						Tlimitsold		= ilimitsold
						Tlimityn		= ilimityn
						If (Tlimityn="Y") then
							If (Tlimitno - Tlimitsold - 5) < 1 Then
								Titemsu = 0
							Else
								Titemsu = Tlimitno - Tlimitsold - 5
							End If
						Else
							Titemsu = 999
						End If
						strSql = ""
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
						strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						strSql = strSql & " VALUES " & VBCRLF
						strSql = strSql & " ('"&iitemid&"', '0000', 'gmarket1010', '', '단일상품', 'Y', '"&Tlimityn&"', '"&Titemsu&"', '0', getdate()) "
						dbget.Execute strSql
					Else
						strSql = ""
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
						strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						strSql = strSql & " SELECT itemid, itemoption, 'gmarket1010', '', optionname "
						strSql = strSql & " ,Case WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold <= 5) THEN 'N' " & VBCRLF
						strSql = strSql & " 	 WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold > 5) THEN optsellyn " & VBCRLF
						strSql = strSql & "	Else optsellyn End, optlimityn, " & VBCRLF
						strSql = strSql & " Case WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold <= 5) THEN '0' " & VBCRLF
						strSql = strSql & " 	 WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold > 5) THEN optlimitno - optlimitsold - 5 " & VBCRLF
						strSql = strSql & " 	 WHEN (optlimityn = 'N') THEN '999' End " & VBCRLF
						strSql = strSql & " , optaddprice, getdate() " & VBCRLF
						strSql = strSql & " FROM db_item.dbo.tbl_item_option " & VBCRLF
						strSql = strSql & " WHERE isUsing='Y' and optsellyn='Y' and itemid= '"&iitemid&"' " & VBCRLF
						dbget.Execute strSql
					End If
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_gmarket_regItem SET "
					strSql = strSql & "	APIaddopt = 'Y'"
					strSql = strSql & " , regedOptCnt = " & ocount
					strSql = strSql & " WHERE itemid = " & iitemid
					dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||[AddOPT]성공"
				Else
					iComment = replace(xmlDOM.getElementsByTagName("AddItemOptionResult ").item(0).getAttribute("Comment"), "'", "")
					iErrStr = "ERR||"&iitemid&"||[AddOPT] "& iComment
				End If
			Else
				iErrStr = "ERR||"&iitemid&"||[AddOPT] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'상품 주문(가격, 재고) 등록/수정 AddPrice
Public Function fnGmarketItemAddPrice(iitemid, strParam, imustPrice, idisplayDate, ichgSellyn, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, iMessage, iResult, iComment
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketAPIURL&"/v1/ItemService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/AddPrice"
		objXML.send(strParam)
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[AddPrice] " & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
'response.write Replace(objXML.responseText,"soap:","")
'response.end
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "AddPriceResponse" Then
				iResult = xmlDOM.getElementsByTagName("AddPriceResult ").item(0).getAttribute("Result")
				If iResult = "Success" Then
					If ichgSellyn = "Y" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	Set displayDate = '"&idisplayDate&"' "
						strSql = strSql & "	,GmarketPrice = '"&imustPrice&"' "
						strSql = strSql & "	,GmarketStatCd = 7"
						strSql = strSql & "	,GmarketSellYn = 'Y'"
						strSql = strSql & "	,GmarketLastUpdate = getdate()"
						strSql = strSql & "	From db_etcmall.dbo.tbl_gmarket_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||[AddPrice]판매"
					Else
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	Set GmarketSellYn = 'N'"
						strSql = strSql & "	,GmarketStatCd = 7"
						'strSql = strSql & "	,GmarketPrice = '"&imustPrice&"' "	'품절로 돌렸을 때 GmarketPrice안 바뀌는듯
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,GmarketLastUpdate = getdate()"
						strSql = strSql & "	From db_etcmall.dbo.tbl_gmarket_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||[AddPrice]품절"
					End If
				Else
					iComment = replace(xmlDOM.getElementsByTagName("AddPriceResult ").item(0).getAttribute("Comment"), "'", "")
					iErrStr = "ERR||"&iitemid&"||[AddPrice] "& iComment
				End If
			Else
				iErrStr = "ERR||"&iitemid&"||[AddPrice] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

''상품 이미지 수정 EditItemImage
Public Function fnGmarketEditImg(iitemid, strParam, byRef iErrStr, iimageNm)
	Dim objXML, xmlDOM, strSql, iMessage, iResult, iComment
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketAPIURL&"/v1/ItemService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/EditItemImage"
		objXML.send(strParam)
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[EditImg] " & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
'response.write Replace(objXML.responseText,"soap:","")
'response.end
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "EditItemImageResponse" Then
				iResult = xmlDOM.getElementsByTagName("EditItemImageResult ").item(0).getAttribute("Result")
				If iResult = "Success" Then
					strSql = ""
					strSql = strSql & " UPDATE R"
					strSql = strSql & "	Set regImageName = '"&iimageNm&"' "
					strSql = strSql & "	From db_etcmall.dbo.tbl_gmarket_regItem  R"
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||[EditImg]성공"
				Else
					iComment = replace(xmlDOM.getElementsByTagName("EditItemImageResult ").item(0).getAttribute("Comment"), "'", "")
					iErrStr = "ERR||"&iitemid&"||[EditImg] "& iComment
				End If
			Else
				iErrStr = "ERR||"&iitemid&"||[EditImg] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

''G9상품 등록 AddG9Item
Public Function fnG9ItemReg(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, iMessage, iResult, iComment, iG9GoodNo
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketAPIURL&"/v1/ItemService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/AddG9Item"
		objXML.send(strParam)
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[REGG9] " & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
'response.write Replace(objXML.responseText,"soap:","")
'response.end
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "AddG9ItemResponse" Then
				iResult = xmlDOM.getElementsByTagName("AddG9ItemResult ").item(0).getAttribute("Result")
				If iResult = "Success" Then
					iG9GoodNo = xmlDOM.getElementsByTagName("AddG9ItemResult ").item(0).getAttribute("GmktItemNo")
					strSql = ""
					strSql = strSql & " UPDATE R"
					strSql = strSql & "	Set G9GoodNo = '"&iG9GoodNo&"' "
					strSql = strSql & "	From db_etcmall.dbo.tbl_gmarket_regItem  R"
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||[REGG9]성공"
				Else
					iComment = replace(xmlDOM.getElementsByTagName("AddG9ItemResult ").item(0).getAttribute("Comment"), "'", "")
					iErrStr = "ERR||"&iitemid&"||[REGG9] "& iComment
				End If
			Else
				iErrStr = "ERR||"&iitemid&"||[REGG9] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'제조사/브랜드 조회 AddMakerBrand
Public Function fnGmarketAddMaker(strParam)
	Dim objXML, xmlDOM, iResult, iComment
	Dim rsMakerName, rsBrandName, rsMakerNo, rsBrandNo

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketAPIURL&"/v1/ItemService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/AddMakerBrand"
		objXML.send(strParam)

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
			'response.write objXML.responseText
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "AddMakerBrandResponse" Then
				iResult = xmlDOM.getElementsByTagName ("AddMakerBrandResult ").item(0).getAttribute("Result")
				iComment = xmlDOM.getElementsByTagName ("AddMakerBrandResult ").item(0).getAttribute("Comment")
				If iResult = "Success" Then
					rsMakerName = xmlDOM.getElementsByTagName ("AddMakerBrandResult ").item(0).getAttribute("MakerName")
					rsBrandName = xmlDOM.getElementsByTagName ("AddMakerBrandResult ").item(0).getAttribute("BrandName")
					rsMakerNo	= xmlDOM.getElementsByTagName ("AddMakerBrandResult ").item(0).getAttribute("MakerNo")
					rsBrandNo	= xmlDOM.getElementsByTagName ("AddMakerBrandResult ").item(0).getAttribute("BrandNo")

					rw "제조사명 : " & rsMakerName
					rw "브랜드명 : " & rsBrandName
					rw "제조사번호 : " & rsMakerNo
					rw "브랜드번호 : " & rsBrandNo
					rw "성공유무 : " & iComment
				ElseIf iResult = "Fail" Then
					rw "제조사명 : " & rsMakerName
					rw "브랜드명 : " & rsBrandName
					rw "제조사번호 : " & rsMakerNo
					rw "브랜드번호 : " & rsBrandNo
					rw "성공유무 : " & iComment
				End If
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'판매자 주소(반품/교환) 배송지 등록 AddAddressBook
Public Function fnGmarketAddAddressBook(strParam)
	Dim objXML, xmlDOM, iResult, iComment
	Dim rsAddressCode, rsAddressTitle, rsName, rsPhone1, rsPhone2, rsZipcode, rsAddress1, rsAddress2, rsBundleNo

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketAPIURL&"/v1/SellerService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/AddAddressBook"
		objXML.send(strParam)

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))

			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "AddAddressBookResponse" Then
				iResult = xmlDOM.getElementsByTagName ("AddAddressBookResult ").item(0).getAttribute("Result")
				iComment = xmlDOM.getElementsByTagName ("AddAddressBookResult ").item(0).getAttribute("Comment")
				If iResult = "Success" Then
					rsAddressCode	= xmlDOM.getElementsByTagName ("AddAddressBookResult ").item(0).getAttribute("AddressCode")
					rsAddressTitle	= xmlDOM.getElementsByTagName ("AddAddressBookResult ").item(0).getAttribute("AddressTitle")
					rsName			= xmlDOM.getElementsByTagName ("AddAddressBookResult ").item(0).getAttribute("Name")
					rsPhone1		= xmlDOM.getElementsByTagName ("AddAddressBookResult ").item(0).getAttribute("Phone1")
					rsPhone2		= xmlDOM.getElementsByTagName ("AddAddressBookResult ").item(0).getAttribute("Phone2")
					rsZipcode		= xmlDOM.getElementsByTagName ("AddAddressBookResult ").item(0).getAttribute("Zipcode")
					rsAddress1		= xmlDOM.getElementsByTagName ("AddAddressBookResult ").item(0).getAttribute("Address1")
					rsAddress2		= xmlDOM.getElementsByTagName ("AddAddressBookResult ").item(0).getAttribute("Address2")
					rsBundleNo		= xmlDOM.getElementsByTagName ("AddAddressBookResult ").item(0).getAttribute("BundleNo")

					rw "배송지번호 : " & rsAddressCode
					rw "주소명 : " & AddressTitle
					rw "이름 : " & rsName
					rw "전화번호 : " & rsPhone1
					rw "핸드폰번호 : " & rsPhone2
					rw "우편번호 : " & rsZipcode
					rw "주소1 : " & rsAddress1
					rw "주소2 : " & rsAddress2
					rw "묶음번호 : " & rsBundleNo
					rw "성공유무 : " & iComment
				ElseIf iResult = "Fail" Then
					rw "배송지번호 : " & rsAddressCode
					rw "주소명 : " & AddressTitle
					rw "이름 : " & rsName
					rw "전화번호 : " & rsPhone1
					rw "핸드폰번호 : " & rsPhone2
					rw "우편번호 : " & rsZipcode
					rw "주소1 : " & rsAddress1
					rw "주소2 : " & rsAddress2
					rw "묶음번호 : " & rsBundleNo
					rw "성공유무 : " & iComment
				End If
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'판매자 주소(반품/교환) 배송지 조회 RequestAddressBook
Public Function fnGmarketRequestAddressBook(strParam)
	Dim objXML, xmlDOM, iResult, iComment

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketAPIURL&"/v1/SellerService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/RequestAddressBook"
		objXML.send(strParam)

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
			response.write Replace(objXML.responseText,"soap:","")
		response.End
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

Public Function fnGmarketCateGet()
	Dim objXML, xmlDOM, iResult, LagrgeNode, MiddleNode, SmallNode, DetailNode, DetailNode2, i, j, k, l, m
	Dim Depth1Name, Depth2Name, Depth3Name, Depth4Name, tmpDepth4Name
	Dim Depth1Code, Depth2Code, Depth3Code, Depth4Code, strSql

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & gmarketAPIURL&"/v1/Category/Category.xml"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.send()

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
			Set LagrgeNode = xmlDOM.SelectNodes("/CATEGORY/LARGE_CATEGORY")
				If Not (LagrgeNode Is Nothing) Then
					strSql = ""
					strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_gmarket_category2 "
					dbget.Execute(strSql)
					For i = 0 To LagrgeNode.length - 1
						Depth1Code = LagrgeNode(i).attributes.GetNamedItem("id").value
						Depth1Name = LagrgeNode(i).attributes.GetNamedItem("name").value
						Set MiddleNode = LagrgeNode(i).SelectNodes("./MIDDLE_CATEGORY")
							If Not (MiddleNode Is Nothing) Then
								For j = 0 To MiddleNode.length - 1
									Depth2Code = MiddleNode(j).attributes.GetNamedItem("id").value
									Depth2Name = MiddleNode(j).attributes.GetNamedItem("name").value
									Set SmallNode = MiddleNode(j).SelectNodes("./SMALL_CATEGORY")
										If Not (SmallNode Is Nothing) Then
											For k = 0 To SmallNode.length - 1
												Depth3Code = SmallNode(k).attributes.GetNamedItem("id").value
												Depth3Name = SmallNode(k).attributes.GetNamedItem("name").value
												strSql = ""
												strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_gmarket_category2 "
												strSql = strSql & "(depthCode, Depth1Nm, depth1Code, Depth2Nm, depth2Code, Depth3Nm, Depth4Nm, depth4Code, isChildrenCate, isLifeCate, isElecCate, regdate) VALUES "
												strSql = strSql & "('"&Depth3Code&"', '"&Depth1Name&"', '"&Depth1Code&"', '"&Depth2Name&"', '"&Depth2Code&"', '"&Depth3Name&"', '', '', '', '', '', getdate())"
												dbget.Execute(strSql)
												Set DetailNode = SmallNode(k).SelectNodes("./CLASS")
													If Not (DetailNode Is Nothing) Then
														For l = 0 To DetailNode.length - 1
															tmpDepth4Name = DetailNode(l).attributes.GetNamedItem("name").value
															Set DetailNode2 = DetailNode(l).SelectNodes("./CLASS_VALUE")
																If Not (DetailNode2 Is Nothing) Then
																	For m = 0 To DetailNode2.length - 1
																		Depth4Code = DetailNode2(m).attributes.GetNamedItem("id").value
																		Depth4Name = tmpDepth4Name & " "& DetailNode2(m).attributes.GetNamedItem("name").value
																		strSql = ""
																		strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_gmarket_category2 "
																		strSql = strSql & "(depthCode, Depth1Nm, depth1Code, Depth2Nm, depth2Code, Depth3Nm, Depth4Nm, depth4Code, isChildrenCate, isLifeCate, isElecCate, regdate) VALUES "
																		strSql = strSql & "('"&Depth3Code&"', '"&Depth1Name&"', '"&Depth1Code&"', '"&Depth2Name&"', '"&Depth2Code&"', '"&Depth3Name&"', '"&Depth4Name&"', '"&Depth4Code&"', '', '', '', getdate())"
																		dbget.Execute(strSql)
																	Next
																End If
															Set DetailNode2 = nothing
														Next
													End If
												Set DetailNode = nothing
											Next
										End If
									Set SmallNode = nothing
								Next
							End If
						Set MiddleNode = nothing
					Next
				End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
		iErrStr = "OK||카테고리 Get"
	Set objXML = nothing
	On Error Goto 0
End Function

'상품 기본정보 수정
Public Function fnGmarketIteminfoEdit(iitemid, iGmarketGoodNo, iItemName, iErrStr, istrParam)
	Dim objXML, xmlDOM, strSql, goodsCd, iResult, iComment
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketAPIURL&"/v1/ItemService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(istrParam)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/AddItem"
		objXML.send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[EditInfo] " & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
'response.write Replace(objXML.responseText,"soap:","")
'response.end
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "AddItemResponse" Then
				iResult = xmlDOM.getElementsByTagName("AddItemResult ").item(0).getAttribute("Result")
				If iResult = "Success" Then
					goodsCd = xmlDOM.getElementsByTagName("AddItemResult ").item(0).getAttribute("GmktItemNo")
					strSql = ""
					strSql = strSql & " UPDATE R"
					strSql = strSql & "	Set GmarketLastUpdate = getdate() "
					strSql = strSql & "	, accFailCnt = 0"
					strSql = strSql & " , regitemname = '"&html2db(iitemname)&"'" & vbcrlf
					strSql = strSql & "	From db_etcmall.dbo.tbl_gmarket_regItem  R"
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||[EditInfo]성공"
				Else
					iComment = replace(xmlDOM.getElementsByTagName("AddItemResult ").item(0).getAttribute("Comment"), "'", "")
					iErrStr = "ERR||"&iitemid&"||[EditInfo] "& iComment
				End If
			Else
				iErrStr = "ERR||"&iitemid&"||[EditInfo] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################

'################################################# 각 기능 별 파라메터 정리시작 ###############################################
'정보고시 Soap XML
Public Function getGmarketInfoCdParameter(iitemid, iGmarketPrdno)
	Dim strSQL, strRst1, strRst2, strRst3
	Dim mallinfodiv, mallinfoCd, infoContent
	strRst1 = ""
	strRst1 = strRst1 & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst1 = strRst1 & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst1 = strRst1 & "	<soap:Header>"
	strRst1 = strRst1 & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst1 = strRst1 & "			<encTicket>"&gmarketTicket&"</encTicket>"
	strRst1 = strRst1 & "		</EncTicket>"
	strRst1 = strRst1 & "	</soap:Header>"
	strRst1 = strRst1 & "	<soap:Body>"
	strRst1 = strRst1 & "		<AddOfficialInfo xmlns=""http://tpl.gmarket.co.kr/"">"
	'--------------------------------  쿼리부분 시작 --------------------------------
	strSQL = ""
	strSQL = strSQL & " SELECT top 100 M.* , " & vbcrlf
	strSQL = strSQL & " CASE WHEN (M.infoCd='00001') THEN '상세정보 별도표기' " & vbcrlf
	strSQL = strSQL & " 	 WHEN (M.infoCd='10000') THEN '관련법 및 소비자분쟁해결기준에 따름' " & vbcrlf
	strSQL = strSQL & " 	 WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035'  " & vbcrlf
	strSQL = strSQL & " 	 WHEN LEN( isNull(F.infocontent, '')) < 2 THEN '상품 상세 참고'  " & vbcrlf
	strSQL = strSQL & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent " & vbcrlf
	strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M  " & vbcrlf
	strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv  " & vbcrlf
	strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid " & vbcrlf
	strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd  " & vbcrlf
	strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&iitemid&"'  " & vbcrlf
	strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='" & iitemid &"' " & vbcrlf
	strSQL = strSQL & " WHERE M.mallid = 'gmarket' and IC.itemid='"&iitemid&"'"
	rsget.CursorLocation = adUseClient
	rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) then
		mallinfodiv = CInt(rsget("mallinfodiv"))
		If mallinfodiv = "47" Then
			mallinfodiv = "40"
		ElseIf mallinfodiv = "48" Then
			mallinfodiv = "41"
		End If
		strRst2 = "					<AddOfficialInfo GmktItemNo="""&iGmarketPrdno&""" GroupCode="""&mallinfodiv&""">"
		Do until rsget.EOF
		    mallinfoCd  = rsget("mallinfoCd")
		    infoContent = rsget("infoContent")
			strRst2 = strRst2 & "		<SubInfoList Code="""&mallinfoCd&""" AddYn=""Y"" AddValue="""&replaceRst(infoContent)&""" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"
			rsget.MoveNext
		Loop
		strRst2 = strRst2 & "		</AddOfficialInfo>"
	End If
	rsget.Close
	'--------------------------------  쿼리부분 끝 --------------------------------
	strRst3 = ""
	strRst3 = strRst3 & "		</AddOfficialInfo>"
	strRst3 = strRst3 & "	</soap:Body>"
	strRst3 = strRst3 & "</soap:Envelope>"
	getGmarketInfoCdParameter = strRst1 & strRst2 & strRst3
End Function

Public Function fnCertCodes(iitemid, iGubun, icertNo, icertDiv, icertDate, imodelName)
	Dim strSql, addSql
	If iGubun = "ELEC" Then
		addSql = addSql & " and r.safetyDiv in ('10', '20', '30') "
	ElseIf iGubun = "LIFE" Then
		addSql = addSql & " and r.safetyDiv in ('40', '50', '60') "
	Else
		addSql = addSql & " and r.safetyDiv in ('70', '80', '90') "
	End If

	strSql = ""
	strSql = strSql & " SELECT TOP 1 r.certNum "
	strSql = strSql & "	,Case When r.safetyDiv in ('10', '40', '70') THEN 'SafeCert' "
	strSql = strSql & "		  When r.safetyDiv in ('20', '50', '80') THEN 'SafeCheck' "
	strSql = strSql & " 	  When r.safetyDiv in ('30', '60', '90') THEN 'SupplierCheck' end as safetyStr "
	strSql = strSql & " ,convert(date, f.certDate) as certDate, f.modelName " & vbcrlf
	strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg as r " & vbcrlf
	strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on r.itemid = f.itemid " & vbcrlf
	strSql = strSql & " WHERE r.itemid='"&iitemid&"' "
	strSql = strSql & addSql
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) then
		icertNo		= rsget("certNum")
		icertDiv	= rsget("safetyStr")
		icertDate	= rsget("certDate")
		imodelName	= rsget("modelName")
	End If
	rsget.Close
End Function

Public Function getGmarketReturnFeeParameter(iitemid, iGmarketPrdno, iReturnFee)
	Dim strRst, buf
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst & "	<soap:Header>"
	strRst = strRst & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "			<encTicket>"&gmarketTicket&"</encTicket>"
	strRst = strRst & "		</EncTicket>"
	strRst = strRst & "	</soap:Header>"
	strRst = strRst & "	<soap:Body>"
	strRst = strRst & "		<AddItemReturnFee xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "			<AddItemReturnFee GmktItemNo="""&iGmarketPrdno&""" ReturnFeeType=""Item"">"		'ReturnFeeType 반품배송비 타입| SellerBasic : 판매자기본 Item :상품별
	strRst = strRst & "				<ItemReturnFee "
	strRst = strRst & "					ReturnChargeType=""ByBuyer"""	'반품배송비과금유형 | 반품배송비를 상품별인 경우 필수 BySeller : 판매자부담-무료 ByBuyer : 구매자부담-유료
	strRst = strRst & "					ReturnShippingFee="""&iReturnFee&""""		'반품배송비(편도) | 상품별/구매자부담 인 경우 필수 (0 원 초과 10 만원이하)
	strRst = strRst & "					ExchangeShippingFee="""&iReturnFee&""" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />" '교환배송비(편도) | 상품별인 경우 필수 (0 원 초과 10 만원이하)
	strRst = strRst & "			</AddItemReturnFee>"
	strRst = strRst & "		</AddItemReturnFee>"
	strRst = strRst & "	</soap:Body>"
	strRst = strRst & "</soap:Envelope>"
	getGmarketReturnFeeParameter = strRst
End Function

'어린이 인증 Soap XML
Public Function getGmarketChildrenParameter(iitemid, iGmarketPrdno, isChildrenCate, isLifeCate, isElecCate)
	Dim strRst, certNo, certDiv, certDate, modelName, buf
	buf = ""
	If isElecCate = "Y" then
		Call fnCertCodes(iitemid, "ELEC", certNo, certDiv, certDate, modelName)
		If certNo <> "" Then
			buf = buf & "			<SafeCertGroupList SafeCertGroupType=""Electric"" CertificationType=""RequireCert"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"">"
			buf = buf & "				<SafeCertInfoList CertificationNo="""&certNo&""" CertificationTargetCode="""&certDiv&""" CertificationStatus=""적합"" CertificationDate="""&certDate&""" ModelName="""&modelName&""" />"
			buf = buf & "			</SafeCertGroupList>"
		Else
			buf = buf & "			<SafeCertGroupList SafeCertGroupType=""Electric"" CertificationType=""AddDescription"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"
		End If
	End If

	If isLifeCate = "Y" then
		Call fnCertCodes(iitemid, "LIFE", certNo, certDiv, certDate, modelName)
		If certNo <> "" Then
			buf = buf & "			<SafeCertGroupList SafeCertGroupType=""Life"" CertificationType=""RequireCert"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"">"
			buf = buf & "				<SafeCertInfoList CertificationNo="""&certNo&""" CertificationTargetCode="""&certDiv&""" CertificationStatus=""적합"" CertificationDate="""&certDate&""" ModelName="""&modelName&""" />"
			buf = buf & "			</SafeCertGroupList>"
		Else
			buf = buf & "			<SafeCertGroupList SafeCertGroupType=""Life"" CertificationType=""AddDescription"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"
		End If
	End If

	If isChildrenCate = "Y" then
		Call fnCertCodes(iitemid, "CHILD", certNo, certDiv, certDate, modelName)
		If certNo <> "" Then
			buf = buf & "			<SafeCertGroupList SafeCertGroupType=""Child"" CertificationType=""RequireCert"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"">"
			buf = buf & "				<SafeCertInfoList CertificationNo="""&certNo&""" CertificationTargetCode="""&certDiv&""" CertificationStatus=""적합"" CertificationDate="""&certDate&""" ModelName="""&modelName&""" />"
			buf = buf & "			</SafeCertGroupList>"
		Else
			buf = buf & "			<SafeCertGroupList SafeCertGroupType=""Child"" CertificationType=""AddDescription"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"
		End If
	End If

	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst & "	<soap:Header>"
	strRst = strRst & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "			<encTicket>"&gmarketTicket&"</encTicket>"
	strRst = strRst & "		</EncTicket>"
	strRst = strRst & "	</soap:Header>"
	strRst = strRst & "	<soap:Body>"
	strRst = strRst & "		<AddIntegrateSafeCert xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "			<AddIntegrateSafeCert GmktItemNo="""&iGmarketPrdno&""">"
	strRst = strRst & buf
	strRst = strRst & "			</AddIntegrateSafeCert>"
	strRst = strRst & "		</AddIntegrateSafeCert>"
	strRst = strRst & "	</soap:Body>"
	strRst = strRst & "</soap:Envelope>"
	getGmarketChildrenParameter = strRst
End Function

'가격/재고 Soap XML
Public Function getGmarketAddPriceParameter(iitemid, iGmarketPrdno, isReged, byref iSellprice, byref iDisplayDate)	''뭔가 XML수정이 필요하다면..incGmarketFunction의 getGmarketAddPriceParameter도 같이 수정
	Dim strSQL, strRst, GetTenTenMargin, ownItemCnt, outmallstandardMargin
	Dim ibuycash, iorgprice, isellcash, ilimityn, ilimitno, ilimitsold, iGmarketFirstPrice, iGmarketPrice, iexpDate, iStockQty, ispecialPrice

	strSQL = ""
	strSQL = strSQL & " SELECT TOP 1 i.itemid, i.buycash, i.orgprice, i.sellcash, i.limityn, i.limitno, i.limitsold "
	strSQL = strSQL & " , isnull(r.GmarketFirstPrice, '0') as GmarketFirstPrice, isnull(r.GmarketPrice, '0') as GmarketPrice, isnull(r.DisplayDate, '') as DisplayDate "
	strSQL = strSQL & " , isnull(mi.mustPrice, 0) as specialPrice, isnull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
	strSQL = strSQL & " FROM db_item.dbo.tbl_item as i "
	strSQL = strSQL & " JOIN db_etcmall.[dbo].[tbl_gmarket_regItem] as r on i.itemid = r.itemid "
	strSQL = strSQL & " LEFT JOIN db_etcmall.[dbo].[tbl_outmall_mustPriceItem] as mi "
	strSQL = strSQL & " 	on i.itemid = mi.itemid "
	strSQL = strSQL & " 	and mi.mallgubun = 'gmarket1010' "
	strSQL = strSQL & " 	and (GETDATE() >= mi.startDate and GETDATE() <= mi.endDate ) "
	strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
	strSQL = strSQL & " WHERE i.itemid = '"&iitemid&"'"
	rsget.CursorLocation = adUseClient
	rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		ibuycash				= rsget("buycash")
		iorgprice				= rsget("orgprice")
		isellcash				= rsget("sellcash")
		ilimityn				= rsget("limityn")
		ilimitno				= rsget("limitno")
		ilimitsold				= rsget("limitsold")
		iGmarketFirstPrice		= rsget("GmarketFirstPrice")
		iGmarketPrice			= rsget("GmarketPrice")
		iexpDate				= rsget("displayDate")
		ispecialPrice			= rsget("specialPrice")
		outmallstandardMargin = rsget("outmallstandardMargin")
	End If
	rsget.Close

	strSQL = ""
	strSQL = strSQL & " SELECT COUNT(*) as CNT "
	strSQL = strSQL & " FROM db_item.dbo.tbl_item i "
	strSQL = strSQL & " JOIN db_partner.dbo.tbl_partner p on i.makerid = p.id "
	strSQL = strSQL & " WHERE p.purchaseType in (3, 5, 6) "		'3 : PB, 5 : ODM, 6 : 수입
	strSQL = strSQL & " and i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		ownItemCnt = rsget("CNT")
	End If
	rsget.Close

	'판매가 구하기
	If ispecialPrice <> "0" Then
		iSellprice = ispecialPrice
	ElseIf ownItemCnt > 0 Then
		iSellprice = iorgprice
	Else
		GetTenTenMargin = CLng(10000 - ibuycash / isellcash * 100 * 100) / 100
		If GetTenTenMargin < outmallstandardMargin Then
			iSellprice = iorgprice
		Else
			iSellprice = isellcash
		End If
	End If

	'노출 기간 설정
	If iexpDate = "" Then
		idisplayDate = DateAdd("yyyy", 1, Date())
	Else
		If DateDiff("m", iexpDate, Date()) <= 3 Then
			idisplayDate = DateAdd("yyyy", 1, Date())
		Else
			'idisplayDate = iexpDate
			idisplayDate = DateAdd("d", 1, Date())
		End If
	End If

	'재고 수량 설정
	If isReged = "N" Then
		iStockQty = 0
	Else
		If ilimityn = "Y" Then
			iStockQty = ilimitno - ilimitsold - 5
			If iStockQty > 1000 Then
				iStockQty = CDEFALUT_STOCK
			End If
		Else
			iStockQty = CDEFALUT_STOCK
		End If
		If (iStockQty < 1) Then iStockQty = 0
	End If

	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst & "	<soap:Header>"
	strRst = strRst & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "			<encTicket>"&gmarketTicket&"</encTicket>"
	strRst = strRst & "		</EncTicket>"
	strRst = strRst & "	</soap:Header>"
	strRst = strRst & "	<soap:Body>"
	strRst = strRst & "		<AddPrice xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "			<AddPrice "
	strRst = strRst & "				GmktItemNo="""&iGmarketPrdno&""""			'#G마켓 상품번호
	strRst = strRst & "				DisplayDate="""&idisplayDate&""""			'#주문기간 | 최대 1년
	strRst = strRst & "				SellPrice="""&iSellprice&""""				'#판매가격 | 최소 100원 이상 1,000,000,000원 미만 (100원 단위)
	strRst = strRst & "				StockQty="""&iStockQty&""""					'#재고수량
	strRst = strRst & "				InventoryNo="""&iitemid&""" />"				'판매자관리코드
	strRst = strRst & "		</AddPrice>"
	strRst = strRst & "	</soap:Body>"
	strRst = strRst & "</soap:Envelope>"
	getGmarketAddPriceParameter = strRst
	' If (session("ssBctID")="kjy8517") Then
	' 	rw Replace(objXML.responseText,"soap:","")
	' End If
End Function

'브랜드 등록/조회 Soap
Public Function getGmarketAddMakerBrandParameter(igMakername, igBrandname)
	Dim strRst
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst & "  <soap:Header>"
	strRst = strRst & "    <EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "      <encTicket>"&gmarketTicket&"</encTicket>"
	strRst = strRst & "    </EncTicket>"
	strRst = strRst & "  </soap:Header>"
	strRst = strRst & "  <soap:Body>"
	strRst = strRst & "    <AddMakerBrand xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "      <AddMakerBrand MakerName="""&igMakername&""" BrandName="""&igBrandname&""" />"
	strRst = strRst & "    </AddMakerBrand>"
	strRst = strRst & "  </soap:Body>"
	strRst = strRst & "</soap:Envelope>"
	getGmarketAddMakerBrandParameter = strRst
End Function

'판매자 주소(반품/교환) 배송지 등록 Soap
Public Function getGmarketAddAddressBookParameter()
	Dim strRst, sqlStr
	Dim AddressTitle, AddressName, Phone1, Phone2, reqzipcode, reqzipaddr, reqaddress, AddressCode, BundleNo
	sqlStr = ""
	sqlStr = sqlStr & " SELECT TOP 1 * FROM db_etcmall.[dbo].[tbl_gmarket_AddressBook] "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		AddressCode		= Trim(rsget("AddressCode"))
		AddressTitle	= Trim(rsget("AddressTitle"))
		AddressName		= Trim(rsget("AddressName"))
		Phone1			= Trim(rsget("Phone1"))
		Phone2			= Trim(rsget("Phone2"))
		reqzipcode		= Trim(rsget("reqzipcode"))
		reqzipaddr		= Trim(rsget("reqzipaddr"))
		reqaddress		= Trim(rsget("reqaddress"))
		BundleNo		= Trim(rsget("BundleNo"))
	End If
	rsget.Close

	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
  	strRst = strRst & "	<soap:Header>"
	strRst = strRst & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "			<encTicket>"&gmarketTicket&"</encTicket>"
	strRst = strRst & "		</EncTicket>"
	strRst = strRst & "	</soap:Header>"
	strRst = strRst & "	<soap:Body>"
	strRst = strRst & "		<AddAddressBook xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "			<AddAddressBook AddressKind=""Add"">"'Add or Change or Remove
	strRst = strRst & "				<AddressBook "
	If AddressCode <> "" Then
		strRst = strRst & "				AddressCode="""&AddressCode&""""
	End If
	strRst = strRst & "					AddressTitle="""&AddressTitle&""""
	strRst = strRst & "					Name="""&AddressName&""""
	strRst = strRst & "					Phone1="""&Phone1&""""
	strRst = strRst & "					Phone2="""&Phone2&""""
	strRst = strRst & "					Zipcode="""&reqzipcode&""""
	strRst = strRst & "					Address1="""&reqzipaddr&""""
	strRst = strRst & "					Address2="""&reqaddress&""""
	strRst = strRst & "					xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"
	strRst = strRst & "			</AddAddressBook>"
	strRst = strRst & "		</AddAddressBook>"
	strRst = strRst & "	</soap:Body>"
	strRst = strRst & "</soap:Envelope>"
	getGmarketAddAddressBookParameter = strRst
End Function

'판매자 주소(반품/교환) 배송지 조회 RequestAddressBook Soap
Public Function getGmarketRequestAddressBookParameter()
	Dim strRst, sqlStr
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
  	strRst = strRst & "	<soap:Header>"
    strRst = strRst & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
    strRst = strRst & "			<encTicket>"&gmarketTicket&"</encTicket>"
    strRst = strRst & "		</EncTicket>"
  	strRst = strRst & "	</soap:Header>"
  	strRst = strRst & "	<soap:Body>"
    strRst = strRst & "		<RequestAddressBook xmlns=""http://tpl.gmarket.co.kr/"">"
    strRst = strRst & "			<AddressBook AddressCode="""" AddressTitle="""" Name="""" Address1="""" Address2="""" BundleNo="""" />"
    strRst = strRst & "		</RequestAddressBook>"
  	strRst = strRst & "	</soap:Body>"
	strRst = strRst & "</soap:Envelope>"
	getGmarketRequestAddressBookParameter = strRst
End Function
'################################################# 각 기능 별 파라메터 정리 끝 ###############################################
%>