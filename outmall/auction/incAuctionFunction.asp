<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'Additem 상품 기본정보 등록
Public Function fnAuctionItemReg(iitemid, strParam, byRef iErrStr, iSellCash, iAuctionSellYn, ilimityn, ilimitno, ilimiysold, iitemname, iimageNm)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & auctionAPIURL&"/APIv1/ShoppingService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/ShoppingService/AddItem"
		objXML.send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[AddItem] " & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))

			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "AddItemResponse" Then
				goodsCd = xmlDOM.getElementsByTagName ("AddItemResult ").item(0).attributes(0).nodeValue
				If goodsCd <> "" Then
					strSql = ""
					strSql = strSql & " UPDATE R"
					strSql = strSql & "	Set AuctionLastUpdate = getdate() "
					strSql = strSql & "	, AuctionGoodNo = '" & goodsCd & "'"
					strSql = strSql & "	, AuctionPrice = " &iSellCash
					strSql = strSql & "	, regImageName = '"&iimageNm&"' "
					strSql = strSql & "	, accFailCnt = 0"
					strSql = strSql & "	, AuctionRegdate = isNULL(AuctionRegdate, getdate())"
				    strSql = strSql & "	, AuctionStatCd=(CASE WHEN isNULL(AuctionStatCd, -1) < 3 then 3 ELSE AuctionStatCd END)"
				    strSql = strSql & "	, APIadditem = 'Y'"
					strSql = strSql & "	From db_etcmall.dbo.tbl_auction_regItem  R"
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||[AddItem]성공"
				Else
					iErrStr = "ERR||"&iitemid&"||[AddItem] "& objXML.responseText
					If (session("ssBctID")="kjy8517") Then
						rw "REQ : <textarea cols=40 rows=10>"&strParam&"</textarea>"
						rw "RES : <textarea cols=40 rows=10>"&objXML.responseText&"</textarea>"
					End If
				End If
			ElseIF xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "Fault" Then
				iMessage = xmlDOM.selectSingleNode("Envelope/Body").firstChild.firstChild.nextSibling.text
				iErrStr = "ERR||"&iitemid&"||[AddItem] "&iMessage
				If (session("ssBctID")="kjy8517") Then
					rw "REQ : <textarea cols=40 rows=10>"&strParam&"</textarea>"
					rw "RES : <textarea cols=40 rows=10>"&objXML.responseText&"</textarea>"
				End If
			Else
				iErrStr = "ERR||"&iitemid&"||[AddItem] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'상품 옵션정보 등록
Public Function fnAuctionOPTReg(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, iMessage, goodsCd
'response.write strParam
'response.end
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & auctionAPIURL&"/APIv1/ShoppingService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/ShoppingService/ReviseItemStock"
		objXML.send(strParam)

'		response.write objXML.responseText
'		response.end
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[AddOPT] " & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "ReviseItemStockResponse" Then
				goodsCd = xmlDOM.getElementsByTagName ("ItemStock ").item(0).attributes(0).nodeValue
				If goodsCd <> "" Then
					strSql = ""
					strSql = strSql & " UPDATE R"
					strSql = strSql & "	Set APIaddopt = 'Y'"
					strSql = strSql & "	From db_etcmall.dbo.tbl_auction_regItem  R"
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||[AddOPT]성공"
				End If
			ElseIF xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "Fault" Then
				iMessage = xmlDOM.selectSingleNode("Envelope/Body").firstChild.firstChild.nextSibling.text
				iErrStr = "ERR||"&iitemid&"||[AddOPT] "&iMessage
			Else
				iErrStr = "ERR||"&iitemid&"||[AddOPT] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
'response.end
End Function

'AddOfficialNotice 상품 정보고시항목 등록
Public Function fnAuctionItemInfoCd(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, iMessage, goodsCd
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & auctionAPIURL&"/APIv1/ShoppingService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/ShoppingService/AddOfficialNotice"
		objXML.send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[AddInfoCd] " & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "AddOfficialNoticeResponse" Then
				goodsCd = xmlDOM.getElementsByTagName ("AddOfficialNoticeResult ").item(0).attributes(0).nodeValue
				If goodsCd <> "" Then
					strSql = ""
					strSql = strSql & " UPDATE R"
					strSql = strSql & "	Set APIaddgosi = 'Y'"
					strSql = strSql & "	From db_etcmall.dbo.tbl_auction_regItem  R"
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||[AddInfoCd]성공"
				End If
			ElseIF xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "Fault" Then
				iMessage = xmlDOM.selectSingleNode("Envelope/Body").firstChild.firstChild.nextSibling.text
				iErrStr = "ERR||"&iitemid&"||[AddInfoCd] "&iMessage
			Else
				iErrStr = "ERR||"&iitemid&"||[AddInfoCd] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'ReviseItemSelling 상품 상태 변경
Public Function fnAuctionSellyn(iitemid, ichgSellyn, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, iMessage, goodsCd
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & auctionAPIURL&"/APIv1/ShoppingService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/ShoppingService/ReviseItemSelling"
		objXML.send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||" & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))

			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "ReviseItemSellingResponse" Then
				goodsCd = xmlDOM.getElementsByTagName ("ReviseItemSellingResult").item(0).attributes(0).nodeValue
				If goodsCd <> "" Then
					If ichgSellyn = "Y" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	Set OnSaleRegdate = getdate()"
						strSql = strSql & "	,AuctionStatCd = 7"
						strSql = strSql & "	,AuctionSellYn = 'Y'"
						strSql = strSql & "	,AuctionLastUpdate = getdate()"
						strSql = strSql & "	From db_etcmall.dbo.tbl_auction_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||판매중으로 변경"
					Else
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	Set AuctionSellYn = 'N'"
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,AuctionLastUpdate = getdate()"
						strSql = strSql & "	From db_etcmall.dbo.tbl_auction_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||품절처리"
					End If
				End If
			ElseIF xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "Fault" Then
				iMessage = xmlDOM.selectSingleNode("Envelope/Body").firstChild.firstChild.nextSibling.text
				iErrStr = "ERR||"&iitemid&"|| "&iMessage
			Else
				iErrStr = "ERR||"&iitemid&"|| 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'기본정보 수정
Public Function fnAuctionIteminfoEdit(iitemid, iauctiongoodno, byRef iErrStr, strParam, iSellcash)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & auctionAPIURL&"/APIv1/ShoppingService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/ShoppingService/ReviseItem"
		objXML.send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[EditInfo]" & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
'response.write Replace(objXML.responseText,"soap:","")
'response.end
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "ReviseItemResponse" Then
				goodsCd = xmlDOM.getElementsByTagName ("ReviseItemResult").item(0).attributes(0).nodeValue
				If goodsCd <> "" Then
					strSql = ""
					strSql = strSql & " UPDATE R"
					strSql = strSql & "	Set AuctionLastUpdate = getdate() "
					strSql = strSql & "	, AuctionPrice = " &iSellCash
					strSql = strSql & "	, returnShippingFee = 3000 "
					strSql = strSql & " , regitemname = '"&html2db(oAuction.FOneItem.FItemname)&"' " & VbCRLF
					If oAuction.FOneItem.isImageChanged Then
						strSql = strSql & " ,regImageName = '"&oAuction.FOneItem.getBasicImage&"' " & VbCRLF
					End If
					strSql = strSql & "	, accFailCnt = 0"
					strSql = strSql & "	From db_etcmall.dbo.tbl_auction_regItem  R"
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||[EditInfo]성공"
				End If
			ElseIF xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "Fault" Then
				iMessage = xmlDOM.selectSingleNode("Envelope/Body").firstChild.firstChild.nextSibling.text
				iMessage = replace(iMessage, "'", "")
				iErrStr = "ERR||"&iitemid&"||[EditInfo] "&db2html(iMessage)
			Else
				iErrStr = "ERR||"&iitemid&"||[EditInfo] 정의되지 않은 오류"
			End IF

			If iErrStr = "" Then
				iErrStr = "ERR||"&iitemid&"||[EditInfo] 정의되지 않은 오류(2)"
			End If
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'상품 옵션 조회
Public Function fnAuctionOPTSTAT(iitemid, strParam, iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, aoptType, aOptcd, aOptnm, aOptprice, aTenOptcd, aOptSellyn, aOptQty
	Dim optlist, SubNodes, isMulti
	Dim para2El, isDanPoom, multicnt, AssignedRow, actCnt
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & auctionAPIURL&"/APIv1/ShoppingService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/ShoppingService/ViewItemStock"
		objXML.send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[OPTSTAT]" & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
'response.write (Replace(objXML.responseText,"soap:",""))
'response.end
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "ViewItemStockResponse" Then
				goodsCd		= xmlDOM.getElementsByTagName ("ItemStock ").item(0).attributes(0).nodeValue	'옥션상품코드
				aoptType	= xmlDOM.getElementsByTagName ("ItemStock ").item(0).attributes(1).nodeValue	'옥션옵션타입
				Select Case aoptType
					Case "BuyerDescriptive"			'단품이면서 주문문구 있음					'테스트완료
						isDanPoom = "Y"
					Case "NotAvailable"				'단품이면서 주문문구 없음					'테스트완료
						isDanPoom = "Y"
					Case "StandAloneMixed"			'일반옵션 이면서 주문문구 있음				'테스트완료
						isDanPoom = "N"
					Case "Mixed"					'복합(2)옵션 이면서 주문문구 있음			'테스트완료
						isDanPoom = "N"
						multicnt = 2
					Case "ThreeCombinationMixed"	'복합(3)옵션 이면서 주문문구 있음			'테스트완료
						isDanPoom = "N"
						multicnt = 3
					Case "StandAlone"				'일반옵션 이면서 주문문구 없음				'테스트완료
						isDanPoom = "N"
					Case "BuyerSelective"			'복합(2)옵션 이면서 주문문구 없음			'테스트완료
						isDanPoom = "N"
						multicnt = 2
					Case "ThreeCombination"			'복합(3)옵션 이면서 주문문구 없음			'테스트완료
						isDanPoom = "N"
						multicnt = 3
				End Select

				If goodsCd <> "" Then
					Set para2El = xmlDOM.selectSingleNode("//StockStandAlone")
					If para2El Is Nothing Then
						isMulti = "Y"
						Set optlist = xmlDOM.getElementsByTagName("OrderStock")
					Else
						isMulti = "N"
						Set optlist = xmlDOM.getElementsByTagName("StockStandAlone")
					End If

					For Each SubNodes in optlist
						If (SubNodes.nodeType = 1 Or SubNodes.nodeType = 2) Then
							If isMulti = "Y" Then
								aOptcd		= SubNodes.attributes.GetNamedItem("StockNo").value
								If multicnt = 2 Then
									aOptnm = SubNodes.attributes.GetNamedItem("Section").value&","&SubNodes.attributes.GetNamedItem("Text").value
								ElseIf multicnt = 3 Then
									aOptnm = SubNodes.attributes.GetNamedItem("Section").value&","&SubNodes.attributes.GetNamedItem("Text").value&","&SubNodes.attributes.GetNamedItem("Text2").value
								End If

								If isDanPoom = "N" Then
									aTenOptcd	= SubNodes.attributes.GetNamedItem("Code").value
								Else
									aTenOptcd	= "0000"
								End If
								aOptprice	= SubNodes.attributes.GetNamedItem("Price").value
								aOptSellyn	= SubNodes.attributes.GetNamedItem("IsDisplayable").value
								aOptQty		= SubNodes.attributes.GetNamedItem("Quantity").value
							Else
								aOptcd		= SubNodes.attributes.GetNamedItem("ItemStockStandAloneNo").value
								aOptnm		= SubNodes.attributes.GetNamedItem("Text").value
								aTenOptcd	= SubNodes.attributes.GetNamedItem("SellerStockCode").value
								aOptprice	= SubNodes.attributes.GetNamedItem("Price").value
								aOptSellyn	= SubNodes.attributes.GetNamedItem("IsSoldOut").value
								If aOptSellyn = "false" Then
									aOptSellyn = "true"
								Else
									aOptSellyn = "false"
								End If
								aOptQty		= SubNodes.attributes.GetNamedItem("StockQty").value
							End If

'rw "옥션옵션코드 : " & aOptcd
'rw "옥션명 : " & aOptnm
'rw "옥션옵션추가금액 : " & aOptprice
'rw "텐텐옵션코드 : " & aTenOptcd
'rw "옥션옵션판매상태 : " & aOptSellyn
'rw "옥션옵션재고 : " & aOptQty
'rw "-------------"
							strSql = ""
							strSql = strSql & " IF Exists(SELECT * FROM db_item.dbo.tbl_OutMall_regedoption WHERE itemid="&iitemid&" and mallid = '"&CMALLNAME&"' and itemoption = '"&aTenOptcd&"' )"
							strSql = strSql & " BEGIN "
							strSql = strSql & " UPDATE db_item.dbo.tbl_OutMall_regedoption SET "
							strSql = strSql & " outmallsellyn='"&CHKIIF(aOptSellyn="false","N","Y")&"'"
							If (aOptcd <> "0000") Then
							    strSql = strSql & " , outmallOptName='"&html2DB(aOptnm)&"'"
							Else
								strSql = strSql & " , outmallOptName='단일상품'"
							End If
							strSql = strSql & " , outmallAddPrice="&aOptprice
							strSql = strSql & " , checkdate = getdate() "
							strSql = strSql & " , outmallOptCode='"&aOptcd&"'"
							strSql = strSql & " WHERE itemid = '"&iitemid&"' and itemoption = '"&aTenOptcd&"' "
							strSql = strSql & " and mallid='"&CMALLNAME&"'"
							strSql = strSql & " END ELSE "
							strSql = strSql & " BEGIN "
							If aTenOptcd = "0000" Then
								strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
								strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastUpdate, checkdate) "
								strSql = strSql & " VALUES "
								strSql = strSql & " ('"&iitemid&"'"
								strSql = strSql & ",  '"&aTenOptcd&"'"
								strSql = strSql & ", '"&CMALLNAME&"'"
								strSql = strSql & ", '"&aOptcd&"'"
								strSql = strSql & ", '단일상품'"
								strSql = strSql & ", '"&CHKIIF(aOptSellyn="false", "N", "Y")&"'"
								strSql = strSql & ", 'Y'"
								strSql = strSql & ", '"&aOptQty&"'"
								strSql = strSql & ", getdate() "
								strSql = strSql & ", getdate()) "
							Else
								strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
								strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastUpdate, checkdate) " & VBCRLF
								strSql = strSql & " SELECT itemid, itemoption, '"& CMALLNAME &"', '"& aOptcd &"', optionname, '"&CHKIIF(aOptSellyn="false", "N", "Y")&"', 'Y', '"&aOptQty&"', getdate(), getdate() "
								strSql = strSql & " FROM db_item.dbo.tbl_item_option " & VBCRLF
								strSql = strSql & " WHERE itemid= '"&iitemid&"' and itemoption = '"& aTenOptcd &"' " & VBCRLF
							End If
							strSql = strSql & " END "
							dbget.Execute strSql, AssignedRow
							actCnt = actCnt+AssignedRow
						End If
					Next

					If (actCnt > 0) Then
						strSql = " update R"   &VbCRLF
						strSql = strSql & " set regedOptCnt=isNULL(T.regedOptCnt,0)"   &VbCRLF
						strSql = strSql & " from db_etcmall.dbo.tbl_auction_regItem R"   &VbCRLF
						strSql = strSql & " 	Join ("   &VbCRLF
						strSql = strSql & " 		select R.itemid,count(*) as CNT "
						strSql = strSql & " 		, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
						strSql = strSql & "        from db_etcmall.dbo.tbl_auction_regItem R"   &VbCRLF
						strSql = strSql & " 			Join db_item.dbo.tbl_OutMall_regedoption Ro"   &VbCRLF
						strSql = strSql & " 			on R.itemid=Ro.itemid"   &VbCRLF
						strSql = strSql & " 			and Ro.mallid='"&CMALLNAME&"'"   &VbCRLF
						strSql = strSql & "             and Ro.itemid="&iitemid&VbCRLF
						strSql = strSql & " 		group by R.itemid"   &VbCRLF
						strSql = strSql & " 	) T on R.itemid=T.itemid"   &VbCRLF
						dbget.Execute strSql
					End If
					Set para2El = nothing
					iErrStr =  "OK||"&iitemid&"||[OPTSTAT]성공"
				End If
			ElseIF xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "Fault" Then
				iMessage = xmlDOM.selectSingleNode("Envelope/Body").firstChild.firstChild.nextSibling.text
				iErrStr = "ERR||"&iitemid&"||[OPTSTAT] "&iMessage
			Else
				iErrStr = "ERR||"&iitemid&"||[OPTSTAT] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'상품 옵션정보 수정
Public Function fnAuctionOPTEDT(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, iMessage, goodsCd
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & auctionAPIURL&"/APIv1/ShoppingService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/ShoppingService/ReviseItemStock"
		objXML.send(strParam)


		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[EDTOPT] " & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "ReviseItemStockResponse" Then
				goodsCd = xmlDOM.getElementsByTagName ("ItemStock ").item(0).attributes(0).nodeValue
				If goodsCd <> "" Then
					iErrStr =  "OK||"&iitemid&"||[EDTOPT]성공"
				End If
			ElseIF xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "Fault" Then
				iMessage = xmlDOM.selectSingleNode("Envelope/Body").firstChild.firstChild.nextSibling.text
				iErrStr = "ERR||"&iitemid&"||[EDTOPT] "&iMessage
			Else
				iErrStr = "ERR||"&iitemid&"||[EDTOPT] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'상품 옵션정보 삭제
Public Function fnAuctionOPTDel(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, iMessage, goodsCd
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & auctionAPIURL&"/APIv1/ShoppingService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strParam)
		objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/ShoppingService/ReviseItemStock"
		objXML.send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[DelOPT] " & Err.Description
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "ReviseItemStockResponse" Then
				goodsCd = xmlDOM.getElementsByTagName ("ItemStock ").item(0).attributes(0).nodeValue
				If goodsCd <> "" Then
					strSql = ""
					strSql = " DELETE FROM db_item.dbo.tbl_outmall_regedoption WHERE mallid = '"&CMALLNAME&"' and itemid = '"&iitemid&"' "
					dbget.Execute(strSql)

					strSql = ""
					strSql = "UPDATE db_etcmall.dbo.tbl_auction_regitem SET regedoptcnt = null WHERE itemid = '"&iitemid&"'"
					dbget.Execute(strSql)

					iErrStr =  "OK||"&iitemid&"||[DelOPT]성공"
				End If
			ElseIF xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "Fault" Then
				iMessage = xmlDOM.selectSingleNode("Envelope/Body").firstChild.firstChild.nextSibling.text
				iErrStr = "ERR||"&iitemid&"||[DelOPT] "&iMessage
			Else
				iErrStr = "ERR||"&iitemid&"||[DelOPT] 정의되지 않은 오류"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'공통코드 확인 및 등록
Public Function fnAuctionCommonCode(iccd, strParam)
	Dim objXML, xmlDOM, SubNodes, strSql
	Dim retCode, iMessage, optlist
	Dim AssignedRow, attr, iOriginCode, iOriginName, iOriginNameDetail

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If iccd = "GetDeliveryList" Then
			objXML.open "POST", "" & auctionAPIURL&"/APIv1/AuctionService.asmx"
			objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
			objXML.setRequestHeader "Content-Length", LenB(strParam)
			objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/AuctionService/"&iccd
		Else
			objXML.open "POST", "" & auctionAPIURL&"/APIv1/ShoppingService.asmx"
			objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
			objXML.setRequestHeader "Content-Length", LenB(strParam)
			objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/ShoppingService/"&iccd
		End If
		objXML.send(strParam)

		If iccd = "GetNationCode" Then
			If objXML.Status = "200" Then
				Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
					xmlDOM.async = False
					xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
					strSql = ""
					strSql = "delete from db_etcmall.dbo.tbl_auction_Nation"
					dbget.Execute(strSql)

					IF xmlDOM.selectSingleNode("Envelope/Body/GetNationCodeResponse/GetNationCodeResult").firstChild.nodeName = "NationListT" Then
						Set optlist = xmlDOM.getElementsByTagName("NationListT")
						For Each SubNodes in optlist
							If (SubNodes.nodeType = 1 Or SubNodes.nodeType = 2) Then
								iOriginCode = SubNodes.attributes.GetNamedItem("OriginCode").value
								iOriginName = SubNodes.attributes.GetNamedItem("OriginName").value
								iOriginNameDetail = SubNodes.attributes.GetNamedItem("OriginNameDetail").value
								strSql = ""
								strSql = strSql & "insert into db_etcmall.dbo.tbl_auction_Nation(code, nationname, nationnameDetail) values ('"&iOriginCode&"', '"&iOriginName&"', '"&iOriginNameDetail&"')"
								dbget.Execute(strSql)
							End If
						Next
						Set optlist = nothing
					End If
				Set xmlDOM = nothing
			End If
		ElseIf iccd = "GetShippingPlaceCode" OR iccd = "GetDeliveryList" Then
			response.write replace(objXML.responseText, "utf-8","euc-kr")
			response.end
		ElseIf iccd = "GetDeliveryPrepareList" Then
			Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
				objXML.open "POST", "" & auctionAPIURL&"/APIv1/AuctionService.asmx"
				objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
				objXML.setRequestHeader "Content-Length", LenB(strParam)
				objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/AuctionService/"&iccd
				objXML.send(strParam)

			response.write replace(objXML.responseText, "utf-8","euc-kr")
			response.end
		End If
	Set objXML = nothing
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################

'################################################# 각 기능 별 파라메터 정리시작 ###############################################
'정보고시 Soap XML
Public Function getAuctionInfoCdParameter(iitemid, iauctionPrdno)
	Dim strSQL, strRst1, strRst2, strRst3
	Dim mallinfodiv, mallinfoCd, infoContent
	strRst1 = ""
	strRst1 = strRst1 & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst1 = strRst1 & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst1 = strRst1 & "	<soap:Header>"
	strRst1 = strRst1 & "		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
	strRst1 = strRst1 & "			<Value>"&auctionTicket&"</Value>"
	strRst1 = strRst1 & "		</EncryptedTicket>"
	strRst1 = strRst1 & "	</soap:Header>"
	strRst1 = strRst1 & "	<soap:Body>"
	strRst1 = strRst1 & "		<AddOfficialNotice xmlns=""http://www.auction.co.kr/APIv1/ShoppingService"">"
	strRst1 = strRst1 & "			<req Version=""1"">"
	strRst1 = strRst1 & "				<ItemOfficialNotice xmlns=""http://schema.auction.co.kr/Arche.Sell3.Service.xsd"">"
	strRst1 = strRst1 & "					<ItemID xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"">"&iauctionPrdno&"</ItemID>"
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
	strSQL = strSQL & " WHERE M.mallid = 'auction' and IC.itemid='"&iitemid&"'"
	rsget.CursorLocation = adUseClient
	rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) then
		mallinfodiv = CInt(rsget("mallinfodiv"))
		If mallinfodiv = "47" Then
			mallinfodiv = "40"
		ElseIf mallinfodiv = "48" Then
			mallinfodiv = "41"
		End If
		strRst2 = "					<NotiItemGroupNo xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"">"&mallinfodiv&"</NotiItemGroupNo>"
		Do until rsget.EOF
		    mallinfoCd  = rsget("mallinfoCd")
		    infoContent = rsget("infoContent")
			strRst2 = strRst2 & "					<ItemOfficialNotiValue NotiItemCode="""&mallinfoCd&""" NotiItemValue="""&replaceRst(infoContent)&""" ExtraMarkIs=""false"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
			rsget.MoveNext
		Loop
	End If
	rsget.Close
	'--------------------------------  쿼리부분 끝 --------------------------------
	strRst3 = ""
	strRst3 = strRst3 & "				</ItemOfficialNotice>"
	strRst3 = strRst3 & "			</req>"
	strRst3 = strRst3 & "		</AddOfficialNotice>"
	strRst3 = strRst3 & "	</soap:Body>"
	strRst3 = strRst3 & "</soap:Envelope>"
	getAuctionInfoCdParameter = strRst1 & strRst2 & strRst3
'	response.write getAuctionInfoCdParameter
'	response.end
End Function

'전부 등록된 것 판매중 XML
Public Function getAuctionSellYnParameter(ionsaleyn, iitemid, iauctionPrdno)
	Dim strRst, PeriodStatus, strSql, overSellDate
	overSellDate = "N"
	If ionsaleyn = "Y" Then
		PeriodStatus = "OnSale"		'판매 중 | 최대한 길게하기위해 ApplyPeriod는 90으로 설정(90일)
	Else
		PeriodStatus = "Stop"		'판매 중지 | 그외에 Waiting(대기)도 있음
	End If

	strSql = ""
	strSql = strSql & " select count(*) as cnt from db_etcmall.dbo.tbl_auction_regitem where lastErrStr like '%최대 판매기간 초과%' and itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("cnt") > 0 Then
		overSellDate = "Y"
	End If
	rsget.Close

	strRst = ""
	strRst = strRst &"<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst &"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst &"	<soap:Header>"
	strRst = strRst &"		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
	strRst = strRst &"			<Value>"&auctionTicket&"</Value>"
	strRst = strRst &"		</EncryptedTicket>"
	strRst = strRst &"	</soap:Header>"
	strRst = strRst &"	<soap:Body>"
	strRst = strRst &"		<ReviseItemSelling xmlns=""http://www.auction.co.kr/APIv1/ShoppingService"">"
	strRst = strRst &"			<req Version=""1"">"
	strRst = strRst &"				<ItemSelling ItemID="""&iauctionPrdno&""" xmlns=""http://schema.auction.co.kr/Arche.Sell3.Service.xsd"">"
	strRst = strRst &"					<Period Status="""&PeriodStatus&""" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"">"
	If ionsaleyn = "Y" and overSellDate = "N" Then
	strRst = strRst &"					<Period ApplyPeriod=""90"" />"
	End If
	strRst = strRst &"					</Period>"
	strRst = strRst &"				</ItemSelling>"
	strRst = strRst &"			</req>"
	strRst = strRst &"		</ReviseItemSelling>"
	strRst = strRst &"	</soap:Body>"
	strRst = strRst &"</soap:Envelope>"
	getAuctionSellYnParameter = strRst
End Function

'상품 옵션 조회 Soap XML
Public Function getAuctionOptSellModParameter(iauctionPrdno)
	Dim strRst, PeriodStatus
	strRst = ""
	strRst = strRst &"<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst &"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst &"	<soap:Header>"
	strRst = strRst &"		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
	strRst = strRst &"			<Value>"&auctionTicket&"</Value>"
	strRst = strRst &"		</EncryptedTicket>"
	strRst = strRst &"	</soap:Header>"
	strRst = strRst &"	<soap:Body>"
	strRst = strRst &"		<ViewItemStock xmlns=""http://www.auction.co.kr/APIv1/ShoppingService"">"
	strRst = strRst &"			<req ItemID="""&iauctionPrdno&""" Version=""1""></req>"
	strRst = strRst &"		</ViewItemStock>"
	strRst = strRst &"	</soap:Body>"
	strRst = strRst &"</soap:Envelope>"
	getAuctionOptSellModParameter = strRst
End Function

'공통코드 중 원산지 Soap XML
Public Function getAuctionCommonCodeList(iccd)
	Dim strRst
	strRst = ""
	strRst = strRst &"<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst &"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst &"	<soap:Header>"
	strRst = strRst &"		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
	strRst = strRst &"			<Value>"&auctionTicket&"</Value>"
	strRst = strRst &"		</EncryptedTicket>"
	strRst = strRst &"	</soap:Header>"
	strRst = strRst &"  <soap:Body>"
	strRst = strRst &"		<"&iccd&" xmlns=""http://www.auction.co.kr/APIv1/ShoppingService"" />"
	strRst = strRst &"	</soap:Body>"
	strRst = strRst &"</soap:Envelope>"
	getAuctionCommonCodeList = strRst
End Function

'공통코드 중 출하지 Soap XML
Public Function getAuctionCommonCodeShipPlace(iccd)
	Dim strRst
	strRst = ""
	strRst = strRst &"<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst &"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst &"	<soap:Header>"
	strRst = strRst &"		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
	strRst = strRst &"			<Value>"&auctionTicket&"</Value>"
	strRst = strRst &"		</EncryptedTicket>"
	strRst = strRst &"	</soap:Header>"
	strRst = strRst &"  <soap:Body>"
	strRst = strRst &"		<GetShippingPlaceCode xmlns=""http://www.auction.co.kr/APIv1/ShoppingService"" >"
	strRst = strRst &"			<req Version=""1"" ShipmentPlace=""텐바이텐 물류센터"" />"
	strRst = strRst &"		</GetShippingPlaceCode >"
	strRst = strRst &"	</soap:Body>"
	strRst = strRst &"</soap:Envelope>"
	getAuctionCommonCodeShipPlace = strRst
End Function

'공통코드 중 택배조회 Soap XML
Public Function getAuctionCommonCodeGetDeliveryList(iccd)
	Dim strRst
	strRst = ""
	strRst = strRst &"<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst &"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst &"	<soap:Header>"
	strRst = strRst &"		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
	strRst = strRst &"			<Value>"&auctionTicket&"</Value>"
	strRst = strRst &"		</EncryptedTicket>"
	strRst = strRst &"	</soap:Header>"
	strRst = strRst &"  <soap:Body>"
	strRst = strRst &"		<"&iccd&" xmlns=""http://www.auction.co.kr/APIv1/AuctionService"" />"
	strRst = strRst &"	</soap:Body>"
	strRst = strRst &"</soap:Envelope>"
	getAuctionCommonCodeGetDeliveryList = strRst
End Function

'주문XML
Public Function getAuctionOrderList(iccd,sday)
	Dim strRst, isday
	strRst = ""
	strRst = strRst &"<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst &"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst &"	<soap:Header>"
	strRst = strRst &"		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
	strRst = strRst &"			<Value>"&auctionTicket&"</Value>"
	strRst = strRst &"		</EncryptedTicket>"
	strRst = strRst &"	</soap:Header>"
	strRst = strRst &"  <soap:Body>"
	strRst = strRst &"		<"&iccd&" xmlns=""http://www.auction.co.kr/APIv1/AuctionService"" >"
	strRst = strRst &"			<req DurationType=""ReceiptDate"" SearchType=""Nothing"">"
	strRst = strRst &"				<SearchDuration StartDate=""2015-08-01"" EndDate="""&sday&""" xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"" />"
	strRst = strRst &"			</req>"
	strRst = strRst &"		</"&iccd&">"
	strRst = strRst &"	</soap:Body>"
	strRst = strRst &"</soap:Envelope>"

	getAuctionOrderList = strRst
End Function

'주문XML2
Public Function getAuctionOrderList2(iccd,sday)
	Dim strRst, isday
	strRst = ""
	strRst = strRst &"<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst &"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst &"	<soap:Header>"
	strRst = strRst &"		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
	strRst = strRst &"			<Value>"&auctionTicket&"</Value>"
	strRst = strRst &"		</EncryptedTicket>"
	strRst = strRst &"	</soap:Header>"
	strRst = strRst &"  <soap:Body>"
	strRst = strRst &"		<"&iccd&" xmlns=""http://www.auction.co.kr/APIv1/AuctionService"" >"
	strRst = strRst &"			<req DurationType=""ReceiptDate"" SearchType=""Nothing"">"
	strRst = strRst &"				<SearchDuration StartDate=""2015-08-01"" EndDate="""&sday&""" xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"" />"
	strRst = strRst &"			</req>"
	strRst = strRst &"		</"&iccd&">"
	strRst = strRst &"	</soap:Body>"
	strRst = strRst &"</soap:Envelope>"
	getAuctionOrderList2 = strRst
End Function
'################################################# 각 기능 별 파라메터 정리 끝 ###############################################
%>