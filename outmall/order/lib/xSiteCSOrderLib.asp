<%
Function getApiUrl(mallid)
	Select Case mallid
		Case "lotteon"
			If application("Svr_Info") = "Dev" Then
				getApiUrl = "https://dev-openapi.lotteon.com"
			Else
				getApiUrl = "https://openapi.lotteon.com"
			End If
	End Select
End Function

Function getApiKey(mallid)
	Select Case mallid
		Case "lotteon"
			If application("Svr_Info") = "Dev" Then
				getApiKey = "5d5b2cb498f3d20001665f4e5451c4d923ac4e2c95df619996f35476"
			Else
				getApiKey = "5d5b2cb498f3d20001665f4e18a41621005d4c1ba262804ec7a10732"
			End If
	End Select
End Function

CONST gmarketTicket = "0A2799EE6A1B65CC78DA96AA52C7546B2181855E48A0A31EDD4F3A77C3C61015856FE3DE5D7828B129A31AAD5914D7060556616D3AB7F2A84008A600C89F5953A0362065429900D0EB25CEBEA0E1CAF9E784FBC4F36E86608F2CF44B40113ADF"
function GetCSOrderCancel_gmarket(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj
	dim i
	dim startdate, enddate
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql, IsDelete

	startdate = Left(DateAdd("d", 0, selldate), 10)
	enddate = Left(DateAdd("d", 1, selldate), 10)

	xmlURL = "https://tpl.gmarket.co.kr/v1/OrderCancelService.asmx"

	strRst = ""
	strRst = strRst + "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst + "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst + " <soap:Header>"
	strRst = strRst + "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "			<encTicket>" & gmarketTicket & "</encTicket>"
	strRst = strRst + "		</EncTicket>"
	strRst = strRst + "	</soap:Header>"
	strRst = strRst + "	<soap:Body>"
	strRst = strRst + "		<RequestOrderCancel xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "			<RequestOrderCancel ClaimStatus=""" & csSubGubun & """ SearchStartDate=""" & startdate & """ SearchEndDate=""" & enddate & """ />"
	strRst = strRst + "		</RequestOrderCancel>"
	strRst = strRst + "	</soap:Body>"
	strRst = strRst + "</soap:Envelope>"
	''response.write strRst

	''response.write "aaa" & startdate & " ~ " & enddate & " " & selldate

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(strRst)
	objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/RequestOrderCancel"
	objXML.send(strRst)

	if objXML.Status <> "200" then
		response.write "ERROR : 통신오류"
		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(Replace(objXML.responseText,"soap:",""))

	iInputCnt = 0
	If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "RequestOrderCancelResponse" Then
		set objArr = xmlDOM.getElementsByTagName("RequestOrderCancelResultT")

		for i = 0 to objArr.length - 1
			set obj = objArr.item(i)

'			<RequestOrderCancelResultT PackNo="long"
'									 ContrNo="long"
'									 ClaimStatus="ClaimReady or ClaimDone or ClaimReject or ClaimDoneG"
'									 ClaimType="ChangeMind or ChangeChoice or DelayDelivery or DamagedItem or ShippingMistake or InformationMistake or OutofStock or CouponNotAccept or Etc"
'									 ClaimReason="string"
'									 ItemSatus="OpenItem or UseItem or UnopenedItem or Etc"
'									 RequestClaimDate="dateTime"
'									 ClaimSolveDate="dateTime"
'									 RefundDate="dateTime"
'									 AdditionalMoney="decimal" />

			OutMallOrderSerial = obj.getAttribute("PackNo")
			OrgDetailKey = obj.getAttribute("ContrNo")
			CSDetailKey = ""
			divcd = "A008"
            IsDelete = "N"
			if (csSubGubun = "ClaimReject") then
				IsDelete = "Y"
			end if
			gubunname = obj.getAttribute("ClaimType")
			select case gubunname
				case "ChangeMind"
					gubunname = "단순변심"
				case "ChangeChoice"
					gubunname = "선택변경"
				case "DelayDelivery"
					gubunname = "배송지연"
				case "DamagedItem"
					gubunname = "상품불량"
				case "ShippingMistake"
					gubunname = "오배송"
				case "InformationMistake"
					gubunname = "상품정보틀림"
				case "OutofStock"
					gubunname = "품절"
				case "CouponNotAccept"
					gubunname = "쿠폰미적용"
				case "Etc"
					gubunname = "기타"
				case else
					gubunname = Replace(gubunname, "'", "")
			end select
			OutMallRegDate = Left(obj.getAttribute("RequestClaimDate"), 10)
			itemno = 0			'// 지마켓은 수량취소 없음, 전부취소

            strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
			strSql = strSql & " BEGIN "
			strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
			strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
			strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
			strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
			strSql = strSql & "		'', '', '', '', '', '' "
			strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
			strSql = strSql & " END "
			''strSql = strSql & " ELSE "
			''strSql = strSql & " BEGIN "
			''strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
			''strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
			''strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
			''strSql = strSql & " END "
			dbget.Execute strSql,iAssignedRow
			''response.write strSql & "<br />"
			''iAssignedRow = 1

            if IsDelete = "Y" then
                strSql = " update db_temp.dbo.tbl_xSite_TMPCS "
                strSql = strSql + " set OutMallCurrState = 'B008' "
                strSql = strSql + " where "
                strSql = strSql + " 	1 = 1 "
                strSql = strSql + " 	and SellSite = '" & sellsite & "' "
                strSql = strSql + " 	and OutMallOrderSerial = '" & OutMallOrderSerial & "' "
                strSql = strSql + " 	and CSDetailKey = '" & CSDetailKey & "' "
                strSql = strSql + " 	and OrgDetailKey = '" & OrgDetailKey & "' "
                strSql = strSql + " 	and divcd = '" & divcd & "' "
                dbget.Execute strSql
            end if

			if (iAssignedRow > 0) then
				iInputCnt = iInputCnt+iAssignedRow

				''주문 입력 이전 내역은 삭제 하자
				strSql = " update c "
				strSql = strSql + " set matchState='D'"
				strSql = strSql + " from db_temp.dbo.tbl_xSite_TMPOrder c "
				strSql = strSql + " WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
				strSql = strSql + " and orderserial is NULL"
				''response.write strSql & "<br />"
				dbget.Execute strSql

				'' CS 마스터정보. 업데이트
				strSql = " update c "
				strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
				strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
				strSql = strSql + " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
				strSql = strSql + " from "
				strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
				strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
				strSql = strSql + " on "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.SellSite = o.SellSite "
				strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
				strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
				strSql = strSql + " where "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.orderserial is NULL "
				strSql = strSql + " 	and o.orderserial is not NULL "
				strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
				''response.write strSql & "<br />"
				dbget.Execute strSql

				If divcd = "A008" Then
					strSql = " update c "
					strSql = strSql + " set c.currstate = 'B007' "
					strSql = strSql + " from "
					strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
					strSql = strSql + " left join db_temp.dbo.tbl_xSite_TMPOrder o "
					strSql = strSql + " on "
					strSql = strSql + "		1 = 1 "
					strSql = strSql + "		and c.SellSite = o.SellSite "
					strSql = strSql + "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
					strSql = strSql + "		and c.OrgDetailKey = o.OrgDetailKey "
					strSql = strSql + " where "
					strSql = strSql + "		1 = 1 "
					strSql = strSql + "		and c.orderserial is NULL "
					strSql = strSql + "		and o.SellSite is NULL "
					strSql = strSql + "		and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
					strSql = strSql + "		and c.currstate = 'B001' "
					strSql = strSql + "		and c.divcd = 'A008' "
					''rw strSql
					dbget.execute strSql
				end if
			end if
		next
	end if

	if (csSubGubun = "ClaimReject") then
		rw "취소철회 CS입력건수:"&iInputCnt
	else
		rw "주문취소 CS입력건수:"&iInputCnt
	end if

end function

function GetCSOrderReturn_gmarket(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj
	dim i
	dim startdate, enddate
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql, IsDelete

	startdate = Left(DateAdd("d", 0, selldate), 10)
	enddate = Left(DateAdd("d", 1, selldate), 10)

	xmlURL = "https://tpl.gmarket.co.kr/v1/ReturnService.asmx"

	strRst = ""
	strRst = strRst + "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst + "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst + "  <soap:Header>"
	strRst = strRst + "    <EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "      <encTicket>" & gmarketTicket & "</encTicket>"
	strRst = strRst + "    </EncTicket>"
	strRst = strRst + "  </soap:Header>"
	strRst = strRst + "  <soap:Body>"
	strRst = strRst + "    <RequestReturn xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "      <RequestReturn ClaimStatus=""" & csSubGubun & """ SearchStartDate=""" & startdate & """ SearchEndDate=""" & enddate & """ />"
	strRst = strRst + "    </RequestReturn>"
	strRst = strRst + "  </soap:Body>"
	strRst = strRst + "</soap:Envelope>"

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(strRst)
	objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/RequestReturn"
	objXML.send(strRst)

	if objXML.Status <> "200" then
		response.write "ERROR : 통신오류"
		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(Replace(objXML.responseText,"soap:",""))

	iInputCnt = 0
	If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "RequestReturnResponse" Then
		set objArr = xmlDOM.getElementsByTagName("RequestReturnResultT")

		for i = 0 to objArr.length - 1
			set obj = objArr.item(i)

'			<RequestReturnResultT
'				PackNo="long"
'				ContrNo="long"
'				ClaimStatus="ClaimReady or ClaimDone or ClaimReject or ClaimDoneG"
'				ClaimType="ChangeMind or ChangeChoice or DelayDelivery or DamagedItem or ShippingMistake or InformationMistake or OutofStock or Etc"
'				ClaimReason="string"
'				ItemStatus="OpenItem or UseItem or UnopenedItem or Etc"
'				RequestClaimDate="dateTime"
'				ClaimSolveDate="dateTime"
'				RefundDate="dateTime"
'				IsReserve="boolean"
'				ReserveType="string"
'				ReserveReason="string"
'				AdditionalMoney="decimal"
'				PayType="Done or MinusRefund or PaySeller or Enclosed or Etc"
'				WhoReturnFee="Buyer or Seller or Gmarket"
'				ReturnFee="decimal"
'				PickupType="Etc or GmktDesigned or SellerDesigned"
'				ExpressName="string"
'				InvoiceNo="string"
'				SenderZipcode="string"
'				SenderName="string"
'				SenderPhone1="string"
'				SenderPhone2="string"
'				SenderAddressFront="string"
'				SenderAddressBack="string"
'				PickupStatus="UnplannedPickup or PreparePickup or RequestPickup or ProgressPickup or CompletePickup or CancelPickup or FailPickup_Gather or FailPickup_Shipping"
'				PickupFailReason="string"
'				IsFastRefund="boolean" />

			OutMallOrderSerial = obj.getAttribute("PackNo")
			OrgDetailKey = obj.getAttribute("ContrNo")
			CSDetailKey = ""
			divcd = "A004"
            IsDelete = "N"
			if (csSubGubun = "ClaimReject") then
				IsDelete = "Y"
			end if
			gubunname = obj.getAttribute("ClaimType")
			select case gubunname
				case "ChangeMind"
					gubunname = "단순변심"
				case "ChangeChoice"
					gubunname = "선택변경"
				case "DelayDelivery"
					gubunname = "배송지연"
				case "DamagedItem"
					gubunname = "상품불량"
				case "ShippingMistake"
					gubunname = "오배송"
				case "InformationMistake"
					gubunname = "상품정보틀림"
				case "OutofStock"
					gubunname = "품절"
				case "CouponNotAccept"
					gubunname = "쿠폰미적용"
				case "Etc"
					gubunname = "기타"
				case else
					gubunname = Replace(gubunname, "'", "")
			end select
			OutMallRegDate = Left(obj.getAttribute("RequestClaimDate"), 10)
			itemno = 0			'// 지마켓은 수량반품 없음, 전부반품

            strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
			strSql = strSql & " BEGIN "
			strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
			strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
			strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
			strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
			strSql = strSql & "		'', '', '', '', '', '' "
			strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
			strSql = strSql & " END "
			''strSql = strSql & " ELSE "
			''strSql = strSql & " BEGIN "
			''strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
			''strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
			''strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
			''strSql = strSql & " END "
			dbget.Execute strSql,iAssignedRow
			''response.write strSql & "<br />"
			''iAssignedRow = 1

            if IsDelete = "Y" then
                strSql = " update db_temp.dbo.tbl_xSite_TMPCS "
                strSql = strSql + " set OutMallCurrState = 'B008' "
                strSql = strSql + " where "
                strSql = strSql + " 	1 = 1 "
                strSql = strSql + " 	and SellSite = '" & sellsite & "' "
                strSql = strSql + " 	and OutMallOrderSerial = '" & OutMallOrderSerial & "' "
                strSql = strSql + " 	and CSDetailKey = '" & CSDetailKey & "' "
                strSql = strSql + " 	and OrgDetailKey = '" & OrgDetailKey & "' "
                strSql = strSql + " 	and divcd = '" & divcd & "' "
                dbget.Execute strSql
            end if

			if (iAssignedRow > 0) then
				iInputCnt = iInputCnt+iAssignedRow

				'' CS 마스터정보. 업데이트
				strSql = " update c "
				strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
				strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
				strSql = strSql + " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
				strSql = strSql + " from "
				strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
				strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
				strSql = strSql + " on "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.SellSite = o.SellSite "
				strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
				strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
				strSql = strSql + " where "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.orderserial is NULL "
				strSql = strSql + " 	and o.orderserial is not NULL "
				strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
				''response.write strSql & "<br />"
				dbget.Execute strSql
			end if
		next
	end if

	if (csSubGubun = "ClaimReject") then
		rw "반품철회 CS입력건수:"&iInputCnt
	else
		rw "반품 CS입력건수:"&iInputCnt
	end if

end function

function GetCSOrderExchange_gmarket(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj
	dim i
	dim startdate, enddate
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql, IsDelete

	startdate = Left(DateAdd("d", 0, selldate), 10)
	enddate = Left(DateAdd("d", 1, selldate), 10)

	xmlURL = "https://tpl.gmarket.co.kr/v1/ExchangeService.asmx"

	strRst = ""
	strRst = strRst + "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst + "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst + "  <soap:Header>"
	strRst = strRst + "    <EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "      <encTicket>" & gmarketTicket & "</encTicket>"
	strRst = strRst + "    </EncTicket>"
	strRst = strRst + "  </soap:Header>"
	strRst = strRst + "  <soap:Body>"
	strRst = strRst + "    <RequestExchange xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "      <RequestExchange ClaimStatus=""" & csSubGubun & """ SearchStartDate=""" & startdate & """ SearchEndDate=""" & enddate & """ />"
	strRst = strRst + "    </RequestExchange>"
	strRst = strRst + "  </soap:Body>"
	strRst = strRst + "</soap:Envelope>"

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(strRst)
	objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/RequestExchange"
	objXML.send(strRst)

	if objXML.Status <> "200" then
		response.write "ERROR : 통신오류"
		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(Replace(objXML.responseText,"soap:",""))

	iInputCnt = 0
	If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "RequestExchangeResponse" Then
		set objArr = xmlDOM.getElementsByTagName("RequestExchangeResultT")

		for i = 0 to objArr.length - 1
			set obj = objArr.item(i)

'			<RequestExchangeResultT
'				PackNo="long"
'				ContrNo="long"
'				ClaimStatus="ClaimReady or ClaimDone or ClaimReject"
'				ClaimType="ChangeMind or ChangeChoice or DelayDelivery or DamagedItem or ShippingMistake or InformationMistake or FaultBySeller or Etc"
'				ClaimReason="string"
'				ItemStatus="OpenItem or UseItem or UnopenedItem or Etc"
'				RequestClaimDate="dateTime"
'				ClaimSolveDate="dateTime"
'				IsReserve="boolean"
'				ReserveType="Etc or FeeOfShipping or UnCheckedPickup"
'				ReserveReason="string"
'				ExchangeShippingFee="decimal"
'				WhoExchangeShippingFee="string"
'				PickupType="Etc or GmktDesigned or SellerDesigned"
'				PickupExpressName="string"
'				PickupInvoiceNo="string"
'				ForwardExpressName="string"
'				ForwardInvoiceNo="string"
'				SenderZipcode="string"
'				SenderName="string"
'				SenderPhone1="string"
'				SenderPhone2="string"
'				SenderAddressFront="string"
'				SenderAddressBack="string"
'				ReceiverZipcode="string"
'				ReceiverName="string"
'				ReceiverPhone1="string"
'				ReceiverPhone2="string"
'				ReceiverAddressFront="string"
'				ReceiverAddressBack="string"
'				PickupStatus="UnplannedPickup or PreparePickup or RequestPickup or ProgressPickup or CompletePickup or CancelPickup or FailPickup_Gather or FailPickup_Shipping"
'				ForwardStatus="UnplannedForward or PrepareForward or RequestForward or ProgressForward or CompleteForward or CancelForward or FailGather"
'				PickupFailReason="string"
'				PayType="Done or MinusRefund or PaySeller or Enclosed or Etc" />

			OutMallOrderSerial = obj.getAttribute("PackNo")
			OrgDetailKey = obj.getAttribute("ContrNo")
			CSDetailKey = ""
			divcd = "A000"
            IsDelete = "N"
			if (csSubGubun = "ClaimReject") then
				IsDelete = "Y"
			end if
			gubunname = obj.getAttribute("ClaimType")
			select case gubunname
				case "ChangeMind"
					gubunname = "단순변심"
				case "ChangeChoice"
					gubunname = "선택변경"
				case "DelayDelivery"
					gubunname = "배송지연"
				case "DamagedItem"
					gubunname = "상품불량"
				case "ShippingMistake"
					gubunname = "오배송"
				case "InformationMistake"
					gubunname = "상품정보틀림"
				case "OutofStock"
					gubunname = "품절"
				case "CouponNotAccept"
					gubunname = "쿠폰미적용"
				case "Etc"
					gubunname = "기타"
				case else
					gubunname = Replace(gubunname, "'", "")
			end select
			OutMallRegDate = Left(obj.getAttribute("RequestClaimDate"), 10)
			itemno = 0			'// 지마켓은 수량교환 없음, 전부교환

            strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
			strSql = strSql & " BEGIN "
			strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
			strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
			strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
			strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
			strSql = strSql & "		'', '', '', '', '', '' "
			strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
			strSql = strSql & " END "
			''strSql = strSql & " ELSE "
			''strSql = strSql & " BEGIN "
			''strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
			''strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
			''strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
			''strSql = strSql & " END "
			dbget.Execute strSql,iAssignedRow
			''response.write strSql & "<br />"
			''iAssignedRow = 1

            if IsDelete = "Y" then
                strSql = " update db_temp.dbo.tbl_xSite_TMPCS "
                strSql = strSql + " set OutMallCurrState = 'B008' "
                strSql = strSql + " where "
                strSql = strSql + " 	1 = 1 "
                strSql = strSql + " 	and SellSite = '" & sellsite & "' "
                strSql = strSql + " 	and OutMallOrderSerial = '" & OutMallOrderSerial & "' "
                strSql = strSql + " 	and CSDetailKey = '" & CSDetailKey & "' "
                strSql = strSql + " 	and OrgDetailKey = '" & OrgDetailKey & "' "
                strSql = strSql + " 	and divcd = '" & divcd & "' "
                dbget.Execute strSql
            end if

			if (iAssignedRow > 0) then
				iInputCnt = iInputCnt+iAssignedRow

				'' CS 마스터정보. 업데이트
				strSql = " update c "
				strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
				strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
				strSql = strSql + " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
				strSql = strSql + " from "
				strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
				strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
				strSql = strSql + " on "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.SellSite = o.SellSite "
				strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
				strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
				strSql = strSql + " where "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.orderserial is NULL "
				strSql = strSql + " 	and o.orderserial is not NULL "
				strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
				''response.write strSql & "<br />"
				dbget.Execute strSql
			end if
		next
	end if

	if (csSubGubun = "ClaimReject") then
		rw "교환철회 CS입력건수:"&iInputCnt
	else
		rw "교환 CS입력건수:"&iInputCnt
	end if

end function

function GetCSOrderAll_interpark(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj, objData, objDetail, objDetailArr
	dim i, j, k
	dim startdate, enddate
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql, IsDelete
	dim xmlSelldate
	dim CLMREQ_STAT

	xmlSelldate = Replace(selldate, "-", "")

	xmlURL = "https://joinapi.interpark.com"
	xmlURL = xmlURL + "/order/OrderClmAPI.do?_method=cnclNClmReqListForSingle&sc.entrId=10X10&sc.supplyEntrNo=3000010614&sc.supplyCtrtSeq=2&sc.strDate=" + xmlSelldate + "000000" + "&sc.endDate=" + xmlSelldate + "235959"

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.setTimeouts 5000,80000,80000,80000
	objXML.Send()

	if objXML.Status = "200" then
		objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		''response.write objData
		''dbget.close : response.end
	else
		response.write "ERROR : 통신오류"
		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML replace(objData,"&","＆")

	Set obj = xmlDOM.selectSingleNode("/ORDER_LIST/RESULT/CODE")

	if obj is Nothing then
		''response.write "내역없음 : 종료"
		GetCSOrderAll_interpark = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	if (obj.text <> "000") then
		response.write "ERROR : 알 수 없는 오류"
		GetCSOrderAll_interpark = False
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	Set objArr = xmlDOM.selectNodes("/ORDER_LIST/ORDER")

	iInputCnt = 0
	for i = 0 to objArr.length - 1
		set obj = objArr.item(i)

		OutMallOrderSerial = obj.selectSingleNode("ORD_NO").text
		CSDetailKey = obj.selectSingleNode("CLMREQ_SEQ").text
		select case obj.selectSingleNode("CLMREQ_TP").text
			case "1"
				divcd = "A008"
			case "2"
				divcd = "A004"
			case "3"
				divcd = "A000"
			case else
				response.write "unknown data[0] : " & obj.selectSingleNode("CLMREQ_TP").text
				dbget.close : response.end
				divcd = 1/0		'// 에러
		end select
		OutMallRegDate = Left(obj.selectSingleNode("CLMREQ_DTS").text, 8)
		OutMallRegDate = Left(OutMallRegDate, 4) & "-" & Mid(OutMallRegDate, 5, 2) & "-" & Right(OutMallRegDate, 2)

		set objDetailArr = obj.selectNodes("PRODUCT/PRD")
		for j = 0 to objDetailArr.length - 1
			set objDetail = objDetailArr.item(j)

			OrgDetailKey = objDetail.selectSingleNode("ORD_SEQ").text

			CLMREQ_STAT = objDetail.selectSingleNode("CLMREQ_STAT").text
            IsDelete = "N"
			if CLMREQ_STAT = "2" then
				select case divcd
					case "A008"
						IsDelete = "Y"
					case "A088"
						IsDelete = "Y"
					case "A004"
						IsDelete = "Y"
					case "A044"
						IsDelete = "Y"
					case "A000"
						IsDelete = "Y"
					case "A090"
						IsDelete = "Y"
					case else
						response.write "unknown data[1] : " & divcd
						dbget.close : response.end
						divcd = 1/0		'// 에러
				end select

				OutMallRegDate = Left(objDetail.selectSingleNode("CLMREQ_CNCL_DTS").text, 8)
				OutMallRegDate = Left(OutMallRegDate, 4) & "-" & Mid(OutMallRegDate, 5, 2) & "-" & Right(OutMallRegDate, 2)
			end if

			itemno = objDetail.selectSingleNode("CLMREQ_QTY").text

			gubunname = objDetail.selectSingleNode("CLMREQ_RSN_TPNM").text
			gubunname = Replace(gubunname, "'", "")

            strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
			strSql = strSql & " BEGIN "
			strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
			strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
			strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
			strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
			strSql = strSql & "		'', '', '', '', '', '' "
			strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
			strSql = strSql & " END "
			''strSql = strSql & " ELSE "
			''strSql = strSql & " BEGIN "
			''strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
			''strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
			''strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
			''strSql = strSql & " END "
			dbget.Execute strSql,iAssignedRow
			''response.write strSql & "<br />"
			''iAssignedRow = 1

            if IsDelete = "Y" then
                strSql = " update db_temp.dbo.tbl_xSite_TMPCS "
                strSql = strSql + " set OutMallCurrState = 'B008' "
                strSql = strSql + " where "
                strSql = strSql + " 	1 = 1 "
                strSql = strSql + " 	and SellSite = '" & sellsite & "' "
                strSql = strSql + " 	and OutMallOrderSerial = '" & OutMallOrderSerial & "' "
                strSql = strSql + " 	and CSDetailKey = '" & CSDetailKey & "' "
                strSql = strSql + " 	and OrgDetailKey = '" & OrgDetailKey & "' "
                strSql = strSql + " 	and divcd = '" & divcd & "' "
                dbget.Execute strSql
            end if

			if (iAssignedRow > 0) then
				iInputCnt = iInputCnt+iAssignedRow

				'' CS 마스터정보. 업데이트
				strSql = " update c "
				strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
				strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
				strSql = strSql + " , c.OrderName = o.OrderName "
				strSql = strSql + " from "
				strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
				strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
				strSql = strSql + " on "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.SellSite = o.SellSite "
				strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
				strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
				strSql = strSql + " where "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.orderserial is NULL "
				strSql = strSql + " 	and o.orderserial is not NULL "
				strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
				''response.write strSql & "<br />"
				dbget.Execute strSql
			end if
		next
	next

	rw "CS취소입력건수1 : "&iInputCnt

end function

function GetCSOrderChgRet_interpark(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj, objData, objDetail, objDetailArr
	dim i, j, k
	dim startdate, enddate
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno, OrgOutMallOrderSerial
	dim iAssignedRow, iInputCnt
	dim strSql, IsDelete
	dim xmlSelldate
	dim CLMREQ_STAT

	xmlSelldate = Replace(selldate, "-", "")

	xmlURL = "https://joinapi.interpark.com"
	xmlURL = xmlURL + "/order/OrderClmAPI.do?_method=clmListForSingle&sc.entrId=10X10&sc.supplyEntrNo=3000010614&sc.supplyCtrtSeq=2&sc.strDate=" + xmlSelldate + "000000" + "&sc.endDate=" + xmlSelldate + "235959"

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.setTimeouts 5000,80000,80000,80000
	objXML.Send()

	if objXML.Status = "200" then
		objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		If session("ssBctID")="kjy8517" Then
			response.write "<textarea cols=100 rows=30>" & objData & "</textarea>"
		End If
'		response.write objData
'		dbget.close : response.end
	else
		response.write "ERROR : 통신오류"
		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML replace(objData,"&","＆")

	Set obj = xmlDOM.selectSingleNode("/ORDER_LIST/RESULT/CODE")

	if obj is Nothing then
		''response.write "내역없음 : 종료"
		GetCSOrderChgRet_interpark = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	if (obj.text <> "000") then
		response.write "ERROR : 알 수 없는 오류"
		GetCSOrderChgRet_interpark = False
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	Set objArr = xmlDOM.selectNodes("/ORDER_LIST/ORDER")

	iInputCnt = 0
	for i = 0 to objArr.length - 1
		set obj = objArr.item(i)
		OrgOutMallOrderSerial = obj.selectSingleNode("CLM_NO").text		'원주문번호가 아니다. ORD_NO가 원주문번호이다. CLM_NO 값으로 정산미매칭 주문번호 찾아야 한다.
		OutMallOrderSerial = obj.selectSingleNode("ORD_NO").text
		CSDetailKey = obj.selectSingleNode("CLM_SEQ").text
		select case obj.selectSingleNode("CLM_CRT_TP").text
			case "01"
				divcd = "A008"
			case "02"
				divcd = "A004"
			case "03"
				divcd = "A000"
			case else
				response.write "unknown data[0] : " & obj.selectSingleNode("CLM_CRT_TP").text
				dbget.close : response.end
				divcd = 1/0		'// 에러
		end select
'		OutMallRegDate = Left(obj.selectSingleNode("CLMREQ_DTS").text, 8)
'		OutMallRegDate = Left(OutMallRegDate, 4) & "-" & Mid(OutMallRegDate, 5, 2) & "-" & Right(OutMallRegDate, 2)

		set objDetailArr = obj.selectNodes("PRODUCT/PRD")
		for j = 0 to objDetailArr.length - 1
			set objDetail = objDetailArr.item(j)
			OrgDetailKey = objDetail.selectSingleNode("ORD_SEQ").text
			CLMREQ_STAT = objDetail.selectSingleNode("CURRENT_CLMPRD_STAT").text
			IsDelete = "N"

			if CLMREQ_STAT = "40" then
				select case divcd
					case "A008"
						IsDelete = "Y"
					case "A088"
						IsDelete = "Y"
					case "A004"
						IsDelete = "Y"
					case "A044"
						IsDelete = "Y"
					case "A000"
						IsDelete = "Y"
					case "A090"
						IsDelete = "Y"
				end select
			End If

			OutMallRegDate = Left(objDetail.selectSingleNode("CLM_DTS").text, 8)
			OutMallRegDate = Left(OutMallRegDate, 4) & "-" & Mid(OutMallRegDate, 5, 2) & "-" & Right(OutMallRegDate, 2)

			itemno = objDetail.selectSingleNode("CLM_QTY").text

			gubunname = objDetail.selectSingleNode("CLM_RSN_TPNM").text
			gubunname = Replace(gubunname, "'", "")

            strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
			strSql = strSql & " BEGIN "
			strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
			strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
			strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno, orgOutMallOrderSerial) VALUES "
			strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
			strSql = strSql & "		'', '', '', '', '', '' "
			strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ", '"& OrgOutMallOrderSerial &"') "
			strSql = strSql & " END "
			dbget.Execute strSql,iAssignedRow

            if IsDelete = "Y" then
                strSql = " update db_temp.dbo.tbl_xSite_TMPCS "
                strSql = strSql + " set OutMallCurrState = 'B008' "
                strSql = strSql + " where "
                strSql = strSql + " 	1 = 1 "
                strSql = strSql + " 	and SellSite = '" & sellsite & "' "
                strSql = strSql + " 	and OutMallOrderSerial = '" & OutMallOrderSerial & "' "
                strSql = strSql + " 	and CSDetailKey = '" & CSDetailKey & "' "
                strSql = strSql + " 	and OrgDetailKey = '" & OrgDetailKey & "' "
                strSql = strSql + " 	and divcd = '" & divcd & "' "
                dbget.Execute strSql
            end if

			if (iAssignedRow > 0) then
				iInputCnt = iInputCnt+iAssignedRow

				'' CS 마스터정보. 업데이트
				strSql = " update c "
				strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
				strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
				strSql = strSql + " , c.OrderName = o.OrderName "
				strSql = strSql + " from "
				strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
				strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
				strSql = strSql + " on "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.SellSite = o.SellSite "
				strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
				strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
				strSql = strSql + " where "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.orderserial is NULL "
				strSql = strSql + " 	and o.orderserial is not NULL "
				strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
				''response.write strSql & "<br />"
				dbget.Execute strSql
			end if
		next
	next
	rw "CS입력건수2 : "&iInputCnt
end function

Function GetCSOrderReturn_coupang(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst, iRbody, strObj
	dim objXML, xmlDOM, objArr, obj
	dim i, j, k
	dim startdate, enddate, retCode, strObjValues, strObjValuesexchangeItems
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno, OrderName, OrderHpNo
	dim iAssignedRow, iInputCnt
	dim strSql, ResultCode, ResultMsg, cancelOrderList, SubNodes
	Dim ClaimSeq, ClaimMemo, RegYMD
'response.end
	startdate = Left(DateAdd("d", 0, selldate), 10)
'RECEIPT : 접수
'PROGRESS : 진행
'SUCCESS : 완료
'REJECT : 불가
'CANCEL : 취소
'startdate = "2018-06-11"
	divcd = "A000"
	iInputCnt = 0
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://xapi.10x10.co.kr:8080/Deliveries/Coupang/exchange/RECEIPT/"&startdate, false
		'objXML.open "GET", "http://xapi.10x10.co.kr:8080/Deliveries/Coupang/exchange/PROGRESS/"&startdate, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			'response.write iRbody
			'response.end
			Set strObj = JSON.parse(iRbody)
				retCode			= strObj.message
				If retCode = "SUCCESS" Then
					set strObjValues = strObj.value
						For i=0 to strObjValues.length-1
							CSDetailKey			= strObjValues.get(i).exchangeId				'교환 아이디
							OutMallOrderSerial	= strObjValues.get(i).orderId					'주문번호
							'rw strObjValues.get(i).vendorId									'벤더 아이디
							'rw strObjValues.get(i).orderDeliveryStatusCode	'주문배송상태 | ACCEPT : 결제완료, INSTRUCT : 상품준비중, DEPARTURE : 배송지시, DELIVERING : 배송중, FINAL_DELIVERY : 배송완료, NONE_TRACKING : 업체직접배송(배송 연동 미적용), 추적불가
							'rw strObjValues.get(i).exchangeStatus			'교환상태 | RECEIPT : 접수, PROGRESS : 진행, SUCCESS : 완료, REJECT : 불가, CANCEL : 철회
							'rw strObjValues.get(i).referType				'접수경로 | VENDOR : 벤더, CS_CENTER : CS, WEB_PC : 웹 PC, WEB_MOBILE : 웹 모바일
							'rw strObjValues.get(i).faultType				'귀책 | COUPANG : 쿠팡과실, VENDOR : 업체과실, CUSTOMER : 고객과실, GENERAL : 일반
							'rw strObjValues.get(i).exchangeAmount			'교환배송비
							'rw strObjValues.get(i).reason					'교환접수사유 | 더이상 사용하지 않는 필드
							'rw strObjValues.get(i).reasonCode				'교환사유코드 | DEFECT : 결함, WRONGITEM : 오배송, OMISSION : 누락, OPTIONCHANGE : 옵션변경, ETC : 기타, BROKEN : 파손, ADDRESSCHANGE : 배송지변경, LOST : 상품분실
							gubunname			= strObjValues.get(i).reasonCodeText			'교환사유설명
							'rw strObjValues.get(i).reasonEtcDetail			'교환사유상세설명
							'rw strObjValues.get(i).cancelReason			'교환철회사유
							'rw strObjValues.get(i).createdByType			'최초 등록자 유형 | CUSTOMER : 고객, COUNSELOR : 상담사, COUPANG : 내부직원, VENDOR : 업체, ETC : 기타
							OutMallRegDate		= strObjValues.get(i).createdAt				'등록일시
							OutMallRegDate		= Replace(OutMallRegDate, "T", " ")
							'rw strObjValues.get(i).modifiedByType			'수정자 | CUSTOMER : 고객, COUNSELOR : 상담사, COUPANG : 내부직원, VENDOR : 업체, ETC : 기타, TRACKING : 배송추적
							'rw strObjValues.get(i).modifiedAt				'수정일시
							set strObjValuesexchangeItems = strObjValues.get(i).exchangeItemDtoV1s
								For j=0 to strObjValuesexchangeItems.length-1
									''OrgDetailKey = strObjValuesexchangeItems.get(j).exchangeItemId	'교환 상품 아이디
									OrgDetailKey = strObjValuesexchangeItems.get(j).orderItemId				'원주문 아이템ID
									'rw strObjValuesexchangeItems.get(j).orderItemUnitPrice			'원주문 아이템 단가
									'rw strObjValuesexchangeItems.get(j).orderItemName				'원주문 아이템 명
									'rw strObjValuesexchangeItems.get(j).orderPackageId				'원주문 패키지 ID
									'rw strObjValuesexchangeItems.get(j).orderPackageName			'원주문 패키지명
									'rw strObjValuesexchangeItems.get(j).targetItemId				'교환 아이템 ID
									'rw strObjValuesexchangeItems.get(j).targetItemUnitPrice		'교환 아이템 단가
									'rw strObjValuesexchangeItems.get(j).targetItemName				'교환 아이템 명
									'rw strObjValuesexchangeItems.get(j).targetPackageId			'교환 패키지 ID
									'rw strObjValuesexchangeItems.get(j).targetPackageName			'교환 패키지 명
									itemno = strObjValuesexchangeItems.get(j).quantity				'교환 수량
									'rw strObjValuesexchangeItems.get(j).orderItemDeliveryComplete	'True /False : 원주문 아이템 배송완료 여부
									'rw strObjValuesexchangeItems.get(j).orderItemReturnComplete	'True /False : 원주문 아이템 반품완료 여부
									'rw strObjValuesexchangeItems.get(j).targetItemDeliveryComplete	'True /False : 교환 아이템 배송완료 여부
									'rw strObjValuesexchangeItems.get(j).createdAt					'생성 일시
									'rw strObjValuesexchangeItems.get(j).modifiedAt					'수정 일시
									'rw strObjValuesexchangeItems.get(j).originalShipmentBoxId		'원배송번호
						            strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
									strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
									strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
									strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
									strSql = strSql & "		'', '', '', '', '', '' "
									strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
									strSql = strSql & " END "
									strSql = strSql & " ELSE "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
									strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
									strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
									strSql = strSql & " END "
									'rw strSql
									dbget.Execute strSql,iAssignedRow

									if (iAssignedRow > 0) then
										iInputCnt = iInputCnt+iAssignedRow

										'' CS 마스터정보. 업데이트
										strSql = " update c "
										strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
										strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
										strSql = strSql + " , c.OrderName = o.OrderName "
										strSql = strSql + " from "
										strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
										strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
										strSql = strSql + " on "
										strSql = strSql + " 	1 = 1 "
										strSql = strSql + " 	and c.SellSite = o.SellSite "
										strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
										strSql = strSql + " 	and c.OrgDetailKey = o.OutMallOptionNo "
										strSql = strSql + " where "
										strSql = strSql + " 	1 = 1 "
										strSql = strSql + " 	and c.orderserial is NULL "
										strSql = strSql + " 	and o.orderserial is not NULL "
										strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										''response.write strSql & "<br />"
										dbget.Execute strSql
									end if
								Next
							set strObjValuesexchangeItems = nothing
							'rw "-----------------------------------"
						Next
					set strObjValues = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing

	rw "CS 교환 입력건수:"&iInputCnt

End Function

Function GetCSOrderCancel_coupang(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst, iRbody, strObj
	dim objXML, xmlDOM, objArr, obj
	dim i, j, k
	dim startdate, enddate, retCode, strObjValues, strObjValuesReturnItems
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno, OrderName, OrderHpNo
	dim iAssignedRow, iInputCnt
	dim strSql, ResultCode, ResultMsg, cancelOrderList, SubNodes
	Dim ClaimSeq, ClaimMemo, RegYMD
    dim OrgOutMallOrderSerial

	startdate = Left(DateAdd("d", 0, selldate), 10)
'RU : 출고중지요청
'CC : 반품완료
'PR : 쿠팡확인요청
'UC : 반품접수
'startdate = "2018-06-11"
	iInputCnt = 0
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://xapi.10x10.co.kr:8080/Deliveries/Coupang/return/UC/"&startdate&"/"&csGubun, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			'response.write iRbody
			'response.end
			Set strObj = JSON.parse(iRbody)
				retCode			= strObj.message
				If retCode = "SUCCESS" Then
					set strObjValues = strObj.value
						For i=0 to strObjValues.length-1

							'// ================================================
							'// *구매자가 배송준비중 ~ 배송전 사이에 취소한 케이스도 반품에서 다루고 있습니다.
							'// ================================================

                            OrgOutMallOrderSerial	= strObjValues.get(i).orderId				'주문번호
							CSDetailKey				= strObjValues.get(i).receiptId				'취소(반품)접수번호
							OutMallOrderSerial		= strObjValues.get(i).orderId				'주문번호
							'rw strObjValues.get(i).paymentId				'결제번호
                            'rw strObjValues.get(i).receiptType									'취소유형 RETURN or CANCEL
                            if (strObjValues.get(i).receiptType = "CANCEL") then
                                divcd = "A008"
                            elseif (strObjValues.get(i).receiptType = "RETURN") then
                                divcd = "A004"
                            else
                                divcd = "AXXX"
                            end if
							'rw strObjValues.get(i).receiptStatus			'취소(반품)진행상태 | RELEASE_STOP_UNCHECKED : 출고중지요청, RETURNS_UNCHECKED : 반품접수, VENDOR_WAREHOUSE_CONFIRM : 입고완료, REQUEST_COUPANG_CHECK : 쿠팡확인요청, RETURNS_COMPLETED : 반품완료
							OutMallRegDate		= strObjValues.get(i).createdAt				'취소(반품) 접수시간
							OutMallRegDate		= Replace(OutMallRegDate, "T", " ")
							'rw strObjValues.get(i).modifiedAt				'취소(반품) 상태 최종 변경시간
							OrderName			= strObjValues.get(i).requesterName				'반품 신청인 이름
							OrderHpNo			= strObjValues.get(i).requesterPhoneNumber		'반품 신청인 전화번호
							gubunname			= strObjValues.get(i).cancelReasonCategory1	'반품 사유 카테고리1
							'rw strObjValues.get(i).cancelReasonCategory2	'반품 사유 카테고리2
							'rw strObjValues.get(i).cancelReason				'취소사유 상세내역
							'rw strObjValues.get(i).cancelCountSum			'총 취소수량
							'rw strObjValues.get(i).returnDeliveryId			'반품배송번호
							'rw strObjValues.get(i).returnDeliveryType		'회수종류 | 전담택배, 연동택배, 수기관리
							'rw strObjValues.get(i).releaseStopStatus		'출고중지처리상태 | 미처리, 처리(이미출고), 처리(출고중지), 자동처리(이미출고), 비대상
							'rw strObjValues.get(i).enclosePrice			'동봉배송비
							'rw strObjValues.get(i).faultByType				'귀책타입 | Coupang 과실 : COUPANG, 협력사 과실 : VENDOR, 고객 과실 : CUSTOMER, 물류과실 : WMS, 일반 : GENERAL
							'rw strObjValues.get(i).preRefund				'선환불야부
							'rw strObjValues.get(i).completeConfirmDate		'완료확인종류 | 파트너확인 : VENDOR_CONFIRM, 미확인 : UNDEFINED, CS 대리확인 : CS_CONFIRM, CS 손실확인 : CS_LOSS_CONFIRM
							'rw strObjValues.get(i).completeConfirmType		'완료확인시간
							set strObjValuesReturnItems = strObjValues.get(i).returnItems
								For j=0 to strObjValuesReturnItems.length-1
									'rw strObjValuesReturnItems.get(j).vendorItemPackageId		'딜번호
									'rw strObjValuesReturnItems.get(j).vendorItemPackageName	'딜명
									OrgDetailKey = strObjValuesReturnItems.get(j).vendorItemId	'벤더아이템번호
									'rw strObjValuesReturnItems.get(j).vendorItemName			'벤더아이템명
									itemno = strObjValuesReturnItems.get(j).purchaseCount		'취소 수량
									'rw strObjValuesReturnItems.get(j).cancelCount				'주문 수량
									'rw strObjValuesReturnItems.get(j).shipmentBoxId			'원 배송번호
									'rw strObjValuesReturnItems.get(j).sellerProductId			'업체등록상품번호
									'rw strObjValuesReturnItems.get(j).sellerProductName		'업체등록상품

						            strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
									strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
									strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno, OrgOutMallOrderSerial) VALUES "
									strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
									strSql = strSql & "		'', '', '', '', '', '' "
									strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ", '" & OrgOutMallOrderSerial & "') "
									strSql = strSql & " END "
									strSql = strSql & " ELSE "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
									strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
									strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
									strSql = strSql & " END "
									'rw strSql
									dbget.Execute strSql,iAssignedRow

									if (iAssignedRow > 0) then
										iInputCnt = iInputCnt+iAssignedRow

										'' CS 마스터정보. 업데이트
										strSql = " update c "
										strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
										strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
										strSql = strSql + " , c.OrderName = o.OrderName "
										strSql = strSql + " from "
										strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
										strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
										strSql = strSql + " on "
										strSql = strSql + " 	1 = 1 "
										strSql = strSql + " 	and c.SellSite = o.SellSite "
										strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
										strSql = strSql + " 	and c.OrgDetailKey = o.OutMallOptionNo "
										strSql = strSql + " where "
										strSql = strSql + " 	1 = 1 "
										strSql = strSql + " 	and c.orderserial is NULL "
										strSql = strSql + " 	and o.orderserial is not NULL "
										strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										''response.write strSql & "<br />"
										dbget.Execute strSql
									end if
								Next
							set strObjValuesReturnItems = nothing
							'rw "-----------------------------------"
						Next
					set strObjValues = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing

	If csGubun = "CANCEL" Then
		rw "CS 취소 입력건수:"&iInputCnt
	Else
		rw "CS 반품 입력건수:"&iInputCnt
	END If
End Function

Function GetCSOrderCancel_11st1010(sellsite, csGubun, csSubGubun, selldate)
	dim objXML, xmlDOM, obj, iRbody
	dim i, j, k
	dim startdate, enddate
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt, beasongNum11st
	dim strSql, ResultCode, SubNodes
	Dim dlvCnclYn, request11Url
	Dim APIkey, Nodes, ordCnRsnCd, ordCnStatCd
	APIkey = "a2319e071dbc304243ee60abd07e9664"
	Select Case csGubun
		Case "ClaimReady"
			request11Url = "http://api.11st.co.kr/rest/claimservice/cancelorders"
		Case "ClaimDone"
			request11Url = "http://api.11st.co.kr/rest/claimservice/canceledorders"
	End Select

	startdate = Left(DateAdd("d", 0, selldate), 10)
	startdate = Replace(startdate, "-", "") & "0000"

	enddate = Left(DateAdd("d", 0, selldate), 10)
	enddate = Replace(enddate, "-", "") & "2359"
	divcd = "A008"
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", request11Url&"/"&startdate&"/"&enddate, false
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				Set Nodes = xmlDOM.getElementsByTagName("ns2:order")
					iInputCnt = 0
					For each SubNodes in Nodes
						OutMallRegDate = LEFT(SubNodes.getElementsByTagName("createDt")(0).text,10)		'클레임 요청 일시
						beasongNum11st = SubNodes.getElementsByTagName("dlvNo")(0).text					'배송번호
						'rw SubNodes.getElementsByTagName("ordCnDtlsRsn")(0).text						'사유코드에 대한 상세내역
						itemno = SubNodes.getElementsByTagName("ordCnQty")(0).text						'클레임 수량
						'rw SubNodes.getElementsByTagName("ordCnMnbdCd")(0).text						'클레임 등록주체 | 01 : 구매자, 02 : 판매자
						ordCnRsnCd = SubNodes.getElementsByTagName("ordCnRsnCd")(0).text				'클레임 사유코드 | 00 : 등록주체 구매자 : 무통장 미입금 취소, 04 : 등록주체 구매자 : 판매자의 배송 처리가 늦음, 06 : 등록주체 구매자 : 판매자의 상품 정보가 잘못됨 등록주체 판매자 : 배송 지연 예상, 07 : 등록주체 구매자 : 동일 상품 재주문(주문정보수정) 등록주체 판매자 : 상품/가격 정보 잘못 입력, 08 : 등록주체 구매자 : 주문상품의 품절/재고없음 등록주체 판매자 : 상품 품절(전체옵션), 09 : 등록주체 구매자 : 11번가 내 다른 상품으로 재주문 등록주체 판매자 : 옵션 품절(해당옵션), 10 : 등록주체 구매자 : 타사이트 상품 주문 등록주체 판매자 : 고객변심, 11 : 등록주체 구매자 : 상품에 이상없으나 구매 의사 없어짐, 12 : 등록주체 구매자 : 기타(구매자 책임사유), 13 : 등록주체 구매자 : 기타(판매자 책임사유), 99 : 등록주체 구매자 : 기타, 14 : 구매의사 없어짐, 15 : 색상/사이즈/주문정보 변경, 16 : 다른 상품 잘못 주문, 17 : 배송지연으로 취소, 18 : 상품품절, 재고없음
						select case ordCnRsnCd
							case "10", "11", "14", "16"
								gubunname = "단순변심"
							case "15"
								gubunname = "선택변경"
							case "04", "17"
								gubunname = "배송지연"
							case "06", "07"
								gubunname = "상품정보틀림"
							case "08", "09", "18"
								gubunname = "품절"
							case "00", "12", "13", "99"
								gubunname = "기타"
							case else
								gubunname = Replace(gubunname, "'", "")
						end select
						ordCnStatCd = SubNodes.getElementsByTagName("ordCnStatCd")(0).text				'클레임 상태 | 01 : 취소요청, 02 : 취소완료
						OutMallOrderSerial = SubNodes.getElementsByTagName("ordNo")(0).text				'11번가 주문번호
						CSDetailKey = SubNodes.getElementsByTagName("ordPrdCnSeq")(0).text				'외부몰 클레임 번호
						OrgDetailKey = SubNodes.getElementsByTagName("ordPrdSeq")(0).text				'주문순번

						strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
						strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
						strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
						strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
						strSql = strSql & "		'', '', '', '', '', '' "
						strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
						strSql = strSql & " END "
						strSql = strSql & " ELSE "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
						strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
						strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
						strSql = strSql & " END "
						'rw strSql
						dbget.Execute strSql,iAssignedRow

						if (iAssignedRow > 0) then
							iInputCnt = iInputCnt+iAssignedRow

							'' CS 마스터정보. 업데이트
							strSql = " update c "
							strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
							strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
							strSql = strSql + " , c.OrderName = o.OrderName "
							strSql = strSql + " from "
							strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
							strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
							strSql = strSql + " on "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.SellSite = o.SellSite "
							strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
							strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
							strSql = strSql + " where "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.orderserial is NULL "
							strSql = strSql + " 	and o.orderserial is not NULL "
							strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
							''response.write strSql & "<br />"
							dbget.Execute strSql
						end if
						' rw OutMallRegDate
						' rw beasongNum11st
						' rw itemno
						' rw ordCnRsnCd
						' rw gubunname
						' rw ordCnStatCd
						' rw OutMallOrderSerial
						' rw CSDetailKey
						' rw OrgDetailKey
						'rw SubNodes.getElementsByTagName("prdNo")(0).text			'상품번호
						'rw SubNodes.getElementsByTagName("slctPrdOptNm")(0).text	'클레임 옵션명
						'rw SubNodes.getElementsByTagName("referSeq")(0).text		'원클릭체크아웃 주문코드
					Next
				Set Nodes = nothing
			Set xmlDOM = nothing
		End If
	SET objXML = nothing
	Select Case csGubun
		Case "ClaimReady"
			rw "CS 취소요청 입력건수:"&iInputCnt
		Case "ClaimDone"
			rw "CS 취소완료 입력건수:"&iInputCnt
	End Select
'	response.end
End Function


Function GetCSOrderExchange_11st1010(sellsite, csGubun, csSubGubun, selldate)
	dim objXML, xmlDOM, obj, iRbody
	dim i, j, k
	dim startdate, enddate
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt, beasongNum11st
	dim strSql, ResultCode, SubNodes
	Dim dlvCnclYn
	Dim APIkey, Nodes, clmReqRsn, ordCnStatCd
	APIkey = "a2319e071dbc304243ee60abd07e9664"
	startdate = Left(DateAdd("d", 0, selldate), 10)
	startdate = Replace(startdate, "-", "") & "0000"

	enddate = Left(DateAdd("d", 0, selldate), 10)
	enddate = Replace(enddate, "-", "") & "2359"
	divcd = "A000"
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://api.11st.co.kr/rest/claimservice/exchangeorders/"&startdate&"/"&enddate, false
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
'rw iRbody
				Set Nodes = xmlDOM.getElementsByTagName("ns2:order")
					iInputCnt = 0
					For each SubNodes in Nodes
						' rw SubNodes.getElementsByTagName("addDlvCst")(0).text
						' rw SubNodes.getElementsByTagName("affliateBndlDlvSeq")(0).text 			'무료교환 여부 | 현재 미사용 필드 0 : 무료교환, 1 : 일반교환(유료)
						' rw SubNodes.getElementsByTagName("appmtDlvCst")(0).text					'11번가 지정반품 택배비
						' rw SubNodes.getElementsByTagName("clmDlvCstMthd")(0).text					'결제방법 | 01 : 신용카드, 02 : 포인트, 03 : 박스에 동봉, 04 : 판매자에게 직접송금, null : 클레임 사유가 판매자일 경우
						' rw SubNodes.getElementsByTagName("clmLstDlvCst")(0).text					'교환배송비
						' rw SubNodes.getElementsByTagName("clmReqCont")(0).text					'사유코드에 대한 상세내역
						itemno = SubNodes.getElementsByTagName("clmReqQty")(0).text					'클레임 수량
						clmReqRsn = SubNodes.getElementsByTagName("clmReqRsn")(0).text				'클레임 사유코드
						select case clmReqRsn
							case "101"
								gubunname = "단순변심"
							case "ChangeChoice"
								gubunname = "선택변경"
							case "112"
								gubunname = "배송지연"
							case "111", "207", "216"
								gubunname = "상품불량"
							case "108", "208"
								gubunname = "오배송"
							case "105", "210"
								gubunname = "상품정보틀림"
							case "OutofStock"
								gubunname = "품절"
							case "CouponNotAccept"
								gubunname = "쿠폰미적용"
							case "113", "114", "115", "116", "117", "118", "119", "198", "199", "206", "209", "212", "213", "211", "214", "215", "217"
								gubunname = "기타"
							case else
								gubunname = Replace(gubunname, "'", "")
						end select

						CSDetailKey = SubNodes.getElementsByTagName("clmReqSeq")(0).text			'외부몰 클레임 번호
						' rw SubNodes.getElementsByTagName("clmStat")(0).text						'클레임 상태 | 103 : 재결제대기중, 104 : 반품보류, 105 : 반품신청, 106 : 반품완료, 107 : 반품거부, 108 : 반품철회, 109 : 반품완료보류, 201 : 교환신청, 212 : 교환승인, 214 : 교환보류, 221 : 교환발송완료, 232 : 교환거부, 233 : 교환철회, 301 : 재배송접수, 302 : 재배송완료
						' rw SubNodes.getElementsByTagName("dlvCstRespnClf")(0).text
						' rw SubNodes.getElementsByTagName("dlvEtprsCd")(0).text					'수거택배사코드 | 00034 : CJ대한통운, 00011 : 한진택배 등등...매뉴얼참고
						' rw SubNodes.getElementsByTagName("dlvNo")(0).text
						' rw SubNodes.getElementsByTagName("exchBaseAddr")(0).text					'교환상품 수령지 기본주소
						' rw SubNodes.getElementsByTagName("exchDtlsAddr")(0).text					'교환상품 수령지 상세주소
						' rw SubNodes.getElementsByTagName("exchMailNo")(0).text					'교환상품 수령지 우편번호
						' rw SubNodes.getElementsByTagName("exchMailNoSeq")(0).text					'교환상품 수령지 우편번호 순번
						' rw SubNodes.getElementsByTagName("exchNm")(0).text						'교환상품 수령지 이름
						' rw SubNodes.getElementsByTagName("exchPrtblTel")(0).text					'교환상품 수령지 휴대폰번호
						' rw SubNodes.getElementsByTagName("exchTlphnNo")(0).text					'교환상품 수령지 전화번호
						' rw SubNodes.getElementsByTagName("exchTypeAdd")(0).text					'교환상품 수령지 주소 유형 | 01 : 지번명, 02 : 도로명
						' rw SubNodes.getElementsByTagName("exchTypeBilNo")(0).text					'교환상품 수령지 건물관리번호
						' rw SubNodes.getElementsByTagName("freeGiftNo")(0).text
						' rw SubNodes.getElementsByTagName("freeGiftQty")(0).text
						' rw SubNodes.getElementsByTagName("kglUseYn")(0).text
						' rw SubNodes.getElementsByTagName("optName")(0).text						'옵션명
						' rw SubNodes.getElementsByTagName("ordNm")(0).text							'수거지 이름
						OutMallOrderSerial = SubNodes.getElementsByTagName("ordNo")(0).text			'11번가 주문번호
						OrgDetailKey = SubNodes.getElementsByTagName("ordPrdSeq")(0).text			'주문순번
						' rw SubNodes.getElementsByTagName("ordPrtblTel")(0).text					'수거지 휴대폰번호
						' rw SubNodes.getElementsByTagName("ordTlphnNo")(0).text					'수거지 전화번호
						' rw SubNodes.getElementsByTagName("prdNo")(0).text							'상품번호
						' rw SubNodes.getElementsByTagName("rcvrBaseAddr")(0).text					'수거지 기본주소
						' rw SubNodes.getElementsByTagName("rcvrDtlsAddr")(0).text					'수거지 상세주소
						' rw SubNodes.getElementsByTagName("rcvrMailNo")(0).text					'수거지 우편번호
						' rw SubNodes.getElementsByTagName("rcvrMailNoSeq")(0).text					'수거지 우편번호 순번
						' rw SubNodes.getElementsByTagName("rcvrTypeAdd")(0).text					'수거지 주소 유형 01 : 지번명, 02 : 도로명
						' rw SubNodes.getElementsByTagName("rcvrTypeBilNo")(0).text					'수거지 건물관리번호
						OutMallRegDate = LEFT(SubNodes.getElementsByTagName("reqDt")(0).text,10)	'클레임 요청 일시
						' rw SubNodes.getElementsByTagName("twMthd")(0).text						'교환상품 발송방법 | 02 : 직접발송, 06 : 11번가 지정반품택배, 07 : 판매자 지정반품택배,
						' rw SubNodes.getElementsByTagName("twPrdInvcNo")(0).text					'수거송장번호

						strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
						strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
						strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
						strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '', '', '', "
						strSql = strSql & "		'', '', '', '', '', '' "
						strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
						strSql = strSql & " END "
						strSql = strSql & " ELSE "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
						strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
						strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
						strSql = strSql & " END "
						dbget.Execute strSql,iAssignedRow

						if (iAssignedRow > 0) then
							iInputCnt = iInputCnt+iAssignedRow
							strSql = " update c "
							strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
							strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
							strSql = strSql + " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
							strSql = strSql + " from "
							strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
							strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
							strSql = strSql + " on "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.SellSite = o.SellSite "
							strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
							strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
							strSql = strSql + " where "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.orderserial is NULL "
							strSql = strSql + " 	and o.orderserial is not NULL "
							strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
							''response.write strSql & "<br />"
							dbget.Execute strSql
						end if
					Next
				Set Nodes = nothing
			Set xmlDOM = nothing
		End If
	SET objXML = nothing
	rw "CS 교환 입력건수:"&iInputCnt
'	response.end
End Function

Function GetCSOrderReturn_11st1010(sellsite, csGubun, csSubGubun, selldate)
	dim objXML, xmlDOM, obj, iRbody
	dim i, j, k
	dim startdate, enddate
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt, beasongNum11st
	dim strSql, ResultCode, SubNodes
	Dim dlvCnclYn
	Dim APIkey, Nodes, clmReqRsn, ordCnStatCd
	APIkey = "a2319e071dbc304243ee60abd07e9664"
	startdate = Left(DateAdd("d", 0, selldate), 10)
	startdate = Replace(startdate, "-", "") & "0000"

	enddate = Left(DateAdd("d", 0, selldate), 10)
	enddate = Replace(enddate, "-", "") & "2359"
	divcd = "A004"
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://api.11st.co.kr/rest/claimservice/returnorders/"&startdate&"/"&enddate, false
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
'rw iRbody
				Set Nodes = xmlDOM.getElementsByTagName("ns2:order")
					iInputCnt = 0
					For each SubNodes in Nodes
						' rw SubNodes.getElementsByTagName("addDlvCst")(0).text						'추가배송비
						' rw SubNodes.getElementsByTagName("affliateBndlDlvSeq")(0).text			'무료반품 여부 | 현재 미사용 필드  0 : 무료반품, 1 : 일반반품(유료)
						' rw SubNodes.getElementsByTagName("appmtDlvCst")(0).text					'11번가 지정반품 택배비
						' rw SubNodes.getElementsByTagName("clmDlvCstMthd")(0).text					'결제방법 | 01 : 신용카드, 02 : 포인트, 03 : 박스에 동봉, 04 : 판매자에게 직접송금, null : 클레임 사유가 판매자일 경우
						' rw SubNodes.getElementsByTagName("clmLstDlvCst")(0).text					'반품배송비
						' rw SubNodes.getElementsByTagName("clmReqCont")(0).text					'사유코드에 대한 상세내역
						itemno = SubNodes.getElementsByTagName("clmReqQty")(0).text					'클레임 수량
						clmReqRsn = SubNodes.getElementsByTagName("clmReqRsn")(0).text				'클레임 사유코드
						select case clmReqRsn
							case "101"
								gubunname = "단순변심"
							case "112"
								gubunname = "배송지연"
							case "111", "207", "122", "123"
								gubunname = "상품불량"
							case "108", "208"
								gubunname = "오배송"
							case "105", "210"
								gubunname = "상품정보틀림"
							case "110", "114", "115", "116", "117", "118", "119", "121", "198", "199", "206", "209", "212", "213", "113", "211", "214"
								gubunname = "기타"
							case else
								gubunname = Replace(gubunname, "'", "")
						end select

						CSDetailKey = SubNodes.getElementsByTagName("clmReqSeq")(0).text			'외부몰 클레임 번호
						' rw SubNodes.getElementsByTagName("clmStat")(0).text						'클레임 상태 | 103 : 재결제대기중, 104 : 반품보류, 105 : 반품신청, 106 : 반품완료, 107 : 반품거부, 108 : 반품철회, 109 : 반품완료보류, 201 : 교환신청, 212 : 교환승인, 214 : 교환보류, 221 : 교환발송완료, 232 : 교환거부, 233 : 교환철회, 301 : 재배송접수, 302 : 재배송완료
						' rw SubNodes.getElementsByTagName("dlvCstRespnClf")(0).text				'배송비 부담여부 | 01 : 구매자, 02 : 판매자
						' rw SubNodes.getElementsByTagName("dlvEtprsCd")(0).text					'수거택배사코드 | 00034 : CJ대한통운, 00011 : 한진택배 등등...매뉴얼참고
						beasongNum11st = SubNodes.getElementsByTagName("dlvNo")(0).text				'배송번호
						' rw SubNodes.getElementsByTagName("exchBaseAddr")(0).text
						' rw SubNodes.getElementsByTagName("exchDtlsAddr")(0).text
						' rw SubNodes.getElementsByTagName("exchMailNo")(0).text
						' rw SubNodes.getElementsByTagName("exchMailNoSeq")(0).text
						' rw SubNodes.getElementsByTagName("exchNm")(0).text
						' rw SubNodes.getElementsByTagName("exchPrtblTel")(0).text
						' rw SubNodes.getElementsByTagName("exchTlphnNo")(0).text
						' rw SubNodes.getElementsByTagName("exchTypeAdd")(0).text
						' rw SubNodes.getElementsByTagName("exchTypeBilNo")(0).text
						' rw SubNodes.getElementsByTagName("freeGiftNo")(0).text
						' rw SubNodes.getElementsByTagName("freeGiftQty")(0).text
						' rw SubNodes.getElementsByTagName("kglUseYn")(0).text						'KGL(해외배송) 택배사용유무 기본값 &#39;N&#39; 해외배송 상품만 해당됨
						' rw SubNodes.getElementsByTagName("optName")(0).text						'옵션명
						' rw SubNodes.getElementsByTagName("ordNm")(0).text							'수거지 이름
						OutMallOrderSerial = SubNodes.getElementsByTagName("ordNo")(0).text			'11번가 주문번호
						OrgDetailKey = SubNodes.getElementsByTagName("ordPrdSeq")(0).text			'주문순번
						' rw SubNodes.getElementsByTagName("ordPrtblTel")(0).text					'수거지 휴대폰번호
						' rw SubNodes.getElementsByTagName("ordTlphnNo")(0).text					'수거지 전화번호
						' rw SubNodes.getElementsByTagName("prdNo")(0).text							'상품번호
						' rw SubNodes.getElementsByTagName("rcvrBaseAddr")(0).text					'수거지 기본주소
						' rw SubNodes.getElementsByTagName("rcvrDtlsAddr")(0).text					'수거지 상세주소
						' rw SubNodes.getElementsByTagName("rcvrMailNo")(0).text					'수거지 우편번호
						' rw SubNodes.getElementsByTagName("rcvrMailNoSeq")(0).text					'수거지 우편번호 순번
						' rw SubNodes.getElementsByTagName("rcvrTypeAdd")(0).text					'수거지 주소 유형 | 01 : 지번명, 02 : 도로명
						' rw SubNodes.getElementsByTagName("rcvrTypeBilNo")(0).text					'수거지 건물관리번호
						OutMallRegDate = LEFT(SubNodes.getElementsByTagName("reqDt")(0).text,10)	'클레임 요청 일시
						' rw SubNodes.getElementsByTagName("twMthd")(0).text						'반품상품 발송방법 | 02 : 직접발송, 06 : 11번가 지정반품택배, 07 : 판매자 지정반품택배, 09 : 판매처 수거
						' rw SubNodes.getElementsByTagName("twPrdInvcNo")(0).text					'수거송장번호

						strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
						strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
						strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
						strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '', '', '', "
						strSql = strSql & "		'', '', '', '', '', '' "
						strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
						strSql = strSql & " END "
						strSql = strSql & " ELSE "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
						strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
						strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
						strSql = strSql & " END "
						'rw strSql
						dbget.Execute strSql,iAssignedRow

						if (iAssignedRow > 0) then
							iInputCnt = iInputCnt+iAssignedRow
							strSql = " update c "
							strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
							strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
							strSql = strSql + " , c.OrderName = o.OrderName "
							strSql = strSql + " from "
							strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
							strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
							strSql = strSql + " on "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.SellSite = o.SellSite "
							strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
							strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
							strSql = strSql + " where "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.orderserial is NULL "
							strSql = strSql + " 	and o.orderserial is not NULL "
							strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
							dbget.Execute strSql
						End If
					Next
				Set Nodes = nothing
			Set xmlDOM = nothing
		End If
	SET objXML = nothing
	rw "CS 반품 입력건수:"&iInputCnt
'	response.end
End Function

Function GetCSOrderCancel_hmall(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst, iRbody, strObj
	dim objXML, xmlDOM, objArr, obj
	dim i, j, k
	dim startdate, enddate, retCode, strObjValues, strObjValuesReturnItems
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno, OrderName, OrderHpNo
	dim iAssignedRow, iInputCnt
	dim strSql, ResultCode, ResultMsg, cancelOrderList, SubNodes
	Dim ClaimSeq, ClaimMemo, RegYMD, dlvCnclYn, orderCount, prgrGb

	startdate = Replace(Left(DateAdd("d", 0, selldate), 10), "-", "")
	divcd = "A008"
	For k = 0 To 1
		If k = 0 Then
			prgrGb = "P0"
		Else
			prgrGb = "P1"
		End If
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			'prgrGb | P0:출고대기, P1:출고진행, P2:출고, P3:배송완료
			objXML.open "GET", "http://xapi.10x10.co.kr:8080/Orders/Hmall/output?startdate="&startdate&"&enddate="&startdate&"&prgrGb="&prgrGb, false
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.Send()

			If objXML.Status = "200" OR objXML.Status = "201" Then
				iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
				Dim obj1
				Set strObj = JSON.parse(iRbody)
					orderCount = strObj.count
					If orderCount > 0 Then
						set obj1 = strObj.lstorder
							for i=0 to obj1.length-1
								CSDetailKey			= obj1.get(i).dlvstNo				'배송지시번호
								OutMallOrderSerial	= obj1.get(i).ordNo					'주문번호
								OrgDetailKey		= obj1.get(i).ordPtcSeq				'주문일련번호
								dlvCnclYn			= obj1.get(i).dlvCnclYn				'배송취소여부
								gubunname			= obj1.get(i).dlvCnclNm				'배송취소
								itemno				= obj1.get(i).custCnclQty			'고객취소수량
								OutMallRegDate 		= LEFT(obj1.get(i).ptcOrdDtm, 10)		'상세주문일자

								If dlvCnclYn = "Y" Then	'배송취소여부가 Y
									strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
									strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
									strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
									strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
									strSql = strSql & "		'', '', '', '', '', '' "
									strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
									strSql = strSql & " END "
									strSql = strSql & " ELSE "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
									strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
									strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
									strSql = strSql & " END "
									'rw strSql
									dbget.Execute strSql,iAssignedRow
								End If
							Next
						set obj1 = nothing
						strSql = " update c "
						strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
						strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
						strSql = strSql + " , c.OrderName = o.OrderName "
						strSql = strSql + " from "
						strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c with (nolock) "
						strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o with (nolock) "
						strSql = strSql + " on "
						strSql = strSql + " 	1 = 1 "
						strSql = strSql + " 	and c.SellSite = o.SellSite "
						strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
						strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
						strSql = strSql + " where "
						strSql = strSql + " 	1 = 1 "
						strSql = strSql + " 	and c.orderserial is NULL "
						strSql = strSql + " 	and o.orderserial is not NULL "
						strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "						
						dbget.Execute strSql

						If divcd = "A008" Then
							strSql = " update c "
							strSql = strSql + " set c.currstate = 'B007' "
							strSql = strSql + " from "
							strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c with (nolock) "
							strSql = strSql + " left join db_temp.dbo.tbl_xSite_TMPOrder o with (nolock) "
							strSql = strSql + " on "
							strSql = strSql + "		1 = 1 "
							strSql = strSql + "		and c.SellSite = o.SellSite "
							strSql = strSql + "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
							strSql = strSql + "		and c.OrgDetailKey = o.OrgDetailKey "
							strSql = strSql + " where "
							strSql = strSql + "		1 = 1 "
							strSql = strSql + "		and c.orderserial is NULL "
							strSql = strSql + "		and o.SellSite is NULL "
							strSql = strSql + "		and c.currstate = 'B001' "
							strSql = strSql + "		and c.divcd = 'A008' "
							strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "						
							''rw strSql
							dbget.execute strSql
						end if
					End If
				Set strObj = nothing
			End If
		Set objXML= nothing
	Next
End Function

Function GetCSOrderReturn_hmall(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst, iRbody, strObj
	dim objXML, xmlDOM, objArr, obj
	dim i, j, k
	dim startdate, enddate, retCode, strObjValues, strObjValuesReturnItems
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno, OrderName, OrderHpNo
	dim iAssignedRow, iInputCnt, dlvTypeGbcd
	dim strSql, ResultCode, ResultMsg, cancelOrderList, SubNodes
	Dim ClaimSeq, ClaimMemo, RegYMD, dlvCnclYn, orderCount

	startdate = Replace(Left(DateAdd("d", 0, selldate), 10), "-", "")

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		'prgrGb | P0:출고대기, P1:출고진행, P2:출고, P3:배송완료
		objXML.open "GET", "http://xapi.10x10.co.kr:8080/Orders/Hmall/return?startdate="&startdate&"&enddate="&startdate&"&prgrGb=P0", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Dim obj1
			Set strObj = JSON.parse(iRbody)
				orderCount = strObj.count
				If orderCount > 0 Then
					set obj1 = strObj.lstorder
						for i=0 to obj1.length-1
							CSDetailKey			= obj1.get(i).dlvstNo				'배송지시번호
							OutMallOrderSerial	= obj1.get(i).ordNo					'주문번호
							OrgDetailKey		= obj1.get(i).befDlvstNo			'이전 배송지시번호
							dlvTypeGbcd			= obj1.get(i).dlvTypeGbcd			'배송유형(주문구분) | 30:반품회수, 45:교환회수, 65:부분교환회수
							gubunname			= obj1.get(i).custVenPaonMsg		'고객협력사전달메시지
							itemno				= obj1.get(i).prrgQty				'대상수량 | 배송지시수량-취소수량-배송수량
							OutMallRegDate 		= LEFT(obj1.get(i).oshpReqnDt, 10)	'출고요청일

							Select Case dlvTypeGbcd
								Case "30"		divcd = "A004"
								Case "45"		divcd = "A000"
								Case "65"		divcd = "A000"
							End Select

							If itemno = "" Then itemno = "0"

							If dlvTypeGbcd <> "" Then	'배송유형(주문구분)에 값이 있다면
								strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
								strSql = strSql & " BEGIN "
								strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
								strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
								strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
								strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
								strSql = strSql & "		'', '', '', '', '', '' "
								strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
								strSql = strSql & " END "
								strSql = strSql & " ELSE "
								strSql = strSql & " BEGIN "
								strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
								strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
								strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
								strSql = strSql & " END "
								'rw strSql
								dbget.Execute strSql,iAssignedRow
							End If
						Next
					set obj1 = nothing
					strSql = " update c "
					strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
					strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
					strSql = strSql + " , c.OrderName = o.OrderName "
					strSql = strSql + " from "
					strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c with (nolock) "
					strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o with (nolock) "
					strSql = strSql + " on "
					strSql = strSql + " 	1 = 1 "
					strSql = strSql + " 	and c.SellSite = o.SellSite "
					strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
					strSql = strSql + " 	and c.OrgDetailKey = o.beasongNum11st "
					strSql = strSql + " where "
					strSql = strSql + " 	1 = 1 "
					strSql = strSql + " 	and c.orderserial is NULL "
					strSql = strSql + " 	and o.orderserial is not NULL "
					strSql = strSql + " 	and c.sellsite = 'hmall1010' "
					dbget.Execute strSql
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

Function GetCSOrderCS_WMP(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst, objXML, xmlDOM, strObj
	dim startdate, enddate
	dim retCode, retMsg, items, item, product, productOption, divcd, gubunname
	dim OutMallOrderSerial, CSDetailKey, OrgDetailKey, dlvTypeGbcd, itemno, OutMallRegDate
	dim i, j, k, canceDoneStr
	dim strSql, iAssignedRow, successCount

	startdate = selldate
	enddate = selldate

	If csGubun = "CANCELDONE" Then
		canceDoneStr = "CANCELDONE"
		xmlURL = "http://110.93.128.100:8090/Wemake/Orders/cleiminquiry?startdate=" & startdate & "%2000%3A00%3A00&enddate=" & enddate & "%2023%3A59%3A59&type=CANCEL&status=APPROVE&searchDateType=REQUEST"
	Else
		xmlURL = "http://110.93.128.100:8090/Wemake/Orders/cleiminquiry?startdate=" & startdate & "%2000%3A00%3A00&enddate=" & enddate & "%2023%3A59%3A59&type=" & csGubun & "&status=REQUEST&searchDateType=REQUEST"
	End If

	strRst = ""

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.send(strRst)

	if objXML.Status <> "200" then
rw xmlURL
rw "-------"
rw objXML.responseText
		response.write "ERROR : 통신오류"
		dbget.close : response.end
	end if

	Set strObj = JSON.parse(objXML.responseText)
	retMsg = strObj.message

	successCount = 0

	if (retMsg = "성공") then
		Set items = strObj.outPutValue.data.claim
		For i = 0 to items.length - 1
			Set item = items.get(i)

			OutMallOrderSerial = item.bundleNo
			CSDetailKey = item.claimBundleNo
			dlvTypeGbcd = item.claimType
			OutMallRegDate = item.requestDate
			gubunname = LEFT(item.claimReason, 10)

			Select Case dlvTypeGbcd
				Case "교환"		divcd = "A000"
				Case "취소"		divcd = "A008"
				Case "반품"		divcd = "A004"
			End Select

			for j = 0 to item.orderProduct.length - 1
				Set product = item.orderProduct.get(j)

				for k = 0 to product.orderOption.length - 1
					Set productOption = product.orderOption.get(k)

					OrgDetailKey = productOption.orderOptionNo
					itemno = productOption.optionQty

					strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
					strSql = strSql & " BEGIN "
					strSql = strSql & "		INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
					strSql = strSql & "		OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
					strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
					strSql = strSql & "		('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
					strSql = strSql & "		'', '', '', '', '', '' "
					strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
					strSql = strSql & " END "
					''rw strSql
					dbget.Execute strSql,iAssignedRow

					if (iAssignedRow > 0) then
						successCount = successCount + iAssignedRow
					end if
				next
			next

			strSql = " update c "
			strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
			strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
			strSql = strSql + " , c.OrderName = o.OrderName "
			strSql = strSql + " from "
			strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
			strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
			strSql = strSql + " on "
			strSql = strSql + "		1 = 1 "
			strSql = strSql + "		and c.SellSite = o.SellSite "
			strSql = strSql + "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
			strSql = strSql + "		and c.OrgDetailKey = o.OrgDetailKey "
			strSql = strSql + " where "
			strSql = strSql + "		1 = 1 "
			strSql = strSql + "		and c.OutMallOrderSerial = '" & OutMallOrderSerial & "' "
			strSql = strSql + "		and c.orderserial is NULL "
			strSql = strSql + "		and o.orderserial is not NULL "
			strSql = strSql + "		and c.sellsite = '" & sellsite & "' "
			dbget.Execute strSql
		next

		If canceDoneStr = "CANCELDONE" Then
			rw "CS " & canceDoneStr & " 입력건수:" & successCount
		Else
			rw "CS " & csGubun & " 입력건수:" & successCount
		End If
	else
		response.write "ERROR : " & retMsg
	end if
end function

Function GetCSOrderCS_wmpfashion(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst, objXML, xmlDOM, strObj
	dim startdate, enddate
	dim retCode, retMsg, items, item, product, productOption, divcd, gubunname
	dim OutMallOrderSerial, CSDetailKey, OrgDetailKey, dlvTypeGbcd, itemno, OutMallRegDate
	dim i, j, k
	dim strSql, iAssignedRow, successCount

	startdate = selldate
	enddate = selldate

	xmlURL = "http://110.93.128.100:8090/fwmp/Orders/cleiminquiry?startdate=" & startdate & "%2000%3A00%3A00&enddate=" & enddate & "%2023%3A59%3A59&type=" & csGubun & "&status=REQUEST&searchDateType=REQUEST"

	strRst = ""

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.send(strRst)

	if objXML.Status <> "200" then
		response.write "ERROR : 통신오류"
	end if

	Set strObj = JSON.parse(objXML.responseText)
	retMsg = strObj.message

	successCount = 0

	if (retMsg = "성공") then
		Set items = strObj.outPutValue.data.claim
		For i = 0 to items.length - 1
			Set item = items.get(i)

			OutMallOrderSerial = item.bundleNo
			CSDetailKey = item.claimBundleNo
			dlvTypeGbcd = item.claimType
			OutMallRegDate = item.requestDate
			gubunname = LEFT(item.claimReason, 10)

			Select Case dlvTypeGbcd
				Case "교환"		divcd = "A000"
				Case "취소"		divcd = "A008"
				Case "반품"		divcd = "A004"
			End Select

			for j = 0 to item.orderProduct.length - 1
				Set product = item.orderProduct.get(j)

				for k = 0 to product.orderOption.length - 1
					Set productOption = product.orderOption.get(k)

					OrgDetailKey = productOption.orderOptionNo
					itemno = productOption.optionQty

					strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
					strSql = strSql & " BEGIN "
					strSql = strSql & "		INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
					strSql = strSql & "		OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
					strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
					strSql = strSql & "		('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
					strSql = strSql & "		'', '', '', '', '', '' "
					strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
					strSql = strSql & " END "
					''rw strSql
					dbget.Execute strSql,iAssignedRow

					if (iAssignedRow > 0) then
						successCount = successCount + iAssignedRow
					end if
				next
			next

			strSql = " update c "
			strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
			strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
			strSql = strSql + " , c.OrderName = o.OrderName "
			strSql = strSql + " from "
			strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
			strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
			strSql = strSql + " on "
			strSql = strSql + "		1 = 1 "
			strSql = strSql + "		and c.SellSite = o.SellSite "
			strSql = strSql + "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
			strSql = strSql + "		and c.OrgDetailKey = o.OrgDetailKey "
			strSql = strSql + " where "
			strSql = strSql + "		1 = 1 "
			strSql = strSql + "		and c.OutMallOrderSerial = '" & OutMallOrderSerial & "' "
			strSql = strSql + "		and c.orderserial is NULL "
			strSql = strSql + "		and o.orderserial is not NULL "
			strSql = strSql + "		and c.sellsite = '" & sellsite & "' "
			dbget.Execute strSql
		next

		rw "CS " & csGubun & " 입력건수:" & successCount
	else
		response.write "ERROR : " & retMsg
	end if
End Function

CONST UPCHECODE = "A5703"								'업체코드
CONST APIKEY = "B6D75816-1F35-4450-8B9B-71137B9212F9"	'API KEY
Function GetCSOrderCancel_halfclub(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj
	dim i
	dim startdate, enddate
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql, ResultCode, ResultMsg, cancelOrderList, SubNodes
	Dim ClaimSeq, ClaimMemo, RegYMD

	startdate = Left(DateAdd("d", 0, selldate), 10)
	startdate = Replace(startdate, "-", "") & "0000"

	enddate = Left(DateAdd("d", 1, selldate), 10)
	enddate = Replace(enddate, "-", "") & "0000"

	xmlURL = "http://api.tricycle.co.kr"

	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
	strRst = strRst & "	<soap12:Header>"
	strRst = strRst & "		<SOAPHeaderAuth xmlns=""http://api.tricycle.co.kr/"">"
	strRst = strRst & "			<User_ID>"&UPCHECODE&"</User_ID>"
	strRst = strRst & "			<User_PWD>"&APIKEY&"</User_PWD>"
	strRst = strRst & "		</SOAPHeaderAuth>"
	strRst = strRst & "	</soap12:Header>"
	strRst = strRst & "	<soap12:Body>"
	strRst = strRst & "		<Get_OrderCancel xmlns=""http://api.tricycle.co.kr/"">"
	strRst = strRst & "			<req_OrderCancel>"
	strRst = strRst & "				<FromYMD>"&startdate&"</FromYMD>"
	strRst = strRst & "				<ToYMD>"&enddate&"</ToYMD>"
	strRst = strRst & "			</req_OrderCancel>"
	strRst = strRst & "		</Get_OrderCancel>"
	strRst = strRst & "	</soap12:Body>"
	strRst = strRst & "</soap12:Envelope>"
'	response.write replace(strRst, "utf-8","euc-kr")
'	response.end

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objXML.open "POST", "" & xmlURL&"/Claim/Claim.asmx"
		objXML.setRequestHeader "Host", "api.tricycle.co.kr"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strRst)
		objXML.setRequestHeader "SOAPMethodName", "Get_OrderCancel"
		objXML.send(strRst)
		if objXML.Status <> "200" then
			response.write "ERROR : 통신오류"
			dbget.close : response.end
		end if
        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
'response.write replace(objXML.responseText, "utf-8","euc-kr")
'response.end
			iInputCnt = 0
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "Get_OrderCancelResponse" Then
				ResultCode	= xmlDOM.getElementsByTagName ("Get_OrderCancelResult ").item(0).attributes(0).nodeValue
				ResultMsg	= xmlDOM.getElementsByTagName ("Get_OrderCancelResult ").item(0).attributes(1).nodeValue
			End If

			If ResultCode = "0000" Then
				Set cancelOrderList = xmlDOM.getElementsByTagName("Response_Order_Cancel")
					For Each SubNodes in cancelOrderList
						OutMallOrderSerial	= SubNodes.getElementsByTagName("OrdNum")(0).Text				'하프클럽 주문번호
						OrgDetailKey		= SubNodes.getElementsByTagName("OrdNum_Nm")(0).Text			'하프클로 주문순번
						itemno				= SubNodes.getElementsByTagName("Qty")(0).Text					'취소 수량
						CSDetailKey			= SubNodes.getElementsByTagName("ClaimSeq")(0).Text				'취소 등록 고유 번호
						ClaimMemo			= SubNodes.getElementsByTagName("ClaimMemo")(0).Text			'취소 등록 사유
						OutMallRegDate		= LEFT(SubNodes.getElementsByTagName("RegYMD")(0).Text, 10)		'취소 일자

						divcd				= "A008"
						gubunname 			= Replace(LEFT(ClaimMemo, 128), "'", "")

				        strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
						strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
						strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
						strSql = strSql & " 	('" & divcd & "', convert(varchar(128), '" & gubunname & "'), '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
						strSql = strSql & "		'', '', '', '', '', '' "
						strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
						strSql = strSql & " END "
						strSql = strSql & " ELSE "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
						strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
						strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
						strSql = strSql & " END "
						dbget.Execute strSql,iAssignedRow
						'response.write strSql & "<br />"

						if (iAssignedRow > 0) then
							iInputCnt = iInputCnt+iAssignedRow

							'' CS 마스터정보. 업데이트
							strSql = " update c "
							strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
							strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
							strSql = strSql + " , c.OrderName = o.OrderName "
							strSql = strSql + " from "
							strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
							strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
							strSql = strSql + " on "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.SellSite = o.SellSite "
							strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
							strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
							strSql = strSql + " where "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.orderserial is NULL "
							strSql = strSql + " 	and o.orderserial is not NULL "
							strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
							'response.write strSql & "<br />"
							dbget.Execute strSql

							If divcd = "A008" Then
								strSql = " update c "
								strSql = strSql + " set c.currstate = 'B007' "
								strSql = strSql + " from "
								strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
								strSql = strSql + " left join db_temp.dbo.tbl_xSite_TMPOrder o "
								strSql = strSql + " on "
								strSql = strSql + " 	1 = 1 "
								strSql = strSql + " 	and c.SellSite = o.SellSite "
								strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
								strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
								strSql = strSql + " where "
								strSql = strSql + " 	1 = 1 "
								strSql = strSql + " 	and c.orderserial is NULL "
								strSql = strSql + " 	and o.SellSite is NULL "
								strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
								strSql = strSql + " 	and c.currstate = 'B001' "
								strSql = strSql + " 	and c.divcd = 'A008' "
								''rw strSql
								dbget.execute strSql
							end if
						end if
					Next
				Set cancelOrderList = nothing
			End If
        Set xmlDOM = nothing
	Set objXML = nothing
	rw "취소 CS입력건수:"&iInputCnt
End Function

Function GetCSOrderReturn_halfclub(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj
	dim i
	dim startdate, enddate
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql, ResultCode, ResultMsg, refundOrderList, SubNodes
	Dim ClaimSeq, ClaimReasonCd, ClaimReason, RegYMD

	startdate = Left(DateAdd("d", 0, selldate), 10)
	startdate = Replace(startdate, "-", "") & "0000"

	enddate = Left(DateAdd("d", 1, selldate), 10)
	enddate = Replace(enddate, "-", "") & "0000"

	xmlURL = "http://api.tricycle.co.kr"

	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
	strRst = strRst & "	<soap12:Header>"
	strRst = strRst & "		<SOAPHeaderAuth xmlns=""http://api.tricycle.co.kr/"">"
	strRst = strRst & "			<User_ID>"&UPCHECODE&"</User_ID>"
	strRst = strRst & "			<User_PWD>"&APIKEY&"</User_PWD>"
	strRst = strRst & "		</SOAPHeaderAuth>"
	strRst = strRst & "	</soap12:Header>"
	strRst = strRst & "	<soap12:Body>"
	strRst = strRst & "		<Get_OrderRefund xmlns=""http://api.tricycle.co.kr/"">"
	strRst = strRst & "			<req_orderRefund>"
	strRst = strRst & "				<FromYMD>"&startdate&"</FromYMD>"
	strRst = strRst & "				<ToYMD>"&enddate&"</ToYMD>"
	strRst = strRst & "			</req_orderRefund>"
	strRst = strRst & "		</Get_OrderRefund>"
	strRst = strRst & "	</soap12:Body>"
	strRst = strRst & "</soap12:Envelope>"
'	response.write replace(strRst, "utf-8","euc-kr")
'	response.end

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objXML.open "POST", "" & xmlURL&"/Claim/Claim.asmx"
		objXML.setRequestHeader "Host", "api.tricycle.co.kr"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strRst)
		objXML.setRequestHeader "SOAPMethodName", "Get_OrderRefund"
		objXML.send(strRst)
		if objXML.Status <> "200" then
			response.write "ERROR : 통신오류"
		end if
        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
'response.write replace(objXML.responseText, "utf-8","euc-kr")
'response.end
			iInputCnt = 0
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "Get_OrderCancelResponse" Then
				ResultCode	= xmlDOM.getElementsByTagName ("Get_OrderRefundResult ").item(0).attributes(0).nodeValue
				ResultMsg	= xmlDOM.getElementsByTagName ("Get_OrderRefundResult ").item(0).attributes(1).nodeValue
			End If

			If ResultCode = "0000" Then
				Set refundOrderList = xmlDOM.getElementsByTagName("Response_Order_Refund")
					For Each SubNodes in refundOrderList
						OutMallOrderSerial	= SubNodes.getElementsByTagName("OrdNum")(0).Text				'하프클럽 주문번호
						OrgDetailKey		= SubNodes.getElementsByTagName("OrdNum_Nm")(0).Text			'하프클로 주문순번
						itemno				= SubNodes.getElementsByTagName("Qty")(0).Text					'반품 수량
						CSDetailKey			= SubNodes.getElementsByTagName("ClaimSeq")(0).Text				'반품 등록 고유 번호
						OutMallRegDate		= LEFT(SubNodes.getElementsByTagName("RegYMD")(0).Text, 10)		'반품 등록일
						ClaimReasonCd		= SubNodes.getElementsByTagName("Claim_ReasonCd")(0).Text		'반품 사유 코드
						ClaimReason			= SubNodes.getElementsByTagName("Claim_Reason")(0).Text			'반품 사유

						select case ClaimReasonCd
							case "1"
								gubunname = "불량"
							case "2"
								gubunname = "사이즈"
							case "3"
								gubunname = "선별오류"
							case "4"
								gubunname = "상품정보불일치"
							case "5"
								gubunname = "고객변심"
							case "6"
								gubunname = "기타"
							case "7"
								gubunname = "추가세일"
							case else
								gubunname = Replace(gubunname, "'", "")
						end select
						divcd				= "A004"
						gubunname 			= Replace(LEFT(ClaimReason, 128), "'", "")

				        strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
						strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
						strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
						strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
						strSql = strSql & "		'', '', '', '', '', '' "
						strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
						strSql = strSql & " END "
						strSql = strSql & " ELSE "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
						strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
						strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
						strSql = strSql & " END "
						dbget.Execute strSql,iAssignedRow
						'response.write strSql & "<br />"

						if (iAssignedRow > 0) then
							iInputCnt = iInputCnt+iAssignedRow

							'' CS 마스터정보. 업데이트
							strSql = " update c "
							strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
							strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
							strSql = strSql + " , c.OrderName = o.OrderName "
							strSql = strSql + " from "
							strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
							strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
							strSql = strSql + " on "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.SellSite = o.SellSite "
							strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
							strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
							strSql = strSql + " where "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.orderserial is NULL "
							strSql = strSql + " 	and o.orderserial is not NULL "
							strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
							'response.write strSql & "<br />"
							dbget.Execute strSql
						end if
					Next
				Set cancelOrderList = nothing
			End If
        Set xmlDOM = nothing
	Set objXML = nothing
	rw "반품 CS입력건수:"&iInputCnt
End Function

'// 222222222222
function GetCSOrderAll_nvstorefarm(sellsite, csGubun, csSubGubun, selldate)
	If sellsite = "nvstorefarm" Then
		sellsite = "nvstorefarm"
	ElseIf sellsite = "nvstoregift" Then
		sellsite = "nvstoregift"
	ElseIf sellsite = "Mylittlewhoopee" Then
		sellsite = "Mylittlewhoopee"
	Else
		sellsite = "nvstoremoonbangu"
	End If
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj, objData, objDetail, objDetailArr
	dim i, j, k
	dim startdate, enddate
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd
	dim reqID, keyGenerated, cryptoLib
	dim ResponseType

	objArr = GetOrderDetailList_nvstorefarm(selldate, csGubun, sellsite)

	response.write "건수(" & UBound(objArr) + 1 & ") " & "<br />"

	if UBound(objArr) < 0 then
		exit function
	end if

	iServ		= "SellerService41"
	iCcd		= "GetProductOrderInfoList"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If
	''response.write xmlURL

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If sellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf sellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">" + vbCrLf
	strRst = strRst & "	<soapenv:Header/>" + vbCrLf
	strRst = strRst & "	<soapenv:Body>" + vbCrLf
	strRst = strRst & "		<sel:GetProductOrderInfoListRequest>" + vbCrLf
	strRst = strRst & "			<sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>" + vbCrLf
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>" + vbCrLf
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>" + vbCrLf
	strRst = strRst & "			</sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>" + vbCrLf
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>" + vbCrLf
	strRst = strRst & "			<sel:Version>4.1</sel:Version>" + vbCrLf
	For each obj in objArr
		strRst = strRst & "			<sel:ProductOrderIDList>" & obj & "</sel:ProductOrderIDList>" + vbCrLf
	next
	strRst = strRst & "		</sel:GetProductOrderInfoListRequest>" + vbCrLf
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	''response.write strRst

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		response.write "ERROR : 통신오류"
		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
	If session("ssBctID")="kjy8517" Then
		response.write "<textarea cols=100 rows=30>" & objXML.responseText & "</textarea>"
	End If
	''response.write objXML.responseText & "<br /><br />"
	''response.flush

	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		response.write "오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	if CLng(xmlDOM.getElementsByTagName("n:ReturnedDataCount").item(0).text) <> (UBound(objArr) + 1) then
		response.write "건수 불일치 오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	set objArr = xmlDOM.getElementsByTagName("n:ProductOrderInfoList")

	iInputCnt = 0
	For each obj in objArr
		OutMallOrderSerial = obj.selectSingleNode("n:Order/n:OrderID").text
		OrgDetailKey = obj.selectSingleNode("n:ProductOrder/n:ProductOrderID").text
		CSDetailKey = obj.selectSingleNode("n:ProductOrder/n:ClaimID").text
		'if (csGubun = "CANCELED") then
		if (csGubun = "CANCEL_REQUESTED") OR (csGubun = "CANCELED") then	'2019-07-23 김진영 CANCEL_REQUESTED로 변경
			divcd = "A008"
			OutMallRegDate = Left(obj.selectSingleNode("n:CancelInfo/n:ClaimRequestDate").text, 10)

			gubunname = obj.selectSingleNode("n:CancelInfo/n:CancelReason").text
			select case gubunname
				case "INTENT_CHANGED"
					gubunname = "단순변심"
				case "PRODUCT_UNSATISFIED"
					gubunname = "상품불만족"
				case "SOLD_OUT"
					gubunname = "품절"
				case "COLOR_AND_SIZE"
					gubunname = "색상사이즈"
				case "WRONG_ORDER"
					gubunname = "주문오류"
				case else
					gubunname = Replace(gubunname, "'", "")
			end select
		'elseif (csGubun = "RETURNED") then
		elseif (csGubun = "RETURN_REQUESTED") then	'2019-07-23 김진영 RETURN_REQUESTED로 변경
			divcd = "A004"
			OutMallRegDate = Left(obj.selectSingleNode("n:ReturnInfo/n:ClaimRequestDate").text, 10)

			gubunname = obj.selectSingleNode("n:ReturnInfo/n:ReturnReason").text
			select case gubunname
				case "INTENT_CHANGED"
					gubunname = "단순변심"
				case "PRODUCT_UNSATISFIED"
					gubunname = "상품불만족"
				case "SOLD_OUT"
					gubunname = "품절"
				case "COLOR_AND_SIZE"
					gubunname = "색상사이즈"
				case "WRONG_ORDER"
					gubunname = "주문오류"
				case else
					gubunname = Replace(gubunname, "'", "")
			end select
		elseif (csGubun = "EXCHANGE_REQUESTED") then	'2022-03-03 김진영 EXCHANGE_REQUESTED추가
			divcd = "A000"
			OutMallRegDate = Left(obj.selectSingleNode("n:ExchangeInfo/n:ClaimRequestDate").text, 10)
			gubunname = obj.selectSingleNode("n:ExchangeInfo/n:ExchangeReason").text
			select case gubunname
				case "INTENT_CHANGED"
					gubunname = "단순변심"
				case "PRODUCT_UNSATISFIED"
					gubunname = "상품불만족"
				case "SOLD_OUT"
					gubunname = "품절"
				case "COLOR_AND_SIZE"
					gubunname = "색상사이즈"
				case "WRONG_ORDER"
					gubunname = "주문오류"
				case else
					gubunname = Replace(gubunname, "'", "")
			end select
		elseif (csGubun = "EXCHANGED") then
			divcd = 1/0			'// 에러
			divcd = "A000"
			OutMallRegDate = "1900-01-01"
		else
			divcd = 1/0			'// 에러
		end if

		itemno = obj.selectSingleNode("n:ProductOrder/n:Quantity").text

        strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
		strSql = strSql & " BEGIN "
		strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
		strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
		strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
		strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
		strSql = strSql & "		'', '', '', '', '', '' "
		strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
		strSql = strSql & " END "
		strSql = strSql & " ELSE "
		strSql = strSql & " BEGIN "
		strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
		strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
		strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
		strSql = strSql & " END "
		dbget.Execute strSql,iAssignedRow
		''response.write strSql & "<br />"

		if (iAssignedRow > 0) then
			iInputCnt = iInputCnt+iAssignedRow

			'' CS 마스터정보. 업데이트
			strSql = " update c "
			strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
			strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
			strSql = strSql + " , c.OrderName = o.OrderName "
			strSql = strSql + " from "
			strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
			strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
			strSql = strSql + " on "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.SellSite = o.SellSite "
			strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
			strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
			strSql = strSql + " where "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.orderserial is NULL "
			strSql = strSql + " 	and o.orderserial is not NULL "
			strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
			''response.write strSql & "<br />"
			dbget.Execute strSql

			If divcd = "A008" Then
				strSql = " update c "
				strSql = strSql + " set c.currstate = 'B007' "
				strSql = strSql + " from "
				strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
				strSql = strSql + " left join db_temp.dbo.tbl_xSite_TMPOrder o "
				strSql = strSql + " on "
				strSql = strSql + "		1 = 1 "
				strSql = strSql + "		and c.SellSite = o.SellSite "
				strSql = strSql + "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
				strSql = strSql + "		and c.OrgDetailKey = o.OrgDetailKey "
				strSql = strSql + " where "
				strSql = strSql + "		1 = 1 "
				strSql = strSql + "		and c.orderserial is NULL "
				strSql = strSql + "		and o.SellSite is NULL "
				strSql = strSql + "		and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
				strSql = strSql + "		and c.currstate = 'B001' "
				strSql = strSql + "		and c.divcd = 'A008' "
				''rw strSql
				dbget.execute strSql
			end if
		end if
	next

	if (csGubun = "CANCELED") then
		rw "취소 CS입력건수:"&iInputCnt
	elseif (csGubun = "RETURNED") then
		rw "반품 CS입력건수:"&iInputCnt
	elseif (csGubun = "EXCHANGE_REQUESTED") then
		rw "교환 CS입력건수:"&iInputCnt
	end if

end function

function GetOrderDetailList_nvstorefarm(selldate, LastChangedStatusCode, isellsite)
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd, reqID, ResponseType
	dim xmlURL
	dim strRst, objXML, xmlDOM
	dim objMasterListXML, objMasterOneXML
	dim PrdOrderList(), i
	dim tmpXml

	redim PrdOrderList(-1)
	GetOrderDetailList_nvstorefarm = PrdOrderList

	iServ		= "SellerService41"
	iCcd		= "GetChangedProductOrderList"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	'// =======================================================================
	'// API URL(기간동안의 주문 가져오기)
	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If
	''response.write xmlURL

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If isellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf isellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">"
	strRst = strRst & "	<soapenv:Header/>"
	strRst = strRst & "	<soapenv:Body>"
	strRst = strRst & "		<sel:GetChangedProductOrderListRequest>"
	strRst = strRst & "			<sel:AccessCredentials>"
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>"
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>"
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>"
	strRst = strRst & "			</sel:AccessCredentials>"
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>"
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>"															'#돌려받는 데이터의 상세 정도(Compact / Full)
	strRst = strRst & "			<sel:Version>4.1</sel:Version>"
	strRst = strRst & "			<sel:InquiryTimeFrom>"&selldate&"T00:00:00</sel:InquiryTimeFrom>"									'#조회 시작 일시(해당 시각 포함)
	strRst = strRst & "			<sel:InquiryTimeTo>"& Left(DateAdd("d", 1, CDate(selldate)), 10)&"T00:00:00</sel:InquiryTimeTo>"	'조회 종료 일시(해당 시각 포함하지 않음)
	strRst = strRst & "			<sel:LastChangedStatusCode>" & LastChangedStatusCode & "</sel:LastChangedStatusCode>"				'최종 상품 주문 상태 코드 (CANCELED | 취소, RETURNED | 반품, EXCHANGED : 교환 | PAYED : 결제완료)
	strRst = strRst & "			<sel:MallID>"&reqID&"</sel:MallID>"																	'판매자 아이디
	strRst = strRst & "		</sel:GetChangedProductOrderListRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	''response.write strRst
	''dbget.close : response.end

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		response.write "ERROR : 통신오류"
		dbget.close : response.end
	end if


	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
	''response.write objXML.responseText & "<br /><br />"

	''dbget.close : response.end

	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		response.write "오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	if CLng(xmlDOM.getElementsByTagName("n:ReturnedDataCount").item(0).text) = 0 then
		response.write "내역없음<br />"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	set objMasterListXML = xmlDOM.getElementsByTagName("n:ChangedProductOrderInfoList")

	i = 0
	redim PrdOrderList(objMasterListXML.length - 1)
	For each objMasterOneXML in objMasterListXML
		PrdOrderList(i) = objMasterOneXML.getElementsByTagName("n:ProductOrderID")(0).Text
		i = i + 1
	next

	GetOrderDetailList_nvstorefarm = PrdOrderList

	set objMasterListXML = Nothing
	Set xmlDOM = Nothing
	Set objXML = Nothing
end function

''1300k 주문취소 조회
Function GetCSOrderCancel_wetoo1300k(sellsite, csGubun, csSubGubun, selldate)
	Dim wetoo1300kAPIURL, company_auth, company_code
	Dim xmlURL, strRst, xmlSelldate
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate, OrgOutMallOrderSerial, orderCancelList
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete, claimType, claimList
	Dim addParam, returnStatus, iMessage
	Dim stDate, edDate, strParam
	Dim CSDetailKeySub
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")

	If application("Svr_Info") = "Dev" Then
		wetoo1300kAPIURL = "https://ts.1300k.com"
		company_auth = "1ac6e7cd04fc587cc26722b1cbaaa75c"
		company_code = "C927"
	Else
		wetoo1300kAPIURL = "http://api.1300k.com"
		company_auth = "f91f60a59e32425e4f22c3d20cf4f7b7"
		company_code = "C927"
	End If

	stDate = Replace(selldate, "-", "")
	edDate = Replace(Dateadd("d", 1, selldate), "-", "")

	Set obj = jsObject()
		Set obj("header") = jsObject()
			obj("header")("company_code") = company_code
			obj("header")("company_auth") = company_auth
			Set obj("claim_date") = jsObject()
				obj("claim_date")("st_date") = stDate&"0000"		'#시작시간 YYYYMMDDHHMM
				obj("claim_date")("ed_date") = edDate&"0000"		'#종료시간 YYYYMMDDHHMM
			strParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", wetoo1300kAPIURL & "/enterstore/api/v2/order_claim.html", false
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
rw iRbody
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "00" Then
					Set claimList = strObj.result.order_claim
						If claimList.length > 0 Then
							For i=0 to claimList.length-1
								CSDetailKey = Trim(claimList.get(i).claim_no)				'클레임번호
								CSDetailKeySub = Trim(claimList.get(i).claim_sub_no)		'클레임SUB번호
								CSDetailKey = CSDetailKey & "-" & CSDetailKeySub

								If InStr(Trim(claimList.get(i).claim_type), "취소") >0 Then
									divcd = "A008"
								ElseIf InStr(Trim(claimList.get(i).claim_type), "반품") >0 Then
									divcd = "A004"
								ElseIf InStr(Trim(claimList.get(i).claim_type), "교환") >0 Then
									divcd = "A000"
								ElseIf InStr(Trim(claimList.get(i).claim_type), "회수") >0 Then
									divcd = "A200"
								End If

'								Trim(claimList.get(i).claim_result)		'STATUS
								OutMallRegDate = Trim(claimList.get(i).claim_request_date)	'클레임요청시간
'								Trim(claimList.get(i).status_change_date)	'상태변경시간
								gubunname = Trim(claimList.get(i).reason)				'클레임사유
'								Trim(claimList.get(i).remark)				'고객코맨트
								OutMallOrderSerial = Trim(claimList.get(i).order_no)			'주문번호
								' Trim(claimList.get(i).order_name)			'주문자명
								OrgDetailKey = Trim(claimList.get(i).seq_no)				'일련번호
								' Trim(claimList.get(i).product_code)		'상품코드
								' Trim(claimList.get(i).product_name)		'상품명
								' Trim(claimList.get(i).opt_no)				'옵션코드
								' Trim(claimList.get(i).opt_name)			'옵션명
								itemno = Trim(claimList.get(i).qty)				'수량
								' Trim(claimList.get(i).company_product_code)	'업체상품코드
								' Trim(claimList.get(i).company_opt_no)		'업체옵션번호

								strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
								strSql = strSql & " BEGIN "
								strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
								strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
								strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
								strSql = strSql & " 	('" & divcd & "', '"& html2db(gubunname) &"', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
								strSql = strSql & "		'', '', '', '', '', '' "
								strSql = strSql & "		, '"& Replace(OutMallRegDate, "/", "-") &"', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
								strSql = strSql & " END "
								dbget.Execute strSql, iAssignedRow

								if (iAssignedRow > 0) then
									iInputCnt = iInputCnt+iAssignedRow
									If divcd = "A008" Then
										''주문 입력 이전 내역은 삭제 하자
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET matchState = 'D'"
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder c "
										strSql = strSql & " WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										strSql = strSql & " and orderserial is NULL"
										dbget.Execute strSql
									End If

									'' CS 마스터정보. 업데이트
									strSql = ""
									strSql = strSql & " UPDATE c "
									strSql = strSql & " SET c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
									strSql = strSql & " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
									strSql = strSql & " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
									strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
									strSql = strSql & " JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
									strSql = strSql & " ON "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.SellSite = o.SellSite "
									strSql = strSql & " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
									strSql = strSql & " 	and c.OrgDetailKey = o.OrgDetailKey "
									strSql = strSql & " WHERE "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.orderserial is NULL "
									strSql = strSql & " 	and o.orderserial is not NULL "
									strSql = strSql & " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
									''response.write strSql & "<br />"
									dbget.Execute strSql

									If divcd = "A008" Then
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET c.currstate = 'B007' "
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
										strSql = strSql & " LEFT JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
										strSql = strSql & " ON "
										strSql = strSql & "		1 = 1 "
										strSql = strSql & "		and c.SellSite = o.SellSite "
										strSql = strSql & "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
										strSql = strSql & "		and c.OrgDetailKey = o.OrgDetailKey "
										strSql = strSql & " WHERE "
										strSql = strSql & "		1 = 1 "
										strSql = strSql & "		and c.orderserial is NULL "
										strSql = strSql & "		and o.SellSite is NULL "
										strSql = strSql & "		and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										strSql = strSql & "		and c.currstate = 'B001' "
										strSql = strSql & "		and c.divcd = 'A008' "
										''rw strSql
										dbget.execute strSql
									End If
								End If
							Next
						End If
					Set claimList = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

''1300k 주문취소 조회
Function GetCSOrderCancel2_wetoo1300k(sellsite, csGubun, csSubGubun, selldate)
	Dim wetoo1300kAPIURL, company_auth, company_code
	Dim xmlURL, strRst, xmlSelldate
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate, OrgOutMallOrderSerial, orderCancelList
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete, claimType, claimList
	Dim addParam, returnStatus, iMessage
	Dim stDate, edDate, strParam
	Dim CSDetailKeySub
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")

	If application("Svr_Info") = "Dev" Then
		wetoo1300kAPIURL = "https://ts.1300k.com"
		company_auth = "1ac6e7cd04fc587cc26722b1cbaaa75c"
		company_code = "C927"
	Else
		wetoo1300kAPIURL = "http://api.1300k.com"
		company_auth = "f91f60a59e32425e4f22c3d20cf4f7b7"
		company_code = "C927"
	End If

	stDate = Replace(selldate, "-", "")
	edDate = Replace(Dateadd("d", 1, selldate), "-", "")

	Set obj = jsObject()
		Set obj("header") = jsObject()
			obj("header")("company_code") = company_code
			obj("header")("company_auth") = company_auth
			Set obj("cancel_date") = jsObject()
				obj("cancel_date")("st_date") = stDate&"0000"		'#시작시간 YYYYMMDDHHMM
				obj("cancel_date")("ed_date") = edDate&"0000"		'#종료시간 YYYYMMDDHHMM
			strParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", wetoo1300kAPIURL & "/enterstore/api/order_cancel.html", false
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
 rw iRbody
' response.end
				If returnCode = "00" Then
					If IsObject(strObj.result.cancel) Then
					Set claimList = strObj.result.cancel
						If claimList.length > 0 Then
							For i=0 to claimList.length-1
								CSDetailKey = ""
								divcd = "A008"
								OutMallRegDate = Trim(claimList.get(i).cancel_date)	'클레임요청시간
								OutMallRegDate = Left(OutMallRegDate,4)&"-"&Mid(OutMallRegDate,5,2)&"-"&Mid(OutMallRegDate,7,2)&" "&Mid(OutMallRegDate,9,2)&":"&Mid(OutMallRegDate,11,2)&":"&Mid(OutMallRegDate,13,2)
								OutMallOrderSerial = Trim(claimList.get(i).order_no)			'주문번호
								OrgDetailKey = Trim(claimList.get(i).seq_no)				'일련번호
								itemno = Trim(claimList.get(i).qty)				'수량

								strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
								strSql = strSql & " BEGIN "
								strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
								strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
								strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
								strSql = strSql & " 	('" & divcd & "', '', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
								strSql = strSql & "		'', '', '', '', '', '' "
								strSql = strSql & "		, '"& Replace(OutMallRegDate, "/", "-") &"', '" & CStr(OrgDetailKey) & "', '', " & itemno & ") "
								strSql = strSql & " END "
								dbget.Execute strSql, iAssignedRow

								if (iAssignedRow > 0) then
									iInputCnt = iInputCnt+iAssignedRow
									If divcd = "A008" Then
										''주문 입력 이전 내역은 삭제 하자
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET matchState = 'D'"
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder c "
										strSql = strSql & " WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										strSql = strSql & " and orderserial is NULL"
										dbget.Execute strSql
									End If

									'' CS 마스터정보. 업데이트
									strSql = ""
									strSql = strSql & " UPDATE c "
									strSql = strSql & " SET c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
									strSql = strSql & " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
									strSql = strSql & " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
									strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
									strSql = strSql & " JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
									strSql = strSql & " ON "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.SellSite = o.SellSite "
									strSql = strSql & " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
									strSql = strSql & " 	and c.OrgDetailKey = o.OrgDetailKey "
									strSql = strSql & " WHERE "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.orderserial is NULL "
									strSql = strSql & " 	and o.orderserial is not NULL "
									strSql = strSql & " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
									''response.write strSql & "<br />"
									dbget.Execute strSql

									If divcd = "A008" Then
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET c.currstate = 'B007' "
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
										strSql = strSql & " LEFT JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
										strSql = strSql & " ON "
										strSql = strSql & "		1 = 1 "
										strSql = strSql & "		and c.SellSite = o.SellSite "
										strSql = strSql & "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
										strSql = strSql & "		and c.OrgDetailKey = o.OrgDetailKey "
										strSql = strSql & " WHERE "
										strSql = strSql & "		1 = 1 "
										strSql = strSql & "		and c.orderserial is NULL "
										strSql = strSql & "		and o.SellSite is NULL "
										strSql = strSql & "		and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										strSql = strSql & "		and c.currstate = 'B001' "
										strSql = strSql & "		and c.divcd = 'A008' "
										''rw strSql
										dbget.execute strSql
									End If
								End If
							Next
						End If
					Set claimList = nothing
					End If
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

''1300k 반품 조회
Function GetCSOrderReturn_wetoo1300k(sellsite, csGubun, csSubGubun, selldate)
	Dim wetoo1300kAPIURL, company_auth, company_code
	Dim xmlURL, strRst, xmlSelldate
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate, OrgOutMallOrderSerial, orderCancelList
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete, claimType, claimList
	Dim addParam, returnStatus, iMessage
	Dim stDate, edDate, strParam
	Dim CSDetailKeySub
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")

	If application("Svr_Info") = "Dev" Then
		wetoo1300kAPIURL = "https://ts.1300k.com"
		company_auth = "1ac6e7cd04fc587cc26722b1cbaaa75c"
		company_code = "C927"
	Else
		wetoo1300kAPIURL = "http://api.1300k.com"
		company_auth = "f91f60a59e32425e4f22c3d20cf4f7b7"
		company_code = "C927"
	End If

	stDate = Replace(selldate, "-", "")
	edDate = Replace(Dateadd("d", 1, selldate), "-", "")

	Set obj = jsObject()
		Set obj("header") = jsObject()
			obj("header")("company_code") = company_code
			obj("header")("company_auth") = company_auth
			Set obj("return_date") = jsObject()
				obj("return_date")("st_date") = stDate&"0000"		'#시작시간 YYYYMMDDHHMM
				obj("return_date")("ed_date") = edDate&"0000"		'#종료시간 YYYYMMDDHHMM
			strParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", wetoo1300kAPIURL & "/enterstore/api/order_return.html", false
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
 rw iRbody
' response.end
				If returnCode = "00" Then
					If IsObject(strObj.result.return) Then
					Set claimList = strObj.result.return
						If claimList.length > 0 Then
							For i=0 to claimList.length-1
								CSDetailKey = ""
								divcd = "A004"
								OutMallRegDate = Trim(claimList.get(i).return_date)	'클레임요청시간
								OutMallRegDate = Left(OutMallRegDate,4)&"-"&Mid(OutMallRegDate,5,2)&"-"&Mid(OutMallRegDate,7,2)&" "&Mid(OutMallRegDate,9,2)&":"&Mid(OutMallRegDate,11,2)&":"&Mid(OutMallRegDate,13,2)
								OutMallOrderSerial = Trim(claimList.get(i).order_no)			'주문번호
								OrgDetailKey = Trim(claimList.get(i).seq_no)				'일련번호
								itemno = Trim(claimList.get(i).qty)				'수량

								strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
								strSql = strSql & " BEGIN "
								strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
								strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
								strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
								strSql = strSql & " 	('" & divcd & "', '', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
								strSql = strSql & "		'', '', '', '', '', '' "
								strSql = strSql & "		, '"& Replace(OutMallRegDate, "/", "-") &"', '" & CStr(OrgDetailKey) & "', '', " & itemno & ") "
								strSql = strSql & " END "
								dbget.Execute strSql, iAssignedRow

								if (iAssignedRow > 0) then
									iInputCnt = iInputCnt+iAssignedRow
									'' CS 마스터정보. 업데이트
									strSql = ""
									strSql = strSql & " UPDATE c "
									strSql = strSql & " SET c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
									strSql = strSql & " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
									strSql = strSql & " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
									strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
									strSql = strSql & " JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
									strSql = strSql & " ON "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.SellSite = o.SellSite "
									strSql = strSql & " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
									strSql = strSql & " 	and c.OrgDetailKey = o.OrgDetailKey "
									strSql = strSql & " WHERE "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.orderserial is NULL "
									strSql = strSql & " 	and o.orderserial is not NULL "
									strSql = strSql & " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
									''response.write strSql & "<br />"
									dbget.Execute strSql
								End If
							Next
						End If
					Set claimList = nothing
					End If
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

''skstoa 주문취소 조회
Function GetCSOrderCancel_skstoa(sellsite, csGubun, csSubGubun, selldate)
	Dim xmlURL, strRst, xmlSelldate
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate, OrgOutMallOrderSerial, orderCancelList
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete, claimType
	Dim addParam, returnStatus, iMessage
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")

	addParam = ""
	addParam = addParam & "linkCode=" & skstoalinkCode		'#연결코드 | SKB에서 부여한 연결코드
	addParam = addParam & "&entpCode=" & skstoaentpCode		'#업체코드 | SKB에서 부여한 업체코드 6자리
	addParam = addParam & "&entpId=" & skstoaentpId			'#업체사용자ID | SKB에서 부여한 업체사용자 ID
	addParam = addParam & "&entpPass=" & skstoaentpPass		'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
	addParam = addParam & "&bDate="& xmlSelldate			'#조회 시작일자 | 최종 처리일 기준 YYYYMMDD 타입. ex) 20140520
	addParam = addParam & "&eDate="& xmlSelldate			'#조회 마지막일자 | 최종 처리일 기준 YYYYMMDD 타입. ex) 20140520
'	addParam = addParam & "&orderNo="						'주문코드 | 주문번호를 이용한 검색. 숫자만 허용되며 사용시 14자리(orderNo) 또는 23자리(orderNo,orderGSeq,orderDSeq,orderWSeq)의 값만 허용
'rw addParam

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", skstoaAPIURL & "/partner/delivery/order-cancel-list?" & addParam , false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnStatus	= strObj.status
				returnCode		= strObj.code
				iMessage		= strObj.message
				rw "==========="
				rw BinaryToText(objXML.ResponseBody,"utf-8")
				rw "==========="
				If returnCode = "200" Then
					Set orderCancelList = strObj.orderCancelList
						If orderCancelList.length > 0 Then
							For i=0 to orderCancelList.length-1
								divcd				= "A008"
								CSDetailKey			= ""
								OutMallOrderSerial	= orderCancelList.get(i).orderNo			'주문번호
								OrgDetailKey		= orderCancelList.get(i).orderGSeq & "-" & orderCancelList.get(i).orderDSeq & "-" & orderCancelList.get(i).orderWSeq	'상품순번 - 세트순번 - 처리순번
								OutMallRegDate		= orderCancelList.get(i).orderDate			'주문접수일
								OutMallRegDate = Left(OutMallRegDate, 4) & "-" & Mid(OutMallRegDate, 5, 2) & "-" & Right(OutMallRegDate, 2)
								' rw orderCancelList.get(i).goodsCode			'판매상품코드
								' rw orderCancelList.get(i).goodsName			'판매상품명
								' rw orderCancelList.get(i).goodsdtCode		'판매단품코드
								' rw orderCancelList.get(i).goodsdtInfo		'판매단품정보
								' rw orderCancelList.get(i).goodsGb			'상품구분
								itemno				= orderCancelList.get(i).orderQty			'원주문수량
								' rw orderCancelList.get(i).salePrice			'개별 판매가
								' rw orderCancelList.get(i).buyPrice			'개별 매입가
								' rw orderCancelList.get(i).custName			'주문자명
'								gubunname			= orderCancelList.get(i).msg				'배송메시지
								gubunname = "취소"
								strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
								strSql = strSql & " BEGIN "
								strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
								strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
								strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
								strSql = strSql & " 	('" & divcd & "', '"& html2db(gubunname) &"', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
								strSql = strSql & "		'', '', '', '', '', '' "
								strSql = strSql & "		, '"& Replace(OutMallRegDate, "/", "-") &"', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
								strSql = strSql & " END "
								dbget.Execute strSql, iAssignedRow

								if (iAssignedRow > 0) then
									iInputCnt = iInputCnt+iAssignedRow
									''주문 입력 이전 내역은 삭제 하자
									strSql = ""
									strSql = strSql & " UPDATE c "
									strSql = strSql & " SET matchState = 'D'"
									strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder c "
									strSql = strSql & " WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
									strSql = strSql & " and orderserial is NULL"
									dbget.Execute strSql

									'' CS 마스터정보. 업데이트
									strSql = ""
									strSql = strSql & " UPDATE c "
									strSql = strSql & " SET c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
									strSql = strSql & " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
									strSql = strSql & " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
									strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
									strSql = strSql & " JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
									strSql = strSql & " ON "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.SellSite = o.SellSite "
									strSql = strSql & " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
									strSql = strSql & " 	and c.OrgDetailKey = o.OrgDetailKey "
									strSql = strSql & " WHERE "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.orderserial is NULL "
									strSql = strSql & " 	and o.orderserial is not NULL "
									strSql = strSql & " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
									''response.write strSql & "<br />"
									dbget.Execute strSql

									If divcd = "A008" Then
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET c.currstate = 'B007' "
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
										strSql = strSql & " LEFT JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
										strSql = strSql & " ON "
										strSql = strSql & "		1 = 1 "
										strSql = strSql & "		and c.SellSite = o.SellSite "
										strSql = strSql & "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
										strSql = strSql & "		and c.OrgDetailKey = o.OrgDetailKey "
										strSql = strSql & " WHERE "
										strSql = strSql & "		1 = 1 "
										strSql = strSql & "		and c.orderserial is NULL "
										strSql = strSql & "		and o.SellSite is NULL "
										strSql = strSql & "		and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										strSql = strSql & "		and c.currstate = 'B001' "
										strSql = strSql & "		and c.divcd = 'A008' "
										''rw strSql
										dbget.execute strSql
									End If
								End If
							Next
						End If
					Set orderCancelList = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

''skstoa 반품교환 회수 대상조회
Function GetCSOrderReturnExchange_skstoa(sellsite, csGubun, csSubGubun, selldate)
	Dim xmlURL, strRst, xmlSelldate
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate, OrgOutMallOrderSerial, returnList
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete, claimType
	Dim addParam, returnStatus, iMessage
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")
	addParam = ""
	addParam = addParam & "linkCode=" & skstoalinkCode		'#연결코드 | SKB에서 부여한 연결코드
	addParam = addParam & "&entpCode=" & skstoaentpCode		'#업체코드 | SKB에서 부여한 업체코드 6자리
	addParam = addParam & "&entpId=" & skstoaentpId			'#업체사용자ID | SKB에서 부여한 업체사용자 ID
	addParam = addParam & "&entpPass=" & skstoaentpPass		'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
	addParam = addParam & "&claimGb=00" 					'회수 구분 값 | 00:전체(default),30:반품회수,45:교환회수
	addParam = addParam & "&bDate="& xmlSelldate			'#조회 시작일자 | 주문승인/교환접수일 기준, YYYYMMDD 타입. ex) 20140520
	addParam = addParam & "&eDate="& xmlSelldate			'#조회 마지막일자 | 주문승인/교환접수일 기준 YYYYMMDD 타입. ex) 20140520
'	addParam = addParam & "&orderNo="						'주문코드 | 주문번호를 이용한 검색. 숫자만 허용되며 사용시 14자리(orderNo) 또는 23자리(orderNo,orderGSeq,orderDSeq,orderWSeq)의 값만 허용
'rw addParam

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", skstoaAPIURL & "/partner/return/return-list?" & addParam , false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnStatus	= strObj.status
				returnCode		= strObj.code
				iMessage		= strObj.message
				rw "==========="
				rw BinaryToText(objXML.ResponseBody,"utf-8")
				rw "==========="
				If returnCode = "200" Then
					Set returnList = strObj.returnList
						If returnList.length > 0 Then
							For i=0 to returnList.length-1
								IsDelete			= "N"
								claimType			= ""
								CSDetailKey			= ""
								If returnList.get(i).claimGb = "30" Then				'회수 구분 값
									divcd = "A004"
								Else
									divcd = "A000"
								End If
								'rw returnList.get(i).claimGbName						'회수 구분 명
								OutMallOrderSerial	= returnList.get(i).orderNo			'주문번호
								OrgDetailKey		= returnList.get(i).orderGSeq & "-" & returnList.get(i).orderDSeq & "-" & returnList.get(i).orderWSeq	'상품순번 - 세트순번 - 처리순번
								'rw returnList.get(i).goodsGb							'상품구분
								'rw returnList.get(i).goodsCode							'판매상품코드
								'rw returnList.get(i).goodsName							'판매상품명
								'rw returnList.get(i).goodsdtCode						'판매단품코드
								'rw returnList.get(i).goodsdtInfo						'판매단품정보
								OutMallRegDate		= returnList.get(i).returnProcDate	'회수지시일
								itemno				= returnList.get(i).syslast			'배송수량
								'rw returnList.get(i).salePrice							'판매가
								'rw returnList.get(i).buyPrice							'매입가(정산예정금액)
								'rw returnList.get(i).shipCostName						'배송비정책명
								'rw returnList.get(i).custName							'주문자명
								'rw returnList.get(i).receiverName						'수취인명
								'rw returnList.get(i).receiverPostNo					'우편번호
								'rw returnList.get(i).receiverTel						'연락처
								'rw returnList.get(i).receiverHp						'휴대폰

								On Error Resume Next
									gubunname			= returnList.get(i).msg			'반품상세사유
									If Err.number <> 0 Then
										gubunname = "사유없음"
									End If
								On Error Goto 0
								gubunname = LEFT(gubunname, 50)

								strSql = ""
								strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
								strSql = strSql & " BEGIN "
								strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
								strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
								strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
								strSql = strSql & " 	('" & divcd & "', '"& html2db(gubunname) &"', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
								strSql = strSql & "		'', '', '', '', '', '' "
								strSql = strSql & "		, '"& Replace(OutMallRegDate, "/", "-") &"', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
								strSql = strSql & " END "
								dbget.Execute strSql, iAssignedRow

								strSql = ""
								strSql = strSql & " UPDATE db_temp.dbo.tbl_xSite_TMPCS "
								strSql = strSql & " SET OutMallCurrState = 'B008' "		''?? 
								strSql = strSql & " WHERE "
								strSql = strSql & " 	1 = 1 "
								strSql = strSql & " 	and SellSite = '" & sellsite & "' "
								strSql = strSql & " 	and OutMallOrderSerial = '" & OutMallOrderSerial & "' "
								strSql = strSql & " 	and CSDetailKey = '" & CSDetailKey & "' "
								strSql = strSql & " 	and OrgDetailKey = '" & OrgDetailKey & "' "
								strSql = strSql & " 	and divcd = '" & divcd & "' "
								dbget.Execute strSql

								If (iAssignedRow > 0) Then
									iInputCnt = iInputCnt+iAssignedRow

									'' CS 마스터정보. 업데이트
									strSql = ""
									strSql = strSql & " UPDATE c "
									strSql = strSql & " SET c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
									strSql = strSql & " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
									strSql = strSql & " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
									strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
									strSql = strSql & " JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
									strSql = strSql & " ON "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.SellSite = o.SellSite "
									strSql = strSql & " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
									strSql = strSql & " 	and LEFT(c.OrgDetailKey, 7) = LEFT(o.OrgDetailKey, 7) "		'신세계만 LEFT 7로 조인해야됨..마지막 3자리가 안 맞음
									strSql = strSql & " WHERE "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.orderserial is NULL "
									strSql = strSql & " 	and o.orderserial is not NULL "
									strSql = strSql & " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' "
									strSql = strSql & " 	and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
									dbget.Execute strSql
								End If
							Next
						End If
					Set returnList = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

''shintvshopping 주문취소 조회
Function GetCSOrderCancel_shintvshopping(sellsite, csGubun, csSubGubun, selldate)
	Dim xmlURL, strRst, xmlSelldate
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate, OrgOutMallOrderSerial, orderCancelList
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete, claimType
	Dim addParam, returnStatus, iMessage
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")

	addParam = ""
	addParam = addParam & "linkCode=" & linkCode			'#연결코드 | 샵링커: SLINK [ TCODE.LGROUP : xxx ]
	addParam = addParam & "&entpCode=" & entpCode			'#업체코드 | 신세계TV쇼핑에서 부여한 업체코드 6자리
	addParam = addParam & "&entpId=" & entpId				'#업체사용자ID | 신세계TV쇼핑에서 부여한 업체사용자 ID
	addParam = addParam & "&entpPass=" & entpPass			'#업체PASSWORD | 신세계TV쇼핑에서 등록한 업체사용자 비밀번호
	addParam = addParam & "&bDate="& xmlSelldate			'#조회 시작일자 | 주문승인/교환접수일 기준, YYYYMMDD 타입. ex) 20140520
	addParam = addParam & "&eDate="& xmlSelldate			'#조회 마지막일자 | 주문승인/교환접수일 기준 YYYYMMDD 타입. ex) 20140520
'	addParam = addParam & "&orderNo="						'주문코드 | 주문번호를 이용한 검색. 숫자만 허용되며 사용시 14자리(orderNo) 또는 23자리(orderNo,orderGSeq,orderDSeq,orderWSeq)의 값만 허용
'rw addParam
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", shintvshoppingAPIURL & "/partner/delivery/order-cancel-list?" & addParam , false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,800000,800000,800000
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnStatus	= strObj.status
				returnCode		= strObj.code
				iMessage		= strObj.message
				rw "==========="
				rw BinaryToText(objXML.ResponseBody,"utf-8")
				rw "==========="
				If returnCode = "200" Then
					Set orderCancelList = strObj.orderCancelList
						If orderCancelList.length > 0 Then
							For i=0 to orderCancelList.length-1
								divcd				= "A008"
								CSDetailKey			= ""
								OutMallOrderSerial	= orderCancelList.get(i).orderNo			'주문번호
								OrgDetailKey		= orderCancelList.get(i).orderGSeq & "-" & orderCancelList.get(i).orderDSeq & "-" & orderCancelList.get(i).orderWSeq	'상품순번 - 세트순번 - 처리순번
								OutMallRegDate		= orderCancelList.get(i).orderDate			'주문접수일
								OutMallRegDate = Left(OutMallRegDate, 4) & "-" & Mid(OutMallRegDate, 5, 2) & "-" & Right(OutMallRegDate, 2)
								' rw orderCancelList.get(i).goodsCode			'판매상품코드
								' rw orderCancelList.get(i).goodsName			'판매상품명
								' rw orderCancelList.get(i).goodsdtCode		'판매단품코드
								' rw orderCancelList.get(i).goodsdtInfo		'판매단품정보
								' rw orderCancelList.get(i).goodsGb			'상품구분
								itemno				= orderCancelList.get(i).orderQty			'원주문수량
								' rw orderCancelList.get(i).salePrice			'개별 판매가
								' rw orderCancelList.get(i).buyPrice			'개별 매입가
								' rw orderCancelList.get(i).custName			'주문자명
								gubunname			= orderCancelList.get(i).msg				'배송메시지

								If ISNULL(gubunname) OR LEN(gubunname) < 2 Then
									gubunname = "취소"
								End If
								strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
								strSql = strSql & " BEGIN "
								strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
								strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
								strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
								strSql = strSql & " 	('" & divcd & "', '"& html2db(gubunname) &"', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
								strSql = strSql & "		'', '', '', '', '', '' "
								strSql = strSql & "		, '"& Replace(OutMallRegDate, "/", "-") &"', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
								strSql = strSql & " END "
								dbget.Execute strSql, iAssignedRow

								if (iAssignedRow > 0) then
									iInputCnt = iInputCnt+iAssignedRow
									''주문 입력 이전 내역은 삭제 하자
									strSql = ""
									strSql = strSql & " UPDATE c "
									strSql = strSql & " SET matchState = 'D'"
									strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder c "
									strSql = strSql & " WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
									strSql = strSql & " and orderserial is NULL"
									dbget.Execute strSql

									'' CS 마스터정보. 업데이트
									strSql = ""
									strSql = strSql & " UPDATE c "
									strSql = strSql & " SET c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
									strSql = strSql & " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
									strSql = strSql & " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
									strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
									strSql = strSql & " JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
									strSql = strSql & " ON "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.SellSite = o.SellSite "
									strSql = strSql & " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
									strSql = strSql & " 	and c.OrgDetailKey = o.OrgDetailKey "
									strSql = strSql & " WHERE "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.orderserial is NULL "
									strSql = strSql & " 	and o.orderserial is not NULL "
									strSql = strSql & " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
									''response.write strSql & "<br />"
									dbget.Execute strSql

									If divcd = "A008" Then
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET c.currstate = 'B007' "
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
										strSql = strSql & " LEFT JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
										strSql = strSql & " ON "
										strSql = strSql & "		1 = 1 "
										strSql = strSql & "		and c.SellSite = o.SellSite "
										strSql = strSql & "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
										strSql = strSql & "		and c.OrgDetailKey = o.OrgDetailKey "
										strSql = strSql & " WHERE "
										strSql = strSql & "		1 = 1 "
										strSql = strSql & "		and c.orderserial is NULL "
										strSql = strSql & "		and o.SellSite is NULL "
										strSql = strSql & "		and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										strSql = strSql & "		and c.currstate = 'B001' "
										strSql = strSql & "		and c.divcd = 'A008' "
										''rw strSql
										dbget.execute strSql
									End If
								End If
							Next
						End If
					Set orderCancelList = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

''shintvshopping 반품교환 회수 대상조회
Function GetCSOrderReturnExchange_shintvshopping(sellsite, csGubun, csSubGubun, selldate)
	Dim xmlURL, strRst, xmlSelldate
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate, OrgOutMallOrderSerial, returnList
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete, claimType
	Dim addParam, returnStatus, iMessage
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")
	addParam = ""
	addParam = addParam & "linkCode=" & linkCode			'#연결코드 | 샵링커: SLINK [ TCODE.LGROUP : xxx ]
	addParam = addParam & "&entpCode=" & entpCode			'#업체코드 | 신세계TV쇼핑에서 부여한 업체코드 6자리
	addParam = addParam & "&entpId=" & entpId				'#업체사용자ID | 신세계TV쇼핑에서 부여한 업체사용자 ID
	addParam = addParam & "&entpPass=" & entpPass			'#업체PASSWORD | 신세계TV쇼핑에서 등록한 업체사용자 비밀번호
	addParam = addParam & "&claimGb=00" 					'회수 구분 값 | 00:전체(default),30:반품회수,45:교환회수
	addParam = addParam & "&bDate="& xmlSelldate			'#조회 시작일자 | 주문승인/교환접수일 기준, YYYYMMDD 타입. ex) 20140520
	addParam = addParam & "&eDate="& xmlSelldate			'#조회 마지막일자 | 주문승인/교환접수일 기준 YYYYMMDD 타입. ex) 20140520
'	addParam = addParam & "&orderNo="						'주문코드 | 주문번호를 이용한 검색. 숫자만 허용되며 사용시 14자리(orderNo) 또는 23자리(orderNo,orderGSeq,orderDSeq,orderWSeq)의 값만 허용
'rw addParam

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", shintvshoppingAPIURL & "/partner/return/return-list?" & addParam , false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,800000,800000,800000
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnStatus	= strObj.status
				returnCode		= strObj.code
				iMessage		= strObj.message
				rw "==========="
				rw BinaryToText(objXML.ResponseBody,"utf-8")
				rw "==========="
				If returnCode = "200" Then
					Set returnList = strObj.returnList
						If returnList.length > 0 Then
							For i=0 to returnList.length-1
								IsDelete			= "N"
								claimType			= ""
								CSDetailKey			= ""
								If returnList.get(i).claimGb = "30" Then				'회수 구분 값
									divcd = "A004"
								Else
									divcd = "A000"
								End If
								'rw returnList.get(i).claimGbName						'회수 구분 명
								OutMallOrderSerial	= returnList.get(i).orderNo			'주문번호
								OrgDetailKey		= returnList.get(i).orderGSeq & "-" & returnList.get(i).orderDSeq & "-" & returnList.get(i).orderWSeq	'상품순번 - 세트순번 - 처리순번
								'rw returnList.get(i).goodsGb							'상품구분
								'rw returnList.get(i).goodsCode							'판매상품코드
								'rw returnList.get(i).goodsName							'판매상품명
								'rw returnList.get(i).goodsdtCode						'판매단품코드
								'rw returnList.get(i).goodsdtInfo						'판매단품정보
								OutMallRegDate		= returnList.get(i).returnProcDate	'회수지시일
								itemno				= returnList.get(i).syslast			'배송수량
								'rw returnList.get(i).salePrice							'판매가
								'rw returnList.get(i).buyPrice							'매입가(정산예정금액)
								'rw returnList.get(i).shipCostName						'배송비정책명
								'rw returnList.get(i).custName							'주문자명
								'rw returnList.get(i).receiverName						'수취인명
								'rw returnList.get(i).receiverPostNo					'우편번호
								'rw returnList.get(i).receiverTel						'연락처
								'rw returnList.get(i).receiverHp						'휴대폰

								On Error Resume Next
									gubunname			= returnList.get(i).msg			'반품상세사유
									If Err.number <> 0 Then
										gubunname = "사유없음"
									End If
								On Error Goto 0
								gubunname = LEFT(gubunname, 50)

								strSql = ""
								strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
								strSql = strSql & " BEGIN "
								strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
								strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
								strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
								strSql = strSql & " 	('" & divcd & "', '"& html2db(gubunname) &"', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
								strSql = strSql & "		'', '', '', '', '', '' "
								strSql = strSql & "		, '"& Replace(OutMallRegDate, "/", "-") &"', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
								strSql = strSql & " END "
								dbget.Execute strSql, iAssignedRow

								strSql = ""
								strSql = strSql & " UPDATE db_temp.dbo.tbl_xSite_TMPCS "
								strSql = strSql & " SET OutMallCurrState = 'B008' "		''?? 
								strSql = strSql & " WHERE "
								strSql = strSql & " 	1 = 1 "
								strSql = strSql & " 	and SellSite = '" & sellsite & "' "
								strSql = strSql & " 	and OutMallOrderSerial = '" & OutMallOrderSerial & "' "
								strSql = strSql & " 	and CSDetailKey = '" & CSDetailKey & "' "
								strSql = strSql & " 	and OrgDetailKey = '" & OrgDetailKey & "' "
								strSql = strSql & " 	and divcd = '" & divcd & "' "
								dbget.Execute strSql

								If (iAssignedRow > 0) Then
									iInputCnt = iInputCnt+iAssignedRow

									'' CS 마스터정보. 업데이트
									strSql = ""
									strSql = strSql & " UPDATE c "
									strSql = strSql & " SET c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
									strSql = strSql & " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
									strSql = strSql & " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
									strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
									strSql = strSql & " JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
									strSql = strSql & " ON "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.SellSite = o.SellSite "
									strSql = strSql & " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
									strSql = strSql & " 	and LEFT(c.OrgDetailKey, 7) = LEFT(o.OrgDetailKey, 7) "		'신세계만 LEFT 7로 조인해야됨..마지막 3자리가 안 맞음
									strSql = strSql & " WHERE "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and c.orderserial is NULL "
									strSql = strSql & " 	and o.orderserial is not NULL "
									strSql = strSql & " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' "
									strSql = strSql & " 	and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
									dbget.Execute strSql
								End If
							Next
						End If
					Set returnList = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

'롯데On 취소요청(완료) 조회
function GetCSOrderCancel_lotteon(sellsite, csGubun, csSubGubun, selldate)
	Dim xmlURL, strRst
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate, OrgOutMallOrderSerial
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete, claimType
	Dim apiUrl, apiKey
	divcd = "A008"
	iInputCnt = 0

	startdate = Left(DateAdd("d", 0, selldate), 10)
	startdate = Replace(startdate, "-", "") & "000000"

	enddate = Left(DateAdd("d", 0, selldate), 10)
	enddate = Replace(enddate, "-", "") & "235959"

	apiUrl = getApiUrl("lotteon")
	apiKey = getApiKey("lotteon")
	apiUrl = apiUrl & "/v1/openapi/claim/v1/cancellationOpenApi/getCancellationRequestAndComplateList"

	SET objJson = jsObject()
		objJson("srchStrtDttm") = startdate		'#검색시작일자 yyyyMMddhh24miss
		objJson("srchEndDttm") = enddate		'#검색종료일자 yyyyMMddhh24miss
		objJson("odNo") = ""					'o주문번호
		objJson("lrtrNo") = ""					'o하위거래처번호가 존재하면 무조건 필수입력
		strRst = objJson.jsString
	SET objJson = nothing

	'// 데이타 가져오기
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strRst)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			'rw iRbody
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				resultmsg	= strObj.message
				If returnCode = "0000" Then
					Set datalist = strObj.data
						For i=0 to datalist.length-1
							IsDelete				= "N"
							claimType				= ""
							OrgOutMallOrderSerial	= datalist.get(i).odNo
							OutMallOrderSerial		= datalist.get(i).odNo				'주문번호
							CSDetailKey				= datalist.get(i).clmNo				'클레임번호
							Set itemList = datalist.get(i).itemList
								For j = 0 to itemList.length-1
									OrgDetailKey	= itemList.get(j).odSeq				'주문순번
'									procSeq			= itemList.get(j).procSeq			'처리순번
'									orglProcSeq		= itemList.get(j).orglProcSeq		'원처리순번
'									odTypCd			= itemList.get(j).odTypCd			'주문유형코드( 20 주문취소 50 AS)
'									odTypDtlCd		= itemList.get(j).odTypDtlCd		'주문유형상세코드( 주문유형이 AS인 경우 필수) | 5011 : 재배송취소, 5021 : 추가배송취소, 5031 : 부분품회수취소
									claimType		= itemList.get(j).odPrgsStepCd		'주문진행단계코드(02 요청 21 취소완료 22 철회(취소요청) )
'									spdNo			= itemList.get(j).spdNo				'판매자상품번호
'									spdNm			= itemList.get(j).spdNm				'판매자상품명
'									sitmNo			= itemList.get(j).sitmNo			'판매자 단품번호
'									sitmNm			= itemList.get(j).sitmNm			'판매자 단품명
									itemno			= itemList.get(j).odQty				'주문수량
'									itmSlPrc		= itemList.get(j).itmSlPrc			'판매가 (할인을 제외한 상품 판매 원가)
'									cnclQty			= itemList.get(j).cnclQty			'취소수량
'									trNo			= itemList.get(j).trNo				'거래처번호
'									lrtrNo			= itemList.get(j).lrtrNo			'하위거래처번호
'									odAccpDttm		= itemList.get(j).odAccpDttm		'주문접수일시 [yyyyMMddHHmmss : 20191201121212]
'									purCfrmDttm		= itemList.get(j).purCfrmDttm		'구매확정일시 [yyyyMMddHHmmss : 20191201121212] 중개상품인경우 구매확정후취소일때 등록됨
									OutMallRegDate	= itemList.get(j).clmReqDttm		'클레임요청일시[yyyyMMddHHmmss : 20191201121212] - 취소요청일시
'									clmCmptDttm		= itemList.get(j).clmCmptDttm		'클레임완료일시[yyyyMMddHHmmss : 20191201121212] - 취소완료일시
									gubunname		= itemList.get(j).clmRsnCd			'클레임사유코드 | 101 : 배송이 늦어짐, 102 : 상품이 품절됨, 103 : 옵션/사이즈 불만 /취소, 104 : 다른 곳보다 비쌈, 105 : 쿠폰/할인혜택 변경, 106 : 구매의사 없어짐, 107 : 사은품 변경 / 취소, 108 : 결제수단 변경(구매사은 적용/카드변경 등), 109 : 유사상품 구매, 110 : 상품정보 미흡, 111 : 판매자 취소(판매자), 112 : 직권 취소(상담사), 113 : 자동 취소(결제부족), 114 : 자동 취소(결품취소), 115 : 자동 취소(선물미수령), 116 : 자동 취소(스마트픽미수령), 117 : 자동 취소(취소요청 정책 접수)

									Select Case gubunname
										Case "103", "104", "106", "109"
											gubunname = "단순변심"
										Case "108"
											gubunname = "선택변경"
										Case "101"
											gubunname = "배송지연"
										Case "110"
											gubunname = "상품정보틀림"
										Case "102"
											gubunname = "품절"
										Case "105"
											gubunname = "쿠폰미적용"
										Case Else
											gubunname = "기타"
									End Select

									If claimType = "22" Then
										IsDelete = "Y"
									End If

'									clmRsnCnts		= itemList.get(j).clmRsnCnts		'클레임사유내용
'									odFvrGrpNo		= itemList.get(j).odFvrGrpNo		'주문혜택그룹번호(N+1, 살수록할인 등 혜택이 묶여 있는 그룹번호)
'									fvrAmt			= itemList.get(j).fvrAmt			'수량 포함한 할인금액의 취소금액, 없으면 0 (취소완료시점에 생성) 1~5차할인금액 취소합계
'									sellerCnYn		= itemList.get(j).sellerCnYn		'판매자직권취소여부
'									frstDvCst		= itemList.get(j).frstDvCst			'초도배송비
'									addDvCst		= itemList.get(j).addDvCst			'추가배송비
'									dvCstBdnMnbdCd	= itemList.get(j).dvCstBdnMnbdCd	'배송비부담주체코드(01:고객 , 02: 업체)
'									excpProcDvsCd	= itemList.get(j).excpProcDvsCd		'예외처리구분코드(03 보류 , 04 거부)
'									rpcSpdNo		= itemList.get(j).rpcSpdNo			'대체판매자상품번호 (마트/슈퍼 존재할경우 제공)
'									rpcSpdNm		= itemList.get(j).rpcSpdNm			'대체판매자상품명 (마트/슈퍼 존재할경우 제공)
'									rpcSitmNo		= itemList.get(j).rpcSitmNo			'대체판매자단품번호(마트/슈퍼 존재할경우 제공)
'									rpcSitmNm		= itemList.get(j).rpcSitmNm			'대체판매자단품명(마트/슈퍼 존재할경우 제공)
'									cmbnDvGrpNo		= itemList.get(j).cmbnDvGrpNo		'합배송그룹번호

						            strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
									strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
									strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno, OrgOutMallOrderSerial) VALUES "
									strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
									strSql = strSql & "		'', '', '', '', '', '' "
									strSql = strSql & "		, CAST(STUFF(STUFF(STUFF('"&OutMallRegDate&"', 9, 0, ' '), 12, 0, ':'), 15, 0, ':') as datetime) , '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ", '" & OrgOutMallOrderSerial & "') "
									strSql = strSql & " END "
									dbget.Execute strSql, iAssignedRow

									If IsDelete = "Y" Then
										strSql = ""
										strSql = strSql & " UPDATE db_temp.dbo.tbl_xSite_TMPCS "
										strSql = strSql & " SET OutMallCurrState = 'B008' "
										strSql = strSql & " WHERE "
										strSql = strSql & " 	1 = 1 "
										strSql = strSql & " 	and SellSite = '" & sellsite & "' "
										strSql = strSql & " 	and OutMallOrderSerial = '" & OutMallOrderSerial & "' "
										strSql = strSql & " 	and CSDetailKey = '" & CSDetailKey & "' "
										strSql = strSql & " 	and OrgDetailKey = '" & OrgDetailKey & "' "
										strSql = strSql & " 	and divcd = '" & divcd & "' "
										dbget.Execute strSql
									End If

									'if (iAssignedRow > 0) then
										iInputCnt = iInputCnt+iAssignedRow
										''주문 입력 이전 내역은 삭제 하자
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET matchState = 'D'"
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder c "
										strSql = strSql & " WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										strSql = strSql & " and orderserial is NULL"
										dbget.Execute strSql

										'' CS 마스터정보. 업데이트
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
										strSql = strSql & " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
										strSql = strSql & " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
										strSql = strSql & " JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
										strSql = strSql & " ON "
										strSql = strSql & " 	1 = 1 "
										strSql = strSql & " 	and c.SellSite = o.SellSite "
										strSql = strSql & " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
										strSql = strSql & " 	and c.OrgDetailKey = o.OrgDetailKey "
										strSql = strSql & " WHERE "
										strSql = strSql & " 	1 = 1 "
										strSql = strSql & " 	and c.orderserial is NULL "
										strSql = strSql & " 	and o.orderserial is not NULL "
										strSql = strSql & " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										''response.write strSql & "<br />"
										dbget.Execute strSql

										If divcd = "A008" Then
											strSql = ""
											strSql = strSql & " UPDATE c "
											strSql = strSql & " SET c.currstate = 'B007' "
											strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
											strSql = strSql & " LEFT JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
											strSql = strSql & " ON "
											strSql = strSql & "		1 = 1 "
											strSql = strSql & "		and c.SellSite = o.SellSite "
											strSql = strSql & "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
											strSql = strSql & "		and c.OrgDetailKey = o.OrgDetailKey "
											strSql = strSql & " WHERE "
											strSql = strSql & "		1 = 1 "
											strSql = strSql & "		and c.orderserial is NULL "
											strSql = strSql & "		and o.SellSite is NULL "
											strSql = strSql & "		and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
											strSql = strSql & "		and c.currstate = 'B001' "
											strSql = strSql & "		and c.divcd = 'A008' "
											''rw strSql
											dbget.execute strSql
										End If
									'End If
								Next
							Set itemList = nothing
						Next
					Set datalist = nothing

					If IsDelete = "Y" then
						rw "취소철회 CS입력건수:"&iInputCnt
					Else
						rw "주문취소 CS입력건수:"&iInputCnt
					End If
				Else
					rw resultmsg
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'롯데On 반품요청/접수 목록조회
function GetCSOrderReturn_lotteon(sellsite, csGubun, csSubGubun, selldate)
	Dim xmlURL, strRst, claimType
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete
	Dim apiUrl, apiKey
	divcd = "A004"
	iInputCnt = 0

	startdate = Left(DateAdd("d", 0, selldate), 10)
	startdate = Replace(startdate, "-", "") & "000000"

	enddate = Left(DateAdd("d", 0, selldate), 10)
	enddate = Replace(enddate, "-", "") & "235959"

	apiUrl = getApiUrl("lotteon")
	apiKey = getApiKey("lotteon")
	apiUrl = apiUrl & "/v1/openapi/claim/v1/returningOpenApi/returnRequestSearch"

	SET objJson = jsObject()
		objJson("srchStrtDttm") = startdate		'#검색시작일자 yyyyMMddhh24miss
		objJson("srchEndDttm") = enddate		'#검색종료일자 yyyyMMddhh24miss
		objJson("odNo") = ""					'o주문번호( only 주문번호일때는 필수)
		objJson("lrtrNo") = ""					'o하위거래처번호
		strRst = objJson.jsString
	SET objJson = nothing

	'// 데이타 가져오기
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strRst)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				resultmsg	= strObj.message
				If returnCode = "0000" Then
					Set datalist = strObj.data
						For i=0 to datalist.length-1
							OutMallOrderSerial	= datalist.get(i).odNo				'주문번호
							CSDetailKey			= datalist.get(i).clmNo				'클레임번호
							Set itemList = datalist.get(i).itemList
								For j = 0 to itemList.length-1
									OrgDetailKey	= itemList.get(j).odSeq				'주문순번
'									procSeq			= itemList.get(j).procSeq			'처리순번
'									orglProcSeq		= itemList.get(j).orglProcSeq		'원처리순번
'									odTypCd			= itemList.get(j).odTypCd			'주문유형 (40 반품 50 AS)
'									odTypDtlCd		= itemList.get(j).odTypDtlCd		'주문유형상세코드 ( 주문유형이 AS인 경우 필수) | 5030 부분품회수
									claimType		= itemList.get(j).odPrgsStepCd		'주문진행단계(02요청 , 03 접수, 27 반품완료)
'									spdNo			= itemList.get(j).spdNo				'판매자상품번호
'									spdNm			= itemList.get(j).spdNm				'판매자상품명
'									sitmNo			= itemList.get(j).sitmNo			'판매자 단품번호
'									sitmNm			= itemList.get(j).sitmNm			'판매자 단품명
									itemno			= itemList.get(j).odQty				'주문수량
'									itmSlPrc		= itemList.get(j).itmSlPrc			'단품판매가
'									rtngQty			= itemList.get(j).rtngQty			'반품수량
'									trNo			= itemList.get(j).trNo				'거래처번호
'									lrtrNo			= itemList.get(j).lrtrNo			'하위거래처번호
'									OutMallRegDate	= itemList.get(j).odAccpDttm		'주문접수일시[yyyyMMddHHmmss : 20191201121212]
'									purCfrmDttm		= itemList.get(j).purCfrmDttm		'구매확정일시[yyyyMMddHHmmss : 20191201121212]
'									OutMallRegDate		= itemList.get(j).clmReqDttm		'클레임요청일시[yyyyMMddHHmmss : 20191201121212] -  반품요청일시
'									clmAccpDttm		= itemList.get(j).clmAccpDttm		'클레임접수일시[yyyyMMddHHmmss : 20191201121212] -  반품접수일시
'									clmCmptDttm		= itemList.get(j).clmCmptDttm		'클레임완료일시[yyyyMMddHHmmss : 20191201121212] - 반품완료일시
									gubunname		= itemList.get(j).clmRsnCd			'클레임사유코드 | 301 : 상품에 하자가 있음(파손/불량), 302 : 구매의사 없어짐, 303 : 구매의사 없어짐(기간 이후), 304 : 다른 상품이 배송됨, 305 : 상품이 생각과 다름(상품정보상이), 306 : 옵션/사이즈 불만, 307 : 결제수단변경(구매사은 적용/카드 변경 등), 308 : 상품/구성품이 안옴, 309 : 다른 곳보다 비쌈, 310 : 자동 반품(크로스픽미수령)

									Select Case gubunname
										Case "301"
											gubunname = "상품불량"
										Case "302", "303", "306", "309"
											gubunname = "단순변심"
										Case "304"
											gubunname = "오배송"
										Case "305"
											gubunname = "상품정보틀림"
										Case "307"
											gubunname = "선택변경"
										Case "308"
											gubunname = "배송지연"
										Case Else
											gubunname = "기타"
									End Select

'									clmRsnCnts		= itemList.get(j).clmRsnCnts		'클레임사유내용
'									spicYn			= itemList.get(j).spicYn			'스마트픽여부 (Y/N)
'									rtrvSeq			= itemList.get(j).rtrvSeq			'회수지순번(배송지순번)
'									rtrvCustNm		= itemList.get(j).rtrvCustNm		'회수지고객명
'									rtrvTelNo		= itemList.get(j).rtrvTelNo			'회수지전화번호
'									rtrvMphnNo		= itemList.get(j).rtrvMphnNo		'회수지휴대폰번호
'									rtrvZipNo		= itemList.get(j).rtrvZipNo			'회수지우편번호
'									rtrvZipNoSeq	= itemList.get(j).rtrvZipNoSeq		'회수지우편번호순번
'									rtrvStnmZipAddr	= itemList.get(j).rtrvStnmZipAddr	'회수지도로명우편주소(마트)
'									rtrvStnmDtlAddr	= itemList.get(j).rtrvStnmDtlAddr	'회수지도로명상세주소
'									dvMsg			= itemList.get(j).dvMsg				'배송메시지
'									spicTypCd		= itemList.get(j).spicTypCd			'스마트픽유형코드 (CRSS 크로스픽 RVS 리버스픽 STR 스토어픽)
'									rnkhSpplcNo		= itemList.get(j).rnkhSpplcNo		'상위픽업처번호
'									rnklSpplcNo		= itemList.get(j).rnklSpplcNo		'하위픽업처번호
'									pkupPlcNo		= itemList.get(j).pkupPlcNo			'픽업장소번호
'									pkupPlcNm		= itemList.get(j).pkupPlcNm			'픽업장소명
'									spicBxchNo		= itemList.get(j).spicBxchNo		'스마트픽교환권번호
'									pkupBgtDttm		= itemList.get(j).pkupBgtDttm		'픽업예정일시[yyyyMMddHHmmss : 20191201121212]
'									excpProcDvsCd	= itemList.get(j).excpProcDvsCd		'예외처리구분코드(03 보류 , 04 거부)
'									rpcSpdNo		= itemList.get(j).rpcSpdNo			'대체판매자상품번호 (마트/슈퍼 존재할경우 제공)
'									rpcSpdNm		= itemList.get(j).rpcSpdNm			'대체판매자상품명 (마트/슈퍼 존재할경우 제공)
'									rpcSitmNo		= itemList.get(j).rpcSitmNo			'대체판매자단품번호 (마트/슈퍼 존재할경우 제공)
'									rpcSitmNm		= itemList.get(j).rpcSitmNm			'대체판매자단품명 (마트/슈퍼 존재할경우 제공)
'									frstDvCst		= itemList.get(j).frstDvCst			'초도배송비
'									addDvCst		= itemList.get(j).addDvCst			'반품추가배송비
'									rcst			= itemList.get(j).rcst				'반품비
'									dvCstBdnMnbdCd	= itemList.get(j).dvCstBdnMnbdCd	'배송비부담주체코드(01:고객 , 02: 업체)
'									shopCnvMsg		= itemList.get(j).shopCnvMsg		'매장전달메세지
'									dvLrtrNo		= itemList.get(j).dvLrtrNo			'배송하위거래처번호-점코드 매핑정보 (슈퍼/마트)
'									hpDvDttm		= itemList.get(j).hpDvDttm			'희망배송일시 yyyymmddhh24miss (슈퍼/마트) | - 점포배송 배송일 및 예약배송상품의 배송일시 및 설치상품 희망배송일
'									afflSqncTypCd	= itemList.get(j).afflSqncTypCd		'계열사회차유형코드(슈퍼/마트) | CRSS_PIC:크로스픽, DN_DV:새벽배송, DRIVE_PIC:드라이브픽, DRT_DV:바로배송, FPRD_DV:정기배송, FS_PKUP:슈퍼픽업, RENTAL_CAR_PKUP:렌터카픽업, SHOP_DV:매장배송, SMT_QCK:스마트퀵, STR_PIC:스토어픽
'									afflCbCd		= itemList.get(j).afflCbCd			'계열사큐브코드(슈퍼/마트)
'									dvSqncCd		= itemList.get(j).dvSqncCd			'회차코드 통합On관리 회차번호(슈퍼/마트)
'									sftNoUseYn		= itemList.get(j).sftNoUseYn		'안심번호사용여부
'									sftDvpTelNo		= itemList.get(j).sftDvpTelNo		'안심배송지전화번호
'									sftDvpMphnNo	= itemList.get(j).sftDvpMphnNo		'안심배송지휴대폰번호
'									cmbnDvGrpNo		= itemList.get(j).frstDvCst			'합배송그룹번호

									strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
									strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
									strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
									strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
									strSql = strSql & "		'', '', '', '', '', '' "
									strSql = strSql & "		, CAST(STUFF(STUFF(STUFF('"&OutMallRegDate&"', 9, 0, ' '), 12, 0, ':'), 15, 0, ':') as datetime), '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
									strSql = strSql & " END "
									dbget.Execute strSql, iAssignedRow

									If (iAssignedRow > 0) Then
										iInputCnt = iInputCnt+iAssignedRow

										'' CS 마스터정보. 업데이트
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
										strSql = strSql & " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
										strSql = strSql & " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
										strSql = strSql & " JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
										strSql = strSql & " ON "
										strSql = strSql & " 	1 = 1 "
										strSql = strSql & " 	and c.SellSite = o.SellSite "
										strSql = strSql & " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
										strSql = strSql & " 	and c.OrgDetailKey = o.OrgDetailKey "
										strSql = strSql & " WHERE "
										strSql = strSql & " 	1 = 1 "
										strSql = strSql & " 	and c.orderserial is NULL "
										strSql = strSql & " 	and o.orderserial is not NULL "
										strSql = strSql & " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										dbget.Execute strSql
									End If
								Next
							Set itemList = nothing
						Next
					Set datalist = nothing
					rw "주문취소 CS입력건수:"&iInputCnt
				Else
					rw resultmsg
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
end function

'반품(요청)취소 목록 조회
function GetCSOrderReturnReject_lotteon(sellsite, csGubun, csSubGubun, selldate)
	Dim xmlURL, strRst
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete
	Dim apiUrl, apiKey
	divcd = "A004"
	iInputCnt = 0

	startdate = Left(DateAdd("d", 0, selldate), 10)
	startdate = Replace(startdate, "-", "") & "000000"

	enddate = Left(DateAdd("d", 0, selldate), 10)
	enddate = Replace(enddate, "-", "") & "235959"

	apiUrl = getApiUrl("lotteon")
	apiKey = getApiKey("lotteon")
	apiUrl = apiUrl & "/v1/openapi/claim/v1/returningOpenApi/returnWithdrawSearch"

	SET objJson = jsObject()
		objJson("srchStrtDttm") = startdate		'#검색시작일자 yyyyMMddhh24miss
		objJson("srchEndDttm") = enddate		'#검색종료일자 yyyyMMddhh24miss
		objJson("odNo") = ""					'o주문번호( only 주문번호일때는 필수)
		objJson("lrtrNo") = ""					'o하위거래처번호
		objJson("odTypCd") = "41"				'주문유형코드(41 반품취소 50 AS)(NULL 인경우 전체조회)
		strRst = objJson.jsString
	SET objJson = nothing

	'// 데이타 가져오기
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strRst)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				resultmsg	= strObj.message
				If returnCode = "0000" Then
					Set datalist = strObj.data
						For i=0 to datalist.length-1
							OutMallOrderSerial	= datalist.get(i).odNo				'주문번호
							CSDetailKey			= datalist.get(i).clmNo				'클레임번호
							Set itemList = datalist.get(i).itemList
								For j = 0 to itemList.length-1
									OrgDetailKey	= itemList.get(j).odSeq				'주문순번
'									procSeq			= itemList.get(j).procSeq			'처리순번
'									orglProcSeq		= itemList.get(j).orglProcSeq		'원처리순번
'									odTypCd			= itemList.get(j).odTypCd			'주문유형코드(41 반품취소 50 AS)
'									odTypDtlCd		= itemList.get(j).odTypDtlCd		'주문유형상세코드(5031 부분품회수취소)
'									odPrgsStepCd	= itemList.get(j).odPrgsStepCd		'주문진행단계코드(02 요청 21 취소완료 22 취소철회 ), 계열사는 21 취소완료만 사용
'									spdNo			= itemList.get(j).spdNo				'판매자상품번호
'									spdNm			= itemList.get(j).spdNm				'판매자상품명
'									sitmNo			= itemList.get(j).sitmNo			'판매자 단품번호
'									sitmNm			= itemList.get(j).sitmNm			'판매자 단품명
									itemno			= itemList.get(j).odQty				'주문수량
'									itmSlPrc		= itemList.get(j).itmSlPrc			'단품판매가, (할인을 제외한 판매원가)
'									rtngQty			= itemList.get(j).rtngQty			'반품수량
'									trNo			= itemList.get(j).trNo				'거래처번호
'									lrtrNo			= itemList.get(j).lrtrNo			'하위거래처번호
									OutMallRegDate	= itemList.get(j).clmReqDttm		'클레임요청일시[yyyyMMddHHmmss : 20191201121212] - 반품요청일시
'									clmCmptDttm		= itemList.get(j).clmCmptDttm		'클레임완료일시[yyyyMMddHHmmss : 20191201121212] - 반품완료일시

									strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
									strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
									strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
									strSql = strSql & " 	('" & divcd & "', '기타', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
									strSql = strSql & "		'', '', '', '', '', '' "
									strSql = strSql & "		, CAST(STUFF(STUFF(STUFF('"&OutMallRegDate&"', 9, 0, ' '), 12, 0, ':'), 15, 0, ':') as datetime), '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
									strSql = strSql & " END "
									dbget.Execute strSql, iAssignedRow

									strSql = ""
									strSql = strSql & " UPDATE db_temp.dbo.tbl_xSite_TMPCS "
									strSql = strSql & " SET OutMallCurrState = 'B008' "
									strSql = strSql & " WHERE "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and SellSite = '" & sellsite & "' "
									strSql = strSql & " 	and OutMallOrderSerial = '" & OutMallOrderSerial & "' "
									strSql = strSql & " 	and CSDetailKey = '" & CSDetailKey & "' "
									strSql = strSql & " 	and OrgDetailKey = '" & OrgDetailKey & "' "
									strSql = strSql & " 	and divcd = '" & divcd & "' "
									dbget.Execute strSql

									If (iAssignedRow > 0) Then
										iInputCnt = iInputCnt+iAssignedRow

										'' CS 마스터정보. 업데이트
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
										strSql = strSql & " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
										strSql = strSql & " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
										strSql = strSql & " JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
										strSql = strSql & " ON "
										strSql = strSql & " 	1 = 1 "
										strSql = strSql & " 	and c.SellSite = o.SellSite "
										strSql = strSql & " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
										strSql = strSql & " 	and c.OrgDetailKey = o.OrgDetailKey "
										strSql = strSql & " WHERE "
										strSql = strSql & " 	1 = 1 "
										strSql = strSql & " 	and c.orderserial is NULL "
										strSql = strSql & " 	and o.orderserial is not NULL "
										strSql = strSql & " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										dbget.Execute strSql
									End If
								Next
							Set itemList = nothing
							rw "반품철회 CS입력건수:"&iInputCnt
						Next
					Set datalist = nothing
				Else
					rw resultmsg
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
end function

'교환요청/접수목록조회
function GetCSOrderExchange_lotteon(sellsite, csGubun, csSubGubun, selldate)
	Dim xmlURL, strRst
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete
	Dim apiUrl, apiKey
	divcd = "A000"
	iInputCnt = 0

	startdate = Left(DateAdd("d", 0, selldate), 10)
	startdate = Replace(startdate, "-", "") & "000000"

	enddate = Left(DateAdd("d", 0, selldate), 10)
	enddate = Replace(enddate, "-", "") & "235959"

	apiUrl = getApiUrl("lotteon")
	apiKey = getApiKey("lotteon")
	apiUrl = apiUrl & "/v1/openapi/claim/v1/exchangeOpenApi/exchangeSearch"

	SET objJson = jsObject()
		objJson("srchStrtDttm") = startdate		'#검색시작일자 yyyyMMddhh24miss
		objJson("srchEndDttm") = enddate		'#검색종료일자 yyyyMMddhh24miss
		objJson("odNo") = ""					'o주문번호( only 주문번호일때는 필수)
		objJson("lrtrNo") = ""					'o하위거래처번호
		strRst = objJson.jsString
	SET objJson = nothing

	'// 데이타 가져오기
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strRst)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				resultmsg	= strObj.message
				If returnCode = "0000" Then
					Set datalist = strObj.data
						For i=0 to datalist.length-1
							OutMallOrderSerial	= datalist.get(i).odNo				'주문번호
							CSDetailKey			= datalist.get(i).clmNo				'클레임번호
rw iRbody
							Set itemList = datalist.get(i).itemList
								For j = 0 to itemList.length-1
									OrgDetailKey	= itemList.get(j).odSeq				'주문순번
'									procSeq			= itemList.get(j).procSeq			'처리순번 : Default 1 단품단위로 처리순서값을 정의최초 입력시 1 이고 클레임이 발생할 경우 1씩 증가함
'									orglProcSeq		= itemList.get(j).orglProcSeq		'원처리순번
'									odTypCd			= itemList.get(j).odTypCd			'주문유형 | 10 : 주문 ,20 : 취소(주문취소), 30 : 교환, 31 : 교환취소, 40 : 반품, 41 : 반품취소, 50 : AS
'									odTypDtlCd		= itemList.get(j).odTypDtlCd		'주문유형상세코드( 5010 재배송 5020 추가배송 (부분품))
'									odPrgsStepCd	= itemList.get(j).odPrgsStepCd		'주문진행단계(교환요청) | 02 : 요청, 03 : 접수, 11 : 출고지시, 12 : 상품준비, 13 : 발송완료, 14 : 배송완료, 15 : 수취완료, 21 : 취소완료, 22 : 취소철회, 23 : 회수지시, 24 : 회수진행, 25 : 회수완료, 26 : 회수확정, 27 : 반품완료
'									spdNo			= itemList.get(j).spdNo				'판매자상품번호
'									spdNm			= itemList.get(j).spdNm				'판매자상품명
'									sitmNo			= itemList.get(j).sitmNo			'판매자 단품번호
'									sitmNm			= itemList.get(j).sitmNm			'판매자 단품명
									itemno			= itemList.get(j).odQty				'주문수량
'									itmSlPrc		= itemList.get(j).itmSlPrc			'단품판매가
'									xchgQty			= itemList.get(j).xchgQty			'교환수량
'									trNo			= itemList.get(j).trNo				'거래처번호
'									lrtrNo			= itemList.get(j).lrtrNo			'하위거래처번호
									OutMallRegDate	= itemList.get(j).odAccpDttm		'주문접수일시 [yyyyMMddHHmmss]
'									purCfrmDttm		= itemList.get(j).purCfrmDttm		'구매확정일시 [yyyyMMddHHmmss]
'									clmReqDttm		= itemList.get(j).clmReqDttm		'클레임요청일시 [yyyyMMddHHmmss]
'									clmAccpDttm		= itemList.get(j).clmAccpDttm		'클레임접수일시[yyyyMMddHHmmss : 20191201121212] - 교환접수일시
'									clmCmptDttm		= itemList.get(j).clmCmptDttm		'클레임완료일시[yyyyMMddHHmmss : 20191201121212] - 교환완료일시
									gubunname		= itemList.get(j).clmRsnCd			'클레임사유코드 | 201 : 상품에 하자가 있음(파손/불량), 202 : 다른 상품이 배송됨, 203 : 상품이 생각과 다름(상품정보상이), 204 : 고객 취향과 상이함

									Select Case gubunname
										Case "201"
											gubunname = "상품불량"
										Case "202"
											gubunname = "오배송"
										Case "203"
											gubunname = "상품정보틀림"
										Case "204"
											gubunname = "단순변심"
										Case Else
											gubunname = "기타"
									End Select

'									clmRsnCnts		= itemList.get(j).clmRsnCnts		'클레임사유내용
'									spicYn			= itemList.get(j).spicYn			'스마트픽여부 (Y/N)
'									dvRtrvDvsCd		= itemList.get(j).dvRtrvDvsCd		'배송회수구분코드 | RTRV:회수, DV:배송
'									frstDvCst		= itemList.get(j).frstDvCst			'초도배송비
'									rtngAddDvCst	= itemList.get(j).rtngAddDvCst		'반품 추가배송비
'									rtngDvCst		= itemList.get(j).rtngDvCst			'반품배송비 - 요청일 경우에는 값이 없음
'									xchgAddDvCst	= itemList.get(j).xchgAddDvCst		'교환 추가배송비
'									xchgDvCst		= itemList.get(j).xchgDvCst			'교환배송비 - 요청일 경우에는 값이 없음
'									dvCstBdnMnbdCd	= itemList.get(j).dvCstBdnMnbdCd	'배송비부담주체코드 (01:고객 / 02: 업체)
'									shopCnvMsg		= itemList.get(j).shopCnvMsg		'매장전달메시지
'									xchgDvsCd		= itemList.get(j).xchgDvsCd			'교환구분코드(01:일반교환,02:맞교환)
'									cmbnDvPsbYn		= itemList.get(j).cmbnDvPsbYn		'합배송가능여부 (Y/N)
'									cmbnDvGrpNo		= itemList.get(j).cmbnDvGrpNo		'합배송그룹번호
'									dvLrtrNo		= itemList.get(j).dvLrtrNo			'배송하위거래처번호-점코드 매핑정보 (슈퍼/마트)
'									hpDvDttm		= itemList.get(j).hpDvDttm			'희망배송일시 yyyymmddhh24miss (슈퍼/마트) - 점포배송 배송일 및 예약배송상품의 배송일시 및 설치상품 희망배송일
'									afflSqncTypCd	= itemList.get(j).afflSqncTypCd		'계열사회차유형코드(슈퍼/마트) | CRSS_PIC:크로스픽, DN_DV:새벽배송, DRIVE_PIC:드라이브픽, DRT_DV:바로배송, FPRD_DV:정기배송, FS_PKUP:슈퍼픽업, RENTAL_CAR_PKUP:렌터카픽업, SHOP_DV:매장배송, SMT_QCK:스마트퀵, STR_PIC:스토어픽
'									afflCbCd		= itemList.get(j).afflCbCd			'계열사큐브코드(슈퍼/마트)
'									dvSqncCd		= itemList.get(j).dvSqncCd			'회차코드 통합On관리 회차번호(슈퍼/마트)
'									rtrvSeq			= itemList.get(j).rtrvSeq			'회수지순번(배송지순번)
'									rtrvCustNm		= itemList.get(j).rtrvCustNm		'회수지고객명
'									rtrvTelNo		= itemList.get(j).rtrvTelNo			'회수지전화번호
'									rtrvMphnNo		= itemList.get(j).rtrvMphnNo		'회수지휴대폰번호
'									rtrvZipNo		= itemList.get(j).rtrvZipNo			'회수지우편번호
'									rtrvZipNoSeq	= itemList.get(j).rtrvZipNoSeq		'회수지우편번호순번(마트)
'									rtrvStnmZipAddr	= itemList.get(j).rtrvStnmZipAddr	'회수지도로명우편주소
'									rtrvStnmDtlAddr	= itemList.get(j).rtrvStnmDtlAddr	'회수지도로명상세주소
'									dvpSeq			= itemList.get(j).dvpSeq			'배송지순번(배송지순번)
'									dvpCustNm		= itemList.get(j).dvpCustNm			'배송지고객명
'									dvpTelNo		= itemList.get(j).dvpTelNo			'배송지전화번호
'									dvpMphnNo		= itemList.get(j).dvpMphnNo			'배송지휴대폰번호
'									dvpZipNo		= itemList.get(j).dvpZipNo			'배송지우편번호
'									dvpZipNoSeq		= itemList.get(j).dvpZipNoSeq		'배송지우편번호순번(마트)
'									dvpStnmZipAddr	= itemList.get(j).dvpStnmZipAddr	'배송지기본주소
'									dvpStnmDtlAddr	= itemList.get(j).dvpStnmDtlAddr	'배송지상세주소
'									dvMsg			= itemList.get(j).dvMsg				'배송메시지
'									sftNoUseYn		= itemList.get(j).sftNoUseYn		'안심번호사용여부 (회수지포함)
'									sftDvpTelNo		= itemList.get(j).sftDvpTelNo		'안심배송지전화번호 (회수지포함)
'									sftDvpMphnNo	= itemList.get(j).sftDvpMphnNo		'안심배송지휴대폰번호 (회수지포함)
'									jntFtdrPwd		= itemList.get(j).jntFtdrPwd		'공동현관비밀번호
'									dnDvRcptOptCd	= itemList.get(j).dnDvRcptOptCd		'배송수령옵션 코드 | 10 : 집앞배달, 20 : 경비실배달
'									dnDvVstMthdCd	= itemList.get(j).dnDvVstMthdCd		'배송방문방법코드 | 10:공동현관비밀번호, 20:자유출입가능, 30:세대호출, 40:경비실통한연락, 50:직접전화
'									rpcSpdNo		= itemList.get(j).rpcSpdNo			'대체판매자상품번호 (마트/슈퍼 존재할경우 제공)
'									rpcSpdNm		= itemList.get(j).rpcSpdNm			'대체판매자상품명 (마트/슈퍼 존재할경우 제공)
'									rpcSitmNo		= itemList.get(j).rpcSitmNo			'대체판매자단품번호 (마트/슈퍼 존재할경우 제공)
'									rpcSitmNm		= itemList.get(j).rpcSitmNm			'대체판매자단품명 (마트/슈퍼 존재할경우 제공)

									strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
									strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
									strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
									strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
									strSql = strSql & "		'', '', '', '', '', '' "
									strSql = strSql & "		, CAST(STUFF(STUFF(STUFF('"&OutMallRegDate&"', 9, 0, ' '), 12, 0, ':'), 15, 0, ':') as datetime), '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
									strSql = strSql & " END "
									dbget.Execute strSql, iAssignedRow

									If (iAssignedRow > 0) Then
										iInputCnt = iInputCnt+iAssignedRow

										'' CS 마스터정보. 업데이트
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
										strSql = strSql & " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
										strSql = strSql & " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
										strSql = strSql & " JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
										strSql = strSql & " ON "
										strSql = strSql & " 	1 = 1 "
										strSql = strSql & " 	and c.SellSite = o.SellSite "
										strSql = strSql & " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
										strSql = strSql & " 	and c.OrgDetailKey = o.OrgDetailKey "
										strSql = strSql & " WHERE "
										strSql = strSql & " 	1 = 1 "
										strSql = strSql & " 	and c.orderserial is NULL "
										strSql = strSql & " 	and o.orderserial is not NULL "
										strSql = strSql & " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										dbget.Execute strSql
									End If
 								Next
							Set itemList = nothing
						Next
					Set datalist = nothing
					rw "교환 CS입력건수:"&iInputCnt
				Else
					rw resultmsg
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
end function

'교환(요청)취소 목록 조회
function GetCSOrderExchangeReject_lotteon(sellsite, csGubun, csSubGubun, selldate)
	Dim xmlURL, strRst
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete
	Dim apiUrl, apiKey
	divcd = "A000"
	iInputCnt = 0

	startdate = Left(DateAdd("d", 0, selldate), 10)
	startdate = Replace(startdate, "-", "") & "000000"

	enddate = Left(DateAdd("d", 0, selldate), 10)
	enddate = Replace(enddate, "-", "") & "235959"

	apiUrl = getApiUrl("lotteon")
	apiKey = getApiKey("lotteon")
	apiUrl = apiUrl & "/v1/openapi/claim/v1/exchangeOpenApi/exchangeWithdrawSearch"

	SET objJson = jsObject()
		objJson("srchStrtDttm") = startdate		'#검색시작일자 yyyyMMddhh24miss
		objJson("srchEndDttm") = enddate		'#검색종료일자 yyyyMMddhh24miss
		objJson("odNo") = ""					'o주문번호( only 주문번호일때는 필수)
		objJson("lrtrNo") = ""					'o하위거래처번호
		strRst = objJson.jsString
	SET objJson = nothing

	'// 데이타 가져오기
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strRst)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				resultmsg	= strObj.message
				If returnCode = "0000" Then
					Set datalist = strObj.data
						For i=0 to datalist.length-1
							OutMallOrderSerial	= datalist.get(i).odNo				'주문번호
							CSDetailKey			= datalist.get(i).clmNo				'클레임번호
							Set itemList = datalist.get(i).itemList
								For j = 0 to itemList.length-1
									OrgDetailKey	= itemList.get(j).odSeq				'주문순번
'									procSeq			= itemList.get(j).procSeq			'처리순번
'									orglProcSeq		= itemList.get(j).orglProcSeq		'원처리순번
'									odTypCd			= itemList.get(j).odTypCd			'주문유형코드(31 교환취소)
'									odPrgsStepCd	= itemList.get(j).odPrgsStepCd		'주문진행단계코드(02 요청 21 취소완료 22 취소철회 ), 계열사는 21 취소완료만 사용
'									dvRtrvDvsCd		= itemList.get(j).dvRtrvDvsCd		'배송회수구분코드 | RTRV:회수, DV:배송
'									spdNo			= itemList.get(j).spdNo				'판매자상품번호
'									spdNm			= itemList.get(j).spdNm				'판매자상품명
'									sitmNo			= itemList.get(j).sitmNo			'판매자 단품번호
'									sitmNm			= itemList.get(j).sitmNm			'판매자 단품명
									itemno			= itemList.get(j).odQty				'주문수량
'									itmSlPrc		= itemList.get(j).itmSlPrc			'단품판매가, (할인을 제외한 상품 판매원가)
'									xchgQty			= itemList.get(j).xchgQty			'교환수량
'									trNo			= itemList.get(j).trNo				'거래처번호
'									odAccpDttm		= itemList.get(j).odAccpDttm		'주문접수일시 [yyyyMMddHHmmss]
									OutMallRegDate	= itemList.get(j).clmReqDttm		'클레임요청일시[yyyyMMddHHmmss : 20191201121212] - 교환철회요청일시
'									clmCmptDttm		= itemList.get(j).clmCmptDttm		'클레임완료일시[yyyyMMddHHmmss : 20191201121212] - 교환철회완료일시

									strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
									strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
									strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
									strSql = strSql & " 	('" & divcd & "', '기타', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
									strSql = strSql & "		'', '', '', '', '', '' "
									strSql = strSql & "		, CAST(STUFF(STUFF(STUFF('"&OutMallRegDate&"', 9, 0, ' '), 12, 0, ':'), 15, 0, ':') as datetime), '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ") "
									strSql = strSql & " END "
									dbget.Execute strSql, iAssignedRow

									strSql = ""
									strSql = strSql & " UPDATE db_temp.dbo.tbl_xSite_TMPCS "
									strSql = strSql & " SET OutMallCurrState = 'B008' "
									strSql = strSql & " WHERE "
									strSql = strSql & " 	1 = 1 "
									strSql = strSql & " 	and SellSite = '" & sellsite & "' "
									strSql = strSql & " 	and OutMallOrderSerial = '" & OutMallOrderSerial & "' "
									strSql = strSql & " 	and CSDetailKey = '" & CSDetailKey & "' "
									strSql = strSql & " 	and OrgDetailKey = '" & OrgDetailKey & "' "
									strSql = strSql & " 	and divcd = '" & divcd & "' "
									dbget.Execute strSql

									If (iAssignedRow > 0) Then
										iInputCnt = iInputCnt+iAssignedRow

										'' CS 마스터정보. 업데이트
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
										strSql = strSql & " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
										strSql = strSql & " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
										strSql = strSql & " JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
										strSql = strSql & " ON "
										strSql = strSql & " 	1 = 1 "
										strSql = strSql & " 	and c.SellSite = o.SellSite "
										strSql = strSql & " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
										strSql = strSql & " 	and c.OrgDetailKey = o.OrgDetailKey "
										strSql = strSql & " WHERE "
										strSql = strSql & " 	1 = 1 "
										strSql = strSql & " 	and c.orderserial is NULL "
										strSql = strSql & " 	and o.orderserial is not NULL "
										strSql = strSql & " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										dbget.Execute strSql
									End If
 								Next
							Set itemList = nothing
						Next
					Set datalist = nothing
					rw "교환철회 CS입력건수:"&iInputCnt
				Else
					rw resultmsg
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
end function

Public Function getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iserv, ioper)
	Dim cryptoLib, oLicense, osecretKey, otimeStamp, osignature
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
		If (application("Svr_Info") = "Dev") Then
			iaccessLicense = "01000100004b035a25d67f991849cad1c7042b8da528d13e9ddce6878f2e43ac88080e0a5e" 'AccessLicense Key 입력, PDF파일참조
			osecretKey = "AQABAAAWPWagCrPjFQnFEtxs5j+oyZFwuzomdNq0XZSricPuMw=="  'SecreKey 입력, PDF파일참조
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		Else
			iaccessLicense = "010001000019133c715650b9c85b820961612f2b90b431ddd8654b42c097c4df1a43d0be09" 'AccessLicense Key 입력, PDF파일참조
			osecretKey = "AQABAADX6Hz/wORFJS5pSIy4KQXkH83gC9G1aXChxBjcnUMqWw=="  'SecreKey 입력, PDF파일참조
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		End If
	Set cryptoLib = nothing
End Function

Public Function generateKey_nvstorefarm(iTimestamp)
	Dim cryptoLib, oLicense, osecretKey, otimeStamp, osignature
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
		If (application("Svr_Info") = "Dev") Then
			osecretKey = "AQABAAAWPWagCrPjFQnFEtxs5j+oyZFwuzomdNq0XZSricPuMw=="  'SecreKey 입력, PDF파일참조
			generateKey_nvstorefarm = cryptoLib.generateKey(iTimestamp, osecretKey)
		Else
			osecretKey = "AQABAADX6Hz/wORFJS5pSIy4KQXkH83gC9G1aXChxBjcnUMqWw=="  'SecreKey 입력, PDF파일참조
			generateKey_nvstorefarm = cryptoLib.generateKey(iTimestamp, osecretKey)
		End If
	Set cryptoLib = nothing
End Function

'// 11111111111111
'// /outmall/gseshop/gseshopItemcls.asp 참조
CONST CGSShopCompanyCode = 1003890	'' 협력사코드
function GetCSOrderCancel_gseshop(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj, objData, objDetail, objDetailArr
	dim i, j, k
	dim startdate, enddate
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt, xmlSelldate
	dim strSql
	dim skipError

	xmlSelldate = Replace(selldate, "-", "")

	'// API URL(기간동안의 주문 가져오기)
	'// tnsType : 주문구분(주문/반품 : S, 취소 : C)
	'// 개발 : test1 운영 : ecb2b
	if (application("Svr_Info") = "Dev") then
		xmlURL = "http://test1.gsshop.com/SupSendOrderInfo.gs?supCd=" + CStr(CGSShopCompanyCode) + "&sdDt=" + CStr(xmlSelldate) + "&tnsType=C"
	else
		xmlURL = "http://ecb2b.gsshop.com/SupSendOrderInfo.gs?supCd=" + CStr(CGSShopCompanyCode) + "&sdDt=" + CStr(xmlSelldate) + "&tnsType=C"
	end if
	''response.write xmlURL & "<br />"
	''dbget.close : response.end

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.setTimeouts 2000,2000,2000,2000

	on error resume next
		objXML.send()


		if objXML.Status <> "200" then
			response.write "ERROR : 통신오류"
			''dbget.close : response.end
		end if

	on error goto 0
	'// 전송요청만 한다.(XML 수신X)

	'// /wapi/outmall/order/xSiteOrder_GSShop_recv_Process.asp 참조

end function

'GSShop 취소 조회
function GetCSOrderNewCancel_gseshop(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, xmlSelldate, obj
	dim objXML, strObj, objData, jParam, requireDetailObj, requireDetail
	dim i, j, k, strsql
	dim successCnt : successCnt = 0
	Dim returnCode, resultMsg
	Dim apiUrl
	'// =======================================================================
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")
	If (application("Svr_Info") = "Dev") Then
		apiUrl = "http://realapi.gsshop.com/b2b/SupSendOrderInfo.gs"
		'apiUrl = "http://testapi.gsshop.com/b2b/SupSendOrderInfo.gs"
	Else
		apiUrl = "http://realapi.gsshop.com/b2b/SupSendOrderInfo.gs"
	End If

	'// =======================================================================
	Set obj = jsObject()
		obj("sender") = "TBT"					'전송자 (협력사별 부여되는 ID로 GS에서 제공)
		obj("receiver") = "GS SHOP"				'수신자 | GS SHOP
		obj("documentId") = "ORDINF"			'문서ID | ORDINF
		obj("processType") = "C"				'전송구분 | A:전체, S:주문/반품, C:취소
		obj("supCd") = ""&CGSShopCompanyCode&""	'협력사코드	| (협력사번호 GS에서 제공)
		obj("sdDt") = xmlSelldate				'조회일자
		jParam = obj.jsString
	Set obj = nothing

	'// 데이타 가져오기
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(jParam)
		If objXML.Status <> "200" Then
			response.write "ERROR : 통신오류"
			dbget.close : response.end
		Else
			objData = BinaryToText(objXML.ResponseBody,"utf-8")
		End If

	'// =======================================================================
	'// Json 파싱
		Set strObj = JSON.parse(objData)
			returnCode		= strObj.resultCd
			resultMsg		= strObj.resultMsg
			If returnCode <> "S" Then
				response.write "ERROR : 오류" & resultMsg
				dbget.close : response.end
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

function GetCSCheckStatus(byVal sellsite, byVal csGubun, byRef LastCheckDate, byRef isSuccess)
	dim strSql

    strSql = " IF NOT Exists("
    strSql = strSql + " 	select LastcheckDate"
    strSql = strSql + " 	from db_temp.[dbo].[tbl_xSite_TMPCS_timestamp]"
    strSql = strSql + " 	where sellsite='" + CStr(sellsite) + "' and csGubun='" + CStr(csGubun) + "'"
	strSql = strSql + " )"
	strSql = strSql + " BEGIN"
	strSql = strSql + "		insert into db_temp.[dbo].[tbl_xSite_TMPCS_timestamp](sellsite, csGubun, lastcheckdate, issuccess, LastUpdate) "
	strSql = strSql + "		values('" & sellsite & "', '" & csGubun & "', '" & Left(DateAdd("d", -1, Now()), 10) & "', 'N', getdate()) "
	strSql = strSql + " END"
	dbget.Execute strSql

	strSql = " select convert(varchar(10), LastCheckDate, 121) as LastCheckDate, isSuccess from db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	strSql = strSql + " 	where sellsite='" + CStr(sellsite) + "' and csGubun='" + CStr(csGubun) + "'"

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		LastCheckDate = rsget("LastCheckDate")
		isSuccess = rsget("isSuccess")
	rsget.Close
end function

function SetCSCheckStatus(sellsite, csGubun, LastCheckDate, isSuccess)
	dim strSql

	strSql = " update db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	strSql = strSql + " set lastcheckdate = '" & LastCheckDate & "', issuccess = '" & isSuccess & "', LastUpdate = getdate() "
	strSql = strSql + " where sellsite = '" + CStr(sellsite) + "' and csGubun='" + CStr(csGubun) + "' "
	''response.write strSql
	dbget.Execute strSql
end function

'// 반품(취소)요청 단건 조회
Function GetCSOrderCancel_One_coupang(sellsite, OutMallOrderSerial, receiptid)
	dim xmlURL, strRst, iRbody, strObj
	dim objXML, xmlDOM, objArr, obj
	dim i, j, k
	dim startdate, enddate, retCode, strObjValues, strObjValuesReturnItems
	dim OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno, OrderName, OrderHpNo
	dim iAssignedRow, iInputCnt
	dim strSql, ResultCode, ResultMsg, cancelOrderList, SubNodes
	Dim ClaimSeq, ClaimMemo, RegYMD
    dim OrgOutMallOrderSerial, OutMallCurrState

    iAssignedRow = 0

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://xapi.10x10.co.kr:8080/return/coupangnew/returnsingle?receiptid=" & receiptid, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			''response.write iRbody & "<br />"
            ''response.write "sellsite :" & sellsite & "<br />"
            ''response.write "OutMallOrderSerial :" & OutMallOrderSerial & "<br />"
            ''response.write "receiptid :" & receiptid & "<br />"
			''response.end

			Set strObj = JSON.parse(iRbody)
				If strObj.success = True Then
					set strObjValues = strObj.outPutValue.data

                    ''RELEASE_STOP_UNCHECKED	출고중지요청
                    ''RETURNS_UNCHECKED	반품접수
                    ''VENDOR_WAREHOUSE_CONFIRM	입고완료
                    ''REQUEST_COUPANG_CHECK	쿠팡확인요청
                    ''RETURNS_COMPLETED	반품완료
                    OutMallCurrState = ""
                    if strObjValues.receiptStatus = "RETURNS_UNCHECKED" then
                        OutMallCurrState = "B001"
                    elseif strObjValues.receiptStatus = "RETURNS_COMPLETED" then
                        OutMallCurrState = "B007"
                    elseif strObjValues.receiptStatus = "VENDOR_WAREHOUSE_CONFIRM" then
                        OutMallCurrState = "B006"
                    end if

                    strSql = " update db_temp.dbo.tbl_xSite_TMPCS "
                    strSql = strSql + " set OutMallCurrState = '" & OutMallCurrState & "' "
                    strSql = strSql + " where "
                    strSql = strSql + " 	1 = 1 "
                    strSql = strSql + " 	and SellSite = '" & sellsite & "' "
                    strSql = strSql + " 	and OutMallOrderSerial = '" & OutMallOrderSerial & "' "
                    strSql = strSql + " 	and CSDetailKey = '" & receiptid & "' "
                    if OutMallCurrState <> "" then
                        dbget.Execute strSql,iAssignedRow
                    end if
					set strObjValues = nothing
                else
                    strSql = " update db_temp.dbo.tbl_xSite_TMPCS "
                    strSql = strSql + " set OutMallCurrState = 'B008' "
                    strSql = strSql + " where "
                    strSql = strSql + " 	1 = 1 "
                    strSql = strSql + " 	and SellSite = '" & sellsite & "' "
                    strSql = strSql + " 	and OutMallOrderSerial = '" & OutMallOrderSerial & "' "
                    strSql = strSql + " 	and CSDetailKey = '" & receiptid & "' "
                    dbget.Execute strSql,iAssignedRow
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing

	''rw "CS 취소(반품) 건수:" & iAssignedRow

    GetCSOrderCancel_One_coupang = iAssignedRow

End Function

function fnMatchCs(sellsite, ioutmallorderserial)
    dim affectedRow, strSql

    strSql = " update T "
	strSql = strSql & " set T.asid = a.id "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_list] a "
	strSql = strSql & " 	on "
	strSql = strSql & " 		1 = 1 "
	strSql = strSql & " 		and a.orderserial = T.OrderSerial "
	strSql = strSql & " 		and a.deleteyn = 'N' "
	strSql = strSql & " 		and ( "
	strSql = strSql & " 			(T.divcd = 'A004' and a.divcd in ('A004', 'A010', 'A008', 'A011', 'A012', 'A112', 'A112')) "
	strSql = strSql & " 			or "
	strSql = strSql & " 			(T.divcd = 'A011' and a.divcd in ('A011', 'A012', 'A112', 'A112')) "
    strSql = strSql & " 			or "
    strSql = strSql & " 			(T.divcd = 'A000' and a.divcd in ('A000', 'A100')) "
    strSql = strSql & " 			or "
    strSql = strSql & " 			(T.divcd = 'A008' and a.divcd in ('A008', 'A004', 'A010')) "
	strSql = strSql & " 		) "
	strSql = strSql & " 		and a.id not in ( "
	strSql = strSql & " 			select T.asid "
	strSql = strSql & " 			from "
	strSql = strSql & " 				[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " 			where "
	strSql = strSql & " 				1 = 1 "
	strSql = strSql & " 				and T.SellSite = '" & sellsite & "' "
	strSql = strSql & " 				and T.OutMallOrderSerial = '" & ioutmallorderserial & "' "
	strSql = strSql & " 				and T.asid is not NULL "
	strSql = strSql & " 		) "
	strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_detail] d "
	strSql = strSql & " 	on "
	strSql = strSql & " 		1 = 1 "
	strSql = strSql & " 		and a.id = d.masterid "
	strSql = strSql & " 		and d.itemid = T.ItemID "
	strSql = strSql & " 		and d.itemoption = T.itemoption "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
	strSql = strSql & " 	and T.OutMallOrderSerial = '" & ioutmallorderserial & "' "
	strSql = strSql & " 	and T.asid is NULL "
    strSql = strSql & " 	and IsNull(T.outmallCurrState, 'B001') <> 'B008' "
    dbget.Execute strSql, affectedRow

    '// 이중매칭 삭제
    if (sellsite = "coupang") then
        strSql = " update A "
        strSql = strSql & " set A.asid = NULL "
        strSql = strSql & " from "
        strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] A "
        strSql = strSql & " 	join [db_temp].[dbo].[tbl_xSite_TMPCS] B "
        strSql = strSql & " 	on "
        strSql = strSql & " 		1 = 1 "
        strSql = strSql & " 		and A.SellSite = B.SellSite "
        strSql = strSql & " 		and A.OutMallOrderSerial = B.OutMallOrderSerial "
        strSql = strSql & " 		and A.OrgDetailKey = B.OrgDetailKey "
        strSql = strSql & " 		and A.CSDetailKey <> B.CSDetailKey "
        strSql = strSql & " 		and A.asid = B.asid "
        strSql = strSql & " where "
        strSql = strSql & " 	1 = 1 "
        strSql = strSql & " 	and A.SellSite = '" & sellsite & "' "
        strSql = strSql & " 	and A.OutMallOrderSerial = '" & ioutmallorderserial & "' "
        strSql = strSql & " 	and A.OutMallCurrState = 'B008' "
        dbget.Execute strSql
    end if

    fnMatchCs = affectedRow
end function

function fnUnmatchDeletedCS(sellsite, ioutmallorderserial)
    dim affectedRow, strSql

    strSql = " update T "
	strSql = strSql & " set T.asid = NULL "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_list] a on T.asid = a.id and a.deleteyn = 'Y' "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
	strSql = strSql & " 	and T.OutMallOrderSerial = '" & ioutmallorderserial & "' "
	strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008', 'A000') "
	strSql = strSql & " 	and T.orderserial is not NULL "
    dbget.Execute strSql

    strSql = " update T "
	strSql = strSql & " set T.asid = NULL "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
	strSql = strSql & " 	and T.OutMallOrderSerial = '" & ioutmallorderserial & "' "
	strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008', 'A000') "
	strSql = strSql & " 	and T.orderserial is not NULL "
    strSql = strSql & " 	and T.asid is not NULL "
    strSql = strSql & " 	and T.outmallCurrState = 'B008' "
    dbget.Execute strSql
end function

function MatchTenCSAsid(sellsite)
    dim strSql, affectedRows
    dim OutMallOrderSerialArr, OutMallOrderSerial

    strSql = " update T "
	strSql = strSql & " set T.asid = NULL "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
	''strSql = strSql & " 	and T.OutMallOrderSerial = '" & ioutmallorderserial & "' "
	strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008', 'A000') "
	strSql = strSql & " 	and T.orderserial is not NULL "
    strSql = strSql & " 	and T.asid is not NULL "
    strSql = strSql & " 	and T.outmallCurrState = 'B008' "
    dbget.Execute strSql

    strSql = " select distinct top 100 T.OutMallOrderSerial "
	strSql = strSql & " from "
	strSql = strSql & " [db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
	strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008', 'A000') "
    strSql = strSql & " 	and T.orderserial is not NULL "
	strSql = strSql & " 	and T.asid is NULL "
    ''strSql = strSql & " 	and T.regdate < convert(varchar(10), getdate(), 121) "
    strSql = strSql & " 	and T.regdate >= DateAdd(day, -90, getdate()) "
    strSql = strSql & " 	and IsNull(T.asidCheckDT, DateAdd(day, -1, getdate())) < DateAdd(hour, -1, getdate()) "
    ''strSql = strSql & " order by newid() "
    ''rw strSql

    OutMallOrderSerialArr = ""

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
        do until rsget.eof
            OutMallOrderSerialArr = OutMallOrderSerialArr + "," + rsget("OutMallOrderSerial")
            rsget.moveNext
        loop
	end if
	rsget.Close

    Response.CharSet="euc-kr"
    Session.codepage="949"
    Response.codepage="949"
    Response.ContentType="text/html;charset=euc-kr"

    '// git 업로드 확인
    if OutMallOrderSerialArr = "" then
        rw "내역없음"
        dbget.close() : response.end
    end if

    affectedRows = 0
    OutMallOrderSerialArr = Split(OutMallOrderSerialArr, ",")
    for i = 0 to UBound(OutMallOrderSerialArr)
        OutMallOrderSerial = OutMallOrderSerialArr(i)
        if OutMallOrderSerial <> "" then
            affectedRows = fnMatchCs(sellsite, OutMallOrderSerial)
            Call fnUnmatchDeletedCS(sellsite, OutMallOrderSerial)

            rw OutMallOrderSerial & " : " & affectedRows & " 건 반영됨"

            if affectedRows = 0 then
                strSql = " update T "
                strSql = strSql & " set T.asidCheckDT = getdate() "
	            strSql = strSql & " from "
	            strSql = strSql & " [db_temp].[dbo].[tbl_xSite_TMPCS] T "
	            strSql = strSql & " where "
	            strSql = strSql & " 	1 = 1 "
	            strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
                strSql = strSql & " 	and T.OutMallOrderSerial = '" & OutMallOrderSerial & "' "
	            strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008', 'A000') "
	            strSql = strSql & " 	and T.asid is NULL "
                dbget.Execute strSql
            end if
        end if
    next
end function

function CheckExtCsState(sellsite)
    dim strSql, affectedRows
    dim divcd, csdetailkey1, csdetailkey2
    dim divcdArr, csdetailkey1Arr, csdetailkey2Arr

    strSql = " select distinct top 100 "
    select case sellsite
        case "coupang"
            strSql = strSql & " T.divcd, T.OutMallOrderSerial as csdetailkey1, T.CSDetailKey as csdetailkey2 "
        case else
            strSql = strSql & " T.divcd, T.OutMallOrderSerial as csdetailkey1, '' as csdetailkey2 "
    end select

	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	''strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_list] a on T.asid = a.id and a.currstate = 'B007' and a.deleteyn = 'N' "		'// 접수이전 내역도 체크
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "

    if (sellsite = "coupang") then
        strSql = strSql & " 	and T.divcd in ('A004') "
    else
        strSql = strSql & " 	and T.divcd in ('A004', 'A011') "
    end if

	strSql = strSql & " 	and T.orderserial is not NULL "
	''strSql = strSql & " 	and T.regdate < convert(varchar(10), DateAdd(day, -0, getdate()), 121) "
	strSql = strSql & " 	and T.regdate >= convert(varchar(10), DateAdd(day, -80, getdate()), 121) "
	strSql = strSql & " 	and IsNull(T.outmallCheckDT, DateAdd(day, -1, getdate())) < DateAdd(hour, -1, getdate()) "
	strSql = strSql & " 	and IsNull(T.OutMallCurrState, 'B001') < 'B007' "
    ''rw strSql

    csdetailkey1Arr = ""
    csdetailkey2Arr = ""
    divcdArr = ""

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
        do until rsget.eof
            divcdArr = divcdArr & rsget("divcd") & ","
            csdetailkey1Arr = csdetailkey1Arr & rsget("csdetailkey1") & ","
            csdetailkey2Arr = csdetailkey2Arr & rsget("csdetailkey2") & ","
            rsget.moveNext
        loop
	end if
	rsget.Close

    Response.CharSet="euc-kr"
    Session.codepage="949"
    Response.codepage="949"
    Response.ContentType="text/html;charset=euc-kr"

    ''Call GetCSOrderCancel_One_coupang("coupang", "10000060357304", "184259720")

    '// git 업로드 확인
    if divcdArr = "" then
        rw "내역없음"
        dbget.close() : response.end
    end if

    affectedRows = 0
    csdetailkey1Arr = Split(csdetailkey1Arr, ",")
    csdetailkey2Arr = Split(csdetailkey2Arr, ",")
    divcdArr = Split(divcdArr, ",")

    for i = 0 to UBound(divcdArr)
        csdetailkey1 = csdetailkey1Arr(i)
        csdetailkey2 = csdetailkey2Arr(i)
        divcd = divcdArr(i)

        if Trim(divcd) <> "" then
        	select case sellsite
                case "coupang"
                    if divcd = "A004" then
                        affectedRows = GetCSOrderCancel_One_coupang(sellsite, csdetailkey1, csdetailkey2)
                        rw csdetailkey1 & " : " & affectedRows & " 건 반영됨"

                        if affectedRows = 0 then
                            strSql = " update T "
                            strSql = strSql & " set T.outmallCheckDT = getdate() "
	                        strSql = strSql & " from "
	                        strSql = strSql & " [db_temp].[dbo].[tbl_xSite_TMPCS] T "
	                        strSql = strSql & " where "
	                        strSql = strSql & " 	1 = 1 "
	                        strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
                            strSql = strSql & " 	and T.OutMallOrderSerial = '" & csdetailkey1 & "' "
	                        strSql = strSql & " 	and T.divcd = '" & divcd & "' "
	                        strSql = strSql & " 	and T.asid is not NULL "
                            dbget.Execute strSql
                        end if
                    end if
                case else
                    response.write "TEST2<br />"
            end select
        end if
    next
end function

function skstoaAPIURL()
	If application("Svr_Info") = "Dev" Then
		skstoaAPIURL = "http://dev-sel.skstoa.com"
	Else
		skstoaAPIURL = "https://open-api.skstoa.com"
	End If
end function

function skstoalinkCode()
	If application("Svr_Info") = "Dev" Then
		skstoalinkCode = "TENBYTEN"
	Else
		skstoalinkCode = "TENBYTEN"
	End If
end function

function skstoaentpCode()
	If application("Svr_Info") = "Dev" Then
		skstoaentpCode = "112644"
	Else
		skstoaentpCode = "112644"
	End If
end function

function skstoaentpId()
	If application("Svr_Info") = "Dev" Then
		skstoaentpId = "E112644"
	Else
		skstoaentpId = "E112644"
	End If
end function

function skstoaentpPass()
	Dim skstoaStrSql
	skstoaStrSql = ""
	skstoaStrSql = skstoaStrSql & " SELECT TOP 1 isnull(iniVal, '') as iniVal "
	skstoaStrSql = skstoaStrSql & " FROM db_etcmall.dbo.tbl_outmall_ini " & VbCRLF
	skstoaStrSql = skstoaStrSql & " where mallid='skstoa' " & VbCRLF
	skstoaStrSql = skstoaStrSql & " and inikey='pass'"
	rsget.CursorLocation = adUseClient
	rsget.Open skstoaStrSql, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.Eof then
		skstoaentpPass	= rsget("iniVal")
	end if
	rsget.close
end function
%>
