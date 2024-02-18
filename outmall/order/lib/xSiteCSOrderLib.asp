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
		response.write "ERROR : ��ſ���"
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
					gubunname = "�ܼ�����"
				case "ChangeChoice"
					gubunname = "���ú���"
				case "DelayDelivery"
					gubunname = "�������"
				case "DamagedItem"
					gubunname = "��ǰ�ҷ�"
				case "ShippingMistake"
					gubunname = "�����"
				case "InformationMistake"
					gubunname = "��ǰ����Ʋ��"
				case "OutofStock"
					gubunname = "ǰ��"
				case "CouponNotAccept"
					gubunname = "����������"
				case "Etc"
					gubunname = "��Ÿ"
				case else
					gubunname = Replace(gubunname, "'", "")
			end select
			OutMallRegDate = Left(obj.getAttribute("RequestClaimDate"), 10)
			itemno = 0			'// �������� ������� ����, �������

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

				''�ֹ� �Է� ���� ������ ���� ����
				strSql = " update c "
				strSql = strSql + " set matchState='D'"
				strSql = strSql + " from db_temp.dbo.tbl_xSite_TMPOrder c "
				strSql = strSql + " WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
				strSql = strSql + " and orderserial is NULL"
				''response.write strSql & "<br />"
				dbget.Execute strSql

				'' CS ����������. ������Ʈ
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
		rw "���öȸ CS�Է°Ǽ�:"&iInputCnt
	else
		rw "�ֹ���� CS�Է°Ǽ�:"&iInputCnt
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
		response.write "ERROR : ��ſ���"
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
					gubunname = "�ܼ�����"
				case "ChangeChoice"
					gubunname = "���ú���"
				case "DelayDelivery"
					gubunname = "�������"
				case "DamagedItem"
					gubunname = "��ǰ�ҷ�"
				case "ShippingMistake"
					gubunname = "�����"
				case "InformationMistake"
					gubunname = "��ǰ����Ʋ��"
				case "OutofStock"
					gubunname = "ǰ��"
				case "CouponNotAccept"
					gubunname = "����������"
				case "Etc"
					gubunname = "��Ÿ"
				case else
					gubunname = Replace(gubunname, "'", "")
			end select
			OutMallRegDate = Left(obj.getAttribute("RequestClaimDate"), 10)
			itemno = 0			'// �������� ������ǰ ����, ���ι�ǰ

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

				'' CS ����������. ������Ʈ
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
		rw "��ǰöȸ CS�Է°Ǽ�:"&iInputCnt
	else
		rw "��ǰ CS�Է°Ǽ�:"&iInputCnt
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
		response.write "ERROR : ��ſ���"
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
					gubunname = "�ܼ�����"
				case "ChangeChoice"
					gubunname = "���ú���"
				case "DelayDelivery"
					gubunname = "�������"
				case "DamagedItem"
					gubunname = "��ǰ�ҷ�"
				case "ShippingMistake"
					gubunname = "�����"
				case "InformationMistake"
					gubunname = "��ǰ����Ʋ��"
				case "OutofStock"
					gubunname = "ǰ��"
				case "CouponNotAccept"
					gubunname = "����������"
				case "Etc"
					gubunname = "��Ÿ"
				case else
					gubunname = Replace(gubunname, "'", "")
			end select
			OutMallRegDate = Left(obj.getAttribute("RequestClaimDate"), 10)
			itemno = 0			'// �������� ������ȯ ����, ���α�ȯ

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

				'' CS ����������. ������Ʈ
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
		rw "��ȯöȸ CS�Է°Ǽ�:"&iInputCnt
	else
		rw "��ȯ CS�Է°Ǽ�:"&iInputCnt
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
		response.write "ERROR : ��ſ���"
		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML replace(objData,"&","��")

	Set obj = xmlDOM.selectSingleNode("/ORDER_LIST/RESULT/CODE")

	if obj is Nothing then
		''response.write "�������� : ����"
		GetCSOrderAll_interpark = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	if (obj.text <> "000") then
		response.write "ERROR : �� �� ���� ����"
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
				divcd = 1/0		'// ����
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
						divcd = 1/0		'// ����
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

				'' CS ����������. ������Ʈ
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

	rw "CS����Է°Ǽ�1 : "&iInputCnt

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
		response.write "ERROR : ��ſ���"
		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML replace(objData,"&","��")

	Set obj = xmlDOM.selectSingleNode("/ORDER_LIST/RESULT/CODE")

	if obj is Nothing then
		''response.write "�������� : ����"
		GetCSOrderChgRet_interpark = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	if (obj.text <> "000") then
		response.write "ERROR : �� �� ���� ����"
		GetCSOrderChgRet_interpark = False
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	Set objArr = xmlDOM.selectNodes("/ORDER_LIST/ORDER")

	iInputCnt = 0
	for i = 0 to objArr.length - 1
		set obj = objArr.item(i)
		OrgOutMallOrderSerial = obj.selectSingleNode("CLM_NO").text		'���ֹ���ȣ�� �ƴϴ�. ORD_NO�� ���ֹ���ȣ�̴�. CLM_NO ������ ����̸�Ī �ֹ���ȣ ã�ƾ� �Ѵ�.
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
				divcd = 1/0		'// ����
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

				'' CS ����������. ������Ʈ
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
	rw "CS�Է°Ǽ�2 : "&iInputCnt
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
'RECEIPT : ����
'PROGRESS : ����
'SUCCESS : �Ϸ�
'REJECT : �Ұ�
'CANCEL : ���
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
							CSDetailKey			= strObjValues.get(i).exchangeId				'��ȯ ���̵�
							OutMallOrderSerial	= strObjValues.get(i).orderId					'�ֹ���ȣ
							'rw strObjValues.get(i).vendorId									'���� ���̵�
							'rw strObjValues.get(i).orderDeliveryStatusCode	'�ֹ���ۻ��� | ACCEPT : �����Ϸ�, INSTRUCT : ��ǰ�غ���, DEPARTURE : �������, DELIVERING : �����, FINAL_DELIVERY : ��ۿϷ�, NONE_TRACKING : ��ü�������(��� ���� ������), �����Ұ�
							'rw strObjValues.get(i).exchangeStatus			'��ȯ���� | RECEIPT : ����, PROGRESS : ����, SUCCESS : �Ϸ�, REJECT : �Ұ�, CANCEL : öȸ
							'rw strObjValues.get(i).referType				'������� | VENDOR : ����, CS_CENTER : CS, WEB_PC : �� PC, WEB_MOBILE : �� �����
							'rw strObjValues.get(i).faultType				'��å | COUPANG : ���ΰ���, VENDOR : ��ü����, CUSTOMER : ������, GENERAL : �Ϲ�
							'rw strObjValues.get(i).exchangeAmount			'��ȯ��ۺ�
							'rw strObjValues.get(i).reason					'��ȯ�������� | ���̻� ������� �ʴ� �ʵ�
							'rw strObjValues.get(i).reasonCode				'��ȯ�����ڵ� | DEFECT : ����, WRONGITEM : �����, OMISSION : ����, OPTIONCHANGE : �ɼǺ���, ETC : ��Ÿ, BROKEN : �ļ�, ADDRESSCHANGE : ���������, LOST : ��ǰ�н�
							gubunname			= strObjValues.get(i).reasonCodeText			'��ȯ��������
							'rw strObjValues.get(i).reasonEtcDetail			'��ȯ�����󼼼���
							'rw strObjValues.get(i).cancelReason			'��ȯöȸ����
							'rw strObjValues.get(i).createdByType			'���� ����� ���� | CUSTOMER : ��, COUNSELOR : ����, COUPANG : ��������, VENDOR : ��ü, ETC : ��Ÿ
							OutMallRegDate		= strObjValues.get(i).createdAt				'����Ͻ�
							OutMallRegDate		= Replace(OutMallRegDate, "T", " ")
							'rw strObjValues.get(i).modifiedByType			'������ | CUSTOMER : ��, COUNSELOR : ����, COUPANG : ��������, VENDOR : ��ü, ETC : ��Ÿ, TRACKING : �������
							'rw strObjValues.get(i).modifiedAt				'�����Ͻ�
							set strObjValuesexchangeItems = strObjValues.get(i).exchangeItemDtoV1s
								For j=0 to strObjValuesexchangeItems.length-1
									''OrgDetailKey = strObjValuesexchangeItems.get(j).exchangeItemId	'��ȯ ��ǰ ���̵�
									OrgDetailKey = strObjValuesexchangeItems.get(j).orderItemId				'���ֹ� ������ID
									'rw strObjValuesexchangeItems.get(j).orderItemUnitPrice			'���ֹ� ������ �ܰ�
									'rw strObjValuesexchangeItems.get(j).orderItemName				'���ֹ� ������ ��
									'rw strObjValuesexchangeItems.get(j).orderPackageId				'���ֹ� ��Ű�� ID
									'rw strObjValuesexchangeItems.get(j).orderPackageName			'���ֹ� ��Ű����
									'rw strObjValuesexchangeItems.get(j).targetItemId				'��ȯ ������ ID
									'rw strObjValuesexchangeItems.get(j).targetItemUnitPrice		'��ȯ ������ �ܰ�
									'rw strObjValuesexchangeItems.get(j).targetItemName				'��ȯ ������ ��
									'rw strObjValuesexchangeItems.get(j).targetPackageId			'��ȯ ��Ű�� ID
									'rw strObjValuesexchangeItems.get(j).targetPackageName			'��ȯ ��Ű�� ��
									itemno = strObjValuesexchangeItems.get(j).quantity				'��ȯ ����
									'rw strObjValuesexchangeItems.get(j).orderItemDeliveryComplete	'True /False : ���ֹ� ������ ��ۿϷ� ����
									'rw strObjValuesexchangeItems.get(j).orderItemReturnComplete	'True /False : ���ֹ� ������ ��ǰ�Ϸ� ����
									'rw strObjValuesexchangeItems.get(j).targetItemDeliveryComplete	'True /False : ��ȯ ������ ��ۿϷ� ����
									'rw strObjValuesexchangeItems.get(j).createdAt					'���� �Ͻ�
									'rw strObjValuesexchangeItems.get(j).modifiedAt					'���� �Ͻ�
									'rw strObjValuesexchangeItems.get(j).originalShipmentBoxId		'����۹�ȣ
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

										'' CS ����������. ������Ʈ
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

	rw "CS ��ȯ �Է°Ǽ�:"&iInputCnt

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
'RU : ���������û
'CC : ��ǰ�Ϸ�
'PR : ����Ȯ�ο�û
'UC : ��ǰ����
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
							'// *�����ڰ� ����غ��� ~ ����� ���̿� ����� ���̽��� ��ǰ���� �ٷ�� �ֽ��ϴ�.
							'// ================================================

                            OrgOutMallOrderSerial	= strObjValues.get(i).orderId				'�ֹ���ȣ
							CSDetailKey				= strObjValues.get(i).receiptId				'���(��ǰ)������ȣ
							OutMallOrderSerial		= strObjValues.get(i).orderId				'�ֹ���ȣ
							'rw strObjValues.get(i).paymentId				'������ȣ
                            'rw strObjValues.get(i).receiptType									'������� RETURN or CANCEL
                            if (strObjValues.get(i).receiptType = "CANCEL") then
                                divcd = "A008"
                            elseif (strObjValues.get(i).receiptType = "RETURN") then
                                divcd = "A004"
                            else
                                divcd = "AXXX"
                            end if
							'rw strObjValues.get(i).receiptStatus			'���(��ǰ)������� | RELEASE_STOP_UNCHECKED : ���������û, RETURNS_UNCHECKED : ��ǰ����, VENDOR_WAREHOUSE_CONFIRM : �԰�Ϸ�, REQUEST_COUPANG_CHECK : ����Ȯ�ο�û, RETURNS_COMPLETED : ��ǰ�Ϸ�
							OutMallRegDate		= strObjValues.get(i).createdAt				'���(��ǰ) �����ð�
							OutMallRegDate		= Replace(OutMallRegDate, "T", " ")
							'rw strObjValues.get(i).modifiedAt				'���(��ǰ) ���� ���� ����ð�
							OrderName			= strObjValues.get(i).requesterName				'��ǰ ��û�� �̸�
							OrderHpNo			= strObjValues.get(i).requesterPhoneNumber		'��ǰ ��û�� ��ȭ��ȣ
							gubunname			= strObjValues.get(i).cancelReasonCategory1	'��ǰ ���� ī�װ�1
							'rw strObjValues.get(i).cancelReasonCategory2	'��ǰ ���� ī�װ�2
							'rw strObjValues.get(i).cancelReason				'��һ��� �󼼳���
							'rw strObjValues.get(i).cancelCountSum			'�� ��Ҽ���
							'rw strObjValues.get(i).returnDeliveryId			'��ǰ��۹�ȣ
							'rw strObjValues.get(i).returnDeliveryType		'ȸ������ | �����ù�, �����ù�, �������
							'rw strObjValues.get(i).releaseStopStatus		'�������ó������ | ��ó��, ó��(�̹����), ó��(�������), �ڵ�ó��(�̹����), ����
							'rw strObjValues.get(i).enclosePrice			'������ۺ�
							'rw strObjValues.get(i).faultByType				'��åŸ�� | Coupang ���� : COUPANG, ���»� ���� : VENDOR, �� ���� : CUSTOMER, �������� : WMS, �Ϲ� : GENERAL
							'rw strObjValues.get(i).preRefund				'��ȯ�Ҿߺ�
							'rw strObjValues.get(i).completeConfirmDate		'�Ϸ�Ȯ������ | ��Ʈ��Ȯ�� : VENDOR_CONFIRM, ��Ȯ�� : UNDEFINED, CS �븮Ȯ�� : CS_CONFIRM, CS �ս�Ȯ�� : CS_LOSS_CONFIRM
							'rw strObjValues.get(i).completeConfirmType		'�Ϸ�Ȯ�νð�
							set strObjValuesReturnItems = strObjValues.get(i).returnItems
								For j=0 to strObjValuesReturnItems.length-1
									'rw strObjValuesReturnItems.get(j).vendorItemPackageId		'����ȣ
									'rw strObjValuesReturnItems.get(j).vendorItemPackageName	'����
									OrgDetailKey = strObjValuesReturnItems.get(j).vendorItemId	'���������۹�ȣ
									'rw strObjValuesReturnItems.get(j).vendorItemName			'���������۸�
									itemno = strObjValuesReturnItems.get(j).purchaseCount		'��� ����
									'rw strObjValuesReturnItems.get(j).cancelCount				'�ֹ� ����
									'rw strObjValuesReturnItems.get(j).shipmentBoxId			'�� ��۹�ȣ
									'rw strObjValuesReturnItems.get(j).sellerProductId			'��ü��ϻ�ǰ��ȣ
									'rw strObjValuesReturnItems.get(j).sellerProductName		'��ü��ϻ�ǰ

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

										'' CS ����������. ������Ʈ
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
		rw "CS ��� �Է°Ǽ�:"&iInputCnt
	Else
		rw "CS ��ǰ �Է°Ǽ�:"&iInputCnt
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
						OutMallRegDate = LEFT(SubNodes.getElementsByTagName("createDt")(0).text,10)		'Ŭ���� ��û �Ͻ�
						beasongNum11st = SubNodes.getElementsByTagName("dlvNo")(0).text					'��۹�ȣ
						'rw SubNodes.getElementsByTagName("ordCnDtlsRsn")(0).text						'�����ڵ忡 ���� �󼼳���
						itemno = SubNodes.getElementsByTagName("ordCnQty")(0).text						'Ŭ���� ����
						'rw SubNodes.getElementsByTagName("ordCnMnbdCd")(0).text						'Ŭ���� �����ü | 01 : ������, 02 : �Ǹ���
						ordCnRsnCd = SubNodes.getElementsByTagName("ordCnRsnCd")(0).text				'Ŭ���� �����ڵ� | 00 : �����ü ������ : ������ ���Ա� ���, 04 : �����ü ������ : �Ǹ����� ��� ó���� ����, 06 : �����ü ������ : �Ǹ����� ��ǰ ������ �߸��� �����ü �Ǹ��� : ��� ���� ����, 07 : �����ü ������ : ���� ��ǰ ���ֹ�(�ֹ���������) �����ü �Ǹ��� : ��ǰ/���� ���� �߸� �Է�, 08 : �����ü ������ : �ֹ���ǰ�� ǰ��/������ �����ü �Ǹ��� : ��ǰ ǰ��(��ü�ɼ�), 09 : �����ü ������ : 11���� �� �ٸ� ��ǰ���� ���ֹ� �����ü �Ǹ��� : �ɼ� ǰ��(�ش�ɼ�), 10 : �����ü ������ : Ÿ����Ʈ ��ǰ �ֹ� �����ü �Ǹ��� : ������, 11 : �����ü ������ : ��ǰ�� �̻������ ���� �ǻ� ������, 12 : �����ü ������ : ��Ÿ(������ å�ӻ���), 13 : �����ü ������ : ��Ÿ(�Ǹ��� å�ӻ���), 99 : �����ü ������ : ��Ÿ, 14 : �����ǻ� ������, 15 : ����/������/�ֹ����� ����, 16 : �ٸ� ��ǰ �߸� �ֹ�, 17 : ����������� ���, 18 : ��ǰǰ��, ������
						select case ordCnRsnCd
							case "10", "11", "14", "16"
								gubunname = "�ܼ�����"
							case "15"
								gubunname = "���ú���"
							case "04", "17"
								gubunname = "�������"
							case "06", "07"
								gubunname = "��ǰ����Ʋ��"
							case "08", "09", "18"
								gubunname = "ǰ��"
							case "00", "12", "13", "99"
								gubunname = "��Ÿ"
							case else
								gubunname = Replace(gubunname, "'", "")
						end select
						ordCnStatCd = SubNodes.getElementsByTagName("ordCnStatCd")(0).text				'Ŭ���� ���� | 01 : ��ҿ�û, 02 : ��ҿϷ�
						OutMallOrderSerial = SubNodes.getElementsByTagName("ordNo")(0).text				'11���� �ֹ���ȣ
						CSDetailKey = SubNodes.getElementsByTagName("ordPrdCnSeq")(0).text				'�ܺθ� Ŭ���� ��ȣ
						OrgDetailKey = SubNodes.getElementsByTagName("ordPrdSeq")(0).text				'�ֹ�����

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

							'' CS ����������. ������Ʈ
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
						'rw SubNodes.getElementsByTagName("prdNo")(0).text			'��ǰ��ȣ
						'rw SubNodes.getElementsByTagName("slctPrdOptNm")(0).text	'Ŭ���� �ɼǸ�
						'rw SubNodes.getElementsByTagName("referSeq")(0).text		'��Ŭ��üũ�ƿ� �ֹ��ڵ�
					Next
				Set Nodes = nothing
			Set xmlDOM = nothing
		End If
	SET objXML = nothing
	Select Case csGubun
		Case "ClaimReady"
			rw "CS ��ҿ�û �Է°Ǽ�:"&iInputCnt
		Case "ClaimDone"
			rw "CS ��ҿϷ� �Է°Ǽ�:"&iInputCnt
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
						' rw SubNodes.getElementsByTagName("affliateBndlDlvSeq")(0).text 			'���ᱳȯ ���� | ���� �̻�� �ʵ� 0 : ���ᱳȯ, 1 : �Ϲݱ�ȯ(����)
						' rw SubNodes.getElementsByTagName("appmtDlvCst")(0).text					'11���� ������ǰ �ù��
						' rw SubNodes.getElementsByTagName("clmDlvCstMthd")(0).text					'������� | 01 : �ſ�ī��, 02 : ����Ʈ, 03 : �ڽ��� ����, 04 : �Ǹ��ڿ��� �����۱�, null : Ŭ���� ������ �Ǹ����� ���
						' rw SubNodes.getElementsByTagName("clmLstDlvCst")(0).text					'��ȯ��ۺ�
						' rw SubNodes.getElementsByTagName("clmReqCont")(0).text					'�����ڵ忡 ���� �󼼳���
						itemno = SubNodes.getElementsByTagName("clmReqQty")(0).text					'Ŭ���� ����
						clmReqRsn = SubNodes.getElementsByTagName("clmReqRsn")(0).text				'Ŭ���� �����ڵ�
						select case clmReqRsn
							case "101"
								gubunname = "�ܼ�����"
							case "ChangeChoice"
								gubunname = "���ú���"
							case "112"
								gubunname = "�������"
							case "111", "207", "216"
								gubunname = "��ǰ�ҷ�"
							case "108", "208"
								gubunname = "�����"
							case "105", "210"
								gubunname = "��ǰ����Ʋ��"
							case "OutofStock"
								gubunname = "ǰ��"
							case "CouponNotAccept"
								gubunname = "����������"
							case "113", "114", "115", "116", "117", "118", "119", "198", "199", "206", "209", "212", "213", "211", "214", "215", "217"
								gubunname = "��Ÿ"
							case else
								gubunname = Replace(gubunname, "'", "")
						end select

						CSDetailKey = SubNodes.getElementsByTagName("clmReqSeq")(0).text			'�ܺθ� Ŭ���� ��ȣ
						' rw SubNodes.getElementsByTagName("clmStat")(0).text						'Ŭ���� ���� | 103 : ����������, 104 : ��ǰ����, 105 : ��ǰ��û, 106 : ��ǰ�Ϸ�, 107 : ��ǰ�ź�, 108 : ��ǰöȸ, 109 : ��ǰ�ϷẸ��, 201 : ��ȯ��û, 212 : ��ȯ����, 214 : ��ȯ����, 221 : ��ȯ�߼ۿϷ�, 232 : ��ȯ�ź�, 233 : ��ȯöȸ, 301 : ��������, 302 : ���ۿϷ�
						' rw SubNodes.getElementsByTagName("dlvCstRespnClf")(0).text
						' rw SubNodes.getElementsByTagName("dlvEtprsCd")(0).text					'�����ù���ڵ� | 00034 : CJ�������, 00011 : �����ù� ���...�Ŵ�������
						' rw SubNodes.getElementsByTagName("dlvNo")(0).text
						' rw SubNodes.getElementsByTagName("exchBaseAddr")(0).text					'��ȯ��ǰ ������ �⺻�ּ�
						' rw SubNodes.getElementsByTagName("exchDtlsAddr")(0).text					'��ȯ��ǰ ������ ���ּ�
						' rw SubNodes.getElementsByTagName("exchMailNo")(0).text					'��ȯ��ǰ ������ �����ȣ
						' rw SubNodes.getElementsByTagName("exchMailNoSeq")(0).text					'��ȯ��ǰ ������ �����ȣ ����
						' rw SubNodes.getElementsByTagName("exchNm")(0).text						'��ȯ��ǰ ������ �̸�
						' rw SubNodes.getElementsByTagName("exchPrtblTel")(0).text					'��ȯ��ǰ ������ �޴�����ȣ
						' rw SubNodes.getElementsByTagName("exchTlphnNo")(0).text					'��ȯ��ǰ ������ ��ȭ��ȣ
						' rw SubNodes.getElementsByTagName("exchTypeAdd")(0).text					'��ȯ��ǰ ������ �ּ� ���� | 01 : ������, 02 : ���θ�
						' rw SubNodes.getElementsByTagName("exchTypeBilNo")(0).text					'��ȯ��ǰ ������ �ǹ�������ȣ
						' rw SubNodes.getElementsByTagName("freeGiftNo")(0).text
						' rw SubNodes.getElementsByTagName("freeGiftQty")(0).text
						' rw SubNodes.getElementsByTagName("kglUseYn")(0).text
						' rw SubNodes.getElementsByTagName("optName")(0).text						'�ɼǸ�
						' rw SubNodes.getElementsByTagName("ordNm")(0).text							'������ �̸�
						OutMallOrderSerial = SubNodes.getElementsByTagName("ordNo")(0).text			'11���� �ֹ���ȣ
						OrgDetailKey = SubNodes.getElementsByTagName("ordPrdSeq")(0).text			'�ֹ�����
						' rw SubNodes.getElementsByTagName("ordPrtblTel")(0).text					'������ �޴�����ȣ
						' rw SubNodes.getElementsByTagName("ordTlphnNo")(0).text					'������ ��ȭ��ȣ
						' rw SubNodes.getElementsByTagName("prdNo")(0).text							'��ǰ��ȣ
						' rw SubNodes.getElementsByTagName("rcvrBaseAddr")(0).text					'������ �⺻�ּ�
						' rw SubNodes.getElementsByTagName("rcvrDtlsAddr")(0).text					'������ ���ּ�
						' rw SubNodes.getElementsByTagName("rcvrMailNo")(0).text					'������ �����ȣ
						' rw SubNodes.getElementsByTagName("rcvrMailNoSeq")(0).text					'������ �����ȣ ����
						' rw SubNodes.getElementsByTagName("rcvrTypeAdd")(0).text					'������ �ּ� ���� 01 : ������, 02 : ���θ�
						' rw SubNodes.getElementsByTagName("rcvrTypeBilNo")(0).text					'������ �ǹ�������ȣ
						OutMallRegDate = LEFT(SubNodes.getElementsByTagName("reqDt")(0).text,10)	'Ŭ���� ��û �Ͻ�
						' rw SubNodes.getElementsByTagName("twMthd")(0).text						'��ȯ��ǰ �߼۹�� | 02 : �����߼�, 06 : 11���� ������ǰ�ù�, 07 : �Ǹ��� ������ǰ�ù�,
						' rw SubNodes.getElementsByTagName("twPrdInvcNo")(0).text					'���ż����ȣ

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
	rw "CS ��ȯ �Է°Ǽ�:"&iInputCnt
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
						' rw SubNodes.getElementsByTagName("addDlvCst")(0).text						'�߰���ۺ�
						' rw SubNodes.getElementsByTagName("affliateBndlDlvSeq")(0).text			'�����ǰ ���� | ���� �̻�� �ʵ�  0 : �����ǰ, 1 : �Ϲݹ�ǰ(����)
						' rw SubNodes.getElementsByTagName("appmtDlvCst")(0).text					'11���� ������ǰ �ù��
						' rw SubNodes.getElementsByTagName("clmDlvCstMthd")(0).text					'������� | 01 : �ſ�ī��, 02 : ����Ʈ, 03 : �ڽ��� ����, 04 : �Ǹ��ڿ��� �����۱�, null : Ŭ���� ������ �Ǹ����� ���
						' rw SubNodes.getElementsByTagName("clmLstDlvCst")(0).text					'��ǰ��ۺ�
						' rw SubNodes.getElementsByTagName("clmReqCont")(0).text					'�����ڵ忡 ���� �󼼳���
						itemno = SubNodes.getElementsByTagName("clmReqQty")(0).text					'Ŭ���� ����
						clmReqRsn = SubNodes.getElementsByTagName("clmReqRsn")(0).text				'Ŭ���� �����ڵ�
						select case clmReqRsn
							case "101"
								gubunname = "�ܼ�����"
							case "112"
								gubunname = "�������"
							case "111", "207", "122", "123"
								gubunname = "��ǰ�ҷ�"
							case "108", "208"
								gubunname = "�����"
							case "105", "210"
								gubunname = "��ǰ����Ʋ��"
							case "110", "114", "115", "116", "117", "118", "119", "121", "198", "199", "206", "209", "212", "213", "113", "211", "214"
								gubunname = "��Ÿ"
							case else
								gubunname = Replace(gubunname, "'", "")
						end select

						CSDetailKey = SubNodes.getElementsByTagName("clmReqSeq")(0).text			'�ܺθ� Ŭ���� ��ȣ
						' rw SubNodes.getElementsByTagName("clmStat")(0).text						'Ŭ���� ���� | 103 : ����������, 104 : ��ǰ����, 105 : ��ǰ��û, 106 : ��ǰ�Ϸ�, 107 : ��ǰ�ź�, 108 : ��ǰöȸ, 109 : ��ǰ�ϷẸ��, 201 : ��ȯ��û, 212 : ��ȯ����, 214 : ��ȯ����, 221 : ��ȯ�߼ۿϷ�, 232 : ��ȯ�ź�, 233 : ��ȯöȸ, 301 : ��������, 302 : ���ۿϷ�
						' rw SubNodes.getElementsByTagName("dlvCstRespnClf")(0).text				'��ۺ� �δ㿩�� | 01 : ������, 02 : �Ǹ���
						' rw SubNodes.getElementsByTagName("dlvEtprsCd")(0).text					'�����ù���ڵ� | 00034 : CJ�������, 00011 : �����ù� ���...�Ŵ�������
						beasongNum11st = SubNodes.getElementsByTagName("dlvNo")(0).text				'��۹�ȣ
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
						' rw SubNodes.getElementsByTagName("kglUseYn")(0).text						'KGL(�ؿܹ��) �ù������� �⺻�� &#39;N&#39; �ؿܹ�� ��ǰ�� �ش��
						' rw SubNodes.getElementsByTagName("optName")(0).text						'�ɼǸ�
						' rw SubNodes.getElementsByTagName("ordNm")(0).text							'������ �̸�
						OutMallOrderSerial = SubNodes.getElementsByTagName("ordNo")(0).text			'11���� �ֹ���ȣ
						OrgDetailKey = SubNodes.getElementsByTagName("ordPrdSeq")(0).text			'�ֹ�����
						' rw SubNodes.getElementsByTagName("ordPrtblTel")(0).text					'������ �޴�����ȣ
						' rw SubNodes.getElementsByTagName("ordTlphnNo")(0).text					'������ ��ȭ��ȣ
						' rw SubNodes.getElementsByTagName("prdNo")(0).text							'��ǰ��ȣ
						' rw SubNodes.getElementsByTagName("rcvrBaseAddr")(0).text					'������ �⺻�ּ�
						' rw SubNodes.getElementsByTagName("rcvrDtlsAddr")(0).text					'������ ���ּ�
						' rw SubNodes.getElementsByTagName("rcvrMailNo")(0).text					'������ �����ȣ
						' rw SubNodes.getElementsByTagName("rcvrMailNoSeq")(0).text					'������ �����ȣ ����
						' rw SubNodes.getElementsByTagName("rcvrTypeAdd")(0).text					'������ �ּ� ���� | 01 : ������, 02 : ���θ�
						' rw SubNodes.getElementsByTagName("rcvrTypeBilNo")(0).text					'������ �ǹ�������ȣ
						OutMallRegDate = LEFT(SubNodes.getElementsByTagName("reqDt")(0).text,10)	'Ŭ���� ��û �Ͻ�
						' rw SubNodes.getElementsByTagName("twMthd")(0).text						'��ǰ��ǰ �߼۹�� | 02 : �����߼�, 06 : 11���� ������ǰ�ù�, 07 : �Ǹ��� ������ǰ�ù�, 09 : �Ǹ�ó ����
						' rw SubNodes.getElementsByTagName("twPrdInvcNo")(0).text					'���ż����ȣ

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
	rw "CS ��ǰ �Է°Ǽ�:"&iInputCnt
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
			'prgrGb | P0:�����, P1:�������, P2:���, P3:��ۿϷ�
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
								CSDetailKey			= obj1.get(i).dlvstNo				'������ù�ȣ
								OutMallOrderSerial	= obj1.get(i).ordNo					'�ֹ���ȣ
								OrgDetailKey		= obj1.get(i).ordPtcSeq				'�ֹ��Ϸù�ȣ
								dlvCnclYn			= obj1.get(i).dlvCnclYn				'�����ҿ���
								gubunname			= obj1.get(i).dlvCnclNm				'������
								itemno				= obj1.get(i).custCnclQty			'����Ҽ���
								OutMallRegDate 		= LEFT(obj1.get(i).ptcOrdDtm, 10)		'���ֹ�����

								If dlvCnclYn = "Y" Then	'�����ҿ��ΰ� Y
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
		'prgrGb | P0:�����, P1:�������, P2:���, P3:��ۿϷ�
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
							CSDetailKey			= obj1.get(i).dlvstNo				'������ù�ȣ
							OutMallOrderSerial	= obj1.get(i).ordNo					'�ֹ���ȣ
							OrgDetailKey		= obj1.get(i).befDlvstNo			'���� ������ù�ȣ
							dlvTypeGbcd			= obj1.get(i).dlvTypeGbcd			'�������(�ֹ�����) | 30:��ǰȸ��, 45:��ȯȸ��, 65:�κб�ȯȸ��
							gubunname			= obj1.get(i).custVenPaonMsg		'�����»����޸޽���
							itemno				= obj1.get(i).prrgQty				'������ | ������ü���-��Ҽ���-��ۼ���
							OutMallRegDate 		= LEFT(obj1.get(i).oshpReqnDt, 10)	'����û��

							Select Case dlvTypeGbcd
								Case "30"		divcd = "A004"
								Case "45"		divcd = "A000"
								Case "65"		divcd = "A000"
							End Select

							If itemno = "" Then itemno = "0"

							If dlvTypeGbcd <> "" Then	'�������(�ֹ�����)�� ���� �ִٸ�
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
		response.write "ERROR : ��ſ���"
		dbget.close : response.end
	end if

	Set strObj = JSON.parse(objXML.responseText)
	retMsg = strObj.message

	successCount = 0

	if (retMsg = "����") then
		Set items = strObj.outPutValue.data.claim
		For i = 0 to items.length - 1
			Set item = items.get(i)

			OutMallOrderSerial = item.bundleNo
			CSDetailKey = item.claimBundleNo
			dlvTypeGbcd = item.claimType
			OutMallRegDate = item.requestDate
			gubunname = LEFT(item.claimReason, 10)

			Select Case dlvTypeGbcd
				Case "��ȯ"		divcd = "A000"
				Case "���"		divcd = "A008"
				Case "��ǰ"		divcd = "A004"
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
			rw "CS " & canceDoneStr & " �Է°Ǽ�:" & successCount
		Else
			rw "CS " & csGubun & " �Է°Ǽ�:" & successCount
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
		response.write "ERROR : ��ſ���"
	end if

	Set strObj = JSON.parse(objXML.responseText)
	retMsg = strObj.message

	successCount = 0

	if (retMsg = "����") then
		Set items = strObj.outPutValue.data.claim
		For i = 0 to items.length - 1
			Set item = items.get(i)

			OutMallOrderSerial = item.bundleNo
			CSDetailKey = item.claimBundleNo
			dlvTypeGbcd = item.claimType
			OutMallRegDate = item.requestDate
			gubunname = LEFT(item.claimReason, 10)

			Select Case dlvTypeGbcd
				Case "��ȯ"		divcd = "A000"
				Case "���"		divcd = "A008"
				Case "��ǰ"		divcd = "A004"
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

		rw "CS " & csGubun & " �Է°Ǽ�:" & successCount
	else
		response.write "ERROR : " & retMsg
	end if
End Function

CONST UPCHECODE = "A5703"								'��ü�ڵ�
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
			response.write "ERROR : ��ſ���"
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
						OutMallOrderSerial	= SubNodes.getElementsByTagName("OrdNum")(0).Text				'����Ŭ�� �ֹ���ȣ
						OrgDetailKey		= SubNodes.getElementsByTagName("OrdNum_Nm")(0).Text			'����Ŭ�� �ֹ�����
						itemno				= SubNodes.getElementsByTagName("Qty")(0).Text					'��� ����
						CSDetailKey			= SubNodes.getElementsByTagName("ClaimSeq")(0).Text				'��� ��� ���� ��ȣ
						ClaimMemo			= SubNodes.getElementsByTagName("ClaimMemo")(0).Text			'��� ��� ����
						OutMallRegDate		= LEFT(SubNodes.getElementsByTagName("RegYMD")(0).Text, 10)		'��� ����

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

							'' CS ����������. ������Ʈ
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
	rw "��� CS�Է°Ǽ�:"&iInputCnt
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
			response.write "ERROR : ��ſ���"
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
						OutMallOrderSerial	= SubNodes.getElementsByTagName("OrdNum")(0).Text				'����Ŭ�� �ֹ���ȣ
						OrgDetailKey		= SubNodes.getElementsByTagName("OrdNum_Nm")(0).Text			'����Ŭ�� �ֹ�����
						itemno				= SubNodes.getElementsByTagName("Qty")(0).Text					'��ǰ ����
						CSDetailKey			= SubNodes.getElementsByTagName("ClaimSeq")(0).Text				'��ǰ ��� ���� ��ȣ
						OutMallRegDate		= LEFT(SubNodes.getElementsByTagName("RegYMD")(0).Text, 10)		'��ǰ �����
						ClaimReasonCd		= SubNodes.getElementsByTagName("Claim_ReasonCd")(0).Text		'��ǰ ���� �ڵ�
						ClaimReason			= SubNodes.getElementsByTagName("Claim_Reason")(0).Text			'��ǰ ����

						select case ClaimReasonCd
							case "1"
								gubunname = "�ҷ�"
							case "2"
								gubunname = "������"
							case "3"
								gubunname = "��������"
							case "4"
								gubunname = "��ǰ��������ġ"
							case "5"
								gubunname = "������"
							case "6"
								gubunname = "��Ÿ"
							case "7"
								gubunname = "�߰�����"
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

							'' CS ����������. ������Ʈ
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
	rw "��ǰ CS�Է°Ǽ�:"&iInputCnt
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

	response.write "�Ǽ�(" & UBound(objArr) + 1 & ") " & "<br />"

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
		response.write "ERROR : ��ſ���"
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
		response.write "���� : ����"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	if CLng(xmlDOM.getElementsByTagName("n:ReturnedDataCount").item(0).text) <> (UBound(objArr) + 1) then
		response.write "�Ǽ� ����ġ ���� : ����"
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
		if (csGubun = "CANCEL_REQUESTED") OR (csGubun = "CANCELED") then	'2019-07-23 ������ CANCEL_REQUESTED�� ����
			divcd = "A008"
			OutMallRegDate = Left(obj.selectSingleNode("n:CancelInfo/n:ClaimRequestDate").text, 10)

			gubunname = obj.selectSingleNode("n:CancelInfo/n:CancelReason").text
			select case gubunname
				case "INTENT_CHANGED"
					gubunname = "�ܼ�����"
				case "PRODUCT_UNSATISFIED"
					gubunname = "��ǰ�Ҹ���"
				case "SOLD_OUT"
					gubunname = "ǰ��"
				case "COLOR_AND_SIZE"
					gubunname = "���������"
				case "WRONG_ORDER"
					gubunname = "�ֹ�����"
				case else
					gubunname = Replace(gubunname, "'", "")
			end select
		'elseif (csGubun = "RETURNED") then
		elseif (csGubun = "RETURN_REQUESTED") then	'2019-07-23 ������ RETURN_REQUESTED�� ����
			divcd = "A004"
			OutMallRegDate = Left(obj.selectSingleNode("n:ReturnInfo/n:ClaimRequestDate").text, 10)

			gubunname = obj.selectSingleNode("n:ReturnInfo/n:ReturnReason").text
			select case gubunname
				case "INTENT_CHANGED"
					gubunname = "�ܼ�����"
				case "PRODUCT_UNSATISFIED"
					gubunname = "��ǰ�Ҹ���"
				case "SOLD_OUT"
					gubunname = "ǰ��"
				case "COLOR_AND_SIZE"
					gubunname = "���������"
				case "WRONG_ORDER"
					gubunname = "�ֹ�����"
				case else
					gubunname = Replace(gubunname, "'", "")
			end select
		elseif (csGubun = "EXCHANGE_REQUESTED") then	'2022-03-03 ������ EXCHANGE_REQUESTED�߰�
			divcd = "A000"
			OutMallRegDate = Left(obj.selectSingleNode("n:ExchangeInfo/n:ClaimRequestDate").text, 10)
			gubunname = obj.selectSingleNode("n:ExchangeInfo/n:ExchangeReason").text
			select case gubunname
				case "INTENT_CHANGED"
					gubunname = "�ܼ�����"
				case "PRODUCT_UNSATISFIED"
					gubunname = "��ǰ�Ҹ���"
				case "SOLD_OUT"
					gubunname = "ǰ��"
				case "COLOR_AND_SIZE"
					gubunname = "���������"
				case "WRONG_ORDER"
					gubunname = "�ֹ�����"
				case else
					gubunname = Replace(gubunname, "'", "")
			end select
		elseif (csGubun = "EXCHANGED") then
			divcd = 1/0			'// ����
			divcd = "A000"
			OutMallRegDate = "1900-01-01"
		else
			divcd = 1/0			'// ����
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

			'' CS ����������. ������Ʈ
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
		rw "��� CS�Է°Ǽ�:"&iInputCnt
	elseif (csGubun = "RETURNED") then
		rw "��ǰ CS�Է°Ǽ�:"&iInputCnt
	elseif (csGubun = "EXCHANGE_REQUESTED") then
		rw "��ȯ CS�Է°Ǽ�:"&iInputCnt
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
	'// API URL(�Ⱓ������ �ֹ� ��������)
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
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>"															'#�����޴� �������� �� ����(Compact / Full)
	strRst = strRst & "			<sel:Version>4.1</sel:Version>"
	strRst = strRst & "			<sel:InquiryTimeFrom>"&selldate&"T00:00:00</sel:InquiryTimeFrom>"									'#��ȸ ���� �Ͻ�(�ش� �ð� ����)
	strRst = strRst & "			<sel:InquiryTimeTo>"& Left(DateAdd("d", 1, CDate(selldate)), 10)&"T00:00:00</sel:InquiryTimeTo>"	'��ȸ ���� �Ͻ�(�ش� �ð� �������� ����)
	strRst = strRst & "			<sel:LastChangedStatusCode>" & LastChangedStatusCode & "</sel:LastChangedStatusCode>"				'���� ��ǰ �ֹ� ���� �ڵ� (CANCELED | ���, RETURNED | ��ǰ, EXCHANGED : ��ȯ | PAYED : �����Ϸ�)
	strRst = strRst & "			<sel:MallID>"&reqID&"</sel:MallID>"																	'�Ǹ��� ���̵�
	strRst = strRst & "		</sel:GetChangedProductOrderListRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	''response.write strRst
	''dbget.close : response.end

	'// =======================================================================
	'// ����Ÿ ��������
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		response.write "ERROR : ��ſ���"
		dbget.close : response.end
	end if


	'// =======================================================================
	'// XML DOM ����
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
	''response.write objXML.responseText & "<br /><br />"

	''dbget.close : response.end

	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		response.write "���� : ����"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	if CLng(xmlDOM.getElementsByTagName("n:ReturnedDataCount").item(0).text) = 0 then
		response.write "��������<br />"
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

''1300k �ֹ���� ��ȸ
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
	'// ��¥����
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
				obj("claim_date")("st_date") = stDate&"0000"		'#���۽ð� YYYYMMDDHHMM
				obj("claim_date")("ed_date") = edDate&"0000"		'#����ð� YYYYMMDDHHMM
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
								CSDetailKey = Trim(claimList.get(i).claim_no)				'Ŭ���ӹ�ȣ
								CSDetailKeySub = Trim(claimList.get(i).claim_sub_no)		'Ŭ����SUB��ȣ
								CSDetailKey = CSDetailKey & "-" & CSDetailKeySub

								If InStr(Trim(claimList.get(i).claim_type), "���") >0 Then
									divcd = "A008"
								ElseIf InStr(Trim(claimList.get(i).claim_type), "��ǰ") >0 Then
									divcd = "A004"
								ElseIf InStr(Trim(claimList.get(i).claim_type), "��ȯ") >0 Then
									divcd = "A000"
								ElseIf InStr(Trim(claimList.get(i).claim_type), "ȸ��") >0 Then
									divcd = "A200"
								End If

'								Trim(claimList.get(i).claim_result)		'STATUS
								OutMallRegDate = Trim(claimList.get(i).claim_request_date)	'Ŭ���ӿ�û�ð�
'								Trim(claimList.get(i).status_change_date)	'���º���ð�
								gubunname = Trim(claimList.get(i).reason)				'Ŭ���ӻ���
'								Trim(claimList.get(i).remark)				'���ڸ�Ʈ
								OutMallOrderSerial = Trim(claimList.get(i).order_no)			'�ֹ���ȣ
								' Trim(claimList.get(i).order_name)			'�ֹ��ڸ�
								OrgDetailKey = Trim(claimList.get(i).seq_no)				'�Ϸù�ȣ
								' Trim(claimList.get(i).product_code)		'��ǰ�ڵ�
								' Trim(claimList.get(i).product_name)		'��ǰ��
								' Trim(claimList.get(i).opt_no)				'�ɼ��ڵ�
								' Trim(claimList.get(i).opt_name)			'�ɼǸ�
								itemno = Trim(claimList.get(i).qty)				'����
								' Trim(claimList.get(i).company_product_code)	'��ü��ǰ�ڵ�
								' Trim(claimList.get(i).company_opt_no)		'��ü�ɼǹ�ȣ

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
										''�ֹ� �Է� ���� ������ ���� ����
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET matchState = 'D'"
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder c "
										strSql = strSql & " WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										strSql = strSql & " and orderserial is NULL"
										dbget.Execute strSql
									End If

									'' CS ����������. ������Ʈ
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

''1300k �ֹ���� ��ȸ
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
	'// ��¥����
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
				obj("cancel_date")("st_date") = stDate&"0000"		'#���۽ð� YYYYMMDDHHMM
				obj("cancel_date")("ed_date") = edDate&"0000"		'#����ð� YYYYMMDDHHMM
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
								OutMallRegDate = Trim(claimList.get(i).cancel_date)	'Ŭ���ӿ�û�ð�
								OutMallRegDate = Left(OutMallRegDate,4)&"-"&Mid(OutMallRegDate,5,2)&"-"&Mid(OutMallRegDate,7,2)&" "&Mid(OutMallRegDate,9,2)&":"&Mid(OutMallRegDate,11,2)&":"&Mid(OutMallRegDate,13,2)
								OutMallOrderSerial = Trim(claimList.get(i).order_no)			'�ֹ���ȣ
								OrgDetailKey = Trim(claimList.get(i).seq_no)				'�Ϸù�ȣ
								itemno = Trim(claimList.get(i).qty)				'����

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
										''�ֹ� �Է� ���� ������ ���� ����
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET matchState = 'D'"
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder c "
										strSql = strSql & " WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										strSql = strSql & " and orderserial is NULL"
										dbget.Execute strSql
									End If

									'' CS ����������. ������Ʈ
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

''1300k ��ǰ ��ȸ
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
	'// ��¥����
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
				obj("return_date")("st_date") = stDate&"0000"		'#���۽ð� YYYYMMDDHHMM
				obj("return_date")("ed_date") = edDate&"0000"		'#����ð� YYYYMMDDHHMM
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
								OutMallRegDate = Trim(claimList.get(i).return_date)	'Ŭ���ӿ�û�ð�
								OutMallRegDate = Left(OutMallRegDate,4)&"-"&Mid(OutMallRegDate,5,2)&"-"&Mid(OutMallRegDate,7,2)&" "&Mid(OutMallRegDate,9,2)&":"&Mid(OutMallRegDate,11,2)&":"&Mid(OutMallRegDate,13,2)
								OutMallOrderSerial = Trim(claimList.get(i).order_no)			'�ֹ���ȣ
								OrgDetailKey = Trim(claimList.get(i).seq_no)				'�Ϸù�ȣ
								itemno = Trim(claimList.get(i).qty)				'����

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
									'' CS ����������. ������Ʈ
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

''skstoa �ֹ���� ��ȸ
Function GetCSOrderCancel_skstoa(sellsite, csGubun, csSubGubun, selldate)
	Dim xmlURL, strRst, xmlSelldate
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate, OrgOutMallOrderSerial, orderCancelList
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete, claimType
	Dim addParam, returnStatus, iMessage
	'// ��¥����
	xmlSelldate = Replace(selldate, "-", "")

	addParam = ""
	addParam = addParam & "linkCode=" & skstoalinkCode		'#�����ڵ� | SKB���� �ο��� �����ڵ�
	addParam = addParam & "&entpCode=" & skstoaentpCode		'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
	addParam = addParam & "&entpId=" & skstoaentpId			'#��ü�����ID | SKB���� �ο��� ��ü����� ID
	addParam = addParam & "&entpPass=" & skstoaentpPass		'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
	addParam = addParam & "&bDate="& xmlSelldate			'#��ȸ �������� | ���� ó���� ���� YYYYMMDD Ÿ��. ex) 20140520
	addParam = addParam & "&eDate="& xmlSelldate			'#��ȸ ���������� | ���� ó���� ���� YYYYMMDD Ÿ��. ex) 20140520
'	addParam = addParam & "&orderNo="						'�ֹ��ڵ� | �ֹ���ȣ�� �̿��� �˻�. ���ڸ� ���Ǹ� ���� 14�ڸ�(orderNo) �Ǵ� 23�ڸ�(orderNo,orderGSeq,orderDSeq,orderWSeq)�� ���� ���
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
								OutMallOrderSerial	= orderCancelList.get(i).orderNo			'�ֹ���ȣ
								OrgDetailKey		= orderCancelList.get(i).orderGSeq & "-" & orderCancelList.get(i).orderDSeq & "-" & orderCancelList.get(i).orderWSeq	'��ǰ���� - ��Ʈ���� - ó������
								OutMallRegDate		= orderCancelList.get(i).orderDate			'�ֹ�������
								OutMallRegDate = Left(OutMallRegDate, 4) & "-" & Mid(OutMallRegDate, 5, 2) & "-" & Right(OutMallRegDate, 2)
								' rw orderCancelList.get(i).goodsCode			'�ǸŻ�ǰ�ڵ�
								' rw orderCancelList.get(i).goodsName			'�ǸŻ�ǰ��
								' rw orderCancelList.get(i).goodsdtCode		'�ǸŴ�ǰ�ڵ�
								' rw orderCancelList.get(i).goodsdtInfo		'�ǸŴ�ǰ����
								' rw orderCancelList.get(i).goodsGb			'��ǰ����
								itemno				= orderCancelList.get(i).orderQty			'���ֹ�����
								' rw orderCancelList.get(i).salePrice			'���� �ǸŰ�
								' rw orderCancelList.get(i).buyPrice			'���� ���԰�
								' rw orderCancelList.get(i).custName			'�ֹ��ڸ�
'								gubunname			= orderCancelList.get(i).msg				'��۸޽���
								gubunname = "���"
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
									''�ֹ� �Է� ���� ������ ���� ����
									strSql = ""
									strSql = strSql & " UPDATE c "
									strSql = strSql & " SET matchState = 'D'"
									strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder c "
									strSql = strSql & " WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
									strSql = strSql & " and orderserial is NULL"
									dbget.Execute strSql

									'' CS ����������. ������Ʈ
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

''skstoa ��ǰ��ȯ ȸ�� �����ȸ
Function GetCSOrderReturnExchange_skstoa(sellsite, csGubun, csSubGubun, selldate)
	Dim xmlURL, strRst, xmlSelldate
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate, OrgOutMallOrderSerial, returnList
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete, claimType
	Dim addParam, returnStatus, iMessage
	'// ��¥����
	xmlSelldate = Replace(selldate, "-", "")
	addParam = ""
	addParam = addParam & "linkCode=" & skstoalinkCode		'#�����ڵ� | SKB���� �ο��� �����ڵ�
	addParam = addParam & "&entpCode=" & skstoaentpCode		'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
	addParam = addParam & "&entpId=" & skstoaentpId			'#��ü�����ID | SKB���� �ο��� ��ü����� ID
	addParam = addParam & "&entpPass=" & skstoaentpPass		'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
	addParam = addParam & "&claimGb=00" 					'ȸ�� ���� �� | 00:��ü(default),30:��ǰȸ��,45:��ȯȸ��
	addParam = addParam & "&bDate="& xmlSelldate			'#��ȸ �������� | �ֹ�����/��ȯ������ ����, YYYYMMDD Ÿ��. ex) 20140520
	addParam = addParam & "&eDate="& xmlSelldate			'#��ȸ ���������� | �ֹ�����/��ȯ������ ���� YYYYMMDD Ÿ��. ex) 20140520
'	addParam = addParam & "&orderNo="						'�ֹ��ڵ� | �ֹ���ȣ�� �̿��� �˻�. ���ڸ� ���Ǹ� ���� 14�ڸ�(orderNo) �Ǵ� 23�ڸ�(orderNo,orderGSeq,orderDSeq,orderWSeq)�� ���� ���
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
								If returnList.get(i).claimGb = "30" Then				'ȸ�� ���� ��
									divcd = "A004"
								Else
									divcd = "A000"
								End If
								'rw returnList.get(i).claimGbName						'ȸ�� ���� ��
								OutMallOrderSerial	= returnList.get(i).orderNo			'�ֹ���ȣ
								OrgDetailKey		= returnList.get(i).orderGSeq & "-" & returnList.get(i).orderDSeq & "-" & returnList.get(i).orderWSeq	'��ǰ���� - ��Ʈ���� - ó������
								'rw returnList.get(i).goodsGb							'��ǰ����
								'rw returnList.get(i).goodsCode							'�ǸŻ�ǰ�ڵ�
								'rw returnList.get(i).goodsName							'�ǸŻ�ǰ��
								'rw returnList.get(i).goodsdtCode						'�ǸŴ�ǰ�ڵ�
								'rw returnList.get(i).goodsdtInfo						'�ǸŴ�ǰ����
								OutMallRegDate		= returnList.get(i).returnProcDate	'ȸ��������
								itemno				= returnList.get(i).syslast			'��ۼ���
								'rw returnList.get(i).salePrice							'�ǸŰ�
								'rw returnList.get(i).buyPrice							'���԰�(���꿹���ݾ�)
								'rw returnList.get(i).shipCostName						'��ۺ���å��
								'rw returnList.get(i).custName							'�ֹ��ڸ�
								'rw returnList.get(i).receiverName						'�����θ�
								'rw returnList.get(i).receiverPostNo					'�����ȣ
								'rw returnList.get(i).receiverTel						'����ó
								'rw returnList.get(i).receiverHp						'�޴���

								On Error Resume Next
									gubunname			= returnList.get(i).msg			'��ǰ�󼼻���
									If Err.number <> 0 Then
										gubunname = "��������"
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

									'' CS ����������. ������Ʈ
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
									strSql = strSql & " 	and LEFT(c.OrgDetailKey, 7) = LEFT(o.OrgDetailKey, 7) "		'�ż��踸 LEFT 7�� �����ؾߵ�..������ 3�ڸ��� �� ����
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

''shintvshopping �ֹ���� ��ȸ
Function GetCSOrderCancel_shintvshopping(sellsite, csGubun, csSubGubun, selldate)
	Dim xmlURL, strRst, xmlSelldate
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate, OrgOutMallOrderSerial, orderCancelList
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete, claimType
	Dim addParam, returnStatus, iMessage
	'// ��¥����
	xmlSelldate = Replace(selldate, "-", "")

	addParam = ""
	addParam = addParam & "linkCode=" & linkCode			'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
	addParam = addParam & "&entpCode=" & entpCode			'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
	addParam = addParam & "&entpId=" & entpId				'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
	addParam = addParam & "&entpPass=" & entpPass			'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
	addParam = addParam & "&bDate="& xmlSelldate			'#��ȸ �������� | �ֹ�����/��ȯ������ ����, YYYYMMDD Ÿ��. ex) 20140520
	addParam = addParam & "&eDate="& xmlSelldate			'#��ȸ ���������� | �ֹ�����/��ȯ������ ���� YYYYMMDD Ÿ��. ex) 20140520
'	addParam = addParam & "&orderNo="						'�ֹ��ڵ� | �ֹ���ȣ�� �̿��� �˻�. ���ڸ� ���Ǹ� ���� 14�ڸ�(orderNo) �Ǵ� 23�ڸ�(orderNo,orderGSeq,orderDSeq,orderWSeq)�� ���� ���
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
								OutMallOrderSerial	= orderCancelList.get(i).orderNo			'�ֹ���ȣ
								OrgDetailKey		= orderCancelList.get(i).orderGSeq & "-" & orderCancelList.get(i).orderDSeq & "-" & orderCancelList.get(i).orderWSeq	'��ǰ���� - ��Ʈ���� - ó������
								OutMallRegDate		= orderCancelList.get(i).orderDate			'�ֹ�������
								OutMallRegDate = Left(OutMallRegDate, 4) & "-" & Mid(OutMallRegDate, 5, 2) & "-" & Right(OutMallRegDate, 2)
								' rw orderCancelList.get(i).goodsCode			'�ǸŻ�ǰ�ڵ�
								' rw orderCancelList.get(i).goodsName			'�ǸŻ�ǰ��
								' rw orderCancelList.get(i).goodsdtCode		'�ǸŴ�ǰ�ڵ�
								' rw orderCancelList.get(i).goodsdtInfo		'�ǸŴ�ǰ����
								' rw orderCancelList.get(i).goodsGb			'��ǰ����
								itemno				= orderCancelList.get(i).orderQty			'���ֹ�����
								' rw orderCancelList.get(i).salePrice			'���� �ǸŰ�
								' rw orderCancelList.get(i).buyPrice			'���� ���԰�
								' rw orderCancelList.get(i).custName			'�ֹ��ڸ�
								gubunname			= orderCancelList.get(i).msg				'��۸޽���

								If ISNULL(gubunname) OR LEN(gubunname) < 2 Then
									gubunname = "���"
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
									''�ֹ� �Է� ���� ������ ���� ����
									strSql = ""
									strSql = strSql & " UPDATE c "
									strSql = strSql & " SET matchState = 'D'"
									strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder c "
									strSql = strSql & " WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
									strSql = strSql & " and orderserial is NULL"
									dbget.Execute strSql

									'' CS ����������. ������Ʈ
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

''shintvshopping ��ǰ��ȯ ȸ�� �����ȸ
Function GetCSOrderReturnExchange_shintvshopping(sellsite, csGubun, csSubGubun, selldate)
	Dim xmlURL, strRst, xmlSelldate
	Dim objXML, xmlDOM, objArr, obj
	Dim i, objJson, strObj, returnCode, resultmsg, j
	Dim startdate, enddate, OrgOutMallOrderSerial, returnList
	Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	Dim iAssignedRow, iInputCnt, iRbody, datalist, itemList
	Dim strSql, IsDelete, claimType
	Dim addParam, returnStatus, iMessage
	'// ��¥����
	xmlSelldate = Replace(selldate, "-", "")
	addParam = ""
	addParam = addParam & "linkCode=" & linkCode			'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
	addParam = addParam & "&entpCode=" & entpCode			'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
	addParam = addParam & "&entpId=" & entpId				'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
	addParam = addParam & "&entpPass=" & entpPass			'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
	addParam = addParam & "&claimGb=00" 					'ȸ�� ���� �� | 00:��ü(default),30:��ǰȸ��,45:��ȯȸ��
	addParam = addParam & "&bDate="& xmlSelldate			'#��ȸ �������� | �ֹ�����/��ȯ������ ����, YYYYMMDD Ÿ��. ex) 20140520
	addParam = addParam & "&eDate="& xmlSelldate			'#��ȸ ���������� | �ֹ�����/��ȯ������ ���� YYYYMMDD Ÿ��. ex) 20140520
'	addParam = addParam & "&orderNo="						'�ֹ��ڵ� | �ֹ���ȣ�� �̿��� �˻�. ���ڸ� ���Ǹ� ���� 14�ڸ�(orderNo) �Ǵ� 23�ڸ�(orderNo,orderGSeq,orderDSeq,orderWSeq)�� ���� ���
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
								If returnList.get(i).claimGb = "30" Then				'ȸ�� ���� ��
									divcd = "A004"
								Else
									divcd = "A000"
								End If
								'rw returnList.get(i).claimGbName						'ȸ�� ���� ��
								OutMallOrderSerial	= returnList.get(i).orderNo			'�ֹ���ȣ
								OrgDetailKey		= returnList.get(i).orderGSeq & "-" & returnList.get(i).orderDSeq & "-" & returnList.get(i).orderWSeq	'��ǰ���� - ��Ʈ���� - ó������
								'rw returnList.get(i).goodsGb							'��ǰ����
								'rw returnList.get(i).goodsCode							'�ǸŻ�ǰ�ڵ�
								'rw returnList.get(i).goodsName							'�ǸŻ�ǰ��
								'rw returnList.get(i).goodsdtCode						'�ǸŴ�ǰ�ڵ�
								'rw returnList.get(i).goodsdtInfo						'�ǸŴ�ǰ����
								OutMallRegDate		= returnList.get(i).returnProcDate	'ȸ��������
								itemno				= returnList.get(i).syslast			'��ۼ���
								'rw returnList.get(i).salePrice							'�ǸŰ�
								'rw returnList.get(i).buyPrice							'���԰�(���꿹���ݾ�)
								'rw returnList.get(i).shipCostName						'��ۺ���å��
								'rw returnList.get(i).custName							'�ֹ��ڸ�
								'rw returnList.get(i).receiverName						'�����θ�
								'rw returnList.get(i).receiverPostNo					'�����ȣ
								'rw returnList.get(i).receiverTel						'����ó
								'rw returnList.get(i).receiverHp						'�޴���

								On Error Resume Next
									gubunname			= returnList.get(i).msg			'��ǰ�󼼻���
									If Err.number <> 0 Then
										gubunname = "��������"
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

									'' CS ����������. ������Ʈ
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
									strSql = strSql & " 	and LEFT(c.OrgDetailKey, 7) = LEFT(o.OrgDetailKey, 7) "		'�ż��踸 LEFT 7�� �����ؾߵ�..������ 3�ڸ��� �� ����
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

'�Ե�On ��ҿ�û(�Ϸ�) ��ȸ
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
		objJson("srchStrtDttm") = startdate		'#�˻��������� yyyyMMddhh24miss
		objJson("srchEndDttm") = enddate		'#�˻��������� yyyyMMddhh24miss
		objJson("odNo") = ""					'o�ֹ���ȣ
		objJson("lrtrNo") = ""					'o�����ŷ�ó��ȣ�� �����ϸ� ������ �ʼ��Է�
		strRst = objJson.jsString
	SET objJson = nothing

	'// ����Ÿ ��������
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
							OutMallOrderSerial		= datalist.get(i).odNo				'�ֹ���ȣ
							CSDetailKey				= datalist.get(i).clmNo				'Ŭ���ӹ�ȣ
							Set itemList = datalist.get(i).itemList
								For j = 0 to itemList.length-1
									OrgDetailKey	= itemList.get(j).odSeq				'�ֹ�����
'									procSeq			= itemList.get(j).procSeq			'ó������
'									orglProcSeq		= itemList.get(j).orglProcSeq		'��ó������
'									odTypCd			= itemList.get(j).odTypCd			'�ֹ������ڵ�( 20 �ֹ���� 50 AS)
'									odTypDtlCd		= itemList.get(j).odTypDtlCd		'�ֹ��������ڵ�( �ֹ������� AS�� ��� �ʼ�) | 5011 : �������, 5021 : �߰�������, 5031 : �κ�ǰȸ�����
									claimType		= itemList.get(j).odPrgsStepCd		'�ֹ�����ܰ��ڵ�(02 ��û 21 ��ҿϷ� 22 öȸ(��ҿ�û) )
'									spdNo			= itemList.get(j).spdNo				'�Ǹ��ڻ�ǰ��ȣ
'									spdNm			= itemList.get(j).spdNm				'�Ǹ��ڻ�ǰ��
'									sitmNo			= itemList.get(j).sitmNo			'�Ǹ��� ��ǰ��ȣ
'									sitmNm			= itemList.get(j).sitmNm			'�Ǹ��� ��ǰ��
									itemno			= itemList.get(j).odQty				'�ֹ�����
'									itmSlPrc		= itemList.get(j).itmSlPrc			'�ǸŰ� (������ ������ ��ǰ �Ǹ� ����)
'									cnclQty			= itemList.get(j).cnclQty			'��Ҽ���
'									trNo			= itemList.get(j).trNo				'�ŷ�ó��ȣ
'									lrtrNo			= itemList.get(j).lrtrNo			'�����ŷ�ó��ȣ
'									odAccpDttm		= itemList.get(j).odAccpDttm		'�ֹ������Ͻ� [yyyyMMddHHmmss : 20191201121212]
'									purCfrmDttm		= itemList.get(j).purCfrmDttm		'����Ȯ���Ͻ� [yyyyMMddHHmmss : 20191201121212] �߰���ǰ�ΰ�� ����Ȯ��������϶� ��ϵ�
									OutMallRegDate	= itemList.get(j).clmReqDttm		'Ŭ���ӿ�û�Ͻ�[yyyyMMddHHmmss : 20191201121212] - ��ҿ�û�Ͻ�
'									clmCmptDttm		= itemList.get(j).clmCmptDttm		'Ŭ���ӿϷ��Ͻ�[yyyyMMddHHmmss : 20191201121212] - ��ҿϷ��Ͻ�
									gubunname		= itemList.get(j).clmRsnCd			'Ŭ���ӻ����ڵ� | 101 : ����� �ʾ���, 102 : ��ǰ�� ǰ����, 103 : �ɼ�/������ �Ҹ� /���, 104 : �ٸ� ������ ���, 105 : ����/�������� ����, 106 : �����ǻ� ������, 107 : ����ǰ ���� / ���, 108 : �������� ����(���Ż��� ����/ī�庯�� ��), 109 : �����ǰ ����, 110 : ��ǰ���� ����, 111 : �Ǹ��� ���(�Ǹ���), 112 : ���� ���(����), 113 : �ڵ� ���(��������), 114 : �ڵ� ���(��ǰ���), 115 : �ڵ� ���(�����̼���), 116 : �ڵ� ���(����Ʈ�ȹ̼���), 117 : �ڵ� ���(��ҿ�û ��å ����)

									Select Case gubunname
										Case "103", "104", "106", "109"
											gubunname = "�ܼ�����"
										Case "108"
											gubunname = "���ú���"
										Case "101"
											gubunname = "�������"
										Case "110"
											gubunname = "��ǰ����Ʋ��"
										Case "102"
											gubunname = "ǰ��"
										Case "105"
											gubunname = "����������"
										Case Else
											gubunname = "��Ÿ"
									End Select

									If claimType = "22" Then
										IsDelete = "Y"
									End If

'									clmRsnCnts		= itemList.get(j).clmRsnCnts		'Ŭ���ӻ�������
'									odFvrGrpNo		= itemList.get(j).odFvrGrpNo		'�ֹ����ñ׷��ȣ(N+1, ��������� �� ������ ���� �ִ� �׷��ȣ)
'									fvrAmt			= itemList.get(j).fvrAmt			'���� ������ ���αݾ��� ��ұݾ�, ������ 0 (��ҿϷ������ ����) 1~5�����αݾ� ����հ�
'									sellerCnYn		= itemList.get(j).sellerCnYn		'�Ǹ���������ҿ���
'									frstDvCst		= itemList.get(j).frstDvCst			'�ʵ���ۺ�
'									addDvCst		= itemList.get(j).addDvCst			'�߰���ۺ�
'									dvCstBdnMnbdCd	= itemList.get(j).dvCstBdnMnbdCd	'��ۺ�δ���ü�ڵ�(01:�� , 02: ��ü)
'									excpProcDvsCd	= itemList.get(j).excpProcDvsCd		'����ó�������ڵ�(03 ���� , 04 �ź�)
'									rpcSpdNo		= itemList.get(j).rpcSpdNo			'��ü�Ǹ��ڻ�ǰ��ȣ (��Ʈ/���� �����Ұ�� ����)
'									rpcSpdNm		= itemList.get(j).rpcSpdNm			'��ü�Ǹ��ڻ�ǰ�� (��Ʈ/���� �����Ұ�� ����)
'									rpcSitmNo		= itemList.get(j).rpcSitmNo			'��ü�Ǹ��ڴ�ǰ��ȣ(��Ʈ/���� �����Ұ�� ����)
'									rpcSitmNm		= itemList.get(j).rpcSitmNm			'��ü�Ǹ��ڴ�ǰ��(��Ʈ/���� �����Ұ�� ����)
'									cmbnDvGrpNo		= itemList.get(j).cmbnDvGrpNo		'�չ�۱׷��ȣ

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
										''�ֹ� �Է� ���� ������ ���� ����
										strSql = ""
										strSql = strSql & " UPDATE c "
										strSql = strSql & " SET matchState = 'D'"
										strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder c "
										strSql = strSql & " WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										strSql = strSql & " and orderserial is NULL"
										dbget.Execute strSql

										'' CS ����������. ������Ʈ
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
						rw "���öȸ CS�Է°Ǽ�:"&iInputCnt
					Else
						rw "�ֹ���� CS�Է°Ǽ�:"&iInputCnt
					End If
				Else
					rw resultmsg
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'�Ե�On ��ǰ��û/���� �����ȸ
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
		objJson("srchStrtDttm") = startdate		'#�˻��������� yyyyMMddhh24miss
		objJson("srchEndDttm") = enddate		'#�˻��������� yyyyMMddhh24miss
		objJson("odNo") = ""					'o�ֹ���ȣ( only �ֹ���ȣ�϶��� �ʼ�)
		objJson("lrtrNo") = ""					'o�����ŷ�ó��ȣ
		strRst = objJson.jsString
	SET objJson = nothing

	'// ����Ÿ ��������
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
							OutMallOrderSerial	= datalist.get(i).odNo				'�ֹ���ȣ
							CSDetailKey			= datalist.get(i).clmNo				'Ŭ���ӹ�ȣ
							Set itemList = datalist.get(i).itemList
								For j = 0 to itemList.length-1
									OrgDetailKey	= itemList.get(j).odSeq				'�ֹ�����
'									procSeq			= itemList.get(j).procSeq			'ó������
'									orglProcSeq		= itemList.get(j).orglProcSeq		'��ó������
'									odTypCd			= itemList.get(j).odTypCd			'�ֹ����� (40 ��ǰ 50 AS)
'									odTypDtlCd		= itemList.get(j).odTypDtlCd		'�ֹ��������ڵ� ( �ֹ������� AS�� ��� �ʼ�) | 5030 �κ�ǰȸ��
									claimType		= itemList.get(j).odPrgsStepCd		'�ֹ�����ܰ�(02��û , 03 ����, 27 ��ǰ�Ϸ�)
'									spdNo			= itemList.get(j).spdNo				'�Ǹ��ڻ�ǰ��ȣ
'									spdNm			= itemList.get(j).spdNm				'�Ǹ��ڻ�ǰ��
'									sitmNo			= itemList.get(j).sitmNo			'�Ǹ��� ��ǰ��ȣ
'									sitmNm			= itemList.get(j).sitmNm			'�Ǹ��� ��ǰ��
									itemno			= itemList.get(j).odQty				'�ֹ�����
'									itmSlPrc		= itemList.get(j).itmSlPrc			'��ǰ�ǸŰ�
'									rtngQty			= itemList.get(j).rtngQty			'��ǰ����
'									trNo			= itemList.get(j).trNo				'�ŷ�ó��ȣ
'									lrtrNo			= itemList.get(j).lrtrNo			'�����ŷ�ó��ȣ
'									OutMallRegDate	= itemList.get(j).odAccpDttm		'�ֹ������Ͻ�[yyyyMMddHHmmss : 20191201121212]
'									purCfrmDttm		= itemList.get(j).purCfrmDttm		'����Ȯ���Ͻ�[yyyyMMddHHmmss : 20191201121212]
'									OutMallRegDate		= itemList.get(j).clmReqDttm		'Ŭ���ӿ�û�Ͻ�[yyyyMMddHHmmss : 20191201121212] -  ��ǰ��û�Ͻ�
'									clmAccpDttm		= itemList.get(j).clmAccpDttm		'Ŭ���������Ͻ�[yyyyMMddHHmmss : 20191201121212] -  ��ǰ�����Ͻ�
'									clmCmptDttm		= itemList.get(j).clmCmptDttm		'Ŭ���ӿϷ��Ͻ�[yyyyMMddHHmmss : 20191201121212] - ��ǰ�Ϸ��Ͻ�
									gubunname		= itemList.get(j).clmRsnCd			'Ŭ���ӻ����ڵ� | 301 : ��ǰ�� ���ڰ� ����(�ļ�/�ҷ�), 302 : �����ǻ� ������, 303 : �����ǻ� ������(�Ⱓ ����), 304 : �ٸ� ��ǰ�� ��۵�, 305 : ��ǰ�� ������ �ٸ�(��ǰ��������), 306 : �ɼ�/������ �Ҹ�, 307 : �������ܺ���(���Ż��� ����/ī�� ���� ��), 308 : ��ǰ/����ǰ�� �ȿ�, 309 : �ٸ� ������ ���, 310 : �ڵ� ��ǰ(ũ�ν��ȹ̼���)

									Select Case gubunname
										Case "301"
											gubunname = "��ǰ�ҷ�"
										Case "302", "303", "306", "309"
											gubunname = "�ܼ�����"
										Case "304"
											gubunname = "�����"
										Case "305"
											gubunname = "��ǰ����Ʋ��"
										Case "307"
											gubunname = "���ú���"
										Case "308"
											gubunname = "�������"
										Case Else
											gubunname = "��Ÿ"
									End Select

'									clmRsnCnts		= itemList.get(j).clmRsnCnts		'Ŭ���ӻ�������
'									spicYn			= itemList.get(j).spicYn			'����Ʈ�ȿ��� (Y/N)
'									rtrvSeq			= itemList.get(j).rtrvSeq			'ȸ��������(���������)
'									rtrvCustNm		= itemList.get(j).rtrvCustNm		'ȸ��������
'									rtrvTelNo		= itemList.get(j).rtrvTelNo			'ȸ������ȭ��ȣ
'									rtrvMphnNo		= itemList.get(j).rtrvMphnNo		'ȸ�����޴�����ȣ
'									rtrvZipNo		= itemList.get(j).rtrvZipNo			'ȸ���������ȣ
'									rtrvZipNoSeq	= itemList.get(j).rtrvZipNoSeq		'ȸ���������ȣ����
'									rtrvStnmZipAddr	= itemList.get(j).rtrvStnmZipAddr	'ȸ�������θ�����ּ�(��Ʈ)
'									rtrvStnmDtlAddr	= itemList.get(j).rtrvStnmDtlAddr	'ȸ�������θ���ּ�
'									dvMsg			= itemList.get(j).dvMsg				'��۸޽���
'									spicTypCd		= itemList.get(j).spicTypCd			'����Ʈ�������ڵ� (CRSS ũ�ν��� RVS �������� STR �������)
'									rnkhSpplcNo		= itemList.get(j).rnkhSpplcNo		'�����Ⱦ�ó��ȣ
'									rnklSpplcNo		= itemList.get(j).rnklSpplcNo		'�����Ⱦ�ó��ȣ
'									pkupPlcNo		= itemList.get(j).pkupPlcNo			'�Ⱦ���ҹ�ȣ
'									pkupPlcNm		= itemList.get(j).pkupPlcNm			'�Ⱦ���Ҹ�
'									spicBxchNo		= itemList.get(j).spicBxchNo		'����Ʈ�ȱ�ȯ�ǹ�ȣ
'									pkupBgtDttm		= itemList.get(j).pkupBgtDttm		'�Ⱦ������Ͻ�[yyyyMMddHHmmss : 20191201121212]
'									excpProcDvsCd	= itemList.get(j).excpProcDvsCd		'����ó�������ڵ�(03 ���� , 04 �ź�)
'									rpcSpdNo		= itemList.get(j).rpcSpdNo			'��ü�Ǹ��ڻ�ǰ��ȣ (��Ʈ/���� �����Ұ�� ����)
'									rpcSpdNm		= itemList.get(j).rpcSpdNm			'��ü�Ǹ��ڻ�ǰ�� (��Ʈ/���� �����Ұ�� ����)
'									rpcSitmNo		= itemList.get(j).rpcSitmNo			'��ü�Ǹ��ڴ�ǰ��ȣ (��Ʈ/���� �����Ұ�� ����)
'									rpcSitmNm		= itemList.get(j).rpcSitmNm			'��ü�Ǹ��ڴ�ǰ�� (��Ʈ/���� �����Ұ�� ����)
'									frstDvCst		= itemList.get(j).frstDvCst			'�ʵ���ۺ�
'									addDvCst		= itemList.get(j).addDvCst			'��ǰ�߰���ۺ�
'									rcst			= itemList.get(j).rcst				'��ǰ��
'									dvCstBdnMnbdCd	= itemList.get(j).dvCstBdnMnbdCd	'��ۺ�δ���ü�ڵ�(01:�� , 02: ��ü)
'									shopCnvMsg		= itemList.get(j).shopCnvMsg		'�������޸޼���
'									dvLrtrNo		= itemList.get(j).dvLrtrNo			'��������ŷ�ó��ȣ-���ڵ� �������� (����/��Ʈ)
'									hpDvDttm		= itemList.get(j).hpDvDttm			'�������Ͻ� yyyymmddhh24miss (����/��Ʈ) | - ������� ����� �� �����ۻ�ǰ�� ����Ͻ� �� ��ġ��ǰ ��������
'									afflSqncTypCd	= itemList.get(j).afflSqncTypCd		'�迭��ȸ�������ڵ�(����/��Ʈ) | CRSS_PIC:ũ�ν���, DN_DV:�������, DRIVE_PIC:����̺���, DRT_DV:�ٷι��, FPRD_DV:������, FS_PKUP:�����Ⱦ�, RENTAL_CAR_PKUP:����ī�Ⱦ�, SHOP_DV:������, SMT_QCK:����Ʈ��, STR_PIC:�������
'									afflCbCd		= itemList.get(j).afflCbCd			'�迭��ť���ڵ�(����/��Ʈ)
'									dvSqncCd		= itemList.get(j).dvSqncCd			'ȸ���ڵ� ����On���� ȸ����ȣ(����/��Ʈ)
'									sftNoUseYn		= itemList.get(j).sftNoUseYn		'�Ƚɹ�ȣ��뿩��
'									sftDvpTelNo		= itemList.get(j).sftDvpTelNo		'�Ƚɹ������ȭ��ȣ
'									sftDvpMphnNo	= itemList.get(j).sftDvpMphnNo		'�Ƚɹ�����޴�����ȣ
'									cmbnDvGrpNo		= itemList.get(j).frstDvCst			'�չ�۱׷��ȣ

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

										'' CS ����������. ������Ʈ
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
					rw "�ֹ���� CS�Է°Ǽ�:"&iInputCnt
				Else
					rw resultmsg
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
end function

'��ǰ(��û)��� ��� ��ȸ
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
		objJson("srchStrtDttm") = startdate		'#�˻��������� yyyyMMddhh24miss
		objJson("srchEndDttm") = enddate		'#�˻��������� yyyyMMddhh24miss
		objJson("odNo") = ""					'o�ֹ���ȣ( only �ֹ���ȣ�϶��� �ʼ�)
		objJson("lrtrNo") = ""					'o�����ŷ�ó��ȣ
		objJson("odTypCd") = "41"				'�ֹ������ڵ�(41 ��ǰ��� 50 AS)(NULL �ΰ�� ��ü��ȸ)
		strRst = objJson.jsString
	SET objJson = nothing

	'// ����Ÿ ��������
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
							OutMallOrderSerial	= datalist.get(i).odNo				'�ֹ���ȣ
							CSDetailKey			= datalist.get(i).clmNo				'Ŭ���ӹ�ȣ
							Set itemList = datalist.get(i).itemList
								For j = 0 to itemList.length-1
									OrgDetailKey	= itemList.get(j).odSeq				'�ֹ�����
'									procSeq			= itemList.get(j).procSeq			'ó������
'									orglProcSeq		= itemList.get(j).orglProcSeq		'��ó������
'									odTypCd			= itemList.get(j).odTypCd			'�ֹ������ڵ�(41 ��ǰ��� 50 AS)
'									odTypDtlCd		= itemList.get(j).odTypDtlCd		'�ֹ��������ڵ�(5031 �κ�ǰȸ�����)
'									odPrgsStepCd	= itemList.get(j).odPrgsStepCd		'�ֹ�����ܰ��ڵ�(02 ��û 21 ��ҿϷ� 22 ���öȸ ), �迭��� 21 ��ҿϷḸ ���
'									spdNo			= itemList.get(j).spdNo				'�Ǹ��ڻ�ǰ��ȣ
'									spdNm			= itemList.get(j).spdNm				'�Ǹ��ڻ�ǰ��
'									sitmNo			= itemList.get(j).sitmNo			'�Ǹ��� ��ǰ��ȣ
'									sitmNm			= itemList.get(j).sitmNm			'�Ǹ��� ��ǰ��
									itemno			= itemList.get(j).odQty				'�ֹ�����
'									itmSlPrc		= itemList.get(j).itmSlPrc			'��ǰ�ǸŰ�, (������ ������ �Ǹſ���)
'									rtngQty			= itemList.get(j).rtngQty			'��ǰ����
'									trNo			= itemList.get(j).trNo				'�ŷ�ó��ȣ
'									lrtrNo			= itemList.get(j).lrtrNo			'�����ŷ�ó��ȣ
									OutMallRegDate	= itemList.get(j).clmReqDttm		'Ŭ���ӿ�û�Ͻ�[yyyyMMddHHmmss : 20191201121212] - ��ǰ��û�Ͻ�
'									clmCmptDttm		= itemList.get(j).clmCmptDttm		'Ŭ���ӿϷ��Ͻ�[yyyyMMddHHmmss : 20191201121212] - ��ǰ�Ϸ��Ͻ�

									strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
									strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
									strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
									strSql = strSql & " 	('" & divcd & "', '��Ÿ', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
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

										'' CS ����������. ������Ʈ
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
							rw "��ǰöȸ CS�Է°Ǽ�:"&iInputCnt
						Next
					Set datalist = nothing
				Else
					rw resultmsg
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
end function

'��ȯ��û/���������ȸ
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
		objJson("srchStrtDttm") = startdate		'#�˻��������� yyyyMMddhh24miss
		objJson("srchEndDttm") = enddate		'#�˻��������� yyyyMMddhh24miss
		objJson("odNo") = ""					'o�ֹ���ȣ( only �ֹ���ȣ�϶��� �ʼ�)
		objJson("lrtrNo") = ""					'o�����ŷ�ó��ȣ
		strRst = objJson.jsString
	SET objJson = nothing

	'// ����Ÿ ��������
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
							OutMallOrderSerial	= datalist.get(i).odNo				'�ֹ���ȣ
							CSDetailKey			= datalist.get(i).clmNo				'Ŭ���ӹ�ȣ
rw iRbody
							Set itemList = datalist.get(i).itemList
								For j = 0 to itemList.length-1
									OrgDetailKey	= itemList.get(j).odSeq				'�ֹ�����
'									procSeq			= itemList.get(j).procSeq			'ó������ : Default 1 ��ǰ������ ó���������� �������� �Է½� 1 �̰� Ŭ������ �߻��� ��� 1�� ������
'									orglProcSeq		= itemList.get(j).orglProcSeq		'��ó������
'									odTypCd			= itemList.get(j).odTypCd			'�ֹ����� | 10 : �ֹ� ,20 : ���(�ֹ����), 30 : ��ȯ, 31 : ��ȯ���, 40 : ��ǰ, 41 : ��ǰ���, 50 : AS
'									odTypDtlCd		= itemList.get(j).odTypDtlCd		'�ֹ��������ڵ�( 5010 ���� 5020 �߰���� (�κ�ǰ))
'									odPrgsStepCd	= itemList.get(j).odPrgsStepCd		'�ֹ�����ܰ�(��ȯ��û) | 02 : ��û, 03 : ����, 11 : �������, 12 : ��ǰ�غ�, 13 : �߼ۿϷ�, 14 : ��ۿϷ�, 15 : ����Ϸ�, 21 : ��ҿϷ�, 22 : ���öȸ, 23 : ȸ������, 24 : ȸ������, 25 : ȸ���Ϸ�, 26 : ȸ��Ȯ��, 27 : ��ǰ�Ϸ�
'									spdNo			= itemList.get(j).spdNo				'�Ǹ��ڻ�ǰ��ȣ
'									spdNm			= itemList.get(j).spdNm				'�Ǹ��ڻ�ǰ��
'									sitmNo			= itemList.get(j).sitmNo			'�Ǹ��� ��ǰ��ȣ
'									sitmNm			= itemList.get(j).sitmNm			'�Ǹ��� ��ǰ��
									itemno			= itemList.get(j).odQty				'�ֹ�����
'									itmSlPrc		= itemList.get(j).itmSlPrc			'��ǰ�ǸŰ�
'									xchgQty			= itemList.get(j).xchgQty			'��ȯ����
'									trNo			= itemList.get(j).trNo				'�ŷ�ó��ȣ
'									lrtrNo			= itemList.get(j).lrtrNo			'�����ŷ�ó��ȣ
									OutMallRegDate	= itemList.get(j).odAccpDttm		'�ֹ������Ͻ� [yyyyMMddHHmmss]
'									purCfrmDttm		= itemList.get(j).purCfrmDttm		'����Ȯ���Ͻ� [yyyyMMddHHmmss]
'									clmReqDttm		= itemList.get(j).clmReqDttm		'Ŭ���ӿ�û�Ͻ� [yyyyMMddHHmmss]
'									clmAccpDttm		= itemList.get(j).clmAccpDttm		'Ŭ���������Ͻ�[yyyyMMddHHmmss : 20191201121212] - ��ȯ�����Ͻ�
'									clmCmptDttm		= itemList.get(j).clmCmptDttm		'Ŭ���ӿϷ��Ͻ�[yyyyMMddHHmmss : 20191201121212] - ��ȯ�Ϸ��Ͻ�
									gubunname		= itemList.get(j).clmRsnCd			'Ŭ���ӻ����ڵ� | 201 : ��ǰ�� ���ڰ� ����(�ļ�/�ҷ�), 202 : �ٸ� ��ǰ�� ��۵�, 203 : ��ǰ�� ������ �ٸ�(��ǰ��������), 204 : �� ����� ������

									Select Case gubunname
										Case "201"
											gubunname = "��ǰ�ҷ�"
										Case "202"
											gubunname = "�����"
										Case "203"
											gubunname = "��ǰ����Ʋ��"
										Case "204"
											gubunname = "�ܼ�����"
										Case Else
											gubunname = "��Ÿ"
									End Select

'									clmRsnCnts		= itemList.get(j).clmRsnCnts		'Ŭ���ӻ�������
'									spicYn			= itemList.get(j).spicYn			'����Ʈ�ȿ��� (Y/N)
'									dvRtrvDvsCd		= itemList.get(j).dvRtrvDvsCd		'���ȸ�������ڵ� | RTRV:ȸ��, DV:���
'									frstDvCst		= itemList.get(j).frstDvCst			'�ʵ���ۺ�
'									rtngAddDvCst	= itemList.get(j).rtngAddDvCst		'��ǰ �߰���ۺ�
'									rtngDvCst		= itemList.get(j).rtngDvCst			'��ǰ��ۺ� - ��û�� ��쿡�� ���� ����
'									xchgAddDvCst	= itemList.get(j).xchgAddDvCst		'��ȯ �߰���ۺ�
'									xchgDvCst		= itemList.get(j).xchgDvCst			'��ȯ��ۺ� - ��û�� ��쿡�� ���� ����
'									dvCstBdnMnbdCd	= itemList.get(j).dvCstBdnMnbdCd	'��ۺ�δ���ü�ڵ� (01:�� / 02: ��ü)
'									shopCnvMsg		= itemList.get(j).shopCnvMsg		'�������޸޽���
'									xchgDvsCd		= itemList.get(j).xchgDvsCd			'��ȯ�����ڵ�(01:�Ϲݱ�ȯ,02:�±�ȯ)
'									cmbnDvPsbYn		= itemList.get(j).cmbnDvPsbYn		'�չ�۰��ɿ��� (Y/N)
'									cmbnDvGrpNo		= itemList.get(j).cmbnDvGrpNo		'�չ�۱׷��ȣ
'									dvLrtrNo		= itemList.get(j).dvLrtrNo			'��������ŷ�ó��ȣ-���ڵ� �������� (����/��Ʈ)
'									hpDvDttm		= itemList.get(j).hpDvDttm			'�������Ͻ� yyyymmddhh24miss (����/��Ʈ) - ������� ����� �� �����ۻ�ǰ�� ����Ͻ� �� ��ġ��ǰ ��������
'									afflSqncTypCd	= itemList.get(j).afflSqncTypCd		'�迭��ȸ�������ڵ�(����/��Ʈ) | CRSS_PIC:ũ�ν���, DN_DV:�������, DRIVE_PIC:����̺���, DRT_DV:�ٷι��, FPRD_DV:������, FS_PKUP:�����Ⱦ�, RENTAL_CAR_PKUP:����ī�Ⱦ�, SHOP_DV:������, SMT_QCK:����Ʈ��, STR_PIC:�������
'									afflCbCd		= itemList.get(j).afflCbCd			'�迭��ť���ڵ�(����/��Ʈ)
'									dvSqncCd		= itemList.get(j).dvSqncCd			'ȸ���ڵ� ����On���� ȸ����ȣ(����/��Ʈ)
'									rtrvSeq			= itemList.get(j).rtrvSeq			'ȸ��������(���������)
'									rtrvCustNm		= itemList.get(j).rtrvCustNm		'ȸ��������
'									rtrvTelNo		= itemList.get(j).rtrvTelNo			'ȸ������ȭ��ȣ
'									rtrvMphnNo		= itemList.get(j).rtrvMphnNo		'ȸ�����޴�����ȣ
'									rtrvZipNo		= itemList.get(j).rtrvZipNo			'ȸ���������ȣ
'									rtrvZipNoSeq	= itemList.get(j).rtrvZipNoSeq		'ȸ���������ȣ����(��Ʈ)
'									rtrvStnmZipAddr	= itemList.get(j).rtrvStnmZipAddr	'ȸ�������θ�����ּ�
'									rtrvStnmDtlAddr	= itemList.get(j).rtrvStnmDtlAddr	'ȸ�������θ���ּ�
'									dvpSeq			= itemList.get(j).dvpSeq			'���������(���������)
'									dvpCustNm		= itemList.get(j).dvpCustNm			'���������
'									dvpTelNo		= itemList.get(j).dvpTelNo			'�������ȭ��ȣ
'									dvpMphnNo		= itemList.get(j).dvpMphnNo			'������޴�����ȣ
'									dvpZipNo		= itemList.get(j).dvpZipNo			'����������ȣ
'									dvpZipNoSeq		= itemList.get(j).dvpZipNoSeq		'����������ȣ����(��Ʈ)
'									dvpStnmZipAddr	= itemList.get(j).dvpStnmZipAddr	'������⺻�ּ�
'									dvpStnmDtlAddr	= itemList.get(j).dvpStnmDtlAddr	'��������ּ�
'									dvMsg			= itemList.get(j).dvMsg				'��۸޽���
'									sftNoUseYn		= itemList.get(j).sftNoUseYn		'�Ƚɹ�ȣ��뿩�� (ȸ��������)
'									sftDvpTelNo		= itemList.get(j).sftDvpTelNo		'�Ƚɹ������ȭ��ȣ (ȸ��������)
'									sftDvpMphnNo	= itemList.get(j).sftDvpMphnNo		'�Ƚɹ�����޴�����ȣ (ȸ��������)
'									jntFtdrPwd		= itemList.get(j).jntFtdrPwd		'����������й�ȣ
'									dnDvRcptOptCd	= itemList.get(j).dnDvRcptOptCd		'��ۼ��ɿɼ� �ڵ� | 10 : ���չ��, 20 : ���ǹ��
'									dnDvVstMthdCd	= itemList.get(j).dnDvVstMthdCd		'��۹湮����ڵ� | 10:����������й�ȣ, 20:�������԰���, 30:����ȣ��, 40:�������ѿ���, 50:������ȭ
'									rpcSpdNo		= itemList.get(j).rpcSpdNo			'��ü�Ǹ��ڻ�ǰ��ȣ (��Ʈ/���� �����Ұ�� ����)
'									rpcSpdNm		= itemList.get(j).rpcSpdNm			'��ü�Ǹ��ڻ�ǰ�� (��Ʈ/���� �����Ұ�� ����)
'									rpcSitmNo		= itemList.get(j).rpcSitmNo			'��ü�Ǹ��ڴ�ǰ��ȣ (��Ʈ/���� �����Ұ�� ����)
'									rpcSitmNm		= itemList.get(j).rpcSitmNm			'��ü�Ǹ��ڴ�ǰ�� (��Ʈ/���� �����Ұ�� ����)

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

										'' CS ����������. ������Ʈ
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
					rw "��ȯ CS�Է°Ǽ�:"&iInputCnt
				Else
					rw resultmsg
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
end function

'��ȯ(��û)��� ��� ��ȸ
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
		objJson("srchStrtDttm") = startdate		'#�˻��������� yyyyMMddhh24miss
		objJson("srchEndDttm") = enddate		'#�˻��������� yyyyMMddhh24miss
		objJson("odNo") = ""					'o�ֹ���ȣ( only �ֹ���ȣ�϶��� �ʼ�)
		objJson("lrtrNo") = ""					'o�����ŷ�ó��ȣ
		strRst = objJson.jsString
	SET objJson = nothing

	'// ����Ÿ ��������
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
							OutMallOrderSerial	= datalist.get(i).odNo				'�ֹ���ȣ
							CSDetailKey			= datalist.get(i).clmNo				'Ŭ���ӹ�ȣ
							Set itemList = datalist.get(i).itemList
								For j = 0 to itemList.length-1
									OrgDetailKey	= itemList.get(j).odSeq				'�ֹ�����
'									procSeq			= itemList.get(j).procSeq			'ó������
'									orglProcSeq		= itemList.get(j).orglProcSeq		'��ó������
'									odTypCd			= itemList.get(j).odTypCd			'�ֹ������ڵ�(31 ��ȯ���)
'									odPrgsStepCd	= itemList.get(j).odPrgsStepCd		'�ֹ�����ܰ��ڵ�(02 ��û 21 ��ҿϷ� 22 ���öȸ ), �迭��� 21 ��ҿϷḸ ���
'									dvRtrvDvsCd		= itemList.get(j).dvRtrvDvsCd		'���ȸ�������ڵ� | RTRV:ȸ��, DV:���
'									spdNo			= itemList.get(j).spdNo				'�Ǹ��ڻ�ǰ��ȣ
'									spdNm			= itemList.get(j).spdNm				'�Ǹ��ڻ�ǰ��
'									sitmNo			= itemList.get(j).sitmNo			'�Ǹ��� ��ǰ��ȣ
'									sitmNm			= itemList.get(j).sitmNm			'�Ǹ��� ��ǰ��
									itemno			= itemList.get(j).odQty				'�ֹ�����
'									itmSlPrc		= itemList.get(j).itmSlPrc			'��ǰ�ǸŰ�, (������ ������ ��ǰ �Ǹſ���)
'									xchgQty			= itemList.get(j).xchgQty			'��ȯ����
'									trNo			= itemList.get(j).trNo				'�ŷ�ó��ȣ
'									odAccpDttm		= itemList.get(j).odAccpDttm		'�ֹ������Ͻ� [yyyyMMddHHmmss]
									OutMallRegDate	= itemList.get(j).clmReqDttm		'Ŭ���ӿ�û�Ͻ�[yyyyMMddHHmmss : 20191201121212] - ��ȯöȸ��û�Ͻ�
'									clmCmptDttm		= itemList.get(j).clmCmptDttm		'Ŭ���ӿϷ��Ͻ�[yyyyMMddHHmmss : 20191201121212] - ��ȯöȸ�Ϸ��Ͻ�

									strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
									strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
									strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
									strSql = strSql & " 	('" & divcd & "', '��Ÿ', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
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

										'' CS ����������. ������Ʈ
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
					rw "��ȯöȸ CS�Է°Ǽ�:"&iInputCnt
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
			iaccessLicense = "01000100004b035a25d67f991849cad1c7042b8da528d13e9ddce6878f2e43ac88080e0a5e" 'AccessLicense Key �Է�, PDF��������
			osecretKey = "AQABAAAWPWagCrPjFQnFEtxs5j+oyZFwuzomdNq0XZSricPuMw=="  'SecreKey �Է�, PDF��������
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		Else
			iaccessLicense = "010001000019133c715650b9c85b820961612f2b90b431ddd8654b42c097c4df1a43d0be09" 'AccessLicense Key �Է�, PDF��������
			osecretKey = "AQABAADX6Hz/wORFJS5pSIy4KQXkH83gC9G1aXChxBjcnUMqWw=="  'SecreKey �Է�, PDF��������
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		End If
	Set cryptoLib = nothing
End Function

Public Function generateKey_nvstorefarm(iTimestamp)
	Dim cryptoLib, oLicense, osecretKey, otimeStamp, osignature
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
		If (application("Svr_Info") = "Dev") Then
			osecretKey = "AQABAAAWPWagCrPjFQnFEtxs5j+oyZFwuzomdNq0XZSricPuMw=="  'SecreKey �Է�, PDF��������
			generateKey_nvstorefarm = cryptoLib.generateKey(iTimestamp, osecretKey)
		Else
			osecretKey = "AQABAADX6Hz/wORFJS5pSIy4KQXkH83gC9G1aXChxBjcnUMqWw=="  'SecreKey �Է�, PDF��������
			generateKey_nvstorefarm = cryptoLib.generateKey(iTimestamp, osecretKey)
		End If
	Set cryptoLib = nothing
End Function

'// 11111111111111
'// /outmall/gseshop/gseshopItemcls.asp ����
CONST CGSShopCompanyCode = 1003890	'' ���»��ڵ�
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

	'// API URL(�Ⱓ������ �ֹ� ��������)
	'// tnsType : �ֹ�����(�ֹ�/��ǰ : S, ��� : C)
	'// ���� : test1 � : ecb2b
	if (application("Svr_Info") = "Dev") then
		xmlURL = "http://test1.gsshop.com/SupSendOrderInfo.gs?supCd=" + CStr(CGSShopCompanyCode) + "&sdDt=" + CStr(xmlSelldate) + "&tnsType=C"
	else
		xmlURL = "http://ecb2b.gsshop.com/SupSendOrderInfo.gs?supCd=" + CStr(CGSShopCompanyCode) + "&sdDt=" + CStr(xmlSelldate) + "&tnsType=C"
	end if
	''response.write xmlURL & "<br />"
	''dbget.close : response.end

	'// =======================================================================
	'// ����Ÿ ��������
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.setTimeouts 2000,2000,2000,2000

	on error resume next
		objXML.send()


		if objXML.Status <> "200" then
			response.write "ERROR : ��ſ���"
			''dbget.close : response.end
		end if

	on error goto 0
	'// ���ۿ�û�� �Ѵ�.(XML ����X)

	'// /wapi/outmall/order/xSiteOrder_GSShop_recv_Process.asp ����

end function

'GSShop ��� ��ȸ
function GetCSOrderNewCancel_gseshop(sellsite, csGubun, csSubGubun, selldate)
	dim xmlURL, xmlSelldate, obj
	dim objXML, strObj, objData, jParam, requireDetailObj, requireDetail
	dim i, j, k, strsql
	dim successCnt : successCnt = 0
	Dim returnCode, resultMsg
	Dim apiUrl
	'// =======================================================================
	'// ��¥����
	xmlSelldate = Replace(selldate, "-", "")
	If (application("Svr_Info") = "Dev") Then
		apiUrl = "http://realapi.gsshop.com/b2b/SupSendOrderInfo.gs"
		'apiUrl = "http://testapi.gsshop.com/b2b/SupSendOrderInfo.gs"
	Else
		apiUrl = "http://realapi.gsshop.com/b2b/SupSendOrderInfo.gs"
	End If

	'// =======================================================================
	Set obj = jsObject()
		obj("sender") = "TBT"					'������ (���»纰 �ο��Ǵ� ID�� GS���� ����)
		obj("receiver") = "GS SHOP"				'������ | GS SHOP
		obj("documentId") = "ORDINF"			'����ID | ORDINF
		obj("processType") = "C"				'���۱��� | A:��ü, S:�ֹ�/��ǰ, C:���
		obj("supCd") = ""&CGSShopCompanyCode&""	'���»��ڵ�	| (���»��ȣ GS���� ����)
		obj("sdDt") = xmlSelldate				'��ȸ����
		jParam = obj.jsString
	Set obj = nothing

	'// ����Ÿ ��������
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(jParam)
		If objXML.Status <> "200" Then
			response.write "ERROR : ��ſ���"
			dbget.close : response.end
		Else
			objData = BinaryToText(objXML.ResponseBody,"utf-8")
		End If

	'// =======================================================================
	'// Json �Ľ�
		Set strObj = JSON.parse(objData)
			returnCode		= strObj.resultCd
			resultMsg		= strObj.resultMsg
			If returnCode <> "S" Then
				response.write "ERROR : ����" & resultMsg
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

'// ��ǰ(���)��û �ܰ� ��ȸ
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

                    ''RELEASE_STOP_UNCHECKED	���������û
                    ''RETURNS_UNCHECKED	��ǰ����
                    ''VENDOR_WAREHOUSE_CONFIRM	�԰�Ϸ�
                    ''REQUEST_COUPANG_CHECK	����Ȯ�ο�û
                    ''RETURNS_COMPLETED	��ǰ�Ϸ�
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

	''rw "CS ���(��ǰ) �Ǽ�:" & iAssignedRow

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

    '// ���߸�Ī ����
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

    '// git ���ε� Ȯ��
    if OutMallOrderSerialArr = "" then
        rw "��������"
        dbget.close() : response.end
    end if

    affectedRows = 0
    OutMallOrderSerialArr = Split(OutMallOrderSerialArr, ",")
    for i = 0 to UBound(OutMallOrderSerialArr)
        OutMallOrderSerial = OutMallOrderSerialArr(i)
        if OutMallOrderSerial <> "" then
            affectedRows = fnMatchCs(sellsite, OutMallOrderSerial)
            Call fnUnmatchDeletedCS(sellsite, OutMallOrderSerial)

            rw OutMallOrderSerial & " : " & affectedRows & " �� �ݿ���"

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
	''strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_list] a on T.asid = a.id and a.currstate = 'B007' and a.deleteyn = 'N' "		'// �������� ������ üũ
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

    '// git ���ε� Ȯ��
    if divcdArr = "" then
        rw "��������"
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
                        rw csdetailkey1 & " : " & affectedRows & " �� �ݿ���"

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
