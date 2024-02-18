<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteCSOrderCls.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/outmall/auction/auctionItemcls.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<%
'// 2014-08-27, skyer9
''Server.ScriptTimeout = 60
'' response.write lotteAuthNo
'' response.end
Dim refer
refer = request.ServerVariables("HTTP_REFERER")

Dim sqlStr, buf
Dim i, j, k

'�ֹ��Ϸ� (1001) /����غ��� (1002) /����� (1003) /����Ϸ� (1004) /�ֹ���� (1005) /��ǰ��û (1007)
'��ǰ�Ϸ� (1008) /��ȯ��û (1011) /��ȯ�Ϸ� (1012) /��ǰ�� �ֹ���� (1009) /���� (1010)/ǰ����ҿ�û (1013)/ǰ����� (1014)

'// ============================================================================
'// [divcd]
'// ============================================================================
'A008			�ֹ����
'
'A004			��ǰ����(��ü���)
'A010			ȸ����û(�ٹ����ٹ��)
'
'A001			������߼�
'A002			���񽺹߼�
'
'A000			�±�ȯ���
'A100			��ǰ���� �±�ȯ���
'
'A009			��Ÿ����
'A006			�������ǻ���
'A700			��ü��Ÿ����
'
'A003			ȯ��
'A005			�ܺθ�ȯ�ҿ�û
'A007			ī��,��ü,�޴�����ҿ�û
'
'A011			�±�ȯȸ��(�ٹ����ٹ��)
'A012			�±�ȯ��ǰ(��ü���)

'A111			��ǰ���� �±�ȯȸ��(�ٹ����ٹ��)
'A112			��ǰ���� �±�ȯ��ǰ(��ü���)
'// ============================================================================

Dim mode
Dim sellsite
Dim reguserid
Dim AssignedRow
Dim ErrMsg

Dim resultCount

Dim divcd, yyyymmdd, idx, finUserid
Dim getDivCD, sDate, eDate

Dim postParam
Dim objXML, xmlDOM, strSql
Dim retCode, goodsCd, iMessage, oMsg, ocount
Dim SubNodes, Nodes
mode		= requestCheckVar(html2db(request("mode")),32)
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
idx			= requestCheckVar(html2db(request("idx")),32)
finUserid	= session("ssBctID")
If finUserid = "" Then
	finUserid = "system"
End If

If (mode = "getxsitecslist") Then
    If (sellsite="auction1010") Then
    	Dim oinputDT
    	ErrMsg = ""
    	oinputDT	= getLastCSInputDT
		''###################### �Ķ� ���� ##########################
		For i = 0 To 10
			if (oinputDT <= Left(Now, 10)) then
				Call getAuctionChangeOrder(0, oinputDT)			'// ���
				Call getAuctionChangeOrder(1, oinputDT)			'// ��ǰ
				Call getAuctionChangeOrder(2, oinputDT)			'// ��ȯ
				oinputDT = Left(DateAdd("d", 1, Left(oinputDT, 10)), 10)
			end if
		Next
		'############################################################
    End If
End If

Function getLastCSInputDT()
	Dim sqlStr
	sqlStr = "select top 1 LastCheckDate as lastCSInputDt"
	sqlStr = sqlStr&" from db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	sqlStr = sqlStr&" where sellsite = 'auction1010'  "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If (Not rsget.Eof) Then
		getLastCSInputDT = rsget("lastCSInputDt")
	Else
		getLastCSInputDT = "2017-10-28"
	End If
	rsget.Close
End Function

Function UpdateLastCSInputDT(dt)
	Dim sqlStr
	sqlStr = " UPDATE db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	sqlStr = sqlStr & " SET LastCheckDate = '" & CStr(dt) & "' "
	sqlStr = sqlStr & " WHERE sellsite = 'auction1010'  "
	dbget.Execute sqlStr
End Function

Public Sub getAuctionChangeOrder(lp, iinputDT)
	Dim strRst, iccd, resFirstChildname, loopOrderName
	Dim stdt, eddt
	stdt = iinputDT
	eddt = DateAdd("d", 1, Left(stdt, 10))
	If lp = 0 Then		'��Ұ�
		getDivCD	= "A008"
		iccd		= "GetOrderCanceledList"
		resFirstChildname = "GetOrderCanceledListResponse"
		loopOrderName = "OrderCanceled"

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "	<soap:Header>"
		strRst = strRst & "		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
		strRst = strRst & "			<Value>"&auctionTicket&"</Value>"
		strRst = strRst & "		</EncryptedTicket>"
		strRst = strRst & "	</soap:Header>"
		strRst = strRst & "	<soap:Body>"
		strRst = strRst & "		<GetOrderCanceledList xmlns=""http://www.auction.co.kr/APIv1/AuctionService"">"
		strRst = strRst & "			<req SearchType=""Nothing"" CategoryID="""">"
		strRst = strRst & "				<SearchDuration StartDate="""&stdt&""" EndDate="""&eddt&""" xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"" />"
		strRst = strRst & "			</req>"
		strRst = strRst & "		</GetOrderCanceledList>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
	ElseIf lp = 1 Then	'��ǰ��
		getDivCD	= "A004"
		iccd		= "GetReturnList"
		resFirstChildname = "GetReturnListResponse"
		loopOrderName = "ReturnList"
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "	<soap:Header>"
		strRst = strRst & "		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
		strRst = strRst & "			<Value>"&auctionTicket&"</Value>"
		strRst = strRst & "		</EncryptedTicket>"
		strRst = strRst & "	</soap:Header>"
		strRst = strRst & "	<soap:Body>"
		strRst = strRst & "		<GetReturnList xmlns=""http://www.auction.co.kr/APIv1/AuctionService"">"
		strRst = strRst & "			<req SearchType=""None"" SearchKeyword="""" SearchDateType=""Request"" PageSize=""5000"">"
		strRst = strRst & "				<SearchDuration StartDate="""&stdt&""" EndDate="""&eddt&""" xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"" />"
		strRst = strRst & "				<SearchFlags xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"">Requested</SearchFlags>"
		strRst = strRst & "			</req>"
		strRst = strRst & "		</GetReturnList>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
	ElseIf lp = 2 Then	'��ȯ��
		getDivCD	= "A000"
		iccd		= "GetExchangeRequestListBySearchCondition"
		resFirstChildname = "GetExchangeRequestListBySearchConditionResponse"
		loopOrderName = "ExchangeBase"
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "	<soap:Header>"
		strRst = strRst & "		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
		strRst = strRst & "			<Value>"&auctionTicket&"</Value>"
		strRst = strRst & "		</EncryptedTicket>"
		strRst = strRst & "	</soap:Header>"
		strRst = strRst & "	<soap:Body>"
		strRst = strRst & "		<GetExchangeRequestListBySearchCondition xmlns=""http://www.auction.co.kr/APIv1/AuctionService"">"
		strRst = strRst & "			<req SearchType=""None"" SearchKeyword="""" SearchDateType=""Request"" PageIndex=""1"">"
		strRst = strRst & "				<SearchDuration StartDate="""&stdt&""" EndDate="""&eddt&""" xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"" />"
		strRst = strRst & "				<SearchFlags xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"">Requested</SearchFlags>"
		strRst = strRst & "			</req>"
		strRst = strRst & "		</GetExchangeRequestListBySearchCondition>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
	End If

	Dim httpRequest, ResponseType, OrderInfoList
	Dim itemid, OrderNo, ItemName, AwardQty, AwardAmount, BuyerName, BuyerID, OutMallOrderSerial
	Dim RecieverName, ReturnReasonCode, BuyerTel, BuyerMobileTel, BuyerPostNo, BuyerAddressPost
	Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		httpRequest.open "POST", "" & auctionAPIURL&"/APIv1/AuctionService.asmx"
		httpRequest.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		httpRequest.setRequestHeader "Content-Length", LenB(strRst)
		httpRequest.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/AuctionService/"&iccd
		httpRequest.send(strRst)
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(httpRequest.responseText,"soap:",""))

response.write "<textarea cols=40 rows=10>"&BinaryToText(httpRequest.ResponseBody,"utf-8")&"</textarea>"
''response.end
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = resFirstChildname Then
				Set OrderInfoList = xmlDOM.getElementsByTagName(loopOrderName)
					If lp = 0 Then
						For Each SubNodes in OrderInfoList
							If (SubNodes.nodeType = 1 Or SubNodes.nodeType = 2) Then
								'On Error Resume Next
								OutMallOrderSerial= ""
								itemid = SubNodes.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("ItemID").value
								OrderNo = SubNodes.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("OrderNo").value
								ItemName = SubNodes.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("ItemName").value
								AwardQty = SubNodes.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("AwardQty").value
								AwardAmount = SubNodes.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("AwardAmount").value
								BuyerName = SubNodes.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("BuyerName").value
								BuyerID = SubNodes.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("BuyerID").value

								strSql = ""
								strSql = " SELECT TOP 1 OutmallOrderSerial FROM db_temp.[dbo].[tbl_xSite_tmporder] where sellsite = 'auction1010' and OrgDetailKey = '"& OrderNo &"' "
								rsget.CursorLocation = adUseClient
								rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
								If not rsget.EOF Then
									OutMallOrderSerial = rsget("OutmallOrderSerial")
								End If
								rsget.Close

								If OutmallOrderSerial <> "" Then
									strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = 'auction1010' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrderNo) & "' ) "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
									strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
									strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
									strSql = strSql & " 	('"&getDivCD&"', '�ܼ�����', 'auction1010', '" & html2db(CStr(OutMallOrderSerial)) & "', '', '', '', '', '', "
									strSql = strSql & "		'', '', '', '', '', '' "
									strSql = strSql & "		, '" & html2db(CStr(iinputDT)) & "', '" & html2db(CStr(OrderNo)) & "', '', '"&AwardQty&"') "
									strSql = strSql & " END "
									dbget.Execute strSql
								End If
							End If
						Next
					ElseIf lp = 1 Then
						For Each SubNodes in OrderInfoList
							If (SubNodes.nodeType = 1 Or SubNodes.nodeType = 2) Then
								'On Error Resume Next
								OutMallOrderSerial = SubNodes.attributes.GetNamedItem("PayNo").value
								RecieverName = SubNodes.attributes.GetNamedItem("RecieverName").value
								ReturnReasonCode = SubNodes.attributes.GetNamedItem("ReturnReasonCode").value

								BuyerTel = SubNodes.getElementsByTagName("Buyer")(0).attributes.GetNamedItem("Tel").value
								BuyerMobileTel = SubNodes.getElementsByTagName("Buyer")(0).attributes.GetNamedItem("MobileTel").value
								BuyerPostNo = SubNodes.getElementsByTagName("Buyer")(0).attributes.GetNamedItem("PostNo").value
								BuyerAddressPost = SubNodes.getElementsByTagName("Buyer")(0).attributes.GetNamedItem("AddressPost").value

								itemid = SubNodes.getElementsByTagName("Order")(0).attributes.GetNamedItem("ItemID").value
								OrderNo = SubNodes.getElementsByTagName("Order")(0).attributes.GetNamedItem("OrderNo").value
								ItemName = SubNodes.getElementsByTagName("Order")(0).attributes.GetNamedItem("ItemName").value
								AwardQty = SubNodes.getElementsByTagName("Order")(0).attributes.GetNamedItem("AwardQty").value
								AwardAmount = SubNodes.getElementsByTagName("Order")(0).attributes.GetNamedItem("AwardAmount").value
								BuyerName = SubNodes.getElementsByTagName("Order")(0).attributes.GetNamedItem("BuyerName").value

								strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = 'auction1010' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrderNo) & "' ) "
								strSql = strSql & " BEGIN "
								strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
								strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
								strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
								strSql = strSql & " 	('"&getDivCD&"', '�ܼ�����', 'auction1010', '" & html2db(CStr(OutMallOrderSerial)) & "', '"& BuyerName &"', '', '"& BuyerTel &"', '"& BuyerMobileTel &"', '"& RecieverName &"', "
								strSql = strSql & "		'', '', '', '', '', '' "
								strSql = strSql & "		, '" & html2db(CStr(iinputDT)) & "', '" & html2db(CStr(OrderNo)) & "', '', '"&AwardQty&"') "
								strSql = strSql & " END "
								dbget.Execute strSql
							End If
						Next
					Else
						For Each SubNodes in OrderInfoList
							If (SubNodes.nodeType = 1 Or SubNodes.nodeType = 2) Then
								'On Error Resume Next
								OutMallOrderSerial= ""
								itemid = SubNodes.attributes.GetNamedItem("ItemID").value
								OrderNo = SubNodes.attributes.GetNamedItem("OrderNo").value
								ItemName = SubNodes.attributes.GetNamedItem("ItemName").value
								AwardQty = SubNodes.attributes.GetNamedItem("Quantity").value
								AwardAmount = SubNodes.attributes.GetNamedItem("AwardAmount").value
								BuyerName = SubNodes.attributes.GetNamedItem("BuyerName").value

								strSql = ""
								strSql = " SELECT TOP 1 OutmallOrderSerial FROM db_temp.[dbo].[tbl_xSite_tmporder] where sellsite = 'auction1010' and OrgDetailKey = '"& OrderNo &"' "
								rsget.CursorLocation = adUseClient
								rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
								If not rsget.EOF Then
									OutMallOrderSerial = rsget("OutmallOrderSerial")
								End If
								rsget.Close

								If OutmallOrderSerial <> "" Then
									strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = 'auction1010' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrderNo) & "' ) "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
									strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
									strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
									strSql = strSql & " 	('"&getDivCD&"', '�ܼ�����', 'auction1010', '" & html2db(CStr(OutMallOrderSerial)) & "', '', '', '', '', '', "
									strSql = strSql & "		'', '', '', '', '', '' "
									strSql = strSql & "		, '" & html2db(CStr(iinputDT)) & "', '" & html2db(CStr(OrderNo)) & "', '', '"&AwardQty&"') "
									strSql = strSql & " END "
									dbget.Execute strSql
								End If
							End If
						Next
					End If
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
					strSql = strSql + " 	and c.OutMallOrderSerial = Replace(o.OutMallOrderSerial, '-', '') "
					strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
					strSql = strSql + " where "
					strSql = strSql + " 	1 = 1 "
					strSql = strSql + " 	and c.orderserial is NULL "
					strSql = strSql + " 	and o.orderserial is not NULL "
					strSql = strSql + " 	and c.sellsite = 'auction1010' "
					dbget.Execute strSql
				Set OrderInfoList = nothing
			End If

			If CDate(Left(stdt, 10)) < Date() Then
				UpdateLastCSInputDT(DateAdd("d", 1, Left(stdt, 10)))
			End If
		Set xmlDOM = nothing
	Set httpRequest = nothing
End Sub
%>
<% rw "OK" %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
