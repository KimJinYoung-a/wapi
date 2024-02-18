<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteCSOrderCls.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/outmall/nvstorefarm/nvstorefarmItemcls.asp"-->
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
    If (sellsite="nvstorefarm") Then
    	ErrMsg = ""
		''###################### �Ķ� ���� ##########################
		Dim oServ, occd, oinputDT
		oServ		= "SellerService41"
		occd		= "GetChangedProductOrderList"
		oinputDT	= getLastCSInputDT
		For i = 0 To 1
			Call getNvstorefarmChangeOrder(oServ, occd, i, oinputDT)
		Next
		'############################################################
    End If
End If

Public Function getsecretKey(iaccessLicense, iTimestamp, isignature, iserv, ioper)
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

Function getLastCSInputDT()
	Dim sqlStr
	sqlStr = "select top 1 LastCheckDate as lastCSInputDt"
	sqlStr = sqlStr&" from db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	sqlStr = sqlStr&" where sellsite = 'nvstorefarm'  "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If (Not rsget.Eof) Then
		getLastCSInputDT = rsget("lastCSInputDt")&"T00:00:00"
	Else
		getLastCSInputDT = "2017-09-20T00:00:00"
	End If
	rsget.Close
End Function

Function UpdateLastCSInputDT(dt)
	Dim sqlStr
	sqlStr = " UPDATE db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	sqlStr = sqlStr & " SET LastCheckDate = '" & CStr(dt) & "' "
	sqlStr = sqlStr & " WHERE sellsite = 'nvstorefarm'  "
	dbget.Execute sqlStr
End Function

Public Sub getNvstorefarmChangeOrder(iServ, iccd, lp, iinputDT)
	Dim reqID, LastChangedStatusCode
	Dim strRst, oaccessLicense, oTimestamp, osignature, oOper, stdt, eddt
	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		reqID = "tenten"
	End If

	Call getsecretKey(oaccessLicense, oTimestamp, osignature, iServ, iccd)

	stdt = iinputDT
	eddt = Left(stdt, 10) & "T23:59:59"

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">"
	strRst = strRst & "	<soapenv:Header/>"
	strRst = strRst & "	<soapenv:Body>"
	strRst = strRst & "		<sel:GetChangedProductOrderListRequest>"
	strRst = strRst & "			<sel:AccessCredentials>"
	strRst = strRst & "				<sel:AccessLicense>"&oaccessLicense&"</sel:AccessLicense>"
	strRst = strRst & "				<sel:Timestamp>"&oTimestamp&"</sel:Timestamp>"
	strRst = strRst & "				<sel:Signature>"&osignature&"</sel:Signature>"
	strRst = strRst & "			</sel:AccessCredentials>"
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>"
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>"							'#�����޴� �������� �� ����(Compact / Full)
	strRst = strRst & "			<sel:Version>4.1</sel:Version>"
	strRst = strRst & "			<sel:InquiryTimeFrom>"&stdt&"</sel:InquiryTimeFrom>"				'#��ȸ ���� �Ͻ�(�ش� �ð� ����)
	strRst = strRst & "			<sel:InquiryTimeTo>"&eddt&"</sel:InquiryTimeTo>"					'��ȸ ���� �Ͻ�(�ش� �ð� �������� ����)
	'	<!--Optional:-->
	'	strRst = strRst & "			<sel:InquiryExtraData>?</sel:InquiryExtraData>"				'��ȸ�� ����� �߰� ������(�� : �ֹ���ȣ)
	'	<!--Optional:-->
	Select Case lp
		Case "0"
			LastChangedStatusCode = "CANCEL_REQUESTED" '"CANCELED"	'2019-07-23 ������ CANCEL_REQUESTED�� ����
			getDivCD = "A008"
		Case "1"
			LastChangedStatusCode = "RETURN_REQUESTED" '"RETURNED"	'2019-07-23 ������ RETURN_REQUESTED�� ����
			getDivCD = "A004"
	End Select

	strRst = strRst & "			<sel:LastChangedStatusCode>"&LastChangedStatusCode&"</sel:LastChangedStatusCode>"	'���� ��ǰ �ֹ� ���� �ڵ� (CANCELED | ���, RETURNED | ��ǰ, EXCHANGED : ��ȯ)
	<!--Optional:-->
	strRst = strRst & "			<sel:MallID>"&reqID&"</sel:MallID>"							'�Ǹ��� ���̵�
	strRst = strRst & "		</sel:GetChangedProductOrderListRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"

	Dim nvstorefarmURL
	If (application("Svr_Info") = "Dev") Then
		nvstorefarmURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		nvstorefarmURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If

	Dim httpRequest, ResponseType, OrderInfoList
	Set httpRequest = CreateObject("MSXML2.XMLHTTP")

	httpRequest.Open "POST", nvstorefarmURL, False
	httpRequest.SetRequestHeader "Content-Type", "text/xml;charset=UTF-8"
	httpRequest.SetRequestHeader "SOAPAction", iServ & "#" & iccd
	httpRequest.send strRst
	If httpRequest.Status = 200 Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(httpRequest.responseText)
			ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
'				response.write (Replace(httpRequest.responseText,"soap:",""))
'				response.end
			If ResponseType = "SUCCESS" Then
				If xmlDOM.getElementsByTagName("n:ChangedProductOrderInfoList").length > 0 Then
					Set Nodes = xmlDOM.getElementsByTagName("n:ChangedProductOrderInfoList")
						Dim retVal, succCnt, failCnt
						Dim OutMallOrderSerial, OrgDetailKey, LastChangedStatus, LastChangedDate, ProductOrderStatus, ClaimType, ClaimStatus, PaymentDate, IsReceiverAddressChanged, GiftReceivingStatus
						Dim CSDetailKey, rcvrNm, rcvrTelNum, rcvrMobile, rcvrPost, rcvrAddr1, rcvrAddr2, orderDt, orderQty, sndNm, sndTelNum, sndMobile, orderReqContent
						succCnt = 0
						failCnt = 0
						For each SubNodes in Nodes
							If SubNodes.selectSingleNode("n:OrderID") is Nothing Then
								OutMallOrderSerial = ""
							Else
								OutMallOrderSerial = SubNodes.getElementsByTagName("n:OrderID")(0).Text							'�ֹ���ȣ
							End If

							If SubNodes.selectSingleNode("n:ProductOrderID") is Nothing Then
								OrgDetailKey = ""
							Else
								OrgDetailKey = SubNodes.getElementsByTagName("n:ProductOrderID")(0).Text						'��ǰ �ֹ� ��ȣ
							End If

							If SubNodes.selectSingleNode("n:LastChangedStatus") is Nothing Then
								LastChangedStatus = ""
							Else
								LastChangedStatus = SubNodes.getElementsByTagName("n:LastChangedStatus")(0).Text				'���� ���� ���� �ڵ�
							End If


							If SubNodes.selectSingleNode("n:LastChangedDate") is Nothing Then
								LastChangedDate = ""
							Else
								LastChangedDate = SubNodes.getElementsByTagName("n:LastChangedDate")(0).Text					'���� ���� �Ͻ�
							End If

							If SubNodes.selectSingleNode("n:ProductOrderStatus") is Nothing Then
								ProductOrderStatus = ""
							Else
								ProductOrderStatus = SubNodes.getElementsByTagName("n:ProductOrderStatus")(0).Text				'��ǰ �ֹ� ���� �ڵ�
							End If

							If SubNodes.selectSingleNode("n:ClaimType") is Nothing Then
								ClaimType = ""
							Else
								ClaimType = SubNodes.getElementsByTagName("n:ClaimType")(0).Text								'Ŭ���� Ÿ�� �ڵ�
							End If

							If SubNodes.selectSingleNode("n:ClaimStatus") is Nothing Then
								ClaimStatus = ""
							Else
								ClaimStatus = SubNodes.getElementsByTagName("n:ClaimStatus")(0).Text							'Ŭ���� ó�� ���� �ڵ�
							End If

							If SubNodes.selectSingleNode("n:PaymentDate") is Nothing Then
								PaymentDate = ""
							Else
								PaymentDate = LEFT(SubNodes.getElementsByTagName("n:PaymentDate")(0).Text,10)					'���� �Ͻ�
							End If

							If SubNodes.selectSingleNode("n:IsReceiverAddressChanged") is Nothing Then
								IsReceiverAddressChanged = ""
							Else
								IsReceiverAddressChanged = SubNodes.getElementsByTagName("n:IsReceiverAddressChanged")(0).Text	'����� ���� ���� ����
							End If

							If SubNodes.selectSingleNode("n:GiftReceivingStatus") is Nothing Then
								GiftReceivingStatus = ""
							Else
								GiftReceivingStatus = SubNodes.getElementsByTagName("n:GiftReceivingStatus")(0).Text			'���� ���� ���� �ڵ�
							End If

'							rw OutMallOrderSerial
'							rw OrgDetailKey
'							rw LastChangedStatus
'							rw LastChangedDate
'							rw ProductOrderStatus
'							rw ClaimType
'							rw ClaimStatus
'							rw PaymentDate
'							rw IsReceiverAddressChanged
'							rw GiftReceivingStatus
'							rw "------------------------------------------------------------" & lp

							strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = 'nvstorefarm' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "') "
							strSql = strSql & " BEGIN "
							strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
							strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
							strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
							strSql = strSql & " 	('"&getDivCD&"', '�ܼ�����', 'nvstorefarm', '" & html2db(CStr(OutMallOrderSerial)) & "', '', '', '', '', '', "
							strSql = strSql & "		'', '', '', '', '', '' "
							strSql = strSql & "		, '" & html2db(CStr(PaymentDate)) & "', '" & html2db(CStr(OrgDetailKey)) & "', '" & html2db(CStr(CSDetailKey)) & "', '') "
							strSql = strSql & " END "
							dbget.Execute strSql
						Next
					Set Nodes = nothing

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
					strSql = strSql + " 	and c.sellsite = 'nvstorefarm' "
					''rw strSql
					dbget.Execute strSql
				End If
			End If
			If CDate(Left(stdt, 10)) < Date() Then
				UpdateLastCSInputDT(DateAdd("d", 1, Left(stdt, 10)))
			End If
		Set xmlDOM = nothing
	End If
End Sub
%>
<% if  (IsAutoScript) then  %>
<% rw "OK" %>
<% else %>
<script>alert('<%= Left(oinputDT, 10) %>�� �����Ͽ����ϴ�');</script>
<script>window.close();</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
