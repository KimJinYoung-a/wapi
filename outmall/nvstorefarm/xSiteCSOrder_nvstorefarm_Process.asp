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

'주문완료 (1001) /출고준비중 (1002) /배송중 (1003) /수취완료 (1004) /주문취소 (1005) /반품요청 (1007)
'반품완료 (1008) /교환요청 (1011) /교환완료 (1012) /반품후 주문취소 (1009) /오류 (1010)/품절취소요청 (1013)/품절취소 (1014)

'// ============================================================================
'// [divcd]
'// ============================================================================
'A008			주문취소
'
'A004			반품접수(업체배송)
'A010			회수신청(텐바이텐배송)
'
'A001			누락재발송
'A002			서비스발송
'
'A000			맞교환출고
'A100			상품변경 맞교환출고
'
'A009			기타사항
'A006			출고시유의사항
'A700			업체기타정산
'
'A003			환불
'A005			외부몰환불요청
'A007			카드,이체,휴대폰취소요청
'
'A011			맞교환회수(텐바이텐배송)
'A012			맞교환반품(업체배송)

'A111			상품변경 맞교환회수(텐바이텐배송)
'A112			상품변경 맞교환반품(업체배송)
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
		''###################### 파람 생성 ##########################
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
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>"							'#돌려받는 데이터의 상세 정도(Compact / Full)
	strRst = strRst & "			<sel:Version>4.1</sel:Version>"
	strRst = strRst & "			<sel:InquiryTimeFrom>"&stdt&"</sel:InquiryTimeFrom>"				'#조회 시작 일시(해당 시각 포함)
	strRst = strRst & "			<sel:InquiryTimeTo>"&eddt&"</sel:InquiryTimeTo>"					'조회 종료 일시(해당 시각 포함하지 않음)
	'	<!--Optional:-->
	'	strRst = strRst & "			<sel:InquiryExtraData>?</sel:InquiryExtraData>"				'조회에 사용할 추가 데이터(예 : 주문번호)
	'	<!--Optional:-->
	Select Case lp
		Case "0"
			LastChangedStatusCode = "CANCEL_REQUESTED" '"CANCELED"	'2019-07-23 김진영 CANCEL_REQUESTED로 변경
			getDivCD = "A008"
		Case "1"
			LastChangedStatusCode = "RETURN_REQUESTED" '"RETURNED"	'2019-07-23 김진영 RETURN_REQUESTED로 변경
			getDivCD = "A004"
	End Select

	strRst = strRst & "			<sel:LastChangedStatusCode>"&LastChangedStatusCode&"</sel:LastChangedStatusCode>"	'최종 상품 주문 상태 코드 (CANCELED | 취소, RETURNED | 반품, EXCHANGED : 교환)
	<!--Optional:-->
	strRst = strRst & "			<sel:MallID>"&reqID&"</sel:MallID>"							'판매자 아이디
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
								OutMallOrderSerial = SubNodes.getElementsByTagName("n:OrderID")(0).Text							'주문번호
							End If

							If SubNodes.selectSingleNode("n:ProductOrderID") is Nothing Then
								OrgDetailKey = ""
							Else
								OrgDetailKey = SubNodes.getElementsByTagName("n:ProductOrderID")(0).Text						'상품 주문 번호
							End If

							If SubNodes.selectSingleNode("n:LastChangedStatus") is Nothing Then
								LastChangedStatus = ""
							Else
								LastChangedStatus = SubNodes.getElementsByTagName("n:LastChangedStatus")(0).Text				'최종 변경 상태 코드
							End If


							If SubNodes.selectSingleNode("n:LastChangedDate") is Nothing Then
								LastChangedDate = ""
							Else
								LastChangedDate = SubNodes.getElementsByTagName("n:LastChangedDate")(0).Text					'최종 변경 일시
							End If

							If SubNodes.selectSingleNode("n:ProductOrderStatus") is Nothing Then
								ProductOrderStatus = ""
							Else
								ProductOrderStatus = SubNodes.getElementsByTagName("n:ProductOrderStatus")(0).Text				'상품 주문 상태 코드
							End If

							If SubNodes.selectSingleNode("n:ClaimType") is Nothing Then
								ClaimType = ""
							Else
								ClaimType = SubNodes.getElementsByTagName("n:ClaimType")(0).Text								'클레임 타입 코드
							End If

							If SubNodes.selectSingleNode("n:ClaimStatus") is Nothing Then
								ClaimStatus = ""
							Else
								ClaimStatus = SubNodes.getElementsByTagName("n:ClaimStatus")(0).Text							'클레임 처리 상태 코드
							End If

							If SubNodes.selectSingleNode("n:PaymentDate") is Nothing Then
								PaymentDate = ""
							Else
								PaymentDate = LEFT(SubNodes.getElementsByTagName("n:PaymentDate")(0).Text,10)					'결제 일시
							End If

							If SubNodes.selectSingleNode("n:IsReceiverAddressChanged") is Nothing Then
								IsReceiverAddressChanged = ""
							Else
								IsReceiverAddressChanged = SubNodes.getElementsByTagName("n:IsReceiverAddressChanged")(0).Text	'배송지 정보 수정 여부
							End If

							If SubNodes.selectSingleNode("n:GiftReceivingStatus") is Nothing Then
								GiftReceivingStatus = ""
							Else
								GiftReceivingStatus = SubNodes.getElementsByTagName("n:GiftReceivingStatus")(0).Text			'선물 수신 상태 코드
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
							strSql = strSql & " 	('"&getDivCD&"', '단순변심', 'nvstorefarm', '" & html2db(CStr(OutMallOrderSerial)) & "', '', '', '', '', '', "
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
<script>alert('<%= Left(oinputDT, 10) %>일 저장하였습니다');</script>
<script>window.close();</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
