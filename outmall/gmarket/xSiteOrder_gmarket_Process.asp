<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 XML 주문처리
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteOrderXMLCls.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/outmall/gmarket/gmarketItemcls.asp"-->
<!-- #include virtual="/outmall/gmarket/incGmarketFunction.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<%
Function fnGmarketConfirmOrder(vOrderserial, vOrgDetailKey)
	Dim objXML, xmlDOM, iRbody, strSql, iResult, resComment
	Dim strRst, POS1, POS2, POS3
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst & "	<soap:Header>"
	strRst = strRst & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "			<encTicket>"&gmarketTicket&"</encTicket>"
	strRst = strRst & "		</EncTicket>"
	strRst = strRst & "	</soap:Header>"
	strRst = strRst & "	<soap:Body>"
	strRst = strRst & "		<ConfirmOrder xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "			<ConfirmOrder ContrNo="""&vOrgDetailKey&""" SendPlanDate="""&DATE()+2&"""  />"
	strRst = strRst & "		</ConfirmOrder>"
	strRst = strRst & "	</soap:Body>"
	strRst = strRst & "</soap:Envelope>"
'	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketAPIURL&"/v1/OrderService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(strRst)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/ConfirmOrder"
		objXML.send(strRst)

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
'				response.write Replace(objXML.responseText,"soap:","")
'				response.end
			response.write "<textarea cols=40 rows=10>"&Replace(objXML.responseText,"soap:","")&"</textarea>"

			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "ConfirmOrderResponse" Then
				iResult = xmlDOM.getElementsByTagName("ConfirmOrderResult ").item(0).getAttribute("Result")
				If iResult = "Success" Then
					strSql = ""
					strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] SET "
					strSql = strSql & " isbaljuConfirmSend = 'Y' "
					strSql = strSql & " , lastUpdate = getdate() "
					strSql = strSql & " WHERE outmallorderserial = '"&vOrderserial&"'  "
					strSql = strSql & " and orgDetailKey = '"&vOrgDetailKey&"' "
					strSql = strSql & " and mallid = 'gmarket1010' "
					dbget.Execute strSql
					fnGmarketConfirmOrder= true
				ElseIf iResult = "Change" Then
					resComment = xmlDOM.getElementsByTagName("ConfirmOrderResult ").item(0).getAttribute("Comment")
					Dim chgName, chgPhone1, chgPhone2, chgZipcode, chgAddress, finAddr1, finAddr2
					chgName		= ""
					chgPhone1	= ""
					chgPhone2	= ""
					chgZipcode	= ""
					chgAddress	= ""

					chgName		= Split(resComment, "^|^")(0)
					chgPhone1	= Split(resComment, "^|^")(1)
					chgPhone2	= Split(resComment, "^|^")(2)
					chgZipcode	= Split(resComment, "^|^")(3)
					chgAddress	= Split(resComment, "^|^")(4)

					POS1 = 0
					POS2 = 0
					POS3 = 0
					POS1 = InStr(chgAddress," ")
					''rw "POS1="&POS1
					IF (POS1>0) then
						POS2 = InStr(MID(chgAddress,POS1+1,512)," ")
						''rw "POS2="&POS2
						IF POS2>0 then
							POS3 = InStr(MID(chgAddress,POS1+POS2+1,512)," ")
							IF POS3>0 then
								finAddr2 = MID(chgAddress,POS1+POS2+POS3+1,512)
								finAddr1 = LEFT(chgAddress, POS1 + POS2 + POS3 - 1)
							END IF
						END IF
					END IF

					strSql = ""
					strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] SET "
					strSql = strSql & " isbaljuConfirmSend = 'C' "
					strSql = strSql & " , lastUpdate = getdate() "
					strSql = strSql & " , bigo = '"& resComment &"' "
					strSql = strSql & " WHERE outmallorderserial = '"&vOrderserial&"'  "
					strSql = strSql & " and orgDetailKey = '"&vOrgDetailKey&"' "
					strSql = strSql & " and mallid = 'gmarket1010' "
					dbget.Execute strSql

					strSql = ""
					strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMPOrder] "
					strSql = strSql & " SET ReceiveName = '"& chgName &"' "
					strSql = strSql & " , ReceiveTelNo ='"& chgPhone1 &"' "
					strSql = strSql & " , ReceiveHpNo ='"& chgPhone2 &"' "
					strSql = strSql & " , ReceiveZipCode ='"& chgZipcode &"' "
					strSql = strSql & " , ReceiveAddr1 ='"& finAddr1 &"' "
					strSql = strSql & " , ReceiveAddr2 ='"& finAddr2 &"' "
					strSql = strSql & " WHERE outmallorderserial = '"&vOrderserial&"'  "
					strSql = strSql & " and orgDetailKey = '"&vOrgDetailKey&"' "
					strSql = strSql & " and SellSite = 'gmarket1010' "
					dbget.Execute strSql
				End If
			End If
		Set xmlDOM = nothing
	Set objXML = nothing
'	On Error Goto 0
End Function

Function getGmarketOrderXMLStr()
	Dim strRst
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst & "	<soap:Header>"
	strRst = strRst & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "			<encTicket>"&gmarketTicket&"</encTicket>"
	strRst = strRst & "		</EncTicket>"
	strRst = strRst & "	</soap:Header>"
	strRst = strRst & "	<soap:Body>"
	strRst = strRst & "		<RequestOrder xmlns=""http://tpl.gmarket.co.kr/"" />"
	strRst = strRst & "	</soap:Body>"
	strRst = strRst & "</soap:Envelope>"
	getGmarketOrderXMLStr = strRst
End function

function saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
        , OrderName, OrderTelNo, OrderHpNo _
        , ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
        , SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
        , deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
        , byref ierrCode, byref ierrStr )
    dim paramInfo, retParamInfo
    dim PayType  : PayType  = "50"
    dim sqlStr
	dim countryCode

	if countryCode="" then countryCode="KR"

    saveOrderOneToTmpTable =false

	if isNULL(OrderTelNo) then OrderTelNo=""
	if isNULL(OrderHpNo) then OrderHpNo=""
	if isNULL(ReceiveTelNo) then ReceiveTelNo=""
	if isNULL(ReceiveHpNo) then ReceiveHpNo=""

    OrderTelNo = replace(OrderTelNo,")","-")
    OrderHpNo = replace(OrderHpNo,")","-")
    ReceiveTelNo = replace(ReceiveTelNo,")","-")
    ReceiveHpNo = replace(ReceiveHpNo,")","-")

    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
        ,Array("@SellSite" , adVarchar	, adParamInput, 32, SellSite)	_
		,Array("@OutMallOrderSerial"	, adVarchar	, adParamInput,32, OutMallOrderSerial)	_
		,Array("@SellDate"	,adDate, adParamInput,, SellDate) _
		,Array("@PayType"	,adVarchar, adParamInput,32, PayType) _
		,Array("@Paydate"	,adDate, adParamInput,, SellDate) _
		,Array("@matchItemID"	,adInteger, adParamInput,, matchItemID) _
		,Array("@matchItemOption"	,adVarchar, adParamInput,4, matchItemOption) _
		,Array("@partnerItemID"	,adVarchar, adParamInput,32, matchItemID) _
		,Array("@partnerItemName"	,adVarchar, adParamInput,128, partnerItemName) _
		,Array("@partnerOption"	,adVarchar, adParamInput,128, matchItemOption) _
		,Array("@partnerOptionName"	,adVarchar, adParamInput,128, partnerOptionName) _
		,Array("@outMallGoodsNo"	,adVarchar, adParamInput,16, outMallGoodsNo) _
		,Array("@OrderUserID"	,adVarchar, adParamInput,32, "") _
		,Array("@OrderName"	,adVarchar, adParamInput,32, OrderName) _
		,Array("@OrderEmail"	,adVarchar, adParamInput,100, "") _
		,Array("@OrderTelNo"	,adVarchar, adParamInput,16, OrderTelNo) _
		,Array("@OrderHpNo"	,adVarchar, adParamInput,16, OrderHpNo) _
		,Array("@ReceiveName"	,adVarchar, adParamInput,32, ReceiveName) _
		,Array("@ReceiveTelNo"	,adVarchar, adParamInput,16, ReceiveTelNo) _
		,Array("@ReceiveHpNo"	,adVarchar, adParamInput,16, ReceiveHpNo) _
		,Array("@ReceiveZipCode"	,adVarchar, adParamInput,7, ReceiveZipCode) _
		,Array("@ReceiveAddr1"	,adVarchar, adParamInput,128, ReceiveAddr1) _
		,Array("@ReceiveAddr2"	,adVarchar, adParamInput,512, ReceiveAddr2) _
		,Array("@SellPrice"	,adCurrency, adParamInput,, SellPrice) _
		,Array("@RealSellPrice"	,adCurrency, adParamInput,, RealSellPrice) _
		,Array("@ItemOrderCount"	,adInteger, adParamInput,, ItemOrderCount) _
		,Array("@OrgDetailKey"	,adVarchar, adParamInput,32, OrgDetailKey) _
		,Array("@DeliveryType"	,adInteger, adParamInput,, 0) _
		,Array("@deliveryprice"	,adCurrency, adParamInput,, 0) _
		,Array("@deliverymemo"	,adVarchar, adParamInput,400, deliverymemo) _
		,Array("@requireDetail"	,adVarchar, adParamInput,400, requireDetail) _
		,Array("@orderDlvPay"	,adCurrency, adParamInput,, orderDlvPay) _
		,Array("@orderCsGbn"	,adInteger, adParamInput,, orderCsGbn) _
    	,Array("@countryCode"	,adVarchar, adParamInput,2, countryCode) _
		,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
	)

    if (matchItemOption<>"") and (matchItemID<>"-1") and (matchItemID<>"") then
        sqlStr = "db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert_FromXML"
        retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

        ierrCode = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
        ierrStr  = GetValue(retParamInfo, "@retErrStr")   ' 에러메세지
    else
        ierrCode = -999
        ierrStr = "상품코드 또는 옵션코드  매칭 실패" & OrgDetailKey & " 상품코드 =" & matchItemID&" 옵션명 = "&partnerOptionName
        rw "["&ierrCode&"]"&ierrStr
        dbget.close() : response.end
    end if

    saveOrderOneToTmpTable = (ierrCode=0)
    if (ierrCode<>0) then
        rw "["&ierrCode&"]"&ierrStr
    end if
end function

Dim sqlStr, buf, i, mode, sellsite
Dim divcd, yyyymmdd, idx, Nodes, Nodes2, SubNodes, SubNodes2
Dim objXML, xmlDOM, retCode, iMessage
Dim xmlStr : xmlStr = getGmarketOrderXMLStr()
mode		= requestCheckVar(html2db(request("mode")),32)
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
idx			= requestCheckVar(html2db(request("idx")),32)

Dim tmpxml, strsql, retVal, deliverymemo, orderCsGbn, errCode, errStr, succCNT, failCNT
Dim OrgDetailKey, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, OrderName, OrderTelNo, OrderHpNo
Dim OutMallOrderSerial, SellDate, outMallGoodsNo, matchItemID, partnerItemName, SellPrice, RealSellPrice, ItemOrderCount, PaymentPrice, orderDlvPay, requireDetail, matchItemOption, OptionInfo, tmpOptionVal
Dim partnerOptionName, SalePrice, AddPrice, CouponDiscountsPrice, AddDiscountsPrice
Dim regOrderCnt

Dim LagrgeNode, MiddleNode, j
If sellsite = "gmarket1010" Then
	'On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & gmarketSSLAPIURL&"/v1/OrderService.asmx"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(xmlStr)
		objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/RequestOrder"
		objXML.send(xmlStr)

		If Err.number <> 0 Then
			iErrStr = "ERR||" & Err.Description
		End If

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
				If session("ssBctID")="kjy8517" Then
					response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"
				End If

				Set LagrgeNode = xmlDOM.SelectNodes("//RequestOrderResultT")
				''response.write LagrgeNode.length & "aaaa<br />"
					If Not (LagrgeNode Is Nothing) Then
						For i = 0 To LagrgeNode.length - 1
							If LagrgeNode(i).getAttribute("IsGiftOrder") = "N" OR ( (LagrgeNode(i).getAttribute("IsGiftOrder") = "Y") AND (LagrgeNode(i).getAttribute("GiftOrderStatus") = "2") ) Then
								orderCsGbn			= 0
								OutMallOrderSerial	= Trim(LagrgeNode(i).getAttribute("PackNo"))
								ReceiveName			= LEFT(Trim(LagrgeNode(i).getAttribute("ReceiverName")), 28)
								ReceiveTelNo		= Trim(LagrgeNode(i).getAttribute("ReceiverPhone1"))
								ReceiveHpNo			= Trim(LagrgeNode(i).getAttribute("ReceiverPhone2"))
								ReceiveZipCode		= Trim(LagrgeNode(i).getAttribute("ReceiverZipcode"))
								ReceiveAddr1		= Trim(LagrgeNode(i).getAttribute("ReceiverAddress1"))
								ReceiveAddr2		= Trim(LagrgeNode(i).getAttribute("ReceiverAddress2"))
								OrderName			= LEFT(Trim(LagrgeNode(i).getAttribute("BuyerName")), 28)
								OrderTelNo			= Trim(LagrgeNode(i).getAttribute("BuyerPhone1"))
								OrderHpNo			= Trim(LagrgeNode(i).getAttribute("BuyerPhone2"))
								OrgDetailKey		= Trim(LagrgeNode(i).getAttribute("ContrNo"))
								SellDate			= Trim(Replace(LagrgeNode(i).getAttribute("ContrDate"), "T", " "))
								outMallGoodsNo		= Trim(LagrgeNode(i).getAttribute("GmktItemNo"))
								matchItemID			= Trim(LagrgeNode(i).getAttribute("OutItemNo"))
								partnerItemName		= Trim(LagrgeNode(i).getAttribute("ItemName"))
								ItemOrderCount		= Trim(LagrgeNode(i).getAttribute("Quantity"))
								SalePrice			= Clng(Trim(LagrgeNode(i).getAttribute("SalePrice")))
								AddPrice			= Clng(Trim(LagrgeNode(i).getAttribute("ItemOptionSelectionPrice")))
								AddPrice			= AddPrice / ItemOrderCount
								CouponDiscountsPrice= Clng(Trim(LagrgeNode(i).getAttribute("CouponDiscountsPrice")))
								AddDiscountsPrice	= Clng(Trim(LagrgeNode(i).getAttribute("AddDiscountsPrice")))
								'RealSellPrice		= Clng(Trim(LagrgeNode(i).getAttribute("SalePrice"))) + Clng(Trim(LagrgeNode(i).getAttribute("ItemOptionSelectionPrice")))
								SellPrice			= SalePrice + AddPrice		'판매가 = 판매금액 + 추가금액
								RealSellPrice		= SellPrice - Clng((CouponDiscountsPrice + AddDiscountsPrice) / ItemOrderCount) '실판매가 = 판매가 - ((쿠폰할인가 + 추가할인가) / 수량) | 2017-08-01 14:48 김진영 수정..할인
								PaymentPrice		= Clng(Trim(LagrgeNode(i).getAttribute("PaymentPrice")))
								orderDlvPay			= Clng(Trim(LagrgeNode(i).getAttribute("ShippingFee")))
								deliverymemo		= Trim(LagrgeNode(i).getAttribute("BuyerMemo"))
								OptionInfo			= Trim(LagrgeNode(i).getAttribute("OptionInfo"))

								If InStr(OptionInfo, "텍스트를 입력하세요") > 0 Then
									requireDetail	= Trim(Split(OptionInfo, "텍스트를 입력하세요;")(1))
									If Right(requireDetail,1) = "," Then
										requireDetail = Left(requireDetail, Len(requireDetail) - 1)
									End If
								Else
									requireDetail	= ""
								End If

								matchItemOption = ""
								partnerOptionName = ""
								tmpOptionVal = ""
								Set MiddleNode = LagrgeNode(i).SelectNodes("./ItemOptionSelect")
									If Not (MiddleNode Is Nothing) Then
										For j = 0 To MiddleNode.length - 1
											matchItemOption		= Trim(MiddleNode(j).getAttribute("ItemOptionCode"))
											tmpOptionVal		= Trim(MiddleNode(j).getAttribute("ItemOptionValue"))
											If Instr(tmpOptionVal, "^|^") > 0 Then
												partnerOptionName	= Trim(Replace(Split(tmpOptionVal, ";")(1), "^|^", ","))
											Else
												partnerOptionName	= Trim(Split(tmpOptionVal, ";")(1))
											End If
										Next
									End If
								Set MiddleNode = nothing
								If matchItemOption = "" Then matchItemOption ="0000"

	'							rw "OrgDetailKey : " & OrgDetailKey
	'							rw "ReceiveName : " & ReceiveName
	'							rw "ReceiveTelNo : " & ReceiveTelNo
	'							rw "ReceiveHpNo : " & ReceiveHpNo
	'							rw "ReceiveZipCode : " & ReceiveZipCode
	'							rw "ReceiveAddr1 : " & ReceiveAddr1
	'							rw "ReceiveAddr2 : " & ReceiveAddr2
	'							rw "OrderName : " & OrderName
	'							rw "OrderTelNo : " & OrderTelNo
	'							rw "OrderHpNo : " & OrderHpNo
	'							rw "OutMallOrderSerial : " & OutMallOrderSerial
	'							rw "SellDate : " & SellDate
	'							rw "outMallGoodsNo : " & outMallGoodsNo
	'							rw "matchItemID : " & matchItemID
	'							rw "partnerItemName : " & partnerItemName
	'							rw "SellPrice : " & SellPrice
	'							rw "SalePrice: " & SalePrice
	'							rw "AddPrice : " & AddPrice
	'							rw "CouponDiscountsPrice : " & CouponDiscountsPrice
	'							rw "AddDiscountsPrice : " & AddDiscountsPrice
	'							rw "RealSellPrice : " & RealSellPrice
	'							rw "ItemOrderCount : " & ItemOrderCount
	'							rw "PaymentPrice : " & PaymentPrice
	'							rw "orderDlvPay : " & orderDlvPay
	'							rw "deliverymemo : " & deliverymemo
	'							rw "requireDetail : " & requireDetail
	'							rw "matchItemOption : " & matchItemOption
	'							rw "partnerOptionName : " & partnerOptionName
	'							rw "Real1 : " & SellPrice - CouponDiscountsPrice - AddDiscountsPrice '실판매가 = 판매가 - 쿠폰할인가 - 추가할인가
	'							rw "Real2 : " & SellPrice - Clng((CouponDiscountsPrice + AddDiscountsPrice) / ItemOrderCount)
	'							rw "---------------------------------"

								retVal= saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
										, OrderName, OrderTelNo, OrderHpNo _
										, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
										, SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
										, deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
										, errCode, errStr )

								'영문 이름 누락이라고 해서 한번 더 검색 후 추후 지켜보기..2017-12-29
								regOrderCnt = 0
								strsql = ""
								strsql = " SELECT COUNT(*) as cnt FROM db_temp.dbo.tbl_xsite_tmporder WHERE outmallorderserial = '"&OutMallOrderSerial&"' and OrgDetailKey = '"&OrgDetailKey&"' and sellsite = 'gmarket1010'  "
								rsget.CursorLocation = adUseClient
								rsget.Open strsql, dbget, adOpenForwardOnly, adLockReadOnly
								if not rsget.Eof then
									regOrderCnt = rsget("cnt")
								end if
								rsget.close

								If (retVal) and (regOrderCnt > 0) Then
									succCNT = succCNT + 1
									strsql = ""
									strsql = strsql & " INSERT INTO db_temp.[dbo].[tbl_xSite_TMP11stOrder] (outmallorderserial, OrgDetailKey, beasongNum11st, isbaljuConfirmSend, regdate, mallid) "
									strsql = strsql & " VALUES ('"&OutMallOrderSerial&"', '"&OrgDetailKey&"', '', 'N', getdate(), 'gmarket1010')"
									dbget.Execute strSql
								Else
									failCNT = failCNT + 1
								End If
							Else
							End If
						Next
					End If
				Set LagrgeNode = nothing

				If (failCNT <> 0) Then
				    rw "["&failCNT&"] 건 실패(주문조회)"
				End if

				If (succCNT <> 0) then
				    rw "["&succCNT&"] 건 성공(주문조회)"
				    Dim arrList, lp, ret1
				    Dim OKcnt, NOcnt
				    OKcnt = 0
				    NOcnt = 0

					strsql = ""
					strsql = strsql & " update T "
					strsql = strsql & " set T.isbaljuConfirmSend='Y' "
					strsql = strsql & " From db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T "
					strsql = strsql & " JOIN db_temp.dbo.tbl_xsite_tmporder as O on T.outmallorderserial = O.OutMallOrderSerial and T.OrgDetailKey = O.OrgDetailKey "
					strsql = strsql & " where T.isbaljuConfirmSend <> 'Y' "
					strsql = strsql & " and O.sendState = 1 "
					strsql = strsql & " and O.matchstate in ('O') "
					strsql = strsql & " and T.mallid = 'gmarket1010' "
					dbget.Execute strsql

					'' 주석처리 2018/11/06
					' strsql = ""
					' strsql = strsql & " update T "
					' strsql = strsql & " set T.isbaljuConfirmSend='Y' "
					' strsql = strsql & " FROM db_order.dbo.tbl_order_master as M "
					' strsql = strsql & " JOIN db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T on M.authcode = T.outmallorderserial "
					' strsql = strsql & " WHERE M.cancelyn ='Y' "
					' strsql = strsql & " and T.isbaljuConfirmSend <> 'Y' "
					' strsql = strsql & " and T.mallid = 'gmarket1010' "
					' dbget.Execute strsql

					strsql = ""
					strsql = strsql & " SELECT TOP 1000 outmallorderserial, OrgDetailKey FROM db_temp.[dbo].[tbl_xSite_TMP11stOrder] "
					strsql = strsql & " WHERE isbaljuConfirmSend = 'N' "
					strsql = strsql & " and mallid = 'gmarket1010' "
					strsql = strsql & " and regdate>dateadd(d,-5,getdate())"  ''최근것만.
					rsget.CursorLocation = adUseClient
					rsget.Open strsql, dbget, adOpenForwardOnly, adLockReadOnly
				    if not rsget.Eof then
				        arrList = rsget.getRows()
				    end if
				    rsget.close

					For lp = 0 To Ubound(arrList, 2)
						ret1 = fnGmarketConfirmOrder(arrList(0, lp), arrList(1, lp))

		                If (ret1) then
		                    OKcnt = OKcnt + 1
		                Else
		                    NOcnt = NOcnt + 1
		                End If
					Next

					If OKcnt <> 0 then
						rw "["&OKcnt&"] 건 성공(발주확인)"
					End If

					If NOcnt <> 0 then
						rw "["&NOcnt&"] 건 실패(발주확인)"
					End If

				End If
			Else
				rw "주문연동 실패..잠시 후 시도 요망"
			End If
		'On Error Goto 0
		Set objXML = nothing
End If

''품절/가격 오류체크
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->