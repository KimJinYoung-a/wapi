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
<!-- #include virtual="/outmall/halfclub/halfclubItemcls.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/halfclub/incHalfclubFunction.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<%
public function getOptionCodByOptionNameOutmall(iitemid, ioptionname)
    dim retStr, sqlStr : retStr=""

    sqlStr = "select top 1 itemoption "
    sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
    sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
    sqlStr = sqlStr & " and optionname = '"&html2db(ioptionname)&"' "
'	response.write sqlstr & "<Br>"
'response.end
    rsget.Open sqlStr,dbget,1
    if (Not rsget.EOF) then
	    retStr = rsget("itemoption")
	end if
    rsget.Close

    If (retStr="") THEN
       ''옵션 매칭이 안되었을때. 수기매칭으로 진행
        sqlStr = "select count(*) as CNT "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
        rsget.Open sqlStr,dbget,1
        if (Not rsget.EOF) then
    	    if (rsget("CNT")>=0) THEN retStr = "0000"
    	end if
        rsget.Close

    END IF
    getOptionCodByOptionNameOutmall = retStr

	if retStr="" then
	    rw sqlStr
	end if
end function

Function fnHalfClubConfirmOrder(vOutmallOrderserial, vOrgDetailKey)
	Dim objXML, xmlDOM, iRbody, strSql, iResult
	Dim strRst
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst & "	<soap:Header>"
	strRst = strRst & "		<SOAPHeaderAuth xmlns=""http://api.tricycle.co.kr/"">"
	strRst = strRst & " 		<User_ID>"&UPCHECODE&"</User_ID>"
	strRst = strRst & " 		<User_PWD>"&APIKEY&"</User_PWD>"
	strRst = strRst & "		</SOAPHeaderAuth>"
	strRst = strRst & "	</soap:Header>"
	strRst = strRst & "	<soap:Body>"
	strRst = strRst & "		<Set_OrderConfirm xmlns=""http://api.tricycle.co.kr/"">"
	strRst = strRst & "			<req_orderConfirm>"
	strRst = strRst & "				<OrdNum>"&vOutmallOrderserial&"</OrdNum>"									'#하프클럽 주문 번호
	strRst = strRst & "				<OrdNum_Nm>"&vOrgDetailKey&"</OrdNum_Nm>"									'#하프클럽 주문 순번
	strRst = strRst & "				<ConfirmYMD>"&FormatDate(now(), "0000-00-00 00:00:00")&"</ConfirmYMD>"		'#제휴사 주문 확인일(yyyy-MM-dd HH:mm:ss)
	strRst = strRst & "			</req_orderConfirm>"
	strRst = strRst & "		</Set_OrderConfirm>"
	strRst = strRst & "	</soap:Body>"
	strRst = strRst & "</soap:Envelope>"
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objXML.open "POST", "" & APIURL&"/Order/Order.asmx"
		objXML.setRequestHeader "Host", "api.tricycle.co.kr"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(xmlStr)
		objXML.setRequestHeader "SOAPMethodName", "Set_OrderConfirm"
		objXML.send(strRst)
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
'			response.write replace(objXML.responseText, "utf-8","euc-kr")
'			response.end
			iResult = xmlDOM.getElementsByTagName("ResultCode").Item(0).Text
			If iResult = "0000" Then
				strSql = ""
				strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] SET "
				strSql = strSql & " isbaljuConfirmSend = 'Y' "
				strSql = strSql & " , lastUpdate = getdate() "
				strSql = strSql & " WHERE outmallorderserial = '"&vOutmallOrderserial&"'  "
				strSql = strSql & " and orgDetailKey = '"&vOrgDetailKey&"' "
				strSql = strSql & " and mallid = 'halfclub' "
				dbget.Execute strSql
				fnHalfClubConfirmOrder= true
			End If
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

Function getHalfclubOrderXMLStr(iFromDate, iNowDate)
	Dim strRst
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
	strRst = strRst & "	<soap12:Header>"
	strRst = strRst & "		<SOAPHeaderAuth xmlns=""http://api.tricycle.co.kr/"">"
	strRst = strRst & " 		<User_ID>"&UPCHECODE&"</User_ID>"
	strRst = strRst & " 		<User_PWD>"&APIKEY&"</User_PWD>"
	strRst = strRst & "		</SOAPHeaderAuth>"
	strRst = strRst & "	</soap12:Header>"
	strRst = strRst & "	<soap12:Body>"
	strRst = strRst & "		<Get_OrderInfo xmlns=""http://api.tricycle.co.kr/"">"
	strRst = strRst & "			<req_OrderInfo>"
	strRst = strRst & "				<FromYMD>"&iFromDate&"</FromYMD>"		'#주문 데이터 시작일 | ex : 201002030101
	strRst = strRst & "				<ToYMD>"&iNowDate&"</ToYMD>"			'#주문 데이터 종료일 | ex : 201002030101
	strRst = strRst & "				<NewOrder>Y</NewOrder>"					'#결제구분 | Y : 결제완료, N : 출고 대기
	strRst = strRst & "			</req_OrderInfo>"
	strRst = strRst & "		</Get_OrderInfo>"
	strRst = strRst & "	</soap12:Body>"
	strRst = strRst & "</soap12:Envelope>"
	getHalfclubOrderXMLStr = strRst
End function

function saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption, prdStckNo, partnerItemName,partnerOptionName,outMallGoodsNo _
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
		,Array("@partnerOption"	,adVarchar, adParamInput,128, prdStckNo) _
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
'rw "ierrCode : " & ierrCode
    saveOrderOneToTmpTable = (ierrCode=0)
    if (ierrCode<>0) then
        rw "["&ierrCode&"]"&ierrStr
    end if
end function

Dim sqlStr, buf, i, mode, sellsite
Dim divcd, yyyymmdd, idx, Nodes, Nodes2, SubNodes, SubNodes2
Dim objXML, xmlDOM, retCode, iMessage, testXML

Dim iSellDate, iIsSuccess, fromDate, nowDate
Call GetCheckStatus("halfclub", iSellDate, iIsSuccess)
fromDate = Replace(iSellDate, "-", "") & "0000"
nowDate = Replace(DATE() + 1, "-", "") & "0000"


Dim xmlStr : xmlStr = getHalfclubOrderXMLStr(fromDate, nowDate)
mode		= requestCheckVar(html2db(request("mode")),32)
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
idx			= requestCheckVar(html2db(request("idx")),32)

Dim tmpxml, strsql, retVal, deliverymemo, orderCsGbn, errCode, errStr, succCNT, failCNT
Dim OrgDetailKey, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, OrderName, OrderTelNo, OrderHpNo, prdStckNo, ToAddr, SumDeliPri, SumPrdQPri, PayPri, SumQty, paymentYmd, Email, EngNm, Jumin, IsGlobal, Etc, SalePst
Dim OutMallOrderSerial, SellDate, outMallGoodsNo, matchItemID, partnerItemName, SellPrice, RealSellPrice, ItemOrderCount, PaymentPrice, orderDlvPay, requireDetail, matchItemOption, OptionInfo, tmpOptionVal
Dim partnerOptionName, SalePrice, AddPrice, CouponDiscountsPrice, AddDiscountsPrice
Dim regOrderCnt

Dim LagrgeNode, MiddleNode, j
If sellsite = "halfclub" Then
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objXML.open "POST", "" & APIURL&"/Order/Order.asmx"
		objXML.setRequestHeader "Host", "api.tricycle.co.kr"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(xmlStr)
		objXML.setRequestHeader "SOAPMethodName", "Get_OrderInfo"
		objXML.send(xmlStr)
		If Err.number <> 0 Then
			iErrStr = "ERR||" & Err.Description
		End If
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
				'response.write replace(objXML.responseText, "utf-8","euc-kr")
				'response.end
				'Test 버전
				'xmlDOM.loadXML(Replace(testXML,"soap:",""))
				'response.write replace(testXML, "utf-8","euc-kr")
'				response.end
				Set LagrgeNode = xmlDOM.SelectNodes("//Request_Order")
					If Not (LagrgeNode Is Nothing) Then
						For i = 0 To LagrgeNode.length - 1
							orderCsGbn			= 0
							OutMallOrderSerial	= Trim(LagrgeNode(i).getElementsByTagName("POrdNum").item(0).text)		'주문번호
							ToAddr				= Trim(LagrgeNode(i).getElementsByTagName("ToAddr").item(0).text)		'?수취인 주소
							ReceiveAddr1		= Split(ToAddr, "_")(0)
							ReceiveAddr2		= Split(ToAddr, "_")(1)
							ReceiveZipCode		= Trim(LagrgeNode(i).getElementsByTagName("ToZiCd").item(0).text)		'수취인 우편번호
							ReceiveName			= Trim(LagrgeNode(i).getElementsByTagName("ToNm").item(0).text)			'수취인
							ReceiveTelNo		= Trim(LagrgeNode(i).getElementsByTagName("ToTel").item(0).text)		'수취인 전화번호
							ReceiveHpNo			= Trim(LagrgeNode(i).getElementsByTagName("ToEmTel").item(0).text)		'수취인 핸드폰번호
							deliverymemo		= Trim(LagrgeNode(i).getElementsByTagName("Memo").item(0).text)			'배송 메모
							OrderTelNo			= Trim(LagrgeNode(i).getElementsByTagName("GuestTel").item(0).text)		'주문자 전화번호
							OrderHpNo			= OrderTelNo
							OrderName			= Trim(LagrgeNode(i).getElementsByTagName("FromNm").item(0).text)		'주문자명
							SumDeliPri			= Trim(LagrgeNode(i).getElementsByTagName("SumDeliPri").item(0).text)	'총배송비 금액
							SumPrdQPri			= Trim(LagrgeNode(i).getElementsByTagName("SumPrdQPri").item(0).text)	'총 상품 금액sum(수량 * 판매가)
							PayPri				= Trim(LagrgeNode(i).getElementsByTagName("PayPri").item(0).text)		'결제금액
							SumQty				= Trim(LagrgeNode(i).getElementsByTagName("SumQty").item(0).text)		'상품 수량 sum(수량)
							SellDate			= Replace(Trim(LagrgeNode(i).getElementsByTagName("OrderYmd").item(0).text), "T", " ")		'주문 번호 생성일
							paymentYmd			= Trim(LagrgeNode(i).getElementsByTagName("paymentYmd").item(0).text)
							Email				= Trim(LagrgeNode(i).getElementsByTagName("Email").item(0).text)
							EngNm				= Trim(LagrgeNode(i).getElementsByTagName("EngNm").item(0).text)
							Jumin				= Trim(LagrgeNode(i).getElementsByTagName("Jumin").item(0).text)
							IsGlobal			= Trim(LagrgeNode(i).getElementsByTagName("IsGlobal").item(0).text)		'해외배송 주문여부(Y / N)
							Set MiddleNode = LagrgeNode(i).SelectNodes("./OrderDetailInfo/Request_OrderDetail")
								If Not (MiddleNode Is Nothing) Then

									For j = 0 To MiddleNode.length - 1
										OrgDetailKey	= Trim(MiddleNode(j).getElementsByTagName("OrdNum_Nm").item(0).text)		'주문 순번
										outMallGoodsNo	= Trim(MiddleNode(j).getElementsByTagName("PCode").item(0).text)			'상품코드
										matchItemID		= Split(outMallGoodsNo, "_")(0)
										prdStckNo		= Trim(MiddleNode(j).getElementsByTagName("OptCd").item(0).text)			'옵션코드
										partnerOptionName	= Trim(MiddleNode(j).getElementsByTagName("OptNm").item(0).text)		'옵션명
										ItemOrderCount	= Trim(MiddleNode(j).getElementsByTagName("Qty").item(0).text)				'수량
										SellPrice		= Trim(MiddleNode(j).getElementsByTagName("SalPri").item(0).text)			'상품 판매가(단품)
										orderDlvPay		= Trim(MiddleNode(j).getElementsByTagName("DeliPri").item(0).text)			'배송비(단품에 대한 배송비)
										RealSellPrice	= Trim(MiddleNode(j).getElementsByTagName("DPayPri").item(0).text)			'쿠폰비를 제외한 상품 판매가
										Etc				= Trim(MiddleNode(j).getElementsByTagName("Etc").item(0).text)
										SalePst			= Trim(MiddleNode(j).getElementsByTagName("SalePst").item(0).text)
										matchItemOption = getOptionCodByOptionNameOutmall(matchItemID, partnerOptionName)

										retVal= saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption, prdStckNo, partnerItemName,partnerOptionName,outMallGoodsNo _
												, OrderName, OrderTelNo, OrderHpNo _
												, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
												, SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
												, deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
												, errCode, errStr )

										If (retVal) Then
											succCNT = succCNT + 1
											strsql = ""
											strsql = strsql & " INSERT INTO db_temp.[dbo].[tbl_xSite_TMP11stOrder] (outmallorderserial, OrgDetailKey, beasongNum11st, isbaljuConfirmSend, regdate, mallid) "
											strsql = strsql & " VALUES ('"&OutMallOrderSerial&"', '"&OrgDetailKey&"', '', 'N', getdate(), 'halfclub')"
											dbget.Execute strSql
										Else
											failCNT = failCNT + 1
										End If

'							rw "orderCsGbn : " & orderCsGbn
'							rw "OutMallOrderSerial : " & OutMallOrderSerial
'							rw "ToAddr : " & ToAddr
'							rw "ReceiveAddr1 : " & ReceiveAddr1
'							rw "ReceiveAddr2 : " & ReceiveAddr2
'							rw "ReceiveZipCode : " & ReceiveZipCode
'							rw "ReceiveName : " & ReceiveName
'							rw "ReceiveTelNo : " & ReceiveTelNo
'							rw "ReceiveHpNo : " & ReceiveHpNo
'							rw "deliverymemo : " & deliverymemo
'							rw "OrderTelNo : " & OrderTelNo
'							rw "OrderHpNo : " & OrderHpNo
'							rw "OrderName : " & OrderName
'							rw "SumDeliPri : " & SumDeliPri
'							rw "SumPrdQPri : " & SumPrdQPri
'							rw "PayPri : " & PayPri
'							rw "SumQty : " & SumQty
'							rw "SellDate : " & SellDate
'							rw "paymentYmd : " & paymentYmd
'							rw "Email : " & Email
'							rw "EngNm : " & EngNm
'							rw "Jumin : " & Jumin
'							rw "IsGlobal : " & IsGlobal
'							rw "OrgDetailKey : " & OrgDetailKey
'							rw "outMallGoodsNo : " & outMallGoodsNo
'							rw "matchItemID : " & matchItemID
'							rw "matchItemOption : " & matchItemOption
'							rw "prdStckNo : " & prdStckNo
'							rw "partnerOptionName : " & partnerOptionName
'							rw "ItemOrderCount : " & ItemOrderCount
'							rw "SellPrice : " & SellPrice
'							rw "orderDlvPay : " & orderDlvPay
'							rw "RealSellPrice : " & RealSellPrice
'							rw "Etc : " & Etc
'							rw "SalePst : " & SalePst
'							rw "------------------------------------------------------------------------------------------------"
									Next
								End If
							Set MiddleNode = nothing
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
					strsql = strsql & " and T.mallid = 'halfclub' "
					dbget.Execute strsql

					strsql = ""
					strsql = strsql & " update T "
					strsql = strsql & " set T.isbaljuConfirmSend='Y' "
					strsql = strsql & " FROM db_order.dbo.tbl_order_master as M "
					strsql = strsql & " JOIN db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T on M.authcode = T.outmallorderserial "
					strsql = strsql & " WHERE M.cancelyn ='Y' "
					strsql = strsql & " and T.isbaljuConfirmSend <> 'Y' "
					strsql = strsql & " and T.mallid = 'halfclub' "
					dbget.Execute strsql

					strsql = ""
					strsql = strsql & " SELECT TOP 1000 outmallorderserial, OrgDetailKey, beasongNum11st FROM db_temp.[dbo].[tbl_xSite_TMP11stOrder] "
					strsql = strsql & " WHERE isbaljuConfirmSend = 'N' "
					strsql = strsql & " and mallid = 'halfclub' "
				    rsget.Open strsql,dbget,1
				    if not rsget.Eof then
				        arrList = rsget.getRows()
				    end if
				    rsget.close

					For lp = 0 To Ubound(arrList, 2)
						ret1 = fnHalfClubConfirmOrder(arrList(0, lp), arrList(1, lp))

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

					If (iSelldate < Left(Now(), 10)) then
						Call SetCheckStatus(sellsite, Left(DateAdd("d", 1, CDate(iSellDate)), 10), "N")
					ElseIf (iSellDate = Left(Now(), 10)) then
						Call SetCheckStatus(sellsite, iSellDate, "Y")
					End If
				End If
		Else
			rw "주문연동 실패..잠시 후 시도 요망"
		End If
	Set objXML = nothing
	On Error Goto 0
End If

''품절/가격 오류체크
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->