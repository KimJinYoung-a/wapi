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
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/wmp/wmpItemcls.asp"-->
<!-- #include virtual="/outmall/wmp/incWmpFunction.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Function fnWMPConfirmOrder(vOrderserial)
	Dim objXML, xmlDOM, iRbody, strSql, istrParam, isSuccessCode, strObj
	'istrParam = "bundleNo="&vOrderserial
	istrParam = "bundleNo="&vOrderserial
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info") = "Dev" Then
			objXML.open "POST", "http://localhost:62569/Wemake/Orders/ordercomplete", false
		Else
			objXML.open "POST", "http://110.93.128.100:8090/wemake/Orders/ordercomplete", false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If Err.number <> 0 Then
			iErrStr = ivendorItemId
			Exit Function
		End If
		rw objXML.Status
		rw BinaryToText(objXML.ResponseBody,"utf-8")

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
			Set strObj = JSON.parse(iRbody)
				isSuccessCode		= strObj.code
				If isSuccessCode = "200" Then
					strSql = ""
					strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] SET "
					strSql = strSql & " isbaljuConfirmSend = 'Y' "
					strSql = strSql & " , lastUpdate = getdate() "
					strSql = strSql & " WHERE outmallorderserial = '"&vOrderserial&"'  "
					strSql = strSql & " and mallid = 'WMP' "
					dbget.Execute strSql
					fnWMPConfirmOrder= true
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

function saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
        , OrderName, OrderTelNo, OrderHpNo _
        , ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
        , SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
        , deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
        , byref ierrCode, byref ierrStr, beasongNum11st, splitbeasongYn, outMallOptionNo, reserve01)
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
		,Array("@reserve01"	,adVarchar, adParamInput,32, reserve01) _
	)

    if (matchItemOption<>"") and (matchItemID<>"-1") and (matchItemID<>"") then
        sqlStr = "db_temp.[dbo].[usp_API_WMP_OrderReg_Add]"
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

Dim sqlStr, buf, i, j, k, mode, sellsite
Dim divcd, yyyymmdd, idx, Nodes, Nodes2, SubNodes, SubNodes2
Dim objXML, xmlDOM, retCode, iMessage, optionQty
mode		= requestCheckVar(html2db(request("mode")),32)
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
idx			= requestCheckVar(html2db(request("idx")),32)

Dim tmpxml, strsql, retVal, deliverymemo, orderCsGbn, errCode, errStr, succCNT, failCNT
Dim OrgDetailKey, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, OrderName, OrderTelNo, OrderHpNo
Dim OutMallOrderSerial, SellDate, outMallGoodsNo, matchItemID, partnerItemName, SellPrice, RealSellPrice, ItemOrderCount, PaymentPrice, orderDlvPay, requireDetail, matchItemOption, OptionInfo, tmpOptionVal, outMallOptionNo
Dim optionAddPrice, optionSalePrice, optionCommissionPrice
Dim partnerOptionName, SalePrice, AddPrice, CouponDiscountsPrice, DiscountsPrice, beasongNum11st, splitbeasongYn, reserve01
Dim regOrderCnt, strObj, iRbody
Dim LagrgeNode, MiddleNode, tmpSellPrice
Dim iSellDate, iIsSuccess, fromDate, nowDate, splitOptName
Call GetCheckStatus("WMP", iSellDate, iIsSuccess)
'rw iSellDate
If sellsite = "WMP" Then
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info") = "Dev" Then
			objXML.open "GET", "http://localhost:62569/Wemake/Orders/orderlist?reqdate="&iSellDate&"&type=NEW", false
		Else
			objXML.open "GET", "http://110.93.128.100:8090/wemake/Orders/orderlist?reqdate="&iSellDate&"&type=NEW", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||" & Err.Description
		End If
		'rw BinaryToText(objXML.ResponseBody,"utf-8")

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody

			Dim obj1, obj2, obj3, isSuccessCode
			Set strObj = JSON.parse(iRbody)
				isSuccessCode		= strObj.code
				iMessage			= strObj.message
				If isSuccessCode = "200" Then
					set obj1 = strObj.outPutValue.data.bundle
						for i=0 to obj1.length-1
							orderCsGbn			= 0
							OutMallOrderSerial	= obj1.get(i).bundleNo				'배송번호
							'obj1.get(i).purchaseNo									'구매번호
							SellDate			= obj1.get(i).orderDate				'주문일
							'obj1.get(i).payDate									'결제일
							'obj1.get(i).originShipDate								'발송기한
							'obj1.get(i).orderConfirmDate							'주문확인일
							'obj1.get(i).orderShippingDate							'배송시작일
							'obj1.get(i).shipCompleteDate							'배송완료일
							OrderName			= obj1.get(i).buyerName				'구매자명
							OrderHpNo			= obj1.get(i).buyerPhone			'구매자연락처
							OrderTelNo			= obj1.get(i).buyerPhone			'구매자연락처
							orderDlvPay			= obj1.get(i).shipPrice				'배송비
							'obj1.get(i).prepayment									'선불여부(선불,착불)
							'obj1.get(i).shipType									'배송유형 (무료, 유료, 조건부무료, 기타)

							'obj1.get(i).delivery.shipStatus						'배송상태 (신규주문 ,상품준비중, 배송중, 배송완료)
							'obj1.get(i).delivery.shipMethod						'배송방법 (일반-택배배송, 일반-직접배송, 일반-우편배송, 해외구매대행, 해외구매대행-택배배송, 해외구매대행-직접배송, 기타배송)
							'obj1.get(i).delivery.shipMethodMessage					'배송방법 메세지 배송방법이 [기타배송]일 경우 입력받는 메세지
							'obj1.get(i).delivery.scheduleShipDate					'배송예정일 배송방법이 [일반-직접배송, 해외구매대행-직접배송]일 경우 입력받는 배송예정일 (yyyy-MM-dd)
							'obj1.get(i).delivery.parcelCompany						'택배사 (주문 기초데이터 조회 - 택배사 조회 API 참고)
							'obj1.get(i).delivery.invoiceNo							'송장번호 (최대 20 자리)
							ReceiveName			= obj1.get(i).delivery.name			'받는사람
							ReceiveHpNo			= obj1.get(i).delivery.phone		'받는사람 연락처
							ReceiveTelNo		= obj1.get(i).delivery.phone		'받는사람 연락처
							'obj1.get(i).delivery.customsPin						'개인통관 고유부호

							ReceiveZipCode		= obj1.get(i).delivery.shipAddress.zipcode		'우편번호
							ReceiveAddr1		= obj1.get(i).delivery.shipAddress.addrFixed	'기본주소
							ReceiveAddr2		= obj1.get(i).delivery.shipAddress.addrDetail	'상세주소
							deliverymemo		= obj1.get(i).delivery.shipAddress.message		'배송메세지(최대 50자)
' rw "orderCsGbn : " & orderCsGbn
' rw "OutMallOrderSerial : " & OutMallOrderSerial
' rw "SellDate : " & SellDate
' rw "OrderName : " & OrderName
' rw "OrderHpNo : " & OrderHpNo
' rw "OrderTelNo : " & OrderTelNo
' rw "orderDlvPay : " & orderDlvPay
' rw "ReceiveName : " & ReceiveName
' rw "ReceiveHpNo : " & ReceiveHpNo
' rw "ReceiveTelNo : " & ReceiveTelNo
' rw "ReceiveZipCode : " & ReceiveZipCode
' rw "ReceiveAddr1 : " & ReceiveAddr1
' rw "ReceiveAddr2 : " & ReceiveAddr2
' rw "deliverymemo : " & deliverymemo
							set obj2 = obj1.get(i).orderProduct
								For j=0 to obj2.length-1
									SellPrice = ""
									tmpSellPrice = ""

									'OrgDetailKey		= obj2.get(j).orderNo			'주문번호
									reserve01			= obj2.get(j).orderNo			'주문번호
									outMallGoodsNo		= obj2.get(j).productNo			'상품번호
									partnerItemName		= obj2.get(j).productName		'상품명
									tmpSellPrice		= obj2.get(j).productPrice		'상품금액
'									RealSellPrice 		= SellPrice						'실판매가가 안 오는 듯..판매가 필드가 obj2.get(j).productPrice 이거 하나임
									ItemOrderCount		= obj2.get(j).productQty		'수량
									matchItemID			= obj2.get(j).sellerProductCode	'업체상품코드

									set obj3 = obj2.get(j).orderOption
										For k=0 to obj3.length-1
											optionAddPrice = ""
											RealSellPrice = ""
											SellPrice = ""

											OrgDetailKey			= obj3.get(k).orderOptionNo		'주문옵션번호
											outMallOptionNo			= obj3.get(k).optionNo			'옵션번호
											partnerOptionName		= obj3.get(k).optionName		'옵션
											ItemOrderCount			= obj3.get(k).optionQty			'수량
											matchItemOption			= obj3.get(k).sellerOptionCode	'업체옵션코드
											optionAddPrice			= obj3.get(k).optionAddPrice		'옵션별추가가격
											optionSalePrice			= obj3.get(k).optionSalePrice		'옵션별판매가격
											optionCommissionPrice	= obj3.get(k).optionCommissionPrice	'옵션별수수료

											SellPrice = CDbl(tmpSellPrice) + CDbl(optionAddPrice)
											RealSellPrice= SellPrice
											requireDetail = ""							'주문제작 문구 변수 초기화..2019-02-18 김진영 추가
											If Instr(obj3.get(k).optionName, "|") > 0 Then
												splitOptName = Split(obj3.get(k).optionName, "텍스트를 입력해주세요 :")
												If Ubound(splitOptName) > 0 Then
													partnerOptionName = Split(obj3.get(k).optionName, "|")(0)
													requireDetail = splitOptName(1)
												End If
											Else
												partnerOptionName	= obj3.get(k).optionName	'옵션
											End If

											retVal= saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
													, OrderName, OrderTelNo, OrderHpNo _
													, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
													, SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
													, deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
													, errCode, errStr, "", "", outMallOptionNo, reserve01 )
											If (retVal) Then
												succCNT = succCNT + 1
												strsql = ""
												strsql = strsql & " IF NOT EXISTS(SELECT * FROM db_temp.[dbo].[tbl_xSite_TMP11stOrder] WHERE outmallorderserial = '"&OutMallOrderSerial&"' and OrgDetailKey = '"& OrgDetailKey &"') "
												strSql = strSql & " 	BEGIN "
												strsql = strsql & " 		INSERT INTO db_temp.[dbo].[tbl_xSite_TMP11stOrder] (outmallorderserial, OrgDetailKey, beasongNum11st, isbaljuConfirmSend, regdate, mallid) "
												strsql = strsql & " 		VALUES ('"&OutMallOrderSerial&"', '"&OrgDetailKey&"', '', 'N', getdate(), 'WMP')"
												strSql = strSql & " 	END "
												dbget.Execute strSql
											Else
												failCNT = failCNT + 1
											End If
										Next
									set obj3 = nothing
								Next
							set obj2 = nothing
						Next
					set obj1 = nothing
				End If
			Set strObj = nothing

			If (failCNT > 0) Then
			    rw "["&failCNT&"] 건 실패(주문조회)"
			End if

			If (succCNT > 0) then
			    rw "["&succCNT&"] 건 성공(주문조회)"
			    Dim arrList, lp, ret1
			    Dim OKcnt, NOcnt
			    OKcnt = 0
			    NOcnt = 0

				strsql = ""
				strsql = strsql & " update T "
				strsql = strsql & " set T.isbaljuConfirmSend='Y' "
				strsql = strsql & " From db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T "
				strsql = strsql & " JOIN db_temp.dbo.tbl_xsite_tmporder as O on T.outmallorderserial = O.OutMallOrderSerial and T.OrgDetailKey = O.outMallOptionNo "
				strsql = strsql & " where T.isbaljuConfirmSend <> 'Y' "
				strsql = strsql & " and O.sendState = 1 "
				strsql = strsql & " and O.matchstate in ('O') "
				strsql = strsql & " and T.mallid = 'WMP' "
				dbget.Execute strsql

				strsql = ""
				strsql = strsql & " update T "
				strsql = strsql & " set T.isbaljuConfirmSend='Y' "
				strsql = strsql & " FROM db_order.dbo.tbl_order_master as M "
				strsql = strsql & " JOIN db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T on M.authcode = T.outmallorderserial "
				strsql = strsql & " WHERE M.cancelyn ='Y' "
				strsql = strsql & " and T.isbaljuConfirmSend <> 'Y' "
				strsql = strsql & " and T.mallid = 'WMP' "
				dbget.Execute strsql

				strsql = ""
				strsql = strsql & " SELECT TOP 1000 outmallorderserial FROM db_temp.[dbo].[tbl_xSite_TMP11stOrder] "
				strsql = strsql & " WHERE isbaljuConfirmSend = 'N' "
				strsql = strsql & " and mallid = 'WMP' "
				strsql = strsql & " GROUP BY outmallorderserial "
				rsget.CursorLocation = adUseClient
				rsget.Open strsql, dbget, adOpenForwardOnly, adLockReadOnly
			    if not rsget.Eof then
			        arrList = rsget.getRows()
			    end if
			    rsget.close

				For lp = 0 To Ubound(arrList, 2)
					ret1 = fnWMPConfirmOrder(arrList(0, lp))

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
'			response.end
			If (iSelldate < Left(Now(), 10)) then
				Call SetCheckStatus(sellsite, Left(DateAdd("d", 1, CDate(iSellDate)), 10), "N")
			ElseIf (iSellDate = Left(Now(), 10)) then
				Call SetCheckStatus(sellsite, iSellDate, "Y")
			End If
		Else
			rw "주문연동 실패..잠시 후 시도 요망"
		End If
	On Error Goto 0
	Set objXML = nothing
End If
%>