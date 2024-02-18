<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� XML �ֹ�ó��
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteOrderXMLCls.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/coupang/coupangItemcls.asp"-->
<!-- #include virtual="/outmall/coupang/incCoupangFunction.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Function fnCoupangConfirmOrder(vOrderserial, vOutMallOptionNo, vBeasongNum11st)
	Dim objXML, xmlDOM, iRbody, strSql, istrParam
	istrParam = "lstshipmentboxIds="&vBeasongNum11st
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "http://xapi.10x10.co.kr:8080/Orders/Coupang/ready", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If Err.number <> 0 Then
			iErrStr = ivendorItemId
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			strSql = ""
			strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] SET "
			strSql = strSql & " isbaljuConfirmSend = 'Y' "
			strSql = strSql & " , lastUpdate = getdate() "
			strSql = strSql & " WHERE outmallorderserial = '"&vOrderserial&"'  "
			strSql = strSql & " and beasongNum11st = '"&vBeasongNum11st&"' "
			strSql = strSql & " and OrgDetailKey = '"&vOutMallOptionNo&"' "
			strSql = strSql & " and mallid = 'coupang' "
			dbget.Execute strSql
			fnCoupangConfirmOrder= true
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

function saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
        , OrderName, OrderTelNo, OrderHpNo _
        , ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
        , SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
        , deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
        , byref ierrCode, byref ierrStr, beasongNum11st, splitbeasongYn, outMallOptionNo)
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
		,Array("@beasongNum11st"	,adVarchar, adParamInput,16, beasongNum11st) _
		,Array("@requireDetail11stYN"	,adVarchar, adParamInput,1, splitbeasongYn) _
		,Array("@outMallOptionNo"	,adVarchar, adParamInput,16, outMallOptionNo) _
	)

    if (matchItemOption<>"") and (matchItemID<>"-1") and (matchItemID<>"") then
        sqlStr = "db_temp.[dbo].[usp_API_Coupang_OrderReg_Add]"
        retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

        ierrCode = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ�
        ierrStr  = GetValue(retParamInfo, "@retErrStr")   ' �����޼���
    else
        ierrCode = -999
        ierrStr = "��ǰ�ڵ� �Ǵ� �ɼ��ڵ�  ��Ī ����" & OrgDetailKey & " ��ǰ�ڵ� =" & matchItemID&" �ɼǸ� = "&partnerOptionName
        rw "["&ierrCode&"]"&ierrStr
        dbget.close() : response.end
    end if

    saveOrderOneToTmpTable = (ierrCode=0)
    if (ierrCode<>0) then
        rw "["&ierrCode&"]"&ierrStr
    end if
end function

function getLastOrderInputDT()
    dim sqlStr
    sqlStr = "select top 1 convert(varchar(10),selldate,21) as lastOrdInputDt"
    sqlStr = sqlStr&" from db_temp.dbo.tbl_XSite_TMpOrder"
    sqlStr = sqlStr&" where sellsite='coupang'"
    sqlStr = sqlStr&" order by selldate desc"

    rsget.Open sqlStr,dbget,1
	if (Not rsget.Eof) then
		getLastOrderInputDT = rsget("lastOrdInputDt")
	end if
	rsget.Close

end function

Dim sqlStr, buf, i, j, mode, sellsite
Dim divcd, yyyymmdd, idx, Nodes, Nodes2, SubNodes, SubNodes2
Dim objXML, xmlDOM, retCode, iMessage
mode		= requestCheckVar(html2db(request("mode")),32)
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
idx			= requestCheckVar(html2db(request("idx")),32)

Dim tmpxml, strsql, retVal, deliverymemo, orderCsGbn, errCode, errStr, succCNT, failCNT
Dim OrgDetailKey, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, OrderName, OrderTelNo, OrderHpNo
Dim OutMallOrderSerial, SellDate, outMallGoodsNo, matchItemID, partnerItemName, SellPrice, RealSellPrice, ItemOrderCount, PaymentPrice, orderDlvPay, requireDetail, matchItemOption, OptionInfo, tmpOptionVal, outMallOptionNo
Dim partnerOptionName, SalePrice, AddPrice, CouponDiscountsPrice, DiscountsPrice, beasongNum11st, splitbeasongYn, discountPrice
Dim regOrderCnt, strObj, iRbody
Dim LagrgeNode, MiddleNode
Dim iSellDate, iIsSuccess, fromDate, nowDate
Call GetCheckStatus("coupang", iSellDate, iIsSuccess)
rw iSellDate
If sellsite = "coupang" Then
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://xapi.10x10.co.kr:8080/Orders/Coupang/ACCEPT/"&iSellDate, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||" & Err.Description
		End If
		'rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'response.write iRbody
'response.end
			Dim obj1, obj2
			Set strObj = JSON.parse(iRbody)
				set obj1 = strObj.value
					for i=0 to obj1.length-1
						orderCsGbn			= 0
						OrgDetailKey		= i
						beasongNum11st		= obj1.get(i).shipmentBoxId					'��۹�ȣ
						OutMallOrderSerial	= obj1.get(i).orderId						'�ֹ���ȣ
						SellDate			= LEFT(obj1.get(i).orderedAt, 10)			'�ֹ��Ͻ�
						OrderName			= obj1.get(i).orderer.name					'�ֹ��� �̸�
						'rw obj1.get(i).orderer.email									'�ֹ��� email
						OrderHpNo			= obj1.get(i).orderer.safeNumber			'�ֹ��� ����ó(�Ƚɹ�ȣ)
						OrderTelNo			= obj1.get(i).orderer.safeNumber			'�ֹ��� ����ó(�Ƚɹ�ȣ)
						'rw obj1.get(i).paidAt											'�����Ͻ�
						'rw obj1.get(i).status											'���ּ����� | ACCEPT : �����Ϸ�, INSTRUCT : ��ǰ�غ���, DEPARTURE : �������, DELIVERING : �����, FINAL_DELIVERY : ��ۿϷ�, NONE_TRACKING : ��ü�������(��ۿ���������), �����Ұ�
						orderDlvPay			= obj1.get(i).shippingPrice					'��ۺ�
						'rw obj1.get(i).remotePrice										'�����갣��ۺ�
						'rw obj1.get(i).remoteArea										'�����갣����
						deliverymemo		= obj1.get(i).parcelPrintMessage			'��۸޼���
						'rw obj1.get(i).splitShipping									'�и���ۿ���
						splitbeasongYn		= Chkiif(obj1.get(i).ableSplitShipping = "True", "Y","N")				'�и���۰��ɿ���
						ReceiveName			= obj1.get(i).receiver.name					'������ �̸�
						ReceiveHpNo			= obj1.get(i).receiver.safeNumber			'������ ����ó(�Ƚɹ�ȣ)
						ReceiveTelNo		= obj1.get(i).receiver.safeNumber			'������ ����ó(�Ƚɹ�ȣ)
						ReceiveAddr1		= obj1.get(i).receiver.addr1				'������ �����1
						ReceiveAddr2		= obj1.get(i).receiver.addr2				'������ �����2
						ReceiveZipCode		= obj1.get(i).receiver.postCode				'������ �����ȣ

						If Len(OrderName) > 16 Then
							'�ֹ��ڰ� �ܱ����Ͻ� ���̰� ��ħ(2019-09-17 �߰�)
							OrderName = Left(OrderName, 16)
						End If

						if Len(ReceiveName) > 16 then
							'// ������ �̸��� �ּҰ� ���� ���̽��� ����(2018-11-08)
							ReceiveName = Left(ReceiveName,16)
						end if

						set obj2 = obj1.get(i).orderItems
							For j=0 to obj2.length-1
								'rw obj2.get(j).vendorItemPackageId						'vendorItemPackageId | optional / ���� ��� 0���� ����
								'rw obj2.get(j).vendorItemPackageName					'vendorItemPackageName | optional
								'rw obj2.get(j).productId								'productId | optional / ���� ��� 0���� ����
								SalePrice = ""
								discountPrice = ""
								outMallOptionNo		= obj2.get(j).vendorItemId			'vendorItemId
								'rw obj2.get(j).vendorItemName							'vendorItemName
								ItemOrderCount		= obj2.get(j).shippingCount			'item count to deliver(It must excludes cancel count)
								SellPrice			= obj2.get(j).salesPrice			'���� ��ǰ ����(price of one item)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'							2018-09-27 11:21 ������ �ּ�ó��
'								SalePrice			= obj2.get(j).discountPrice			'���� ����..2018-07-24 ������ ����
'								If isnull(SalePrice) OR SalePrice = "" OR SalePrice = "null" Then
'									SalePrice = 0
'								End If
'
'								If SalePrice <> 0 Then
'									SalePrice			= CLng(SalePrice / ItemOrderCount)
'								End If
'								'rw obj2.get(j).orderPrice								'���� ���� salesPrice * shippingCount : ������ �׻� ������ ����
'							2018-09-27 11:21 ������ �ּ�ó�� ��
								'RealSellPrice = SellPrice - SalePrice					'���ǸŰ��� = SalePrice - DiscountsPrice Ȯ���ʿ�!!
'							2018-09-27 11:21 ������ �ϴ����� ����
'								If ItemOrderCount = 0 Then
'									RealSellPrice = SellPrice
'								Else
'									RealSellPrice = obj2.get(j).orderPrice / ItemOrderCount
'								End If
''''''''''''''''''''''''''''''''''''''''''''''2019-09-10 13:45 ������ �Ʒ��� ����''''''''''''''''''''''''''''''''''''''''''''''''''''''
								If ItemOrderCount = 0 Then
									RealSellPrice = SellPrice
								Else
									discountPrice			= obj2.get(j).discountPrice			'���� ����..2019-09-10 / �ٽ� �� �ʵ� ����غ�
									If isnull(discountPrice) OR discountPrice = "" OR discountPrice = "null" Then
										SalePrice = 0
									Else
										SalePrice = CLng(discountPrice / ItemOrderCount)	'���ΰ����� ������ ���ؼ� ����..���� ������ ���� ����
									End If
									RealSellPrice = SellPrice - SalePrice					'���ǸŰ���
								End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
								matchItemID			= Split(obj2.get(j).externalVendorSkuCode, "_")(0)		'external code / optional
								matchItemOption		= Split(obj2.get(j).externalVendorSkuCode, "_")(1)		'external code / optional
								'rw obj2.get(j).etcInfoHeader							'��ǰ�� ���� �Է� �׸� / optional
								'rw obj2.get(j).etcInfoValue							'��ǰ�� ���� �Է� �׸� ���� ������� �Է°� / optional : �ʵ�� �����ϳ� ���� ���� �����Դϴ�. �ʿ�ÿ��� �Ʒ��� etcInfoValues�� ����Ͻñ� �ٶ��ϴ�.
								'rw obj2.get(j).etcInfoValues							'��ǰ�� ���� �Է� �׸� ���� ������� �Է°� ����Ʈ / optional : v4 version���θ� ��ȸ����
								outMallGoodsNo		= obj2.get(j).sellerProductId		'��ü��ǰ ���̵�
								partnerItemName		= obj2.get(j).sellerProductName		'��ü��ǰ��
								partnerOptionName	= obj2.get(j).sellerProductItemName	'��ϿɼǸ�
								'rw obj2.get(j).firstSellerProductItemName	'���ʵ�ϿɼǸ�
								'rw obj2.get(j).cancelCount					'��Ҽ���
								'rw obj2.get(j).holdCountForCancel			'ȯ�Ҵ�����
								'rw obj2.get(j).estimatedShippingDate		'�ֹ��� ������� / optional / yyyy-mm-dd
								'rw obj2.get(j).plannedShippingDate			'���� ��� ������(�и���۽�)
								'rw obj2.get(j).invoiceNumberUploadDate		'������ȣ ���ε� �Ͻ� | optional / yyyy-mm-dd
								'rw obj2.get(j).pricingBadge				'������ ��ǰ ���� | true/false : v4 version���θ� ��ȸ����
								'rw obj2.get(j).usedProduct					'�߰� ��ǰ ���� | true/false : v4 version���θ� ��ȸ����
								'rw obj2.get(j).confirmDate					'����Ȯ������ | yyyy-MM-dd HH:mm:ss : v4 version���θ� ��ȸ����
								'rw obj2.get(j).deliveryChargeTypeName		'��ۺ񱸺� | ����, ���� : v4 version���θ� ��ȸ����
								'rw obj2.get(j).canceled					'�ֹ� ��� ���� | true/false

'								rw "beasongNum11st : " & beasongNum11st
'								rw "OutMallOrderSerial : " & OutMallOrderSerial
'								rw "SellDate : " & SellDate
'								rw "OrderName : " & OrderName
'								rw "OrderHpNo : " & OrderHpNo
'								rw "orderDlvPay : " & orderDlvPay
'								rw "deliverymemo : " & deliverymemo
'								rw "splitbeasongYn : " & splitbeasongYn
'								rw "ReceiveName : " & ReceiveName
'								rw "ReceiveHpNo : " & ReceiveHpNo
'								rw "ReceiveAddr1 : " & ReceiveAddr1
'								rw "ReceiveAddr2 : " & ReceiveAddr2
'								rw "ReceiveZipCode : " & ReceiveZipCode
'
'								rw "outMallOptionNo : " & outMallOptionNo
'								rw "ItemOrderCount : " & ItemOrderCount
'								rw "SalePrice : " & SalePrice
'								rw "DiscountsPrice : " & DiscountsPrice
'								rw "RealSellPrice : " & RealSellPrice
'								rw "matchItemID : " & matchItemID
'								rw "matchItemOption : " & matchItemOption
'								rw "outMallGoodsNo : " & outMallGoodsNo
'								rw "partnerItemName : " & partnerItemName
'								rw "partnerOptionName : " & partnerOptionName

								retVal= saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
										, OrderName, OrderTelNo, OrderHpNo _
										, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
										, SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
										, deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
										, errCode, errStr, beasongNum11st, splitbeasongYn, outMallOptionNo )

								If (retVal) Then
									succCNT = succCNT + 1
									strsql = ""
									strsql = strsql & " INSERT INTO db_temp.[dbo].[tbl_xSite_TMP11stOrder] (outmallorderserial, OrgDetailKey, beasongNum11st, isbaljuConfirmSend, regdate, mallid) "
									strsql = strsql & " VALUES ('"&OutMallOrderSerial&"', '"&outMallOptionNo&"', '"&beasongNum11st&"', 'N', getdate(), 'coupang')"
									dbget.Execute strSql
								Else
									failCNT = failCNT + 1
								End If
							Next
						set obj2 = nothing
						'rw "--------------------------------------------------------"


						'rw obj1.get(i).overseaShippingInfoDto.personalCustomsClearanceCode
						'rw obj1.get(i).overseaShippingInfoDto.ordererSsn
						'rw obj1.get(i).overseaShippingInfoDto.ordererPhoneNumber

						'rw obj1.get(i).deliveryCompanyName					'�ù�� | CJ �������, �����ù� : v4 version���θ� ��ȸ����
						'rw obj1.get(i).invoiceNumber						'������ȣ : v4 version���θ� ��ȸ����
						'rw obj1.get(i).inTrasitDateTime					'�����(�߼���) | yyyy-MM-dd HH:mm:ss : v4 version���θ� ��ȸ����
						'rw obj1.get(i).deliveredDate						'��ۿϷ��� | yyyy-MM-dd HH:mm:ss : v4 version���θ� ��ȸ����
						'rw obj1.get(i).refer								'������ġ | ��������, �ȵ���̵��, PC�� : v4 version���θ� ��ȸ����
						'rw "------------------------------------"
					Next
				set obj1 = nothing
			Set strObj = nothing

			If (failCNT <> 0) Then
			    rw "["&failCNT&"] �� ����(�ֹ���ȸ)"
			End if

			If (succCNT <> 0) then
			    rw "["&succCNT&"] �� ����(�ֹ���ȸ)"
			    ' Dim arrList, lp, ret1
			    ' Dim OKcnt, NOcnt
			    ' OKcnt = 0
			    ' NOcnt = 0

				' strsql = ""
				' strsql = strsql & " update T "
				' strsql = strsql & " set T.isbaljuConfirmSend='Y' "
				' strsql = strsql & " From db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T "
				' strsql = strsql & " JOIN db_temp.dbo.tbl_xsite_tmporder as O on T.outmallorderserial = O.OutMallOrderSerial and T.OrgDetailKey = O.outMallOptionNo "
				' strsql = strsql & " where T.isbaljuConfirmSend <> 'Y' "
				' strsql = strsql & " and O.sendState = 1 "
				' strsql = strsql & " and O.matchstate in ('O') "
				' strsql = strsql & " and T.mallid = 'coupang' "
				' dbget.Execute strsql

				' strsql = ""
				' strsql = strsql & " update T "
				' strsql = strsql & " set T.isbaljuConfirmSend='Y' "
				' strsql = strsql & " FROM db_order.dbo.tbl_order_master as M "
				' strsql = strsql & " JOIN db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T on M.authcode = T.outmallorderserial "
				' strsql = strsql & " WHERE M.cancelyn ='Y' "
				' strsql = strsql & " and T.isbaljuConfirmSend <> 'Y' "
				' strsql = strsql & " and T.mallid = 'coupang' "
				' dbget.Execute strsql

				' strsql = ""
				' strsql = strsql & " SELECT TOP 1000 outmallorderserial, OrgDetailKey, beasongNum11st FROM db_temp.[dbo].[tbl_xSite_TMP11stOrder] "
				' strsql = strsql & " WHERE isbaljuConfirmSend = 'N' "
				' strsql = strsql & " and mallid = 'coupang' "
			    ' rsget.Open strsql,dbget,1
			    ' if not rsget.Eof then
			    '     arrList = rsget.getRows()
			    ' end if
			    ' rsget.close

				' For lp = 0 To Ubound(arrList, 2)
				' 	ret1 = fnCoupangConfirmOrder(arrList(0, lp), arrList(1, lp), arrList(2, lp))

	            '     If (ret1) then
	            '         OKcnt = OKcnt + 1
	            '     Else
	            '         NOcnt = NOcnt + 1
	            '     End If
				' Next

				' If OKcnt <> 0 then
				' 	rw "["&OKcnt&"] �� ����(����Ȯ��)"
				' End If

				' If NOcnt <> 0 then
				' 	rw "["&NOcnt&"] �� ����(����Ȯ��)"
				' End If
			End If
'			response.end
			If (iSelldate < Left(Now(), 10)) then
				Call SetCheckStatus(sellsite, Left(DateAdd("d", 1, CDate(iSellDate)), 10), "N")
			ElseIf (iSellDate = Left(Now(), 10)) then
				Call SetCheckStatus(sellsite, iSellDate, "Y")
			End If
		Else
			rw "�ֹ����� ����..��� �� �õ� ���"
		End If
	On Error Goto 0
	Set objXML = nothing
End If


''ǰ��/���� ����üũ
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->