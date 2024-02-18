<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="javascript" runat="server">
var confirmDt = (new Date()).valueOf();
</script>
<style>
body {
  font-size: small;
}
</style>
</head>
<body bgcolor="#F4F4F4" >
<%


Dim sqlStr
Dim isOnlyTodayBaljuView : isOnlyTodayBaljuView=false  ''발주된내역만 (주문수신 view)
Dim isDlvConfirmProc 	 : isDlvConfirmProc=false  ''주문확인 Proc
Dim isDlvInputProc 	 	 : isDlvInputProc=false    ''입력 Proc
Dim eddt : eddt=requestCheckvar(request("eddt"),10)

Dim IS_TEST_MODE : IS_TEST_MODE=FALSE

Dim istyyyymmdd, iedyyyymmdd
    iedyyyymmdd = LEFT(dateadd("d",-2,now()),10)

    if eddt<>"" then
        if isDate(eddt) then
            iedyyyymmdd=eddt  '''yyyy-mm-dd
        end if
    end if
    istyyyymmdd = LEFT(dateadd("d",-2,iedyyyymmdd),10)

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'coupang','"&confirmDt&"'"
dbget.Execute sqlStr
rw "초기화작업"

'' ACCEPT:결제완료, INSTRUCT:상품준비중, DEPARTURE:배송지시, DELIVERING:배송중, NONE_TRACKING:업체직접배송
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"ACCEPT","주문통보")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"INSTRUCT","주문확인")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"DEPARTURE","출고완료")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"DELIVERING","배송중")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"NONE_TRACKING","업체직접배송")
response.flush

call Get_CoupangExchangeListByStatus(LEFT(dateadd("d", 0,istyyyymmdd),10))
call Get_CoupangExchangeListByStatus(LEFT(dateadd("d", 1,istyyyymmdd),10))
call Get_CoupangExchangeListByStatus(LEFT(dateadd("d", 2,istyyyymmdd),10))
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-3,iedyyyymmdd),10)
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"ACCEPT","주문통보")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"INSTRUCT","주문확인")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"DEPARTURE","출고완료")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"DELIVERING","배송중")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"NONE_TRACKING","업체직접배송")
response.flush

call Get_CoupangExchangeListByStatus(LEFT(dateadd("d", 0,istyyyymmdd),10))
call Get_CoupangExchangeListByStatus(LEFT(dateadd("d", 1,istyyyymmdd),10))
call Get_CoupangExchangeListByStatus(LEFT(dateadd("d", 2,istyyyymmdd),10))
call Get_CoupangExchangeListByStatus(LEFT(dateadd("d", 3,istyyyymmdd),10))
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-10,iedyyyymmdd),10)
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"ACCEPT","주문통보")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"INSTRUCT","주문확인")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"DEPARTURE","출고완료")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"DELIVERING","배송중")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"NONE_TRACKING","업체직접배송")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-10,iedyyyymmdd),10)
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"ACCEPT","주문통보")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"INSTRUCT","주문확인")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"DEPARTURE","출고완료")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"DELIVERING","배송중")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"NONE_TRACKING","업체직접배송")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-10,iedyyyymmdd),10)
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"ACCEPT","주문통보")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"INSTRUCT","주문확인")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"DEPARTURE","출고완료")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"DELIVERING","배송중")
response.flush
call Get_CoupangOrderListByStatus(istyyyymmdd,iedyyyymmdd,"NONE_TRACKING","업체직접배송")
response.flush

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'coupang','"&confirmDt&"'"
dbget.Execute sqlStr
rw "주문매핑"

rw "완료"
'response.write("<script>setTimeout(alert('완료'),1000);self.close();</script>")

function Get_CoupangOrderListByStatus(stdate,eddate,iorderStatus,istatusName)
	dim sellsite : sellsite = "coupang"
	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData
	dim masterCnt, detailCnt, resultcode, obj
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim i, j, k
	dim tmpStr, pos
	dim successCnt : successCnt = 0
	dim strRst
	dim tmpOptionSeq : tmpOptionSeq = 0
	dim postParam
	dim tmpXML, oSql

    dim strSql, bufStr

	Get_CoupangOrderListByStatus = False

	'// =======================================================================
	'// 날짜형식
	''selldate = "2017-11-10"
	''xmlSelldate = Replace(selldate, "-", "")

	'// API URL(기간동안의 주문 가져오기)
	xmlURL = "http://xapi.10x10.co.kr:8080/coupangnew/etc/oderlist"


	postParam = ""
	postParam = postParam & "startdate=" & stdate
	postParam = postParam & "&enddate=" & Left(DateAdd("d", 1, CDate(eddate)), 10)
	postParam = postParam & "&status="&iorderStatus
	''response.write postParam

    rw "기간검색:"&stdate&"~"&eddate&" 상태:"&iorderStatus&"("&istatusName&")"
	'// =======================================================================
	'// 데이타 가져오기


	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL&"?"&postParam, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.send()

	if objXML.Status <> "200" then
		response.write "ERROR : 통신오류" & objXML.Status
		dbget.close : response.end
	end if

    Dim iRbody, strObj, orderCount, obj1, obj2


    Dim ordNo, ordItemSeq, shppNo, shppSeq, reOrderYn, delayNts
    Dim cspGoodsCd, goodsCd, uitemId, orderQty, shppDivDtlNm
    Dim optionContent, shppRsvtDt, whoutCritnDt, autoShortgYn
    Dim orderStatus, dlvrCd, dlvrNo, dlvrDt, dlvrFinishDt, cancelDt
    Dim paramInfo, retParamInfo, RetErr

    Dim shppTypeDtlNm, delicoVenId, delicoVenNm, wblNo
	Dim invoiceUpDt, outjFixedDt

	iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'rw"<textarea cols=80 rows=20>"&iRbody&"</textarea>"
'exit function

    Set strObj = JSON.parse(iRbody)
    if Not isObject(strObj.outPutValue) then
        rw "No outPutValue"
        exit function
    end if

        set obj1 = strObj.outPutValue

        'rw strObj.totalcount & ":" &obj1.length

        If obj1.length >0 Then
            response.write "주문건수(" & obj1.length & ") " & "<br />"
            ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] '"&sellsite&"','"&confirmDt&"'"
            ' dbget.Execute strSql

            for i=0 to obj1.length-1
                ordNo           = obj1.get(i).orderId				'주문번호
                ordItemSeq      = -1				'주문일련번호 없는듯.
                shppNo		    = obj1.get(i).shipmentBoxId			'배송번호(묶음배송번호)
                shppSeq			= ""			'배송지시상세번호
                reOrderYn ="N" ''재주문여부
                delayNts  =""  ''지연일수

                shppTypeDtlNm = ""
                delicoVenId     = ""            								'택배배송사코드
                wblNo           = obj1.get(i).invoiceNumber									'운송장번호
                delicoVenNm     = obj1.get(i).deliveryCompanyName
                orderStatus     = obj1.get(i).status              '발주서상태 | ACCEPT/INSTRUCT/DEPARTURE/DELIVERING/FINAL_DELIVERY/NONE_TRACKING

                set obj2 = obj1.get(i).orderItems
                    For j=0 to obj2.length-1
                        cspGoodsCd      = Split(obj2.get(j).externalVendorSkuCode, "_")(0)				'TEN 상품관리코드
                        goodsCd         = obj2.get(j).sellerProductId		                    '판매상품코드
                        uitemId         = obj2.get(j).vendorItemId				        '상품속성코드
                        orderQty        = obj2.get(j).shippingCount				    '출고수량(취소수량반영됨.)

                        shppDivDtlNm = ""
                        if (obj2.get(j).canceled="true") then
                            shppDivDtlNm = "취소"
                        end if
                        if (obj2.get(j).cancelCount<>0) then
                            shppDivDtlNm = shppDivDtlNm & CHKIIF(shppDivDtlNm<>"","/","") & obj2.get(j).cancelCount      ''취소수량
                        end if

                        optionContent   = obj2.get(j).sellerProductItemName				'등록옵션명
                        shppRsvtDt      = ""''예정일
                        whoutCritnDt    = obj2.get(j).estimatedShippingDate	 ''estimatedShippingDate	    '최종출고마감일자  (''출고기준일)
                        autoShortgYn    = "" ''자동결품여부

                        invoiceUpDt = replace(Null2Blank(obj2.get(j).invoiceNumberUploadDate),"T"," ") ''운송장번호 업로드 일시 (이게 오래된거면 추적(집하)이 안된거 일 수 있다.)
                        outjFixedDt = replace(Null2Blank(obj2.get(j).confirmDate),"T"," ") ''구매확정일자  - 업체직송인경우 7일후 완료된다. 정산이 안되면 업체직송으로 수정해야한다.

                        bufStr = ""
                        bufStr = sellsite&"|"&ordNo
                        bufStr = bufStr &"|"&ordItemSeq
                        bufStr = bufStr &"|"&shppNo
                        bufStr = bufStr &"|"&shppSeq
                        bufStr = bufStr &"|"&cspGoodsCd
                        bufStr = bufStr &"|"&goodsCd

                        bufStr = bufStr &"|"&uitemId
                        bufStr = bufStr &"|"&orderQty
                        bufStr = bufStr &"|"&shppDivDtlNm

                        bufStr = bufStr &"|"&optionContent
                        bufStr = bufStr &"|"&whoutCritnDt


                        bufStr = bufStr &"|"&orderStatus
                        bufStr = bufStr &"|"&shppTypeDtlNm
                        bufStr = bufStr &"|"&delicoVenId
                        bufStr = bufStr &"|"&wblNo
                        bufStr = bufStr &"|"&delicoVenNm
    'rw bufStr
                        ' if (whoutCritnDt<>"") then
                        '     whoutCritnDt = LEFT(whoutCritnDt,4)&"-"&MID(whoutCritnDt,5,2)&"-"&RIGHT(whoutCritnDt,2)
                        ' end if


                        sqlStr = "db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Input]"
                        paramInfo = Array(Array("@RETURN_VALUE",adInteger	,adParamReturnValue	,,0) _
                            ,Array("@SellSite" 				, adVarchar		, adParamInput		, 	32, sellsite)	_
                            ,Array("@OutMallOrderSerial"	, adVarchar		, adParamInput		,	32, Trim(ordNo)) _
                            ,Array("@OrgDetailKey"			, adVarchar		, adParamInput		,   32, Trim(ordItemSeq)) _

                            ,Array("@confirmDt"				, adVarchar     , adParamInput		,	16, Trim(confirmDt)) _
                            ,Array("@shppNo"				, adVarchar		, adParamInput		,   32, Trim(shppNo)) _
                            ,Array("@shppSeq"				, adVarchar		, adParamInput		,   10, Trim(shppSeq)) _
                            ,Array("@reOrderYn"				, adVarchar		, adParamInput		,    1, Trim(reOrderYn)) _
                            ,Array("@delayNts"			    , adInteger		, adParamInput		,     , Trim(delayNts)) _
                            ,Array("@splVenItemId"			, adInteger		, adParamInput		,     , Trim(cspGoodsCd)) _
                            ,Array("@outMallGoodsNo"		, adVarchar		, adParamInput		,   16, Trim(goodsCd)) _
                            ,Array("@outMallOptionNo"		, adVarchar		, adParamInput		,   32, Trim(uitemId)) _
                            ,Array("@ordQty"			    , adInteger		, adParamInput		,     , Trim(orderQty)) _
                            ,Array("@shppDivDtlNm"		    , adVarchar		, adParamInput		,   20, Trim(shppDivDtlNm)) _
                            ,Array("@uitemNm"		        , adVarchar		, adParamInput		,   128, Trim(optionContent)) _
                            ,Array("@shppRsvtDt"			, adDate		, adParamInput		,	  , Trim(shppRsvtDt)) _
                            ,Array("@whoutCritnDt"			, adDate		, adParamInput		,	  , Trim(whoutCritnDt)) _
                            ,Array("@autoShortgYn"			, adVarchar		, adParamInput		,    1, Trim(autoShortgYn)) _
                            ,Array("@outorderstatus"		, adVarchar		, adParamInput		,   30, Trim(orderStatus)) _

                            ,Array("@shppTypeDtlNm"		, adVarchar		, adParamInput		,   16, Trim(shppTypeDtlNm)) _
                            ,Array("@delicoVenId"		, adVarchar		, adParamInput		,   16, Trim(delicoVenId)) _
                            ,Array("@delicoVenNm"		, adVarchar		, adParamInput		,   32, Trim(delicoVenNm)) _
                            ,Array("@wblNo"		        , adVarchar		, adParamInput		,   32, Trim(wblNo)) _

                            ,Array("@invoiceUpDt"	    , adVarchar		, adParamInput		,   19, Trim(invoiceUpDt)) _
                            ,Array("@outjFixedDt"		, adVarchar		, adParamInput		,   19, Trim(outjFixedDt)) _

                        )

                        'On Error RESUME Next
                        retParamInfo = fnExecSPOutput(sqlStr, paramInfo)
                        ' If ERR then
                        '     rw invoiceUpDt
                        '     rw outjFixedDt
                        '     response.end
                        ' end if
                        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드

                        successCnt = successCnt+1
                    next
                set obj2 = nothing
            next
            set obj1 = nothing
        End If
    Set strObj = nothing

    '' 주문번호 매핑.
    ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] '"&sellsite&"','"&confirmDt&"'"
    ' dbget.Execute strSql

    rw "상세건수:"&successCnt
    rw "======================================"

	Get_CoupangOrderListByStatus = True

end function

function Get_CoupangExchangeListByStatus(stdate)
	dim sellsite : sellsite = "coupang"
	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData
	dim masterCnt, detailCnt, resultcode, obj
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim i, j, k
	dim tmpStr, pos
	dim successCnt : successCnt = 0
	dim strRst
	dim tmpOptionSeq : tmpOptionSeq = 0
	dim postParam
	dim tmpXML, oSql

    dim strSql, bufStr

	Get_CoupangExchangeListByStatus = False

	'// =======================================================================
	'// 날짜형식
	''selldate = "2017-11-10"

	'// API URL(기간동안의 교환출고/교환회수 가져오기)
	xmlURL = "http://xapi.10x10.co.kr:8080/Deliveries/Coupang/exchange/PROGRESS/" & stdate
    'objXML.open "GET", "http://xapi.10x10.co.kr:8080/Deliveries/Coupang/exchange/RECEIPT/"&startdate, false
	'objXML.open "GET", "http://xapi.10x10.co.kr:8080/Deliveries/Coupang/exchange/PROGRESS/"&startdate, false

    rw "기간검색:" & stdate
	'// =======================================================================
	'// 데이타 가져오기


	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.send()

	if objXML.Status <> "200" then
		response.write "ERROR : 통신오류" & objXML.Status
		dbget.close : response.end
	end if


    Dim iRbody, strObj, strObjValues, orderCount, obj1, obj2, retCode


    Dim ordNo, ordItemSeq, shppNo, shppSeq, reOrderYn, delayNts, orordNo, orordItemSeq
    Dim cspGoodsCd, goodsCd, uitemId, orderQty, shppDivDtlNm
    Dim optionContent, shppRsvtDt, whoutCritnDt, autoShortgYn
    Dim orderStatus, dlvrCd, dlvrNo, dlvrDt, dlvrFinishDt, cancelDt
    Dim paramInfo, retParamInfo, RetErr

    Dim shppTypeDtlNm, delicoVenId, delicoVenNm, wblNo, strObjValuesexchangeItems
	Dim invoiceUpDt, outjFixedDt
    dim targetItemId, targetItemName

	iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
    ''rw"<textarea cols=80 rows=20>"&iRbody&"</textarea>"
    ''exit function

    Set strObj = JSON.parse(iRbody)
    retCode			= strObj.message

    if (retCode = "SUCCESS") then
        set strObjValues = strObj.value

        For i=0 to strObjValues.length-1
            '' orderId : 원주문번호
            '' exchangeId : 교환주문번호
            '' exchangeStatus : 교환상태

            '' exchangeItemDtoV1s.exchangeItemId : 교환주문디테일번호
            '' exchangeItemDtoV1s.orderItemId : 회수 상품코드(outmalloptionno)
            '' exchangeItemDtoV1s.targetItemId : 출고 상품코드
            '' exchangeItemDtoV1s.quantity : 수량
            '' exchangeItemDtoV1s.targetItemName : 출고 상품명

            '' deliveryInvoiceGroupDtos.shipmentBoxId
            '' deliveryInvoiceGroupDtos.deliveryInvoiceDtos.invoiceNumber
            '' deliveryInvoiceGroupDtos.deliveryInvoiceDtos.deliverCode
            '' deliveryInvoiceGroupDtos.deliveryInvoiceDtos.
            '' deliveryInvoiceGroupDtos.deliveryInvoiceDtos.
            '' deliveryInvoiceGroupDtos.
            '' deliveryInvoiceGroupDtos.

            '' deliveryStatus : 출고상태
            '' collectStatus : 회수상태

            '' returnDeliveryDtos.deliveryCompanyCode : 회수 택배사
            '' returnDeliveryDtos.deliveryInvoiceNo : 회수 송장번호

            orordNo = strObjValues.get(i).orderId
            ordNo = strObjValues.get(i).exchangeId
            shppNo		    = ""
            shppSeq			= ""
            reOrderYn 		= "N"
            delayNts  		= ""
            shppTypeDtlNm 	= ""
            delicoVenId     = ""
            wblNo           = ""
            delicoVenNm     = ""
            wblNo           = ""
            orderStatus     = strObjValues.get(i).exchangeStatus					'교환상태 | RECEIPT : 접수, PROGRESS : 진행, SUCCESS : 완료, REJECT : 불가, CANCEL : 철회

            set strObjValuesexchangeItems = strObjValues.get(i).exchangeItemDtoV1s
            For j=0 to strObjValuesexchangeItems.length-1
                ordItemSeq = strObjValuesexchangeItems.get(j).exchangeItemId
                orordItemSeq = strObjValuesexchangeItems.get(j).orderItemId		'// outmalloptionno 가 온다.

                strSql = " select top 1 orgdetailkey, matchitemid "
                strSql = strSql & " from "
                strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPOrder] "
                strSql = strSql & " where "
                strSql = strSql & " 	1 = 1 "
                strSql = strSql & " 	and OutMallOrderSerial = '" & orordNo & "' "
                strSql = strSql & " 	and SellSite = '" & sellsite & "' "
                strSql = strSql & " 	and outmalloptionno = '" & orordItemSeq & "' "
                dbget.CursorLocation = adUseClient
                rsget.Open strSql,dbget,adOpenForwardOnly,adLockReadOnly
                if  not rsget.EOF  then
                    orordItemSeq 	= rsget("orgdetailkey")
                    cspGoodsCd 		= rsget("matchitemid")
                end if
                rsget.close()

                goodsCd = strObjValuesexchangeItems.get(j).orderItemId
                uitemId = ""
                orderQty = strObjValuesexchangeItems.get(j).quantity
                shppDivDtlNm = ""
                optionContent = strObjValuesexchangeItems.get(j).targetItemName
                shppRsvtDt      = ""
                whoutCritnDt    = ""
                autoShortgYn	= ""
                invoiceUpDt		= ""
                outjFixedDt 	= ""

                bufStr = ""
                bufStr = sellsite&"|"&ordNo
                bufStr = bufStr &"|"&ordItemSeq
                bufStr = bufStr &"|"&shppNo
                bufStr = bufStr &"|"&shppSeq
                bufStr = bufStr &"|"&cspGoodsCd
                bufStr = bufStr &"|"&goodsCd

                bufStr = bufStr &"|"&uitemId
                bufStr = bufStr &"|"&orderQty
                bufStr = bufStr &"|"&shppDivDtlNm

                bufStr = bufStr &"|"&optionContent
                bufStr = bufStr &"|"&whoutCritnDt


                bufStr = bufStr &"|"&orderStatus
                bufStr = bufStr &"|"&shppTypeDtlNm
                bufStr = bufStr &"|"&delicoVenId
                bufStr = bufStr &"|"&wblNo
                bufStr = bufStr &"|"&delicoVenNm

                bufStr = bufStr &"|"&orordNo
                bufStr = bufStr &"|"&orordItemSeq
                rw bufStr
            next

            successCnt = successCnt + 1
        Next
    end if

    rw "상세건수:"&successCnt
    rw "======================================"

	Get_CoupangExchangeListByStatus = True

end function

%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
