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
    iedyyyymmdd = LEFT(dateadd("d",-1,now()),10)

    if eddt<>"" then
        if isDate(eddt) then
            iedyyyymmdd=eddt  '''yyyy-mm-dd
        end if
    end if
    istyyyymmdd = LEFT(dateadd("d",-30,iedyyyymmdd),10)
''

'' 조회 유형 (NEW:신규주문 ,CONFIRM:발송처리대상, DELIVERY:배송중, COMPLETE:배송완료)

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'WMP','"&confirmDt&"'"
dbget.Execute sqlStr
rw "초기화작업"

dim datelen : datelen=datediff("d",istyyyymmdd, iedyyyymmdd)
dim thedate , k
for k=0 to datelen-1
    thedate=dateadd("d",-1*k,iedyyyymmdd)
    if k<5 then
    call Get_WMPOrderListByStatus(thedate,thedate,"NEW","주문통보")
    response.flush
    end if
    call Get_WMPOrderListByStatus(thedate,thedate,"CONFIRM","주문확인")
    response.flush
    call Get_WMPOrderListByStatus(thedate,thedate,"DELIVERY","출고완료")
    response.flush

    '' call Get_WMPOrderListByStatus(istyyyymmdd,iedyyyymmdd,"COMPLETE","배송완료")
    '' response.flush
next

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'WMP','"&confirmDt&"'"
dbget.Execute sqlStr
rw "주문매핑"

rw "완료"
'response.write("<script>setTimeout(alert('완료'),1000);self.close();</script>")

function Get_WMPOrderListByStatus(stdate,eddate,iorderStatus,istatusName)
	dim sellsite : sellsite = "WMP"
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

	Get_WMPOrderListByStatus = False

	'// =======================================================================
	'// 날짜형식
	''selldate = "2017-11-10"
	''xmlSelldate = Replace(selldate, "-", "")

	'// API URL(기간동안의 주문 가져오기)
	xmlURL = "http://110.93.128.100:8090/wemake/Orders/orderlist"


	postParam = ""
	postParam = postParam & "reqdate=" & stdate
	''postParam = postParam & "&enddate=" & Left(DateAdd("d", 1, CDate(eddate)), 10)
	postParam = postParam & "&type="&iorderStatus
    if (iorderStatus="NEW") then
        postParam = postParam & "&DateType=NEW"     ''결제완료일
    else
        postParam = postParam & "&DateType=CONFIRM" ''주문확인일
    end if
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

    Dim iRbody, strObj, orderCount, obj1, obj2, obj3


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
    if Not isObject(strObj.outPutValue.data.bundle) then
        rw "No outPutValue"
        exit function
    end if

        set obj1 = strObj.outPutValue.data.bundle

        'rw strObj.totalcount & ":" &obj1.length

        If obj1.length >0 Then
            response.write "주문건수(" & obj1.length & ") " & "<br />"
            for i=0 to obj1.length-1
                ordNo           = obj1.get(i).bundleNo				'주문번호(배송번호)

                shppSeq			= ""			'배송지시상세번호
                reOrderYn ="N" ''재주문여부
                delayNts  =""  ''지연일수

                shppTypeDtlNm   = obj1.get(i).delivery.shipMethod
                delicoVenId     = ""	           								'택배배송사코드
                wblNo           = obj1.get(i).delivery.invoiceNo									'운송장번호
                if (shppTypeDtlNm="기타배송") then
                    wblNo = wblNo & obj1.get(i).delivery.shipMethodMessage					'배송방법 메세지 배송방법이 [기타배송]일 경우 입력받는 메세지
                end if
                delicoVenNm     = obj1.get(i).delivery.parcelCompany
                orderStatus     = obj1.get(i).delivery.shipStatus              '발주서상태 | ACCEPT/INSTRUCT/DEPARTURE/DELIVERING/FINAL_DELIVERY/NONE_TRACKING

                whoutCritnDt    = obj1.get(i).originShipDate	 '' 발송기한.
                outjFixedDt     = obj1.get(i).shipCompleteDate ''구매확정일자  - 업체직송인경우 7일후 완료된다. 정산이 안되면 업체직송으로 수정해야한다.



                set obj2 = obj1.get(i).orderProduct
                    For j=0 to obj2.length-1
                        shppNo		    = obj2.get(j).orderNo			    'reserve01(주문번호)

                        cspGoodsCd      = obj2.get(j).sellerProductCode	'업체상품코드
                        goodsCd         = obj2.get(j).productNo		                    '판매상품코드
                        uitemId         = obj2.get(j).sellerProductCode				                '상품속성코드

                        shppDivDtlNm = ""
                        ' if (obj2.get(j).canceled="true") then
                        '     shppDivDtlNm = "취소"
                        ' end if
                        ' if (obj2.get(j).cancelCount<>0) then
                        '     shppDivDtlNm = shppDivDtlNm & CHKIIF(shppDivDtlNm<>"","/","") & obj2.get(j).cancelCount      ''취소수량
                        ' end if

                        shppRsvtDt      = ""''예정일
                        autoShortgYn    = "" ''자동결품여부
                        invoiceUpDt = "" ''운송장번호 업로드 일시 (이게 오래된거면 추적(집하)이 안된거 일 수 있다.)


                        set obj3 = obj2.get(j).orderOption
						    For k=0 to obj3.length-1
                                ordItemSeq      = obj3.get(k).orderOptionNo		'주문옵션번호
                                uitemId		    = obj3.get(k).optionNo			'옵션번호
                                optionContent	= obj3.get(k).optionName		'옵션
                                orderQty		= obj3.get(k).optionQty			'수량

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

                            next



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

	Get_WMPOrderListByStatus = True

end function
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->