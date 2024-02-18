<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
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
function getHmallDlvCode2Name(idlvCd)
    if isNULL(idlvCd) then Exit function

    SELECT CASE idlvCd

        CASE "12" : getHmallDlvCode2Name = "CJ대한통운"
        CASE "11" : getHmallDlvCode2Name = "롯데택배"
        CASE "13" : getHmallDlvCode2Name = "한진택배"
        CASE "33" : getHmallDlvCode2Name = "로젠택배"
        CASE "35" : getHmallDlvCode2Name = "우체국"

        CASE "61" : getHmallDlvCode2Name = "경동택배"
        CASE "38" : getHmallDlvCode2Name = "일양택배"
        CASE "70" : getHmallDlvCode2Name = "건영택배"
        CASE "64" : getHmallDlvCode2Name = "천일택배"
        CASE "63" : getHmallDlvCode2Name = "호남택배"
        CASE "69" : getHmallDlvCode2Name = "대신택배"
        CASE "68" : getHmallDlvCode2Name = "합동택배"
        CASE "71" : getHmallDlvCode2Name = "GTX로지스"
        CASE "65" : getHmallDlvCode2Name = "CU POST"
        CASE "29" : getHmallDlvCode2Name = "KGB택배"
        CASE "74" : getHmallDlvCode2Name = "FLF퍼레버택배"
        CASE "60" : getHmallDlvCode2Name = "퀵서비스"

        CASE "16" : getHmallDlvCode2Name = "자가배송"


        ' CASE "1082" : getHmallDlvCode2Name = "기타택배"
        ' CASE "1001" : getHmallDlvCode2Name = "DHL"
        ' CASE "1011" : getHmallDlvCode2Name = "옐로우캡"
        ' CASE "1012" : getHmallDlvCode2Name = "우체국택배EMS"
        ' CASE "1080" : getHmallDlvCode2Name = "KG로지스택배"
        ' CASE "1081" : getHmallDlvCode2Name = "업체직배송"
        ' CASE "1103" : getHmallDlvCode2Name = "한의사랑택배"
        ' CASE "1104" : getHmallDlvCode2Name = "다드림"
        ' CASE "1105" : getHmallDlvCode2Name = "굿투럭"
        ' CASE "1108" : getHmallDlvCode2Name = "CJ대한통운국제특송"
        ' CASE "1109" : getHmallDlvCode2Name = "EMS"
        ' CASE "1110" : getHmallDlvCode2Name = "한덱스"
        ' CASE "1111" : getHmallDlvCode2Name = "FedEx"
        ' CASE "1112" : getHmallDlvCode2Name = "UPS"
        ' CASE "1113" : getHmallDlvCode2Name = "TNT"
        ' CASE "1114" : getHmallDlvCode2Name = "USPS"
        ' CASE "1115" : getHmallDlvCode2Name = "i-parcel"
        ' CASE "1116" : getHmallDlvCode2Name = "GSM NtoN"
        ' CASE "1117" : getHmallDlvCode2Name = "성원글로벌"
        ' CASE "1118" : getHmallDlvCode2Name = "범한판토스"
        ' CASE "1119" : getHmallDlvCode2Name = "ACI Express"
        ' CASE "1121" : getHmallDlvCode2Name = "대운글로벌"
        ' CASE "1122" : getHmallDlvCode2Name = "에어보이익스프레스"
        ' CASE "1123" : getHmallDlvCode2Name = "KGL네트웍스"
        ' CASE "1124" : getHmallDlvCode2Name = "LineExpress"
        ' CASE "1125" : getHmallDlvCode2Name = "2fast익스프레스"
        ' CASE "1126" : getHmallDlvCode2Name = "GSI익스프레스"

        CASE ELSE : getHmallDlvCode2Name = idlvCd
    END SELECT
end function


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
    istyyyymmdd = LEFT(dateadd("d",-5,iedyyyymmdd),10)

Dim strSql : strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'hmall1010','"&confirmDt&"'"
dbget.Execute strSql
rw "초기화작업"

'' 최대 7일간 가능하다. 7*3 =21 일

'' P0:출고대기, P1:출고진행, P2:출고, P3:배송완료
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P0","주문미확인")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P1","주문확인")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P2","출고완료")

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-5,iedyyyymmdd),10)

'' P0:출고대기, P1:출고진행, P2:출고, P3:배송완료
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P0","주문미확인")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P1","주문확인")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P2","출고완료")

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-5,iedyyyymmdd),10)

'' P0:출고대기, P1:출고진행, P2:출고, P3:배송완료
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P0","주문미확인")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P1","주문확인")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P2","출고완료")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-5,iedyyyymmdd),10)

'' P0:출고대기, P1:출고진행, P2:출고, P3:배송완료
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P0","주문미확인")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P1","주문확인")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P2","출고완료")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-5,iedyyyymmdd),10)

'' P0:출고대기, P1:출고진행, P2:출고, P3:배송완료
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P0","주문미확인")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P1","주문확인")  
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P2","출고완료")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-5,iedyyyymmdd),10)

'' P0:출고대기, P1:출고진행, P2:출고, P3:배송완료
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P0","주문미확인")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P1","주문확인")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P2","출고완료")
response.flush

strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'hmall1010','"&confirmDt&"'"
dbget.Execute strSql
rw "주문매핑"

rw "완료"
'response.write("<script>setTimeout(alert('완료'),1000);self.close();</script>")

function Get_HmallOrderListByStatus(stdate,eddate,iorderStatus,istatusName)
	dim sellsite : sellsite = "hmall1010"
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

	Get_HmallOrderListByStatus = False

	'// =======================================================================
	'// 날짜형식
	''selldate = "2017-11-10"
	''xmlSelldate = Replace(selldate, "-", "")

	'// API URL(기간동안의 주문 가져오기)
	xmlURL = "http://xapi.10x10.co.kr:8080/Orders/Hmall/output"


	postParam = ""
	postParam = postParam & "startdate=" & Replace(stdate, "-", "")
	postParam = postParam & "&enddate=" & Replace(Left(DateAdd("d", 1, CDate(eddate)), 10), "-", "")
	postParam = postParam & "&prgrGb="&iorderStatus
	''response.write postParam

    rw "기간검색:"&stdate&"~"&eddate&" 상태:"&iorderStatus&"("&istatusName&")"
	'// =======================================================================
	'// 데이타 가져오기


	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL&"?"&postParam, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
	objXML.send()

	if objXML.Status <> "200" then
		response.write "ERROR : 통신오류" & objXML.Status
		dbget.close : response.end
	end if

    Dim iRbody, strObj, orderCount, obj1


    Dim ordNo, ordItemSeq, orOrdNo, orordItemSeq, shppNo, shppSeq, reOrderYn, delayNts
    Dim cspGoodsCd, goodsCd, uitemId, orderQty, shppDivDtlNm
    Dim  optionContent, shppRsvtDt, whoutCritnDt, autoShortgYn
    Dim  orderStatus, dlvrCd, dlvrNo, dlvrDt, dlvrFinishDt, cancelDt
    Dim paramInfo, retParamInfo, RetErr

    Dim shppTypeDtlNm, delicoVenId, delicoVenNm, wblNo


	iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

        Set strObj = JSON.parse(iRbody)
        orderCount = strObj.count
        If orderCount >0 Then
            'response.write "건수(" & orderCount & ") " & "<br />"
            ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] '"&sellsite&"','"&confirmDt&"'"
            ' dbget.Execute strSql

            set obj1 = strObj.lstorder
            	for i=0 to obj1.length-1
                    orOrdNo="": orordItemSeq=""

                    ordNo           = obj1.get(i).ordNo					        '주문번호
                    ordItemSeq      = obj1.get(i).ordPtcSeq				    '주문일련번호
                    shppNo		    = obj1.get(i).dlvstNo				'배송지시번호
                    shppSeq			= obj1.get(i).dlvstPtcSeq			'배송지시상세번호
                    reOrderYn ="N" ''재주문여부
                    delayNts  =""  ''지연일수
                    cspGoodsCd      = obj1.get(i).venItemCd				'협력사 상품관리코드
                    goodsCd         = obj1.get(i).slitmCd				    '판매상품코드
                    uitemId         = obj1.get(i).uitmCd				'상품속성코드
                    orderQty        = obj1.get(i).ordQty				'주문 수량

                    shppDivDtlNm = ""
                    if (obj1.get(i).dlvTypeGbcd="40") then
                        shppDivDtlNm = "교환출고"
                    end if
                    if (obj1.get(i).dlvCnclNm<>"") then
                        shppDivDtlNm = shppDivDtlNm & CHKIIF(shppDivDtlNm<>"","/","") & obj1.get(i).dlvCnclNm             '배송취소 (''배송구분상세명 (출고/교환출고..))  //nshipTypeGbcd 미출고유형구분코드	출고지연:10, 품절취소:20
                    end if

                    if (shppDivDtlNm = "전체취소") then
                        shppDivDtlNm = "주문취소"
                    elseif (shppDivDtlNm = "교환출고/전체취소") then
                        shppDivDtlNm = "교환출고철회"
                    end if

                    optionContent   = obj1.get(i).uitmTotNm				'상품속성명
                    shppRsvtDt      = ""''예정일
                    whoutCritnDt    = obj1.get(i).lastOshpDlineDt		    '최종출고마감일자  (''출고기준일)
                    autoShortgYn    = "" ''자동결품여부

                    orderStatus     = obj1.get(i).lastDlvstPrgrGbcd		        '최종배송지시진행구분코드 | 25:출고대기, 30:출고진행, 45:출고, 50:배송완료

                    shppTypeDtlNm = ""
                    delicoVenId     = obj1.get(i).dsrvDlvcoCd								'택배배송사코드
                    wblNo           = obj1.get(i).invcNo									'운송장번호
                    delicoVenNm     = getHmallDlvCode2Name(delicoVenId)


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
                    if (whoutCritnDt<>"") then
                        whoutCritnDt = LEFT(whoutCritnDt,4)&"-"&MID(whoutCritnDt,5,2)&"-"&RIGHT(whoutCritnDt,2)
                    end if


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

                        ,Array("@shppTypeDtlNm"		    , adVarchar		, adParamInput		,   16, Trim(shppTypeDtlNm)) _
                        ,Array("@delicoVenId"		    , adVarchar		, adParamInput		,   16, Trim(delicoVenId)) _
                        ,Array("@delicoVenNm"		    , adVarchar		, adParamInput		,   32, Trim(delicoVenNm)) _
                        ,Array("@wblNo"		            , adVarchar		, adParamInput		,   32, Trim(wblNo)) _
                        ,Array("@invoiceUpDt"		    , adVarchar		, adParamInput		,   19, Trim("")) _
                        ,Array("@outjFixedDt"		    , adVarchar		, adParamInput		,   19, Trim("")) _

                        ,Array("@OrgOutMallOrderSerial"	, adVarchar		, adParamInput		,   32, Trim(orordNo)) _
                        ,Array("@OrgOrgDetailKey"		, adVarchar		, adParamInput		,   32, Trim(orordItemSeq)) _
                    )

                    'On Error RESUME Next
                    retParamInfo = fnExecSPOutput(sqlStr, paramInfo)
                    RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드

                     successCnt = successCnt+1
                Next
            set obj1 = nothing
        End If
    Set strObj = nothing

    '' 주문번호 매핑.
    ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] '"&sellsite&"','"&confirmDt&"'"
    ' dbget.Execute strSql

    rw "건수:"&successCnt&"======================================"

	Get_HmallOrderListByStatus = True

end function
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
