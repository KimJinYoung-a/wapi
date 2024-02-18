<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp" -->
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
function getEzWelDlvCode2Name(idlvCd)
    if isNULL(idlvCd) then Exit function

    SELECT CASE idlvCd
        CASE "1007" : getEzWelDlvCode2Name = "CJ대한통운"
        CASE "1017" : getEzWelDlvCode2Name = "롯데택배"
        CASE "1016" : getEzWelDlvCode2Name = "한진택배"
        CASE "1008" : getEzWelDlvCode2Name = "로젠택배"
        CASE "1161" : getEzWelDlvCode2Name = "우편등기"

        CASE "1180" : getEzWelDlvCode2Name = "일양로지스"
        CASE "1163" : getEzWelDlvCode2Name = "이노지스"
        CASE "1200" : getEzWelDlvCode2Name = "대신택배"
        CASE "1082" : getEzWelDlvCode2Name = "기타택배"
        CASE "1001" : getEzWelDlvCode2Name = "DHL"
        CASE "1002" : getEzWelDlvCode2Name = "KGB택배"
        CASE "1005" : getEzWelDlvCode2Name = "경동택배"
        CASE "1011" : getEzWelDlvCode2Name = "옐로우캡"
        CASE "1012" : getEzWelDlvCode2Name = "우체국택배EMS"
        CASE "1014" : getEzWelDlvCode2Name = "천일택배"
        CASE "1080" : getEzWelDlvCode2Name = "KG로지스택배"
        CASE "1081" : getEzWelDlvCode2Name = "업체직배송"
        CASE "1260" : getEzWelDlvCode2Name = "GTX로지스"
        
        CASE "1102" : getEzWelDlvCode2Name = "합동택배"
        CASE "1103" : getEzWelDlvCode2Name = "한의사랑택배"
        CASE "1104" : getEzWelDlvCode2Name = "다드림"
        CASE "1105" : getEzWelDlvCode2Name = "굿투럭"
        CASE "1106" : getEzWelDlvCode2Name = "건영택배"
        CASE "1107" : getEzWelDlvCode2Name = "호남택배"
        CASE "1108" : getEzWelDlvCode2Name = "CJ대한통운국제특송"
        CASE "1109" : getEzWelDlvCode2Name = "EMS"
        CASE "1110" : getEzWelDlvCode2Name = "한덱스"
        CASE "1111" : getEzWelDlvCode2Name = "FedEx"
        CASE "1112" : getEzWelDlvCode2Name = "UPS"
        CASE "1113" : getEzWelDlvCode2Name = "TNT"
        CASE "1114" : getEzWelDlvCode2Name = "USPS"
        CASE "1115" : getEzWelDlvCode2Name = "i-parcel"
        CASE "1116" : getEzWelDlvCode2Name = "GSM NtoN"
        CASE "1117" : getEzWelDlvCode2Name = "성원글로벌"
        CASE "1118" : getEzWelDlvCode2Name = "범한판토스"
        CASE "1119" : getEzWelDlvCode2Name = "ACI Express"
        CASE "1121" : getEzWelDlvCode2Name = "대운글로벌"
        CASE "1122" : getEzWelDlvCode2Name = "에어보이익스프레스"
        CASE "1123" : getEzWelDlvCode2Name = "KGL네트웍스"
        CASE "1124" : getEzWelDlvCode2Name = "LineExpress"
        CASE "1125" : getEzWelDlvCode2Name = "2fast익스프레스"
        CASE "1126" : getEzWelDlvCode2Name = "GSI익스프레스"
        CASE "1240" : getEzWelDlvCode2Name = "편의점택배"
        CASE ELSE : getEzWelDlvCode2Name =""
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
    istyyyymmdd = LEFT(dateadd("d",-14,iedyyyymmdd),10)

 

CONST cspCd		= "10040413"							'CP업체코드(이지웰 발급)
CONST crtCd		= "8e5a6dbdd27efb49fc600c293884ef47"	'보안코드(이지웰 발급)

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'ezwel','"&confirmDt&"'"
dbget.Execute sqlStr
rw "초기화작업"

'' 1001:주문완료 / 1002:출고준비중 / 1003:배송중 / 1004:수취완료 / 1005:주문취소 / 1007:반품요청 ....
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1001","주문미확인")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1002","주문확인")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1003","배송중")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-14,iedyyyymmdd),10)

call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1001","주문미확인")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1002","주문확인")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1003","배송중")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)

call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1001","주문미확인")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1002","주문확인")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1003","배송중")
response.flush

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'ezwel','"&confirmDt&"'"
dbget.Execute sqlStr
rw "주문매핑"

rw "완료"

' 일별정산시 이용할 수 있을듯함.
' call Get_ezwelOrderListByStatus("2019-10-25","2019-10-25","1004","수취완료")
' response.flush
'response.write("<script>setTimeout(alert('완료'),1000);self.close();</script>")

function Get_ezwelOrderListByStatus(stdate,eddate,iorderStatus,istatusName)
	dim sellsite : sellsite = "ezwel"
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

	Get_ezwelOrderListByStatus = False

	'// =======================================================================
	'// 날짜형식
	''selldate = "2017-11-10"
	''xmlSelldate = Replace(selldate, "-", "")

	'// API URL(기간동안의 주문 가져오기)
	xmlURL = "http://api.ezwel.com/if/api/orderListAPI.ez"
	''response.write xmlURL

	postParam = "cspCd=" & cspCd & "&crtCd=" & crtCd
	postParam = postParam & "&startDate=" & Replace(stdate, "-", "") & "000000"
	postParam = postParam & "&endDate=" & Replace(Left(DateAdd("d", 1, CDate(eddate)), 10), "-", "") & "000000"
	postParam = postParam & "&orderStatus="&iorderStatus
	''response.write postParam

    rw "기간검색:"&stdate&"~"&eddate&" 상태:"&iorderStatus&"("&istatusName&")"
	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
	objXML.send(postParam)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

    Dim ordNo, ordItemSeq, shppNo, shppSeq, reOrderYn, delayNts
    Dim cspGoodsCd, goodsCd, uitemId, orderQty, shppDivDtlNm
    Dim  optionContent, shppRsvtDt, whoutCritnDt, autoShortgYn
    Dim  orderStatus, dlvrCd, dlvrNo, dlvrDt, dlvrFinishDt, cancelDt
    Dim paramInfo, retParamInfo, RetErr

    Dim shppTypeDtlNm, delicoVenId, delicoVenNm, wblNo

	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False

    if (application("Svr_Info")="Dev") then

        bufStr = "<?xml version='1.0' encoding='EUC-KR'?>"
        bufStr = bufStr & "<resultSet class='array'><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>8640</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>1234873</cspGoodsCd><dccpnPrice type='string'>0</dccpnPrice><dlvrCd class='object' null='true'/><dlvrDt class='object' null='true'/><dlvrFinishDt class='object' null='true'/><dlvrNo class='object' null='true'/><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>0</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1016905554</goodsCd><goodsNm type='string'><![CDATA[[텐바이텐]헬로키티 다용도대야대 딸기]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>12100</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[:^]]></optionContent><orderGoodsNum type='number'>6</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1002</orderStatus><salePrice type='number'>9600</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053458145</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191018181827</orderDt><orderNum type='string'>1026227679</orderNum><orderReqContent type='string'><![CDATA[]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[경기도 안산시 단원구 적금로 164 (고잔동, 고잔롯데캐슬골드파크)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[108동 2102호]]></rcvrAddr2><rcvrMobile type='string'>010-9648-0820</rcvrMobile><rcvrNm type='string'>김완?</rcvrNm><rcvrPost type='string'>15347</rcvrPost><rcvrTelNum type='string'>02--</rcvrTelNum><sndMobile type='string'>010-9648-0820</sndMobile><sndNm type='string'>김완택</sndNm><sndTelNum type='string'>02-</sndTelNum></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>177300</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>2376775</cspGoodsCd><dccpnPrice type='string'>0</dccpnPrice><dlvrCd class='object' null='true'/><dlvrDt class='object' null='true'/><dlvrFinishDt class='object' null='true'/><dlvrNo class='object' null='true'/><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>0</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1019145570</goodsCd><goodsNm type='string'><![CDATA[[텐바이텐]라자가구 오브 쿠페 800 전신거울 수납형A 옷장 NA8744]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>323000</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[선택:그레이^]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1002</orderStatus><salePrice type='number'>197000</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053491658</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191020211957</orderDt><orderNum type='string'>1026244168</orderNum><orderReqContent type='string'><![CDATA[배송전 미리 연락 바랍니다. ]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[경기도 안산시 단원구 당곡2로 29 (고잔동, 주공8단지아파트)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[801동 606호]]></rcvrAddr2><rcvrMobile type='string'>010-7270-9109</rcvrMobile><rcvrNm type='string'>황광현</rcvrNm><rcvrPost type='string'>15338</rcvrPost><rcvrTelNum type='string'>02--</rcvrTelNum><sndMobile type='string'>010-7270-9109</sndMobile><sndNm type='string'>황광현</sndNm><sndTelNum type='string'>02-</sndTelNum></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>60300</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>1784132</cspGoodsCd><dccpnPrice type='string'>0</dccpnPrice><dlvrCd class='object' null='true'/><dlvrDt class='object' null='true'/><dlvrFinishDt class='object' null='true'/><dlvrNo class='object' null='true'/><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>0</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1015413316</goodsCd><goodsNm type='string'><![CDATA[[텐바이텐]밀크스타 자수 암막커튼]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>67000</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[:^]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>2</orderQty><orderStatus type='string'>1002</orderStatus><salePrice type='number'>67000</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053494255</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191020230642</orderDt><orderNum type='string'>1026245580</orderNum><orderReqContent type='string'><![CDATA[현관앞배송해주세요 ]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[04355]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[서울특별시 용산구 백범로 250 (효창동, 세양청마루아파트) 101동 803호]]></rcvrAddr2><rcvrMobile type='string'>010-3023-5688</rcvrMobile><rcvrNm type='string'>김경희</rcvrNm><rcvrPost type='string'>04355</rcvrPost><rcvrTelNum type='string'>02--</rcvrTelNum><sndMobile type='string'>010-3023-5688</sndMobile><sndNm type='string'>김경희</sndNm><sndTelNum type='string'>-</sndTelNum></arrOrderList><orderCnt type='number'>3</orderCnt><resultCode type='string'>200</resultCode><resultMsg type='string'>성공</resultMsg></resultSet>"
        
        bufStr = "<?xml version='1.0' encoding='EUC-KR'?>"
        bufStr = bufStr & "<resultSet class='array'><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>128160</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>1738833</cspGoodsCd><dccpnPrice class='object' null='true'/><dlvrCd type='string'>1007</dlvrCd><dlvrDt type='string'>20191010161501</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>625820114020</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>2500</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1015247678</goodsCd><goodsNm type='string'><![CDATA[[텐바이텐]목화솜 면 홑겹 이불 퀸 (Q)]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>0</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[선택:베개세트,솜샤시추가안함^]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>142400</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053219404</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191007161131</orderDt><orderNum type='string'>1026118893</orderNum><orderReqContent type='string'><![CDATA[]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[서울특별시 강남구  광평로19길 15 (일원동, 목련타운아파트)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[109동 1403호]]></rcvrAddr2><rcvrMobile type='string'>010-8459-3620</rcvrMobile><rcvrNm type='string'>성경자</rcvrNm><rcvrPost type='string'>06355</rcvrPost><rcvrTelNum class='object' null='true'/><sndMobile type='string'>010-8459-3620</sndMobile><sndNm type='string'>성경자</sndNm><sndTelNum type='string'>02-459-3620</sndTelNum></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>7740</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>2468588</cspGoodsCd><dccpnPrice type='string'>0</dccpnPrice><dlvrCd type='string'>1007</dlvrCd><dlvrDt type='string'>20191011154502</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>625902859056</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>0</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1019787277</goodsCd><goodsNm type='string'><![CDATA[[텐바이텐][젤시스슬라임] 사과수 / 지글리슬라임 / 230ml]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>0</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[:^]]></optionContent><orderGoodsNum type='number'>2</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>8600</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053227540</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191007215851</orderDt><orderNum type='string'>1026122245</orderNum><orderReqContent type='string'><![CDATA[무인택배함에 넣어주세요.]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[인천광역시 연수구 송도과학로27번길 55 (송도동, 롯데캐슬 캠퍼스타운)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[101-1504]]></rcvrAddr2><rcvrMobile type='string'>010-5573-7371</rcvrMobile><rcvrNm type='string'>박지혜</rcvrNm><rcvrPost type='string'>21982</rcvrPost><rcvrTelNum type='string'>02-0000-0000</rcvrTelNum><sndMobile type='string'>010-5573-7371</sndMobile><sndNm type='string'>박지혜</sndNm><sndTelNum class='object' null='true'/></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>41310</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>2400003</cspGoodsCd><dccpnPrice class='object' null='true'/><dlvrCd type='string'>1007</dlvrCd><dlvrDt type='string'>20191010141504</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>625868717385</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>2500</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1019261135</goodsCd><goodsNm type='string'><![CDATA[[텐바이텐]모던 심플 전자렌지선반(블랙,화이트)_(2172003)]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>0</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[선택:화이트^]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>45900</salePrice><useAmt type='string'>2300</useAmt></arrOrderGoods><aspOrderNum type='string'>1053264243</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191009175555</orderDt><orderNum type='string'>1026137775</orderNum><orderReqContent type='string'><![CDATA[]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[강원도 춘천시 후만로24번길 13 (후평동, 현대5차아파트)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[501동 414호]]></rcvrAddr2><rcvrMobile type='string'>010-9086-9004</rcvrMobile><rcvrNm type='string'>김석현</rcvrNm><rcvrPost type='string'>24285</rcvrPost><rcvrTelNum class='object' null='true'/><sndMobile type='string'>010-9086-9004</sndMobile><sndNm type='string'>김석현</sndNm><sndTelNum class='object' null='true'/></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>18000</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>1392969</cspGoodsCd><dccpnPrice class='object' null='true'/><dlvrCd type='string'>1007</dlvrCd><dlvrDt type='string'>20191011104503</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>625871366653</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>2500</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1016258682</goodsCd><goodsNm type='string'><![CDATA[[텐바이텐]브리니클(Brinicle) 풀 스테인레스 304 마늘다지기/ 마늘분쇄기]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>0</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>20000</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053267154</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191009203155</orderDt><orderNum type='string'>1026139025</orderNum><orderReqContent type='string'><![CDATA[]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[인천광역시 서구 원당대로840번길 21 (당하동, 풍림아이원아파트)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[824동1403호]]></rcvrAddr2><rcvrMobile type='string'>010-2602-4682</rcvrMobile><rcvrNm type='string'>고건</rcvrNm><rcvrPost type='string'>22682</rcvrPost><rcvrTelNum class='object' null='true'/><sndMobile type='string'>010-2602-4682</sndMobile><sndNm type='string'>고건</sndNm><sndTelNum class='object' null='true'/></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>29700</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>2277613</cspGoodsCd><dccpnPrice type='string'>0</dccpnPrice><dlvrCd type='string'>1008</dlvrCd><dlvrDt type='string'>20191016124501</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>95184514994</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>2500</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1018868242</goodsCd><goodsNm type='string'><![CDATA[[텐바이텐]딩고 방수앞치마]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>33000</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[선택:핑크^]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>33000</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053289051</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191010181801</orderDt><orderNum type='string'>1026148586</orderNum><orderReqContent type='string'><![CDATA[]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[경기도 화성시 융건로 99 (기안동, 풍성신미주아파트)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[114-402]]></rcvrAddr2><rcvrMobile type='string'>010-8917-5627</rcvrMobile><rcvrNm type='string'>최수지</rcvrNm><rcvrPost type='string'>18342</rcvrPost><rcvrTelNum type='string'>02-8917-5627</rcvrTelNum><sndMobile type='string'>010-8917-5627</sndMobile><sndNm type='string'>최수지</sndNm><sndTelNum type='string'>010-8917-5627</sndTelNum></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>40500</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>2189217</cspGoodsCd><dccpnPrice class='object' null='true'/><dlvrCd type='string'>1007</dlvrCd><dlvrDt type='string'>20191018202205</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>625927013676</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>2500</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1018127841</goodsCd><goodsNm type='string'><![CDATA[[텐바이텐]프라이머 UV코팅 암막 도어커튼 95X210]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>60000</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>45000</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053346247</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191014000135</orderDt><orderNum type='string'>1026174793</orderNum><orderReqContent type='string'><![CDATA[100cm (가로) x 150cm (세로) /아일렛]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[서울특별시 동작구 사당로16길 131-5 (사당동)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[301호]]></rcvrAddr2><rcvrMobile type='string'>010-9948-8313</rcvrMobile><rcvrNm type='string'>김영선</rcvrNm><rcvrPost type='string'>07022</rcvrPost><rcvrTelNum class='object' null='true'/><sndMobile type='string'>010-9948-8313</sndMobile><sndNm type='string'>김영선</sndNm><sndTelNum type='string'>010-9948-8313</sndTelNum></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>16200</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>1649904</cspGoodsCd><dccpnPrice class='object' null='true'/><dlvrCd type='string'>1008</dlvrCd><dlvrDt type='string'>20191015164603</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>95202908673</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>0</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1016673436</goodsCd><goodsNm type='string'><![CDATA[[텐바이텐]솔리드 멜란 실내화 - 베이지_(1860391)]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>18000</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[]]></optionContent><orderGoodsNum type='number'>3</orderGoodsNum><orderQty type='number'>2</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>18000</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053369946</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191014235220</orderDt><orderNum type='string'>1026185859</orderNum><orderReqContent type='string'><![CDATA[부재시 문앞에 두고가주세요~]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[경기도 광주시 초월읍 경충대로 1023-10]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[백합동 303호]]></rcvrAddr2><rcvrMobile type='string'>010-8915-9103</rcvrMobile><rcvrNm type='string'>원유진</rcvrNm><rcvrPost type='string'>12736</rcvrPost><rcvrTelNum class='object' null='true'/><sndMobile type='string'>010-8915-9103</sndMobile><sndNm type='string'>원유진</sndNm><sndTelNum class='object' null='true'/></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>24120</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>1997414</cspGoodsCd><dccpnPrice type='string'>0</dccpnPrice><dlvrCd type='string'>1014</dlvrCd><dlvrDt type='string'>20191021144504</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>51969716790</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>0</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1017342754</goodsCd><goodsNm type='string'><![CDATA[[텐바이텐][이노센트] 리코 300 전신 거치 거울_(1237091)]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>29000</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[선택:화이트^]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>26800</salePrice><useAmt type='string'>1000</useAmt></arrOrderGoods><aspOrderNum type='string'>1053481532</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191020111041</orderDt><orderNum type='string'>1026239132</orderNum><orderReqContent type='string'><![CDATA[문앞에 놔주세요]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[전라남도 순천시 풍덕새길 22 (풍덕동)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[리원빌 201호]]></rcvrAddr2><rcvrMobile type='string'>010-6669-6416</rcvrMobile><rcvrNm type='string'>김민휘</rcvrNm><rcvrPost type='string'>57995</rcvrPost><rcvrTelNum type='string'>02--</rcvrTelNum><sndMobile type='string'>010-6669-6416</sndMobile><sndNm type='string'>김민휘</sndNm><sndTelNum type='string'>-</sndTelNum></arrOrderList><orderCnt type='number'>8</orderCnt><resultCode type='string'>200</resultCode><resultMsg type='string'>성공</resultMsg></resultSet>"


        xmlDOM.LoadXML(bufStr)

    else
	    xmlDOM.LoadXML(objXML.responseText)
    end if
'response.write "<textarea cols='40' rows=10'>"&objXML.responseText & "</textarea><br /><br />"


	if (xmlDOM.getElementsByTagName("resultSet/arrOrderList").length < 1) then
		''if IsAutoScript then
			response.write "내역없음 : 종료" & "<br />"
		''end if

        if (xmlDOM.getElementsByTagName("resultSet/resultMsg").length>0) then
            rw "resultMsg:"&xmlDOM.getElementsByTagName("resultSet/resultMsg")(0).Text
        end if

		Get_ezwelOrderListByStatus = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	else
		response.write "건수(" & xmlDOM.getElementsByTagName("resultSet/arrOrderList").length & ") " & "<br />"

        ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'ezwel','"&confirmDt&"'"
        ' dbget.Execute strSql
	end if

	set objMasterListXML = xmlDOM.getElementsByTagName("resultSet/arrOrderList")
	For each objMasterOneXML in objMasterListXML
        
        
		set objDetailListXML = objMasterOneXML.getElementsByTagName("arrOrderGoods")
		For each objDetailOneXML in objDetailListXML
            ordNo = objMasterOneXML.getElementsByTagName("orderNum")(0).Text      ''주문번호
            ordItemSeq = objDetailOneXML.getElementsByTagName("orderGoodsNum")(0).Text  ''주문순번
            shppNo = "" ''배송번호
            shppSeq ="" ''배송Seq
            reOrderYn ="N" ''재주문여부 
            delayNts  =""  ''지연일수
            cspGoodsCd = objDetailOneXML.getElementsByTagName("cspGoodsCd")(0).Text   ''업체상품코드
            goodsCd = objDetailOneXML.getElementsByTagName("goodsCd")(0).Text         ''(제휴)상품코드
            uitemId = "" ''제휴 단품ID
            orderQty = objDetailOneXML.getElementsByTagName("orderQty")(0).Text       ''주문수량
            shppDivDtlNm = "" ''배송구분상세명 (출고/교환출고..)
            optionContent = objDetailOneXML.getElementsByTagName("optionContent")(0).Text        ''옵션명 */arrOrderGoods/optionContent
            shppRsvtDt = ""  ''예정일?
            whoutCritnDt ="" ''출고기준일
            autoShortgYn ="" ''자동결품여부

            orderStatus = objDetailOneXML.getElementsByTagName("orderStatus")(0).Text         ''주문상태
            dlvrCd = objDetailOneXML.getElementsByTagName("dlvrCd")(0).Text         ''택배사
            dlvrNo = objDetailOneXML.getElementsByTagName("dlvrNo")(0).Text         ''송장번호
            dlvrDt = objDetailOneXML.getElementsByTagName("dlvrDt")(0).Text         ''배송일
            dlvrFinishDt = objDetailOneXML.getElementsByTagName("dlvrFinishDt")(0).Text   ''배송완료일
            cancelDt = objDetailOneXML.getElementsByTagName("cancelDt")(0).Text       ''취소일

            bufStr = ""
            bufStr = sellsite&"|"&ordNo
            bufStr = bufStr &"|"&ordItemSeq
            bufStr = bufStr &"|"&cspGoodsCd
            bufStr = bufStr &"|"&goodsCd
            
            bufStr = bufStr &"|"&orderQty

            bufStr = bufStr &"|"&orderStatus
            bufStr = bufStr &"|"&dlvrCd
			bufStr = bufStr &"|"&dlvrNo

            bufStr = bufStr &"|"&dlvrDt
            bufStr = bufStr &"|"&dlvrFinishDt
            bufStr = bufStr &"|"&cancelDt

            shppTypeDtlNm = ""
            delicoVenId   = dlvrCd
            delicoVenNm   = getEzWelDlvCode2Name(dlvrCd)
            wblNo         = dlvrNo

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

                ,Array("@invoiceUpDt"	    , adVarchar		, adParamInput		,   19, "") _
                ,Array("@outjFixedDt"		, adVarchar		, adParamInput		,   19, Trim(dlvrFinishDt)) _
            )

            'On Error RESUME Next
            retParamInfo = fnExecSPOutput(sqlStr, paramInfo)
            RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
		next

	next

    '' 주문번호 매핑.
    ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'ezwel','"&confirmDt&"'"
    ' dbget.Execute strSql

	''if IsAutoScript then
		'response.write "주문입력(" & successCnt & ")" & "<br />"
	''end if

	Get_ezwelOrderListByStatus = True
	Set xmlDOM = Nothing
	Set objXML = Nothing

end function
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->