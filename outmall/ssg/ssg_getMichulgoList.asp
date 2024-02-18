<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/ssg/ssgItemcls.asp"-->
<script language="javascript" runat="server">
var confirmDt = (new Date()).valueOf();
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style>
body {
  font-size: small;
}
</style>
</head>
<body bgcolor="#F4F4F4" >
<%

'' TLS 1.2를 지원하지 않는 서버가 있는듯함..
const Option_TLS12 = 2048
const Option_TLS1 = 512
const Option_TLS = 128

'' 1. 배송지시목록조회
'' 2. 주문 확인 처리
'' 3. 촐고 대상목록 조회

Dim sqlStr
Dim isOnlyTodayBaljuView : isOnlyTodayBaljuView=false  ''발주된내역만 (주문수신 view)
Dim isDlvConfirmProc 	 : isDlvConfirmProc=false  ''주문확인 Proc
Dim isDlvInputProc 	 	 : isDlvInputProc=false    ''입력 Proc
Dim eddt : eddt=requestCheckvar(request("eddt"),10)

Dim IS_TEST_MODE : IS_TEST_MODE=FALSE '' SaveOrderToDB

Dim istyyyymmdd, iedyyyymmdd
    iedyyyymmdd = LEFT(dateadd("d",-2,now()),10)
    if eddt<>"" then
        if isDate(eddt) then
            iedyyyymmdd=eddt  '''yyyymmdd
        end if
    end if

    istyyyymmdd = dateadd("d",-7,iedyyyymmdd)

''초기화.
sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'ssg','"&confirmDt&"'"
dbget.Execute sqlStr
rw "초기화작업"

''isOnlyTodayBaljuView = True

call getSsgDlvReqList(istyyyymmdd,iedyyyymmdd)  ''배송지시목록
response.flush
call getSsgDlvConfirmList(istyyyymmdd,iedyyyymmdd)  ''주문확인(송장입력)
response.flush
call getSsglistNonDelivery(istyyyymmdd,iedyyyymmdd) ''(출고완료)
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)

call getSsgDlvReqList(istyyyymmdd,iedyyyymmdd)  ''배송지시목록
response.flush
call getSsgDlvConfirmList(istyyyymmdd,iedyyyymmdd)  ''주문확인(송장입력)
response.flush
call getSsglistNonDelivery(istyyyymmdd,iedyyyymmdd) ''(출고완료)
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)

call getSsgDlvReqList(istyyyymmdd,iedyyyymmdd)  ''배송지시목록
response.flush
call getSsgDlvConfirmList(istyyyymmdd,iedyyyymmdd)  ''주문확인(송장입력)
response.flush
call getSsglistNonDelivery(istyyyymmdd,iedyyyymmdd) ''(출고완료)
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)

call getSsgDlvReqList(istyyyymmdd,iedyyyymmdd)  ''배송지시목록
response.flush
call getSsgDlvConfirmList(istyyyymmdd,iedyyyymmdd)  ''주문확인(송장입력)
response.flush
call getSsglistNonDelivery(istyyyymmdd,iedyyyymmdd) ''(출고완료)
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)

call getSsgDlvReqList(istyyyymmdd,iedyyyymmdd)  ''배송지시목록
response.flush
call getSsgDlvConfirmList(istyyyymmdd,iedyyyymmdd)  ''주문확인(송장입력)
response.flush
call getSsglistNonDelivery(istyyyymmdd,iedyyyymmdd) ''(출고완료)
response.flush

 '' 주문번호 매핑.
sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'ssg','"&confirmDt&"'"
dbget.Execute sqlStr
rw "주문매핑"

rw "완료"
'response.write("<script>setTimeout(alert('완료'),1000);self.close();</script>")

''출고되었으나 미배송
public function getSsglistNonDelivery(byVal styyyymmdd,byVal edyyyymmdd)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    rw "기간검색:"&styyyymmdd&"~"&edyyyymmdd&" 상태:"&"출고완료"
    styyyymmdd = replace(styyyymmdd,"-","")
    edyyyymmdd = replace(edyyyymmdd,"-","")

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listNonDelivery.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestNonDelivery>"
    requestBody = requestBoDy&"<perdType>01</perdType>"  ''01출고완료일, 02결제완료일
    requestBody = requestBoDy&"<perdStrDts>"&styyyymmdd&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&edyyyymmdd&"</perdEndDts>"  ''하루를 더해야?
    requestBody = requestBoDy&"</requestNonDelivery>"
	objXML.send(requestBody)

    if (isOnlyTodayBaljuView) then
        response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"
      ' response.end
    end if

    Dim successCnt : successCnt=0

    Dim shppNo,shppSeq,shppTabProgStatCd,evntSeq,shppDivDtlCd,shppDivDtlNm,reOrderYn,delayNts,ordNo,ordItemSeq,ordCmplDts
    Dim lastShppProgStatDtlNm,lastShppProgStatDtlCd,salestrNo,shppVenId,shppVenNm,shppTypeNm,shppTypeCd,shppTypeDtlCd,shppTypeDtlNm,boxNo
    Dim shppcst,shppcstCodYn,itemNm,splVenItemId,itemId,uitemId,dircItemQty,cnclItemQty,ordQty,sellprc,frgShppYn
    Dim ordpeNm,rcptpeNm,rcptpeHpno,rcptpeTelno,shpplocAddr,shpplocZipcd,shpplocOldZipcd,shpplocRoadAddr,itemChrctDivCd,shppStatCd,shppStatNm
    Dim orordNo,orordItemSeq,shppMainCd,siteNo,siteNm,shppRsvtDt,splprc,shortgYn,newWblNoData,newRow,itemDiv
    Dim shpplocBascAddr,shpplocDtlAddr,ordItemDivNm
    Dim ordpeHpno, ordMemoCntt, pCus, frebieNm ,shortgProgStatCd, shortgProgStatNm, uitemNm
    Dim iBufrequireDetail, procItemQty

    Dim whoutCritnDt, autoShortgYn
    Dim delicoVenId ''택배사ID
    Dim delicoVenNm	''택배사명
    Dim wblNo	    ''운송장번호

    dim retBody : retBody=objXML.responseText
    Dim paramInfo, RetparamInfo, RetErr
    retBody = replace(retBody,"&","")
	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(retBody) ''objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text


			Set LagrgeNode = xmlDOM.SelectNodes("/result/nonDeliverys/nonDelivery")
			If Not (LagrgeNode Is Nothing) Then
                ''초기화(기간별)
                ' if (LagrgeNode.length>0) then
                '      strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'ssg','"&confirmDt&"'"
                '      dbget.Execute strSql
                ' end if

			    For i = 0 To LagrgeNode.length - 1
			        ''변수초기화.
			        shppNo ="": shppSeq = "": shppTabProgStatCd ="": evntSeq ="": shppDivDtlCd =""
                    shppDivDtlNm ="": reOrderYn ="": delayNts ="": ordNo ="": ordItemSeq =""
                    ordCmplDts ="": lastShppProgStatDtlNm = "": lastShppProgStatDtlCd ="": salestrNo ="": shppVenId =""
                    shppVenNm ="": shppTypeNm ="": shppTypeCd ="": shppTypeDtlCd ="": shppTypeDtlNm =""
                    delicoVenId ="": boxNo ="": shppcst ="": shppcstCodYn ="": itemNm =""
                    splVenItemId ="":itemId ="":uitemId ="": dircItemQty ="": cnclItemQty =""
                    ordQty ="" :sellprc ="": frgShppYn ="": ordpeNm =""
                    rcptpeNm ="" :rcptpeHpno ="": rcptpeTelno ="": shpplocAddr =""
                    shpplocZipcd ="": shpplocOldZipcd ="": shpplocRoadAddr ="": itemChrctDivCd =""
                    shppStatCd ="": shppStatNm ="": orordNo ="": orordItemSeq ="": shppMainCd =""
                    siteNo ="": siteNm ="": shppRsvtDt ="": splprc ="": shortgYn =""
                    newWblNoData ="": newRow ="": itemDiv ="": shpplocBascAddr ="": shpplocDtlAddr ="": ordItemDivNm =""

                    ordpeHpno = "": ordMemoCntt = "": pCus = "": frebieNm = "": shortgProgStatCd ="": shortgProgStatNm ="" : uitemNm=""
                    iBufrequireDetail = ""
                    whoutCritnDt =""
                    delicoVenNm	="" ''택배사명
                    wblNo	    ="" ''운송장번호


			        shppNo              = LagrgeNode(i).SelectSingleNode("shppNo").Text                 ''*배송번호
                    shppSeq             = LagrgeNode(i).SelectSingleNode("shppSeq").Text                ''*배송순번
                    ' shppTabProgStatCd   = LagrgeNode(i).SelectSingleNode("shppTabProgStatCd").Text      ''최종배송상세진행상태코드(배송단위) 11 배송지시 21 피킹지시 22 피킹완료 31 패킹완료 41 출고보류 42 출고지연 43 출고완료 51 배송완료 52 배송거절
					' if NOT (LagrgeNode(i).SelectSingleNode("evntSeq") is Nothing) then
                    ' 	evntSeq             = LagrgeNode(i).SelectSingleNode("evntSeq").Text                ''이벤트순번
					' end if
                    shppDivDtlCd        = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text           ''*배송구분상세코드 11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고
                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''배송구분상세명
                    reOrderYn           = LagrgeNode(i).SelectSingleNode("reOrderYn").Text              ''*재지시여부구분
                    delayNts            = "" ''LagrgeNode(i).SelectSingleNode("delayNts").Text               ''지연횟수
                    ordNo               = LagrgeNode(i).SelectSingleNode("ordNo").Text                  ''*주문번호 [20171123128379]
                    ordItemSeq          = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text             ''*주문순번
                    ordCmplDts          = LagrgeNode(i).SelectSingleNode("ordCmplDts").Text             ''*주문완료일시 [2017-11-23 10:39:42.0]
                    lastShppProgStatDtlNm   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlNm").Text  ''최종배송상세진행상태명(배송상품단위) [피킹완료]
                    lastShppProgStatDtlCd   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''최종배송상세진행상태코드(배송상품단위) 11 배송지시 21 피킹지시 22 피킹완료 31 패킹완료 41 출고보류 42 출고지연 43 출고완료 51 배송완료 52 배송거절
                    'salestrNo           = LagrgeNode(i).SelectSingleNode("salestrNo").Text              '' [6004]
                    'shppVenId           = LagrgeNode(i).SelectSingleNode("shppVenId").Text      ''공급업체아이디 [0000003198]
                    'shppVenNm           = LagrgeNode(i).SelectSingleNode("shppVenNm").Text      ''공급업체명
                    shppTypeNm          = LagrgeNode(i).SelectSingleNode("shppTypeNm").Text     ''배송유형명    [택배배송]
                    shppTypeCd          = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text     ''배송유형코드 10 자사배송 20 택배배송 30 매장방문 40 등기 50 미배송 60 미발송
                    shppTypeDtlCd       = LagrgeNode(i).SelectSingleNode("shppTypeDtlCd").Text  ''배송유형상세코드 14 업체자사배송 22 업체택배배송 25 해외택배배송 31 매장방문 41 등기 51 SMS 52 EMAIL 61 미발송
                    shppTypeDtlNm       = LagrgeNode(i).SelectSingleNode("shppTypeDtlNm").Text  ''배송유형상세명 [업체택배배송]

                    if NOT (LagrgeNode(i).SelectSingleNode("delicoVenId") is Nothing) then
                        delicoVenId         = LagrgeNode(i).SelectSingleNode("delicoVenId").Text    ''택배사ID [0000033011]
                    end if
                    'boxNo               = LagrgeNode(i).SelectSingleNode("boxNo").Text          ''박스번호 [398327952]
                    'shppcst             = LagrgeNode(i).SelectSingleNode("shppcst").Text        '' 배송비? [303] ??
                    'shppcstCodYn        = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text   ''*배송비 착불여부 Y: 착불 N: 선불
                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*상품명
                    itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*상품번호 [1000024811163]

					If NOT (LagrgeNode(i).SelectSingleNode("splVenItemId") is Nothing) Then
	    			    splVenItemId    = LagrgeNode(i).SelectSingleNode("splVenItemId").Text        ''업체상품번호 [1024019]
					Else
						strSql = ""
						strSql = strSql & " select top 1 itemid "
						strSql = strSql & " FROM db_etcmall.dbo.tbl_ssg_regitem"
						strSql = strSql & " where ssgGoodNo = '"& itemId &"' "
						rsget.CursorLocation = adUseClient
						rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
						If Not rsget.Eof Then
							splVenItemId = rsget("itemid")
						Else
							rw "오류 주문번호 : " & ordNo
						End If
						rsget.Close

						If session("ssBctID")="kjy8517" Then
							response.write "<textarea cols=100 rows=30>**"&objXML.responseText&"</textarea>"
						End If
					End If

                    'uitemId             = LagrgeNode(i).SelectSingleNode("uitemId").Text            ''*단품ID [00000]

                    'dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''지시수량 [2]
                    'cnclItemQty         = LagrgeNode(i).SelectSingleNode("cnclItemQty").Text        ''취소수량 [0]
                    'ordQty              = LagrgeNode(i).SelectSingleNode("ordQty").Text             ''주문수량 [2]
                    procItemQty         = LagrgeNode(i).SelectSingleNode("procItemQty").Text             ''처리수량 [2]
                    sellprc             = LagrgeNode(i).SelectSingleNode("sellprc").Text            ''판매가 [1000]
                    'frgShppYn           = LagrgeNode(i).SelectSingleNode("frgShppYn").Text          ''국내/외 구분 [국내]
                    ordpeNm             = LagrgeNode(i).SelectSingleNode("ordpeNm").Text            ''*주문자

                    rcptpeNm            = LagrgeNode(i).SelectSingleNode("rcptpeNm").Text           ''*수취인
                    'rcptpeHpno          = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text         ''*수취인 휴대폰번호
                    'if NOT (LagrgeNode(i).SelectSingleNode("rcptpeTelno") is Nothing) then
                    '    rcptpeTelno         = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text        ''*수취인 집전화번호
                    'end if
                    'shpplocAddr         = LEFT(LagrgeNode(i).SelectSingleNode("shpplocAddr").Text, 500)        ''수취인 상세주소
					'if NOT (LagrgeNode(i).SelectSingleNode("shpplocZipcd") is Nothing) then
                    '	shpplocZipcd        = LagrgeNode(i).SelectSingleNode("shpplocZipcd").Text       ''*수취인 우편번호          [04733]
					'end if
                    'if NOT (LagrgeNode(i).SelectSingleNode("shpplocOldZipcd") is Nothing) then
                    '    shpplocOldZipcd     = LagrgeNode(i).SelectSingleNode("shpplocOldZipcd").Text    ''*수취인 구우편번호(6자리)  [133750]
                    'end if
                    'shpplocRoadAddr     = LEFT(LagrgeNode(i).SelectSingleNode("shpplocRoadAddr").Text, 500)    ''수취인도로명주소
                    'itemChrctDivCd      = LagrgeNode(i).SelectSingleNode("itemChrctDivCd").Text     ''상품특성구분코드 10 일반 20 몰인몰 30 해외구매대행상품 40 미가공귀금속 50 모바일기프트 60 상품권 70 쇼핑충전금 80 모바일상품권 91 이벤트
                    'shppStatCd          = LagrgeNode(i).SelectSingleNode("shppStatCd").Text         ''*배송상태코드 10 정상 30 대기
                    'shppStatNm          = LagrgeNode(i).SelectSingleNode("shppStatNm").Text         ''배송상태명
                    orordNo             = LagrgeNode(i).SelectSingleNode("orordNo").Text            ''원주문번호 [20171123128379]
                    orordItemSeq        = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''원주문순번 [2]
                    'shppMainCd          = LagrgeNode(i).SelectSingleNode("shppMainCd").Text         ''배송주체코드 32 업체창고 41 협력업체 42 브랜드직배  [41]
                    'siteNo              = LagrgeNode(i).SelectSingleNode("siteNo").Text             ''사이트번호 6001 이마트몰 6002 트레이더스몰 6003 분스몰 6004 신세계몰 6005 S.COM몰 6009 신세계백화점몰
                    'siteNm              = LagrgeNode(i).SelectSingleNode("siteNm").Text             ''사이트명
                    'shppRsvtDt          = LagrgeNode(i).SelectSingleNode("shppRsvtDt").Text
                    splprc              = LagrgeNode(i).SelectSingleNode("splprc").Text             ''공급가
                    'shortgYn            = LagrgeNode(i).SelectSingleNode("shortgYn").Text
                    'newWblNoData        = LagrgeNode(i).SelectSingleNode("newWblNoData").Text
                    'newRow              = LagrgeNode(i).SelectSingleNode("newRow").Text
                    'itemDiv             = LagrgeNode(i).SelectSingleNode("itemDiv").Text                ''판매불가신청상태 10:일반 20: 명절 GIFT 일반 30: 명절 GIFT 센터 40: 명절 GIFT 냉장
                    'shpplocBascAddr     = LagrgeNode(i).SelectSingleNode("shpplocBascAddr").Text        ''수취인주소 20170712
                    'shpplocDtlAddr      = LEFT(LagrgeNode(i).SelectSingleNode("shpplocDtlAddr").Text, 500)         ''수취인상세주소	20170712
                    'ordItemDivNm        = LagrgeNode(i).SelectSingleNode("ordItemDivNm").Text           ''주문상품구분	20170809  // 주문, 부분배송주문


                    ''//필수값 아닌경우 .
                    ' if NOT (LagrgeNode(i).SelectSingleNode("ordpeHpno") is Nothing) then
                    '     ordpeHpno         = LagrgeNode(i).SelectSingleNode("ordpeHpno").Text           ''주문자휴대폰번호  //선택값
                    ' end if

                    ' if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
                    '     ordMemoCntt         = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[고객배송메모]","")            ''고객배송메모  //선택값
                    ' end if

                    ' if NOT (LagrgeNode(i).SelectSingleNode("pCus") is Nothing) then
                    '     pCus         = LagrgeNode(i).SelectSingleNode("pCus").Text           ''개인통관고유번호  //선택값
                    ' end if

                    ' if NOT (LagrgeNode(i).SelectSingleNode("frebieNm") is Nothing) then
                    '     frebieNm         = LagrgeNode(i).SelectSingleNode("frebieNm").Text    ''사은품  //선택값
                    ' end if

                    ' if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatCd") is Nothing) then
                    '     shortgProgStatCd         = LagrgeNode(i).SelectSingleNode("shortgProgStatCd").Text    ''판매불가신청상태  //선택값 11 결품등록 12 결품CS처리중 13 결품확정 21 상품정보오류등록 22 상품정보오류CS처리중 23 상품정보오류확정 41 입고지연등록 43 입고지연완료 51 배송지연등록
                    ' end if

                    ' if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatNm") is Nothing) then
                    '     shortgProgStatNm         = LagrgeNode(i).SelectSingleNode("shortgProgStatNm").Text    ''결품진행상태명  //선택값
                    ' end if

                    ' ''옵션명
                    ' if NOT (LagrgeNode(i).SelectSingleNode("uitemNm") is Nothing) then
                    '     uitemNm         = LagrgeNode(i).SelectSingleNode("uitemNm").Text                 ''[,1:^:asdasd:^:주문문구1,2:^:asdasdddd:^:주문문구2]
                    ' end if

                    ''if (orordNo<>ordNo) then ''원주문번호로 업데이트 ''부분출고처리시 주문번호가 바뀜
                    ''    ordNo=orordNo
                    ''end if

                    ''if (orordItemSeq<>ordItemSeq) then  ''2018/03/05 추가  <ordItemDivNm>부분배송주문</ordItemDivNm> 20180305585498
                    ''    ordItemSeq=orordItemSeq
                    ''end if

                    if (orordNo=ordNo) and (orordItemSeq=ordItemSeq) then
                        orordNo=""
                        orordItemSeq=""
                    end if

                    ''' 출고기준일
                    if NOT (LagrgeNode(i).SelectSingleNode("whoutCritnDt") is Nothing) then
                        whoutCritnDt         = LagrgeNode(i).SelectSingleNode("whoutCritnDt").Text                 ''출고기준일
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("delicoVenNm") is Nothing) then
                        delicoVenNm         = LagrgeNode(i).SelectSingleNode("delicoVenNm").Text                 ''택배사명
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("wblNo") is Nothing) then
                        wblNo         = LagrgeNode(i).SelectSingleNode("wblNo").Text                 ''운송장번호
                    end if


                    '' 자동결품여부
                    ' if NOT (LagrgeNode(i).SelectSingleNode("autoShortgYn") is Nothing) then
                    '     autoShortgYn         = LagrgeNode(i).SelectSingleNode("autoShortgYn").Text                 ''자동결품여부
                    ' end if

                    ' response.write "<br>"
                    ' response.write ordNo&":"&shppDivDtlCd&":"&shppNo&":"&ordItemSeq
                    ' response.write ":출고기준일:"&whoutCritnDt&":shppRsvtDt:"&shppRsvtDt&":자동결품여부:"&autoShortgYn&":상품명:"&itemNm&":옵션명:"&uitemNm

                    ' if (shppRsvtDt<>"") then
                    '     shppRsvtDt   = LEFT(shppRsvtDt,4)&"-"&MID(shppRsvtDt,5,2)&"-"&RIGHT(shppRsvtDt,2)
                    ' end if
                    ' if (whoutCritnDt<>"") then
                    '     whoutCritnDt = LEFT(whoutCritnDt,4)&"-"&MID(whoutCritnDt,5,2)&"-"&RIGHT(whoutCritnDt,2)
                    ' end if

                    ' rw ordNo&":"&ordItemSeq&":"&confirmDt&":"&shppNo&":"&shppSeq&":"&reOrderYn&":"&delayNts&":"&splVenItemId
                    ' rw itemId&":"&uitemId&":"&ordQty&":"&shppDivDtlNm&":"&uitemNm&":"&shppRsvtDt&":"&whoutCritnDt&":"&autoShortgYn



                    sqlStr = "db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Input]"
                    paramInfo = Array(Array("@RETURN_VALUE",adInteger	,adParamReturnValue	,,0) _
                        ,Array("@SellSite" 				, adVarchar		, adParamInput		, 	32, Trim("ssg"))	_
                        ,Array("@OutMallOrderSerial"	, adVarchar		, adParamInput		,	32, Trim(ordNo)) _
                        ,Array("@OrgDetailKey"			, adVarchar		, adParamInput		,   32, Trim(ordItemSeq)) _

                        ,Array("@confirmDt"				, adVarchar     , adParamInput		,	16, Trim(confirmDt)) _
                        ,Array("@shppNo"				, adVarchar		, adParamInput		,   32, Trim(shppNo)) _
                        ,Array("@shppSeq"				, adVarchar		, adParamInput		,   10, Trim(shppSeq)) _
                        ,Array("@reOrderYn"				, adVarchar		, adParamInput		,    1, Trim(reOrderYn)) _
                        ,Array("@delayNts"			    , adInteger		, adParamInput		,     , Trim(delayNts)) _
                        ,Array("@splVenItemId"			, adInteger		, adParamInput		,     , Trim(splVenItemId)) _
                        ,Array("@outMallGoodsNo"		, adVarchar		, adParamInput		,   16, Trim(itemId)) _
                        ,Array("@outMallOptionNo"		, adVarchar		, adParamInput		,   32, Trim(uitemId)) _
                        ,Array("@ordQty"			    , adInteger		, adParamInput		,     , Trim(procItemQty)) _
                        ,Array("@shppDivDtlNm"		    , adVarchar		, adParamInput		,   20, Trim(shppDivDtlNm)) _
                        ,Array("@uitemNm"		        , adVarchar		, adParamInput		,   128, Trim(uitemNm)) _
                        ,Array("@shppRsvtDt"			, adDate		, adParamInput		,	  , Trim(shppRsvtDt)) _
                        ,Array("@whoutCritnDt"			, adDate		, adParamInput		,	  , Trim(whoutCritnDt)) _
                        ,Array("@autoShortgYn"			, adVarchar		, adParamInput		,    1, Trim(autoShortgYn)) _
                        ,Array("@outorderstatus"		, adVarchar		, adParamInput		,   30, Trim(lastShppProgStatDtlNm)) _

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

			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing


	rw "건수:"&successCnt
    rw "======================================"
end function

''촐고 대상목록 조회
public function getSsgDlvConfirmList(byVal styyyymmdd,byVal edyyyymmdd)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim shppNo,shppSeq,shppTabProgStatCd,evntSeq,shppDivDtlCd,shppDivDtlNm,reOrderYn,delayNts,ordNo,ordItemSeq,ordCmplDts
    Dim lastShppProgStatDtlNm,lastShppProgStatDtlCd,salestrNo,shppVenId,shppVenNm,shppTypeNm,shppTypeCd,shppTypeDtlCd,shppTypeDtlNm,boxNo
    Dim shppcst,shppcstCodYn,itemNm,splVenItemId,itemId,uitemId,dircItemQty,cnclItemQty,ordQty,sellprc,frgShppYn
    Dim ordpeNm,rcptpeNm,rcptpeHpno,rcptpeTelno,shpplocAddr,shpplocZipcd,shpplocOldZipcd,shpplocRoadAddr,itemChrctDivCd,shppStatCd,shppStatNm
    Dim orordNo,orordItemSeq,shppMainCd,siteNo,siteNm,shppRsvtDt,splprc,shortgYn,newWblNoData,newRow,itemDiv
    Dim shpplocBascAddr,shpplocDtlAddr,ordItemDivNm
    Dim ordpeHpno, ordMemoCntt, pCus, frebieNm ,shortgProgStatCd, shortgProgStatNm, uitemNm
    Dim iBufrequireDetail

    Dim delicoVenId ''택배사ID
    Dim delicoVenNm	''택배사명
    Dim wblNo	    ''운송장번호

    Dim whoutCritnDt, autoShortgYn

    Dim oMaster, oDetailArr(0)
    Dim successCnt : successCnt=0
    Dim failCnt : failCnt=0

    rw "기간검색:"&styyyymmdd&"~"&edyyyymmdd&" 상태:"&"주문확인"
    styyyymmdd = replace(styyyymmdd,"-","")
    edyyyymmdd = replace(edyyyymmdd,"-","")

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listWarehouseOut.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestWarehouseOut>"
    requestBody = requestBoDy&"<perdType>01</perdType>"  ''주문확인일
    requestBody = requestBoDy&"<perdStrDts>"&styyyymmdd&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&edyyyymmdd&"</perdEndDts>"  ''하루를 더해야?
    ''requestBody = requestBoDy&"<wblNoRegYn>N</wblNoRegYn>" ''운송장등록여부
    requestBody = requestBoDy&"</requestWarehouseOut>"
	objXML.send(requestBody)

'rw objXML.status
if (isOnlyTodayBaljuView) then
    response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"
    'response.end
end if
    dim retBody : retBody=objXML.responseText
    Dim paramInfo, RetparamInfo, RetErr
    retBody = replace(retBody,"&","")
	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(retBody) ''objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text


			Set LagrgeNode = xmlDOM.SelectNodes("/result/warehouseOuts/warehouseOut")
			If Not (LagrgeNode Is Nothing) Then
                ''초기화(기간별)
                ' if (LagrgeNode.length>0) then
                '     strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'ssg','"&confirmDt&"'"
                '     dbget.Execute strSql
                ' end if

			    For i = 0 To LagrgeNode.length - 1
			        ''변수초기화.
			        shppNo ="": shppSeq = "": shppTabProgStatCd ="": evntSeq ="": shppDivDtlCd =""
                    shppDivDtlNm ="": reOrderYn ="": delayNts ="": ordNo ="": ordItemSeq =""
                    ordCmplDts ="": lastShppProgStatDtlNm = "": lastShppProgStatDtlCd ="": salestrNo ="": shppVenId =""
                    shppVenNm ="": shppTypeNm ="": shppTypeCd ="": shppTypeDtlCd ="": shppTypeDtlNm =""
                    delicoVenId ="": boxNo ="": shppcst ="": shppcstCodYn ="": itemNm =""
                    splVenItemId ="":itemId ="":uitemId ="": dircItemQty ="": cnclItemQty =""
                    ordQty ="" :sellprc ="": frgShppYn ="": ordpeNm =""
                    rcptpeNm ="" :rcptpeHpno ="": rcptpeTelno ="": shpplocAddr =""
                    shpplocZipcd ="": shpplocOldZipcd ="": shpplocRoadAddr ="": itemChrctDivCd =""
                    shppStatCd ="": shppStatNm ="": orordNo ="": orordItemSeq ="": shppMainCd =""
                    siteNo ="": siteNm ="": shppRsvtDt ="": splprc ="": shortgYn =""
                    newWblNoData ="": newRow ="": itemDiv ="": shpplocBascAddr ="": shpplocDtlAddr ="": ordItemDivNm =""

                    ordpeHpno = "": ordMemoCntt = "": pCus = "": frebieNm = "": shortgProgStatCd ="": shortgProgStatNm ="" : uitemNm=""
                    iBufrequireDetail = ""
                    whoutCritnDt =""

                    delicoVenNm	="" ''택배사명
                    wblNo	    ="" ''운송장번호
                    delicoVenId =""

			        shppNo              = LagrgeNode(i).SelectSingleNode("shppNo").Text                 ''*배송번호
                    shppSeq             = LagrgeNode(i).SelectSingleNode("shppSeq").Text                ''*배송순번
                    shppTabProgStatCd   = LagrgeNode(i).SelectSingleNode("shppTabProgStatCd").Text      ''최종배송상세진행상태코드(배송단위) 11 배송지시 21 피킹지시 22 피킹완료 31 패킹완료 41 출고보류 42 출고지연 43 출고완료 51 배송완료 52 배송거절
					if NOT (LagrgeNode(i).SelectSingleNode("evntSeq") is Nothing) then
                    	evntSeq             = LagrgeNode(i).SelectSingleNode("evntSeq").Text                ''이벤트순번
					end if
                    shppDivDtlCd        = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text           ''*배송구분상세코드 11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고
                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''배송구분상세명
                    reOrderYn           = LagrgeNode(i).SelectSingleNode("reOrderYn").Text              ''*재지시여부구분
                    delayNts            = LagrgeNode(i).SelectSingleNode("delayNts").Text               ''지연횟수
                    ordNo               = LagrgeNode(i).SelectSingleNode("ordNo").Text                  ''*주문번호 [20171123128379]
                    ordItemSeq          = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text             ''*주문순번
                    ordCmplDts          = LagrgeNode(i).SelectSingleNode("ordCmplDts").Text             ''*주문완료일시 [2017-11-23 10:39:42.0]
                    lastShppProgStatDtlNm   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlNm").Text  ''최종배송상세진행상태명(배송상품단위) [피킹완료]
                    lastShppProgStatDtlCd   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''최종배송상세진행상태코드(배송상품단위) 11 배송지시 21 피킹지시 22 피킹완료 31 패킹완료 41 출고보류 42 출고지연 43 출고완료 51 배송완료 52 배송거절
                    salestrNo           = LagrgeNode(i).SelectSingleNode("salestrNo").Text              '' [6004]
                    shppVenId           = LagrgeNode(i).SelectSingleNode("shppVenId").Text      ''공급업체아이디 [0000003198]
                    shppVenNm           = LagrgeNode(i).SelectSingleNode("shppVenNm").Text      ''공급업체명
                    shppTypeNm          = LagrgeNode(i).SelectSingleNode("shppTypeNm").Text     ''배송유형명    [택배배송]
                    shppTypeCd          = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text     ''배송유형코드 10 자사배송 20 택배배송 30 매장방문 40 등기 50 미배송 60 미발송
                    shppTypeDtlCd       = LagrgeNode(i).SelectSingleNode("shppTypeDtlCd").Text  ''배송유형상세코드 14 업체자사배송 22 업체택배배송 25 해외택배배송 31 매장방문 41 등기 51 SMS 52 EMAIL 61 미발송
                    shppTypeDtlNm       = LagrgeNode(i).SelectSingleNode("shppTypeDtlNm").Text  ''배송유형상세명 [업체택배배송]
                    if NOT (LagrgeNode(i).SelectSingleNode("delicoVenId") is Nothing) then
                        delicoVenId         = LagrgeNode(i).SelectSingleNode("delicoVenId").Text    ''택배사ID [0000033011]
                    end if
                    boxNo               = LagrgeNode(i).SelectSingleNode("boxNo").Text          ''박스번호 [398327952]
                    shppcst             = LagrgeNode(i).SelectSingleNode("shppcst").Text        '' 배송비? [303] ??
                    shppcstCodYn        = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text   ''*배송비 착불여부 Y: 착불 N: 선불
                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*상품명
                    itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*상품번호 [1000024811163]

					If NOT (LagrgeNode(i).SelectSingleNode("splVenItemId") is Nothing) Then
	    			    splVenItemId    = LagrgeNode(i).SelectSingleNode("splVenItemId").Text        ''업체상품번호 [1024019]
					Else
						strSql = ""
						strSql = strSql & " select top 1 itemid "
						strSql = strSql & " FROM db_etcmall.dbo.tbl_ssg_regitem"
						strSql = strSql & " where ssgGoodNo = '"& itemId &"' "
						rsget.CursorLocation = adUseClient
						rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
						If Not rsget.Eof Then
							splVenItemId = rsget("itemid")
						Else
							rw "오류 주문번호 : " & ordNo
						End If
						rsget.Close

						If session("ssBctID")="kjy8517" Then
							response.write "<textarea cols=100 rows=30>!!"&objXML.responseText&"</textarea>"
						End If
					End If

                    uitemId             = LagrgeNode(i).SelectSingleNode("uitemId").Text            ''*단품ID [00000]
                    dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''지시수량 [2]
                    cnclItemQty         = LagrgeNode(i).SelectSingleNode("cnclItemQty").Text        ''취소수량 [0]
                    ordQty              = LagrgeNode(i).SelectSingleNode("ordQty").Text             ''주문수량 [2]
                    sellprc             = LagrgeNode(i).SelectSingleNode("sellprc").Text            ''판매가 [1000]
                    'frgShppYn           = LagrgeNode(i).SelectSingleNode("frgShppYn").Text          ''국내/외 구분 [국내]
                    ordpeNm             = LagrgeNode(i).SelectSingleNode("ordpeNm").Text            ''*주문자

                    rcptpeNm            = LagrgeNode(i).SelectSingleNode("rcptpeNm").Text           ''*수취인
                    rcptpeHpno          = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text         ''*수취인 휴대폰번호
                    if NOT (LagrgeNode(i).SelectSingleNode("rcptpeTelno") is Nothing) then
                        rcptpeTelno         = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text        ''*수취인 집전화번호
                    end if
                    shpplocAddr         = LEFT(LagrgeNode(i).SelectSingleNode("shpplocAddr").Text, 500)        ''수취인 상세주소
					if NOT (LagrgeNode(i).SelectSingleNode("shpplocZipcd") is Nothing) then
                    	shpplocZipcd        = LagrgeNode(i).SelectSingleNode("shpplocZipcd").Text       ''*수취인 우편번호          [04733]
					end if
                    if NOT (LagrgeNode(i).SelectSingleNode("shpplocOldZipcd") is Nothing) then
                        shpplocOldZipcd     = LagrgeNode(i).SelectSingleNode("shpplocOldZipcd").Text    ''*수취인 구우편번호(6자리)  [133750]
                    end if
                    shpplocRoadAddr     = LEFT(LagrgeNode(i).SelectSingleNode("shpplocRoadAddr").Text, 500)    ''수취인도로명주소
                    itemChrctDivCd      = LagrgeNode(i).SelectSingleNode("itemChrctDivCd").Text     ''상품특성구분코드 10 일반 20 몰인몰 30 해외구매대행상품 40 미가공귀금속 50 모바일기프트 60 상품권 70 쇼핑충전금 80 모바일상품권 91 이벤트
                    shppStatCd          = LagrgeNode(i).SelectSingleNode("shppStatCd").Text         ''*배송상태코드 10 정상 30 대기
                    shppStatNm          = LagrgeNode(i).SelectSingleNode("shppStatNm").Text         ''배송상태명
                    orordNo             = LagrgeNode(i).SelectSingleNode("orordNo").Text            ''원주문번호 [20171123128379]
                    orordItemSeq        = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''원주문순번 [2]
                    shppMainCd          = LagrgeNode(i).SelectSingleNode("shppMainCd").Text         ''배송주체코드 32 업체창고 41 협력업체 42 브랜드직배  [41]
                    siteNo              = LagrgeNode(i).SelectSingleNode("siteNo").Text             ''사이트번호 6001 이마트몰 6002 트레이더스몰 6003 분스몰 6004 신세계몰 6005 S.COM몰 6009 신세계백화점몰
                    siteNm              = LagrgeNode(i).SelectSingleNode("siteNm").Text             ''사이트명
                    shppRsvtDt          = LagrgeNode(i).SelectSingleNode("shppRsvtDt").Text
                    splprc = 0
                    if NOT (LagrgeNode(i).SelectSingleNode("splprc") is Nothing) then
                        splprc              = LagrgeNode(i).SelectSingleNode("splprc").Text             ''공급가
                    end if
                    shortgYn            = LagrgeNode(i).SelectSingleNode("shortgYn").Text
                    newWblNoData        = LagrgeNode(i).SelectSingleNode("newWblNoData").Text
                    newRow              = LagrgeNode(i).SelectSingleNode("newRow").Text
                    itemDiv             = LagrgeNode(i).SelectSingleNode("itemDiv").Text                ''판매불가신청상태 10:일반 20: 명절 GIFT 일반 30: 명절 GIFT 센터 40: 명절 GIFT 냉장
                    shpplocBascAddr     = LagrgeNode(i).SelectSingleNode("shpplocBascAddr").Text        ''수취인주소 20170712
                    shpplocDtlAddr      = LEFT(LagrgeNode(i).SelectSingleNode("shpplocDtlAddr").Text, 500)         ''수취인상세주소	20170712
                    ordItemDivNm        = LagrgeNode(i).SelectSingleNode("ordItemDivNm").Text           ''주문상품구분	20170809  // 주문, 부분배송주문


                    ''//필수값 아닌경우 .
                    if NOT (LagrgeNode(i).SelectSingleNode("ordpeHpno") is Nothing) then
                        ordpeHpno         = LagrgeNode(i).SelectSingleNode("ordpeHpno").Text           ''주문자휴대폰번호  //선택값
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
                        ordMemoCntt         = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[고객배송메모]","")            ''고객배송메모  //선택값
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("pCus") is Nothing) then
                        pCus         = LagrgeNode(i).SelectSingleNode("pCus").Text           ''개인통관고유번호  //선택값
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("frebieNm") is Nothing) then
                        frebieNm         = LagrgeNode(i).SelectSingleNode("frebieNm").Text    ''사은품  //선택값
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatCd") is Nothing) then
                        shortgProgStatCd         = LagrgeNode(i).SelectSingleNode("shortgProgStatCd").Text    ''판매불가신청상태  //선택값 11 결품등록 12 결품CS처리중 13 결품확정 21 상품정보오류등록 22 상품정보오류CS처리중 23 상품정보오류확정 41 입고지연등록 43 입고지연완료 51 배송지연등록
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatNm") is Nothing) then
                        shortgProgStatNm         = LagrgeNode(i).SelectSingleNode("shortgProgStatNm").Text    ''결품진행상태명  //선택값
                    end if

                    ''옵션명
                    if NOT (LagrgeNode(i).SelectSingleNode("uitemNm") is Nothing) then
                        uitemNm         = LagrgeNode(i).SelectSingleNode("uitemNm").Text                 ''[,1:^:asdasd:^:주문문구1,2:^:asdasdddd:^:주문문구2]
                    end if

                    ''if (orordNo<>ordNo) then ''원주문번호로 업데이트 ''부분출고처리시 주문번호가 바뀜
                    ''    ordNo=orordNo
                    ''end if

                    ''if (orordItemSeq<>ordItemSeq) then  ''2018/03/05 추가  <ordItemDivNm>부분배송주문</ordItemDivNm> 20180305585498
                    ''    ordItemSeq=orordItemSeq
                    ''end if

                    if (orordNo=ordNo) and (orordItemSeq=ordItemSeq) then
                        orordNo=""
                        orordItemSeq=""
                    end if

                    ''' 출고기준일
                    if NOT (LagrgeNode(i).SelectSingleNode("whoutCritnDt") is Nothing) then
                        whoutCritnDt         = LagrgeNode(i).SelectSingleNode("whoutCritnDt").Text                 ''출고기준일
                    end if

                    '' 자동결품여부
                    if NOT (LagrgeNode(i).SelectSingleNode("autoShortgYn") is Nothing) then
                        autoShortgYn         = LagrgeNode(i).SelectSingleNode("autoShortgYn").Text                 ''자동결품여부
                    end if


                    if NOT (LagrgeNode(i).SelectSingleNode("delicoVenNm") is Nothing) then
                        delicoVenNm         = LagrgeNode(i).SelectSingleNode("delicoVenNm").Text                 ''택배사명
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("wblNo") is Nothing) then
                        wblNo         = LagrgeNode(i).SelectSingleNode("wblNo").Text                 ''운송장번호
                    end if


                    ' response.write "<br>"
                    ' response.write ordNo&":"&shppDivDtlCd&":"&shppNo&":"&ordItemSeq
                    ' response.write ":출고기준일:"&whoutCritnDt&":shppRsvtDt:"&shppRsvtDt&":자동결품여부:"&autoShortgYn&":상품명:"&itemNm&":옵션명:"&uitemNm

                    if (shppRsvtDt<>"") then
                        shppRsvtDt   = LEFT(shppRsvtDt,4)&"-"&MID(shppRsvtDt,5,2)&"-"&RIGHT(shppRsvtDt,2)
                    end if
                    if (whoutCritnDt<>"") then
                        whoutCritnDt = LEFT(whoutCritnDt,4)&"-"&MID(whoutCritnDt,5,2)&"-"&RIGHT(whoutCritnDt,2)
                    end if



                    ' rw ordNo&":"&ordItemSeq&":"&confirmDt&":"&shppNo&":"&shppSeq&":"&reOrderYn&":"&delayNts&":"&splVenItemId
                    ' rw itemId&":"&uitemId&":"&ordQty&":"&shppDivDtlNm&":"&uitemNm&":"&shppRsvtDt&":"&whoutCritnDt&":"&autoShortgYn



                    sqlStr = "db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Input]"
                    paramInfo = Array(Array("@RETURN_VALUE",adInteger	,adParamReturnValue	,,0) _
                        ,Array("@SellSite" 				, adVarchar		, adParamInput		, 	32, Trim("ssg"))	_
                        ,Array("@OutMallOrderSerial"	, adVarchar		, adParamInput		,	32, Trim(ordNo)) _
                        ,Array("@OrgDetailKey"			, adVarchar		, adParamInput		,   32, Trim(ordItemSeq)) _

                        ,Array("@confirmDt"				, adVarchar     , adParamInput		,	16, Trim(confirmDt)) _
                        ,Array("@shppNo"				, adVarchar		, adParamInput		,   32, Trim(shppNo)) _
                        ,Array("@shppSeq"				, adVarchar		, adParamInput		,   10, Trim(shppSeq)) _
                        ,Array("@reOrderYn"				, adVarchar		, adParamInput		,    1, Trim(reOrderYn)) _
                        ,Array("@delayNts"			    , adInteger		, adParamInput		,     , Trim(delayNts)) _
                        ,Array("@splVenItemId"			, adInteger		, adParamInput		,     , Trim(splVenItemId)) _
                        ,Array("@outMallGoodsNo"		, adVarchar		, adParamInput		,   16, Trim(itemId)) _
                        ,Array("@outMallOptionNo"		, adVarchar		, adParamInput		,   32, Trim(uitemId)) _
                        ,Array("@ordQty"			    , adInteger		, adParamInput		,     , Trim(ordQty)) _
                        ,Array("@shppDivDtlNm"		    , adVarchar		, adParamInput		,   20, Trim(shppDivDtlNm)) _
                        ,Array("@uitemNm"		        , adVarchar		, adParamInput		,   128, Trim(uitemNm)) _
                        ,Array("@shppRsvtDt"			, adDate		, adParamInput		,	  , Trim(shppRsvtDt)) _
                        ,Array("@whoutCritnDt"			, adDate		, adParamInput		,	  , Trim(whoutCritnDt)) _
                        ,Array("@autoShortgYn"			, adVarchar		, adParamInput		,    1, Trim(autoShortgYn)) _
                        ,Array("@outorderstatus"		, adVarchar		, adParamInput		,   30, Trim(lastShppProgStatDtlNm)) _

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

                '' 주문번호 매핑.
                ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'ssg','"&confirmDt&"'"
                ' dbget.Execute strSql

			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing


	rw "건수:"&successCnt
    rw "======================================"
end function


''배송지시목록 조회
public function getSsgDlvReqList(byVal styyyymmdd,byVal edyyyymmdd)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim ordNo, ordItemSeq, orOrdNo, orordItemSeq, shppNo, shppSeq, ordStatCd, shppStatCd, shppStatNm, itemId, itemNm, splVenItemId
    Dim ordCstId, ordCstOccCd, shppcst, shppcstCodYn, ordRcpDts, ordpeNm, rcptpeNm, rcptpeHpno, rcptpeTelno, shppDivDtlCd, shppProgStatDtlCd, shppRsvtDt
    Dim uitemId, uitemNm, siteNo, rsvtItemYn, frgShppYn, dircItemQty, cnclItemQty, ordQty, splprc, sellprc, ordCmplDts, ordpeHpno
    Dim shpplocAddr, shpplocZipcd, shpplocOldZipcd, ordMemoCntt, ordpeRoadAddr, ordShpplocId, shppTypeDtlCd, reOrderYn, itemDiv, shpplocBascAddr, shpplocDtlAddr, ordItemDivNm
    Dim delayNts, shppDivDtlNm, whoutCritnDt, autoShortgYn

    Dim ArrShppNo, ArrShppSeq, ArrshppStatCd, lastShppProgStatDtlNm
    Dim paramInfo, RetparamInfo, RetErr
    Dim successCnt : successCnt=0
    Dim shppTypeDtlNm
    Dim delicoVenId ''택배사ID
    Dim delicoVenNm	''택배사명
    Dim wblNo	    ''운송장번호

    rw "기간검색:"&styyyymmdd&"~"&edyyyymmdd&" 상태:"&"주문미확인"
    styyyymmdd = replace(styyyymmdd,"-","")
    edyyyymmdd = replace(edyyyymmdd,"-","")

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listShppDirection.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestShppDirection>"
    requestBody = requestBoDy&"<perdType>01</perdType>"
    requestBody = requestBoDy&"<perdStrDts>"&styyyymmdd&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&edyyyymmdd&"</perdEndDts>"
    requestBody = requestBoDy&"</requestShppDirection>"

	objXML.send(requestBody)
	''rw objXML.status
'response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"
'response.end

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text

			Set LagrgeNode = xmlDOM.SelectNodes("/result/shppDirections/shppDirection")
			If Not (LagrgeNode Is Nothing) Then
			    ''response.write "건수:" & LagrgeNode.length
			    redim ArrShppNo(LagrgeNode.length-1)
			    redim ArrShppSeq(LagrgeNode.length-1)
			    redim ArrshppStatCd(LagrgeNode.length-1)

			    For i = 0 To LagrgeNode.length - 1
    			    ordNo="": ordItemSeq="": orOrdNo="": orordItemSeq="": shppNo="": shppSeq="": ordStatCd="": shppStatCd=""
                    shppStatNm="": itemId="": itemNm="": splVenItemId="": ordCstId="": ordCstOccCd="": shppcst="": shppcstCodYn=""
                    ordRcpDts="": ordpeNm="": rcptpeNm="": rcptpeHpno="": rcptpeTelno="": shppDivDtlCd="": shppProgStatDtlCd="": shppRsvtDt=""
                    uitemId="": uitemNm="": siteNo="": rsvtItemYn="": frgShppYn="": dircItemQty="": cnclItemQty="": ordQty="": splprc="": sellprc=""
                    ordCmplDts="": ordpeHpno="": shpplocAddr="": shpplocZipcd="": shpplocOldZipcd="": ordMemoCntt="": ordpeRoadAddr="": ordShpplocId=""
                    shppTypeDtlCd="": reOrderYn="": itemDiv="": shpplocBascAddr="": shpplocDtlAddr="": ordItemDivNm=""
                    delayNts = ""

                    shppTypeDtlNm=""
                    delicoVenId =""
                    delicoVenNm	=""
                    wblNo	    =""


                    whoutCritnDt = ""
                    autoShortgYn = ""
                    lastShppProgStatDtlNm = "주문통보"

                    shppNo           = LagrgeNode(i).SelectSingleNode("shppNo").Text                 ''**배송번호 [D2125835493]
    			    shppSeq         = LagrgeNode(i).SelectSingleNode("shppSeq").Text               ''**배송순번 [1]
    			    ordNo            = LagrgeNode(i).SelectSingleNode("ordNo").Text                 ''*주문번호 [20171127616023]
    			    ordItemSeq       = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text            ''*주문순번? [1]
                    If NOT (LagrgeNode(i).SelectSingleNode("orOrdNo") is Nothing) Then
                        orOrdNo          = LagrgeNode(i).SelectSingleNode("orOrdNo").Text                ''원주문번호
                    End If
    			    orordItemSeq     = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text                ''원주문순번  orOrdNo
    			    shppStatCd      = LagrgeNode(i).SelectSingleNode("shppStatCd").Text            '' *배송상태코드 10 정상 30 대기[10]

    			    ArrShppNo(i) = shppNo
    			    ArrShppSeq(i) = shppSeq
    			    ArrshppStatCd(i) = shppStatCd
                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''배송구분상세명
    			    ordStatCd       = LagrgeNode(i).SelectSingleNode("ordStatCd").Text              ''          [120]
    			    shppStatNm      = LagrgeNode(i).SelectSingleNode("shppStatNm").Text            '' 배송상태명         [정상]
    			    itemId          = LagrgeNode(i).SelectSingleNode("itemId").Text                  ''상품번호  [1000024811163]
    			    itemNm          = LagrgeNode(i).SelectSingleNode("itemNm").Text                  ''상품명    [문주란스티커]

					If NOT (LagrgeNode(i).SelectSingleNode("splVenItemId") is Nothing) Then
	    			    splVenItemId    = LagrgeNode(i).SelectSingleNode("splVenItemId").Text        ''업체상품번호 [1024019]
					Else
						strSql = ""
						strSql = strSql & " select top 1 itemid "
						strSql = strSql & " FROM db_etcmall.dbo.tbl_ssg_regitem"
						strSql = strSql & " where ssgGoodNo = '"& itemId &"' "
						rsget.CursorLocation = adUseClient
						rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
						If Not rsget.Eof Then
							splVenItemId = rsget("itemid")
						Else
							rw "오류 주문번호 : " & ordNo
						End If
						rsget.Close

						If session("ssBctID")="kjy8517" Then
							response.write "<textarea cols=100 rows=30>##"&objXML.responseText&"</textarea>"
						End If
					End If

    			    ordCstId        = LagrgeNode(i).SelectSingleNode("ordCstId").Text                ''주문비용아이디
    			    ordCstOccCd     = LagrgeNode(i).SelectSingleNode("ordCstOccCd").Text             ''주문비용발생코드 [부과] :: 01,02가 아님
    			    shppcst         = LagrgeNode(i).SelectSingleNode("shppcst").Text                 ''배송비?
    			    shppcstCodYn    = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text            ''배송비착불여부 : Y :착불,N :선불 [N]
    			    ordRcpDts       = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text            ''주문접수일시 [2017-11-27 09:32:31.0]
    			    ordpeNm         = LagrgeNode(i).SelectSingleNode("ordpeNm").Text                   ''주문자
    			    rcptpeNm        = LagrgeNode(i).SelectSingleNode("rcptpeNm").Text                ''수령인
    			    rcptpeHpno      = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text                ''수령인 휴대폰
    			    if NOT (LagrgeNode(i).SelectSingleNode("rcptpeTelno") is Nothing) then
        			    rcptpeTelno     = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text                ''수령인 전화 [--]
        			end if
    			    shppDivDtlCd    = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text               ''배송구분상세코드 11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고 [11]
    			    shppProgStatDtlCd = LagrgeNode(i).SelectSingleNode("shppProgStatDtlCd").Text        ' 최종배송상세진행상태코드(배송상품단위) 11 배송지시 21 피킹지시 22 피킹완료 31 패킹완료 41 출고보류 42 출고지연 43 출고완료 51 배송완료 52 배송거절	[11]
    			    shppRsvtDt      = LagrgeNode(i).SelectSingleNode("shppRsvtDt").Text                 ''[20171128]
    			    uitemId         = LagrgeNode(i).SelectSingleNode("uitemId").Text                 ''단품ID [00000]

    			    siteNo          = LagrgeNode(i).SelectSingleNode("siteNo").Text                 ''사이트번호 6001 이마트몰 6002 트레이더스몰 6003 분스몰 6004 신세계몰 6005 S.COM몰 6009 신세계백화점몰[6004]
    			    rsvtItemYn      = LagrgeNode(i).SelectSingleNode("rsvtItemYn").Text                 ''예약판매구분 [N]
'    			    frgShppYn       = LagrgeNode(i).SelectSingleNode("frgShppYn").Text                 ''국내/외 구분 [N]
    			    dircItemQty     = LagrgeNode(i).SelectSingleNode("dircItemQty").Text                 ''지시수량 [2]
    			    cnclItemQty     = LagrgeNode(i).SelectSingleNode("cnclItemQty").Text                 ''취소수량 [2]
    			    ordQty          = LagrgeNode(i).SelectSingleNode("ordQty").Text                 ''주문수량 [2]
    			    splprc          = LagrgeNode(i).SelectSingleNode("splprc").Text                 ''공급가 [755]
    			    sellprc         = LagrgeNode(i).SelectSingleNode("sellprc").Text                 ''판매가 [1000]


    			    if NOT (LagrgeNode(i).SelectSingleNode("ordCmplDts") is Nothing) then
    			        ordCmplDts      = LagrgeNode(i).SelectSingleNode("ordCmplDts").Text                 ''주문완료일시 [2017-11-27 09:32:31.0]
    			    end if
    			    if NOT (LagrgeNode(i).SelectSingleNode("ordpeHpno") is Nothing) then
    			        ordpeHpno       = LagrgeNode(i).SelectSingleNode("ordpeHpno").Text                 ''주문자휴대폰번호 [01091603979]
    			    end if
    			    shpplocAddr     = LagrgeNode(i).SelectSingleNode("shpplocAddr").Text                 ''[서울 성동구 옥수동 561번지 래미안옥수리버젠 104동 103호]
					if NOT (LagrgeNode(i).SelectSingleNode("shpplocZipcd") is Nothing) then
    			    	shpplocZipcd    = LagrgeNode(i).SelectSingleNode("shpplocZipcd").Text                 ''*수취인 우편번호 [04733]
					end if
    			    if NOT (LagrgeNode(i).SelectSingleNode("shpplocOldZipcd") is Nothing) then
    			        shpplocOldZipcd = LagrgeNode(i).SelectSingleNode("shpplocOldZipcd").Text                 ''수취인(구) 우편번호[133750]
    			    end if

    			    ordpeRoadAddr   = LagrgeNode(i).SelectSingleNode("ordpeRoadAddr").Text                 ''[서울 성동구 매봉길 15, 104동 103호 (옥수동, 래미안옥수리버젠)]
    			    ordShpplocId    = LagrgeNode(i).SelectSingleNode("ordShpplocId").Text                 ''주문배송지ID [1102603504]
    			    shppTypeDtlCd   = LagrgeNode(i).SelectSingleNode("shppTypeDtlCd").Text                 ''배송유형상세코드 14 업체자사배송 22 업체택배배송 25 해외택배배송 31 매장방문 41 등기 51 SMS 52 EMAIL 61 미발송  [22]
    			    reOrderYn       = LagrgeNode(i).SelectSingleNode("reOrderYn").Text                 ''*재지시여부구분  [N]
    			    itemDiv         = LagrgeNode(i).SelectSingleNode("itemDiv").Text                 ''판매불가신청상태 10:일반 20: 명절 GIFT 일반 30: 명절 GIFT 센터 40: 명절 GIFT 냉장 [10]
                    If NOT (LagrgeNode(i).SelectSingleNode("shpplocBascAddr") is Nothing) then
    			        shpplocBascAddr = LagrgeNode(i).SelectSingleNode("shpplocBascAddr").Text                 '' [서울 성동구 매봉길]
                    End If

                    If NOT (LagrgeNode(i).SelectSingleNode("shpplocDtlAddr") is Nothing) then
                        shpplocDtlAddr  = LEFT(LagrgeNode(i).SelectSingleNode("shpplocDtlAddr").Text, 500)                 ''[15, 104동 103호 (옥수동, 래미안옥수리버젠)]
                    End If
    			    ordItemDivNm    = LagrgeNode(i).SelectSingleNode("ordItemDivNm").Text                 ''[주문]

    			    ' if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
    			    '     ordMemoCntt     = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[고객배송메모]","")                 ''[[고객배송메모]배송메세지]
    			    ' end if

    			    if NOT (LagrgeNode(i).SelectSingleNode("uitemNm") is Nothing) then
    			        uitemNm         = LagrgeNode(i).SelectSingleNode("uitemNm").Text                 ''[,1:^:asdasd:^:주문문구1,2:^:asdasdddd:^:주문문구2]
    			    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("shppTypeDtlNm") is Nothing) then
                        shppTypeDtlNm       = LagrgeNode(i).SelectSingleNode("shppTypeDtlNm").Text  ''배송유형상세명 [업체택배배송]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("delicoVenId") is Nothing) then
                        delicoVenId         = LagrgeNode(i).SelectSingleNode("delicoVenId").Text                 ''택배사ID [0000033011]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("delicoVenNm") is Nothing) then
                        delicoVenNm         = LagrgeNode(i).SelectSingleNode("delicoVenNm").Text                 ''택배사명
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("wblNo") is Nothing) then
                        wblNo         = LagrgeNode(i).SelectSingleNode("wblNo").Text                 ''운송장번호
                    end if

                    if (orordNo=ordNo) and (orordItemSeq=ordItemSeq) then
                        orordNo=""
                        orordItemSeq=""
                    end if

                    if (shppRsvtDt<>"") then
                        shppRsvtDt   = LEFT(shppRsvtDt,4)&"-"&MID(shppRsvtDt,5,2)&"-"&RIGHT(shppRsvtDt,2)
                    end if
                    if (whoutCritnDt<>"") then
                        whoutCritnDt = LEFT(whoutCritnDt,4)&"-"&MID(whoutCritnDt,5,2)&"-"&RIGHT(whoutCritnDt,2)
                    end if

                    ' rw ordNo&":"&ordItemSeq&":"&confirmDt&":"&shppNo&":"&shppSeq&":"&reOrderYn&":"&delayNts&":"&splVenItemId
                    ' rw itemId&":"&uitemId&":"&ordQty&":"&shppDivDtlNm&":"&uitemNm&":"&shppRsvtDt&":"&whoutCritnDt&":"&autoShortgYn

                    sqlStr = "db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Input]"
                    paramInfo = Array(Array("@RETURN_VALUE",adInteger	,adParamReturnValue	,,0) _
                        ,Array("@SellSite" 				, adVarchar		, adParamInput		, 	32, Trim("ssg"))	_
                        ,Array("@OutMallOrderSerial"	, adVarchar		, adParamInput		,	32, Trim(ordNo)) _
                        ,Array("@OrgDetailKey"			, adVarchar		, adParamInput		,   32, Trim(ordItemSeq)) _

                        ,Array("@confirmDt"				, adVarchar     , adParamInput		,	16, Trim(confirmDt)) _
                        ,Array("@shppNo"				, adVarchar		, adParamInput		,   32, Trim(shppNo)) _
                        ,Array("@shppSeq"				, adVarchar		, adParamInput		,   10, Trim(shppSeq)) _
                        ,Array("@reOrderYn"				, adVarchar		, adParamInput		,    1, Trim(reOrderYn)) _
                        ,Array("@delayNts"			    , adInteger		, adParamInput		,     , Trim(delayNts)) _
                        ,Array("@splVenItemId"			, adInteger		, adParamInput		,     , Trim(splVenItemId)) _
                        ,Array("@outMallGoodsNo"		, adVarchar		, adParamInput		,   16, Trim(itemId)) _
                        ,Array("@outMallOptionNo"		, adVarchar		, adParamInput		,   32, Trim(uitemId)) _
                        ,Array("@ordQty"			    , adInteger		, adParamInput		,     , Trim(ordQty)) _
                        ,Array("@shppDivDtlNm"		    , adVarchar		, adParamInput		,   20, Trim(shppDivDtlNm)) _
                        ,Array("@uitemNm"		        , adVarchar		, adParamInput		,   128, Trim(uitemNm)) _
                        ,Array("@shppRsvtDt"			, adDate		, adParamInput		,	  , Trim(shppRsvtDt)) _
                        ,Array("@whoutCritnDt"			, adDate		, adParamInput		,	  , Trim(whoutCritnDt)) _
                        ,Array("@autoShortgYn"			, adVarchar		, adParamInput		,    1, Trim(autoShortgYn)) _
                        ,Array("@outorderstatus"		, adVarchar		, adParamInput		,   30, Trim(lastShppProgStatDtlNm)) _

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

			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing


	rw "건수:"&successCnt
    rw "======================================"
end function

%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
