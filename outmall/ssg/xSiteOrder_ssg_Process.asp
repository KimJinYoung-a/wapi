<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/ssg/ssgItemcls.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
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

Dim IS_TEST_MODE : IS_TEST_MODE=FALSE '' SaveOrderToDB

Dim istyyyymmdd, iedyyyymmdd
    iedyyyymmdd = replace(LEFT(now(),10),"-","")
    istyyyymmdd = replace(dateadd("d",-7,LEFT(iedyyyymmdd,4)&"-"&Mid(iedyyyymmdd,5,2)&"-"&Mid(iedyyyymmdd,7,2)),"-","")

'response.write istyyyymmdd&":"&iedyyyymmdd
'response.end

'istyyyymmdd = "20180226"
'iedyyyymmdd = "20180226"

''if (istyyyymmdd<"20180219") then istyyyymmdd="20180219" ''2018/02/19  ''설 연휴라 내역이 많아서 입력이 안되는듯하여 하루씩 입력하였음.


isOnlyTodayBaljuView = (request("rcvtp")="1") '' justView
isDlvConfirmProc = (request("rcvtp")="2") '주문확인 먼저.
isDlvInputProc   = (request("rcvtp")="3") '주문입력

if (request("targetdt")<>"") then
	istyyyymmdd = request("targetdt")
	iedyyyymmdd = istyyyymmdd
end if

if (isOnlyTodayBaljuView) then  ''주문확인한 내역 View
	call getSsgDlvConfirmList(iedyyyymmdd,iedyyyymmdd)
elseif (isDlvInputProc) then    ''주문확인한 내역만가져와 주문입력
    call getSsgDlvConfirmList(iedyyyymmdd,iedyyyymmdd)
elseif (isDlvConfirmProc) then  ''주문미확인건 확인처리Proc
	call getSsgDlvReqList(istyyyymmdd,iedyyyymmdd)
else
    call getSsgDlvReqList(istyyyymmdd,iedyyyymmdd)     '' 7일간 //ActConfirmDlvReq 를 너무 많이 호출하면 SSL 관련 오류가 발생 확인처리를 다 한 후에 이곳을 주석처리후 주문을 땡겨오면 가능..
    call getSsgDlvConfirmList(iedyyyymmdd,iedyyyymmdd) '' 주문확인일 기준.(당일 확인한것으로 해도 무방.)  // 오류발생시 iedyyyymmdd 값을 조정.
end if

''call getSsgDlvConfirmList(istyyyymmdd,iedyyyymmdd)

''배송지시목록 조회
public function getSsgDlvReqList(styyyymmdd,edyyyymmdd)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim ordNo, ordItemSeq, orOrdNo, orordItemSeq, shppNo, shppSeq, ordStatCd, shppStatCd, shppStatNm, itemId, itemNm, splVenItemId
    Dim ordCstId, ordCstOccCd, shppcst, shppcstCodYn, ordRcpDts, ordpeNm, rcptpeNm, rcptpeHpno, rcptpeTelno, shppDivDtlCd, shppProgStatDtlCd, shppRsvtDt
    Dim uitemId, uitemNm, siteNo, rsvtItemYn, frgShppYn, dircItemQty, cnclItemQty, ordQty, splprc, sellprc, ordCmplDts, ordpeHpno
    Dim shpplocAddr, shpplocZipcd, shpplocOldZipcd, ordMemoCntt, ordpeRoadAddr, ordShpplocId, shppTypeDtlCd, reOrderYn, itemDiv, shpplocBascAddr, shpplocDtlAddr, ordItemDivNm

    Dim ArrShppNo, ArrShppSeq, ArrshppStatCd

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

'    			    ordStatCd       = LagrgeNode(i).SelectSingleNode("ordStatCd").Text              ''          [120]
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
							response.write "<textarea cols=100 rows=30>!!"&objXML.responseText&"</textarea>"
						End If

					End If
    			    ordCstId        = LagrgeNode(i).SelectSingleNode("ordCstId").Text                ''주문비용아이디
    			    ordCstOccCd     = LagrgeNode(i).SelectSingleNode("ordCstOccCd").Text             ''주문비용발생코드 [부과] :: 01,02가 아님
    			    shppcst         = LagrgeNode(i).SelectSingleNode("shppcst").Text                 ''배송비?
    			    shppcstCodYn    = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text            ''배송비착불여부 : Y :착불,N :선불 [N]
    			    ordRcpDts       = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text            ''주문접수일시 [2017-11-27 09:32:31.0]
					ordpeNm			= LagrgeNode(i).SelectSingleNode("ordpeNm").Text				 ''주문자
					rcptpeNm 		= LagrgeNode(i).SelectSingleNode("rcptpeNm").Text				 ''수령인

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
					' If NOT (LagrgeNode(i).SelectSingleNode("frgShppYn") is Nothing) then
    			    ' 	frgShppYn       = LagrgeNode(i).SelectSingleNode("frgShppYn").Text                 ''국내/외 구분 [N]
					' End If

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
					if NOT (LagrgeNode(i).SelectSingleNode("shpplocBascAddr") is Nothing) then
						shpplocBascAddr = LagrgeNode(i).SelectSingleNode("shpplocBascAddr").Text                 '' [서울 성동구 매봉길]
					End If
					if NOT (LagrgeNode(i).SelectSingleNode("shpplocDtlAddr") is Nothing) then
						shpplocDtlAddr  = LEFT(LagrgeNode(i).SelectSingleNode("shpplocDtlAddr").Text, 500)                 ''[15, 104동 103호 (옥수동, 래미안옥수리버젠)]
					End If
    			    ordItemDivNm    = LagrgeNode(i).SelectSingleNode("ordItemDivNm").Text                 ''[주문]

    			    if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
    			        ordMemoCntt     = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[고객배송메모]","")                 ''[[고객배송메모]배송메세지]
    			    end if

    			    if NOT (LagrgeNode(i).SelectSingleNode("uitemNm") is Nothing) then
    			        uitemNm         = LagrgeNode(i).SelectSingleNode("uitemNm").Text                 ''[,1:^:asdasd:^:주문문구1,2:^:asdasdddd:^:주문문구2]
    			    end if

			    Next

			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing

	dim newOrdCnt, succConfirmCnt, irefErrStr
	newOrdCnt = 0 : succConfirmCnt = 0 : irefErrStr = ""
	if IsArray(ArrShppNo) then
	    for i=LBound(ArrShppNo) to UBound(ArrShppNo)
	        response.flush

	        '' response.write ArrShppNo(i)&":"&ArrShppSeq(i)
	        if (ArrShppNo(i)<>"") and (ArrShppSeq(i)<>"") then
	            if (ArrshppStatCd(i)="10") then
    	            if (ActConfirmDlvReq(ArrShppNo(i),ArrShppSeq(i),irefErrStr)) then
    	                succConfirmCnt=succConfirmCnt+1
    	            else
    	                irefErrStr = irefErrStr&":"&ArrShppNo(i)&":"&ArrShppSeq(i)&"::"
    	            end if
    	            newOrdCnt = newOrdCnt+1
    	        else
	                rw "대기주문:"&ArrshppStatCd(i)&":"&ArrShppNo(i)
	            end if

	        end if
	    next
	end if

    rw "========================================="
	rw "신규주문확인:"&styyyymmdd&"~"&edyyyymmdd
	rw "신규주문:"&newOrdCnt&"(건)"
	rw "발주확인:"&succConfirmCnt&"(건)"
	if (irefErrStr<>"") then
	    rw irefErrStr
	end if
end function

''주문확인처리
public function ActConfirmDlvReq(iShppno, iShppSeq, byref iErrStr)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/updateOrderSubjectManage.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestOrderSubjectManage>"
    requestBody = requestBoDy&"<shppNo>"&iShppno&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&iShppSeq&"</shppSeq>"
    requestBody = requestBoDy&"</requestOrderSubjectManage>"

	objXML.send(requestBody)

	'rw objXML.status
	'response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
		Set xmlDOM = nothing
	Set objXML = nothing

	'response.write 	ssgresultCode&":"&ssgresultMessage&":"&ssgresultMessage&":"&ssgresultDesc

	if (ssgresultCode<>"00") then
	    iErrStr = "["&ssgresultMessage&"]"&ssgresultDesc
	end if
	ActConfirmDlvReq = (ssgresultCode="00")
end function

''촐고 대상목록 조회
public function getSsgDlvConfirmList(styyyymmdd,edyyyymmdd)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim shppNo,shppSeq,shppTabProgStatCd,evntSeq,shppDivDtlCd,shppDivDtlNm,reOrderYn,delayNts,ordNo,ordItemSeq,ordCmplDts
    Dim lastShppProgStatDtlNm,lastShppProgStatDtlCd,salestrNo,shppVenId,shppVenNm,shppTypeNm,shppTypeCd,shppTypeDtlCd,shppTypeDtlNm,delicoVenId,boxNo
    Dim shppcst,shppcstCodYn,itemNm,splVenItemId,itemId,uitemId,dircItemQty,cnclItemQty,ordQty,sellprc,frgShppYn
    Dim ordpeNm,rcptpeNm,rcptpeHpno,rcptpeTelno,shpplocAddr,shpplocZipcd,shpplocOldZipcd,shpplocRoadAddr,itemChrctDivCd,shppStatCd,shppStatNm
    Dim orordNo,orordItemSeq,shppMainCd,siteNo,siteNm,shppRsvtDt,splprc,shortgYn,newWblNoData,newRow,itemDiv
    Dim shpplocBascAddr,shpplocDtlAddr,ordItemDivNm
    Dim ordpeHpno, ordMemoCntt, pCus, frebieNm ,shortgProgStatCd, shortgProgStatNm, uitemNm
    Dim iBufrequireDetail

    Dim oMaster, oDetailArr(0)
    Dim successCnt : successCnt=0
    Dim failCnt : failCnt=0

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listWarehouseOut.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestWarehouseOut>"
    requestBody = requestBoDy&"<perdType>01</perdType>"
    requestBody = requestBoDy&"<perdStrDts>"&styyyymmdd&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&edyyyymmdd&"</perdEndDts>"  ''하루를 더해야?
    requestBody = requestBoDy&"</requestWarehouseOut>"
	objXML.send(requestBody)

'rw objXML.status
if (isOnlyTodayBaljuView) then
    response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"
    'response.end
end if
    dim retBody : retBody=objXML.responseText
    retBody = replace(retBody,"&","")
	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(retBody) ''objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text

			Set LagrgeNode = xmlDOM.SelectNodes("/result/warehouseOuts/warehouseOut")
			If Not (LagrgeNode Is Nothing) Then
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
'                    salestrNo           = LagrgeNode(i).SelectSingleNode("salestrNo").Text              '' [6004]
                    shppVenId           = LagrgeNode(i).SelectSingleNode("shppVenId").Text      ''공급업체아이디 [0000003198]
                    shppVenNm           = LagrgeNode(i).SelectSingleNode("shppVenNm").Text      ''공급업체명
                    shppTypeNm          = LagrgeNode(i).SelectSingleNode("shppTypeNm").Text     ''배송유형명    [택배배송]
                    shppTypeCd          = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text     ''배송유형코드 10 자사배송 20 택배배송 30 매장방문 40 등기 50 미배송 60 미발송
                    shppTypeDtlCd       = LagrgeNode(i).SelectSingleNode("shppTypeDtlCd").Text  ''배송유형상세코드 14 업체자사배송 22 업체택배배송 25 해외택배배송 31 매장방문 41 등기 51 SMS 52 EMAIL 61 미발송
                    shppTypeDtlNm       = LagrgeNode(i).SelectSingleNode("shppTypeDtlNm").Text  ''배송유형상세명 [업체택배배송]
					if NOT (LagrgeNode(i).SelectSingleNode("delicoVenId") is Nothing) then
                    	delicoVenId         = LagrgeNode(i).SelectSingleNode("delicoVenId").Text    ''택배사ID [0000033011]
					End If
                    boxNo               = LagrgeNode(i).SelectSingleNode("boxNo").Text          ''박스번호 [398327952]
                    shppcst             = LagrgeNode(i).SelectSingleNode("shppcst").Text        '' 배송비? [303] ??
                    shppcstCodYn        = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text   ''*배송비 착불여부 Y: 착불 N: 선불
                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*상품명
					itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*상품번호 [1000024811163]
					if NOT (LagrgeNode(i).SelectSingleNode("splVenItemId") is Nothing) then
	                    splVenItemId        = LagrgeNode(i).SelectSingleNode("splVenItemId").Text       ''*업체상품번호 [1024019]
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

                    uitemId             = LagrgeNode(i).SelectSingleNode("uitemId").Text            ''*단품ID [00000]
                    dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''지시수량 [2]
                    cnclItemQty         = LagrgeNode(i).SelectSingleNode("cnclItemQty").Text        ''취소수량 [0]
                    ordQty              = LagrgeNode(i).SelectSingleNode("ordQty").Text             ''주문수량 [2]
                    sellprc             = LagrgeNode(i).SelectSingleNode("sellprc").Text            ''판매가 [1000]
					If NOT (LagrgeNode(i).SelectSingleNode("frgShppYn") is Nothing) then
    			    	frgShppYn       = LagrgeNode(i).SelectSingleNode("frgShppYn").Text          ''국내/외 구분 [국내]
					End If
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
'                    shppRsvtDt          = LagrgeNode(i).SelectSingleNode("shppRsvtDt").Text
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

                    if (orordNo<>ordNo) then ''원주문번호로 업데이트 ''부분출고처리시 주문번호가 바뀜
                        ordNo=orordNo
                    end if

                    if (orordItemSeq<>ordItemSeq) then  ''2018/03/05 추가  <ordItemDivNm>부분배송주문</ordItemDivNm> 20180305585498
                        ordItemSeq=orordItemSeq
                    end if

                    ''주문입력.
                    Set oMaster = new COrderMasterItem
                        oMaster.FSellSite 			= CMALLNAME  ''ssgcom
                        oMaster.FOutMallOrderSerial = ordNo

                        oMaster.FbeasongNum11st     = shppNo     ''배송번호
                        oMaster.Freserve01          = shppSeq    ''배송번호-seq

                        oMaster.FSellDate 			= Left(ordCmplDts, 19)
                		oMaster.FPayType			= "50"
                		oMaster.FPaydate			= oMaster.FSellDate
                		oMaster.FOrderUserID		= ""
                		oMaster.FOrderName			= LEFT(html2db(ordpeNm),32) '' 주문번호 20180106923841
                		oMaster.FOrderTelNo			= ""
                		oMaster.FOrderHpNo			= ordpeHpno
                		oMaster.FOrderEmail			= ""
                		oMaster.FReceiveName		= LEFT(html2db(rcptpeNm),32)
                		oMaster.FReceiveTelNo		= html2db(rcptpeTelno)
                		oMaster.FReceiveHpNo		= html2db(rcptpeHpno)

                		oMaster.Fdeliverymemo		= html2db(ordMemoCntt)
                		oMaster.FdeliverPay			= shppcst
                        ''if (oMaster.FdeliverPay>0) then oMaster.FdeliverPay=2500 ''배송비가 안분되어 들어감 //아래서 처리

                		oMaster.FReceiveZipCode		= shpplocZipcd
		                oMaster.FReceiveAddr1		= html2db(shpplocBascAddr)
			            oMaster.FReceiveAddr2	    = html2db(shpplocDtlAddr)

                		Set oDetailArr(0) = new COrderDetail
                		oDetailArr(0).FdetailSeq = ordItemSeq
                		oDetailArr(0).FItemID = splVenItemId    ''업체상품번호
                		oDetailArr(0).FItemOption = ""
                		oDetailArr(0).FOutMallItemID = itemId           ''ssg 상품코드
                		oDetailArr(0).FOutMallItemOption = uitemId      ''ssg 단품코드
                		oDetailArr(0).FOutMallItemName = html2db(itemNm)
                		oDetailArr(0).FOutMallItemOptionName = (uitemNm)  ''옵션명이 없음? html2db(objMasterOneXML.attributes.GetNamedItem("RequestOption").text)

                		oDetailArr(0).FItemNo = CLng(dircItemQty) ''주문수량-ordQty, 지시수량- dircItemQty

                		oDetailArr(0).Fitemcost = CLNG(sellprc) '' 단가 맞음. 단가인지 확인 Clng(objMasterOneXML.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("AwardAmount").text) / oDetailArr(0).FItemNo

                		oDetailArr(0).FReducedPrice = oDetailArr(0).Fitemcost  ''쿠폰 할인값(실판매가) 확인. ToDo :: 없는듯함.
                		oDetailArr(0).FOutMallCouponPrice = 0
                		oDetailArr(0).FTenCouponPrice = 0


                		if oDetailArr(0).FOutMallItemOption = "00000" then
                			oDetailArr(0).FItemOption = "0000"
                			''oDetailArr(0).FOutMallItemOption = "0000"  ''주석처리 2018/03/06
                		end if

                        '' ToDo : 옵션이 있는경우 : 옵션명으로 옵션을 가져와야함..
                		if ((oDetailArr(0).FItemOption <> "0000") or (oDetailArr(0).FOutMallItemOptionName<>"")) then
                		    oDetailArr(0).FItemOption = getOptionCodByOptionNameSSG(oDetailArr(0).FItemID,oDetailArr(0).FOutMallItemOptionName,iBufrequireDetail, uitemId)

							'2020-05-14 김진영..&amp;를 replace하지 않은 이유는 위에 body &을 공백으로 처리해서임
							iBufrequireDetail = replace(iBufrequireDetail, "amp;", "&")
							iBufrequireDetail = replace(iBufrequireDetail, "quot;", """")
							iBufrequireDetail = replace(iBufrequireDetail, "lt;", "<")
							iBufrequireDetail = replace(iBufrequireDetail, "gt;", ">")

             		        oDetailArr(0).FrequireDetail = iBufrequireDetail
'                			if Not GetCheckItemOptionValid(oDetailArr(0).FItemID, oDetailArr(0).FItemOption) then
'                				'// 잘못된 옵션.
'                				tmpOptionSeq = tmpOptionSeq + 1
'                				oDetailArr(0).FItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
'                				oDetailArr(0).FOutMallItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
'                			end if
                		end if
 ''rw "oDetailArr(0).FItemOption:"&oDetailArr(0).FItemOption&":"&oDetailArr(0).FOutMallItemOptionName
                        ''rw "---------------------------------------"
                        ''실디비 입력.


                        IF (isOnlyTodayBaljuView) then
                            response.write oMaster.FOutMallOrderSerial&":"&shppDivDtlCd&":"&oMaster.FbeasongNum11st&":"&oDetailArr(0).FdetailSeq&"<br>"
                        else
                            if NOT ((shppDivDtlCd="11") or (shppDivDtlCd="12")) then
                                rw oMaster.FOutMallOrderSerial&":"&shppDivDtlCd&":CS"
                            else
                                if (SaveOrderToDB(oMaster, oDetailArr) = True) then
                                    successCnt = successCnt + 1

                                    '// siteNo 저장 => 아래 프로시져에 포함
        							' strSql = " update db_temp.[dbo].[tbl_xSite_TMPOrder] "
        							' strSql = strSql + " set subSellSite = '" & siteNo & "' "
        							' strSql = strSql + " where sellsite = '" & CMALLNAME & "' "
        							' strSql = strSql + " and OutMallOrderSerial = '" + CStr(oMaster.FOutMallOrderSerial) + "' "
        							' strSql = strSql + " and OrgDetailKey = '" + CStr(oDetailArr(0).FdetailSeq) + "' "
        							' ''response.write strSql
        							' dbget.Execute strSql

									''배송비를 합해서 다시 넣자
									strSql = " Exec db_temp.[dbo].[sp_TEN_xSite_TMPOrder_ssg_DlvPayUp] '" + CStr(oMaster.FOutMallOrderSerial) + "','"&CStr(oDetailArr(0).FdetailSeq)&"','"&siteNo&"',"&oMaster.FdeliverPay
									dbget.Execute strSql
                                else
                                    failCnt = failCnt + 1
                                end if
                            end if
                        end if

                        SEt oDetailArr(0) = Nothing
                    SEt oMaster = Nothing
			    Next
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing

	rw "======================================"
	rw "총미출고건수:"&failCnt+successCnt
	rw "주문입력건수:"&successCnt
end function

''품절/가격 오류체크
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->