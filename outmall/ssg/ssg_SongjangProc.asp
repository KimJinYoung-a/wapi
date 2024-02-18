<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbDatamartOpen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/ssg/ssgItemcls.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body bgcolor="#F4F4F4" >

<%
''ToDo 4. 설치/제작 등 택배사를 이용하지 않는 경우, 직접 배송완료 처리를 해야 합니다. //2019/10/08 , 추적이 안되는케이스도.

'' TLS 1.2를 지원하지 않는 서버가 있는듯함..
const Option_TLS12 = 2048
const Option_TLS1 = 512
const Option_TLS = 128

'response.write "TT"
'response.end


'' 1. 운송장등록
'' 2. 출고처리
'' 3.
'dim IsAutoScript : IsAutoScript =false

dim shppNo  : shppNo = requestCheckvar(request("shppNo"),32)              ''배송번호
dim shppSeq : shppSeq = requestCheckvar(request("shppSeq"),10)            ''배송순번
dim wblno   : wblno = TRIM(requestCheckvar(replace(request("wblno"),"-",""),20))                ''운송장번호
dim delicoVenId : delicoVenId = requestCheckvar(request("delicoVenId"),20)    ''택배사번호
dim itemno : itemno = requestCheckvar(request("itemno"),10)    ''출고수량
dim outmallorderserial : outmallorderserial = requestCheckvar(request("outmallorderserial"),20)    ''제휴주문번호.
dim orgdetailKey : orgdetailKey = requestCheckvar(request("orgdetailKey"),20)    ''제휴주문상세번호.

dim prctp : prctp = requestCheckvar(request("prctp"),20)    ''처리 Action (1:운송장등록, )

dim dlvfinishdt : dlvfinishdt = requestCheckvar(request("dlvfinishdt"),10)
if (prctp="") then prctp="1"

dim mode      : mode=request("mode")
dim tenorderserial : tenorderserial=request("tenorderserial")
dim tenitemid : tenitemid=request("tenitemid")
dim tenitemoption : tenitemoption=request("tenitemoption")



Dim sqlStr, iAssignedRow
If mode = "updateSendState" Then

	sqlStr = "Update db_temp.dbo.tbl_xSite_TMPOrder "&VBCRLF
	sqlStr = sqlStr & "	Set sendState='"&requestCheckvar(request("updateSendState"),10)&"'"&VBCRLF
	sqlStr = sqlStr & "	,sendReqCnt=sendReqCnt+1"&VBCRLF
	sqlStr = sqlStr & "	where OutMallOrderSerial='"&outmallorderserial&"'"&VBCRLF
	sqlStr = sqlStr & "	and beasongNum11st='"&shppNo&"'"&VBCRLF
	sqlStr = sqlStr & "	and reserve01='"&shppSeq&"'"&VBCRLF
	sqlStr = sqlStr & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
	sqlStr = sqlStr & "	and sellsite='ssg'"


    'response.write sqlStr
	dbget.Execute sqlStr,iAssignedRow
	response.write "<script>alert('"&iAssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If


''TEST
'shppNo      = "D2126007228" ''"D2126019324"
'shppSeq     = "2"
'wblno       = "340566203190"
'delicoVenId = TenDlvCode2SSGDlvCode("4")
'itemno      = 2
'outmallorderserial = "20171123128379"
'orgdetailKey = "2"
''response.write request.serverVariables("REMOTE_ADDR")
'if ( request.serverVariables("REMOTE_ADDR")="110.93.128.113") or ( request.serverVariables("REMOTE_ADDR")="110.93.128.99")  then
'    response.write "잠시대기"
'    dbget.close() : response.end
'end if

'if (outmallorderserial="__20180710269242")  then
'     response.write "잠시대기"
'    dbget.close() : response.end
'end if

'' 주문을 나눠 입력하는 케이스.
Dim GG_ORG_outmallorderserial : GG_ORG_outmallorderserial = outmallorderserial
IF (InStr(outmallorderserial,"_")>0) then
	outmallorderserial = getOutmallRefOrgOrderNO(outmallorderserial,orgdetailKey,"ssg")
end if

if (delicoVenId="기타") then
    delicoVenId="0000033028" ''기타택배사
end if

if (delicoVenId="0000033052") and (LEFT(wblno,1)="1") then ''우체국택배 / 1로 시작.
    delicoVenId="0000033051" ''우체국 등기
end if

if (prctp="1") then  ''운송장 등록
    if (IsPortionSongjangSendRequireNAssign(outmallorderserial,orgdetailKey,shppNo,shppSeq,wblno)) then ''부분출고가 필요한경우
        response.write "부분출고 처리 후 송장입력PROC<br>"
        if ((delicoVenId="0000033028") and NOT IsNumeric(wblno)) then
            wblno=RIGHT(outmallorderserial,12)
        elseif ((delicoVenId="0000033028") and LEN(wblno)<11) then
            wblno=wblno&RIGHT(outmallorderserial,6)
        end if

        ''주문번호에 영문이 추가되었음.
        if NOT(IsNumeric(wblno)) then
            if IsNumeric(LEFT(wblno,LEN(wblno)-6)) and NOT IsNumeric(MID(wblno,LEN(wblno)-6+1,6)) then
                wblno = LEFT(wblno,LEN(wblno)-6) & cdbl("&H"&MID(wblno,LEN(wblno)-6+1,6))
            end if
        end if
        Call saveSSGWblNo(shppNo,shppSeq,wblno,delicoVenId,outmallorderserial,orgdetailKey, IsAutoScript)
    else
        response.write "송장입력PROC<br>"
        if ((delicoVenId="0000033028") and NOT IsNumeric(wblno)) then
            wblno=RIGHT(outmallorderserial,12)
        elseif ((delicoVenId="0000033028") and LEN(wblno)<11) then
            wblno=wblno&RIGHT(outmallorderserial,6)
        end if

        ''rw LEFT(wblno,LEN(wblno)-5)&":"&MID(wblno,LEN(wblno)-5+1,5)
        ''주문번호에 영문이 추가되었음.
'rw  wblno
'rw  IsNumeric(wblno)
        if NOT(IsNumeric(wblno)) then
            if IsNumeric(LEFT(wblno,LEN(wblno)-6)) and NOT IsNumeric(MID(wblno,LEN(wblno)-6+1,6)) then
                on error resume next
                wblno = LEFT(wblno,LEN(wblno)-6) & cdbl("&H"&MID(wblno,LEN(wblno)-6+1,6))
                if err.number <> 0 then
                    wblno = getNumeric(wblno)
                    err.clear
                end if
                on error goto 0
            end if
        end if
'rw  wblno
'response.end
        Call saveSSGWblNo(shppNo,shppSeq,wblno,delicoVenId,outmallorderserial,orgdetailKey, IsAutoScript)
    end if
elseif (prctp="2") then ''출고처리
    Call Savessgwhoutcompleteprocess(shppNo,shppSeq, itemno, outmallorderserial,orgdetailKey,wblno, IsAutoScript)
elseif (prctp="3") then ''미배송완료 목록 조회

    sqlStr = "select top(1) beasongNum11st,reserve01,OrgDetailKey, isNULL(sendSongjangNo,'') as sendSongjangNo "
    sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder WITH(NOLOCK)"
    sqlStr = sqlStr & " where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
    sqlStr = sqlStr & " and orderserial='"&tenorderserial&"'"&VBCRLF
    sqlStr = sqlStr & " and matchitemid='"&tenitemid&"'"&VBCRLF
    sqlStr = sqlStr & " and matchitemoption='"&tenitemoption&"'"&VBCRLF

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        shppNo          = rsget("beasongNum11st")
        shppSeq         = rsget("reserve01")
        orgdetailKey    = rsget("OrgDetailKey")
        wblno           = rsget("sendSongjangNo")
    end if
    rsget.close

    if (shppNo<>"") and (shppSeq<>"") and (orgdetailKey<>"") then

        if (ChecklistNonDeliveryByShppNo(shppNo,shppSeq,outmallorderserial,orgdetailKey,wblno, IsAutoScript)>0) then
            Call saveDeliveryEnd(shppNo,shppSeq,outmallorderserial,orgdetailKey,IsAutoScript) ''배송완료처리
        else
            rw "None"
        end if
    else
        rw "ERR"
    end if
elseif (prctp="33") then ''미배송완료 CS 목록 조회

    sqlStr = " select top 1 shppNo, shppSeq, OrgDetailKey "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " [db_temp].[dbo].[tbl_xSite_TMPMiChulList] c "
    sqlStr = sqlStr & " where "
    sqlStr = sqlStr & " 	1 = 1 "
    sqlStr = sqlStr & " 	and c.OutMallOrderSerial = '"&GG_ORG_outmallorderserial&"' "
    sqlStr = sqlStr & " 	and c.Matchorderserial = '"&tenorderserial&"' "
    sqlStr = sqlStr & " 	and c.matchItemID = '"&tenitemid&"' "
    sqlStr = sqlStr & " 	and c.matchitemoption = '"&tenitemoption&"' "

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        shppNo          = rsget("shppNo")
        shppSeq         = rsget("shppSeq")
        orgdetailKey    = rsget("shppSeq")
    end if
    rsget.close

    if (shppNo<>"") and (shppSeq<>"") and (orgdetailKey<>"") then

        Call saveDeliveryEndCS(shppNo,shppSeq,outmallorderserial,orgdetailKey,IsAutoScript) ''배송완료처리
    else
        rw "ERR"
    end if
elseif (prctp="5") then ''부분출고처리
    rw "ERR"
    'Call saveSSGPortionWarehouseOutProcess(outmallorderserial,orgdetailKey,shppNo,shppSeq,itemno)
    'CAll getSsgChulgoListByOutmallOrderserialUpshppNo(outmallorderserial)
elseif (prctp="6") then ''주문확인
    Call saveSsgOrderChulgoConfirm(shppNo,shppSeq)
elseif (prctp="9") then ''결품처리 (출고불가)

elseif (prctp="999") then ''출고처리대상주문 작성.
    Call makeSSGChulgoTargetMake()
else
    response.write "미지정:"&prctp
    dbget.close():response.end
end if


''부분출고 처리 : 부분출고 하면. 동일 배송번호의 운송장번호가 빈값으로 바뀌고, 배송번호도 바뀜? /주문번호도 새로따짐, 원주문번호는 바뀌지 않음/  (부분출고 입력한것의?) [실제출고처리가 아님?] => 다시 배송번호를 조회
''상품단에서 출고처가 따로 존재해야할듯함.
''이걸쓰면 출고되는게 아니라 운송장번호가 없어지고, 배송번호가 바뀌는것임.
''Call saveSSGPortionWarehouseOutProcess(shppNo,shppSeq,itemno)

''주문번호로 배송목록조회 후 배송번호업데이트  : 배송번호를 다시가져오기 위함.
''CAll getSsgChulgoListByOutmallOrderserialUpshppNo(outmallorderserial)

''출고처리
''Call saveSSGWhOutCompleteProcess(shppNo,shppSeq, itemno, outmallorderserial,orgdetailKey)

''배송완료처리 ''택배는 자동으로 된다.?

''부분출고 필요한지 체크
public function IsPortionSongjangSendRequireNAssign(outmallorderserial,orgdetailKey,shppNo,ishppSeq,wblno)
    Dim strSql, ArrRows, i
    dim isPartitalSendReq : isPartitalSendReq = FALSE
    IsPortionSongjangSendRequireNAssign = FALSE

'    strSql = "select beasongNum11st, sendsongjangno"&vbCRLF
'    strSql = strSql & " from db_temp.dbo.tbl_xsite_TmpOrder"&vbCRLF
'    strSql = strSql & " where outmallorderserial='"&outmallorderserial&"'"&vbCRLF
'    strSql = strSql & " and orgdetailkey<>'"&orgdetailKey&"'"&vbCRLF
'    strSql = strSql & " and beasongNum11st='"&shppNo&"'"&vbCRLF
'    strSql = strSql & " and isNULL(sendsongjangno,'"&wblno&"')<>'"&wblno&"'"&vbCRLF

    strSql = "exec [db_temp].[dbo].[sp_TEN_xSite_PartialOutCheck_SSG] '"&outmallorderserial&"','"&orgdetailKey&"','"&shppNo&"','"&wblno&"'"


    rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        isPartitalSendReq = true
        arrRows = rsget.getRows
    end if
    rsget.close

    if (isPartitalSendReq) then
        response.write "부분출고처리PROC1"
        if isArray(arrRows) then
            For i = 0 To UBound(ArrRows,2)
                Call saveSSGPortionWarehouseOutProcess(ArrRows(0,i),ArrRows(1,i),ArrRows(2,i),ArrRows(3,i),ArrRows(4,i))
            Next
        end if

        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
    	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
        strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
        strSql = strSql & "	and beasongNum11st='"&shppNo&"'"&VBCRLF
        strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
        strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
        strSql = strSql & "	and matchstate in ('O','C','Q','A')"

    	dbget.Execute strSql


        CAll getSsgChulgoListByOutmallOrderserialUpshppNo(outmallorderserial)
    end if

    IsPortionSongjangSendRequireNAssign =isPartitalSendReq
end function

'// 주문확인
public function saveSsgOrderChulgoConfirm(shppNo, shppSeq)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim errCount

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/updateOrderSubjectManage.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestOrderSubjectManage>"
    requestBody = requestBoDy&"<shppNo>"&shppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&shppSeq&"</shppSeq>"
    requestBody = requestBoDy&"</requestOrderSubjectManage>"

	objXML.send(requestBody)
	''rw objXML.status
	''response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text

	Set objXML = nothing

	rw ssgresultCode&":"&ssgresultMessage&":"&ssgresultDesc
    rw shppNo&":"&shppSeq
end function

''부분출고 처리
public function saveSSGPortionWarehouseOutProcess(outmallorderserial,orgdetailKey,ishppNo,ishppSeq, iitemno)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim errCount
    'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/savePortionWarehouseOutProcess.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestWhOutCompleteProcess>"
    requestBody = requestBoDy&"<shppNo>"&ishppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&ishppSeq&"</shppSeq>"
    requestBody = requestBoDy&"<procItemQty>"&iitemno&"</procItemQty>"
    requestBody = requestBoDy&"</requestWhOutCompleteProcess>"
''response.write "<textarea cols=60 rows=10>"&requestBody&"</textarea>"
	objXML.send(requestBody)

'response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text
			Set LagrgeNode = nothing
	Set objXML = nothing

	rw 	ssgresultCode&":"&ssgresultMessage&":"&ssgresultDesc

	if (ssgresultCode<>"00") then
	    strSql = ""
    	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
    	strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"
    	strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"
    	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
    	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
    	strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
    	rsget.CursorLocation = adUseClient
        rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
    	If Not rsget.Eof Then
    		errCount = rsget("cnt")
    	End If
    	rsget.Close

    	If errCount > 0 Then
    	    response.write "오류회수 초과."

    		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
    						"	<option value=''>선택</option>" &_
    						"	<option value='951'>기전송 내역</option>" &_
    						"	<option value='952'>취소주문</option>" &_
    						"</select>&nbsp;&nbsp;"
    		response.write "<input type='button' value='완료처리' onClick=""fnSetSendState('"&GG_ORG_outmallorderserial&"','"&ishppNo&"','"&ishppSeq&"','"&orgdetailKey&"',document.getElementById('updateSendState').value)"">"
    		response.write "<script language='javascript'>"&VbCRLF
    		response.write "function fnSetSendState(outmallorderserial,ishppNo,ishppSeq,iorgdetailKey,selectValue){"&VbCRLF
    		response.write "    if(selectValue == ''){"&VbCRLF
    		response.write "    	alert('선택해주세요');"&VbCRLF
    		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
    		response.write "    	return;"&VbCRLF
    		response.write "    }"&VbCRLF
    		response.write "    var uri = 'ssg_SongjangProc.asp?mode=updateSendState&outmallorderserial='+outmallorderserial+'&shppNo='+ishppNo+'&shppSeq='+ishppSeq+'&orgdetailKey='+iorgdetailKey+'&updateSendState='+selectValue;"&VbCRLF
    		response.write "    location.replace(uri);"&VbCRLF
    		response.write "}"&VbCRLF
    		response.write "</script>"&VbCRLF
    	End If
	end if

end function

''배송완료된 목록
public function CheckOrderFinishedAssign(ishppNo,ishppSeq, outmallorderserial,orgdetailKey, iwblno,IsAutoScript)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim errCount : errCount = 0
    Dim AssignedCNT : AssignedCNT=0
    Dim successCnt, AssignedRow : AssignedRow=0
    Dim perdStrDts, perdEndDts

    Dim ordNo, ordItemSeq, orOrdNo, orordItemSeq
    Dim reOrderYn, shppNo, shppSeq
    Dim shppDt, whoutDt, settlRfltDts, lastShppProgStatDtlCd, wblNo, shppTypeCd
    Dim tenSendstate : tenSendstate = "0"

    CheckOrderFinishedAssign = 0

    perdStrDts = replace(LEFT(CStr(dateadd("d",-14,NOW())),10),"-","")
    perdEndDts = replace(LEFT(CStr(dateadd("d",+1,NOW())),10),"-","")

    'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listDeliveryEnd.ssg"   ''배송완료목록

	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestDeliveryEnd>"
    requestBody = requestBoDy&"<perdType>02</perdType>"  ''02 출고완료일
    requestBody = requestBoDy&"<perdStrDts>"&perdStrDts&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&perdEndDts&"</perdEndDts>"
    requestBody = requestBoDy&"<commType>01</commType>"  ''01 원주문번호 / 02 배송번호
    requestBody = requestBoDy&"<commValue>"&outmallorderserial&"</commValue>"  ''빠른조회값, 빠른조회일 경우 필수 값
    requestBody = requestBoDy&"</requestDeliveryEnd>"
	objXML.send(requestBody)

	Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
		xmlDOM.async = False
		xmlDOM.loadXML(objXML.responseText)
		ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
		ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
		ssgresultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text
'rw "<textarea cols=20 rows=10>"&objXML.responseText&"</textarea>"
	    Set LagrgeNode = xmlDOM.SelectNodes("/result/deliveryEnds/deliveryEnd")
			If Not (LagrgeNode Is Nothing) Then
			    For i = 0 To LagrgeNode.length - 1
			        ''변수초기화.
			        ordNo = ""
			        ordItemSeq = ""
                    orOrdNo = ""
                    orordItemSeq = ""
                    reOrderYn = ""
                    shppNo = ""
                    shppSeq = ""
                    shppDt = ""
                    whoutDt = ""
                    lastShppProgStatDtlCd = ""
                    shppTypeCd =""
                    wblNo = ""

                    ordNo           = LagrgeNode(i).SelectSingleNode("ordNo").Text '*주문번호 [20171123128379]
			        ordItemSeq      = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text ''*주문순번
			        if NOT (LagrgeNode(i).SelectSingleNode("orOrdNo") is Nothing) then
                        orOrdNo         = LagrgeNode(i).SelectSingleNode("orOrdNo").Text '*원주문번호 [20171123128379]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("orordItemSeq") is Nothing) then
                        orordItemSeq    = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text ''원주문순번 [2]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("shppTypeCd") is Nothing) then
                        shppTypeCd    = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text ''배송유형코드
                    end if
                    reOrderYn       = LagrgeNode(i).SelectSingleNode("reOrderYn").Text '' 재주문여부
                    shppNo          = LagrgeNode(i).SelectSingleNode("shppNo").Text '' 배송번호
                    shppSeq         = LagrgeNode(i).SelectSingleNode("shppSeq").Text '' 배송순번
                    shppDt          = LagrgeNode(i).SelectSingleNode("shppDt").Text '' 배송완료일시

                    if NOT (LagrgeNode(i).SelectSingleNode("whoutDt") is Nothing) then
                        whoutDt         = LagrgeNode(i).SelectSingleNode("whoutDt").Text '' 출고완료일시
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("wblNo") is Nothing) then
                        wblNo    = LagrgeNode(i).SelectSingleNode("wblNo").Text ''송장번호
                    end if
                    ''settlRfltDts    = LagrgeNode(i).SelectSingleNode("shppNo").Text  '' 정산반영일시
                    lastShppProgStatDtlCd  = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''최종배송상세진행상태코드 11 배송지시, 21 피킹지시, 22 피킹완료, 31 패킹완료, 41 출고보류, 42 출고지연, 43 출고완료, 51 배송완료, 52 배송거절



                    if (lastShppProgStatDtlCd="43" or lastShppProgStatDtlCd="51") then
                        if ((ordNo=outmallorderserial) and (ordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) _
                            or ((orOrdNo=outmallorderserial) and (orordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) _
                            or ((orOrdNo="") and (orordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) then

                            if (shppTypeCd="10") then
                                rw "shppTypeCd:자사배송"
                            elseif (shppTypeCd="20") then
                                rw "shppTypeCd:택배배송:"&wblNo
                            else
                                rw "shppTypeCd:"&shppTypeCd
                            end if

                            if (lastShppProgStatDtlCd="22") then
                                rw "피킹완료 상태"
                                tenSendstate = "0"
                            elseif (lastShppProgStatDtlCd="43") then
                                rw "출고완료 상태"
                                tenSendstate = "3"
                            elseif (lastShppProgStatDtlCd="51") then
                                rw "배송완료 상태"
                                tenSendstate = "4"
                            end if

                            rw ordNo&"|"&ordItemSeq&"|"&orOrdNo&"|"&orordItemSeq&"|"&reOrderYn&"|"&shppNo&"|"&shppSeq&"|"&shppDt&"|"&whoutDt&"|"&lastShppProgStatDtlCd&"|"&wblNo
                            rw outmallorderserial&"|"&orgdetailKey&"|"&ishppNo&"|"&ishppSeq&"|"&iwblno

                            ''if (wblNo=iwblno) or (wblNo="") then ''wblNo="" 자사배송
                                strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
                            	strSql = strSql & "	Set sendstate='"&tenSendstate&"'" ''배송완료는 4번으로 하자.
                                strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
                                strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
                            	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
                            	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
                                strSql = strSql & "	and matchstate in ('O','A')"  ''A 추가 2019/10/08
                                strSql = strSql & "	and sendstate in (0,2)"
                            	dbget.Execute strSql,AssignedRow

                                successCnt = successCnt + AssignedRow
                            ''end if
                        else
                            if (shppNo=ishppNo) then
                                rw "-------------------------------------"
                                rw ordNo&"|"&ordItemSeq&"|"&orOrdNo&"|"&orordItemSeq&"|"&reOrderYn&"|"&shppNo&"|"&shppSeq&"|"&shppDt&"|"&whoutDt&"|"&lastShppProgStatDtlCd&"|"&wblNo
                                rw outmallorderserial&"|"&orgdetailKey&"|"&ishppNo&"|"&ishppSeq&"|"&iwblno
                                rw "-------------------------------------"
                            end if
                        end if
                    end if


			    Next
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing
	CheckOrderFinishedAssign = successCnt
end function

''미 배송완료된 목록(출고완료 상태?)
public function CheckOrderSendedAssign(ishppNo,ishppSeq, outmallorderserial,orgdetailKey, iwblno,IsAutoScript)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim errCount : errCount = 0
    Dim AssignedCNT : AssignedCNT=0
    Dim AssignedRow, successCnt : successCnt=0
    Dim perdStrDts, perdEndDts

    Dim ordNo, ordItemSeq, orOrdNo, orordItemSeq
    Dim reOrderYn, shppNo, shppSeq
    Dim shppDt, whoutDt, settlRfltDts, lastShppProgStatDtlCd, wblNo, shppTypeCd
    Dim tenSendstate : tenSendstate="0"

    CheckOrderSendedAssign = 0

    perdStrDts = replace(LEFT(CStr(dateadd("d",-14,NOW())),10),"-","")
    perdEndDts = replace(LEFT(CStr(dateadd("d",+1,NOW())),10),"-","")

    'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listNonDelivery.ssg"   ''미 배송완료목록

	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestNonDelivery>"
    requestBody = requestBoDy&"<perdType>01</perdType>"  ''01 출고완료일 / 02 결제완료일
    requestBody = requestBoDy&"<perdStrDts>"&perdStrDts&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&perdEndDts&"</perdEndDts>"
    requestBody = requestBoDy&"<commType>01</commType>"  ''01 원주문번호 / 02 배송번호
    requestBody = requestBoDy&"<commValue>"&outmallorderserial&"</commValue>"  ''빠른조회값, 빠른조회일 경우 필수 값
    requestBody = requestBoDy&"</requestNonDelivery>"
	objXML.send(requestBody)

	Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
		xmlDOM.async = False
		xmlDOM.loadXML(objXML.responseText)
		ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
		ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
		ssgresultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text

	    Set LagrgeNode = xmlDOM.SelectNodes("/result/nonDeliverys/nonDelivery")
			If Not (LagrgeNode Is Nothing) Then
			    For i = 0 To LagrgeNode.length - 1
			        ''변수초기화.
			        ordNo = ""
			        ordItemSeq = ""
                    orOrdNo = ""
                    orordItemSeq = ""
                    reOrderYn = ""
                    shppNo = ""
                    shppSeq = ""
                    shppDt = ""
                    whoutDt = ""
                    lastShppProgStatDtlCd = ""
                    shppTypeCd = ""
                    wblNo = ""


                    ordNo           = LagrgeNode(i).SelectSingleNode("ordNo").Text '*주문번호 [20171123128379]
			        ordItemSeq      = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text ''*주문순번
			        if NOT (LagrgeNode(i).SelectSingleNode("orOrdNo") is Nothing) then
                        orOrdNo         = LagrgeNode(i).SelectSingleNode("orOrdNo").Text '*원주문번호 [20171123128379]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("orordItemSeq") is Nothing) then
                        orordItemSeq    = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text ''원주문순번 [2]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("shppTypeCd") is Nothing) then
                        shppTypeCd    = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text ''배송유형코드
                    end if

                    reOrderYn       = LagrgeNode(i).SelectSingleNode("reOrderYn").Text '' 재주문여부
                    shppNo          = LagrgeNode(i).SelectSingleNode("shppNo").Text '' 배송번호
                    shppSeq         = LagrgeNode(i).SelectSingleNode("shppSeq").Text '' 배송순번

                    ''shppDt          = LagrgeNode(i).SelectSingleNode("shppDt").Text '' 배송완료일시
                    ' whoutDt         = LagrgeNode(i).SelectSingleNode("whoutDt").Text '' 출고완료일시
                    ''settlRfltDts    = LagrgeNode(i).SelectSingleNode("shppNo").Text  '' 정산반영일시
                    lastShppProgStatDtlCd  = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''최종배송상세진행상태코드 11 배송지시, 21 피킹지시, 22 피킹완료, 31 패킹완료, 41 출고보류, 42 출고지연, 43 출고완료, 51 배송완료, 52 배송거절
                    if NOT (LagrgeNode(i).SelectSingleNode("wblNo") is Nothing) then
                        wblNo    = LagrgeNode(i).SelectSingleNode("wblNo").Text ''송장번호
                    end if

                    ''rw ordNo&"|"&ordItemSeq&"|"&orOrdNo&"|"&orordItemSeq&"|"&reOrderYn&"|"&shppNo&"|"&shppSeq&"|"&shppDt&"|"&whoutDt&"|"&lastShppProgStatDtlCd&"|"&wblNo
                    ''rw outmallorderserial&"|"&orgdetailKey&"|"&ishppNo&"|"&ishppSeq&"|"&iwblno


                    '' 43:출고완료,51:배송완료,(22:피킹완료)
                    if (lastShppProgStatDtlCd="43" or lastShppProgStatDtlCd="51") then
                        ''if (ordNo=outmallorderserial) and (ordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq) then
                        if ((ordNo=outmallorderserial) and (ordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) _
                            or ((orOrdNo=outmallorderserial) and (orordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq))  _
                            or ((orOrdNo="") and (orordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) then

                            if (shppTypeCd="10") then
                                rw "shppTypeCd:자사배송"
                            elseif (shppTypeCd="20") then
                                rw "shppTypeCd:택배배송:"&wblNo
                            else
                                rw "shppTypeCd:"&shppTypeCd
                            end if

                            if (lastShppProgStatDtlCd="22") then
                                rw "피킹완료 상태"
                                tenSendstate = "0"
                            elseif (lastShppProgStatDtlCd="43") then
                                rw "출고완료 상태"
                                tenSendstate = "3"
                            elseif (lastShppProgStatDtlCd="51") then
                                rw "배송완료 상태"
                                tenSendstate = "4"
                            end if

                            rw ordNo&"|"&ordItemSeq&"|"&orOrdNo&"|"&orordItemSeq&"|"&reOrderYn&"|"&shppNo&"|"&shppSeq&"|"&shppDt&"|"&whoutDt&"|"&lastShppProgStatDtlCd&"|"&wblNo
                            rw outmallorderserial&"|"&orgdetailKey&"|"&ishppNo&"|"&ishppSeq&"|"&iwblno

                            ''if (wblNo=iwblno) then
                                strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
                            	strSql = strSql & "	Set sendstate='"&tenSendstate&"'"
                                strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
                                strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
                            	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
                            	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
                                strSql = strSql & "	and matchstate in ('O')"
                                strSql = strSql & "	and sendstate in (0,2)"
                            	dbget.Execute strSql,AssignedRow

                                successCnt = successCnt + AssignedRow
                            ''end if
                        else
                            if (shppNo=ishppNo) then
                                rw "-------------------------------------"
                                rw ordNo&"|"&ordItemSeq&"|"&orOrdNo&"|"&orordItemSeq&"|"&reOrderYn&"|"&shppNo&"|"&shppSeq&"|"&shppDt&"|"&whoutDt&"|"&lastShppProgStatDtlCd&"|"&wblNo
                                rw outmallorderserial&"|"&orgdetailKey&"|"&ishppNo&"|"&ishppSeq&"|"&iwblno
                                rw "-------------------------------------"
                            end if
                        end if
                    end if

			    Next
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing
	CheckOrderSendedAssign = successCnt
end function

''미 배송완료된 목록(출고완료 상태) 배송번호로 조회 2019/10/16
public function ChecklistNonDeliveryByShppNo(ishppNo,ishppSeq, outmallorderserial,orgdetailKey, iwblno,IsAutoScript)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim errCount : errCount = 0
    Dim AssignedCNT : AssignedCNT=0
    Dim AssignedRow, successCnt : successCnt=0
    Dim perdStrDts, perdEndDts

    Dim ordNo, ordItemSeq, orOrdNo, orordItemSeq
    Dim reOrderYn, shppNo, shppSeq
    Dim shppDt, whoutDt, settlRfltDts, lastShppProgStatDtlCd, wblNo, shppTypeCd
    Dim shppTypeDtlNm, delicoVenNm
    Dim tenSendstate : tenSendstate="0"

    ChecklistNonDeliveryByShppNo = 0

    perdStrDts = replace(LEFT(CStr(dateadd("d",-14,NOW())),10),"-","")
    perdEndDts = replace(LEFT(CStr(dateadd("d",+1,NOW())),10),"-","")

    'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listNonDelivery.ssg"   ''미 배송완료목록

	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestNonDelivery>"
    requestBody = requestBoDy&"<perdType>01</perdType>"  ''01 출고완료일 / 02 결제완료일
    requestBody = requestBoDy&"<perdStrDts>"&perdStrDts&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&perdEndDts&"</perdEndDts>"
    requestBody = requestBoDy&"<commType>02</commType>"  ''01 원주문번호 / 02 배송번호
    requestBody = requestBoDy&"<commValue>"&ishppNo&"</commValue>"  ''빠른조회값, 빠른조회일 경우 필수 값
    requestBody = requestBoDy&"</requestNonDelivery>"
	objXML.send(requestBody)

	Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
		xmlDOM.async = False
		xmlDOM.loadXML(objXML.responseText)
		ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
		ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
		ssgresultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text

	    Set LagrgeNode = xmlDOM.SelectNodes("/result/nonDeliverys/nonDelivery")
			If Not (LagrgeNode Is Nothing) Then
			    For i = 0 To LagrgeNode.length - 1
			        ''변수초기화.
			        ordNo = ""
			        ordItemSeq = ""
                    orOrdNo = ""
                    orordItemSeq = ""
                    reOrderYn = ""
                    shppNo = ""
                    shppSeq = ""
                    shppDt = ""
                    whoutDt = ""
                    lastShppProgStatDtlCd = ""
                    shppTypeCd = ""
                    wblNo = ""
                    shppTypeDtlNm = ""
                    delicoVenNm =""


                    ordNo           = LagrgeNode(i).SelectSingleNode("ordNo").Text '*주문번호 [20171123128379]
			        ordItemSeq      = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text ''*주문순번
			        if NOT (LagrgeNode(i).SelectSingleNode("orOrdNo") is Nothing) then
                        orOrdNo         = LagrgeNode(i).SelectSingleNode("orOrdNo").Text '*원주문번호 [20171123128379]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("orordItemSeq") is Nothing) then
                        orordItemSeq    = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text ''원주문순번 [2]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("shppTypeCd") is Nothing) then
                        shppTypeCd    = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text ''배송유형코드
                    end if

                    reOrderYn       = LagrgeNode(i).SelectSingleNode("reOrderYn").Text '' 재주문여부
                    shppNo          = LagrgeNode(i).SelectSingleNode("shppNo").Text '' 배송번호
                    shppSeq         = LagrgeNode(i).SelectSingleNode("shppSeq").Text '' 배송순번

                    ''shppDt          = LagrgeNode(i).SelectSingleNode("shppDt").Text '' 배송완료일시
                    ' whoutDt         = LagrgeNode(i).SelectSingleNode("whoutDt").Text '' 출고완료일시
                    ''settlRfltDts    = LagrgeNode(i).SelectSingleNode("shppNo").Text  '' 정산반영일시
                    lastShppProgStatDtlCd  = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''최종배송상세진행상태코드 11 배송지시, 21 피킹지시, 22 피킹완료, 31 패킹완료, 41 출고보류, 42 출고지연, 43 출고완료, 51 배송완료, 52 배송거절

                    if (orOrdNo<>"") and (ordNo<>orOrdNo) then ordNo=orOrdNo
                    if (orordItemSeq<>"") and (ordItemSeq<>orordItemSeq) then ordItemSeq=orordItemSeq

                    if NOT (LagrgeNode(i).SelectSingleNode("delicoVenNm") is Nothing) then
                        delicoVenNm    = LagrgeNode(i).SelectSingleNode("delicoVenNm").Text ''택배사
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("wblNo") is Nothing) then
                        wblNo    = LagrgeNode(i).SelectSingleNode("wblNo").Text ''송장번호
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("shppTypeDtlNm") is Nothing) then
                        shppTypeDtlNm    = LagrgeNode(i).SelectSingleNode("shppTypeDtlNm").Text ''배송유형상세명
                    end if


                    ' rw ordNo&"|"&ordItemSeq&"|"&orOrdNo&"|"&orordItemSeq&"|"&reOrderYn&"|"&shppNo&"|"&shppSeq&"|"&shppDt&"|"&whoutDt&"|"&lastShppProgStatDtlCd&"|"&wblNo
                    ' rw outmallorderserial&"|"&orgdetailKey&"|"&ishppNo&"|"&ishppSeq&"|"&shppTypeDtlNm&"|"&iwblno


                    '' 43:출고완료,51:배송완료,(22:피킹완료)
                    if (lastShppProgStatDtlCd="43" or lastShppProgStatDtlCd="51") then
                        if ((ordNo=outmallorderserial) and (ordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) _
                            or ((orOrdNo=outmallorderserial) and (orordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq))  _
                            or ((orOrdNo="") and (orordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) then

                            if (shppTypeCd="10") then
                                rw "shppTypeCd:자사배송"
                            elseif (shppTypeCd="20") then
                                rw "shppTypeCd:택배배송:"&wblNo
                            else
                                rw "shppTypeCd:"&shppTypeCd
                            end if

                            if (lastShppProgStatDtlCd="22") then
                                rw "피킹완료 상태"
                                tenSendstate = "0"
                            elseif (lastShppProgStatDtlCd="43") then
                                rw "출고완료 상태"
                                tenSendstate = "3"
                            elseif (lastShppProgStatDtlCd="51") then
                                rw "배송완료 상태"
                                tenSendstate = "4"
                            end if

                            rw ordNo&"|"&ordItemSeq&"|"&orOrdNo&"|"&orordItemSeq&"|"&reOrderYn&"|"&shppNo&"|"&shppSeq&"|"&shppDt&"|"&whoutDt&"|"&lastShppProgStatDtlCd&"|"&delicoVenNm&"|"&shppTypeDtlNm&"|"&wblNo
                            rw outmallorderserial&"|"&orgdetailKey&"|"&ishppNo&"|"&ishppSeq&"|"&iwblno

                            successCnt = successCnt + 1
                            ' ''if (wblNo=iwblno) then
                            '     strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
                            ' 	strSql = strSql & "	Set sendstate='"&tenSendstate&"'"
                            '     strSql = strSql & "	where OutMallOrderSerial='"&outmallorderserial&"'"&VBCRLF
                            '     strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
                            ' 	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
                            ' 	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
                            '     strSql = strSql & "	and matchstate in ('O')"
                            '     strSql = strSql & "	and sendstate in (0,2)"
                            ' 	dbget.Execute strSql,AssignedRow

                            '     successCnt = successCnt + AssignedRow
                            ' ''end if
                        else
                            if (shppNo=ishppNo) then
                                rw "-------------------------------------"
                                rw ordNo&"|"&ordItemSeq&"|"&orOrdNo&"|"&orordItemSeq&"|"&reOrderYn&"|"&shppNo&"|"&shppSeq&"|"&shppDt&"|"&whoutDt&"|"&lastShppProgStatDtlCd&"|"&delicoVenNm&"|"&shppTypeDtlNm&"|"&wblNo
                                rw outmallorderserial&"|"&orgdetailKey&"|"&ishppNo&"|"&ishppSeq&"|"&iwblno
                                rw "-------------------------------------"
                            end if
                        end if
                    end if

			    Next
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing
	ChecklistNonDeliveryByShppNo = successCnt
end function

''상품출고 처리
public function saveSSGWhOutCompleteProcess(ishppNo,ishppSeq, iitemno,outmallorderserial,orgdetailKey, iwblno, IsAutoScript)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim errCount : errCount = 0
    Dim AssignedCNT : AssignedCNT=0

    'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/saveWhOutCompleteProcess.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestWhOutCompleteProcess>"
    requestBody = requestBoDy&"<shppNo>"&ishppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&ishppSeq&"</shppSeq>"
    requestBody = requestBoDy&"<procItemQty>"&iitemno&"</procItemQty>"
    requestBody = requestBoDy&"</requestWhOutCompleteProcess>"

	objXML.send(requestBody)
'rw objXML.status
'response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
            if (xmlDOM.getElementsByTagName("resultDesc").length > 0) then  ''2019-09-04 11:38 오류 증가로 수정
			    ssgresultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text
            else
                ssgresultDesc = " nothing"
            end if
			Set LagrgeNode = nothing
	Set objXML = nothing

	rw 	ssgresultCode&":"&ssgresultMessage&":"&ssgresultDesc

	response.flush
	if (ssgresultCode="00") then
        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
    	strSql = strSql & "	Set sendstate=3"
        strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
        strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
    	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
    	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
        strSql = strSql & "	and matchstate in ('O')"
        strSql = strSql & "	and sendstate=2"
    	dbget.Execute strSql,AssignedCNT

        IF (AssignedCNT>0) then
    	    if (IsAutoScript) then
    	        rw "OK|"&outmallorderserial&" "&orgdetailKey
    	    ELSE
        	    response.write "OK"
        	ENd IF
        ENd IF
    else
        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
    	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
        strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
        strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
        strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
        strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
        strSql = strSql & "	and matchstate in ('O','C','Q','A')"

    	dbget.Execute strSql

        rw "<font color=red>"&ssgresultMessage&":"&ssgresultDesc&"</font>"

        rw outmallorderserial
        rw ishppNo
        rw ishppSeq


    	'만약 에러횟수가 3회가 넘으면 수기처리 가능
    	'updateSendState = 951		기전송 내역
    	'updateSendState = 952		취소주문

    	strSql = ""
    	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder WITH(NOLOCK) " & VBCRLF
    	strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"
    	strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"
    	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
    	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
    	strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
    	rsget.CursorLocation = adUseClient
        rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
    	If Not rsget.Eof Then
    		errCount = rsget("cnt")
    	End If
    	rsget.Close

        Dim reChkCnt : reChkCnt=0
        if (errCount>0) then
            if (ssgresultDesc<>"해당 데이터가 없습니다.") then '' 취소일개연성이 많음.
                reChkCnt = CheckOrderSendedAssign(ishppNo,ishppSeq, outmallorderserial,orgdetailKey, iwblno,IsAutoScript)
                if (reChkCnt<1) then
                    reChkCnt = CheckOrderFinishedAssign(ishppNo,ishppSeq, outmallorderserial,orgdetailKey, iwblno, IsAutoScript)
                end if
                rw reChkCnt
                if reChkCnt>0 then errCount=0  ''이미 배송완료 진행 되었음.
            end if
        end if

    	If errCount > 0 Then
    	    response.write "오류회수 초과."

    		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
    						"	<option value=''>선택</option>" &_
    						"	<option value='951'>기전송 내역</option>" &_
    						"	<option value='952'>취소주문</option>" &_
    						"</select>&nbsp;&nbsp;"
    		response.write "<input type='button' value='완료처리' onClick=""fnSetSendState('"&GG_ORG_outmallorderserial&"','"&ishppNo&"','"&ishppSeq&"','"&orgdetailKey&"',document.getElementById('updateSendState').value)"">"
    		response.write "<script language='javascript'>"&VbCRLF
    		response.write "function fnSetSendState(outmallorderserial,ishppNo,ishppSeq,iorgdetailKey,selectValue){"&VbCRLF
    		response.write "    if(selectValue == ''){"&VbCRLF
    		response.write "    	alert('선택해주세요');"&VbCRLF
    		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
    		response.write "    	return;"&VbCRLF
    		response.write "    }"&VbCRLF
    		response.write "    var uri = 'ssg_SongjangProc.asp?mode=updateSendState&outmallorderserial='+outmallorderserial+'&shppNo='+ishppNo+'&shppSeq='+ishppSeq+'&orgdetailKey='+iorgdetailKey+'&updateSendState='+selectValue;"&VbCRLF
    		response.write "    location.replace(uri);"&VbCRLF
    		response.write "}"&VbCRLF
    		response.write "</script>"&VbCRLF
    	End If

    end if
end function

'' 운송장 등록 API // 운송장등록API 와 출고처리 API가 별개이다, 출고처리도 부분출고처리/전체출고가 있음;;
'' 운송장은 [배송번호] 별로 하나씩 등록 가능하다? - 중복송장불가? == 업체배송은 배송지를 별도로 해야할듯?
'' 하나넣어도 송번호가 같으면 같이 배들어감
public function saveSSGWblNo(ishppNo,ishppSeq, iwblno, idelicoVenId, outmallorderserial, orgdetailKey, IsAutoScript)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim shppTypeCd: shppTypeCd="20"         ''택배배송
    Dim shppTypeDtlCd: shppTypeDtlCd="22"   ''업체택배배송
    Dim AssignedCNT

    'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/saveWblNo.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestWhOutCompleteProcess>"
    requestBody = requestBoDy&"<shppNo>"&ishppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&ishppSeq&"</shppSeq>"
    requestBody = requestBoDy&"<wblNo>"&iwblno&"</wblNo>"
    requestBody = requestBoDy&"<delicoVenId>"&idelicoVenId&"</delicoVenId>"
    requestBody = requestBoDy&"<shppTypeCd>"&shppTypeCd&"</shppTypeCd>"
    requestBody = requestBoDy&"<shppTypeDtlCd>"&shppTypeDtlCd&"</shppTypeDtlCd>"
    requestBody = requestBoDy&"</requestWhOutCompleteProcess>"

	objXML.send(requestBody)
'rw objXML.status
'response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"
	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			ssgresultDesc = Trim(xmlDOM.getElementsByTagName("resultDesc").Item(0).Text)
			Set LagrgeNode = nothing
	Set objXML = nothing

	rw 	ssgresultCode&":"&ssgresultMessage&":"&ssgresultDesc

	if (ssgresultCode="00" AND ssgresultDesc="성공") then
        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "&VBCRLF
    	strSql = strSql & "	Set sendState=1"&VBCRLF
    	strSql = strSql & "	,sendReqCnt=1"&VBCRLF ''sendReqCnt+1 =>1 , 1로 초기화
    	strSql = strSql & "	,sendSongjangNo='"&iwblno&"'"&VBCRLF        ''2017/01/03 추가
        strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
        strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
    	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
    	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
        'strSql = strSql & "	and matchstate in ('O')"
        strSql = strSql & "	and matchstate in ('O', 'A')" '2019-05-16 김진영 A도 추가
    	dbget.Execute strSql,AssignedCNT
    else
        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
    	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
        strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
        strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
        strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
        strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
        strSql = strSql & "	and matchstate in ('O','C','Q','A')"

    	dbget.Execute strSql

        rw "<font color=red>"&ssgresultMessage&":"&ssgresultDesc&"</font>"

        rw outmallorderserial
        rw ishppNo
        rw ishppSeq

        Dim errCount
        strSql = ""
    	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder WITH(NOLOCK)" & VBCRLF
    	strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"
    	strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"
    	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
    	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
    	strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
    	rsget.CursorLocation = adUseClient
        rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
    	If Not rsget.Eof Then
    		errCount = rsget("cnt")
    	End If
    	rsget.Close

        ''ssgresultDesc
        ''운송장 번호가 유효하지 않습니다.
        ''기등록된 운송장번호입니다.
        Dim reChkCnt : reChkCnt=0
        Dim IsreqForceSend : IsreqForceSend = false
        if (errCount>0) and (NOT IsAutoScript) then
            reChkCnt = getSsgChulgoListByshppNo(ishppNo,ishppSeq, outmallorderserial,orgdetailKey, iwblno, IsreqForceSend, IsAutoScript)

            if (IsreqForceSend) then ''배송기일이 임박하여 자가배송으로 송장입력처리 한다..
                call saveSSGWblNoForce(ishppNo,ishppSeq, iwblno, idelicoVenId, outmallorderserial, orgdetailKey, IsAutoScript)
                Exit function
            end if

            if (ssgresultDesc<>"해당 데이터가 없습니다.") then '' 취소일개연성이 많음.
                if (reChkCnt<1) then
                    reChkCnt = CheckOrderSendedAssign(ishppNo,ishppSeq, outmallorderserial,orgdetailKey, iwblno,IsAutoScript)  '' iitemno=>1 not using
                    if (reChkCnt<1) then
                        reChkCnt = CheckOrderFinishedAssign(ishppNo,ishppSeq, outmallorderserial,orgdetailKey, iwblno, IsAutoScript)
                    end if
                    rw reChkCnt
                    if reChkCnt>0 then errCount=0  ''이미 배송완료 진행 되었음.
                end if
            end if
        end if

    	If errCount > 0 Then
    	    response.write "오류회수 초과.."

    		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
    						"	<option value=''>선택</option>" &_
    						"	<option value='951'>기전송 내역</option>" &_
    						"	<option value='952'>취소주문</option>" &_
    						"</select>&nbsp;&nbsp;"
    		response.write "<input type='button' value='완료처리' onClick=""fnSetSendState('"&GG_ORG_outmallorderserial&"','"&ishppNo&"','"&ishppSeq&"','"&orgdetailKey&"',document.getElementById('updateSendState').value)"">"
    		response.write "<script language='javascript'>"&VbCRLF
    		response.write "function fnSetSendState(outmallorderserial,ishppNo,ishppSeq,iorgdetailKey,selectValue){"&VbCRLF
    		response.write "    if(selectValue == ''){"&VbCRLF
    		response.write "    	alert('선택해주세요');"&VbCRLF
    		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
    		response.write "    	return;"&VbCRLF
    		response.write "    }"&VbCRLF
    		response.write "    var uri = 'ssg_SongjangProc.asp?mode=updateSendState&outmallorderserial='+outmallorderserial+'&shppNo='+ishppNo+'&shppSeq='+ishppSeq+'&orgdetailKey='+iorgdetailKey+'&updateSendState='+selectValue;"&VbCRLF
    		response.write "    location.replace(uri);"&VbCRLF
    		response.write "}"&VbCRLF
    		response.write "</script>"&VbCRLF
    	End If
    end if
end function

public function saveDeliveryEnd(ishppNo,ishppSeq, outmallorderserial, orgdetailKey, IsAutoScript)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim shppTypeCd: shppTypeCd="10"         ''자사배송
    Dim shppTypeDtlCd: shppTypeDtlCd="14"   ''업체자사배송
    Dim AssignedCNT

    'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/saveDeliveryEnd.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestDeliveryEnd>"
    requestBody = requestBoDy&"<shppNo>"&ishppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&ishppSeq&"</shppSeq>"
    'requestBody = requestBoDy&"<wblNo>"&iwblno&"</wblNo>"                      ''자사배송인경우 넣지 않는다.
    'requestBody = requestBoDy&"<delicoVenId>"&idelicoVenId&"</delicoVenId>"    ''자사배송인경우 넣지 않는다.
    'requestBody = requestBoDy&"<shppTypeCd>"&shppTypeCd&"</shppTypeCd>"
    'requestBody = requestBoDy&"<shppTypeDtlCd>"&shppTypeDtlCd&"</shppTypeDtlCd>"
    requestBody = requestBoDy&"</requestDeliveryEnd>"

	objXML.send(requestBody)
'rw objXML.status
'response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			ssgresultDesc = Trim(xmlDOM.getElementsByTagName("resultDesc").Item(0).Text)
			Set LagrgeNode = nothing
	Set objXML = nothing

	rw 	"<font color='blue'>배송완료처리:"&ssgresultCode&":"&ssgresultMessage&":"&ssgresultDesc&"<font>"

	if (ssgresultCode="00" AND ssgresultDesc="성공") then
        sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_OUTAMLL_Jungsan_Comment_add] '" & tenorderserial & "', " & tenitemid & ", '" & tenitemoption & "','배송완료진행','"&session("ssBctId")&"'"
		dbDatamart_dbget.Execute sqlStr

        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "&VBCRLF
    	strSql = strSql & "	Set sendState=3"&VBCRLF
        strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
        strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
    	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
    	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
        strSql = strSql & "	and matchstate in ('O', 'A')" '2019-05-16 김진영 A도 추가
    	dbget.Execute strSql,AssignedCNT
    else
        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
    	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
        strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
        strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
        strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
        strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
        strSql = strSql & "	and matchstate in ('O','C','Q','A')"

    	dbget.Execute strSql

        rw "<font color=red>"&ssgresultMessage&":"&ssgresultDesc&"</font>"
    end if
end function

public function saveDeliveryEndCS(ishppNo,ishppSeq, outmallorderserial, orgdetailKey, IsAutoScript)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim shppTypeCd: shppTypeCd="10"         ''자사배송
    Dim shppTypeDtlCd: shppTypeDtlCd="14"   ''업체자사배송
    Dim AssignedCNT

    'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/saveDeliveryEnd.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestDeliveryEnd>"
    requestBody = requestBoDy&"<shppNo>"&ishppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&ishppSeq&"</shppSeq>"
    'requestBody = requestBoDy&"<wblNo>"&iwblno&"</wblNo>"                      ''자사배송인경우 넣지 않는다.
    'requestBody = requestBoDy&"<delicoVenId>"&idelicoVenId&"</delicoVenId>"    ''자사배송인경우 넣지 않는다.
    'requestBody = requestBoDy&"<shppTypeCd>"&shppTypeCd&"</shppTypeCd>"
    'requestBody = requestBoDy&"<shppTypeDtlCd>"&shppTypeDtlCd&"</shppTypeDtlCd>"
    requestBody = requestBoDy&"</requestDeliveryEnd>"

	objXML.send(requestBody)
'rw objXML.status
'response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			ssgresultDesc = Trim(xmlDOM.getElementsByTagName("resultDesc").Item(0).Text)
			Set LagrgeNode = nothing
	Set objXML = nothing

	rw 	"<font color='blue'>배송완료처리:"&ssgresultCode&":"&ssgresultMessage&":"&ssgresultDesc&"<font>"

	if (ssgresultCode="00" AND ssgresultDesc="성공") then
        '// do nothing
    else
        rw "<font color=red>"&ssgresultMessage&":"&ssgresultDesc&"</font>"
    end if
end function

public function saveSSGWblNoForce(ishppNo,ishppSeq, iwblno, idelicoVenId, outmallorderserial, orgdetailKey, IsAutoScript)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim shppTypeCd: shppTypeCd="10"         ''자사배송
    Dim shppTypeDtlCd: shppTypeDtlCd="14"   ''업체자사배송
    Dim AssignedCNT

    'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/saveWblNo.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestWhOutCompleteProcess>"
    requestBody = requestBoDy&"<shppNo>"&ishppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&ishppSeq&"</shppSeq>"
    'requestBody = requestBoDy&"<wblNo>"&iwblno&"</wblNo>"                      ''자사배송인경우 넣지 않는다.
    'requestBody = requestBoDy&"<delicoVenId>"&idelicoVenId&"</delicoVenId>"    ''자사배송인경우 넣지 않는다.
    requestBody = requestBoDy&"<shppTypeCd>"&shppTypeCd&"</shppTypeCd>"
    requestBody = requestBoDy&"<shppTypeDtlCd>"&shppTypeDtlCd&"</shppTypeDtlCd>"
    requestBody = requestBoDy&"</requestWhOutCompleteProcess>"

	objXML.send(requestBody)
'rw objXML.status
'response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			ssgresultDesc = Trim(xmlDOM.getElementsByTagName("resultDesc").Item(0).Text)
			Set LagrgeNode = nothing
	Set objXML = nothing

	rw 	"<font color='blue'>자사배송으로전송:"&ssgresultCode&":"&ssgresultMessage&":"&ssgresultDesc&"<font>"

	if (ssgresultCode="00" AND ssgresultDesc="성공") then
        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "&VBCRLF
    	strSql = strSql & "	Set sendState=1"&VBCRLF
    	strSql = strSql & "	,sendReqCnt=1"&VBCRLF ''sendReqCnt+1 =>1 , 1로 초기화
    	strSql = strSql & "	,sendSongjangNo='자사배송"&iwblno&"'"&VBCRLF        ''2017/01/03 추가
        strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
        strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
    	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
    	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
        'strSql = strSql & "	and matchstate in ('O')"
        strSql = strSql & "	and matchstate in ('O', 'A')" '2019-05-16 김진영 A도 추가
    	dbget.Execute strSql,AssignedCNT
    else
        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
    	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
        strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
        strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
        strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
        strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
        strSql = strSql & "	and matchstate in ('O','C','Q','A')"

    	dbget.Execute strSql

        rw "<font color=red>"&ssgresultMessage&":"&ssgresultDesc&"</font>"
    end if
end function

''촐고 대상목록 조회
public function getSsgChulgoListByOutmallOrderserialUpshppNo(ioutmallorderserial)
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
    Dim AssignedRow : AssignedRow =0
    Dim successCnt : successCnt=0
    Dim failCnt : failCnt=0
    Dim totalcnt: totalcnt=0

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
	objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listWarehouseOut.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestWarehouseOut>"
    requestBody = requestBoDy&"<commType>01</commType>"
    requestBody = requestBoDy&"<commValue>"&ioutmallorderserial&"</commValue>"
    requestBody = requestBoDy&"</requestWarehouseOut>"

	objXML.send(requestBody)
'rw objXML.status
'response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text

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
                    ordNo               = LagrgeNode(i).SelectSingleNode("ordNo").Text                  ''*주문번호 [20171123128379]
                    ordItemSeq          = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text             ''*주문순번
                    orordNo             = LagrgeNode(i).SelectSingleNode("orordNo").Text            ''원주문번호 [20171123128379]

                    On Error Resume Next
                        orordItemSeq    = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''원주문순번 [2]
                        If Err.number <> 0 Then
                            orordItemSeq = ordItemSeq
                        End If
                    On Error Goto 0

                    strSql = "update db_temp.dbo.tbl_xSite_TMPOrder"&vbCRLF
                    strSql = strSql & " set beasongNum11st='"&shppNo&"'"&vbCRLF
                    strSql = strSql & " ,reserve01='"&shppSeq&"'"&vbCRLF
                    strSql = strSql & " where OutMallOrderSerial='"&orordNo&"'"&vbCRLF
                    strSql = strSql & " and OrgDetailKey='"&orordItemSeq&"'"&vbCRLF
                    dbget.Execute strSql, AssignedRow

                    successCnt = successCnt + AssignedRow
                    totalCnt = totalcnt + 1

'                    shppTabProgStatCd   = LagrgeNode(i).SelectSingleNode("shppTabProgStatCd").Text      ''최종배송상세진행상태코드(배송단위) 11 배송지시 21 피킹지시 22 피킹완료 31 패킹완료 41 출고보류 42 출고지연 43 출고완료 51 배송완료 52 배송거절
'                    evntSeq             = LagrgeNode(i).SelectSingleNode("evntSeq").Text                ''이벤트순번
'                    shppDivDtlCd        = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text           ''*배송구분상세코드 11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고
'                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''배송구분상세명
'                    reOrderYn           = LagrgeNode(i).SelectSingleNode("reOrderYn").Text              ''*재지시여부구분
'                    delayNts            = LagrgeNode(i).SelectSingleNode("delayNts").Text               ''지연횟수
'                    ordCmplDts          = LagrgeNode(i).SelectSingleNode("ordCmplDts").Text             ''*주문완료일시 [2017-11-23 10:39:42.0]
'                    lastShppProgStatDtlNm   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlNm").Text  ''최종배송상세진행상태명(배송상품단위) [피킹완료]
'                    lastShppProgStatDtlCd   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''최종배송상세진행상태코드(배송상품단위) 11 배송지시 21 피킹지시 22 피킹완료 31 패킹완료 41 출고보류 42 출고지연 43 출고완료 51 배송완료 52 배송거절
'                    salestrNo           = LagrgeNode(i).SelectSingleNode("salestrNo").Text              '' [6004]
'                    shppVenId           = LagrgeNode(i).SelectSingleNode("shppVenId").Text      ''공급업체아이디 [0000003198]
'                    shppVenNm           = LagrgeNode(i).SelectSingleNode("shppVenNm").Text      ''공급업체명
'                    shppTypeNm          = LagrgeNode(i).SelectSingleNode("shppTypeNm").Text     ''배송유형명    [택배배송]
'                    shppTypeCd          = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text     ''배송유형코드 10 자사배송 20 택배배송 30 매장방문 40 등기 50 미배송 60 미발송
'                    shppTypeDtlCd       = LagrgeNode(i).SelectSingleNode("shppTypeDtlCd").Text  ''배송유형상세코드 14 업체자사배송 22 업체택배배송 25 해외택배배송 31 매장방문 41 등기 51 SMS 52 EMAIL 61 미발송
'                    shppTypeDtlNm       = LagrgeNode(i).SelectSingleNode("shppTypeDtlNm").Text  ''배송유형상세명 [업체택배배송]
'                    delicoVenId         = LagrgeNode(i).SelectSingleNode("delicoVenId").Text    ''택배사ID [0000033011]
'                    boxNo               = LagrgeNode(i).SelectSingleNode("boxNo").Text          ''박스번호 [398327952]
'                    shppcst             = LagrgeNode(i).SelectSingleNode("shppcst").Text        '' 배송비? [303] ??
'                    shppcstCodYn        = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text   ''*배송비 착불여부 Y: 착불 N: 선불
'                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*상품명
'                    splVenItemId        = LagrgeNode(i).SelectSingleNode("splVenItemId").Text       ''*업체상품번호 [1024019]
'                    itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*상품번호 [1000024811163]
'                    uitemId             = LagrgeNode(i).SelectSingleNode("uitemId").Text            ''*단품ID [00000]
'                    dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''지시수량 [2]
'                    cnclItemQty         = LagrgeNode(i).SelectSingleNode("cnclItemQty").Text        ''취소수량 [0]
'                    ordQty              = LagrgeNode(i).SelectSingleNode("ordQty").Text             ''주문수량 [2]
'                    sellprc             = LagrgeNode(i).SelectSingleNode("sellprc").Text            ''판매가 [1000]
'                    frgShppYn           = LagrgeNode(i).SelectSingleNode("frgShppYn").Text          ''국내/외 구분 [국내]
'                    ordpeNm             = LagrgeNode(i).SelectSingleNode("ordpeNm").Text            ''*주문자
'                    rcptpeNm            = LagrgeNode(i).SelectSingleNode("rcptpeNm").Text           ''*수취인
'                    rcptpeHpno          = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text         ''*수취인 휴대폰번호
'                    rcptpeTelno         = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text        ''*수취인 집전화번호
'                    shpplocAddr         = LagrgeNode(i).SelectSingleNode("shpplocAddr").Text        ''수취인 상세주소
'                    shpplocZipcd        = LagrgeNode(i).SelectSingleNode("shpplocZipcd").Text       ''*수취인 우편번호          [04733]
'                    shpplocOldZipcd     = LagrgeNode(i).SelectSingleNode("shpplocOldZipcd").Text    ''*수취인 구우편번호(6자리)  [133750]
'                    shpplocRoadAddr     = LagrgeNode(i).SelectSingleNode("shpplocRoadAddr").Text    ''수취인도로명주소
'                    itemChrctDivCd      = LagrgeNode(i).SelectSingleNode("itemChrctDivCd").Text     ''상품특성구분코드 10 일반 20 몰인몰 30 해외구매대행상품 40 미가공귀금속 50 모바일기프트 60 상품권 70 쇼핑충전금 80 모바일상품권 91 이벤트
'                    shppStatCd          = LagrgeNode(i).SelectSingleNode("shppStatCd").Text         ''*배송상태코드 10 정상 30 대기
'                    shppStatNm          = LagrgeNode(i).SelectSingleNode("shppStatNm").Text         ''배송상태명
'                    shppMainCd          = LagrgeNode(i).SelectSingleNode("shppMainCd").Text         ''배송주체코드 32 업체창고 41 협력업체 42 브랜드직배  [41]
'                    siteNo              = LagrgeNode(i).SelectSingleNode("siteNo").Text             ''사이트번호 6001 이마트몰 6002 트레이더스몰 6003 분스몰 6004 신세계몰 6005 S.COM몰 6009 신세계백화점몰
'                    siteNm              = LagrgeNode(i).SelectSingleNode("siteNm").Text             ''사이트명
'                    shppRsvtDt          = LagrgeNode(i).SelectSingleNode("shppRsvtDt").Text
'                    splprc              = LagrgeNode(i).SelectSingleNode("splprc").Text             ''공급가
'                    shortgYn            = LagrgeNode(i).SelectSingleNode("shortgYn").Text
'                    newWblNoData        = LagrgeNode(i).SelectSingleNode("newWblNoData").Text
'                    newRow              = LagrgeNode(i).SelectSingleNode("newRow").Text
'                    itemDiv             = LagrgeNode(i).SelectSingleNode("itemDiv").Text                ''판매불가신청상태 10:일반 20: 명절 GIFT 일반 30: 명절 GIFT 센터 40: 명절 GIFT 냉장
'                    shpplocBascAddr     = LagrgeNode(i).SelectSingleNode("shpplocBascAddr").Text        ''수취인주소 20170712
'                    shpplocDtlAddr      = LagrgeNode(i).SelectSingleNode("shpplocDtlAddr").Text         ''수취인상세주소	20170712
'                    ordItemDivNm        = LagrgeNode(i).SelectSingleNode("ordItemDivNm").Text           ''주문상품구분	20170809
'
'
'                    ''//필수값 아닌경우 .
'                    if NOT (LagrgeNode(i).SelectSingleNode("ordpeHpno") is Nothing) then
'                        ordpeHpno         = LagrgeNode(i).SelectSingleNode("ordpeHpno").Text           ''주문자휴대폰번호  //선택값
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
'                        ordMemoCntt         = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[고객배송메모]","")            ''고객배송메모  //선택값
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("pCus") is Nothing) then
'                        pCus         = LagrgeNode(i).SelectSingleNode("pCus").Text           ''개인통관고유번호  //선택값
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("frebieNm") is Nothing) then
'                        frebieNm         = LagrgeNode(i).SelectSingleNode("frebieNm").Text    ''사은품  //선택값
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatCd") is Nothing) then
'                        shortgProgStatCd         = LagrgeNode(i).SelectSingleNode("shortgProgStatCd").Text    ''판매불가신청상태  //선택값 11 결품등록 12 결품CS처리중 13 결품확정 21 상품정보오류등록 22 상품정보오류CS처리중 23 상품정보오류확정 41 입고지연등록 43 입고지연완료 51 배송지연등록
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatNm") is Nothing) then
'                        shortgProgStatNm         = LagrgeNode(i).SelectSingleNode("shortgProgStatNm").Text    ''결품진행상태명  //선택값
'                    end if
'
'                    ''옵션명
'                    if NOT (LagrgeNode(i).SelectSingleNode("uitemNm") is Nothing) then
'                        uitemNm         = LagrgeNode(i).SelectSingleNode("uitemNm").Text                 ''[,1:^:asdasd:^:주문문구1,2:^:asdasdddd:^:주문문구2]
'                    end if

			    Next
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing

	rw "======================================"
	rw "총상세건수:"&totalcnt
	rw "배송번호 업데이트:"&successCnt
end function

''촐고 대상목록 조회 by 배송번호
public function getSsgChulgoListByshppNo(ishppNo,ishppSeq, ioutmallorderserial,iorgdetailKey, iwblno,IsreqForceSend, IsAutoScript)
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
    Dim whoutCritnDt
    Dim iBufrequireDetail

    Dim oMaster, oDetailArr(0)
    Dim AssignedRow : AssignedRow =0
    Dim successCnt : successCnt=0
    Dim failCnt : failCnt=0
    Dim totalcnt: totalcnt=0

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
	objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listWarehouseOut.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestWarehouseOut>"
    requestBody = requestBoDy&"<commType>02</commType>"
    requestBody = requestBoDy&"<commValue>"&ishppNo&"</commValue>"
    requestBody = requestBoDy&"</requestWarehouseOut>"

	objXML.send(requestBody)
'rw objXML.status
'response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text

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

                    ordpeHpno = "": ordMemoCntt = "": pCus = "": frebieNm = "": shortgProgStatCd ="": shortgProgStatNm ="" : uitemNm="" : whoutCritnDt=""
                    iBufrequireDetail = ""

			        shppNo              = LagrgeNode(i).SelectSingleNode("shppNo").Text                 ''*배송번호
                    shppSeq             = LagrgeNode(i).SelectSingleNode("shppSeq").Text                ''*배송순번
                    ordNo               = LagrgeNode(i).SelectSingleNode("ordNo").Text                  ''*주문번호 [20171123128379]
                    ordItemSeq          = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text             ''*주문순번
                    orordNo             = LagrgeNode(i).SelectSingleNode("orordNo").Text            ''원주문번호 [20171123128379]
                    orordItemSeq        = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''원주문순번 [2]


                    shppTabProgStatCd   = LagrgeNode(i).SelectSingleNode("shppTabProgStatCd").Text      ''최종배송상세진행상태코드(배송단위) 11 배송지시 21 피킹지시 22 피킹완료 31 패킹완료 41 출고보류 42 출고지연 43 출고완료 51 배송완료 52 배송거절
                    delayNts            = LagrgeNode(i).SelectSingleNode("delayNts").Text               ''지연횟수
                    lastShppProgStatDtlNm   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlNm").Text  ''최종배송상세진행상태명(배송상품단위) [피킹완료]
                    lastShppProgStatDtlCd   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''최종배송상세진행상태코드(배송상품단위) 11 배송지시 21 피킹지시 22 피킹완료 31 패킹완료 41 출고보류 42 출고지연 43 출고완료 51 배송완료 52 배송거절
                    shppTypeNm          = LagrgeNode(i).SelectSingleNode("shppTypeNm").Text     ''배송유형명    [택배배송]
                    shppStatCd          = LagrgeNode(i).SelectSingleNode("shppStatCd").Text         ''*배송상태코드 10 정상 30 대기
                    shppStatNm          = LagrgeNode(i).SelectSingleNode("shppStatNm").Text         ''배송상태명

                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatNm") is Nothing) then
                        shortgProgStatNm         = LagrgeNode(i).SelectSingleNode("shortgProgStatNm").Text    ''결품진행상태명  //선택값
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("whoutCritnDt") is Nothing) then
                        whoutCritnDt         = LagrgeNode(i).SelectSingleNode("whoutCritnDt").Text    ''출고기준일
                    end if


                    if ((ordNo=ioutmallorderserial) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) _
                        or ((orordNo=ioutmallorderserial) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) then

                        rw ordNo&"|"&ordItemSeq&"|"&orordNo&"|"&orordItemSeq&"|"&shppNo&"|"&shppSeq
                        rw shppTabProgStatCd&"|"&delayNts&"|"&lastShppProgStatDtlNm&"|"&lastShppProgStatDtlCd&"|"&shppTypeNm&"|"&shppStatCd&"|"&CHKIIF(shppStatNm="취소","<strong><font color=red>","")&shppStatNm&CHKIIF(shppStatNm="취소","</font></strong>","")&"|"&shortgProgStatNm&"|"&whoutCritnDt
                        successCnt = successCnt + 1 ''AssignedRow

                        if (shppStatNm<>"취소") and (lastShppProgStatDtlNm="피킹완료") then
                            ''출고기준일이 오늘/내일/모레 이라면..
                            if (replace(CSTR(dateadd("d",3,NOW())),"-","")>=whoutCritnDt) then
                                IsreqForceSend = true
                            end if

                            if (NOT IsreqForceSend) and (dlvfinishdt<>"") then
                                ''우리쪽 배송완료일 N일 지난경우라면. => 우리쪽 송장 배송완료이면 전송한다.
                                ''if (DateDiff("d",dlvfinishdt,now())>=2) then
                                    IsreqForceSend = true
                                ''end if
                            end if

                            if (NOT IsreqForceSend) then
                                if (request("isfrcsend")="1") then
                                    IsreqForceSend = true
                                ELSE
                                    Dim reqURI : reqURI="?shppNo="&request("shppNo")&"&shppSeq="&request("shppSeq")&"&delicoVenId="&request("delicoVenId")&"&wblno="&request("wblno")&"&itemno="&request("itemno")&"&outmallorderserial="&request("outmallorderserial")&"&orgdetailKey="&request("orgdetailKey")&"&dlvfinishdt="&request("dlvfinishdt")&"&prctp="&request("prctp")&"&isfrcsend=1"
                                    rw "<br><input type='button' value='자사배송 전송' onClick=""location.href='"&reqURI&"'"">"
                                END IF
                            end if

                        end if
                    end if
                    ' strSql = "update db_temp.dbo.tbl_xSite_TMPOrder"&vbCRLF
                    ' strSql = strSql & " set beasongNum11st='"&shppNo&"'"&vbCRLF
                    ' strSql = strSql & " ,reserve01='"&shppSeq&"'"&vbCRLF
                    ' strSql = strSql & " where OutMallOrderSerial='"&orordNo&"'"&vbCRLF
                    ' strSql = strSql & " and OrgDetailKey='"&orordItemSeq&"'"&vbCRLF
                    ' dbget.Execute strSql, AssignedRow

                    ' successCnt = successCnt + AssignedRow
                    ' totalCnt = totalcnt + 1

'
'                    evntSeq             = LagrgeNode(i).SelectSingleNode("evntSeq").Text                ''이벤트순번
'                    shppDivDtlCd        = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text           ''*배송구분상세코드 11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고
'                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''배송구분상세명
'                    reOrderYn           = LagrgeNode(i).SelectSingleNode("reOrderYn").Text              ''*재지시여부구분
'
'                    ordCmplDts          = LagrgeNode(i).SelectSingleNode("ordCmplDts").Text             ''*주문완료일시 [2017-11-23 10:39:42.0]
'
'
'                    salestrNo           = LagrgeNode(i).SelectSingleNode("salestrNo").Text              '' [6004]
'                    shppVenId           = LagrgeNode(i).SelectSingleNode("shppVenId").Text      ''공급업체아이디 [0000003198]
'                    shppVenNm           = LagrgeNode(i).SelectSingleNode("shppVenNm").Text      ''공급업체명
'
'                    shppTypeCd          = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text     ''배송유형코드 10 자사배송 20 택배배송 30 매장방문 40 등기 50 미배송 60 미발송
'                    shppTypeDtlCd       = LagrgeNode(i).SelectSingleNode("shppTypeDtlCd").Text  ''배송유형상세코드 14 업체자사배송 22 업체택배배송 25 해외택배배송 31 매장방문 41 등기 51 SMS 52 EMAIL 61 미발송
'                    shppTypeDtlNm       = LagrgeNode(i).SelectSingleNode("shppTypeDtlNm").Text  ''배송유형상세명 [업체택배배송]
'                    delicoVenId         = LagrgeNode(i).SelectSingleNode("delicoVenId").Text    ''택배사ID [0000033011]
'                    boxNo               = LagrgeNode(i).SelectSingleNode("boxNo").Text          ''박스번호 [398327952]
'                    shppcst             = LagrgeNode(i).SelectSingleNode("shppcst").Text        '' 배송비? [303] ??
'                    shppcstCodYn        = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text   ''*배송비 착불여부 Y: 착불 N: 선불
'                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*상품명
'                    splVenItemId        = LagrgeNode(i).SelectSingleNode("splVenItemId").Text       ''*업체상품번호 [1024019]
'                    itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*상품번호 [1000024811163]
'                    uitemId             = LagrgeNode(i).SelectSingleNode("uitemId").Text            ''*단품ID [00000]
'                    dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''지시수량 [2]
'                    cnclItemQty         = LagrgeNode(i).SelectSingleNode("cnclItemQty").Text        ''취소수량 [0]
'                    ordQty              = LagrgeNode(i).SelectSingleNode("ordQty").Text             ''주문수량 [2]
'                    sellprc             = LagrgeNode(i).SelectSingleNode("sellprc").Text            ''판매가 [1000]
'                    frgShppYn           = LagrgeNode(i).SelectSingleNode("frgShppYn").Text          ''국내/외 구분 [국내]
'                    ordpeNm             = LagrgeNode(i).SelectSingleNode("ordpeNm").Text            ''*주문자
'                    rcptpeNm            = LagrgeNode(i).SelectSingleNode("rcptpeNm").Text           ''*수취인
'                    rcptpeHpno          = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text         ''*수취인 휴대폰번호
'                    rcptpeTelno         = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text        ''*수취인 집전화번호
'                    shpplocAddr         = LagrgeNode(i).SelectSingleNode("shpplocAddr").Text        ''수취인 상세주소
'                    shpplocZipcd        = LagrgeNode(i).SelectSingleNode("shpplocZipcd").Text       ''*수취인 우편번호          [04733]
'                    shpplocOldZipcd     = LagrgeNode(i).SelectSingleNode("shpplocOldZipcd").Text    ''*수취인 구우편번호(6자리)  [133750]
'                    shpplocRoadAddr     = LagrgeNode(i).SelectSingleNode("shpplocRoadAddr").Text    ''수취인도로명주소
'                    itemChrctDivCd      = LagrgeNode(i).SelectSingleNode("itemChrctDivCd").Text     ''상품특성구분코드 10 일반 20 몰인몰 30 해외구매대행상품 40 미가공귀금속 50 모바일기프트 60 상품권 70 쇼핑충전금 80 모바일상품권 91 이벤트
'
'
'                    shppMainCd          = LagrgeNode(i).SelectSingleNode("shppMainCd").Text         ''배송주체코드 32 업체창고 41 협력업체 42 브랜드직배  [41]
'                    siteNo              = LagrgeNode(i).SelectSingleNode("siteNo").Text             ''사이트번호 6001 이마트몰 6002 트레이더스몰 6003 분스몰 6004 신세계몰 6005 S.COM몰 6009 신세계백화점몰
'                    siteNm              = LagrgeNode(i).SelectSingleNode("siteNm").Text             ''사이트명
'                    shppRsvtDt          = LagrgeNode(i).SelectSingleNode("shppRsvtDt").Text
'                    splprc              = LagrgeNode(i).SelectSingleNode("splprc").Text             ''공급가
'                    shortgYn            = LagrgeNode(i).SelectSingleNode("shortgYn").Text
'                    newWblNoData        = LagrgeNode(i).SelectSingleNode("newWblNoData").Text
'                    newRow              = LagrgeNode(i).SelectSingleNode("newRow").Text
'                    itemDiv             = LagrgeNode(i).SelectSingleNode("itemDiv").Text                ''판매불가신청상태 10:일반 20: 명절 GIFT 일반 30: 명절 GIFT 센터 40: 명절 GIFT 냉장
'                    shpplocBascAddr     = LagrgeNode(i).SelectSingleNode("shpplocBascAddr").Text        ''수취인주소 20170712
'                    shpplocDtlAddr      = LagrgeNode(i).SelectSingleNode("shpplocDtlAddr").Text         ''수취인상세주소	20170712
'                    ordItemDivNm        = LagrgeNode(i).SelectSingleNode("ordItemDivNm").Text           ''주문상품구분	20170809
'
'
'                    ''//필수값 아닌경우 .
'                    if NOT (LagrgeNode(i).SelectSingleNode("ordpeHpno") is Nothing) then
'                        ordpeHpno         = LagrgeNode(i).SelectSingleNode("ordpeHpno").Text           ''주문자휴대폰번호  //선택값
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
'                        ordMemoCntt         = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[고객배송메모]","")            ''고객배송메모  //선택값
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("pCus") is Nothing) then
'                        pCus         = LagrgeNode(i).SelectSingleNode("pCus").Text           ''개인통관고유번호  //선택값
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("frebieNm") is Nothing) then
'                        frebieNm         = LagrgeNode(i).SelectSingleNode("frebieNm").Text    ''사은품  //선택값
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatCd") is Nothing) then
'                        shortgProgStatCd         = LagrgeNode(i).SelectSingleNode("shortgProgStatCd").Text    ''판매불가신청상태  //선택값 11 결품등록 12 결품CS처리중 13 결품확정 21 상품정보오류등록 22 상품정보오류CS처리중 23 상품정보오류확정 41 입고지연등록 43 입고지연완료 51 배송지연등록
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatNm") is Nothing) then
'                        shortgProgStatNm         = LagrgeNode(i).SelectSingleNode("shortgProgStatNm").Text    ''결품진행상태명  //선택값
'                    end if
'
'                    ''옵션명
'                    if NOT (LagrgeNode(i).SelectSingleNode("uitemNm") is Nothing) then
'                        uitemNm         = LagrgeNode(i).SelectSingleNode("uitemNm").Text                 ''[,1:^:asdasd:^:주문문구1,2:^:asdasdddd:^:주문문구2]
'                    end if

			    Next
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing

	rw "======================================"
	rw "총상세건수:"&totalcnt
	''rw "배송번호 업데이트:"&successCnt
end function

function makeSSGChulgoTargetMake()
    dim sqlStr
    sqlStr = "exec [db_etcmall].[dbo].[usp_Ten_OutMall_Ssg_ChulgoTargetMake] "
    dbget.Execute sqlStr

    rw "OK"
end function
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbDatamartclose.asp" -->
