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
''ToDo 4. ��ġ/���� �� �ù�縦 �̿����� �ʴ� ���, ���� ��ۿϷ� ó���� �ؾ� �մϴ�. //2019/10/08 , ������ �ȵǴ����̽���.

'' TLS 1.2�� �������� �ʴ� ������ �ִµ���..
const Option_TLS12 = 2048
const Option_TLS1 = 512
const Option_TLS = 128

'response.write "TT"
'response.end


'' 1. �������
'' 2. ���ó��
'' 3.
'dim IsAutoScript : IsAutoScript =false

dim shppNo  : shppNo = requestCheckvar(request("shppNo"),32)              ''��۹�ȣ
dim shppSeq : shppSeq = requestCheckvar(request("shppSeq"),10)            ''��ۼ���
dim wblno   : wblno = TRIM(requestCheckvar(replace(request("wblno"),"-",""),20))                ''������ȣ
dim delicoVenId : delicoVenId = requestCheckvar(request("delicoVenId"),20)    ''�ù���ȣ
dim itemno : itemno = requestCheckvar(request("itemno"),10)    ''������
dim outmallorderserial : outmallorderserial = requestCheckvar(request("outmallorderserial"),20)    ''�����ֹ���ȣ.
dim orgdetailKey : orgdetailKey = requestCheckvar(request("orgdetailKey"),20)    ''�����ֹ��󼼹�ȣ.

dim prctp : prctp = requestCheckvar(request("prctp"),20)    ''ó�� Action (1:�������, )

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
	response.write "<script>alert('"&iAssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
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
'    response.write "��ô��"
'    dbget.close() : response.end
'end if

'if (outmallorderserial="__20180710269242")  then
'     response.write "��ô��"
'    dbget.close() : response.end
'end if

'' �ֹ��� ���� �Է��ϴ� ���̽�.
Dim GG_ORG_outmallorderserial : GG_ORG_outmallorderserial = outmallorderserial
IF (InStr(outmallorderserial,"_")>0) then
	outmallorderserial = getOutmallRefOrgOrderNO(outmallorderserial,orgdetailKey,"ssg")
end if

if (delicoVenId="��Ÿ") then
    delicoVenId="0000033028" ''��Ÿ�ù��
end if

if (delicoVenId="0000033052") and (LEFT(wblno,1)="1") then ''��ü���ù� / 1�� ����.
    delicoVenId="0000033051" ''��ü�� ���
end if

if (prctp="1") then  ''����� ���
    if (IsPortionSongjangSendRequireNAssign(outmallorderserial,orgdetailKey,shppNo,shppSeq,wblno)) then ''�κ���� �ʿ��Ѱ��
        response.write "�κ���� ó�� �� �����Է�PROC<br>"
        if ((delicoVenId="0000033028") and NOT IsNumeric(wblno)) then
            wblno=RIGHT(outmallorderserial,12)
        elseif ((delicoVenId="0000033028") and LEN(wblno)<11) then
            wblno=wblno&RIGHT(outmallorderserial,6)
        end if

        ''�ֹ���ȣ�� ������ �߰��Ǿ���.
        if NOT(IsNumeric(wblno)) then
            if IsNumeric(LEFT(wblno,LEN(wblno)-6)) and NOT IsNumeric(MID(wblno,LEN(wblno)-6+1,6)) then
                wblno = LEFT(wblno,LEN(wblno)-6) & cdbl("&H"&MID(wblno,LEN(wblno)-6+1,6))
            end if
        end if
        Call saveSSGWblNo(shppNo,shppSeq,wblno,delicoVenId,outmallorderserial,orgdetailKey, IsAutoScript)
    else
        response.write "�����Է�PROC<br>"
        if ((delicoVenId="0000033028") and NOT IsNumeric(wblno)) then
            wblno=RIGHT(outmallorderserial,12)
        elseif ((delicoVenId="0000033028") and LEN(wblno)<11) then
            wblno=wblno&RIGHT(outmallorderserial,6)
        end if

        ''rw LEFT(wblno,LEN(wblno)-5)&":"&MID(wblno,LEN(wblno)-5+1,5)
        ''�ֹ���ȣ�� ������ �߰��Ǿ���.
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
elseif (prctp="2") then ''���ó��
    Call Savessgwhoutcompleteprocess(shppNo,shppSeq, itemno, outmallorderserial,orgdetailKey,wblno, IsAutoScript)
elseif (prctp="3") then ''�̹�ۿϷ� ��� ��ȸ

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
            Call saveDeliveryEnd(shppNo,shppSeq,outmallorderserial,orgdetailKey,IsAutoScript) ''��ۿϷ�ó��
        else
            rw "None"
        end if
    else
        rw "ERR"
    end if
elseif (prctp="33") then ''�̹�ۿϷ� CS ��� ��ȸ

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

        Call saveDeliveryEndCS(shppNo,shppSeq,outmallorderserial,orgdetailKey,IsAutoScript) ''��ۿϷ�ó��
    else
        rw "ERR"
    end if
elseif (prctp="5") then ''�κ����ó��
    rw "ERR"
    'Call saveSSGPortionWarehouseOutProcess(outmallorderserial,orgdetailKey,shppNo,shppSeq,itemno)
    'CAll getSsgChulgoListByOutmallOrderserialUpshppNo(outmallorderserial)
elseif (prctp="6") then ''�ֹ�Ȯ��
    Call saveSsgOrderChulgoConfirm(shppNo,shppSeq)
elseif (prctp="9") then ''��ǰó�� (���Ұ�)

elseif (prctp="999") then ''���ó������ֹ� �ۼ�.
    Call makeSSGChulgoTargetMake()
else
    response.write "������:"&prctp
    dbget.close():response.end
end if


''�κ���� ó�� : �κ���� �ϸ�. ���� ��۹�ȣ�� ������ȣ�� ������ �ٲ��, ��۹�ȣ�� �ٲ�? /�ֹ���ȣ�� ���ε���, ���ֹ���ȣ�� �ٲ��� ����/  (�κ���� �Է��Ѱ���?) [�������ó���� �ƴ�?] => �ٽ� ��۹�ȣ�� ��ȸ
''��ǰ�ܿ��� ���ó�� ���� �����ؾ��ҵ���.
''�̰ɾ��� ���Ǵ°� �ƴ϶� ������ȣ�� ��������, ��۹�ȣ�� �ٲ�°���.
''Call saveSSGPortionWarehouseOutProcess(shppNo,shppSeq,itemno)

''�ֹ���ȣ�� ��۸����ȸ �� ��۹�ȣ������Ʈ  : ��۹�ȣ�� �ٽð������� ����.
''CAll getSsgChulgoListByOutmallOrderserialUpshppNo(outmallorderserial)

''���ó��
''Call saveSSGWhOutCompleteProcess(shppNo,shppSeq, itemno, outmallorderserial,orgdetailKey)

''��ۿϷ�ó�� ''�ù�� �ڵ����� �ȴ�.?

''�κ���� �ʿ����� üũ
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
        response.write "�κ����ó��PROC1"
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

'// �ֹ�Ȯ��
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

''�κ���� ó��
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
    	    response.write "����ȸ�� �ʰ�."

    		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
    						"	<option value=''>����</option>" &_
    						"	<option value='951'>������ ����</option>" &_
    						"	<option value='952'>����ֹ�</option>" &_
    						"</select>&nbsp;&nbsp;"
    		response.write "<input type='button' value='�Ϸ�ó��' onClick=""fnSetSendState('"&GG_ORG_outmallorderserial&"','"&ishppNo&"','"&ishppSeq&"','"&orgdetailKey&"',document.getElementById('updateSendState').value)"">"
    		response.write "<script language='javascript'>"&VbCRLF
    		response.write "function fnSetSendState(outmallorderserial,ishppNo,ishppSeq,iorgdetailKey,selectValue){"&VbCRLF
    		response.write "    if(selectValue == ''){"&VbCRLF
    		response.write "    	alert('�������ּ���');"&VbCRLF
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

''��ۿϷ�� ���
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
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listDeliveryEnd.ssg"   ''��ۿϷ���

	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestDeliveryEnd>"
    requestBody = requestBoDy&"<perdType>02</perdType>"  ''02 ���Ϸ���
    requestBody = requestBoDy&"<perdStrDts>"&perdStrDts&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&perdEndDts&"</perdEndDts>"
    requestBody = requestBoDy&"<commType>01</commType>"  ''01 ���ֹ���ȣ / 02 ��۹�ȣ
    requestBody = requestBoDy&"<commValue>"&outmallorderserial&"</commValue>"  ''������ȸ��, ������ȸ�� ��� �ʼ� ��
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
			        ''�����ʱ�ȭ.
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

                    ordNo           = LagrgeNode(i).SelectSingleNode("ordNo").Text '*�ֹ���ȣ [20171123128379]
			        ordItemSeq      = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text ''*�ֹ�����
			        if NOT (LagrgeNode(i).SelectSingleNode("orOrdNo") is Nothing) then
                        orOrdNo         = LagrgeNode(i).SelectSingleNode("orOrdNo").Text '*���ֹ���ȣ [20171123128379]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("orordItemSeq") is Nothing) then
                        orordItemSeq    = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text ''���ֹ����� [2]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("shppTypeCd") is Nothing) then
                        shppTypeCd    = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text ''��������ڵ�
                    end if
                    reOrderYn       = LagrgeNode(i).SelectSingleNode("reOrderYn").Text '' ���ֹ�����
                    shppNo          = LagrgeNode(i).SelectSingleNode("shppNo").Text '' ��۹�ȣ
                    shppSeq         = LagrgeNode(i).SelectSingleNode("shppSeq").Text '' ��ۼ���
                    shppDt          = LagrgeNode(i).SelectSingleNode("shppDt").Text '' ��ۿϷ��Ͻ�

                    if NOT (LagrgeNode(i).SelectSingleNode("whoutDt") is Nothing) then
                        whoutDt         = LagrgeNode(i).SelectSingleNode("whoutDt").Text '' ���Ϸ��Ͻ�
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("wblNo") is Nothing) then
                        wblNo    = LagrgeNode(i).SelectSingleNode("wblNo").Text ''�����ȣ
                    end if
                    ''settlRfltDts    = LagrgeNode(i).SelectSingleNode("shppNo").Text  '' ����ݿ��Ͻ�
                    lastShppProgStatDtlCd  = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''������ۻ���������ڵ� 11 �������, 21 ��ŷ����, 22 ��ŷ�Ϸ�, 31 ��ŷ�Ϸ�, 41 �����, 42 �������, 43 ���Ϸ�, 51 ��ۿϷ�, 52 ��۰���



                    if (lastShppProgStatDtlCd="43" or lastShppProgStatDtlCd="51") then
                        if ((ordNo=outmallorderserial) and (ordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) _
                            or ((orOrdNo=outmallorderserial) and (orordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) _
                            or ((orOrdNo="") and (orordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) then

                            if (shppTypeCd="10") then
                                rw "shppTypeCd:�ڻ���"
                            elseif (shppTypeCd="20") then
                                rw "shppTypeCd:�ù���:"&wblNo
                            else
                                rw "shppTypeCd:"&shppTypeCd
                            end if

                            if (lastShppProgStatDtlCd="22") then
                                rw "��ŷ�Ϸ� ����"
                                tenSendstate = "0"
                            elseif (lastShppProgStatDtlCd="43") then
                                rw "���Ϸ� ����"
                                tenSendstate = "3"
                            elseif (lastShppProgStatDtlCd="51") then
                                rw "��ۿϷ� ����"
                                tenSendstate = "4"
                            end if

                            rw ordNo&"|"&ordItemSeq&"|"&orOrdNo&"|"&orordItemSeq&"|"&reOrderYn&"|"&shppNo&"|"&shppSeq&"|"&shppDt&"|"&whoutDt&"|"&lastShppProgStatDtlCd&"|"&wblNo
                            rw outmallorderserial&"|"&orgdetailKey&"|"&ishppNo&"|"&ishppSeq&"|"&iwblno

                            ''if (wblNo=iwblno) or (wblNo="") then ''wblNo="" �ڻ���
                                strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
                            	strSql = strSql & "	Set sendstate='"&tenSendstate&"'" ''��ۿϷ�� 4������ ����.
                                strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
                                strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
                            	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
                            	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
                                strSql = strSql & "	and matchstate in ('O','A')"  ''A �߰� 2019/10/08
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

''�� ��ۿϷ�� ���(���Ϸ� ����?)
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
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listNonDelivery.ssg"   ''�� ��ۿϷ���

	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestNonDelivery>"
    requestBody = requestBoDy&"<perdType>01</perdType>"  ''01 ���Ϸ��� / 02 �����Ϸ���
    requestBody = requestBoDy&"<perdStrDts>"&perdStrDts&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&perdEndDts&"</perdEndDts>"
    requestBody = requestBoDy&"<commType>01</commType>"  ''01 ���ֹ���ȣ / 02 ��۹�ȣ
    requestBody = requestBoDy&"<commValue>"&outmallorderserial&"</commValue>"  ''������ȸ��, ������ȸ�� ��� �ʼ� ��
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
			        ''�����ʱ�ȭ.
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


                    ordNo           = LagrgeNode(i).SelectSingleNode("ordNo").Text '*�ֹ���ȣ [20171123128379]
			        ordItemSeq      = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text ''*�ֹ�����
			        if NOT (LagrgeNode(i).SelectSingleNode("orOrdNo") is Nothing) then
                        orOrdNo         = LagrgeNode(i).SelectSingleNode("orOrdNo").Text '*���ֹ���ȣ [20171123128379]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("orordItemSeq") is Nothing) then
                        orordItemSeq    = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text ''���ֹ����� [2]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("shppTypeCd") is Nothing) then
                        shppTypeCd    = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text ''��������ڵ�
                    end if

                    reOrderYn       = LagrgeNode(i).SelectSingleNode("reOrderYn").Text '' ���ֹ�����
                    shppNo          = LagrgeNode(i).SelectSingleNode("shppNo").Text '' ��۹�ȣ
                    shppSeq         = LagrgeNode(i).SelectSingleNode("shppSeq").Text '' ��ۼ���

                    ''shppDt          = LagrgeNode(i).SelectSingleNode("shppDt").Text '' ��ۿϷ��Ͻ�
                    ' whoutDt         = LagrgeNode(i).SelectSingleNode("whoutDt").Text '' ���Ϸ��Ͻ�
                    ''settlRfltDts    = LagrgeNode(i).SelectSingleNode("shppNo").Text  '' ����ݿ��Ͻ�
                    lastShppProgStatDtlCd  = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''������ۻ���������ڵ� 11 �������, 21 ��ŷ����, 22 ��ŷ�Ϸ�, 31 ��ŷ�Ϸ�, 41 �����, 42 �������, 43 ���Ϸ�, 51 ��ۿϷ�, 52 ��۰���
                    if NOT (LagrgeNode(i).SelectSingleNode("wblNo") is Nothing) then
                        wblNo    = LagrgeNode(i).SelectSingleNode("wblNo").Text ''�����ȣ
                    end if

                    ''rw ordNo&"|"&ordItemSeq&"|"&orOrdNo&"|"&orordItemSeq&"|"&reOrderYn&"|"&shppNo&"|"&shppSeq&"|"&shppDt&"|"&whoutDt&"|"&lastShppProgStatDtlCd&"|"&wblNo
                    ''rw outmallorderserial&"|"&orgdetailKey&"|"&ishppNo&"|"&ishppSeq&"|"&iwblno


                    '' 43:���Ϸ�,51:��ۿϷ�,(22:��ŷ�Ϸ�)
                    if (lastShppProgStatDtlCd="43" or lastShppProgStatDtlCd="51") then
                        ''if (ordNo=outmallorderserial) and (ordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq) then
                        if ((ordNo=outmallorderserial) and (ordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) _
                            or ((orOrdNo=outmallorderserial) and (orordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq))  _
                            or ((orOrdNo="") and (orordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) then

                            if (shppTypeCd="10") then
                                rw "shppTypeCd:�ڻ���"
                            elseif (shppTypeCd="20") then
                                rw "shppTypeCd:�ù���:"&wblNo
                            else
                                rw "shppTypeCd:"&shppTypeCd
                            end if

                            if (lastShppProgStatDtlCd="22") then
                                rw "��ŷ�Ϸ� ����"
                                tenSendstate = "0"
                            elseif (lastShppProgStatDtlCd="43") then
                                rw "���Ϸ� ����"
                                tenSendstate = "3"
                            elseif (lastShppProgStatDtlCd="51") then
                                rw "��ۿϷ� ����"
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

''�� ��ۿϷ�� ���(���Ϸ� ����) ��۹�ȣ�� ��ȸ 2019/10/16
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
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listNonDelivery.ssg"   ''�� ��ۿϷ���

	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestNonDelivery>"
    requestBody = requestBoDy&"<perdType>01</perdType>"  ''01 ���Ϸ��� / 02 �����Ϸ���
    requestBody = requestBoDy&"<perdStrDts>"&perdStrDts&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&perdEndDts&"</perdEndDts>"
    requestBody = requestBoDy&"<commType>02</commType>"  ''01 ���ֹ���ȣ / 02 ��۹�ȣ
    requestBody = requestBoDy&"<commValue>"&ishppNo&"</commValue>"  ''������ȸ��, ������ȸ�� ��� �ʼ� ��
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
			        ''�����ʱ�ȭ.
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


                    ordNo           = LagrgeNode(i).SelectSingleNode("ordNo").Text '*�ֹ���ȣ [20171123128379]
			        ordItemSeq      = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text ''*�ֹ�����
			        if NOT (LagrgeNode(i).SelectSingleNode("orOrdNo") is Nothing) then
                        orOrdNo         = LagrgeNode(i).SelectSingleNode("orOrdNo").Text '*���ֹ���ȣ [20171123128379]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("orordItemSeq") is Nothing) then
                        orordItemSeq    = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text ''���ֹ����� [2]
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("shppTypeCd") is Nothing) then
                        shppTypeCd    = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text ''��������ڵ�
                    end if

                    reOrderYn       = LagrgeNode(i).SelectSingleNode("reOrderYn").Text '' ���ֹ�����
                    shppNo          = LagrgeNode(i).SelectSingleNode("shppNo").Text '' ��۹�ȣ
                    shppSeq         = LagrgeNode(i).SelectSingleNode("shppSeq").Text '' ��ۼ���

                    ''shppDt          = LagrgeNode(i).SelectSingleNode("shppDt").Text '' ��ۿϷ��Ͻ�
                    ' whoutDt         = LagrgeNode(i).SelectSingleNode("whoutDt").Text '' ���Ϸ��Ͻ�
                    ''settlRfltDts    = LagrgeNode(i).SelectSingleNode("shppNo").Text  '' ����ݿ��Ͻ�
                    lastShppProgStatDtlCd  = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''������ۻ���������ڵ� 11 �������, 21 ��ŷ����, 22 ��ŷ�Ϸ�, 31 ��ŷ�Ϸ�, 41 �����, 42 �������, 43 ���Ϸ�, 51 ��ۿϷ�, 52 ��۰���

                    if (orOrdNo<>"") and (ordNo<>orOrdNo) then ordNo=orOrdNo
                    if (orordItemSeq<>"") and (ordItemSeq<>orordItemSeq) then ordItemSeq=orordItemSeq

                    if NOT (LagrgeNode(i).SelectSingleNode("delicoVenNm") is Nothing) then
                        delicoVenNm    = LagrgeNode(i).SelectSingleNode("delicoVenNm").Text ''�ù��
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("wblNo") is Nothing) then
                        wblNo    = LagrgeNode(i).SelectSingleNode("wblNo").Text ''�����ȣ
                    end if
                    if NOT (LagrgeNode(i).SelectSingleNode("shppTypeDtlNm") is Nothing) then
                        shppTypeDtlNm    = LagrgeNode(i).SelectSingleNode("shppTypeDtlNm").Text ''��������󼼸�
                    end if


                    ' rw ordNo&"|"&ordItemSeq&"|"&orOrdNo&"|"&orordItemSeq&"|"&reOrderYn&"|"&shppNo&"|"&shppSeq&"|"&shppDt&"|"&whoutDt&"|"&lastShppProgStatDtlCd&"|"&wblNo
                    ' rw outmallorderserial&"|"&orgdetailKey&"|"&ishppNo&"|"&ishppSeq&"|"&shppTypeDtlNm&"|"&iwblno


                    '' 43:���Ϸ�,51:��ۿϷ�,(22:��ŷ�Ϸ�)
                    if (lastShppProgStatDtlCd="43" or lastShppProgStatDtlCd="51") then
                        if ((ordNo=outmallorderserial) and (ordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) _
                            or ((orOrdNo=outmallorderserial) and (orordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq))  _
                            or ((orOrdNo="") and (orordItemSeq=orgdetailKey) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) then

                            if (shppTypeCd="10") then
                                rw "shppTypeCd:�ڻ���"
                            elseif (shppTypeCd="20") then
                                rw "shppTypeCd:�ù���:"&wblNo
                            else
                                rw "shppTypeCd:"&shppTypeCd
                            end if

                            if (lastShppProgStatDtlCd="22") then
                                rw "��ŷ�Ϸ� ����"
                                tenSendstate = "0"
                            elseif (lastShppProgStatDtlCd="43") then
                                rw "���Ϸ� ����"
                                tenSendstate = "3"
                            elseif (lastShppProgStatDtlCd="51") then
                                rw "��ۿϷ� ����"
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

''��ǰ��� ó��
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
            if (xmlDOM.getElementsByTagName("resultDesc").length > 0) then  ''2019-09-04 11:38 ���� ������ ����
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


    	'���� ����Ƚ���� 3ȸ�� ������ ����ó�� ����
    	'updateSendState = 951		������ ����
    	'updateSendState = 952		����ֹ�

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
            if (ssgresultDesc<>"�ش� �����Ͱ� �����ϴ�.") then '' ����ϰ������� ����.
                reChkCnt = CheckOrderSendedAssign(ishppNo,ishppSeq, outmallorderserial,orgdetailKey, iwblno,IsAutoScript)
                if (reChkCnt<1) then
                    reChkCnt = CheckOrderFinishedAssign(ishppNo,ishppSeq, outmallorderserial,orgdetailKey, iwblno, IsAutoScript)
                end if
                rw reChkCnt
                if reChkCnt>0 then errCount=0  ''�̹� ��ۿϷ� ���� �Ǿ���.
            end if
        end if

    	If errCount > 0 Then
    	    response.write "����ȸ�� �ʰ�."

    		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
    						"	<option value=''>����</option>" &_
    						"	<option value='951'>������ ����</option>" &_
    						"	<option value='952'>����ֹ�</option>" &_
    						"</select>&nbsp;&nbsp;"
    		response.write "<input type='button' value='�Ϸ�ó��' onClick=""fnSetSendState('"&GG_ORG_outmallorderserial&"','"&ishppNo&"','"&ishppSeq&"','"&orgdetailKey&"',document.getElementById('updateSendState').value)"">"
    		response.write "<script language='javascript'>"&VbCRLF
    		response.write "function fnSetSendState(outmallorderserial,ishppNo,ishppSeq,iorgdetailKey,selectValue){"&VbCRLF
    		response.write "    if(selectValue == ''){"&VbCRLF
    		response.write "    	alert('�������ּ���');"&VbCRLF
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

'' ����� ��� API // �������API �� ���ó�� API�� �����̴�, ���ó���� �κ����ó��/��ü��� ����;;
'' ������� [��۹�ȣ] ���� �ϳ��� ��� �����ϴ�? - �ߺ�����Ұ�? == ��ü����� ������� ������ �ؾ��ҵ�?
'' �ϳ��־ �۹�ȣ�� ������ ���� ���
public function saveSSGWblNo(ishppNo,ishppSeq, iwblno, idelicoVenId, outmallorderserial, orgdetailKey, IsAutoScript)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim shppTypeCd: shppTypeCd="20"         ''�ù���
    Dim shppTypeDtlCd: shppTypeDtlCd="22"   ''��ü�ù���
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

	if (ssgresultCode="00" AND ssgresultDesc="����") then
        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "&VBCRLF
    	strSql = strSql & "	Set sendState=1"&VBCRLF
    	strSql = strSql & "	,sendReqCnt=1"&VBCRLF ''sendReqCnt+1 =>1 , 1�� �ʱ�ȭ
    	strSql = strSql & "	,sendSongjangNo='"&iwblno&"'"&VBCRLF        ''2017/01/03 �߰�
        strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
        strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
    	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
    	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
        'strSql = strSql & "	and matchstate in ('O')"
        strSql = strSql & "	and matchstate in ('O', 'A')" '2019-05-16 ������ A�� �߰�
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
        ''����� ��ȣ�� ��ȿ���� �ʽ��ϴ�.
        ''���ϵ� ������ȣ�Դϴ�.
        Dim reChkCnt : reChkCnt=0
        Dim IsreqForceSend : IsreqForceSend = false
        if (errCount>0) and (NOT IsAutoScript) then
            reChkCnt = getSsgChulgoListByshppNo(ishppNo,ishppSeq, outmallorderserial,orgdetailKey, iwblno, IsreqForceSend, IsAutoScript)

            if (IsreqForceSend) then ''��۱����� �ӹ��Ͽ� �ڰ�������� �����Է�ó�� �Ѵ�..
                call saveSSGWblNoForce(ishppNo,ishppSeq, iwblno, idelicoVenId, outmallorderserial, orgdetailKey, IsAutoScript)
                Exit function
            end if

            if (ssgresultDesc<>"�ش� �����Ͱ� �����ϴ�.") then '' ����ϰ������� ����.
                if (reChkCnt<1) then
                    reChkCnt = CheckOrderSendedAssign(ishppNo,ishppSeq, outmallorderserial,orgdetailKey, iwblno,IsAutoScript)  '' iitemno=>1 not using
                    if (reChkCnt<1) then
                        reChkCnt = CheckOrderFinishedAssign(ishppNo,ishppSeq, outmallorderserial,orgdetailKey, iwblno, IsAutoScript)
                    end if
                    rw reChkCnt
                    if reChkCnt>0 then errCount=0  ''�̹� ��ۿϷ� ���� �Ǿ���.
                end if
            end if
        end if

    	If errCount > 0 Then
    	    response.write "����ȸ�� �ʰ�.."

    		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
    						"	<option value=''>����</option>" &_
    						"	<option value='951'>������ ����</option>" &_
    						"	<option value='952'>����ֹ�</option>" &_
    						"</select>&nbsp;&nbsp;"
    		response.write "<input type='button' value='�Ϸ�ó��' onClick=""fnSetSendState('"&GG_ORG_outmallorderserial&"','"&ishppNo&"','"&ishppSeq&"','"&orgdetailKey&"',document.getElementById('updateSendState').value)"">"
    		response.write "<script language='javascript'>"&VbCRLF
    		response.write "function fnSetSendState(outmallorderserial,ishppNo,ishppSeq,iorgdetailKey,selectValue){"&VbCRLF
    		response.write "    if(selectValue == ''){"&VbCRLF
    		response.write "    	alert('�������ּ���');"&VbCRLF
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
    Dim shppTypeCd: shppTypeCd="10"         ''�ڻ���
    Dim shppTypeDtlCd: shppTypeDtlCd="14"   ''��ü�ڻ���
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
    'requestBody = requestBoDy&"<wblNo>"&iwblno&"</wblNo>"                      ''�ڻ����ΰ�� ���� �ʴ´�.
    'requestBody = requestBoDy&"<delicoVenId>"&idelicoVenId&"</delicoVenId>"    ''�ڻ����ΰ�� ���� �ʴ´�.
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

	rw 	"<font color='blue'>��ۿϷ�ó��:"&ssgresultCode&":"&ssgresultMessage&":"&ssgresultDesc&"<font>"

	if (ssgresultCode="00" AND ssgresultDesc="����") then
        sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_OUTAMLL_Jungsan_Comment_add] '" & tenorderserial & "', " & tenitemid & ", '" & tenitemoption & "','��ۿϷ�����','"&session("ssBctId")&"'"
		dbDatamart_dbget.Execute sqlStr

        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "&VBCRLF
    	strSql = strSql & "	Set sendState=3"&VBCRLF
        strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
        strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
    	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
    	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
        strSql = strSql & "	and matchstate in ('O', 'A')" '2019-05-16 ������ A�� �߰�
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
    Dim shppTypeCd: shppTypeCd="10"         ''�ڻ���
    Dim shppTypeDtlCd: shppTypeDtlCd="14"   ''��ü�ڻ���
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
    'requestBody = requestBoDy&"<wblNo>"&iwblno&"</wblNo>"                      ''�ڻ����ΰ�� ���� �ʴ´�.
    'requestBody = requestBoDy&"<delicoVenId>"&idelicoVenId&"</delicoVenId>"    ''�ڻ����ΰ�� ���� �ʴ´�.
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

	rw 	"<font color='blue'>��ۿϷ�ó��:"&ssgresultCode&":"&ssgresultMessage&":"&ssgresultDesc&"<font>"

	if (ssgresultCode="00" AND ssgresultDesc="����") then
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
    Dim shppTypeCd: shppTypeCd="10"         ''�ڻ���
    Dim shppTypeDtlCd: shppTypeDtlCd="14"   ''��ü�ڻ���
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
    'requestBody = requestBoDy&"<wblNo>"&iwblno&"</wblNo>"                      ''�ڻ����ΰ�� ���� �ʴ´�.
    'requestBody = requestBoDy&"<delicoVenId>"&idelicoVenId&"</delicoVenId>"    ''�ڻ����ΰ�� ���� �ʴ´�.
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

	rw 	"<font color='blue'>�ڻ�����������:"&ssgresultCode&":"&ssgresultMessage&":"&ssgresultDesc&"<font>"

	if (ssgresultCode="00" AND ssgresultDesc="����") then
        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "&VBCRLF
    	strSql = strSql & "	Set sendState=1"&VBCRLF
    	strSql = strSql & "	,sendReqCnt=1"&VBCRLF ''sendReqCnt+1 =>1 , 1�� �ʱ�ȭ
    	strSql = strSql & "	,sendSongjangNo='�ڻ���"&iwblno&"'"&VBCRLF        ''2017/01/03 �߰�
        strSql = strSql & "	where OutMallOrderSerial='"&GG_ORG_outmallorderserial&"'"&VBCRLF
        strSql = strSql & "	and beasongNum11st='"&ishppNo&"'"&VBCRLF
    	strSql = strSql & "	and reserve01='"&ishppSeq&"'"&VBCRLF
    	strSql = strSql & "	and OrgDetailKey='"&orgdetailKey&"'"&VBCRLF
        'strSql = strSql & "	and matchstate in ('O')"
        strSql = strSql & "	and matchstate in ('O', 'A')" '2019-05-16 ������ A�� �߰�
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

''�Ͱ� ����� ��ȸ
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
			        ''�����ʱ�ȭ.
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

			        shppNo              = LagrgeNode(i).SelectSingleNode("shppNo").Text                 ''*��۹�ȣ
                    shppSeq             = LagrgeNode(i).SelectSingleNode("shppSeq").Text                ''*��ۼ���
                    ordNo               = LagrgeNode(i).SelectSingleNode("ordNo").Text                  ''*�ֹ���ȣ [20171123128379]
                    ordItemSeq          = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text             ''*�ֹ�����
                    orordNo             = LagrgeNode(i).SelectSingleNode("orordNo").Text            ''���ֹ���ȣ [20171123128379]

                    On Error Resume Next
                        orordItemSeq    = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''���ֹ����� [2]
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

'                    shppTabProgStatCd   = LagrgeNode(i).SelectSingleNode("shppTabProgStatCd").Text      ''������ۻ���������ڵ�(��۴���) 11 ������� 21 ��ŷ���� 22 ��ŷ�Ϸ� 31 ��ŷ�Ϸ� 41 ����� 42 ������� 43 ���Ϸ� 51 ��ۿϷ� 52 ��۰���
'                    evntSeq             = LagrgeNode(i).SelectSingleNode("evntSeq").Text                ''�̺�Ʈ����
'                    shppDivDtlCd        = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text           ''*��۱��л��ڵ� 11 �Ϲ���� 12 �κ���� 14 ���� 15 ��ȯ��� 16 AS���
'                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''��۱��л󼼸�
'                    reOrderYn           = LagrgeNode(i).SelectSingleNode("reOrderYn").Text              ''*�����ÿ��α���
'                    delayNts            = LagrgeNode(i).SelectSingleNode("delayNts").Text               ''����Ƚ��
'                    ordCmplDts          = LagrgeNode(i).SelectSingleNode("ordCmplDts").Text             ''*�ֹ��Ϸ��Ͻ� [2017-11-23 10:39:42.0]
'                    lastShppProgStatDtlNm   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlNm").Text  ''������ۻ�������¸�(��ۻ�ǰ����) [��ŷ�Ϸ�]
'                    lastShppProgStatDtlCd   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''������ۻ���������ڵ�(��ۻ�ǰ����) 11 ������� 21 ��ŷ���� 22 ��ŷ�Ϸ� 31 ��ŷ�Ϸ� 41 ����� 42 ������� 43 ���Ϸ� 51 ��ۿϷ� 52 ��۰���
'                    salestrNo           = LagrgeNode(i).SelectSingleNode("salestrNo").Text              '' [6004]
'                    shppVenId           = LagrgeNode(i).SelectSingleNode("shppVenId").Text      ''���޾�ü���̵� [0000003198]
'                    shppVenNm           = LagrgeNode(i).SelectSingleNode("shppVenNm").Text      ''���޾�ü��
'                    shppTypeNm          = LagrgeNode(i).SelectSingleNode("shppTypeNm").Text     ''���������    [�ù���]
'                    shppTypeCd          = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text     ''��������ڵ� 10 �ڻ��� 20 �ù��� 30 ����湮 40 ��� 50 �̹�� 60 �̹߼�
'                    shppTypeDtlCd       = LagrgeNode(i).SelectSingleNode("shppTypeDtlCd").Text  ''����������ڵ� 14 ��ü�ڻ��� 22 ��ü�ù��� 25 �ؿ��ù��� 31 ����湮 41 ��� 51 SMS 52 EMAIL 61 �̹߼�
'                    shppTypeDtlNm       = LagrgeNode(i).SelectSingleNode("shppTypeDtlNm").Text  ''��������󼼸� [��ü�ù���]
'                    delicoVenId         = LagrgeNode(i).SelectSingleNode("delicoVenId").Text    ''�ù��ID [0000033011]
'                    boxNo               = LagrgeNode(i).SelectSingleNode("boxNo").Text          ''�ڽ���ȣ [398327952]
'                    shppcst             = LagrgeNode(i).SelectSingleNode("shppcst").Text        '' ��ۺ�? [303] ??
'                    shppcstCodYn        = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text   ''*��ۺ� ���ҿ��� Y: ���� N: ����
'                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*��ǰ��
'                    splVenItemId        = LagrgeNode(i).SelectSingleNode("splVenItemId").Text       ''*��ü��ǰ��ȣ [1024019]
'                    itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*��ǰ��ȣ [1000024811163]
'                    uitemId             = LagrgeNode(i).SelectSingleNode("uitemId").Text            ''*��ǰID [00000]
'                    dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''���ü��� [2]
'                    cnclItemQty         = LagrgeNode(i).SelectSingleNode("cnclItemQty").Text        ''��Ҽ��� [0]
'                    ordQty              = LagrgeNode(i).SelectSingleNode("ordQty").Text             ''�ֹ����� [2]
'                    sellprc             = LagrgeNode(i).SelectSingleNode("sellprc").Text            ''�ǸŰ� [1000]
'                    frgShppYn           = LagrgeNode(i).SelectSingleNode("frgShppYn").Text          ''����/�� ���� [����]
'                    ordpeNm             = LagrgeNode(i).SelectSingleNode("ordpeNm").Text            ''*�ֹ���
'                    rcptpeNm            = LagrgeNode(i).SelectSingleNode("rcptpeNm").Text           ''*������
'                    rcptpeHpno          = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text         ''*������ �޴�����ȣ
'                    rcptpeTelno         = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text        ''*������ ����ȭ��ȣ
'                    shpplocAddr         = LagrgeNode(i).SelectSingleNode("shpplocAddr").Text        ''������ ���ּ�
'                    shpplocZipcd        = LagrgeNode(i).SelectSingleNode("shpplocZipcd").Text       ''*������ �����ȣ          [04733]
'                    shpplocOldZipcd     = LagrgeNode(i).SelectSingleNode("shpplocOldZipcd").Text    ''*������ �������ȣ(6�ڸ�)  [133750]
'                    shpplocRoadAddr     = LagrgeNode(i).SelectSingleNode("shpplocRoadAddr").Text    ''�����ε��θ��ּ�
'                    itemChrctDivCd      = LagrgeNode(i).SelectSingleNode("itemChrctDivCd").Text     ''��ǰƯ�������ڵ� 10 �Ϲ� 20 ���θ� 30 �ؿܱ��Ŵ����ǰ 40 �̰����ͱݼ� 50 ����ϱ���Ʈ 60 ��ǰ�� 70 ���������� 80 ����ϻ�ǰ�� 91 �̺�Ʈ
'                    shppStatCd          = LagrgeNode(i).SelectSingleNode("shppStatCd").Text         ''*��ۻ����ڵ� 10 ���� 30 ���
'                    shppStatNm          = LagrgeNode(i).SelectSingleNode("shppStatNm").Text         ''��ۻ��¸�
'                    shppMainCd          = LagrgeNode(i).SelectSingleNode("shppMainCd").Text         ''�����ü�ڵ� 32 ��üâ�� 41 ���¾�ü 42 �귣������  [41]
'                    siteNo              = LagrgeNode(i).SelectSingleNode("siteNo").Text             ''����Ʈ��ȣ 6001 �̸�Ʈ�� 6002 Ʈ���̴����� 6003 �н��� 6004 �ż���� 6005 S.COM�� 6009 �ż����ȭ����
'                    siteNm              = LagrgeNode(i).SelectSingleNode("siteNm").Text             ''����Ʈ��
'                    shppRsvtDt          = LagrgeNode(i).SelectSingleNode("shppRsvtDt").Text
'                    splprc              = LagrgeNode(i).SelectSingleNode("splprc").Text             ''���ް�
'                    shortgYn            = LagrgeNode(i).SelectSingleNode("shortgYn").Text
'                    newWblNoData        = LagrgeNode(i).SelectSingleNode("newWblNoData").Text
'                    newRow              = LagrgeNode(i).SelectSingleNode("newRow").Text
'                    itemDiv             = LagrgeNode(i).SelectSingleNode("itemDiv").Text                ''�ǸźҰ���û���� 10:�Ϲ� 20: ���� GIFT �Ϲ� 30: ���� GIFT ���� 40: ���� GIFT ����
'                    shpplocBascAddr     = LagrgeNode(i).SelectSingleNode("shpplocBascAddr").Text        ''�������ּ� 20170712
'                    shpplocDtlAddr      = LagrgeNode(i).SelectSingleNode("shpplocDtlAddr").Text         ''�����λ��ּ�	20170712
'                    ordItemDivNm        = LagrgeNode(i).SelectSingleNode("ordItemDivNm").Text           ''�ֹ���ǰ����	20170809
'
'
'                    ''//�ʼ��� �ƴѰ�� .
'                    if NOT (LagrgeNode(i).SelectSingleNode("ordpeHpno") is Nothing) then
'                        ordpeHpno         = LagrgeNode(i).SelectSingleNode("ordpeHpno").Text           ''�ֹ����޴�����ȣ  //���ð�
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
'                        ordMemoCntt         = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[����۸޸�]","")            ''����۸޸�  //���ð�
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("pCus") is Nothing) then
'                        pCus         = LagrgeNode(i).SelectSingleNode("pCus").Text           ''�������������ȣ  //���ð�
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("frebieNm") is Nothing) then
'                        frebieNm         = LagrgeNode(i).SelectSingleNode("frebieNm").Text    ''����ǰ  //���ð�
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatCd") is Nothing) then
'                        shortgProgStatCd         = LagrgeNode(i).SelectSingleNode("shortgProgStatCd").Text    ''�ǸźҰ���û����  //���ð� 11 ��ǰ��� 12 ��ǰCSó���� 13 ��ǰȮ�� 21 ��ǰ����������� 22 ��ǰ��������CSó���� 23 ��ǰ��������Ȯ�� 41 �԰�������� 43 �԰������Ϸ� 51 ����������
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatNm") is Nothing) then
'                        shortgProgStatNm         = LagrgeNode(i).SelectSingleNode("shortgProgStatNm").Text    ''��ǰ������¸�  //���ð�
'                    end if
'
'                    ''�ɼǸ�
'                    if NOT (LagrgeNode(i).SelectSingleNode("uitemNm") is Nothing) then
'                        uitemNm         = LagrgeNode(i).SelectSingleNode("uitemNm").Text                 ''[,1:^:asdasd:^:�ֹ�����1,2:^:asdasdddd:^:�ֹ�����2]
'                    end if

			    Next
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing

	rw "======================================"
	rw "�ѻ󼼰Ǽ�:"&totalcnt
	rw "��۹�ȣ ������Ʈ:"&successCnt
end function

''�Ͱ� ����� ��ȸ by ��۹�ȣ
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
			        ''�����ʱ�ȭ.
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

			        shppNo              = LagrgeNode(i).SelectSingleNode("shppNo").Text                 ''*��۹�ȣ
                    shppSeq             = LagrgeNode(i).SelectSingleNode("shppSeq").Text                ''*��ۼ���
                    ordNo               = LagrgeNode(i).SelectSingleNode("ordNo").Text                  ''*�ֹ���ȣ [20171123128379]
                    ordItemSeq          = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text             ''*�ֹ�����
                    orordNo             = LagrgeNode(i).SelectSingleNode("orordNo").Text            ''���ֹ���ȣ [20171123128379]
                    orordItemSeq        = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''���ֹ����� [2]


                    shppTabProgStatCd   = LagrgeNode(i).SelectSingleNode("shppTabProgStatCd").Text      ''������ۻ���������ڵ�(��۴���) 11 ������� 21 ��ŷ���� 22 ��ŷ�Ϸ� 31 ��ŷ�Ϸ� 41 ����� 42 ������� 43 ���Ϸ� 51 ��ۿϷ� 52 ��۰���
                    delayNts            = LagrgeNode(i).SelectSingleNode("delayNts").Text               ''����Ƚ��
                    lastShppProgStatDtlNm   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlNm").Text  ''������ۻ�������¸�(��ۻ�ǰ����) [��ŷ�Ϸ�]
                    lastShppProgStatDtlCd   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''������ۻ���������ڵ�(��ۻ�ǰ����) 11 ������� 21 ��ŷ���� 22 ��ŷ�Ϸ� 31 ��ŷ�Ϸ� 41 ����� 42 ������� 43 ���Ϸ� 51 ��ۿϷ� 52 ��۰���
                    shppTypeNm          = LagrgeNode(i).SelectSingleNode("shppTypeNm").Text     ''���������    [�ù���]
                    shppStatCd          = LagrgeNode(i).SelectSingleNode("shppStatCd").Text         ''*��ۻ����ڵ� 10 ���� 30 ���
                    shppStatNm          = LagrgeNode(i).SelectSingleNode("shppStatNm").Text         ''��ۻ��¸�

                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatNm") is Nothing) then
                        shortgProgStatNm         = LagrgeNode(i).SelectSingleNode("shortgProgStatNm").Text    ''��ǰ������¸�  //���ð�
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("whoutCritnDt") is Nothing) then
                        whoutCritnDt         = LagrgeNode(i).SelectSingleNode("whoutCritnDt").Text    ''��������
                    end if


                    if ((ordNo=ioutmallorderserial) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) _
                        or ((orordNo=ioutmallorderserial) and (shppNo=ishppNo) and (shppSeq=ishppSeq)) then

                        rw ordNo&"|"&ordItemSeq&"|"&orordNo&"|"&orordItemSeq&"|"&shppNo&"|"&shppSeq
                        rw shppTabProgStatCd&"|"&delayNts&"|"&lastShppProgStatDtlNm&"|"&lastShppProgStatDtlCd&"|"&shppTypeNm&"|"&shppStatCd&"|"&CHKIIF(shppStatNm="���","<strong><font color=red>","")&shppStatNm&CHKIIF(shppStatNm="���","</font></strong>","")&"|"&shortgProgStatNm&"|"&whoutCritnDt
                        successCnt = successCnt + 1 ''AssignedRow

                        if (shppStatNm<>"���") and (lastShppProgStatDtlNm="��ŷ�Ϸ�") then
                            ''���������� ����/����/�� �̶��..
                            if (replace(CSTR(dateadd("d",3,NOW())),"-","")>=whoutCritnDt) then
                                IsreqForceSend = true
                            end if

                            if (NOT IsreqForceSend) and (dlvfinishdt<>"") then
                                ''�츮�� ��ۿϷ��� N�� ���������. => �츮�� ���� ��ۿϷ��̸� �����Ѵ�.
                                ''if (DateDiff("d",dlvfinishdt,now())>=2) then
                                    IsreqForceSend = true
                                ''end if
                            end if

                            if (NOT IsreqForceSend) then
                                if (request("isfrcsend")="1") then
                                    IsreqForceSend = true
                                ELSE
                                    Dim reqURI : reqURI="?shppNo="&request("shppNo")&"&shppSeq="&request("shppSeq")&"&delicoVenId="&request("delicoVenId")&"&wblno="&request("wblno")&"&itemno="&request("itemno")&"&outmallorderserial="&request("outmallorderserial")&"&orgdetailKey="&request("orgdetailKey")&"&dlvfinishdt="&request("dlvfinishdt")&"&prctp="&request("prctp")&"&isfrcsend=1"
                                    rw "<br><input type='button' value='�ڻ��� ����' onClick=""location.href='"&reqURI&"'"">"
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
'                    evntSeq             = LagrgeNode(i).SelectSingleNode("evntSeq").Text                ''�̺�Ʈ����
'                    shppDivDtlCd        = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text           ''*��۱��л��ڵ� 11 �Ϲ���� 12 �κ���� 14 ���� 15 ��ȯ��� 16 AS���
'                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''��۱��л󼼸�
'                    reOrderYn           = LagrgeNode(i).SelectSingleNode("reOrderYn").Text              ''*�����ÿ��α���
'
'                    ordCmplDts          = LagrgeNode(i).SelectSingleNode("ordCmplDts").Text             ''*�ֹ��Ϸ��Ͻ� [2017-11-23 10:39:42.0]
'
'
'                    salestrNo           = LagrgeNode(i).SelectSingleNode("salestrNo").Text              '' [6004]
'                    shppVenId           = LagrgeNode(i).SelectSingleNode("shppVenId").Text      ''���޾�ü���̵� [0000003198]
'                    shppVenNm           = LagrgeNode(i).SelectSingleNode("shppVenNm").Text      ''���޾�ü��
'
'                    shppTypeCd          = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text     ''��������ڵ� 10 �ڻ��� 20 �ù��� 30 ����湮 40 ��� 50 �̹�� 60 �̹߼�
'                    shppTypeDtlCd       = LagrgeNode(i).SelectSingleNode("shppTypeDtlCd").Text  ''����������ڵ� 14 ��ü�ڻ��� 22 ��ü�ù��� 25 �ؿ��ù��� 31 ����湮 41 ��� 51 SMS 52 EMAIL 61 �̹߼�
'                    shppTypeDtlNm       = LagrgeNode(i).SelectSingleNode("shppTypeDtlNm").Text  ''��������󼼸� [��ü�ù���]
'                    delicoVenId         = LagrgeNode(i).SelectSingleNode("delicoVenId").Text    ''�ù��ID [0000033011]
'                    boxNo               = LagrgeNode(i).SelectSingleNode("boxNo").Text          ''�ڽ���ȣ [398327952]
'                    shppcst             = LagrgeNode(i).SelectSingleNode("shppcst").Text        '' ��ۺ�? [303] ??
'                    shppcstCodYn        = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text   ''*��ۺ� ���ҿ��� Y: ���� N: ����
'                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*��ǰ��
'                    splVenItemId        = LagrgeNode(i).SelectSingleNode("splVenItemId").Text       ''*��ü��ǰ��ȣ [1024019]
'                    itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*��ǰ��ȣ [1000024811163]
'                    uitemId             = LagrgeNode(i).SelectSingleNode("uitemId").Text            ''*��ǰID [00000]
'                    dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''���ü��� [2]
'                    cnclItemQty         = LagrgeNode(i).SelectSingleNode("cnclItemQty").Text        ''��Ҽ��� [0]
'                    ordQty              = LagrgeNode(i).SelectSingleNode("ordQty").Text             ''�ֹ����� [2]
'                    sellprc             = LagrgeNode(i).SelectSingleNode("sellprc").Text            ''�ǸŰ� [1000]
'                    frgShppYn           = LagrgeNode(i).SelectSingleNode("frgShppYn").Text          ''����/�� ���� [����]
'                    ordpeNm             = LagrgeNode(i).SelectSingleNode("ordpeNm").Text            ''*�ֹ���
'                    rcptpeNm            = LagrgeNode(i).SelectSingleNode("rcptpeNm").Text           ''*������
'                    rcptpeHpno          = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text         ''*������ �޴�����ȣ
'                    rcptpeTelno         = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text        ''*������ ����ȭ��ȣ
'                    shpplocAddr         = LagrgeNode(i).SelectSingleNode("shpplocAddr").Text        ''������ ���ּ�
'                    shpplocZipcd        = LagrgeNode(i).SelectSingleNode("shpplocZipcd").Text       ''*������ �����ȣ          [04733]
'                    shpplocOldZipcd     = LagrgeNode(i).SelectSingleNode("shpplocOldZipcd").Text    ''*������ �������ȣ(6�ڸ�)  [133750]
'                    shpplocRoadAddr     = LagrgeNode(i).SelectSingleNode("shpplocRoadAddr").Text    ''�����ε��θ��ּ�
'                    itemChrctDivCd      = LagrgeNode(i).SelectSingleNode("itemChrctDivCd").Text     ''��ǰƯ�������ڵ� 10 �Ϲ� 20 ���θ� 30 �ؿܱ��Ŵ����ǰ 40 �̰����ͱݼ� 50 ����ϱ���Ʈ 60 ��ǰ�� 70 ���������� 80 ����ϻ�ǰ�� 91 �̺�Ʈ
'
'
'                    shppMainCd          = LagrgeNode(i).SelectSingleNode("shppMainCd").Text         ''�����ü�ڵ� 32 ��üâ�� 41 ���¾�ü 42 �귣������  [41]
'                    siteNo              = LagrgeNode(i).SelectSingleNode("siteNo").Text             ''����Ʈ��ȣ 6001 �̸�Ʈ�� 6002 Ʈ���̴����� 6003 �н��� 6004 �ż���� 6005 S.COM�� 6009 �ż����ȭ����
'                    siteNm              = LagrgeNode(i).SelectSingleNode("siteNm").Text             ''����Ʈ��
'                    shppRsvtDt          = LagrgeNode(i).SelectSingleNode("shppRsvtDt").Text
'                    splprc              = LagrgeNode(i).SelectSingleNode("splprc").Text             ''���ް�
'                    shortgYn            = LagrgeNode(i).SelectSingleNode("shortgYn").Text
'                    newWblNoData        = LagrgeNode(i).SelectSingleNode("newWblNoData").Text
'                    newRow              = LagrgeNode(i).SelectSingleNode("newRow").Text
'                    itemDiv             = LagrgeNode(i).SelectSingleNode("itemDiv").Text                ''�ǸźҰ���û���� 10:�Ϲ� 20: ���� GIFT �Ϲ� 30: ���� GIFT ���� 40: ���� GIFT ����
'                    shpplocBascAddr     = LagrgeNode(i).SelectSingleNode("shpplocBascAddr").Text        ''�������ּ� 20170712
'                    shpplocDtlAddr      = LagrgeNode(i).SelectSingleNode("shpplocDtlAddr").Text         ''�����λ��ּ�	20170712
'                    ordItemDivNm        = LagrgeNode(i).SelectSingleNode("ordItemDivNm").Text           ''�ֹ���ǰ����	20170809
'
'
'                    ''//�ʼ��� �ƴѰ�� .
'                    if NOT (LagrgeNode(i).SelectSingleNode("ordpeHpno") is Nothing) then
'                        ordpeHpno         = LagrgeNode(i).SelectSingleNode("ordpeHpno").Text           ''�ֹ����޴�����ȣ  //���ð�
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
'                        ordMemoCntt         = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[����۸޸�]","")            ''����۸޸�  //���ð�
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("pCus") is Nothing) then
'                        pCus         = LagrgeNode(i).SelectSingleNode("pCus").Text           ''�������������ȣ  //���ð�
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("frebieNm") is Nothing) then
'                        frebieNm         = LagrgeNode(i).SelectSingleNode("frebieNm").Text    ''����ǰ  //���ð�
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatCd") is Nothing) then
'                        shortgProgStatCd         = LagrgeNode(i).SelectSingleNode("shortgProgStatCd").Text    ''�ǸźҰ���û����  //���ð� 11 ��ǰ��� 12 ��ǰCSó���� 13 ��ǰȮ�� 21 ��ǰ����������� 22 ��ǰ��������CSó���� 23 ��ǰ��������Ȯ�� 41 �԰�������� 43 �԰������Ϸ� 51 ����������
'                    end if
'
'                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatNm") is Nothing) then
'                        shortgProgStatNm         = LagrgeNode(i).SelectSingleNode("shortgProgStatNm").Text    ''��ǰ������¸�  //���ð�
'                    end if
'
'                    ''�ɼǸ�
'                    if NOT (LagrgeNode(i).SelectSingleNode("uitemNm") is Nothing) then
'                        uitemNm         = LagrgeNode(i).SelectSingleNode("uitemNm").Text                 ''[,1:^:asdasd:^:�ֹ�����1,2:^:asdasdddd:^:�ֹ�����2]
'                    end if

			    Next
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing

	rw "======================================"
	rw "�ѻ󼼰Ǽ�:"&totalcnt
	''rw "��۹�ȣ ������Ʈ:"&successCnt
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
