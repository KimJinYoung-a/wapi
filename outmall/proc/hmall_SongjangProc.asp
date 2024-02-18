<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/hmall/hmallItemcls.asp"-->
<!-- #include virtual="/outmall/hmall/inchmallFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body bgcolor="#F4F4F4" >
<%
Dim mode : mode=request("mode")
Dim sqlStr, AssignedRow
If mode = "updateSendState" Then
	sqlStr = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr & "	Set sendState='"&request("updateSendState")&"'"
	sqlStr = sqlStr & "	,sendReqCnt=sendReqCnt+1"

	if (request("updateSendState") = "952") then
		'// 취소주문은 인수전송도 skip
		sqlStr = sqlStr & " , recvSendState = 100 "
		sqlStr = sqlStr & " , recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1 "
	end if

	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ord_no")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	sqlStr = sqlStr & "	and sellsite='hmall1010'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
elseif mode = "updateSendStateCS" then
	sqlStr = "Update db_temp.dbo.tbl_xSite_TMPMiChulList "
	sqlStr = sqlStr & "	Set sendState='"&request("updateSendStateCS")&"'"
	sqlStr = sqlStr & "	,sendReqCnt=sendReqCnt+1"
	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ord_no")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	sqlStr = sqlStr & "	and sellsite='hmall1010'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If

'###############################################################################################################################################################
Dim strSql, actCnt, iRbody, xmlDOM, istrParam
Dim AssignedCNT, objXML, retCode, iMessage
Dim ord_no		: ord_no		= request("OutMallOrderSerial")
Dim ord_dtl_sn	: ord_dtl_sn	= request("OrgDetailKey")
Dim hdc_cd		: hdc_cd		= request("hdc_cd")
Dim inv_no		: inv_no		= Left(request("songjangNo"), 15)					'// 15자 넘으면 에러
Dim beasongNum	: beasongNum	= request("beasongNum")
Dim reserve01	: reserve01		= request("reserve01")

actCnt = 0			'실갱신건수
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no
Dim retDoc, isSuccess, strObj
Dim successYn, errorMsg
dim prctp : prctp = requestCheckvar(request("prctp"),20)

'' 주문을 나눠 입력하는 케이스.
IF (InStr(ord_no,"_")>0) then
	ord_no = getOutmallRefOrgOrderNO(ord_no,ord_dtl_sn,"hmall1010")
end if

''if (prctp="3") then ''배송완료처리
''	istrParam = "DlvstNo="&beasongNum&"&DlvstPtcSeq=" & reserve01 & "&OrdNo=" & ord_no & "&OrdPtcSeq=" & ord_dtl_sn & "&ProcGb=P3&DsrvDlvcoCd="& hdc_cd &"&InvcNo="& inv_no &" "
''elseif (prctp="6") then ''주문확인
''    istrParam = "DlvstNo="&beasongNum&"&DlvstPtcSeq=" & reserve01 & "&OrdNo=" & ord_no & "&OrdPtcSeq=" & ord_dtl_sn & "&ProcGb=P1&DsrvDlvcoCd="& hdc_cd &"&InvcNo="& inv_no &" "
''else
''	istrParam = "DlvstNo="&beasongNum&"&DlvstPtcSeq=" & reserve01 & "&OrdNo=" & ord_no & "&OrdPtcSeq=" & ord_dtl_sn & "&ProcGb=P2&DsrvDlvcoCd="& hdc_cd &"&InvcNo="& inv_no &" "
''end if

istrParam = ""
istrParam = istrParam & "{"
istrParam = istrParam & "  ""DlvstNo"": """ & beasongNum & ""","
istrParam = istrParam & "  ""DlvstPtcSeq"": """ & reserve01 & ""","
istrParam = istrParam & "  ""OrdNo"": """ & ord_no & ""","
istrParam = istrParam & "  ""OrdPtcSeq"": """ & ord_dtl_sn & ""","
select case prctp
    case "3"
        ''배송완료처리
        istrParam = istrParam & "  ""ProcGb"": ""P3"","
    case "22"
        ''CS출고 출고완료
        istrParam = istrParam & "  ""ProcGb"": ""P2"","
    case "33"
        ''CS출고 배송완료처리
        istrParam = istrParam & "  ""ProcGb"": ""P3"","
    case "6"
        ''주문확인
        istrParam = istrParam & "  ""ProcGb"": ""P1"","
    case "7"
        ''주문취소(주문확인취소)
        rw "작업이전"
        dbget.close() : response.end
    case else
        istrParam = istrParam & "  ""ProcGb"": ""P2"","
end select
istrParam = istrParam & "  ""DsrvDlvcoCd"": """ & hdc_cd & ""","
istrParam = istrParam & "  ""InvcNo"": """ & inv_no & """"
istrParam = istrParam & "}"

'/////////////////////////////////////

'ProcGb | P1:주문확인, P2:출고완료, P3:배송완료

'istrParam = "DlvstNo=20181119220317&DlvstPtcSeq=1&OrdNo=20181119400210&OrdPtcSeq=1&ProcGb=P2&DsrvDlvcoCd=123123123&InvcNo=3434343434"
On Error Resume Next
Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "POST", "http://xapi.10x10.co.kr:8080/Orders/Hmall/actionoutput", false
	objXML.setRequestHeader "Content-Type", "application/json"
	objXML.Send(istrParam)

	If objXML.Status = "200" OR objXML.Status = "201" Then
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		response.write iRbody
		Set strObj = JSON.parse(iRbody)
			isSuccess		= strObj.success
		Set strObj = nothing
	Else
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			isSuccess		= strObj.success
			iMessage		= strObj.message
		Set strObj = nothing
	End If
Set objXML = nothing
On Error Goto 0
'////////////////////////////////////
'rw successYn  (true, false)
'rw errorMsg
'rw successYn
'rw errorMsg

Dim errCount : errCount = 0

if (isSuccess) then
    if (prctp="6") or (prctp="22") or (prctp="33") then
	    if (IsAutoScript) then
	        rw "OK|"&ord_no&" "&ord_dtl_sn
	    ELSE
    	    response.write "OK"
    	ENd IF
    else
        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	    strSql = strSql & "	Set sendState="&CHKIIF(prctp="3","2","1")
	    strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
        strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
	    strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
        strSql = strSql & "	and matchstate in ('O')"
	    dbget.Execute strSql,AssignedCNT
        IF (AssignedCNT>0) then
	        if (IsAutoScript) then
	            rw "OK|"&ord_no&" "&ord_dtl_sn
	        ELSE
    	        response.write "OK"
    	    ENd IF
        ENd IF
    end if
else
    rw "<font color=red>"&iMessage&"</font>"
    rw istrParam

    rw ord_no
    rw ord_dtl_sn
    rw hdc_cd
    rw inv_no

    if (prctp="6") or (prctp="22") or (prctp="33") then
        '// CS
        strSql = "Update db_temp.dbo.tbl_xSite_TMPMiChulList "
	    strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
        strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
        strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
        strSql = strSql & "	and sendState in ('N')"
	    dbget.Execute strSql

	    '만약 에러횟수가 3회가 넘으면 수기처리 가능
	    'updateSendState = 951		기전송(취소) 내역
	    strSql = ""
	    strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPMiChulList " & VBCRLF
	    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
	    strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
	    strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
	    rsget.CursorLocation = adUseClient
	    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	    If Not rsget.Eof Then
		    errCount = rsget("cnt")
	    End If
	    rsget.Close

	    If errCount > 0 Then
			response.write  "<select name='updateSendStateCS' id=""updateSendStateCS"">" &_
							"	<option value=''>선택</option>" &_
							"	<option value='X'>기전송(취소) 내역</option>" &_
							"</select>&nbsp;&nbsp;"
			response.write "<input type='button' value='완료처리' onClick=""fnSetSendState('"&ORG_ord_no&"','"&ord_dtl_sn&"',document.getElementById('updateSendStateCS').value)"">"
			response.write "<script language='javascript'>"&VbCRLF
			response.write "function fnSetSendState(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
			response.write "    if(selectValue == ''){"&VbCRLF
			response.write "    	alert('선택해주세요');"&VbCRLF
			response.write "    	document.getElementById('updateSendStateCS').focus();"&VbCRLF
			response.write "    	return;"&VbCRLF
			response.write "    }"&VbCRLF
			response.write "    var uri = 'hmall_SongjangProc.asp?mode=updateSendStateCS&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendStateCS='+selectValue;"&VbCRLF
			response.write "    location.replace(uri);"&VbCRLF
			response.write "}"&VbCRLF
			response.write "</script>"&VbCRLF
	    End If
    else
        strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	    strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
        strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
        strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
        strSql = strSql & "	and matchstate in ('O','C','Q','A')"
	    dbget.Execute strSql


	    '만약 에러횟수가 3회가 넘으면 수기처리 가능
	    'updateSendState = 951		기전송 내역
	    'updateSendState = 952		취소주문
	    strSql = ""
	    strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
	    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
	    strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
	    strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
	    rsget.CursorLocation = adUseClient
	    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	    If Not rsget.Eof Then
		    errCount = rsget("cnt")
	    End If
	    rsget.Close

	    If errCount > 0 Then
			response.write  "<select name='updateSendState' id=""updateSendState"">" &_
							"	<option value=''>선택</option>" &_
							"	<option value='951'>기전송 내역</option>" &_
							"	<option value='952'>취소주문</option>" &_
							"</select>&nbsp;&nbsp;"
			response.write "<input type='button' value='완료처리' onClick=""fnSetSendState('"&ORG_ord_no&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
			response.write "<script language='javascript'>"&VbCRLF
			response.write "function fnSetSendState(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
			response.write "    if(selectValue == ''){"&VbCRLF
			response.write "    	alert('선택해주세요');"&VbCRLF
			response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
			response.write "    	return;"&VbCRLF
			response.write "    }"&VbCRLF
			response.write "    var uri = 'hmall_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
			response.write "    location.replace(uri);"&VbCRLF
			response.write "}"&VbCRLF
			response.write "</script>"&VbCRLF
	    End If
    end if
end if
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
