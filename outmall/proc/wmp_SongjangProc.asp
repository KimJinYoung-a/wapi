<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/wmp/wmpItemcls.asp"-->
<!-- #include virtual="/outmall/wmp/incWmpFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Function fnWMPConfirmOrder(vOrderserial)
	Dim objXML, xmlDOM, iRbody, strSql, istrParam, isSuccessCode, strObj
	'istrParam = "bundleNo="&vOrderserial
	istrParam = "bundleNo="&vOrderserial
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info") = "Dev" Then
			objXML.open "POST", "http://localhost:62569/Wemake/Orders/ordercomplete", false
		Else
			objXML.open "POST", "http://110.93.128.100:8090/wemake/Orders/ordercomplete", false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If Err.number <> 0 Then
			iErrStr = ivendorItemId
			Exit Function
		End If
		rw objXML.Status
		rw BinaryToText(objXML.ResponseBody,"utf-8")

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
			Set strObj = JSON.parse(iRbody)
				isSuccessCode		= strObj.code
				If isSuccessCode = "200" Then
					strSql = ""
					strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] SET "
					strSql = strSql & " isbaljuConfirmSend = 'Y' "
					strSql = strSql & " , lastUpdate = getdate() "
					strSql = strSql & " WHERE outmallorderserial = '"&vOrderserial&"'  "
					strSql = strSql & " and mallid = 'WMP' "
					dbget.Execute strSql
					fnWMPConfirmOrder= true
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
	On Error Goto 0
End Function


Dim mode : mode=request("mode")
If mode = "updateSendState" Then
	Dim sqlStr, AssignedRow
	sqlStr = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr & "	Set sendState='"&request("updateSendState")&"'"
	sqlStr = sqlStr & "	,sendReqCnt=sendReqCnt+1"

	if (request("updateSendState") = "952") then
		'// 취소주문은 인수전송도 skip
		sqlStr = sqlStr & " , recvSendState = 100 "
		sqlStr = sqlStr & " , recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1 "
	end if

	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ord_no")&"'"
	sqlStr = sqlStr & "	and sellsite='WMP'"
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
Dim successYn, errorMsg, ret1


IF (InStr(ord_no,"_")>0) then
	ord_no = getOutmallRefOrgOrderNO(ord_no,ord_dtl_sn,"WMP")
end if

'/////////////////////////////////////

'shipMethod | 택배배송:PARCEL, 직접배송:DIRECT, 우편배송:POST, 기타배송:ETC)
If hdc_cd = "0" Then
	istrParam = "bundleNo=" & ord_no & "&shipMethod=ETC&shipMethodMessage=" & inv_no
Else
	istrParam = "bundleNo=" & ord_no & "&shipMethod=PARCEL&parcelCompanyCode=" & hdc_cd & "&invoiceNo="& inv_no
End If
'istrParam = "DlvstNo=20181119220317&DlvstPtcSeq=1&OrdNo=20181119400210&OrdPtcSeq=1&ProcGb=P2&DsrvDlvcoCd=123123123&InvcNo=3434343434"
On Error Resume Next
Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	If application("Svr_Info") = "Dev" Then
		objXML.open "POST", "http://localhost:62569/Wemake/Orders/outputproc", false
	Else
		objXML.open "POST", "http://110.93.128.100:8090/wemake/Orders/outputproc", false
	End If
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send(istrParam)

	If objXML.Status = "200" OR objXML.Status = "201" Then
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		response.write iRbody
		Set strObj = JSON.parse(iRbody)
			isSuccess		= strObj.success
		Set strObj = nothing
	Else
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		response.write iRbody
		Set strObj = JSON.parse(iRbody)
			isSuccess		= strObj.success
			iMessage		= strObj.message
		Set strObj = nothing
	End If
Set objXML = nothing
On Error Goto 0

' If session("ssBctID")="kjy8517" then
' 	response.end
' End If
'////////////////////////////////////
'rw successYn  (true, false)
'rw errorMsg
'rw successYn
'rw errorMsg

If (INSTR(iMessage, "주문확인건만 발송처리 가능합니다") > 0) Then
	ret1 = fnWMPConfirmOrder(ord_no)
	If (ret1) then
		On Error Resume Next
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			If application("Svr_Info") = "Dev" Then
				objXML.open "POST", "http://localhost:62569/Wemake/Orders/outputproc", false
			Else
				objXML.open "POST", "http://110.93.128.100:8090/wemake/Orders/outputproc", false
			End If
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
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
		'///////////////////
	End If
End If

if (isSuccess) OR (INSTR(iMessage, "이미 발송된 상품입니다") > 0)  Then
	If (INSTR(iMessage, "이미 발송된 상품입니다") > 0) Then
		strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
		strSql = strSql & "	Set sendState=1"
		strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
		strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
		dbget.Execute strSql,AssignedCNT
	Else
		strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
		strSql = strSql & "	Set sendState=1"
		strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
		strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
		strSql = strSql & "	and matchstate in ('O')"
		dbget.Execute strSql,AssignedCNT
	End If

    IF (AssignedCNT>0) then
	    if (IsAutoScript) then
	        rw "OK|"&ord_no&" "&ord_dtl_sn
	    ELSE
    	    response.write "OK"
    	ENd IF
    ENd IF
else
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
'   strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
    strSql = strSql & "	and matchstate in ('O','C','Q','A')"
	dbget.Execute strSql

    rw "<font color=red>"&iMessage&"</font>"

    rw ord_no
'    rw ord_dtl_sn
    rw hdc_cd
    rw inv_no

	'만약 에러횟수가 3회가 넘으면 수기처리 가능
	'updateSendState = 951		기전송 내역
	'updateSendState = 952		취소주문
	Dim errCount : errCount = 0
	strSql = ""
	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
	strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
'	strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
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
		response.write "<input type='button' value='완료처리' onClick=""fnSetSendState('"&ORG_ord_no&"',document.getElementById('updateSendState').value)"">"
		response.write "<script language='javascript'>"&VbCRLF
		response.write "function fnSetSendState(ORG_ord_no,selectValue){"&VbCRLF
		response.write "    if(selectValue == ''){"&VbCRLF
		response.write "    	alert('선택해주세요');"&VbCRLF
		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
		response.write "    	return;"&VbCRLF
		response.write "    }"&VbCRLF
		response.write "    var uri = 'wmp_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If

end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->