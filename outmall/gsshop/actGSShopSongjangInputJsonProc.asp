<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/gsshop/gsshopItemcls.asp"-->
<!-- #include virtual="/outmall/gsshop/incGSShopFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Function getGSShopSongjangJsonStr(ordclmNo, ordSeq, delvEntrNo, invoNo)
	Dim ordNo, ordItemNo
	Dim obj
	'2015-09-17 김진영 하단 If문 추가
	If Ubound(Split(ordclmNo,"_")) > 0 Then
		ordNo = Split(ordclmNo,"_")(0)
		ordNo = Right(("0000000000" & ordNo), 10)
	Else
		ordNo = Right(("0000000000" & ordclmNo), 10)
	End If
	ordItemNo = ordSeq

	Set obj = jsObject()
		'obj("sender") = "10X10"
		obj("sender") = "TBT"
		obj("receiver") = "GS SHOP"
		obj("documentId") = "DLVINF"
		obj("processType") = "C"
		obj("ordNo") = CStr(ordNo)
		obj("ordItemNo") = CStr(ordItemNo)
		obj("deliveryCd") = CStr(delvEntrNo)
		obj("deliveryNo") = CStr(invoNo)
		getGSShopSongjangJsonStr = obj.jsString
	Set obj = nothing
End Function

dim ordclmNo    : ordclmNo=request("ordclmNo")      ''인터파크 주문번호
dim ordSeq      : ordSeq=request("ordSeq")          ''인터파크 주문순번
dim delvEntrNo  : delvEntrNo=request("delvEntrNo")  ''택배사코드
dim invoNo      : invoNo=request("invoNo")          ''운송장번호 숫자만 가능함.
dim reqJson
dim errCount

invoNo= trim(replace(invoNo,"-",""))


'2013/02/28 진영추가
dim mode      : mode=request("mode")
If mode = "updateSendState" Then
	Dim sqlStr, AssignedRow
	sqlStr = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr & "	Set sendState='"&request("updateSendState")&"'"
	sqlStr = sqlStr & "	,sendReqCnt=sendReqCnt+1"
	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ordclmNo")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ordSeq")&"'"
	sqlStr = sqlStr & "	and sellsite='gseshop'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If
'2013/02/28 진영추가 끝

reqJson = getGSShopSongjangJsonStr(ordclmNo, ordSeq, delvEntrNo, invoNo)

'// 변수선언
Dim strSql, actCnt
Dim AssignedCNT
actCnt = 0			'실갱신건수

Dim iResult, iMessage
Dim iUrl, replyXML, ErrMsg
iUrl = "http://realapi.gsshop.com/b2b/aliaSupCommonReceiveOrderInfo.gs"

Dim objXML, iRbody, strObj, resultcode, resultmsg
Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", iUrl, false
	objXML.setRequestHeader "Content-Type", "application/json"
	objXML.Send(reqJson)
	If objXML.Status = "200" OR objXML.Status = "201" Then
		iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
			rw "req : " & reqJson
			rw "ret : " & iRbody
		End If
		Set strObj = JSON.parse(iRbody)
			resultcode	= strObj.resultCd
			resultmsg	= strObj.resultMsg
			If resultcode <> "S" Then
				ErrMsg = resultmsg
			Else
				ErrMsg = ""
			End If
		Set strObj = nothing
	End If
Set objXML = nothing

If (ErrMsg <> "") Then
	If (IsAutoScript) Then
		rw "GS샵 송장입력중  오류가 발생했습니다. "&ordclmNo&" "&ordclmNo&"_"&ordSeq
	Else
		Response.Write "<script language=javascript>alert('GS샵 송장입력중  오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
		rw ErrMsg
	End If

'' 시도 회수 추가 sendReqCnt
	strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
	strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
	strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
	strSql = strSql & "	and sellsite='gseshop'"
	strSql = strSql & "	and matchstate in ('O','C','Q','A')"
	dbget.Execute strSql

	strSql = ""
	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
	strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
	strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
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
		"	<option value='901'>발송처리누락 수기등록건</option>" &_
		"	<option value='902'>취소후 제결제건</option>" &_
		"	<option value='903'>반품처리건</option>" &_
		"</select>&nbsp;&nbsp;"
		response.write "<input type='button' value='완료처리' onClick=""finCancelOrd2('"&ordclmNo&"','"&ordSeq&"',document.getElementById('updateSendState').value)"">"
		response.write "<script language='javascript'>"&VbCRLF
		response.write "function finCancelOrd2(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
		response.write "    if(selectValue == ''){"&VbCRLF
		response.write "    	alert('선택해주세요');"&VbCRLF
		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
		response.write "    	return;"&VbCRLF
		response.write "    }"&VbCRLF
		response.write "    var uri = 'actGSShopSongjangInputJsonProc.asp?mode=updateSendState&ordclmNo='+ORG_ord_no+'&ordSeq='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If
Else
	IF (resultcode = "S") Then
		rw "성공"

		strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
		strSql = strSql & "	Set sendState=1"
		strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
		strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
		strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
		strSql = strSql & "	and sellsite='gseshop'"
		strSql = strSql & "	and matchstate in ('O')"
		''rw strSql
		dbget.Execute strSql,AssignedCNT
		actCnt = actCnt+AssignedCNT

		iMessage = "성공"
	Else
		'' 시도 회수 추가 sendReqCnt
		strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
		strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
		strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
		strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
		strSql = strSql & "	and sellsite='gseshop'"
		strSql = strSql & "	and matchstate in ('O','C','Q','A')"
		dbget.Execute strSql

		iMessage = "<font color=red>ERROR</font>"

		strSql = ""
		strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
		strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
		strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
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
							"	<option value='901'>발송처리누락 수기등록건</option>" &_
							"	<option value='902'>취소후 제결제건</option>" &_
							"	<option value='903'>반품처리건</option>" &_
							"</select>&nbsp;&nbsp;"
			response.write "<input type='button' value='완료처리' onClick=""finCancelOrd2('"&ordclmNo&"','"&ordSeq&"',document.getElementById('updateSendState').value)"">"
			response.write "<script language='javascript'>"&VbCRLF
			response.write "function finCancelOrd2(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
			response.write "    if(selectValue == ''){"&VbCRLF
			response.write "    	alert('선택해주세요');"&VbCRLF
			response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
			response.write "    	return;"&VbCRLF
			response.write "    }"&VbCRLF
			response.write "    var uri = 'actGSShopSongjangInputJsonProc.asp?mode=updateSendState&ordclmNo='+ORG_ord_no+'&ordSeq='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
			response.write "    location.replace(uri);"&VbCRLF
			response.write "}"&VbCRLF
			response.write "</script>"&VbCRLF
		End if
	End If

	If (IsAutoScript) Then
		rw "iMessage="&iMessage&":"&ordclmNo&" "&ordclmNo&"_"&ordSeq
	Else
		rw "iMessage="&iMessage
	End If
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
