<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/lotteon/lotteonItemcls.asp"-->
<!-- #include virtual="/outmall/lotteon/inclotteonFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->

<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body bgcolor="#F4F4F4" >
<%
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
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	sqlStr = sqlStr & "	and sellsite='lotteon'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If
'###############################################################################################################################################################
Dim strSql, actCnt, iRbody, xmlDOM, strObj
Dim AssignedCNT, objXML, retCode, iMessage
Dim ord_no     : ord_no     = request("OutMallOrderSerial")
Dim ord_dtl_sn : ord_dtl_sn = request("OrgDetailKey")
Dim deliveryCompanyCode     : deliveryCompanyCode     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("songjangNo"), 15)					'// 15자 넘으면 에러
Dim outmallGoodNo		: outmallGoodNo = request("outmallGoodNo")
Dim outmallOptionCode	: outmallOptionCode = request("outmallOptionCode")
Dim beasongNum			: beasongNum = request("beasongNum")
Dim sendQnt				: sendQnt = request("sendQnt")
Dim objJson

actCnt = 0			'실갱신건수
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no
Dim retDoc
Dim errorMsg, istrParam
Dim IsSuccss
Dim errCount : errCount = 0
Dim isOk, xmlURL
Dim kjytest				: kjytest = request("kjytest")
'/////////////////////////////////////

SET objJson = jsObject()
	Set objJson("deliveryProgressStateList")= jsArray()
		Set objJson("deliveryProgressStateList")(null) = jsObject()
			objJson("deliveryProgressStateList")(null)("odNo") = ORG_ord_no					'#주문번호 : 주문테이블의 PK속성
			objJson("deliveryProgressStateList")(null)("odSeq") = ord_dtl_sn				'#주문순번 : 주문내역에 대해서 단품별로 부여되는 속성값 1
			objJson("deliveryProgressStateList")(null)("procSeq") = beasongNum				'#처리순번 : Default 1 단품단위로 처리순서값을 정의함. 최초 입력시 1 이고 클레임이 발생할 경우 1씩 증가함
			objJson("deliveryProgressStateList")(null)("odPrgsStepCd") = "13"				'#주문진행단계 | 11 : 출고지시, 12 : 상품준비, 13 : 발송완료, 14 : 배송완료, 15 : 수취완료, 23 : 회수지시, 24 : 회수진행, 25 : 회수완료, 26 : 회수확정
			objJson("deliveryProgressStateList")(null)("dvTrcStatDttm") = FormatDate(now(), "00000000000000")	'#배송상태발생일시
			objJson("deliveryProgressStateList")(null)("invcNbr") = 1						'송장개수 : 하나의 단품에 대해서 송장이 분리될 경우 송장수를 의미
			objJson("deliveryProgressStateList")(null)("dvCoCd") = deliveryCompanyCode		'배송사코드
			objJson("deliveryProgressStateList")(null)("invcNo") = inv_no					'송장번호 : Packing단위로 붙여진 운송장번호 주문진행단계가 13:발송완료,25:회수완료인 경우 송장개수 필수
			objJson("deliveryProgressStateList")(null)("spdNo") = outmallGoodNo				'#상품번호 : 롯데ON에서 관리되는 상품번호
			objJson("deliveryProgressStateList")(null)("sitmNo") = outmallOptionCode		'#단품번호 : 롯데ON에서 관리되는 단품번호
			objJson("deliveryProgressStateList")(null)("slQty") = sendQnt					'수량 : 단품에 대한 주문수량
	istrParam = objJson.jsString
SET objJson = nothing
'response.end

Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "POST", APIURL & "/v1/openapi/delivery/v1/SellerDeliveryProgressStateInform", false
	objXML.setRequestHeader "Authorization", "Bearer " & APIkey
	objXML.setRequestHeader "Accept", "application/json"
	objXML.setRequestHeader "Accept-Language", "ko"
	objXML.setRequestHeader "X-Timezone", "GMT+09:00"
	objXML.setRequestHeader "Content-Type", "application/json"
	objXML.Send(istrParam)

'	If kjytest = "Y" Then
		rw objXML.Status
		rw istrParam
		rw iRbody
'	End If

	If objXML.Status <> "200" Then
		IsSuccss = false
		iMessage = "전송 오류"
	Else
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

'		If kjytest = "Y" Then
			rw istrParam
			rw iRbody
'		End If

		Set strObj = JSON.parse(iRbody)
			'If strObj.returnCode <> "0000" Then
			If strObj.data.rsltCd <> "0000" Then
				IsSuccss = false
				iMessage = replaceMsg(strObj.data.rsltMsg)
			Else
				IsSuccss = true
			End If
		Set strObj = nothing
	End If
Set objXML = nothing
'rw iMessage
'response.end
'////////////////////////////////////
if (IsSuccss) then
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendState=1"
	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
    strSql = strSql & "	and beasongNum11st='"&beasongNum&"'"&VBCRLF
	strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
    strSql = strSql & "	and matchstate in ('O')"
	dbget.Execute strSql,AssignedCNT
'rw strSql

    IF (AssignedCNT>0) then
	    if (IsAutoScript) then
	        rw "OK|"&ord_no&" "&ord_dtl_sn
	    ELSE
    	    response.write "OK"
    	ENd IF
	Else
		strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
		strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
		strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
		strSql = strSql & "	and beasongNum11st='"&beasongNum&"'"&VBCRLF
		strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
		strSql = strSql & "	and matchstate in ('O','C','Q','A')"
		dbget.Execute strSql

		'만약 에러횟수가 3회가 넘으면 수기처리 가능
		'updateSendState = 951		기전송 내역
		'updateSendState = 952		취소주문
		strSql = ""
		strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
		strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
		strSql = strSql & "	and beasongNum11st='"&beasongNum&"'"
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
			response.write "<input type='button' value='완료처리' onClick=""fnSetSendState('"&ORG_ord_no&"','"&beasongNum&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
			response.write "<script language='javascript'>"&VbCRLF
			response.write "function fnSetSendState(ORG_ord_no, beasongNum, ord_dtl_sn, selectValue){"&VbCRLF
			response.write "    if(selectValue == ''){"&VbCRLF
			response.write "    	alert('선택해주세요');"&VbCRLF
			response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
			response.write "    	return;"&VbCRLF
			response.write "    }"&VbCRLF
			response.write "    var uri = 'Lotteon_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&beasongNum='+beasongNum+'&updateSendState='+selectValue;"&VbCRLF
			response.write "    location.replace(uri);"&VbCRLF
			response.write "}"&VbCRLF
			response.write "</script>"&VbCRLF
		End If
    ENd IF
else
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
    strSql = strSql & "	and beasongNum11st='"&beasongNum&"'"&VBCRLF
    strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
    strSql = strSql & "	and matchstate in ('O','C','Q','A')"
	dbget.Execute strSql

    rw "<font color=red>"&iMessage&"</font>"

    rw ord_no
    rw ord_dtl_sn
'    rw hdc_cd
    rw inv_no

	'만약 에러횟수가 3회가 넘으면 수기처리 가능
	'updateSendState = 951		기전송 내역
	'updateSendState = 952		취소주문
	strSql = ""
	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
	strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
	strSql = strSql & "	and beasongNum11st='"&beasongNum&"'"
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
		response.write "<input type='button' value='완료처리' onClick=""fnSetSendState('"&ORG_ord_no&"','"&beasongNum&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
		response.write "<script language='javascript'>"&VbCRLF
		response.write "function fnSetSendState(ORG_ord_no, beasongNum, ord_dtl_sn, selectValue){"&VbCRLF
		response.write "    if(selectValue == ''){"&VbCRLF
		response.write "    	alert('선택해주세요');"&VbCRLF
		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
		response.write "    	return;"&VbCRLF
		response.write "    }"&VbCRLF
		response.write "    var uri = 'Lotteon_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&beasongNum='+beasongNum+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If
end if
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->