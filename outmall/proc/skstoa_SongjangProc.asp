<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/skstoa/skstoaItemcls.asp"-->
<!-- #include virtual="/outmall/skstoa/incskstoaFunction.asp"-->
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
	sqlStr = sqlStr & "	and sellsite='skstoa'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If

'###############################################################################################################################################################
Dim strSql, actCnt, iRbody, xmlDOM
Dim AssignedCNT, objXML, retCode, iMessage
Dim ord_no     : ord_no     = request("OutMallOrderSerial")
Dim ord_dtl_sn : ord_dtl_sn = request("OrgDetailKey")
Dim deliveryCompanyCode     : deliveryCompanyCode     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("songjangNo"), 15)					'// 15자 넘으면 에러
Dim outmallGoodNo		: outmallGoodNo = request("outmallGoodNo")
Dim outmallOptionCode	: outmallOptionCode = request("outmallOptionCode")
Dim objJson

actCnt = 0			'실갱신건수
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no
Dim retDoc, strObj
Dim errorMsg, istrParam
Dim errCount : errCount = 0
Dim isOk, xmlURL, IsSuccss
'/////////////////////////////////////
istrParam = ""
istrParam = istrParam & "linkCode=" & skstoalinkCode				'#연결코드 | SKB에서 부여한 연결코드
istrParam = istrParam & "&entpCode=" & skstoaentpCode				'#업체코드 | SKB에서 부여한 업체코드 6자리
istrParam = istrParam & "&entpId=" & skstoaentpId					'#업체사용자ID | SKB에서 부여한 업체사용자 ID
istrParam = istrParam & "&entpPass=" & skstoaentpPass				'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
istrParam = istrParam & "&orderGb=10"								'#배송구분 | 10:주문, 40:교환
istrParam = istrParam & "&orderNo=" & ord_no						'#주문번호
istrParam = istrParam & "&orderGSeq=" &	Split(ord_dtl_sn,"-")(0)	'#상품순번
istrParam = istrParam & "&orderDSeq=" &	Split(ord_dtl_sn,"-")(1)	'#세트순번
istrParam = istrParam & "&orderWSeq=" &	Split(ord_dtl_sn,"-")(2)	'#처리순번
istrParam = istrParam & "&goodsCode=" & outmallGoodNo				'#판매상품코드
istrParam = istrParam & "&goodsdtCode=" & outmallOptionCode			'#판매단품코드
istrParam = istrParam & "&slipNo=" & inv_no							'#운송장 번호
istrParam = istrParam & "&delyGb=" & deliveryCompanyCode			'#배송사 코드
IsSuccss = false
Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "POST", skstoaAPIURL & "/partner/delivery/delivery-out", false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send(istrParam)

	If objXML.Status = "200" OR objXML.Status = "201" Then
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			retCode			= strObj.code
			iMessage		= strObj.message
			If retCode = "200" Then
				IsSuccss = true
			End If
		Set strObj = nothing
	End If
Set objXML = nothing
'////////////////////////////////////
'rw successYn  (true, false)
'rw iMessage
'rw successYn
'rw errorMsg

if (IsSuccss) then
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendState=1"
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
else
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
    strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
    strSql = strSql & "	and matchstate in ('O','C','Q','A')"

	dbget.Execute strSql

    rw "response : " & iRbody

    rw ord_no
    rw ord_dtl_sn
    rw deliveryCompanyCode
    rw inv_no

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
		response.write "    var uri = 'skstoa_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If

end if
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->