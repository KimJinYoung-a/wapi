<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/gmarket/gmarketItemcls.asp"-->
<!-- #include virtual="/outmall/gmarket/incGmarketFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<% if NOT (IsAutoScript) then %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body bgcolor="#F4F4F4" >
<% end if %>
<%
Function getGmarketSongjangXMLStr(masterno, detailno, delicompCd, wbNo, isDiv)
'delicompCd : 주문번호
'wbNo		: 송장
'delicompCd	: 택배코드
'rw isDiv
'response.end
	Dim strRst
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
	strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst & "	<soap:Header>"
	strRst = strRst & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "			<encTicket>"&gmarketTicket&"</encTicket>"
	strRst = strRst & "		</EncTicket>"
	strRst = strRst & "	</soap:Header>"
	strRst = strRst & "	<soap:Body>"
	strRst = strRst & "		<AddShipping xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst & "			<AddShipping PackNo="""&masterno&""" ContrNo="""&detailno&""" ExpressName="""&delicompCd&""" InvoiceNo="""&wbNo&""" ShippingDate="""&Date()&""" />"
	strRst = strRst & "		</AddShipping>"
	strRst = strRst & "	</soap:Body>"
	strRst = strRst & "</soap:Envelope>"
	getGmarketSongjangXMLStr = strRst
End function

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
	sqlStr = sqlStr & "	and sellsite='gmarket1010'"
	dbget.Execute sqlStr,AssignedRow
	''response.write sqlStr
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If

'###############################################################################################################################################################
Dim strSql, actCnt
Dim AssignedCNT, objXML, retCode, iMessage, xmlDOM
Dim ord_no     : ord_no     = request("ord_no")
Dim ord_dtl_sn : ord_dtl_sn = request("ord_dtl_sn")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("inv_no"), 15)					'// 15자 넘으면 에러
Dim s_Div : s_Div = request("songjangDiv")
actCnt = 0			'실갱신건수
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no

'' 주문을 나눠 입력하는 케이스.
IF (InStr(ord_no,"_")>0) then
	ord_no = getOutmallRefOrgOrderNO(ord_no,ord_dtl_sn,"gmarket1010")
end if

Dim xmlStr : xmlStr = getGmarketSongjangXMLStr(ord_no, ord_dtl_sn, hdc_cd, inv_no, s_Div)
Dim retDoc, sURL
Dim successYn, errorMsg

'////////////////////////////////////
Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "POST", "" & gmarketAPIURL&"/v1/ShippingService.asmx"
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(xmlStr)
	objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/AddShipping"
	objXML.send(xmlStr)
'	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
'			response.write Replace(objXML.responseText,"soap:","")
'			response.end
			On Error Resume Next
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "AddShippingResponse" Then
				retCode = xmlDOM.getElementsByTagName("AddShippingResult ").item(0).getAttribute("Result")
				iMessage = xmlDOM.getElementsByTagName("AddShippingResult ").item(0).getAttribute("Comment")
			End IF
			On Error Goto 0
		Set xmlDOM = nothing
'	End If
Set objXML = nothing

Dim IsSuccss : IsSuccss=(retCode="Success")
response.write "IsSuccss:"&IsSuccss&"<BR>"

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

    rw "<font color=red>"&iMessage&"</font>"

    rw ord_no
    rw ord_dtl_sn
    rw hdc_cd
    rw inv_no

	'만약 에러횟수가 3회가 넘으면 수기처리 가능
	'updateSendState = 951		기전송 내역
	'updateSendState = 952		취소주문
	Dim errCount : errCount = 0
	strSql = ""
	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
	strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
	strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
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
		''response.write "    alert(selectValue); "&VbCRLF
		response.write "    var uri = 'Gmarket_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If

end if
%>
<% if NOT (IsAutoScript) then %>
</body>
</html>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->