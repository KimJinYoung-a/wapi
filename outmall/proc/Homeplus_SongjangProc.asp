<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
Public homeplusAPIURL
Public strInterface
Public homeplusVenderID
Public homepluspasswd

IF application("Svr_Info") = "Dev" THEN
	homeplusAPIURL = "http://112.108.7.201:7006/services/API2?wsdl"
	strInterface = "http://112.108.7.201:7006/api/services/API2"
	homeplusVenderID = "292811"
	homepluspasswd = "qwer1234"
Else
	homeplusAPIURL = "http://api.direct.homeplus.co.kr:17004/services/API2?wsdl"
	strInterface = "http://api.direct.homeplus.co.kr:17004/api/services/API2"
	homeplusVenderID = "292811"
	homepluspasswd = "cube1010!!"
End if

Function getHomeplusSongjangXMLStr(masterno, detailno, delicompCd, wbNo)
	Dim strRst
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
	strRst = strRst & "	<SOAP-ENV:Body>"
	strRst = strRst & "		<m:setReleaseEnd xmlns:m=""" & strInterface & """>"
	strRst = strRst & "		<ReleaseEnd>"
	strRst = strRst & "			<i_ORDNO>"&masterno&"</i_ORDNO>"
	strRst = strRst & "			<i_ORDDETNO>"&detailno&"</i_ORDDETNO>"
	strRst = strRst & "			<s_DELICOMP>"&delicompCd&"</s_DELICOMP>"
	strRst = strRst & "			<s_PARCELNO>"&wbNo&"</s_PARCELNO>"
	strRst = strRst & "		</ReleaseEnd>"
	strRst = strRst & "		</m:setReleaseEnd>"
	strRst = strRst & "	</SOAP-ENV:Body>"
	strRst = strRst & "</SOAP-ENV:Envelope>"
    getHomeplusSongjangXMLStr = strRst
End function

Function getXMLString(mode)
	Dim strRst
	If mode = "login" Then
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/""  xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:ns1=""http://xml.apache.org/axis/"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:"&mode&" xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<venderId>"&homeplusVenderID&"</venderId>"
		strRst = strRst & "			<passwd>"&homepluspasswd&"</passwd>"
		strRst = strRst & "		</m:"&mode&">"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
	End If
	getXMLString = strRst
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
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	sqlStr = sqlStr & "	and sellsite='homeplus'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If

'###############################################################################################################################################################
Dim strSql, actCnt
Dim AssignedCNT, objXML, retCode, iMessage
Dim ord_no     : ord_no     = request("ord_no")
Dim ord_dtl_sn : ord_dtl_sn = request("ord_dtl_sn")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("inv_no"), 15)					'// 15자 넘으면 에러
actCnt = 0			'실갱신건수
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no
Dim xmlStr : xmlStr = getHomeplusSongjangXMLStr(ord_no, ord_dtl_sn, hdc_cd, inv_no)
Dim retDoc, sURL
Dim successYn, errorMsg
'/////////////////////////////////////
Dim xmlDOM, xmlDOM2
Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "POST", "" & homeplusAPIURL & "", False
	objXML.setRequestHeader "CONTENT_TYPE", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", Len(getXMLString("login"))
	objXML.setRequestHeader "SOAPAction", strInterface & "#login"
	objXML.setTimeouts 5000,90000,90000,90000
	objXML.send(getXMLString("login"))
If objXML.Status = "200" Then
	Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
		xmlDOM.async = False
		xmlDOM.ValidateOnParse= True
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
	If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then	'로그인 성공이라면
		objXML.open "post", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#"&mode
		objXML.send(xmlStr)
		If objXML.Status = "200" Then
			Set xmlDOM2 = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM2.async = False
				xmlDOM2.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
				'response.end
				retCode		= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/setReleaseEndResponse/ns1:setReleaseEndReturn/ns1:code").text
				iMessage	= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/setReleaseEndResponse/ns1:setReleaseEndReturn/ns1:message").text
			Set xmlDOM2 = nothing
		End If
	End If
	Set objXML = nothing
End If
'////////////////////////////////////
'rw successYn  (true, false)
'rw errorMsg
'rw successYn
'rw retCode
'response.end
Dim IsSuccss : IsSuccss=(retCode="E0000")
Dim errCount
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
	Else

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
		strSql = ""
		strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
		strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
		strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
		strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
		rsget.Open strSql,dbget,1
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
			response.write "    var uri = 'Homeplus_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
			response.write "    location.replace(uri);"&VbCRLF
			response.write "}"&VbCRLF
			response.write "</script>"&VbCRLF
		End If


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
	strSql = ""
	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
	strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
	strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
	strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
	rsget.Open strSql,dbget,1
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
		response.write "    var uri = 'Homeplus_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If

end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
