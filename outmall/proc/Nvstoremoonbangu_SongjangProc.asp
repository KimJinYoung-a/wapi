<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/nvstoremoonbangu/nvstoremoonbanguItemcls.asp"-->
<!-- #include virtual="/outmall/nvstoremoonbangu/incNvstoremoonbanguFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Function getNvstorefarmSongjangXMLStr(masterno, detailno, delicompCd, wbNo, isDiv)
'delicompCd : 주문번호
'wbNo		: 송장
'delicompCd	: 택배코드
'rw isDiv
'response.end
	Dim oaccessLicense, oTimestamp, osignature

	If delicompCd = "" Then
		delicompCd = "CH1"		'택배코드가 안 넘어왔으면 기타(CH1)로 넘기게..
	End If
	Call getsecretKey(oaccessLicense, oTimestamp, osignature, "SellerService41", "ShipProductOrder")

	Dim strRst, sSql, dName
	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">"
	strRst = strRst & "	<soapenv:Header/>"
	strRst = strRst & "	<soapenv:Body>"
	strRst = strRst & "		<sel:ShipProductOrderRequest>"
	strRst = strRst & "			<sel:AccessCredentials>"
	strRst = strRst & "				<sel:AccessLicense>"&oaccessLicense&"</sel:AccessLicense>"
	strRst = strRst & "				<sel:Timestamp>"&oTimestamp&"</sel:Timestamp>"
	strRst = strRst & "				<sel:Signature>"&osignature&"</sel:Signature>"
	strRst = strRst & "			</sel:AccessCredentials>"
	strRst = strRst & "			<sel:RequestID>ncp_1np6kl_01</sel:RequestID>"
'	strRst = strRst & "			<sel:DetailLevel>?</sel:DetailLevel>"
	strRst = strRst & "			<sel:Version>2.0</sel:Version>"
	strRst = strRst & "			<sel:ProductOrderID>"&detailno&"</sel:ProductOrderID>"								'상품 주문 번호
	If delicompCd = "ETC1" OR delicompCd = "ETC2" Then
		If delicompCd = "ETC1" Then		'퀵서비스
			strRst = strRst & "			<sel:DeliveryMethodCode>QUICK_SVC</sel:DeliveryMethodCode>"					'배송 방법 코드 | QUICK_SVC : 퀵서비스
		ElseIf delicompCd = "ETC2" Then	'기타
			strRst = strRst & "			<sel:DeliveryMethodCode>DIRECT_DELIVERY</sel:DeliveryMethodCode>"			'배송 방법 코드 | DIRECT_DELIVERY : 직접전달
		End If
	Else
		strRst = strRst & "			<sel:DeliveryMethodCode>DELIVERY</sel:DeliveryMethodCode>"						'배송 방법 코드 | DELIVERY : 택배, 등기, 소포
		strRst = strRst & "			<sel:DeliveryCompanyCode>"&delicompCd&"</sel:DeliveryCompanyCode>"				'택배사 코드
		strRst = strRst & "			<sel:TrackingNumber>"&wbNo&"</sel:TrackingNumber>"								'송장 번호
	End If
	strRst = strRst & "			<sel:DispatchDate>"&FormatDate(now(), "0000-00-00T00:00:00")&"</sel:DispatchDate>"	'배송일
	strRst = strRst & "		</sel:ShipProductOrderRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	getNvstorefarmSongjangXMLStr = strRst
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
	sqlStr = sqlStr & "	and sellsite='nvstoremoonbangu'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If

'###############################################################################################################################################################
Dim strSql, actCnt
Dim AssignedCNT, objXML, iMessage, xmlDOM
Dim ord_no     : ord_no     = request("ord_no")
Dim ord_dtl_sn : ord_dtl_sn = request("ord_dtl_sn")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("inv_no"), 15)					'// 15자 넘으면 에러
Dim s_Div : s_Div = request("songjangDiv")
actCnt = 0			'실갱신건수
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no
Dim xmlStr : xmlStr = getNvstorefarmSongjangXMLStr(ord_no, ord_dtl_sn, hdc_cd, inv_no, s_Div)
Dim retDoc, sURL
Dim successYn, errorMsg, nvstorefarmURL
Dim ResponseType
'/////////////////////////////////////
nvstorefarmURL = "http://ec.api.naver.com/ShopN/SellerService41"
Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", nvstorefarmURL, False
	objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
	objXML.setRequestHeader "SOAPAction", "SellerService41#ShipProductOrder"
	objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			On Error Resume Next
			ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
			If ResponseType = "SUCCESS" Then
			Else
				iMessage = xmlDOM.getElementsByTagName("n:Message")(0).Text
			End If
			On Error Goto 0
		Set xmlDOM = nothing
	End If
Set objXML = nothing
'////////////////////////////////////

Dim IsSuccss : IsSuccss=(ResponseType="SUCCESS")
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
		response.write "    var uri = 'Nvstoremoonbangu_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If

end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->