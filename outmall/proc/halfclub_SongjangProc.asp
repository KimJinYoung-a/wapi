<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/halfclub/halfclubItemcls.asp"-->
<!-- #include virtual="/outmall/halfclub/inchalfclubFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
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
	sqlStr = sqlStr & "	and sellsite='halfclub'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If

Public Function getHalfclubDeliveryXMLStr(iord_no, iord_dtl_sn, ioutmallGoodNo, ioutmallOptionCode, ioutmallOptionName, iitemno, ihdc_cd, iinv_no)
	Dim strRst
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst & "	<soap:Header>"
	strRst = strRst & "		<SOAPHeaderAuth xmlns=""http://api.tricycle.co.kr/"">"
	strRst = strRst & " 		<User_ID>"&UPCHECODE&"</User_ID>"
	strRst = strRst & " 		<User_PWD>"&APIKEY&"</User_PWD>"
	strRst = strRst & "		</SOAPHeaderAuth>"
	strRst = strRst & "	</soap:Header>"
	strRst = strRst & "	<soap:Body>"
	strRst = strRst & "		<Set_Delivery xmlns=""http://api.tricycle.co.kr/"">"
	strRst = strRst & "			<req_Delivery>"
	strRst = strRst & "				<OrdNum>"&iord_no&"</OrdNum>"
	strRst = strRst & "				<OrdNum_Nm>"&iord_dtl_sn&"</OrdNum_Nm>"
	strRst = strRst & "				<PCode>"&ioutmallGoodNo&"</PCode>"
	strRst = strRst & "				<OptCd>"&ioutmallOptionCode&"</OptCd>"
	strRst = strRst & "				<OptNm>"&ioutmallOptionName&"</OptNm>"
	strRst = strRst & "				<PrdQty>"&iitemno&"</PrdQty>"
	strRst = strRst & "				<OffNo>"&ihdc_cd&"</OffNo>"
	strRst = strRst & "				<Invoice>"&iinv_no&"</Invoice>"
	strRst = strRst & "			</req_Delivery>"
	strRst = strRst & "		</Set_Delivery>"
	strRst = strRst & "	</soap:Body>"
	strRst = strRst & "</soap:Envelope>"
	getHalfclubDeliveryXMLStr = strRst
End Function

'###############################################################################################################################################################
Dim strSql, actCnt, iRbody, xmlDOM
Dim AssignedCNT, objXML, retCode, iMessage
Dim ord_no     : ord_no     = request("OutMallOrderSerial")
Dim ord_dtl_sn : ord_dtl_sn = request("OrgDetailKey")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("songjangNo"), 15)					'// 15자 넘으면 에러
Dim outmallGoodNo		: outmallGoodNo = request("outmallGoodNo")
Dim outmallOptionCode	: outmallOptionCode = request("outmallOptionCode")
Dim outmallOptionName	: outmallOptionName = request("outmallOptionName")
Dim itemno				: itemno = request("itemno")

actCnt = 0			'실갱신건수
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no
Dim retDoc, sURL
Dim successYn, errorMsg
'/////////////////////////////////////
Dim xmlStr : xmlStr = getHalfclubDeliveryXMLStr(ord_no, ord_dtl_sn, outmallGoodNo, outmallOptionCode, outmallOptionName, itemno, hdc_cd, inv_no)
'response.write replace(xmlStr, "utf-8","euc-kr")
'response.end
Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
	objXML.open "POST", "" & APIURL&"/Delivery/Delivery.asmx"
	objXML.setRequestHeader "Host", "api.tricycle.co.kr"
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(xmlStr)
	objXML.setRequestHeader "SOAPMethodName", "Set_Delivery"
	objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
'			response.write replace(objXML.responseText, "utf-8","euc-kr")
'			response.end
			On Error Resume Next
				retCode		= xmlDOM.getElementsByTagName("ResultCode").Item(0).Text
				iMessage	= Trim(xmlDOM.getElementsByTagName("ResultMsg").item(0).text)
			On Error Goto 0
		Set xmlDOM = nothing
	End If
Set objXML = nothing
'////////////////////////////////////
'rw successYn  (true, false)
'rw errorMsg
'rw successYn
'rw errorMsg
Dim IsSuccss : IsSuccss=(retCode="0000")

if (IsSuccss) OR (INSTR(iMessage, "출고 정보 중복 오류") > 0)  Then
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
	strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
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
		response.write "    var uri = 'halfclub_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If

end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->