<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/11st/11stItemcls.asp"-->
<!-- #include virtual="/outmall/11st/inc11stFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
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
	sqlStr = sqlStr & "	and beasongNum11st='"&request("beasongNum")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	sqlStr = sqlStr & "	and sellsite='11st1010'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If

'###############################################################################################################################################################
Dim strSql, actCnt, iRbody, xmlDOM
Dim AssignedCNT, objXML, retCode, iMessage
Dim ord_no     : ord_no     = request("ord_no")
Dim ord_dtl_sn : ord_dtl_sn = request("ord_dtl_sn")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("inv_no"), 15)					'// 15자 넘으면 에러
Dim beasongNum : beasongNum = request("songjangDiv")

'' 기사용송장 오류 인경우 기타배송 전송
if (request("isfrcsend")="1") then
	hdc_cd = "00099" '' 택배사 기타로 바꿈.
end if

actCnt = 0			'실갱신건수
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no
Dim retDoc, sURL
Dim successYn, errorMsg, sendDt
'/////////////////////////////////////
sendDt = CStr(Replace(Date(), "-", ""))&Num2Str(hour(now()),2,"0","R")&Num2Str(minute(now()),2,"0","R")


'response.write "테스트중"&"<br>"
'response.write ord_no&"<br>"
'response.write ord_dtl_sn&"<br>"
'response.write APISSLURL & "/ordservices/reqdelivery/" & sendDt & "/01/" & hdc_cd & "/" & inv_no & "/" & beasongNum& "/Y/"& ord_no& "/" & ord_dtl_sn
'response.end

'https://api.11st.co.kr/rest/ordservices/reqdelivery/[sendDt]/[dlvMthdCd]/[dlvEtprsCd]/[invcNo]/[dlvNo]/[partDlvYn]/[ordNo]/[ordPrdSeq] : 부분발송처리
'https://api.11st.co.kr/rest/ordservices/reqdelivery/[sendDt]/[dlvMthdCd]/[dlvEtprsCd]/[invcNo]/[dlvNo] : 일반발송처리

'	https://api.11st.co.kr/rest/ordservices/reqdelivery/[sendDt]/[dlvMthdCd]/[dlvEtprsCd]/[invcNo]/[dlvNo]
Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	if (request("dlvMthdCd")<>"") then  ''퀵서비스(04).  직접(화물배달)(03).
		objXML.open "GET", "" & APISSLURL & "/ordservices/reqdelivery/" & sendDt & "/"&request("dlvMthdCd")&"/" & hdc_cd & "/" & inv_no & "/" & beasongNum& "/Y/"& ord_no& "/" & ord_dtl_sn
	else
		objXML.open "GET", "" & APISSLURL & "/ordservices/reqdelivery/" & sendDt & "/01/" & hdc_cd & "/" & inv_no & "/" & beasongNum& "/Y/"& ord_no& "/" & ord_dtl_sn
	end if
	objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
	objXML.setRequestHeader "openapikey",""&APIkey&""
	objXML.send()
	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
			xmlDOM.LoadXML iRbody
'response.write replace(iRbody, "xml","aaaa")

			If session("ssBctID")="kjy8517" Then
				response.write "<textarea cols=100 rows=30>"&iRbody&"</textarea>"
			End If

			On Error Resume Next
			retCode		= Trim(xmlDOM.getElementsByTagName("result_code").item(0).text)
			iMessage	= Trim(xmlDOM.getElementsByTagName("result_text").item(0).text)
			On Error Goto 0
		Set xmlDOM = nothing
	End If
Set objXML = nothing

'////////////////////////////////////
'rw successYn  (true, false)
'rw errorMsg
'rw successYn
'rw errorMsg
Dim IsSuccss : IsSuccss=(retCode="0")

if (IsSuccss) then
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendState=1"
	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
    strSql = strSql & "	and beasongNum11st='"&beasongNum&"'"&VBCRLF
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
    strSql = strSql & "	and beasongNum11st='"&beasongNum&"'"&VBCRLF
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
		if (iMessage="다른 판매자가 사용한 송장번호입니다. 송장번호 재확인 후 판매자센터로 연락주세요.") or (LEFT(iMessage,LEN("이미 사용된 송장번호 입니다."))="이미 사용된 송장번호 입니다.") then
			Dim reqURI : reqURI="?ord_no="&request("ord_no")&"&ord_dtl_sn="&request("ord_dtl_sn")&"&hdc_cd="&request("hdc_cd")&"&inv_no="&request("inv_no")&"&songjangDiv="&request("songjangDiv")&"&isfrcsend=1"
            rw "<br><input type='button' value='기타배송 전송' onClick=""location.href='"&reqURI&"'"">"
		end if

		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
						"	<option value=''>선택</option>" &_
						"	<option value='951'>기전송 내역</option>" &_
						"	<option value='952'>취소주문</option>" &_
						"</select>&nbsp;&nbsp;"
		response.write "<input type='button' value='완료처리' onClick=""fnSetSendState('"&ORG_ord_no&"','"&beasongNum&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
		response.write "<script language='javascript'>"&VbCRLF
		response.write "function fnSetSendState(ORG_ord_no,beasongNum,ord_dtl_sn,selectValue){"&VbCRLF
		response.write "    if(selectValue == ''){"&VbCRLF
		response.write "    	alert('선택해주세요');"&VbCRLF
		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
		response.write "    	return;"&VbCRLF
		response.write "    }"&VbCRLF
		response.write "    var uri = '11st_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&beasongNum='+beasongNum+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If

end if
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->