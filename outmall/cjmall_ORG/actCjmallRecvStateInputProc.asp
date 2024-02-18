<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/outmall/cjmall/incCJmallFunction.asp"-->
<!-- #include virtual="/outmall/incOutMallRecvCheckFunction.asp"-->

<%

function getCjMallRecvXMLStr(mallordernoAll, delicompCd, wbNo, vendorOrdId, recvNm)
    Dim strRst
    Dim mallorderno,mallOrderSeq1,mallOrderSeq2,mallOrderSeq3

    mallorderno     = splitValue(mallordernoAll,"-",0)
    mallOrderSeq1   = splitValue(mallordernoAll,"-",1)
    mallOrderSeq2   = splitValue(mallordernoAll,"-",2)
    mallOrderSeq3   = splitValue(mallordernoAll,"-",3)

    strRst = ""
    strRst = strRst &"<?xml version=""1.0"" encoding=""EUC-KR""?>"
    strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_04_05"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_04_05.xsd"">"
    strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
    strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
	strRst = strRst &"<tns:receiveComplete>"
    strRst = strRst &"<tns:ordNo>"&mallorderno&"</tns:ordNo>" ''주문정보 - 주문번호
    strRst = strRst &"<tns:ordGSeq>"&mallOrderSeq1&"</tns:ordGSeq>" ''주문정보 - 주문상품순번
    strRst = strRst &"<tns:ordDSeq>"&mallOrderSeq2&"</tns:ordDSeq>" ''주문정보 - 주문상세순번
    strRst = strRst &"<tns:ordWSeq>"&mallOrderSeq3&"</tns:ordWSeq>" ''주문정보 - 주문처리순번
    strRst = strRst &"<tns:recvNm>"&recvNm&"</tns:recvNm>" ''받는사람
    strRst = strRst &"</tns:receiveComplete>"
    strRst = strRst &"</tns:ifRequest>"

    getCjMallRecvXMLStr = strRst
end function

dim mode      : mode=request("mode")
If mode = "updateSendState" Then
	Dim sqlStr, AssignedRow
	sqlStr = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr & "	Set recvSendState='"&request("updateSendState")&"'"
	sqlStr = sqlStr & "	,recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1"
	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ord_no")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	sqlStr = sqlStr & "	and sellsite='cjmall'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If

	'// 변수선언
	dim LotteGoodNo, LotteStatCd
	dim strSql, actCnt, lp
    dim AssignedCNT

	actCnt = 0			'실갱신건수

Dim proc_gubun : proc_gubun = "sfin"

Dim ten_ord_no : ten_ord_no     = request("ten_ord_no")
Dim ord_no     : ord_no     = request("ord_no")
Dim ord_dtl_sn : ord_dtl_sn = request("ord_dtl_sn")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = request("inv_no")
Dim rcv_nm     : rcv_nm     = request("rcv_nm")
Dim paramAdd
Dim tenOrderSerial,orgsubtotalprice ,minusOrderSerial ,minusSubtotalprice
dim recvCheckResult, IsAutoCheckAvail, IsTooManyFail
dim errCode, errMSG

inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")

Dim ORG_ord_no : ORG_ord_no = ord_no

IsAutoCheckAvail = True
recvCheckResult = True

if (hdc_cd="30") or (hdc_cd = "10") then
    rw hdc_cd & " 이노지스택배사, 아주택배 자동조회 불가<br>"
    ''dbget.Close() : response.end
	IsAutoCheckAvail = False
else
	if (IsAutoScript) and (hdc_cd = "21") then
		'// 경동택배 자동전송시 임시로 SKIP
	    rw "SKIP|"&ord_no&" "&ord_dtl_sn
		dbget.Close() : response.end
	else
    	recvCheckResult = FnCheckNSaveRecvState(hdc_cd, inv_no, errCode, errMSG)
    end if
end if



if (Not recvCheckResult) or (Not IsAutoCheckAvail) then
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1"
    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
    strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
    strSql = strSql & "	and matchstate in ('O','C','Q','A')"

	dbget.Execute strSql

	IsTooManyFail = False

	strSql = ""
	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
	strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
	strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
	strSql = strSql & " and recvSendReqCnt >= 3" & VBCRLF
	rsget.Open strSql,dbget,1
	If Not rsget.Eof Then
		IsTooManyFail = (rsget("cnt") > 0)
	End If
	rsget.Close

	if (IsTooManyFail) or (Not IsAutoCheckAvail) then

		response.write "배송조회 : <a href='" + CStr(FnGetRecvCheckURL(hdc_cd, inv_no)) + "'>" + CStr(FnGetRecvCheckURL(hdc_cd, inv_no)) + "</a><br>"
		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
		"	<option value=''>선택</option>" &_
		"	<option value='100'>기전송 내역</option>" &_
		"</select>&nbsp;&nbsp;"
		response.write "<input type='button' value='완료처리' onClick=""fnSetSendState('"&ORG_ord_no&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
		response.write "<script language='javascript'>"&VbCRLF
		response.write "function fnSetSendState(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
		response.write "    if(selectValue == ''){"&VbCRLF
		response.write "    	alert('선택해주세요');"&VbCRLF
		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
		response.write "    	return;"&VbCRLF
		response.write "    }"&VbCRLF
		response.write "    var uri = 'actCjmallRecvStateInputProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF

		dbget.close()
		response.end
	else
		response.write "<script>alert('배송완료 이전입니다.'); location.href='" + CStr(FnGetRecvCheckURL(hdc_cd, inv_no)) + "';</script>"
		dbget.close()
		response.end
	end if


end if

'' response.write "TEST" & hdc_cd
'' response.end



Dim xmlStr : xmlStr = getCjMallRecvXMLStr(ord_dtl_sn, hdc_cd, inv_no, ten_ord_no, rcv_nm)


Dim retDoc, sURL
Dim successYn, errorMsg
sURL = cjMallAPIURL

SET retDoc = xmlSend(sURL, xmlStr)

'On Error Resume next
successYn = retDoc.getElementsByTagName("ns1:successYn").item(0).text
errorMsg = retDoc.getElementsByTagName("ns1:errorMsg").item(0).text
'On Error Goto 0
SET retDoc = Nothing

'rw successYn  (true, false)
'rw errorMsg
rw successYn
rw errorMsg
Dim IsSuccss : IsSuccss=(successYn="true")

if (IsSuccss) then
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set recvSendState = 100"
	strSql = strSql & "	, recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1 "
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
	if (errorMsg = "고객인수가 완료된 상태입니다.") then
		strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
		strSql = strSql & "	Set recvSendState = 100"
		strSql = strSql & "	, recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1 "
		strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
		strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
		strSql = strSql & "	and matchstate in ('O','A')"
	else
		strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
		strSql = strSql & "	Set recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1 "
		strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
		strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
		strSql = strSql & "	and matchstate in ('O','C','Q','A')"
	end if

	dbget.Execute strSql

    rw "<font color=red>"&errorMsg&"</font>"

    rw ten_ord_no
    rw ord_no
    rw ord_dtl_sn
    rw hdc_cd
    rw inv_no
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
