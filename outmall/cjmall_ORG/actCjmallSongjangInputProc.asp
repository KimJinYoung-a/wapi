<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/outmall/cjmall/incCJmallFunction.asp"-->
<%
function getCjMallSongjangXMLStr(mallordernoAll,delicompCd,wbNo,vendorOrdId)
    Dim strRst
    Dim mallorderno,mallOrderSeq1,mallOrderSeq2,mallOrderSeq3

    mallorderno     = splitValue(mallordernoAll,"-",0)
    mallOrderSeq1   = splitValue(mallordernoAll,"-",1)
    mallOrderSeq2   = splitValue(mallordernoAll,"-",2)
    mallOrderSeq3   = splitValue(mallordernoAll,"-",3)

    strRst = ""
    strRst = strRst &"<?xml version=""1.0"" encoding=""EUC-KR""?>"
    strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_04_04"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_04_04.xsd"">"
    strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
    strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
    strRst = strRst &"<tns:takeout>"
    strRst = strRst &"<tns:ordNo>"&mallorderno&"</tns:ordNo>" ''주문정보 - 주문번호
    strRst = strRst &"<tns:ordGSeq>"&mallOrderSeq1&"</tns:ordGSeq>" ''주문정보 - 주문상품순번
    strRst = strRst &"<tns:ordDSeq>"&mallOrderSeq2&"</tns:ordDSeq>" ''주문정보 - 주문상세순번
    strRst = strRst &"<tns:ordWSeq>"&mallOrderSeq3&"</tns:ordWSeq>" ''주문정보 - 주문처리순번
    strRst = strRst &"<tns:delicompCd>"&delicompCd&"</tns:delicompCd>" ''택배사
    strRst = strRst &"<tns:wbNo>"&wbNo&"</tns:wbNo>" ''운송장번호
    strRst = strRst &"<tns:vendorOrdId>"&vendorOrdId&"</tns:vendorOrdId>" ''협력사주문번호
    strRst = strRst &"</tns:takeout>"
    strRst = strRst &"</tns:ifRequest>"

    getCjMallSongjangXMLStr = strRst
end function

dim mode      : mode=request("mode")
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
Dim inv_no     : inv_no     = Left(request("inv_no"), 15)					'// 15자 넘으면 에러
Dim paramAdd
Dim tenOrderSerial,orgsubtotalprice ,minusOrderSerial ,minusSubtotalprice

inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")

Dim ORG_ord_no : ORG_ord_no = ord_no

'ord_no = replace(ord_no,":01","")  '''2011-11-30-4100282:01
'ord_no = replace(ord_no,"_1","")  ''2012-03-11-8159343_1
'ord_no = replace(ord_no,"_2","")
'ord_no = replace(ord_no,"_3","")
'ord_no = replace(ord_no,"_4","")

'if (inv_no="보코통 순면 오가닉 원형패드") then inv_no="기타"
'paramAdd = "&ord_no="+Replace(ord_no,"-","")
'paramAdd = paramAdd + "&ord_dtl_sn="+ord_dtl_sn
'paramAdd = paramAdd + "&proc_gubun="+proc_gubun
'paramAdd = paramAdd + "&hdc_cd="+hdc_cd
'paramAdd = paramAdd + "&inv_no="+server.UrlEncode(replace(inv_no,"-",""))

'ord_no = "20130523069391"
'ord_dtl_sn = "20130523069391-001-001-001"

'20130523069250-001-001-001
'20130523069304-001-001-001
'20130523069391-001-001-001

Dim xmlStr : xmlStr = getCjMallSongjangXMLStr(ord_dtl_sn,hdc_cd,inv_no,ten_ord_no)

'' response.write xmlStr
'' response.end

Dim retDoc, sURL
Dim successYn, errorMsg
sURL = cjMallAPIURL

SET retDoc = xmlSend(sURL, xmlStr)

'    If (isCJ_DebugMode) Then
'        CALL XMLFileSave(retDoc.XML, "RET_SONGJANG", ord_dtl_sn)
'    End If

'On Error Resume next
successYn = retDoc.getElementsByTagName("ns1:successYn").item(0).text
errorMsg = retDoc.getElementsByTagName("ns1:errorMsg").item(0).text
'On Error Goto 0
SET retDoc = Nothing

'rw successYn  (true, false)
'rw errorMsg
'rw successYn
'rw errorMsg
Dim IsSuccss : IsSuccss=(successYn="true")

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

    rw "<font color=red>"&errorMsg&"</font>"

    rw ten_ord_no
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
		response.write "    var uri = 'actCjmallSongjangInputProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If

end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
