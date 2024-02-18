<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/ltimall/inc_dailyAuthCheck.asp"-->
<% if NOT (IsAutoScript) then %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body bgcolor="#F4F4F4" >
<% end if %>
<%
public function GetxSiteDateFormat(dt)
	GetxSiteDateFormat = Replace(dt, "-", "")
end function

Dim strParam, sURL
Dim ord_no, ord_dtl_sn, sendQnt, sendDate, outmallGoodsID, hdc_cd, inv_no
Dim buf, objMasterListXML, strSql, ErrMsg, objMasterOneXML, sendOK

ord_no			= requestCheckVar(html2db(request("ord_no")),32)
ord_dtl_sn		= requestCheckVar(html2db(request("ord_dtl_sn")),32)
sendQnt			= requestCheckVar(html2db(request("sendQnt")),32)
sendDate		= requestCheckVar(html2db(request("sendDate")),32)
outmallGoodsID	= requestCheckVar(html2db(request("outmallGoodsID")),32)
hdc_cd			= requestCheckVar(html2db(request("hdc_cd")),32)
inv_no			= requestCheckVar(html2db(request("inv_no")),32)

If (hdc_cd="99") and Len(replace(inv_no,"-",""))>15 then inv_no=Left(replace(inv_no,"-",""),15)
If (inv_no="11시배송완료") then inv_no="11시배송"
If (inv_no="핸드폰으로전송예정:)") then inv_no="기타"

If instr(inv_no,"-") > 0 Then			'2015-07-10 김진영 추가
	inv_no = replace(inv_no, "-", "")
End If

sURL = ltiMallAPIURL & "/openapi/registDeliver.lotte?subscriptionId=" & CStr(ltiMallAuthNo) & "&ord_no=" & CStr(ord_no) & "&ord_dtl_sn=" & CStr(ord_dtl_sn) & "&proc_gubun=sfin&hdc_cd=" & CStr(hdc_cd) & "&inv_no=" & CStr(inv_no) & "&dlv_fin_dtime=" & CStr(GetxSiteDateFormat(sendDate))
Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", sURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()

	If objXML.Status = "200" Then
		buf = BinaryToText(objXML.ResponseBody, "euc-kr")
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML replace(buf,"&","＆")
			Set objMasterListXML = xmlDOM.selectNodes("/Response/Result")
			If (objMasterListXML.length > 0) then
				sendOK = "Y"
				set objMasterListXML = nothing
			Else
				sendOK = "N"
				set objMasterListXML = nothing

				set objMasterListXML = xmlDOM.selectNodes("/Response/Errors")
				set objMasterOneXML = objMasterListXML.item(0)
					ErrMsg = objMasterOneXML.selectSingleNode("Error/Message").text
				set objMasterListXML = nothing
				set objMasterOneXML = nothing
			End If
		Set xmlDOM = Nothing
	End If
Set objXML = nothing

if sendOK = "Y" then
	'// 성공
	strSql = " update db_temp.dbo.tbl_xSite_TMPOrder"
	strSql = strSql & " set sendstate=1"
	strSql = strSql & " ,sendreqCnt=IsNULL(sendreqCnt,0)+1"
	strSql = strSql & " where outmallorderserial='"&ord_no&"'"
	strSql = strSql & " and orgdetailkey='"&ord_dtl_sn&"'"
	strSql = strSql & " and IsNULL(sendstate,0)=0"
	strSql = strSql & " and IsNULL(matchstate,'') <> 'D' and ordercsgbn = 0"
	dbget.Execute strSql
    if (IsAutoScript) then
        rw "OK|"&ord_no&" "&ord_dtl_sn
    ELSE
	    response.write "OK"
	ENd IF
else
    if (IsAutoScript) then
        rw "ErrMsg="&ErrMsg&":"&ord_no&" "&ord_dtl_sn
    else
        rw "ErrMsg="&ErrMsg
    ENd IF

	strSql = " update db_temp.dbo.tbl_xSite_TMPOrder"
	strSql = strSql & " set sendreqCnt=IsNULL(sendreqCnt,0)+1"
	strSql = strSql & " where outmallorderserial='"&ord_no&"'"
	strSql = strSql & " and orgdetailkey='"&ord_dtl_sn&"'"
	strSql = strSql & " and IsNULL(sendstate,0)=0"
	strSql = strSql & " and IsNULL(matchstate,'') <> 'D' and IsNULL(ordercsgbn, 0) = 0"
	''response.write strSql
	dbget.Execute strSql

	'// 에러 3회 이상이면 수기처리
	Dim errCount
	strSql = ""
	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
	strSql = strSql & "	where OutMallOrderSerial='"&ord_no&"'"
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
						"	<option value='901'>발송처리누락 수기등록건</option>" &_
						"	<option value='902'>취소후 제결제건</option>" &_
						"	<option value='903'>반품처리건</option>" &_
						"</select>&nbsp;&nbsp;"
		response.write "<input type='button' value='완료처리' onClick=""finCancelOrd2('"&ord_no&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)""><br>"
		response.write "<script language='javascript'>"&VbCRLF
		response.write "function finCancelOrd2(ord_no,ord_dtl_sn,selectValue){"&VbCRLF
		response.write "    if(selectValue == ''){"&VbCRLF
		response.write "    	alert('선택해주세요');"&VbCRLF
		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
		response.write "    	return;"&VbCRLF
		response.write "    }"&VbCRLF
		response.write "    var uri = 'LtimallProc.asp?mode=updateSendState&ord_no='+ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    var popwin = window.open(uri,'finCancelOrd2','width=200,height=200');"&VbCRLF
		response.write "    popwin.focus()"&VbCRLF
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