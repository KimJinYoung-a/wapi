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
		'// ����ֹ��� �μ����۵� skip
		sqlStr = sqlStr & " , recvSendState = 100 "
		sqlStr = sqlStr & " , recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1 "
	end if

	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ord_no")&"'"
	sqlStr = sqlStr & "	and beasongNum11st='"&request("beasongNum")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	sqlStr = sqlStr & "	and sellsite='11st1010'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
	response.end
End If

'###############################################################################################################################################################
Dim strSql, actCnt, iRbody, xmlDOM
Dim AssignedCNT, objXML, retCode, iMessage
Dim ord_no     : ord_no     = request("ord_no")
Dim ord_dtl_sn : ord_dtl_sn = request("ord_dtl_sn")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("inv_no"), 15)					'// 15�� ������ ����
Dim beasongNum : beasongNum = request("songjangDiv")

'' ������� ���� �ΰ�� ��Ÿ��� ����
if (request("isfrcsend")="1") then
	hdc_cd = "00099" '' �ù�� ��Ÿ�� �ٲ�.
end if

actCnt = 0			'�ǰ��ŰǼ�
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no
Dim retDoc, sURL
Dim successYn, errorMsg, sendDt
'/////////////////////////////////////
sendDt = CStr(Replace(Date(), "-", ""))&Num2Str(hour(now()),2,"0","R")&Num2Str(minute(now()),2,"0","R")


'response.write "�׽�Ʈ��"&"<br>"
'response.write ord_no&"<br>"
'response.write ord_dtl_sn&"<br>"
'response.write APISSLURL & "/ordservices/reqdelivery/" & sendDt & "/01/" & hdc_cd & "/" & inv_no & "/" & beasongNum& "/Y/"& ord_no& "/" & ord_dtl_sn
'response.end

'https://api.11st.co.kr/rest/ordservices/reqdelivery/[sendDt]/[dlvMthdCd]/[dlvEtprsCd]/[invcNo]/[dlvNo]/[partDlvYn]/[ordNo]/[ordPrdSeq] : �κй߼�ó��
'https://api.11st.co.kr/rest/ordservices/reqdelivery/[sendDt]/[dlvMthdCd]/[dlvEtprsCd]/[invcNo]/[dlvNo] : �Ϲݹ߼�ó��

'	https://api.11st.co.kr/rest/ordservices/reqdelivery/[sendDt]/[dlvMthdCd]/[dlvEtprsCd]/[invcNo]/[dlvNo]
Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	if (request("dlvMthdCd")<>"") then  ''������(04).  ����(ȭ�����)(03).
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

	'���� ����Ƚ���� 3ȸ�� ������ ����ó�� ����
	'updateSendState = 951		������ ����
	'updateSendState = 952		����ֹ�
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
		if (iMessage="�ٸ� �Ǹ��ڰ� ����� �����ȣ�Դϴ�. �����ȣ ��Ȯ�� �� �Ǹ��ڼ��ͷ� �����ּ���.") or (LEFT(iMessage,LEN("�̹� ���� �����ȣ �Դϴ�."))="�̹� ���� �����ȣ �Դϴ�.") then
			Dim reqURI : reqURI="?ord_no="&request("ord_no")&"&ord_dtl_sn="&request("ord_dtl_sn")&"&hdc_cd="&request("hdc_cd")&"&inv_no="&request("inv_no")&"&songjangDiv="&request("songjangDiv")&"&isfrcsend=1"
            rw "<br><input type='button' value='��Ÿ��� ����' onClick=""location.href='"&reqURI&"'"">"
		end if

		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
						"	<option value=''>����</option>" &_
						"	<option value='951'>������ ����</option>" &_
						"	<option value='952'>����ֹ�</option>" &_
						"</select>&nbsp;&nbsp;"
		response.write "<input type='button' value='�Ϸ�ó��' onClick=""fnSetSendState('"&ORG_ord_no&"','"&beasongNum&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
		response.write "<script language='javascript'>"&VbCRLF
		response.write "function fnSetSendState(ORG_ord_no,beasongNum,ord_dtl_sn,selectValue){"&VbCRLF
		response.write "    if(selectValue == ''){"&VbCRLF
		response.write "    	alert('�������ּ���');"&VbCRLF
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