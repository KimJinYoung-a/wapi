<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/lotteon/lotteonItemcls.asp"-->
<!-- #include virtual="/outmall/lotteon/inclotteonFunction.asp"-->
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
		'// ����ֹ��� �μ����۵� skip
		sqlStr = sqlStr & " , recvSendState = 100 "
		sqlStr = sqlStr & " , recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1 "
	end if

	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ord_no")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	sqlStr = sqlStr & "	and sellsite='lotteon'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
	response.end
End If
'###############################################################################################################################################################
Dim strSql, actCnt, iRbody, xmlDOM, strObj
Dim AssignedCNT, objXML, retCode, iMessage
Dim ord_no     : ord_no     = request("OutMallOrderSerial")
Dim ord_dtl_sn : ord_dtl_sn = request("OrgDetailKey")
Dim deliveryCompanyCode     : deliveryCompanyCode     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("songjangNo"), 15)					'// 15�� ������ ����
Dim outmallGoodNo		: outmallGoodNo = request("outmallGoodNo")
Dim outmallOptionCode	: outmallOptionCode = request("outmallOptionCode")
Dim beasongNum			: beasongNum = request("beasongNum")
Dim sendQnt				: sendQnt = request("sendQnt")
Dim objJson

actCnt = 0			'�ǰ��ŰǼ�
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no
Dim retDoc
Dim errorMsg, istrParam
Dim IsSuccss
Dim errCount : errCount = 0
Dim isOk, xmlURL
Dim kjytest				: kjytest = request("kjytest")
'/////////////////////////////////////

SET objJson = jsObject()
	Set objJson("deliveryProgressStateList")= jsArray()
		Set objJson("deliveryProgressStateList")(null) = jsObject()
			objJson("deliveryProgressStateList")(null)("odNo") = ORG_ord_no					'#�ֹ���ȣ : �ֹ����̺��� PK�Ӽ�
			objJson("deliveryProgressStateList")(null)("odSeq") = ord_dtl_sn				'#�ֹ����� : �ֹ������� ���ؼ� ��ǰ���� �ο��Ǵ� �Ӽ��� 1
			objJson("deliveryProgressStateList")(null)("procSeq") = beasongNum				'#ó������ : Default 1 ��ǰ������ ó���������� ������. ���� �Է½� 1 �̰� Ŭ������ �߻��� ��� 1�� ������
			objJson("deliveryProgressStateList")(null)("odPrgsStepCd") = "13"				'#�ֹ�����ܰ� | 11 : �������, 12 : ��ǰ�غ�, 13 : �߼ۿϷ�, 14 : ��ۿϷ�, 15 : ����Ϸ�, 23 : ȸ������, 24 : ȸ������, 25 : ȸ���Ϸ�, 26 : ȸ��Ȯ��
			objJson("deliveryProgressStateList")(null)("dvTrcStatDttm") = FormatDate(now(), "00000000000000")	'#��ۻ��¹߻��Ͻ�
			objJson("deliveryProgressStateList")(null)("invcNbr") = 1						'���尳�� : �ϳ��� ��ǰ�� ���ؼ� ������ �и��� ��� ������� �ǹ�
			objJson("deliveryProgressStateList")(null)("dvCoCd") = deliveryCompanyCode		'��ۻ��ڵ�
			objJson("deliveryProgressStateList")(null)("invcNo") = inv_no					'�����ȣ : Packing������ �ٿ��� ������ȣ �ֹ�����ܰ谡 13:�߼ۿϷ�,25:ȸ���Ϸ��� ��� ���尳�� �ʼ�
			objJson("deliveryProgressStateList")(null)("spdNo") = outmallGoodNo				'#��ǰ��ȣ : �Ե�ON���� �����Ǵ� ��ǰ��ȣ
			objJson("deliveryProgressStateList")(null)("sitmNo") = outmallOptionCode		'#��ǰ��ȣ : �Ե�ON���� �����Ǵ� ��ǰ��ȣ
			objJson("deliveryProgressStateList")(null)("slQty") = sendQnt					'���� : ��ǰ�� ���� �ֹ�����
	istrParam = objJson.jsString
SET objJson = nothing
'response.end

Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "POST", APIURL & "/v1/openapi/delivery/v1/SellerDeliveryProgressStateInform", false
	objXML.setRequestHeader "Authorization", "Bearer " & APIkey
	objXML.setRequestHeader "Accept", "application/json"
	objXML.setRequestHeader "Accept-Language", "ko"
	objXML.setRequestHeader "X-Timezone", "GMT+09:00"
	objXML.setRequestHeader "Content-Type", "application/json"
	objXML.Send(istrParam)

'	If kjytest = "Y" Then
		rw objXML.Status
		rw istrParam
		rw iRbody
'	End If

	If objXML.Status <> "200" Then
		IsSuccss = false
		iMessage = "���� ����"
	Else
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

'		If kjytest = "Y" Then
			rw istrParam
			rw iRbody
'		End If

		Set strObj = JSON.parse(iRbody)
			'If strObj.returnCode <> "0000" Then
			If strObj.data.rsltCd <> "0000" Then
				IsSuccss = false
				iMessage = replaceMsg(strObj.data.rsltMsg)
			Else
				IsSuccss = true
			End If
		Set strObj = nothing
	End If
Set objXML = nothing
'rw iMessage
'response.end
'////////////////////////////////////
if (IsSuccss) then
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendState=1"
	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
    strSql = strSql & "	and beasongNum11st='"&beasongNum&"'"&VBCRLF
	strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
    strSql = strSql & "	and matchstate in ('O')"
	dbget.Execute strSql,AssignedCNT
'rw strSql

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
		strSql = strSql & "	and beasongNum11st='"&beasongNum&"'"&VBCRLF
		strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
		strSql = strSql & "	and matchstate in ('O','C','Q','A')"
		dbget.Execute strSql

		'���� ����Ƚ���� 3ȸ�� ������ ����ó�� ����
		'updateSendState = 951		������ ����
		'updateSendState = 952		����ֹ�
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
			response.write  "<select name='updateSendState' id=""updateSendState"">" &_
							"	<option value=''>����</option>" &_
							"	<option value='951'>������ ����</option>" &_
							"	<option value='952'>����ֹ�</option>" &_
							"</select>&nbsp;&nbsp;"
			response.write "<input type='button' value='�Ϸ�ó��' onClick=""fnSetSendState('"&ORG_ord_no&"','"&beasongNum&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
			response.write "<script language='javascript'>"&VbCRLF
			response.write "function fnSetSendState(ORG_ord_no, beasongNum, ord_dtl_sn, selectValue){"&VbCRLF
			response.write "    if(selectValue == ''){"&VbCRLF
			response.write "    	alert('�������ּ���');"&VbCRLF
			response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
			response.write "    	return;"&VbCRLF
			response.write "    }"&VbCRLF
			response.write "    var uri = 'Lotteon_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&beasongNum='+beasongNum+'&updateSendState='+selectValue;"&VbCRLF
			response.write "    location.replace(uri);"&VbCRLF
			response.write "}"&VbCRLF
			response.write "</script>"&VbCRLF
		End If
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
'    rw hdc_cd
    rw inv_no

	'���� ����Ƚ���� 3ȸ�� ������ ����ó�� ����
	'updateSendState = 951		������ ����
	'updateSendState = 952		����ֹ�
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
		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
						"	<option value=''>����</option>" &_
						"	<option value='951'>������ ����</option>" &_
						"	<option value='952'>����ֹ�</option>" &_
						"</select>&nbsp;&nbsp;"
		response.write "<input type='button' value='�Ϸ�ó��' onClick=""fnSetSendState('"&ORG_ord_no&"','"&beasongNum&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
		response.write "<script language='javascript'>"&VbCRLF
		response.write "function fnSetSendState(ORG_ord_no, beasongNum, ord_dtl_sn, selectValue){"&VbCRLF
		response.write "    if(selectValue == ''){"&VbCRLF
		response.write "    	alert('�������ּ���');"&VbCRLF
		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
		response.write "    	return;"&VbCRLF
		response.write "    }"&VbCRLF
		response.write "    var uri = 'Lotteon_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&beasongNum='+beasongNum+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If
end if
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->