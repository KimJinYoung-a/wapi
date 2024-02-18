<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/coupang/coupangItemcls.asp"-->
<!-- #include virtual="/outmall/coupang/incCoupangFunction.asp"-->
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
	If (request("updateSendState") = "953") Then
		Dim jsonObj, resObj, jBody, jStatus, isAlreadyReq, jSql, jAssignedRow
		isAlreadyReq = "N"
		Set jsonObj= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			jsonObj.open "GET", "http://xapi.10x10.co.kr:8080/Orders/Coupang/orderId/"&request("ord_no"), false
			jsonObj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			jsonObj.Send()
			If jsonObj.Status = "200" OR jsonObj.Status = "201" Then
				jBody = BinaryToText(jsonObj.ResponseBody,"utf-8")
				Set resObj = JSON.parse(jBody)
					'response.write jBody & "<br/ >"
					'response.end
					jStatus = resObj.data.get(0).status
					Select Case jStatus
						'DEPARTURE - �������, DELIVERING - �����, FINAL_DELIVERY - ��ۿϷ�, NONE_TRACKING - ��ü���� : ��ü�����߰�
						Case "DEPARTURE", "DELIVERING", "FINAL_DELIVERY", "NONE_TRACKING"
							isAlreadyReq = "Y"
						Case Else
							isAlreadyReq = "N"
					End Select
				Set resObj = nothing
			Else
				isAlreadyReq = "N"
			End If
		Set jsonObj = nothing

		If isAlreadyReq = "Y" Then
			jSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
			jSql = jSql & "	Set sendState='951' "
			jSql = jSql & "	,sendReqCnt=sendReqCnt+1"
			jSql = jSql & "	where OutMallOrderSerial='"&request("ord_no")&"'"
			jSql = jSql & "	and beasongNum11st='"&request("beasongNum")&"'"
			jSql = jSql & "	and outmallOptionNo='"&request("outmallOptionCode")&"'"
			jSql = jSql & "	and sellsite='coupang'"
			dbget.Execute jSql,jAssignedRow
			response.write "<script>alert('"&jAssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
			response.end
		End If
	Else
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
		sqlStr = sqlStr & "	and outmallOptionNo='"&request("outmallOptionCode")&"'"
		sqlStr = sqlStr & "	and sellsite='coupang'"
		dbget.Execute sqlStr,AssignedRow
		response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
		response.end
	End If
End If

'###############################################################################################################################################################
Function fnCoupangOrderCheck(iORG_ord_no, ibeasongNum, ioutmallOptionCode, iisOk)
	Dim jsonObj, resObj, jBody, jStatus, isAlreadyReq, jSql, jAssignedRow
	isAlreadyReq = "N"
	Set jsonObj= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		jsonObj.open "GET", "http://xapi.10x10.co.kr:8080/Orders/Coupang/orderId/"&iORG_ord_no, false
		jsonObj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		jsonObj.Send()
		If jsonObj.Status = "200" OR jsonObj.Status = "201" Then
			jBody = BinaryToText(jsonObj.ResponseBody,"utf-8")
			Set resObj = JSON.parse(jBody)
				'response.write jBody & "<br/ >"
				'response.end
				jStatus = resObj.data.get(0).status
				Select Case jStatus
					'DEPARTURE - �������, DELIVERING - �����, FINAL_DELIVERY - ��ۿϷ�
					Case "DEPARTURE", "DELIVERING", "FINAL_DELIVERY"
						isAlreadyReq = "Y"
					Case Else
						isAlreadyReq = "N"
				End Select
			Set resObj = nothing
		Else
			isAlreadyReq = "N"
		End If
	Set jsonObj = nothing

	If isAlreadyReq = "Y" Then
		jSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
		jSql = jSql & "	Set sendState='951' "
		jSql = jSql & "	,sendReqCnt=sendReqCnt+1"
		jSql = jSql & "	where OutMallOrderSerial='"& iORG_ord_no &"'"
		jSql = jSql & "	and beasongNum11st='"& ibeasongNum &"'"
		jSql = jSql & "	and outmallOptionNo='"& ioutmallOptionCode &"'"
		jSql = jSql & "	and sellsite='coupang'"
		dbget.Execute jSql,jAssignedRow
		response.write "<br />" & jAssignedRow & "�� �Ϸ� ó��.(API���ּ��ܰ���ȸ)"
		response.end
		'response.write "<script>alert('"&jAssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
		'response.end
	Else
		iisOk = "N"
	End If
End Function

Dim strSql, actCnt, iRbody, xmlDOM, strObj
Dim AssignedCNT, objXML, retCode, iMessage
Dim ord_no     : ord_no     = request("OutMallOrderSerial")
Dim ord_dtl_sn : ord_dtl_sn = request("OrgDetailKey")
Dim deliveryCompanyCode     : deliveryCompanyCode     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("songjangNo"), 15)					'// 15�� ������ ����
Dim outmallGoodNo		: outmallGoodNo = request("outmallGoodNo")
Dim outmallOptionCode	: outmallOptionCode = request("outmallOptionCode")
Dim outmallOptionName	: outmallOptionName = request("outmallOptionName")
Dim beasongNum			: beasongNum = request("beasongNum")
Dim splitrequire		: splitrequire = request("splitrequire")
Dim objJson

actCnt = 0			'�ǰ��ŰǼ�
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no
Dim retDoc, sURL
Dim errorMsg, istrParam
Dim IsSuccss
Dim errCount : errCount = 0
Dim isOk

if (mode="directsend") then  ''��ü�������� ����
	deliveryCompanyCode = "DIRECT"
end if
'/////////////////////////////////////
SET objJson = jsObject()
	objJson("shipmentBoxId") = ""&beasongNum&""
	objJson("orderId") = ""&ord_no&""
	objJson("vendorItemId") = ""&outmallOptionCode&""
	objJson("deliveryCompanyCode") = ""&deliveryCompanyCode&""
	objJson("invoiceNumber") = ""&inv_no&""
	objJson("splitShipping") = false
	objJson("preSplitShipped") = false
	objJson("estimatedShippingDate") = ""
	istrParam = objJson.jsString
SET objJson = nothing
'response.end
Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "POST", "http://xapi.10x10.co.kr:8080/Orders/Coupang/invoice", false
	objXML.setRequestHeader "Content-Type", "application/json"
	objXML.Send(istrParam)

	If objXML.Status = "200" OR objXML.Status = "201" Then
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			'response.write iRbody & "<br/ >"
			If strObj.data.responseCode <> "0" Then
				IsSuccss = false
				iMessage = strObj.data.responseList.get(0).resultMessage
			Else
				IsSuccss = true
			End If
		Set strObj = nothing
	Else
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			iMessage = "���� ����"
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
	if (mode="directsend") then ''2019/11/08 �б�
		strSql = strSql & "	,sendSongjangNo = 'DIRECT "&inv_no&"'"
	else
		strSql = strSql & "	,sendSongjangNo = '"&inv_no&"'"		'2018-09-19 17:34 ������ �߰�
	end if
    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
    strSql = strSql & "	and beasongNum11st='"&beasongNum&"'"&VBCRLF
	strSql = strSql & "	and outmallOptionNo='"&outmallOptionCode&"'"&VBCRLF
    strSql = strSql & "	and matchstate in ('O')"
	dbget.Execute strSql,AssignedCNT
'rw strSql
    IF (AssignedCNT>0) then
	    if (IsAutoScript) then
	        rw "OK|"&ord_no&" "&ord_dtl_sn
	    ELSE
    	    response.write "OK"
    	ENd IF

		if (mode="directsend") then
			' sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_OUTAMLL_Jungsan_Comment_add] '" & tenorderserial & "', " & tenitemid & ", '" & tenitemoption & "','�ڻ�������','"&session("ssBctId")&"'"
			' dbDatamart_dbget.Execute sqlStr
		end if
	ELSE
	    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
		strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
	    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
	    strSql = strSql & "	and beasongNum11st='"&beasongNum&"'"&VBCRLF
	    strSql = strSql & "	and outmallOptionNo='"&outmallOptionCode&"'"&VBCRLF
	    strSql = strSql & "	and matchstate in ('O','C','Q','A')"
		dbget.Execute strSql

		'���� ����Ƚ���� 3ȸ�� ������ ����ó�� ����
		'updateSendState = 951		������ ����
		'updateSendState = 952		����ֹ�
		strSql = ""
		strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
		strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
		strSql = strSql & "	and beasongNum11st='"&beasongNum&"'"
		strSql = strSql & "	and outmallOptionNo='"&outmallOptionCode&"'"&VBCRLF
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
							"	<option value='953'>API���ּ��ܰ���ȸ</option>" &_
							"</select>&nbsp;&nbsp;"
			response.write "<input type='button' value='�Ϸ�ó��' onClick=""fnSetSendState('"&ORG_ord_no&"','"&beasongNum&"','"&outmallOptionCode&"',document.getElementById('updateSendState').value)"">"
			response.write "<script language='javascript'>"&VbCRLF
			response.write "function fnSetSendState(ORG_ord_no,beasongNum,outmallOptionCode,selectValue){"&VbCRLF
			response.write "    if(selectValue == ''){"&VbCRLF
			response.write "    	alert('�������ּ���');"&VbCRLF
			response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
			response.write "    	return;"&VbCRLF
			response.write "    }"&VbCRLF
			response.write "    var uri = 'coupang_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&outmallOptionCode='+outmallOptionCode+'&beasongNum='+beasongNum+'&updateSendState='+selectValue;"&VbCRLF
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
    strSql = strSql & "	and outmallOptionNo='"&outmallOptionCode&"'"&VBCRLF
    strSql = strSql & "	and matchstate in ('O','C','Q','A')"

	dbget.Execute strSql

    rw "<font color=red>"&iMessage&"</font>"

    rw ord_no
    rw ord_dtl_sn
    rw deliveryCompanyCode
    rw inv_no

	if InStr(iMessage,"�̹� ����� �����ȣ�� �־�, �����ȣ ����� �Ұ����մϴ�.")>0 then
		rw "<input type='button' value='�ڻ�����������' onClick=""fnSetDirectSend('"&ORG_ord_no&"','"&beasongNum&"','"&outmallOptionCode&"')""> (�������� �ȵǴ°�� ��ü�������� ����, 7���������)"

		response.write "<script language='javascript'>"&VbCRLF
		response.write "function fnSetDirectSend(ORG_ord_no,beasongNum,outmallOptionCode){"&VbCRLF
		response.write "    var uri = 'coupang_SongjangProc.asp?mode=directsend&OutMallOrderSerial="&ORG_ord_no&"&OrgDetailKey="&ord_dtl_sn&"&hdc_cd="&deliveryCompanyCode&"&songjangNo="&inv_no&"&outmallGoodNo="&outmallGoodNo&"&outmallOptionCode="&outmallOptionCode&"&beasongNum="&beasongNum&"&splitrequire="&splitrequire&"';"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF

		rw ""
	end if
	'���� ����Ƚ���� 3ȸ�� ������ ����ó�� ����
	'updateSendState = 951		������ ����
	'updateSendState = 952		����ֹ�
	strSql = ""
	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
	strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
	strSql = strSql & "	and beasongNum11st='"&beasongNum&"'"
	strSql = strSql & "	and outmallOptionNo='"&outmallOptionCode&"'"&VBCRLF
	strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not rsget.Eof Then
		errCount = rsget("cnt")
	End If
	rsget.Close

	If errCount > 0 Then
		'Call ��ȸ �ؼ� dim xx�� y�� ��, N�̸� �ϴ� response.write ȣ��
		Call fnCoupangOrderCheck(ORG_ord_no, beasongNum, outmallOptionCode, isOk)
		If isOk = "N" Then
			response.write  "<select name='updateSendState' id=""updateSendState"">" &_
							"	<option value=''>����</option>" &_
							"	<option value='951'>������ ����</option>" &_
							"	<option value='952'>����ֹ�</option>" &_
							"	<option value='953'>API���ּ��ܰ���ȸ</option>" &_
							"</select>&nbsp;&nbsp;"
			response.write "<input type='button' value='�Ϸ�ó��' onClick=""fnSetSendState('"&ORG_ord_no&"','"&beasongNum&"','"&outmallOptionCode&"',document.getElementById('updateSendState').value)"">"
			response.write "<script language='javascript'>"&VbCRLF
			response.write "function fnSetSendState(ORG_ord_no,beasongNum,outmallOptionCode,selectValue){"&VbCRLF
			response.write "    if(selectValue == ''){"&VbCRLF
			response.write "    	alert('�������ּ���');"&VbCRLF
			response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
			response.write "    	return;"&VbCRLF
			response.write "    }"&VbCRLF
			response.write "    var uri = 'coupang_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&outmallOptionCode='+outmallOptionCode+'&beasongNum='+beasongNum+'&updateSendState='+selectValue;"&VbCRLF
			response.write "    location.replace(uri);"&VbCRLF
			response.write "}"&VbCRLF
			response.write "</script>"&VbCRLF
		End If
	End If
end if
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->