<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim mode : mode=request("mode")
Dim sellsite : sellsite = request("sellsite")
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
	sqlStr = sqlStr & "	and sellsite='"& sellsite &"'"
	dbget.Execute sqlStr,AssignedRow
    ''response.write sqlStr
	response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
	response.end
End If

''response.write "aaaa"
''dbget.close : response.end

Public Function getSabangNetDeliveryXMLStr(iord_no, ihdc_cd, iinv_no)
	Dim strRst

    strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    strRst = strRst & "<SABANG_INV_REGI>"
 	strRst = strRst & "	<HEADER>"
	strRst = strRst & "		<SEND_COMPAYNY_ID>"&sabangnetID&"</SEND_COMPAYNY_ID>"				'#���� �α��� ���̵�
	strRst = strRst & "		<SEND_AUTH_KEY>"&sabangnetAPIKEY&"</SEND_AUTH_KEY>"					'#���ݿ��� �߱� ���� ����Ű
	strRst = strRst & "		<SEND_DATE>"&Replace(Date(), "-", "")&"</SEND_DATE>"				'#�������� | YYYYMMDD
 	strRst = strRst & "		<SEND_INV_EDIT_YN>N</SEND_INV_EDIT_YN>"
 	strRst = strRst & "	</HEADER>"
 	strRst = strRst & "	<DATA>"
 	strRst = strRst & "		<SABANGNET_IDX><![CDATA[" & iord_no & "]]></SABANGNET_IDX>"
 	strRst = strRst & "		<TAK_CODE><![CDATA[" & ihdc_cd & "]]></TAK_CODE>"
 	strRst = strRst & "		<TAK_INVOICE><![CDATA[" & iinv_no & "]]></TAK_INVOICE>"
 	strRst = strRst & "	</DATA>"
 	strRst = strRst & "</SABANG_INV_REGI>"

	getSabangNetDeliveryXMLStr = strRst
End Function

'###############################################################################################################################################################
Dim strSql, actCnt, iRbody, xmlDOM
Dim AssignedCNT, objXML, retCode, iMessage
Dim ord_no     : ord_no     = request("OutMallOrderSerial")
Dim ord_dtl_sn : ord_dtl_sn = request("OrgDetailKey")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("songjangNo"), 15)					'// 15�� ������ ����
Dim outmallGoodNo		: outmallGoodNo = request("outmallGoodNo")
Dim outmallOptionCode	: outmallOptionCode = request("outmallOptionCode")
Dim outmallOptionName	: outmallOptionName = request("outmallOptionName")
Dim itemno				: itemno = request("itemno")
Dim shoplinkerorderid	: shoplinkerorderid = request("shoplinkerorderid")

actCnt = 0			'�ǰ��ŰǼ�
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no
Dim retDoc, sURL
Dim successYn, errorMsg
'/////////////////////////////////////
Dim xmlStr : xmlStr = getSabangNetDeliveryXMLStr(shoplinkerorderid, hdc_cd, inv_no)
''response.write "shoplinkerorderid : " & shoplinkerorderid
''response.write xmlStr
'response.end

dim xmlURL : xmlURL = sabangnetAPIURL
xmlURL = xmlURL + "/RTL_API/xml_order_invoice.html"
''https://r.sabangnet.co.kr/RTL_API/xml_order_invoice.html?xml_url=����xml�ּ�
''response.write xmlURL

Dim opath : opath = "/outmall/sabangnet/sabangnetXML/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
Dim defaultPath : defaultPath = server.mappath(opath) + "\"
CALL CheckFolderCreate(defaultPath)
Dim fileName

fileName = "SendSongjang" &"_"& getCurrDateTimeFormat&".xml"

dim fso, tFile
Set fso = CreateObject("Scripting.FileSystemObject")
    Set tFile = fso.CreateTextFile(defaultPath & FileName )
        tFile.WriteLine xmlStr
    Set tFile = nothing
Set fso = nothing

dim dataURL : dataURL = "?xml_url="&sabangnetWapiURL&opath&FileName

Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
    objXML.open "POST", "" & xmlURL & dataURL, false
    objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objXML.Send(xmlStr)

    Call DelAPITMPFile(sabangnetWapiURL&opath&FileName)

	If objXML.Status = "200" Then
        iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
        if (InStr(iRbody, "���� : " & shoplinkerorderid)) then
            retCode = "0000"
        else
            iMessage = iRbody
        end if
	End If
Set objXML = nothing
'////////////////////////////////////
'rw successYn  (true, false)
'rw errorMsg
'rw successYn
'rw errorMsg
Dim IsSuccss : IsSuccss=(retCode="0000")

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

	'���� ����Ƚ���� 3ȸ�� ������ ����ó�� ����
	'updateSendState = 951		������ ����
	'updateSendState = 952		����ֹ�
	Dim errCount : errCount = 0
	Dim isellsite
	strSql = ""
	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
	strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
	strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
	strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not rsget.Eof Then
		errCount = rsget("cnt")
	End If
	rsget.Close

	strSql = ""
	strSql = strSql & " SELECT TOP 1 sellsite FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
	strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
	strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
	strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not rsget.Eof Then
		isellsite = rsget("sellsite")
	End If
	rsget.Close


	If errCount > 0 Then
		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
						"	<option value=''>����</option>" &_
						"	<option value='951'>������ ����</option>" &_
						"	<option value='952'>����ֹ�</option>" &_
						"</select>&nbsp;&nbsp;"
		response.write "<input type='button' value='�Ϸ�ó��' onClick=""fnSetSendState('"&ORG_ord_no&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
		response.write "<script language='javascript'>"&VbCRLF
		response.write "function fnSetSendState(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
		response.write "    if(selectValue == ''){"&VbCRLF
		response.write "    	alert('�������ּ���');"&VbCRLF
		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
		response.write "    	return;"&VbCRLF
		response.write "    }"&VbCRLF
		response.write "    var uri = 'sabangnet_SongjangProc.asp?mode=updateSendState&sellsite='"& isellsite &"'&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If

end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->