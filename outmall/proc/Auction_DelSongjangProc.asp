<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/auction/auctionItemcls.asp"-->
<!-- #include virtual="/outmall/auction/incAuctionFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Function getAuctionSongjangXMLStr(masterno, detailno, delicompCd, wbNo, isDiv)
'delicompCd : �ֹ���ȣ
'wbNo		: ����
'delicompCd	: �ù��ڵ�
'rw isDiv
'response.end

	Dim strRst, sSql, dName
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst & "	<soap:Header>"
	strRst = strRst & "		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
	strRst = strRst &"			<Value>"&auctionTicket&"</Value>"
	strRst = strRst & "		</EncryptedTicket>"
	strRst = strRst & "	</soap:Header>"
	strRst = strRst & "	<soap:Body>"
	strRst = strRst & "		<DoShippingGeneral xmlns=""http://www.auction.co.kr/APIv1/AuctionService"">"
	strRst = strRst & "			<req OrderNo="""&detailno&""">"
	strRst = strRst & "				<RemittanceMethod RemittanceMethodType=""Emoney"" RemittanceAccountName="""" RemittanceAccountNumber="""" RemittanceBankCode="""" xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"" />"
If delicompCd <> "etc" Then
	strRst = strRst & "				<ShippingMethod SendDate="""&Date()&""" InvoiceNo="""&wbNo&""" MessageForBuyer="""" ShippingMethodClassficationType=""Door2Door"" DeliveryAgency="""&delicompCd&""" DeliveryAgencyName="""" ShippingEtcMethod=""Nothing"" ShippingEtcAgencyName="""" xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"" />"
Else
	sSql = ""
	sSql = sSql & " select top 1 divname from db_order.[dbo].[tbl_songjang_div] where divcd = '"&isDiv&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open sSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not rsget.Eof Then
		dName = rsget("divname")
	End If
	rsget.Close

	strRst = strRst & "				<ShippingMethod SendDate="""&Date()&""" InvoiceNo="""&wbNo&""" MessageForBuyer="""" ShippingMethodClassficationType=""Door2Door"" DeliveryAgency="""&delicompCd&""" DeliveryAgencyName="""&dName&""" xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"" />"
End If
	strRst = strRst & "			</req>"
	strRst = strRst & "		</DoShippingGeneral>"
	strRst = strRst & "	</soap:Body>"
	strRst = strRst & "</soap:Envelope>"
	getAuctionSongjangXMLStr = strRst
End function

Dim mode : mode=request("mode")
If mode = "updateSendState" Then
	Dim sqlStr, AssignedRow
	sqlStr = "Update db_temp.dbo.tbl_xsite_mayDelOrder "
	sqlStr = sqlStr & "	Set sendState='"&request("updateSendState")&"'"
	sqlStr = sqlStr & "	,sendReqCnt=sendReqCnt+1"

	if (request("updateSendState") = "952") then
		'// ����ֹ��� �μ����۵� skip
		sqlStr = sqlStr & " , recvSendState = 100 "
		sqlStr = sqlStr & " , recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1 "
	end if

	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ord_no")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	sqlStr = sqlStr & "	and sellsite='auction1010'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
	response.end
End If

'###############################################################################################################################################################
Dim strSql, actCnt
Dim AssignedCNT, objXML, retCode, iMessage, xmlDOM
Dim ord_no     : ord_no     = request("ord_no")
Dim ord_dtl_sn : ord_dtl_sn = request("ord_dtl_sn")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("inv_no"), 15)					'// 15�� ������ ����
Dim s_Div : s_Div = request("songjangDiv")
actCnt = 0			'�ǰ��ŰǼ�
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no
Dim xmlStr : xmlStr = getAuctionSongjangXMLStr(ord_no, ord_dtl_sn, hdc_cd, inv_no, s_Div)
Dim retDoc, sURL
Dim successYn, errorMsg
'/////////////////////////////////////
Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "POST", "" & auctionAPIURL&"/APIv1/AuctionService.asmx"
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(xmlStr)
	objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/AuctionService/DoShippingGeneral"
	objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(objXML.responseText,"soap:",""))
'			response.write (Replace(objXML.responseText,"soap:",""))
'			response.end
			On Error Resume Next
			If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "DoShippingGeneralResponse" Then
				retCode		= xmlDOM.getElementsByTagName ("DoShippingGeneralResult ").item(0).attributes(0).nodeValue	'���ǻ�ǰ�ڵ�
				'iMessage	= xmlDOM.getElementsByTagName ("DoShippingGeneralResult ").item(0).attributes(1).nodeValue	'���ǿɼ�Ÿ��
			End If
			On Error Goto 0
		Set xmlDOM = nothing
	End If
Set objXML = nothing
'////////////////////////////////////

Dim IsSuccss : IsSuccss=(retCode="Success")
response.write "IsSuccss:"&IsSuccss&"<BR>"

if (IsSuccss) then
    strSql = "Update db_temp.dbo.tbl_xsite_mayDelOrder "
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
    strSql = "Update db_temp.dbo.tbl_xsite_mayDelOrder "
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
	strSql = ""
	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xsite_mayDelOrder " & VBCRLF
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
		response.write "    var uri = 'Auction_SongjangProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If

end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->