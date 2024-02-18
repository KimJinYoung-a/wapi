<%@  codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<%


Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

'Response.ContentType = "application/json"
'Response.AddHeader "Accept", "application/json"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim sauthkey, sepType, yyyymmdd, hh, snvMids, mode, strParam
Dim apiURL, objXML, xmlDOM, buf
If (application("Svr_Info")="Dev") Then
	apiURL = "http://dev.api.shopping.naver.com"
	sauthkey = "s_dc4b9209dcb"
Else
	apiURL = "http://api.shopping.naver.com"
	sauthkey = "s_dc4b9209dcb"
End If

mode 		= request("mode")
sepType		= request("epType")
yyyymmdd	= request("yyyymmdd")
hh			= request("hh")
snvMids		= request("nvMids")

Select Case mode
	Case "A"			'��ü EP
		If sepType <> "all" Then
			strParam = "?authkey="&sauthkey&"&type=FILE"													'��� �⺻������ ä���� ���ϳ�¥�� ��üEP ����
		Else
			strParam = "/api/lowestPrice/?authkey="&sauthkey&"&type=FILE&epType="&sepType&"&yyyy-mm-dd="&yyyymmdd			'yyyy-mm-dd���� ��üEP ����
		End If
	Case "C"			'��� EP
		If yyyymmdd = "" AND hh = "" Then
			strParam = "/api/lowestPrice/?authkey="&sauthkey&"&type=FILE&epType=summary"									'���ϳ�¥ + ����ð��� ���EP ����
		Else
			strParam = "/api/lowestPrice/?authkey="&sauthkey&"&type=FILE&epType=summary&yyyy-mm-dd="&yyyymmdd&"&hh="&hh		'yyyy-mm-dd���� hh�� ���EP ����
		End If
	Case "R"			'�ǽð�
		strParam = "/api/lowestPrice/model?authkey="&sauthkey&"&nvMids="&snvMids											'snvMids�� ���� �ǽð� ������
End Select

Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", apiURL & strparam, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"
	objXML.Send()
	If objXML.Status = "200" Then
		buf = BinaryToText(objXML.ResponseBody, "utf-8")
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML replace(buf,"&","��")

response.write replace(buf,"&","��")



		Set xmlDOM = nothing
	End If
Set objXML = nothing
%>