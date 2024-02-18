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
	Case "A"			'전체 EP
		If sepType <> "all" Then
			strParam = "?authkey="&sauthkey&"&type=FILE"													'모두 기본값으로 채워져 당일날짜의 전체EP 전송
		Else
			strParam = "/api/lowestPrice/?authkey="&sauthkey&"&type=FILE&epType="&sepType&"&yyyy-mm-dd="&yyyymmdd			'yyyy-mm-dd일자 전체EP 전송
		End If
	Case "C"			'요약 EP
		If yyyymmdd = "" AND hh = "" Then
			strParam = "/api/lowestPrice/?authkey="&sauthkey&"&type=FILE&epType=summary"									'당일날짜 + 현재시간의 요약EP 전송
		Else
			strParam = "/api/lowestPrice/?authkey="&sauthkey&"&type=FILE&epType=summary&yyyy-mm-dd="&yyyymmdd&"&hh="&hh		'yyyy-mm-dd일자 hh시 요약EP 전송
		End If
	Case "R"			'실시간
		strParam = "/api/lowestPrice/model?authkey="&sauthkey&"&nvMids="&snvMids											'snvMids에 대한 실시간 최저가
End Select

Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", apiURL & strparam, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"
	objXML.Send()
	If objXML.Status = "200" Then
		buf = BinaryToText(objXML.ResponseBody, "utf-8")
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML replace(buf,"&","＆")

response.write replace(buf,"&","＆")



		Set xmlDOM = nothing
	End If
Set objXML = nothing
%>