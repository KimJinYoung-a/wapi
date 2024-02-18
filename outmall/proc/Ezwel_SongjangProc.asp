<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/ezwel/ezwelItemcls.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body bgcolor="#F4F4F4" >

<%
Function getEzwelSongjangXMLStr(masterno, detailno, delicompCd, wbNo)
	Dim strRst
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
	strRst = strRst & "	<dataSet>"
	strRst = strRst & "		<arrSheetNoInfo>"
	strRst = strRst & "			<orderNum>"&masterno&"</orderNum>"
	strRst = strRst & "			<orderGoodsNum>"&detailno&"</orderGoodsNum>"
	strRst = strRst & "			<dlvrCd>"&delicompCd&"</dlvrCd>"
	strRst = strRst & "			<sheetNo>"&wbNo&"</sheetNo>"
	strRst = strRst & "		</arrSheetNoInfo>"
	strRst = strRst & "	</dataSet>"
    getEzwelSongjangXMLStr = strRst
End function

''수취완료전송 : 기타택배/업체직배송.
function setEzwelDlvFinish(orderNum, orderGoodsNum,icspCd,icrtCd)
	dim xmlURL, postParam
	dim strRst
	dim objXML, xmlDOM

	setEzwelDlvFinish = False

	xmlURL = "http://api.ezwel.com/if/api/orderStatusInfoAPI.ez"
	postParam = "cspCd=" & icspCd & "&crtCd=" & icrtCd & "&dataSet="

	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
	strRst = strRst & "<dataSet>"
	strRst = strRst & "       <arrOrderStatusInfo>"
	strRst = strRst & "              <orderNum>" & orderNum & "</orderNum>"
	strRst = strRst & "              <orderGoodsNum>" & orderGoodsNum & "</orderGoodsNum>"
	strRst = strRst & "              <orderStatus>1004</orderStatus>"
	strRst = strRst & "              <orderMemo>forceFin</orderMemo>"
	strRst = strRst & "       </arrOrderStatusInfo>"
	strRst = strRst & "</dataSet>"
	''response.write strRst
	''dbget.close : response.end

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL ''& "?" & postParam
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
	objXML.send(postParam & strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
	''response.write objXML.responseText & "<br /><br />"

	If xmlDOM.getElementsByTagName("resultSet/resultCode").item(0).text <> "200" Then
		response.write "주문상태(수취완료) 전송오류 : " & xmlDOM.getElementsByTagName("resultSet/resultMsg").item(0).text & "<br />"
		exit function
	ELSE
		response.write xmlDOM.getElementsByTagName("resultSet/resultMsg").item(0).text
	end if

	setEzwelDlvFinish = True

end function



Dim strSql, actCnt, xmlDOM
Dim AssignedCNT, objXML, retCode, iMessage
Dim ord_no     : ord_no     = request("ord_no")
Dim ord_dtl_sn : ord_dtl_sn = request("ord_dtl_sn")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("inv_no"), 15)					'// 15자 넘으면 에러
actCnt = 0			'실갱신건수
inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")
Dim ORG_ord_no : ORG_ord_no = ord_no
dim prctp : prctp = requestCheckvar(request("prctp"),20)    ''처리 Action (3:수취완료등록, )

'' 주문을 나눠 입력하는 케이스.
IF (InStr(ord_no,"_")>0) then
	ord_no = getOutmallRefOrgOrderNO(ord_no,ord_dtl_sn,"ezwel")
end if

''2019/11/05 수취완료전송 추가.
IF (prctp="3") then
	Call setEzwelDlvFinish(ord_no, ord_dtl_sn, cspCd, crtCd)
	dbget.Close() 
	response.end
End IF

Dim xmlStr : xmlStr = getEzwelSongjangXMLStr(ord_no, ord_dtl_sn, hdc_cd, inv_no)
Dim retDoc, sURL
Dim successYn, errorMsg
Dim ezwelsongjangURL
ezwelsongjangURL = "http://api.ezwel.com/if/api/sheetNoInfoAPI.ez?cspCd="&cspCd&"&crtCd="&crtCd&"&dataSet="

Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "POST", ezwelsongjangURL & xmlStr, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
	objXML.send()

	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
			retCode		= xmlDOM.getElementsByTagName("resultCode").item(0).text
			iMessage	= xmlDOM.getElementsByTagName("resultMsg").item(0).text
		On Error Goto 0
		Set xmlDOM = nothing
	End If
Set objXML = nothing
'////////////////////////////////////
'rw successYn  (true, false)
'rw errorMsg
'rw successYn
'rw errorMsg
Dim IsSuccss : IsSuccss=(retCode="200")

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

	'만약 에러횟수가 3회가 넘으면 수기처리 가능
	'updateSendState = 951		기전송 내역
	'updateSendState = 952		취소주문
	Dim errCount : errCount = 0
	strSql = ""
	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
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
		Dim reqURI 
		if (InStr(iMessage,"잘못된 송장번호 입니다")>0) then
			reqURI="?ord_no="&request("ord_no")&"&ord_dtl_sn="&request("ord_dtl_sn")&"&hdc_cd=1082&inv_no="&request("inv_no")&"&isfrcsend=1"
        	response.write "<br><input type='button' value='기타배송 전송' onClick=""location.href='"&reqURI&"'""><br>"
		end if
		
		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
						"	<option value=''>선택</option>" &_
						"	<option value='951'>기전송 내역</option>" &_
						"	<option value='952'>취소주문</option>" &_
						"</select>&nbsp;&nbsp;"
		''Ezwel_SongjangProc.asp?ord_no=1028177513&ord_dtl_sn=1&hdc_cd=1007&inv_no=35852324358523247854
		
		response.write "<input type='button' value='완료처리' onClick=""fnSetSendState('"&ORG_ord_no&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
		response.write "<script language='javascript'>"&VbCRLF
		response.write "function fnSetSendState(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
		response.write "    if(selectValue == ''){"&VbCRLF
		response.write "    	alert('선택해주세요');"&VbCRLF
		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
		response.write "    	return;"&VbCRLF
		response.write "    }"&VbCRLF
		response.write "    var uri = 'EzwelProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If

end if
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->