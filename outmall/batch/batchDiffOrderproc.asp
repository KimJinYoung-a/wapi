<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim mallid      : mallid = requestCheckVar(request("mallid"),32)
Dim apiAction   : apiAction = requestCheckVar(request("apiAction"),32)
Dim topN        : topN=requestCheckVar(request("topN"),10)
Dim retVal, paramData, iParams
Dim rowsVal, i, oneRow, colsVal
Dim iIdx, iMallid, iItemid
Dim arrItemidIdx

paramData = "mallid="&mallid
arrItemidIdx = ""

If (application("Svr_Info") = "Dev") Then
	retVal = SendReqOutmall("http://localhost:11117/outmall/proc/diffOrderJavaRead.asp",paramData)
Else
	retVal = SendReqOutmall("http://wapi.10x10.co.kr/outmall/proc/diffOrderJavaRead.asp",paramData)
End If

If (IsSuccess(retVal)) Then
	rowsVal = Split(retVal, VBCRLF)
	For i = Lbound(rowsVal) To Ubound(rowsVal)
		oneRow = rowsVal(i)
		colsVal = Split(oneRow, "||")
		If Ubound(colsVal) >= 2 Then
			iIdx	= colsVal(0)
			iMallid	= colsVal(1)
			iItemid	= colsVal(2)
			arrItemidIdx = arrItemidIdx & iItemid & "##" & iIdx & ","
		End If
	Next

	If Right(Trim(arrItemidIdx),1) = "," Then
		arrItemidIdx = Left(arrItemidIdx, Len(arrItemidIdx) - 1)
	End If
End If
response.write arrItemidIdx
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->