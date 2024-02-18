<%
Function CheckVaildBatchIP(ref)
	'2022-08-11 김진영 | CheckVaildBatchIP = False -> CheckVaildBatchIP = True로 변경 후 하단 전부 주석처리
	CheckVaildBatchIP = True
	'2022-08-11 김진영 | CheckVaildBatchIP = False -> CheckVaildBatchIP = True로 변경 후 하단 전부 주석처리 끝
	' CheckVaildBatchIP = False

	' Dim VaildIP, i
	' If (application("Svr_Info") = "Dev") Then
	' 	VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.70","61.252.133.10","61.252.133.80","110.93.128.114","110.93.128.113","192.168.1.70", "192.168.1.67")
	' Else
	' 	VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.70","61.252.133.10","61.252.133.80","110.93.128.114","110.93.128.113","61.252.133.67","112.218.65.244", "192.168.1.72")
	' End If

	' For i=0 to UBound(VaildIP)
	' 	if (VaildIP(i) = ref) Then
	' 		CheckVaildBatchIP = True
	' 		Exit Function
	' 	End If
	' Next
End Function

Function SendReqOutmall(call_url, sedata)
	Dim objHttp, ret_txt, status
	Set objHttp = CreateObject("Msxml2.ServerXMLHTTP")
		On Error Resume Next
			objHttp.Open "POST", call_url, False
			objHttp.setRequestHeader "Connection", "close"
			objHttp.setRequestHeader "Content-Length", Len(sedata)
			objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objHttp.setTimeouts 5000,90000,90000,90000
			objHttp.Send  sedata
			'지정한 경로의 서버상태값을 가지고 옵니다.
			status = objHttp.status

			'오류가 있거나 (오류가 없을경우 err.number가 0 값을 돌림) status 값이 200 (로딩 성공) 이 아닐경우
			If err.number <> 0 or status <> 200 Then
				If status = 404 Then
					ret_txt = "[404]존재하지 않는 페이지 입니다."
				Elseif status >= 401 and status < 402 Then
					ret_txt = "[401]접근이 금지된 페이지 입니다."
				Elseif status >= 500 and status <= 600 Then
					ret_txt = "[500]내부 서버 오류 입니다."
				Else
					ret_txt = "[err]서버가 다운되었거나 올바른 경로가 아닙니다."
				End If
				'오류가 없음 (문서를 성공적으로 로딩함)
			Else
				ret_txt = objHttp.ResponseBody
			End If
		On Error Goto 0
	Set objHttp = Nothing
	SendReqOutmall = Trim(BinToText(ret_txt,8192))
End Function

Function sendOneApi(iIdx, iMallid, iItemid, iApiAction)
	Dim ret, retSplit
	Dim iUrl
	Dim iParams, retResult, retParam, iErrCode, iErrMsg
	iParams = "itemid="& iItemid & "&mallid="& iMallid & "&action="& iApiAction & "&redSsnKey=system"
	Select Case iMallid
		Case "ezwel"		iUrl = "http://wapi.10x10.co.kr/outmall/proc/EzwelProc.asp"
		Case "interpark"	iUrl = "http://wapi.10x10.co.kr/outmall/proc/InterparkProc.asp"
	End Select

	If (iUrl <> "") AND (iParams <> "") Then
		iErrCode	= ""
		iErrMsg		= ""
		retParam	= ""

		ret = SendReqOutmall(iUrl, iParams)

		retSplit = Split(ret, "||")
		If Ubound(retSplit) >= 2 Then
			iErrCode = retSplit(0)
			iErrMsg	 = Trim(retSplit(2))
		End If
		retParam = "redSsnKey=system&idx="& iIdx & "&itemid="& iItemid & "&ErrCode="& iErrCode & "&ErrMsg=" & iErrMsg

		If (application("Svr_Info") = "Dev") Then
			retResult = SendReqOutmall("http://localhost:11117/outmall/proc/QueResultWrite.asp", retParam)
		Else
			retResult = SendReqOutmall("http://wapi.10x10.co.kr/outmall/proc/QueResultWrite.asp", retParam)
		End If

		If IsJenkinsScript then
		 	If retVal <> "S_OK" then
		 		If jenkinsResponseStatus = "0000" Then
		 			jenkinsResponseStatus = "2000"
		 		End If

		 		If jenkinsResponseText = "" Then
		 			jenkinsResponseText = "ITEMID : "
		 		End If
		 		jenkinsResponseText = jenkinsResponseText & iItemid & ", "
		 	End If
		Else
		 	response.write retResult & VbCRLF
		End If
	End If
	Sleep(0.1)
End Function

Function IsSuccess(iRetVal)
	IsSuccess = Left(iRetVal, Len("S_OK")) = "S_OK"
End Function

Function BinToText(varBinData, intDataSizeBytes)
	Const adFldLong = &H00000080
	Const adVarChar = 200
	Dim objRS, strV, tmpMsg,isError

	Set objRS = CreateObject("ADODB.Recordset")
		objRS.Fields.Append "txt", adVarChar, intDataSizeBytes, adFldLong
		objRS.Open
		objRS.AddNew
		objRS.Fields("txt").AppendChunk varBinData
		strV = objRS("txt").Value
		BinToText = strV
		objRS.Close
	Set objRS=Nothing
End Function

Function Sleep(seconds)
	Dim oShell, cmd
	set oShell = CreateObject("Wscript.Shell")

	cmd = "%COMSPEC% /c timeout " & seconds & " /nobreak"
	oShell.Run cmd, 0, 1
End Function

Dim ref : ref = Request.ServerVariables("REMOTE_ADDR")
If (Not CheckVaildBatchIP(ref)) and (Not CheckJenkinsServerIP(ref)) Then
	response.write ref
	dbget.Close()
	response.end
End If

Dim IsJenkinsScript : IsJenkinsScript = False
Dim jenkinsResponseStatus : jenkinsResponseStatus = "0000"
Dim jenkinsResponseText : jenkinsResponseText = ""
If CheckJenkinsServerIP(ref) Then
	IsJenkinsScript = True
End If
%>