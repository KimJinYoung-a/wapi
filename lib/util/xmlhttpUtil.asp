<%
Function SendReq(call_url, sedata)
    dim objHttp, ret_txt, status
    Set objHttp = CreateObject("Msxml2.ServerXMLHTTP")

    on error resume next
    objHttp.Open "POST", call_url, False
    objHttp.setRequestHeader "Connection", "close"
    objHttp.setRequestHeader "Content-Length", Len(sedata)
    objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHttp.setTimeouts 5000,90000,90000,90000
    objHttp.Send  sedata
    '지정한 경로의 서버상태값을 가지고 옵니다.
    status = objHttp.status
	'오류가 있거나 (오류가 없을경우 err.number가 0 값을 돌림) status 값이 200 (로딩 성공) 이 아닐경우
	if err.number <> 0 or status <> 200 then
	      if status = 404 then
	            ret_txt = "[404]존재하지 않는 페이지 입니다."
	      elseif status >= 401 and status < 402 then
	            ret_txt = "[401]접근이 금지된 페이지 입니다."
	      elseif status >= 500 and status <= 600 then
	            ret_txt = "[500]내부 서버 오류 입니다."
	      else
	            ret_txt = "[err]서버가 다운되었거나 올바른 경로가 아닙니다."
	      end if
	'오류가 없음 (문서를 성공적으로 로딩함)
	else
	      ret_txt = objHttp.ResponseBody
	end if
	on Error Goto 0
    set objHttp = Nothing
    SendReq = Trim(BinToText(ret_txt,8192))
end function


Function SendReqGet(call_url, sedata)
    dim objHttp, ret_txt, status
    Set objHttp = CreateObject("Msxml2.ServerXMLHTTP")

    on error resume next
    objHttp.Open "GET", call_url & "?" & sedata , False
'    objHttp.setRequestHeader "Connection", "close"
'    objHttp.setRequestHeader "Content-Length", Len(sedata)
'    objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHttp.setTimeouts 5000,90000,90000,90000
    objHttp.Send 

    '지정한 경로의 서버상태값을 가지고 옵니다.
    status = objHttp.status

	'오류가 있거나 (오류가 없을경우 err.number가 0 값을 돌림) status 값이 200 (로딩 성공) 이 아닐경우
	if err.number <> 0 or status <> 200 then
	      if status = 404 then
	            ret_txt = "[404]존재하지 않는 페이지 입니다."
	      elseif status >= 401 and status < 402 then
	            ret_txt = "[401]접근이 금지된 페이지 입니다."
	      elseif status >= 500 and status <= 600 then
	            ret_txt = "[500]내부 서버 오류 입니다."
	      else
	            ret_txt = "[err]서버가 다운되었거나 올바른 경로가 아닙니다."
	      end if
	'오류가 없음 (문서를 성공적으로 로딩함)
	else
	      ret_txt = objHttp.ResponseBody
	end if
	on Error Goto 0
    set objHttp = Nothing
    
    
    SendReqGet = Trim(BinToText(ret_txt,8192))
end function

Function BinToText(varBinData, intDataSizeBytes)
	Const adFldLong = &H00000080
	Const adVarChar = 200

	dim objRS, strV, tmpMsg,isError

	Set objRS = CreateObject("ADODB.Recordset")
	objRS.Fields.Append "txt", adVarChar, intDataSizeBytes, adFldLong
	objRS.Open
	objRS.AddNew
	objRS.Fields("txt").AppendChunk varBinData
	strV=objRS("txt").Value
	BinToText = strV
	objRS.Close
	Set objRS=Nothing
End Function

Function StripTags(htmlDoc)
	Dim rex
	Set rex = new Regexp
	rex.Pattern= "<[^>]+>"
	rex.Global=True
	StripTags =rex.Replace(htmlDoc,"")
	Set rex = Nothing
End Function
%>
