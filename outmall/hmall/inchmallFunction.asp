<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'상품 등록
Public Function fnHmallItemReg(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccess
	istrParam = "itemid="&iitemid
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "http://xapi.10x10.co.kr:8080/Product/Hmall", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상품등록] " & html2db(Err.Description)
			Exit Function
		End If
		'rw objXML.Status
		'rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			'response.write iRbody
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success

				If isSuccess = true Then
					strSql = " EXEC db_etcmall.[dbo].[usp_API_Hmall_RegItemInfo_Upd] '"&iitemid&"', 'I' "
					dbget.execute strSql
					iErrStr = "OK||"&iitemid&"||성공[상품등록]"
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품등록] 통신은 성공"
					If (session("ssBctID")="kjy8517") Then
						response.write BinaryToText(objXML.ResponseBody,"utf-8")
					End If
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[상품등록] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품등록_NO] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'상품만 등록
Public Function fnHmallOnlyItemReg(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccess
	istrParam = "itemid="&iitemid
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "http://xapi.10x10.co.kr:8080/Product/Hmall/singlereg", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상품등록add] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success

				If isSuccess = true Then
					strSql = " EXEC db_etcmall.[dbo].[usp_API_Hmall_RegItemInfo_Upd] '"&iitemid&"', 'I' "
					dbget.execute strSql
					iErrStr = "OK||"&iitemid&"||성공[상품등록add]"
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품등록add] 통신은 성공"
					If (session("ssBctID")="kjy8517") Then
						response.write BinaryToText(objXML.ResponseBody,"utf-8")
					End If
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[상품등록add] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품등록add_NO] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'이미지 등록
Public Function fnHmallImage(iitemid, ichgImageNm, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccess
	istrParam = "itemid="&iitemid
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "PUT", "http://xapi.10x10.co.kr:8080/Product/Hmall/imagereg", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[이미지등록] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success

				If isSuccess = true Then
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regitem "
					strSql = strSql & " SET APIaddImg = 'Y' "
					If (ichgImageNm <> "") Then
						strSql = strSql & " ,regimageName='"&ichgImageNm&"'"& VbCrlf
					End If
					strSql = strSql & " WHERE itemid = '"& iitemid &"' "
					dbget.execute strSql
					iErrStr = "OK||"&iitemid&"||성공[이미지등록]"
				Else
					iErrStr = "ERR||"&iitemid&"||실패[이미지등록] 통신은 성공"
					If (session("ssBctID")="kjy8517") Then
						response.write BinaryToText(objXML.ResponseBody,"utf-8")
					End If
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[이미지등록] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[이미지등록_NO] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'이미지 확인
Public Function fnHmallImageConfirm(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccess
	istrParam = "itemid="&iitemid
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "PUT", "http://xapi.10x10.co.kr:8080/Product/Hmall/imageconfirm", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[이미지확인] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success

				If isSuccess = true Then
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regitem "
					strSql = strSql & " SET APIconfirmImg = 'Y' "
					strSql = strSql & " WHERE itemid = '"& iitemid &"' "
					dbget.execute strSql
					iErrStr = "OK||"&iitemid&"||성공[이미지확인]"
				Else
					iErrStr = "ERR||"&iitemid&"||실패[이미지확인] 통신은 성공"
					If (session("ssBctID")="kjy8517") Then
						response.write BinaryToText(objXML.ResponseBody,"utf-8")
					End If
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[이미지확인] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[이미지확인_NO] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'상품 정보 수정
Public Function fnHmallOnlyItemEdit(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccess
	istrParam = "itemid="&iitemid
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "PUT", "http://xapi.10x10.co.kr:8080/Product/Hmall/singleupd", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[정보수정] " & html2db(Err.Description)
			Exit Function
		End If
		'rw objXML.Status
		'rw BinaryToText(objXML.ResponseBody,"utf-8")
		'response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success

				If isSuccess = true Then
					strSql = ""
					strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Hmall_RegItemInfo_Upd] '"&iitemid&"', 'R' "
					dbget.Execute(strSql)
					iErrStr = "OK||"&iitemid&"||성공[정보수정]"
				Else
					iErrStr = "ERR||"&iitemid&"||실패[정보수정] 통신은 성공"
					If (session("ssBctID")="kjy8517") Then
						response.write BinaryToText(objXML.ResponseBody,"utf-8")
					End If
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[정보수정] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[정보수정] "& html2db(replace(iRbody, """", ""))
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'상품 등록(신규등록)
Public Function fnHmallItemOnlyReg(iitemid, istrParam, iErrStr, igetMustprice, ihmallSellYn, iLimityn, iLimitNo, iLimitSold, iItemName, ibasicimageNm)
    Dim objXML, strSql, i, iRbody, iMessage, strObj, isSuccess
	Dim xmlDOM, retCode, detailMessage, slitmCd, sitmCd
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "https://openapi.hmall.com//front/pd/pdb/multiItem.do", false
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "oauserId", "002569"
		objXML.setRequestHeader "oauseKey", "23439A336B4FC812A1ED415489F185A2"
		objXML.Send(strParam)
		If (session("ssBctID")="kjy8517") Then
			response.write "<textarea cols=40 rows=10>"&strParam&"</textarea>"
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = objXML.ResponseText
				xmlDOM.LoadXML iRbody

				If (session("ssBctID")="kjy8517") Then
					response.write "<textarea cols=40 rows=10>"&iRbody&"</textarea>"
				End If

				retCode  = xmlDOM.getElementsByTagName("code").item(0).text
				If retCode = "0000" Then
					slitmCd = xmlDOM.getElementsByTagName("slitmCd").item(0).text
					sitmCd = xmlDOM.getElementsByTagName("sitmCd").item(0).text

					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET R.hmallRegdate = getdate()" & VbCrlf
					If (slitmCd <> "") Then
					    strSql = strSql & "	, R.hmallStatCd = '3'"& VbCRLF					'승인대기
					Else
						strSql = strSql & "	, R.hmallStatCd = '1'"& VbCRLF					'전송시도
					End If
					strSql = strSql & " ,R.hmallGoodNo = '" & slitmCd & "'" & VbCrlf
					strSql = strSql & " ,R.hmallGoodNo2 = '" & sitmCd & "'" & VbCrlf
					strSql = strSql & " ,R.hmallLastUpdate = getdate()"
					strSql = strSql & " ,R.hmallsellyn = 'Y' "
					strSql = strSql & " ,R.hmallPrice = '"&igetMustprice&"' " & VbCrlf
					strSql = strSql & " ,R.accFailCNT = 0" & VbCrlf                 ''실패회수 초기화
					strSql = strSql & " ,R.regitemname = i.itemname " & VbCRLF
					strSql = strSql & " ,R.regimageName = '"&ibasicimageNm&"'"& VbCrlf
					strSql = strSql & " ,R.updateCoupon = 'Y' "& VbCrlf
					strSql = strSql & " FROM db_etcmall.dbo.tbl_hmall_regitem R" & VbCrlf
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " WHERE R.itemid = " & iitemid
					dbget.Execute(strSql)
					iErrStr = "OK||"&iitemid&"||성공[상품등록add]"
				Else
					detailMessage = xmlDOM.getElementsByTagName("detail").item(0).text
					iErrStr = "ERR||"&iitemid&"||실패[상품등록add] " & detailMessage
					If (session("ssBctID")="kjy8517") Then
						response.write iRbody
					End If
				End If
			Set xmlDOM = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
				xmlDOM.LoadXML iRbody
				iMessage  = xmlDOM.getElementsByTagName("message").item(0).text
				If iMessage <> "Success" Then
					iErrStr = "ERR||"&iitemid&"||실패[상품등록add] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품등록add] "& html2db(replace(iRbody, """", ""))
				End If
			Set xmlDOM = nothing
		End If
	Set objXML= nothing
End Function

'상품 정보 수정
Public Function fnHmallItemOnlyEdit(iitemid, strParam, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccess
	Dim xmlDOM, retCode, detailMessage
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "https://openapi.hmall.com//front/pd/pdb/updateItemByVen.do", false
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "oauserId", "002569"
		objXML.setRequestHeader "oauseKey", "23439A336B4FC812A1ED415489F185A2"
		objXML.Send(strParam)
		If (session("ssBctID")="kjy8517")  AND iitemid = "678905" Then
			response.write "req : <textarea cols=40 rows=10>"&strParam&"</textarea>"
			'response.end
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = objXML.ResponseText
				xmlDOM.LoadXML iRbody
				On Error Resume Next
					retCode  = xmlDOM.getElementsByTagName("code").item(0).text
					If Err.number <> 0 Then
						retCode = ""
					End If
				On Error Goto 0

				If retCode = "0000" Then
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regitem "
					strSql = strSql & " SET updateCoupon = 'Y' "
					strSql = strSql & " WHERE itemid = '"& iitemid &"' "
					dbget.Execute(strSql)

					strSql = ""
					strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Hmall_RegItemInfo_Upd] '"&iitemid&"', 'R' "
					dbget.Execute(strSql)
					iErrStr = "OK||"&iitemid&"||성공[정보수정]"
				ElseIf retCode = "" Then
					iErrStr = "ERR||"&iitemid&"||실패[정보수정] Hmall 통신 안 됨."
					If (session("ssBctID")="kjy8517") Then
						response.write iRbody
					End If
				Else
					detailMessage = xmlDOM.getElementsByTagName("detail").item(0).text
					iErrStr = "ERR||"&iitemid&"||실패[정보수정] " & detailMessage
					If (session("ssBctID")="kjy8517") Then
						response.write iRbody
					End If
				End If
			Set xmlDOM = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
				xmlDOM.LoadXML iRbody


		If (session("ssBctID")="kjy8517") Then
			response.write "<textarea cols=40 rows=10>"&iRbody&"</textarea>"
			response.end
		End If

				iMessage  = xmlDOM.getElementsByTagName("message").item(0).text
				If iMessage <> "Success" Then
					iErrStr = "ERR||"&iitemid&"||실패[정보수정] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[정보수정] "& html2db(replace(iRbody, """", ""))
				End If
			Set xmlDOM = nothing
		End If
	Set objXML= nothing
End Function

'상품 수정
Public Function fnHmallItemEdit(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccess
	istrParam = "itemid="&iitemid
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "PUT", "http://xapi.10x10.co.kr:8080/Product/Hmall", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상품수정] " & html2db(Err.Description)
			Exit Function
		End If
		'rw objXML.Status
		'rw BinaryToText(objXML.ResponseBody,"utf-8")
		'response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success

				If isSuccess = true Then
					strSql = ""
					strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Hmall_RegItemInfo_Upd] '"&iitemid&"', 'R' "
					dbget.Execute(strSql)
					iErrStr = "OK||"&iitemid&"||성공[상품수정]"
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품수정] 통신은 성공"
					If (session("ssBctID")="kjy8517") Then
						response.write BinaryToText(objXML.ResponseBody,"utf-8")
					End If
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[상품수정] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품수정] "& html2db(replace(iRbody, """", ""))
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

' '상품 가격 수정
' Public Function fnHmallPrice(iitemid, imustPrice, imrgnRate, byRef iErrStr)
'     Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode, hmallStatcd
' 	istrParam = ""
' 	istrParam = istrParam & "{"
' 	istrParam = istrParam & "  ""itemid"": """&iitemid&""","
' 	istrParam = istrParam & "  ""mrgnRate"": "&imrgnRate&","
' 	istrParam = istrParam & "  ""sellPrc"": """&imustPrice&""" "
' 	istrParam = istrParam & "}"

' 	On Error Resume Next
' 	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
' 		objXML.open "PUT", "http://xapi.10x10.co.kr:8080/Product/Hmall/price", false
' 		objXML.setRequestHeader "Content-Type", "application/json"
' 		objXML.Send(istrParam)

' 		If Err.number <> 0 Then
' 			iErrStr = "ERR||"&iitemid&"||실패[가격수정] " & html2db(Err.Description)
' 			Exit Function
' 		End If
' '		rw objXML.Status
' '		rw BinaryToText(objXML.ResponseBody,"utf-8")
' '		response.end
' 		If objXML.Status = "200" OR objXML.Status = "201" Then
' 			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
' '			response.write iRbody
' '			response.end
' 			Set strObj = JSON.parse(iRbody)
' 				isSuccessCode		= strObj.outValue.code
' 				iMessage			= strObj.outValue.message

' 				If isSuccessCode = "0000" Then
' 					strSql = ""
' 					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regitem " & VbCRLF
' 					strSql = strSql & "	SET hmallLastUpdate = getdate() " & VbCRLF
' 					strSql = strSql & "	,hmallPrice = " & imustPrice & VbCRLF
' 					strSql = strSql & "	,setMargin = " & imrgnRate & VbCRLF
' 					strSql = strSql & "	,accFailCnt = 0"& VbCRLF
' 					strSql = strSql & " WHERE itemid='" & iitemid & "'"
' 					dbget.Execute(strSql)
' 					iErrStr =  "OK||"&iitemid&"||성공[가격수정]"
' 				Else
' 					iErrStr = "ERR||"&iitemid&"||실패[가격수정] "& html2db(iMessage)
' 				End If
' 			Set strObj = nothing
' 		Else
' 			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
' 			Set strObj = JSON.parse(iRbody)
' 				iMessage			= strObj.message
' 				'rw iRbody
' 				If Instr(iMessage, "동일한 가격이 존재") > 0 Then
' 					strSql = ""
' 					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regitem " & VbCRLF
' 					strSql = strSql & "	SET hmallLastUpdate = getdate() " & VbCRLF
' 					strSql = strSql & "	,hmallPrice = " & imustPrice & VbCRLF
' 					strSql = strSql & "	,setMargin = " & imrgnRate & VbCRLF
' 					strSql = strSql & "	,accFailCnt = 0"& VbCRLF
' 					strSql = strSql & " WHERE itemid='" & iitemid & "'"
' 					dbget.Execute(strSql)
' 					iErrStr =  "OK||"&iitemid&"||성공[가격수정]"
' 				Else
' 					If iMessage <> "" Then
' 						iErrStr = "ERR||"&iitemid&"||실패[가격수정] "& html2db(iMessage)
' 					Else
' 						iErrStr = "ERR||"&iitemid&"||실패[가격수정] 통신오류"
' 					End If
' 				End If
' 			Set strObj = nothing
' 		End If
' 	Set objXML= nothing
' End Function

'상품 가격 수정
Public Function fnHmallPrice(iitemid, imustPrice, imrgnRate, istrParam, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, strObj, isSuccess
	Dim xmlDOM, retCode, detailMessage, slitmCd, sitmCd

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "https://openapi.hmall.com//front/pd/pdh/updateItemPrcHist.do", false
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "oauserId", "002569"
		objXML.setRequestHeader "oauseKey", "23439A336B4FC812A1ED415489F185A2"
		objXML.Send(istrParam)
		If (session("ssBctID")="kjy8517") Then
			response.write "req : <textarea cols=40 rows=10>"&istrParam&"</textarea>"
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = objXML.ResponseText
				xmlDOM.LoadXML iRbody

				If (session("ssBctID")="kjy8517") Then
					response.write "res : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
				End If

				retCode  = xmlDOM.getElementsByTagName("code").item(0).text
				If retCode = "0000" Then
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regitem " & VbCRLF
					strSql = strSql & "	SET hmallLastUpdate = getdate() " & VbCRLF
					strSql = strSql & "	,hmallPrice = " & imustPrice & VbCRLF
					strSql = strSql & "	,setMargin = " & imrgnRate & VbCRLF
					strSql = strSql & "	,accFailCnt = 0"& VbCRLF
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||성공[가격수정]"
				Else
					detailMessage = xmlDOM.getElementsByTagName("detail").item(0).text
					If Instr(detailMessage, "동일한 가격이 존재") > 0 Then
						strSql = ""
						strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regitem " & VbCRLF
						strSql = strSql & "	SET hmallLastUpdate = getdate() " & VbCRLF
						strSql = strSql & "	,hmallPrice = " & imustPrice & VbCRLF
						strSql = strSql & "	,setMargin = " & imrgnRate & VbCRLF
						strSql = strSql & "	,accFailCnt = 0"& VbCRLF
						strSql = strSql & " WHERE itemid='" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||성공[가격수정]"
					Else
						iErrStr = "ERR||"&iitemid&"||실패[가격수정] " & detailMessage
					End If
				End If
			Set xmlDOM = nothing
		Else
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
				xmlDOM.LoadXML iRbody

				If (session("ssBctID")="kjy8517") Then
					response.write "<textarea cols=40 rows=10>"&iRbody&"</textarea>"
				End If

				iMessage  = xmlDOM.getElementsByTagName("message").item(0).text
				If iMessage <> "Success" Then
					iErrStr = "ERR||"&iitemid&"||실패[가격수정] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[가격수정] "& html2db(replace(iRbody, """", ""))
				End If
			Set xmlDOM = nothing
		End If
	Set objXML= nothing
End Function

'상품 상태 수정
Public Function fnHmallSellYN(iitemid, ichgSellYn, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode, itemSellGbcd
	If ichgSellYn = "Y" Then
		itemSellGbcd = "00"		'판매진행
	Else
		itemSellGbcd = "11"		'일시중단
	End If
	istrParam = ""
	istrParam = istrParam & "{"
	istrParam = istrParam & "  ""itemid"": """& iitemid &""", "
	istrParam = istrParam & "  ""itemSellGbcd"": """& itemSellGbcd &"""  "
	istrParam = istrParam & "}"

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "PUT", "http://xapi.10x10.co.kr:8080/Product/Hmall/sale", false
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상태수정] " & html2db(Err.Description)
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
'		response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end
			Set strObj = JSON.parse(iRbody)
				isSuccessCode		= strObj.outValue.code
				iMessage			= strObj.outValue.message
				If isSuccessCode = "0000" Then
					If ichgSellyn = "Y" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET hmallSellYn = 'Y'"
						strSql = strSql & "	,hmallLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_hmall_regitem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||판매[상태수정]"
					ElseIf ichgSellyn = "N" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET hmallSellYn = 'N'"
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,hmallLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_hmall_regitem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||품절처리[상태수정]"
					End If
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상태수정] "& html2db(iMessage)
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[상태수정] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상태수정] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'상품 상세 조회
Public Function fnHmallStatChk(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode, hmallStatcd
	Dim slitmAprvlGbcd, slitmAprvlGbcdNm, itemSellGbcd, ihmallSellYn, ihmallPrice, ostkYn, itemAthzGbcd, itemAthzGbcdNm
	istrParam = "itemid="& iitemid

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://xapi.10x10.co.kr:8080/Product/Hmall?" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[조회] " & html2db(Err.Description)
			Exit Function
		End If
'		rw objXML.Status
'		response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end
			Set strObj = JSON.parse(iRbody)
				isSuccessCode		= strObj.outValue.result.code
				iMessage			= strObj.outValue.result.message
				If isSuccessCode = "0000" Then
					itemAthzGbcd		= strObj.outValue.dsitem.itemAthzGbcd
					itemAthzGbcdNm		= strObj.outValue.dsitem.itemAthzGbcdNm
					slitmAprvlGbcd		= strObj.outValue.dsitem.slitmAprvlGbcd
					slitmAprvlGbcdNm	= strObj.outValue.dsitem.slitmAprvlGbcdNm
					itemSellGbcd		= strObj.outValue.dsitem.itemSellGbcd
					ostkYn				= strObj.outValue.dsitem.ostkYn	'품절여부
					ihmallPrice			= strObj.outValue.slitmPrcAthzHis.sellPrc

					Select Case slitmAprvlGbcd
						Case "0"	hmallStatcd = "3"	'임시저장
						Case "1"	hmallStatcd = "3"	'팀장승인대기
						Case "2"	hmallStatcd = "7"	'승인완료
						Case "3"	hmallStatcd = "4"	'취소
						Case "4"	hmallStatcd = "3"	'시스템처리중
						Case "5"	hmallStatcd = "2"	'반송
						Case "6"	hmallStatcd = "3"	'MD승인대기
						Case "7"	hmallStatcd = "3"	'시스템 처리 대기
						Case "9"	hmallStatcd = "4"	'실패
					End Select

'					Select Case itemAthzGbcd
'						Case "00"	hmallStatcd = "3"	'등록(MD승인대기)
'						Case "11"	hmallStatcd = "3"	'팀장승인대기
'						Case "21"	hmallStatcd = "2"	'보류
'						Case "31"	hmallStatcd = "2"	'반송
'						Case "41"	hmallStatcd = "2"	'취소
'						Case "80"	hmallStatcd = "7"	'승인완료
'						Case "99"	hmallStatcd = "2"	'예외
'					End Select

					Select Case itemSellGbcd
						Case "00"	ihmallSellYn = "Y"	'진행
						Case "11"	ihmallSellYn = "N"	'일시중단
						Case "19"	ihmallSellYn = "X"	'영구중단
					End Select

					If ostkYn = "Y" Then
						ihmallSellYn = "N"
					End If

					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regitem " & VbCRLF
					strSql = strSql & " SET lastConfirmdate = getdate() " & VbCRLF
					strSql = strSql & "	,hmallStatcd = '"&hmallStatcd&"' " & VbCRLF
'					strSql = strSql & "	,hmallSellYn = '"&ihmallSellYn&"' " & VbCRLF
'					strSql = strSql & "	,hmallPrice = '"&ihmallPrice&"' " & VbCRLF
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr = "OK||"&iitemid&"||성공[조회("&slitmAprvlGbcdNm&")]"
				Else
					iErrStr = "ERR||"&iitemid&"||실패[조회] "& html2db(iMessage)
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[조회] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[조회] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'상품 재고 조회
Public Function fnHmallOptionStatCheck(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode, isSingleOption
	Dim uitmCd, uitmTotNm, sellGbcd, ihmallSellYn, currSaleQty, objOption, optTypeName, optionName
	istrParam = "itemid="& iitemid
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://xapi.10x10.co.kr:8080/Product/Hmall/optionqty?" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[재고조회] " & html2db(Err.Description)
			Exit Function
		End If
'		rw objXML.Status
'		response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end
			Set strObj = JSON.parse(iRbody)
				isSuccessCode		= strObj.outValue.result.code
				iMessage			= strObj.outValue.result.message
				If isSuccessCode = "0000" Then
					strSql = ""
					strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_hmall_regedOption] "
					strSql = strSql & " WHERE itemid = " & iitemid
					dbget.Execute strSql

					Set objOption = strObj.outValue.dslist
						For i=0 to objOption.length-1
							uitmCd				= objOption.get(i).uitmCd				'속성코드
							uitmTotNm			= html2db(objOption.get(i).uitmTotNm)	'속성명
							sellGbcd			= objOption.get(i).sellGbcd				'상품판매구분코드 | 00:판매진행, 11:일시중단
							currSaleQty			= objOption.get(i).currSaleQty			'남은수량

							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_hmall_regedOption] (itemid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimitno, lastupdate) VALUES "
							strSql = strSql & " ('"&iitemid&"', '"&uitmCd&"', '"&uitmTotNm&"', '"& CHKIIF(sellGbcd="00", "Y", "N") &"', '"&currSaleQty&"', getdate())"
							dbget.Execute strSql
						Next
					Set objOption = nothing
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_hmall_regedOption] "
					strSql = strSql & " SET itemoption = o.itemoption "
					strSql = strSql & " ,outmallOptName = o.optionname "	'2019-05-08 11:11 김진영 옵션명 수정 추가
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_hmall_regedOption] as Ro "
					strSql = strSql & " 	join db_item.dbo.tbl_item_option o "
					strSql = strSql & " 	on Ro.itemid=o.itemid "
					strSql = strSql & " 	and replace(replace(replace(Ro.outmalloptName,'∼','~'),'&amp;','&'),',','/')=(CASE WHEN o.optionTypename='복합옵션' THEN '' ELSE ltrim(rtrim(o.optionTypename))+'/' END)+(CASE WHEN o.optionTypename='복합옵션' THEN replace(o.optionname,',','/') else o.optionname end) "
					strSql = strSql & " where outmallOptName<>'단일옵션' "
					strSql = strSql & " and ro.itemid = '"& iitemid &"' "
					dbget.Execute strSql

					strSql = ""
					strSql = strSql & " UPDATE R " & VBCRLF
					strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VBCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_hmall_regItem R " & VBCRLF
					strSql = strSql & " Join ( " & VBCRLF
					strSql = strSql & " 	SELECT R.itemid, count(*) as CNT, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt "
					strSql = strSql & " 	FROM db_etcmall.dbo.tbl_hmall_regItem R " & VBCRLF
					strSql = strSql & " 	JOIN db_etcmall.dbo.tbl_hmall_regedOption Ro on R.itemid = Ro.itemid and Ro.itemid = '"&iitemid&"' " & VBCRLF
					strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid " & VBCRLF
					dbget.Execute strSql
					iErrStr = "OK||"&iitemid&"||성공[재고조회]"
				Else
					iErrStr = "ERR||"&iitemid&"||실패[재고조회] "& html2db(iMessage)
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[재고조회] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[재고조회] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'상품 옵션 수정
Public Function fnHmallOptionEdit(iitemid, istrparam, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, strObj, isSuccessCode, hmallStatcd
	If istrparam = "" Then
		iErrStr = "ERR||"&iitemid&"||실패[옵션수정] 옵션가져오기 오류"
	Else
		On Error Resume Next
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.open "PUT", "http://xapi.10x10.co.kr:8080/Product/Hmall/option", false
			objXML.setRequestHeader "Content-Type", "application/json"
			objXML.Send(istrParam)
			If Err.number <> 0 Then
				iErrStr = "ERR||"&iitemid&"||실패[옵션수정] " & html2db(Err.Description)
				Exit Function
			End If
	'		rw objXML.Status
	'		rw BinaryToText(objXML.ResponseBody,"utf-8")
	'		response.end

			If (session("ssBctID")="kjy8517") Then
				response.write istrparam
			End If

			If objXML.Status = "200" OR objXML.Status = "201" Then
				iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
	'			response.write iRbody
	'			response.end
				Set strObj = JSON.parse(iRbody)
					isSuccessCode		= strObj.outValue.code
					iMessage			= strObj.outValue.message
					If isSuccessCode = "0000" Then
						iErrStr =  "OK||"&iitemid&"||성공[옵션수정]"
					Else
						iErrStr = "ERR||"&iitemid&"||실패[옵션수정] "& html2db(iMessage)
					End If
				Set strObj = nothing
			Else
				iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
				Set strObj = JSON.parse(iRbody)
					iMessage			= strObj.message
					'rw iRbody
					If iMessage <> "" Then
						iErrStr = "ERR||"&iitemid&"||실패[옵션수정] "& html2db(iMessage)
					Else
						iErrStr = "ERR||"&iitemid&"||실패[옵션수정] 통신오류"
					End If
				Set strObj = nothing
			End If
		Set objXML= nothing
	End If
End Function

'상품 재고 조회
Public Function fnHmallOptionStatChk(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode, isSingleOption
	Dim uitmCd, uitmTotNm, sellGbcd, ihmallSellYn, currSaleQty, objOption, optTypeName, optionName
	istrParam = "itemid="& iitemid
	isSingleOption = "N"
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://xapi.10x10.co.kr:8080/Product/Hmall/optionqty?" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[재고조회] " & html2db(Err.Description)
			Exit Function
		End If
'		rw objXML.Status
'		response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end
			Set strObj = JSON.parse(iRbody)
				isSuccessCode		= strObj.outValue.result.code
				iMessage			= strObj.outValue.result.message
				If isSuccessCode = "0000" Then
					strSql = ""
					strSql = strSql & " SELECT TOP 1 optionTypeName "
					strSql = strSql & " FROM db_item.dbo.tbl_item_option "
					strSql = strSql & " WHERE itemid = " & iitemid
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If Not(rsget.EOF or rsget.BOF) then
						optTypeName = rsget("optionTypeName")
					End If
					rsget.Close

					Set objOption = strObj.outValue.dslist
						For i=0 to objOption.length-1
							If i = 0 and Ubound(Split(objOption.get(0).uitmTotNm, "/")) > 0 Then
								If optTypeName = Split(objOption.get(0).uitmTotNm, "/")(0) Then
									isSingleOption = "Y"
								Else
									isSingleOption = "N"
								End If
							End If
							uitmCd				= objOption.get(i).uitmCd		'속성코드
							uitmTotNm			= objOption.get(i).uitmTotNm	'속성명
							sellGbcd			= objOption.get(i).sellGbcd		'상품판매구분코드 | 00:판매진행, 11:일시중단
							currSaleQty			= objOption.get(i).currSaleQty	'남은수량

							If uitmTotNm = "단일옵션" Then
								strSql = ""
								strSql = strSql & " IF Exists(SELECT * FROM db_etcmall.dbo.tbl_hmall_regedOption where itemid='"&iitemid&"' and itemoption = '0000' and mallid = 'hmall1010') "
								strSql = strSql & " BEGIN"& VbCRLF
								strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regedOption SET "
								strSql = strSql & " outmalllimitno =  "
								strSql = strSql & " Case WHEN B.limityn = 'Y' and B.limitno - B.limitsold <= 5 THEN '0'  "
								strSql = strSql & " 	 WHEN B.limityn = 'Y' and B.limitno - B.limitsold > 5 THEN B.limitno - B.limitsold - 5 "
								strSql = strSql & " 	 WHEN B.limityn = 'N' THEN '999' END "
								strSql = strSql & " FROM db_etcmall.dbo.tbl_hmall_regedOption A  "
								strSql = strSql & " JOIN db_item.dbo.tbl_item B on A.itemid = B.itemid "
								strSql = strSql & " WHERE A.itemid = '"&iitemid&"' and A.itemoption = '0000' and A.mallid = 'hmall1010' "
								strSql = strSql & " END ELSE "
								strSql = strSql & " BEGIN"& VbCRLF
								strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_hmall_regedOption " & VBCRLF
								strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
								strSql = strSql & " SELECT itemid, '0000', 'hmall1010', '"& uitmCd &"', '"& uitmTotNm &"', sellyn, limityn, " & VBCRLF
								strSql = strSql & " Case WHEN limityn = 'Y' AND limitno - limitsold <= 5 THEN '0' " & VBCRLF
								strSql = strSql & " 	 WHEN limityn = 'Y' AND limitno - limitsold > 5 THEN limitno - limitsold - 5 " & VBCRLF
								strSql = strSql & " 	 WHEN limityn = 'N' THEN '999' End " & VBCRLF
								strSql = strSql & " , '0', getdate() " & VBCRLF
								strSql = strSql & " FROM db_item.dbo.tbl_item " & VBCRLF
								strSql = strSql & " WHERE itemid= '"&iitemid&"' " & VBCRLF
								strSql = strSql & " END "
								dbget.Execute strSql
							Else
								If isSingleOption = "Y" Then
									optionName = Split(objOption.get(i).uitmTotNm, "/")(1)
								Else
									optionName = uitmTotNm
								End If

								strSql = ""
								strSql = strSql & " IF Exists(SELECT * FROM db_etcmall.dbo.tbl_hmall_regedOption where itemid='"&iitemid&"' and outmallOptCode = '"&uitmCd&"' and mallid = 'hmall1010') "
								strSql = strSql & " BEGIN"& VbCRLF
								strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regedOption " & VbCRLF
								strSql = strSql & " SET outmalllimitno = " & VbCRLF
								strSql = strSql & " Case WHEN optlimityn = 'Y' AND optlimitno - optlimitsold <= 5 THEN '0' " & VbCRLF
								strSql = strSql & " 	 WHEN optlimityn = 'Y' AND optlimitno - optlimitsold > 5 THEN optlimitno - optlimitsold - 5" & VbCRLF
								strSql = strSql & " 	 WHEN optlimityn = 'N' THEN '999' End" & VbCRLF
								strSql = strSql & " ,outmalllimityn = B.optlimityn " & VbCRLF
								strSql = strSql & " FROM db_etcmall.dbo.tbl_hmall_regedOption A  " & VbCRLF
								strSql = strSql & " JOIN db_item.dbo.tbl_item_option B on A.itemid = B.itemid and A.itemoption = B.itemoption " & VbCRLF
								strSql = strSql & " WHERE B.itemid = '"&iitemid&"' and replace(B.optionname, ',', '/') = '"&optionName&"' and A.mallid = 'hmall1010' "
								strSql = strSql & " END ELSE "
								strSql = strSql & " BEGIN"& VbCRLF
								strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_hmall_regedOption " & VBCRLF
								strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
								strSql = strSql & " SELECT itemid, itemoption, 'hmall1010', '"&uitmCd&"', optionname, optsellyn, optlimityn, " & VBCRLF
								strSql = strSql & " Case WHEN optlimityn = 'Y' AND optlimitno - optlimitsold <= 5 THEN '0' " & VBCRLF
								strSql = strSql & " 	 WHEN optlimityn = 'Y' AND optlimitno - optlimitsold > 5 THEN optlimitno - optlimitsold - 5 " & VBCRLF
								strSql = strSql & " 	 WHEN optlimityn = 'N' THEN '999' End " & VBCRLF
								strSql = strSql & " , '0', getdate() " & VBCRLF
								strSql = strSql & " FROM db_item.dbo.tbl_item_option " & VBCRLF
								strSql = strSql & " WHERE itemid= '"&iitemid&"' " & VBCRLF
								strSql = strSql & " and replace(optionname, ',', '/') = '"& optionName &"' "
								strSql = strSql & " END "
								dbget.Execute strSql
							End If
						Next
					Set objOption = nothing

					strSql = ""
					strSql = strSql & " UPDATE R " & VBCRLF
					strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VBCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_hmall_regItem R " & VBCRLF
					strSql = strSql & " Join ( " & VBCRLF
					strSql = strSql & " 	SELECT R.itemid, count(*) as CNT, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt "
					strSql = strSql & " 	FROM db_etcmall.dbo.tbl_hmall_regItem R " & VBCRLF
					strSql = strSql & " 	JOIN db_etcmall.dbo.tbl_hmall_regedOption Ro on R.itemid = Ro.itemid and Ro.mallid = 'hmall1010' and Ro.itemid = '"&iitemid&"' " & VBCRLF
					strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid " & VBCRLF
					dbget.Execute strSql
					iErrStr = "OK||"&iitemid&"||성공[재고조회]"
				Else
					iErrStr = "ERR||"&iitemid&"||실패[재고조회] "& html2db(iMessage)
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[재고조회] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[재고조회] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'상품 상세 조회
Public Function fnHmallStatChk2(iitemid, strparam, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, strObj, isSuccessCode, hmallStatcd, xmlDOM, oNode
	Dim slitmAprvlGbcd, slitmAprvlGbcdNm, itemSellGbcd, ihmallSellYn, ihmallPrice, ostkYn, itemAthzGbcd, itemAthzGbcdNm

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "https://openapi.hmall.com//front/pd/pdb/selectItem.do"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "oauserId","002569"
		objXML.setRequestHeader "oauseKey","23439A336B4FC812A1ED415489F185A2"
		objXML.send(strparam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[조회] " & html2db(Err.Description)
			Exit Function
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			For Each oNode In xmlDOM.SelectNodes("/Response2XML/Dataset")
				If oNode.GetAttribute("id") = "result" Then
					isSuccessCode	= oNode.getElementsByTagName("code")(0).text
					iMessage		= oNode.getElementsByTagName("message")(0).text
				End If
			Next

			If isSuccessCode = "0000" Then
				For Each oNode In xmlDOM.SelectNodes("/Response2XML/Dataset")
					If oNode.GetAttribute("id") = "dsItem" Then
						itemSellGbcd 		= oNode.getElementsByTagName("itemSellGbcd")(0).text
						itemAthzGbcd		= oNode.getElementsByTagName("itemSellGbcd")(0).text
						itemAthzGbcdNm		= oNode.getElementsByTagName("itemAthzGbcdNm")(0).text
						slitmAprvlGbcd		= oNode.getElementsByTagName("slitmAprvlGbcd")(0).text
						slitmAprvlGbcdNm	= oNode.getElementsByTagName("slitmAprvlGbcdNm")(0).text
						itemSellGbcd		= oNode.getElementsByTagName("itemSellGbcd")(0).text
						ostkYn				= oNode.getElementsByTagName("ostkYn")(0).text

						Select Case slitmAprvlGbcd
							Case "0"	hmallStatcd = "3"	'임시저장
							Case "1"	hmallStatcd = "3"	'팀장승인대기
							Case "2"	hmallStatcd = "7"	'승인완료
							Case "3"	hmallStatcd = "4"	'취소
							Case "4"	hmallStatcd = "3"	'시스템처리중
							Case "5"	hmallStatcd = "2"	'반송
							Case "6"	hmallStatcd = "3"	'MD승인대기
							Case "7"	hmallStatcd = "3"	'시스템 처리 대기
							Case "9"	hmallStatcd = "4"	'실패
						End Select

						Select Case itemSellGbcd
							Case "00"	ihmallSellYn = "Y"	'진행
							Case "11"	ihmallSellYn = "N"	'일시중단
							Case "19"	ihmallSellYn = "X"	'영구중단
						End Select
					End If
				Next

				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regitem " & VbCRLF
				strSql = strSql & " SET lastConfirmdate = getdate() " & VbCRLF
				strSql = strSql & "	,hmallStatcd = '"&hmallStatcd&"' " & VbCRLF
				strSql = strSql & " WHERE itemid='" & iitemid & "'"
				dbget.Execute(strSql)
				iErrStr = "OK||"&iitemid&"||성공[조회("&slitmAprvlGbcdNm&")]"
			Else
				iErrStr = "ERR||"&iitemid&"||실패[조회] "& html2db(iMessage)
			End If
		Set xmlDOM = nothing
	Set objXML= nothing
End Function

Public Function fnHmallSectView()
    Dim objXML, strSql, i, iRbody, iMessage, strParam, strObj, isSuccess
	Dim xmlDOM, retCode, detailMessage
	Dim Nodes, SubNodes
	Dim sectId, sectNmPath, venDispYn
'	On Error Resume Next
	strParam = ""
	strParam = strParam & "<?xml version=""1.0"" encoding=""EUC-KR""?>"
	strParam = strParam & "<Response2XML>"
	strParam = strParam & "<Dataset id=""dsInput"">"
	strParam = strParam & "<rows>"
	strParam = strParam & "    <row>"
	strParam = strParam & "			<vtltStatGbcd>A</vtltStatGbcd>"		'활성 상태 구분
	strParam = strParam & "			<tlndYn>Y</tlndYn>"					'최하위
	strParam = strParam & "    </row>"
	strParam = strParam & "</rows>"
	strParam = strParam & "</Dataset>"
	strParam = strParam & "</Response2XML>"

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "https://openapi.hmall.com//front/om/oma/selectSectMstList.do", false
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "oauserId", "002569"
		objXML.setRequestHeader "oauseKey", "23439A336B4FC812A1ED415489F185A2"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = Replace(objXML.ResponseText, "&", "_")
'				iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
				xmlDOM.LoadXML iRbody
				If xmlDOM.getElementsByTagName("Response2XML/Dataset/rows/row").length > 0 Then
					Set Nodes = xmlDOM.getElementsByTagName("Response2XML/Dataset/rows/row")
						strSql = ""
						strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_hmall_sectId] "
						dbget.Execute(strSql)
						For each SubNodes in Nodes
							sectId = SubNodes.getElementsByTagName("sectId")(0).Text			'매장 아이디
							sectNmPath = SubNodes.getElementsByTagName("sectNmPath")(0).Text		'매장명 패스
							venDispYn = SubNodes.getElementsByTagName("venDispYn")(0).Text		'협력사 전시여부

							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_hmall_sectId] (sectId, sectNmPath, venDispYn, regdate) VALUES ('"&sectId&"', '"&html2db(sectNmPath)&"', '"&venDispYn&"', GETDATE()) "
							dbget.Execute(strSql)
						Next
					Set Nodes = nothing
					rw xmlDOM.getElementsByTagName("Response2XML/Dataset/rows/row").length & "건 입력"
				End If
			Set xmlDOM = nothing
		Else
			rw "ERR : " & objXML.ResponseText
			response.end
		End If
	Set objXML= nothing
	rw "##### END #####"
	response.end
End Function

Public Function fnHmallInfoDivView()
    Dim objXML, strSql, i, iRbody, iMessage, strParam, strObj, isSuccess
	Dim xmlDOM, retCode, detailMessage
	Dim Nodes, SubNodes
	Dim sectId, sectNmPath, venDispYn, infNotfBsicCd
'	On Error Resume Next
' 1 : R1010101
' 2 : R3020101
' 3 : R3010101
' 4 : R3030101
' 5 : R1010107
' 6 : R6030101
' 7 : S1010101
' 8 : R5030101
' 9 : S1060102
' 10 : S1080105
' 11 : S2030102
' 12 : R6070402
' 13 : S2020201
' 14 : S1020101
' 15 : S7070101
' 16 : R5020205
' 17 : R6040701
' 18 : R5010101
' 19 : R3030201
' 20 : S5010101
' 21 : S5010201
' 22 : S5050101
' 23 : R6020701
' 24 : S8050101
' 25 : S6060502
' 26 : R6050502
' 27 : S9020102
' 28 : S9020101
' 29: 
' 30 : S9011201
' 31 : S9010101
' 32 : S9011101
' 33 : S8080102
' 34 : S8060501
' 35 : S9040108
' 36 : 
' 37 : 
' 38 : R3020109
' 39 : 
' 40 : R5010604
' 41 : S4030201

	strParam = ""
	strParam = strParam & "<?xml version=""1.0"" encoding=""EUC-KR""?>"
	strParam = strParam & "<Root>"
	strParam = strParam & "<Dataset id=""dsCond"">"
	strParam = strParam & "<rows> "
	strParam = strParam & "<row>"
	strParam = strParam & "<itemCsfCd>S4030201</itemCsfCd>"
	strParam = strParam & "</row>"
	strParam = strParam & "</rows>"
	strParam = strParam & "</Dataset>"
	strParam = strParam & "</Root>"

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "https://openapi.hmall.com//front/pd/pdb/selectInfNotfBsicDtlList.do", false
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "oauserId", "002569"
		objXML.setRequestHeader "oauseKey", "23439A336B4FC812A1ED415489F185A2"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
'				iRbody = Replace(objXML.ResponseText, "&", "_")
				iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
				xmlDOM.LoadXML iRbody
				If xmlDOM.getElementsByTagName("Dataset").item(0).attributes(0).nodeValue = "dsInfNotfBsicDtl" Then
					If xmlDOM.getElementsByTagName("Response2XML/Dataset/rows/row").length > 0 Then
						Set Nodes = xmlDOM.getElementsByTagName("Response2XML/Dataset/rows/row")
							For each SubNodes in Nodes
								On Error Resume Next
								rw "(" & SubNodes.getElementsByTagName("itstCd")(0).Text & ")" & SubNodes.getElementsByTagName("itstTitl")(0).Text
								rw "sortOrdg : " & SubNodes.getElementsByTagName("sortOrdg")(0).Text
								rw "---------------------------"
							Next
						Set Nodes = nothing
						response.end
					End If
				End If
			Set xmlDOM = nothing
		Else
			rw "ERR : " & objXML.ResponseText
			response.end
		End If
	Set objXML= nothing
	rw "##### END #####"
	response.end
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################
%>