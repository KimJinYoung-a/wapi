<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
Public Function fnItemCopy(iitemid, imakerid, iitemdiv, ividx, iiErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccess, copyItemid
	istrParam = ""
	istrParam = istrParam & "itemid="&iitemid
	istrParam = istrParam & "&brandId="&imakerid
	istrParam = istrParam & "&itemdiv="&iitemdiv
	istrParam = istrParam & "&idx="&ividx

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://192.168.50.4:8090/scmapi/item/itemcopy", false
		Else
			objXML.open "POST", "http://110.93.128.100:8090/scmapi/item/itemcopy", false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[err] " & replace(replace(replace(html2db(Err.Description), vbCrLf, ""), vbCr, ""), vbLf, "")
			strSql = "EXEC [db_item].[dbo].[usp_API_itemcopy_Upd] 'U', '"& ividx &"', '', '', 'ERR', '실패[err] " & html2db(Err.Description)&"', ''"
			dbget.execute strSql
		Else
			If objXML.Status = "200" OR objXML.Status = "201" Then
				iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
				Set strObj = JSON.parse(iRbody)
					isSuccess		= strObj.success
					iMessage		= strObj.message
					copyItemid		= strObj.outPutValue
					If isSuccess = true Then
						iErrStr = "OK||"&iitemid&"||성공["& copyItemid &"] "
					Else
						iErrStr = "ERR||"&iitemid&"||실패[복제] " & iMessage
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
						iErrStr = "ERR||"&iitemid&"||실패[복제] "& html2db(iMessage)
						strSql = "EXEC [db_item].[dbo].[usp_API_itemcopy_Upd] 'U', '"& ividx &"', '', '', 'ERR', '실패[msg] "&html2db(iMessage)&"', ''"
						dbget.execute strSql
					Else
						iErrStr = "ERR||"&iitemid&"||실패[복제_NO] 통신오류"
						strSql = "EXEC [db_item].[dbo].[usp_API_itemcopy_Upd] 'U', '"& ividx &"', '', '', 'ERR', '실패[notmsg] 통신오류', ''"
						dbget.execute strSql
					End If
				Set strObj = nothing
			End If
		End If
	Set objXML= nothing
	On Error Goto 0
End Function

Public Function fnFindMakerid(imakerid)
	Dim sqlStr
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as CNT "
	sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c with (nolock) "
	sqlStr = sqlStr & " WHERE userid = '"& imakerid &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("CNT") > 0 Then
		fnFindMakerid = "Y"
	Else
		fnFindMakerid = "N"
	End If
	rsget.Close
End Function

Public Function fnFindItemdiv(iitemdiv)
	Dim sqlStr
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as CNT "
	sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_item_div] with (nolock) "
	sqlStr = sqlStr & " WHERE code = '"& iitemdiv &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("CNT") > 0 Then
		fnFindItemdiv = "Y"
	Else
		fnFindItemdiv = "N"
	End If
	rsget.Close
End Function

Public Function fnFindItemid(iitemid, retCopyitemid)
	Dim sqlStr
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as CNT, max(copyitemid) as copyitemid "
	sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_api_item_copy] with (nolock) "
	sqlStr = sqlStr & " WHERE itemid = '"& iitemid &"' "
	sqlStr = sqlStr & " and isNull(copyitemid, '') <> ''  "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("CNT") > 0 Then
		retCopyitemid = rsget("copyitemid")
		fnFindItemid = "Y"
	Else
		fnFindItemid = "N"
	End If
	rsget.Close
End Function

%>
