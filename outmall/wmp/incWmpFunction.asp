<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'상품 등록
Public Function fnWmpItemReg(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj
	If application("Svr_Info") = "Dev" Then
		istrParam = "itemid="&iitemid&"&ScmId=kjy8517"
	Else
		istrParam = "itemid="&iitemid&"&ScmId="&session("SSBctID")
	End If

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info") = "Dev" Then
			objXML.open "POST", "http://localhost:62569/Wemake/Products", false
		Else
			objXML.open "POST", "http://110.93.128.100:8090/wemake/Products", false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상품등록] " & Err.Description
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			'response.write iRbody
			Set strObj = JSON.parse(iRbody)
				strSql = " EXEC db_etcmall.[dbo].[usp_API_Wemake_RegItemInfo_Upd] '"&iitemid&"' "
				dbget.execute strSql

				iErrStr = "OK||"&iitemid&"||성공[상품등록]"
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[상품등록] "&html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품등록] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

'상품 상태 변경
Public Function fnWmpSellyn(iitemid, byRef iErrStr, mustprice, stockCount, ichgSellyn)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode
	If application("Svr_Info") = "Dev" Then
		istrParam = "itemid="&iitemid&"&scmid=kjy8517"
	Else
		istrParam = "itemid="&iitemid&"&scmid="&session("SSBctID")
	End If

	'관리자 강제 삭제처리(따로 API가 있는 것이 아님)
	If ichgSellyn = "X" Then
		strSql = ""
		strSql = strSql &" INSERT INTO [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] " & VBCRLF
		strSql = strSql &" SELECT TOP 1 'WMP', i.itemid, r.wemakeGoodNo, r.wemakeregdate, getdate(), r.lastErrStr" & VBCRLF
		strSql = strSql &" FROM db_item.dbo.tbl_item as i " & VBCRLF
		strSql = strSql &" JOIN db_etcmall.dbo.tbl_wemake_regItem as r on i.itemid = r.itemid " & VBCRLF
		strSql = strSql &" WHERE i.itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)

		strSql = ""
		strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_wemake_regItem " & vbcrlf
		strSql = strSql & " WHERE itemid = '"&iitemid&"' "
		dbget.Execute(strSql)

		strSql = ""
		strSql = strSql & " DELETE FROM db_item.dbo.tbl_outmall_regedoption " & vbcrlf
		strSql = strSql & " WHERE itemid = '"&iitemid&"' " & vbcrlf
		strSql = strSql & " and mallid = '"&CMALLNAME&"' " & vbcrlf
		dbget.Execute(strSql)

		strSql = ""
		strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_outmall_API_Que " & vbcrlf
		strSql = strSql & " WHERE itemid = '"&iitemid&"' " & vbcrlf
		strSql = strSql & " and mallid = '"&CMALLNAME&"' " & vbcrlf
		dbget.Execute(strSql)
		iErrStr =  "OK||"&iitemid&"||삭제[상태수정]"
		Exit Function
	End If

	'판매중으로 돌리려는 데, 재고가 0이면 품절처리
	If ichgSellyn = "Y" and stockCount < 1 Then
		ichgSellyn = "N"
	End If

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info") = "Dev" Then
			If ichgSellyn = "Y" Then
				objXML.open "POST", "http://localhost:62569/Wemake/Products/start", false
			ElseIf ichgSellyn = "N" Then
				objXML.open "POST", "http://localhost:62569/Wemake/Products/stop", false
			End If
		Else
			If ichgSellyn = "Y" Then
				objXML.open "POST", "http://110.93.128.100:8090/wemake/Products/start", false
			ElseIf ichgSellyn = "N" Then
				objXML.open "POST", "http://110.93.128.100:8090/wemake/Products/stop", false
			End If
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상태변경] " & Err.Description
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
'		response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				isSuccessCode		= strObj.code
				iMessage			= strObj.message
				If isSuccessCode = "200" Then
					If ichgSellyn = "Y" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET wemakeSellYn = 'Y'"
						strSql = strSql & "	,wemakePrice = '"& mustprice &"' "
						strSql = strSql & "	,wemakeLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_wemake_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||판매[상태수정]"
					ElseIf ichgSellyn = "N" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET wemakeSellYn = 'N'"
						strSql = strSql & "	,wemakePrice = '"& mustprice &"' "
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,wemakeLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_wemake_regItem  R"
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
	Set objXML = nothing
End Function

'상품 가격 변경
Public Function fnWmpPrice(iitemid, byRef iErrStr, mustprice, stockCount, iOptSellValid)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode, ichgSellyn
	If application("Svr_Info") = "Dev" Then
		istrParam = "itemid="&iitemid&"&scmid=kjy8517"
	Else
		istrParam = "itemid="&iitemid&"&scmid="&session("SSBctID")
	End If

	ichgSellyn = "Y"
	'판매중으로 돌리려는 데, 재고가 0이면 품절처리
	If stockCount < 1 or iOptSellValid = False Then
		ichgSellyn = "N"
	End If

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info") = "Dev" Then
			If ichgSellyn = "Y" Then
				objXML.open "POST", "http://localhost:62569/Wemake/Products/start", false
			ElseIf ichgSellyn = "N" Then
				objXML.open "POST", "http://localhost:62569/Wemake/Products/stop", false
			End If
		Else
			If ichgSellyn = "Y" Then
				objXML.open "POST", "http://110.93.128.100:8090/wemake/Products/start", false
			ElseIf ichgSellyn = "N" Then
				objXML.open "POST", "http://110.93.128.100:8090/wemake/Products/stop", false
			End If
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[가격수정] " & Err.Description
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
'		response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				isSuccessCode		= strObj.code
				iMessage			= strObj.message
				If isSuccessCode = "200" Then
					If ichgSellyn = "Y" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET wemakeSellYn = 'Y'"
						strSql = strSql & "	,wemakePrice = '"& mustprice &"' "
						strSql = strSql & "	,wemakeLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_wemake_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||성공[가격수정]"
					ElseIf ichgSellyn = "N" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET wemakeSellYn = 'N'"
						strSql = strSql & "	,wemakePrice = '"& mustprice &"' "
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,wemakeLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_wemake_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||품절처리[가격수정]"
					End If
				Else
					iErrStr = "ERR||"&iitemid&"||실패[가격수정] "& html2db(iMessage)
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[가격수정] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[가격수정] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

'상품 수정
Public Function fnWmpItemEdit(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj
	If application("Svr_Info") = "Dev" Then
		istrParam = "itemid="&iitemid&"&scmid=kjy8517"
	Else
		istrParam = "itemid="&iitemid&"&scmid="&session("SSBctID")
	End If
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info") = "Dev" Then
			objXML.open "PUT", "http://localhost:62569/Wemake/Products", false
		Else
			objXML.open "PUT", "http://110.93.128.100:8090/wemake/Products", false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상품수정] " & Err.Description
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			strSql = ""
			strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Wemake_EditItemInfo_Upd] '"&iitemid&"' "
			dbget.Execute(strSql)
			iErrStr = "OK||"&iitemid&"||성공[상품수정]"
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					If Instr(iMessage, "텍스트옵션사용여부") > 0 Then
						iMessage = "주문문구추가 불가(등록제외상품에 추가 by.진영)"
						strSql = ""
						strSql = strSql & "	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = 'WMP' AND itemid = '" & iitemid & "') "
						strSql = strSql & "		BEGIN "
						strSql = strSql & "			INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid (itemid, mallgubun) VALUES('" & iitemid & "', 'WMP') "
						strSql = strSql & "		END	"
						dbget.execute strSql
					End If
					iErrStr = "ERR||"&iitemid&"||실패[상품수정] "&html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품수정] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

'상품 조회
Public Function fnWmpStatCheck(iitemid, byRef iErrStr, ilimityn)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode
	Dim objOption, optionValue1, optionValue2, saleStatus, stockCount, displayYn, sellerOptionCode, outmallOptName
	Dim salePrice, itemStockCount, itemsellyn, wemakeGoodNo, productStatusName, saleStartDate
	If application("Svr_Info") = "Dev" Then
		istrParam = "itemid="&iitemid&"&scmid=kjy8517"
	Else
		istrParam = "itemid="&iitemid&"&scmid="&session("SSBctID")
	End If
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info") = "Dev" Then
			objXML.open "GET", "http://localhost:62569/Wemake/Products?"&istrParam, false
		Else
			objXML.open "GET", "http://110.93.128.100:8090/wemake/Products?"&istrParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상품조회] " & Err.Description
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			'rw objXML.Status
			'rw BinaryToText(objXML.ResponseBody,"utf-8")
			'response.end
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

			If (session("ssBctID")="kjy8517") Then
				response.write "<textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If

			Set strObj = JSON.parse(iRbody)
				isSuccessCode		= strObj.code
				iMessage			= strObj.message
				If isSuccessCode = "200" Then
					wemakeGoodNo	= strObj.outPutValue.data.productNo
					salePrice		= strObj.outPutValue.data.sale.salePrice
					itemStockCount	= strObj.outPutValue.data.sale.stockCount
					productStatusName = strObj.outPutValue.data.basic.productStatusName
					saleStartDate	= strObj.outPutValue.data.sale.saleStartDate

					If itemStockCount > 0 Then
						itemsellyn = "Y"
					Else
						itemsellyn = "N"
					End If

					If productStatusName = "판매종료" Then
						itemsellyn = "N"
					End If

					strSQL = ""
					strSQL = strSQL & " DELETE FROM db_item.dbo.tbl_Outmall_regedoption WHERE itemid = '"&iitemid&"' and mallid = '"&CMALLNAME&"' "
					dbget.Execute strSQL

					Set objOption = strObj.outPutValue.data.option.selectOptionValueList
						For i=0 to objOption.length-1
							optionValue1		= objOption.get(i).optionValue1			'옵션값1
							optionValue2		= objOption.get(i).optionValue2			'옵션값2
							saleStatus			= objOption.get(i).saleStatus			'옵션 판매상태(A:판매중, S:품절)
							stockCount			= objOption.get(i).stockCount			'옵션 재고수량 (0 ~ 99999)
							displayYn			= objOption.get(i).displayYn			'옵션 노출여부(Y:노출, N:비노출)
							sellerOptionCode	= objOption.get(i).sellerOptionCode		'업체옵션코드(최대50자)

							If optionValue2 <> "" Then
								outmallOptName = optionValue1 &","&optionValue2
							Else
								outmallOptName = optionValue1
							End If

							If sellerOptionCode = "0000" Then
								strSql = ""
								strSql = strSql & " INSERT INTO db_item.[dbo].[tbl_Outmall_regedoption] (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastupdate) VALUES "
								strSql = strSql & " ('"&iitemid&"', '"&sellerOptionCode&"', '"&CMALLNAME&"', '"&i+1&"', '"&outmallOptName&"', '"& CHKIIF(saleStatus="A", "Y", "N") &"', '"&ilimityn&"', '"&stockCount&"', getdate())"
								dbget.Execute strSql
							Else
								strSql = ""
								strSql = strSql & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_OutMall_regedoption WHERE itemid="&itemid&" and mallid = '"&CMALLNAME&"' and itemoption = '"&sellerOptionCode&"' )"
								strSql = strSql & " BEGIN "
								strSql = strSql & " 	INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
								strSql = strSql & " 	(itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastUpdate, checkdate) "
								strSql = strSql & " 	SELECT itemid, itemoption, '"&CMALLNAME&"', '"&i+1&"', optionname, '"&Chkiif(saleStatus="A","Y","N")&"', '"&ilimityn&"', '"&stockCount&"', getdate(), getdate() "
								strSql = strSql & " 	FROM db_item.dbo.tbl_item_option "
								strSql = strSql & " 	WHERE itemid = '"& itemid &"' "
								strSql = strSql & " 	and itemoption = '"& sellerOptionCode &"' "
								strSql = strSql & " END "
								dbget.Execute strSql
							End If
						Next
						strSql = ""
						strSql = strSql & " UPDATE R " & VBCRLF
						strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VBCRLF
						strSql = strSql & " FROM db_etcmall.dbo.tbl_wemake_regItem R " & VBCRLF
						strSql = strSql & " JOIN ( " & VBCRLF
						strSql = strSql & " 	SELECT R.itemid,count(*) as CNT " & VBCRLF
						strSql = strSql & " 	, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt " & VBCRLF
						strSql = strSql & "     FROM db_etcmall.dbo.tbl_wemake_regItem R " & VBCRLF
						strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid=Ro.itemid " & VBCRLF
						strSql = strSql & "			and Ro.mallid='"&CMALLNAME&"' " & VBCRLF
						strSql = strSql & "			and Ro.itemid= '"& iitemid &"' " & VBCRLF
						strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
						strSql = strSql & " ) T on R.itemid=T.itemid "
						dbget.Execute strSql

						strSql = ""
						strSql = strSql & " UPDATE R" & VbCRLF
						strSql = strSql & " SET wemakePrice = " & salePrice & VbCRLF
						strSql = strSql & " ,wemakeSellyn='"&itemsellyn&"'" & VbCRLF
						strSql = strSql & " ,regitemname = i.itemname " & VbCRLF
						strSql = strSql & " ,regImageName = i.basicImage " & VbCRLF
						strSql = strSql & " ,wemakeRegdate = Case When wemakeRegdate is null then '"& saleStartDate &"' else wemakeRegdate end " & VbCRLF
						strSql = strSql & " ,sellStartDate = Case When sellStartDate is null then '"& saleStartDate &"' else sellStartDate end " & VbCRLF
						strSql = strSql & " FROM db_etcmall.dbo.tbl_wemake_regItem R" & VbCRLF
						strSql = strSql & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid " & VbCRLF
						strSql = strSql & " WHERE R.itemid="&iitemid & VbCRLF
						strSql = strSql & " and isNULL(wemakeGoodNo,'') in ('','"&wemakeGoodNo&"')"&VbCRLF    ''중복등록된CaSE 대비
						strSql = strSql & " and (isNULL(wemakePrice,0)<>"&salePrice&"" & VbCRLF
						strSql = strSql & "     or isNULL(wemakeSellyn,'') <> '"&itemsellyn&"'"& VbCRLF
						strSql = strSql & "     or isNULL(regitemname,'') <> i.basicImage "& VbCRLF
						strSql = strSql & "     or isNULL(wemakeGoodNo,'') <> '"&wemakeGoodNo&"'"& VbCRLF
						strSql = strSql & " )"
						dbget.Execute strSql
						iErrStr = "OK||"&iitemid&"||성공[상품조회]"
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[상품조회] "&html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품조회] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

Public Function fnGetCateList
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode
	Dim retCode, datalist, lcateName, mcateName, scateName, dcateName, dcateCode

	strSql = ""
	strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_wemake_category "
	dbget.Execute(strSql)

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "https://w-api.wemakeprice.com/product/out/getCategory"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "apiKey", "b32bfc8ae3d522eb729e96a60d9d277aeb242302c8f9b33fd51dcc3ee739f19b9d974e2e0a8e1ef683ef3a76e4927378"
		objXML.send()
 		If objXML.Status = "200" OR objXML.Status = "201" Then
 			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			rw iRbody
			Set strObj = JSON.parse(iRbody)
				retCode		= strObj.resultCode			'서버 응답 코드
				If retCode = "200" Then
					Set datalist = strObj.data		'결과리스트 | 결과가 없을 때는 빈 리스트가 리턴
						For i=0 to datalist.length-1
							lcateName = html2db(datalist.get(i).lcateName)
							mcateName = html2db(datalist.get(i).mcateName)
							scateName = html2db(datalist.get(i).scateName)
							dcateName = html2db(datalist.get(i).dcateName)
							dcateCode = datalist.get(i).dcateCode

							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_wemake_category (CateKey, depth1Name, depth2Name, depth3Name, depth4Name, isusing, lastupdate) VALUES  "
							strSql = strSql & " ('"& dcateCode &"', '"& lcateName &"', '"& mcateName &"', '"& scateName &"', '"& dcateName &"', 'Y', GETDATE()) "
							dbget.Execute(strSql)
						Next
						rw "카테고리 " & datalist.length & " 건 등록"
					Set datalist = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function
%>