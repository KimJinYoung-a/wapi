<%
'############################################## ���� �����ϴ� API �Լ� ���� ���� ############################################
'��ǰ ���
Public Function fnWmpfashionItemReg(iitemid, byRef iErrStr)
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
			objXML.open "POST", "http://110.93.128.100:8090/fwmp/Products", false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ǰ���] " & Err.Description
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			'response.write iRbody
			Set strObj = JSON.parse(iRbody)
				strSql = " EXEC db_etcmall.[dbo].[usp_API_wfWemake_RegItemInfo_Upd] '"&iitemid&"' "
				dbget.execute strSql

				iErrStr = "OK||"&iitemid&"||����[��ǰ���]"
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||����[��ǰ���] "&html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||����[��ǰ���] ��ſ���"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

'��ǰ ���� ����
Public Function fnWmpfashionSellyn(iitemid, byRef iErrStr, mustprice, stockCount, ichgSellyn)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode
	If application("Svr_Info") = "Dev" Then
		istrParam = "itemid="&iitemid&"&scmid=kjy8517"
	Else
		istrParam = "itemid="&iitemid&"&scmid="&session("SSBctID")
	End If

	'������ ���� ����ó��(���� API�� �ִ� ���� �ƴ�)
	If ichgSellyn = "X" Then
		strSql = ""
		strSql = strSql &" INSERT INTO [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] " & VBCRLF
		strSql = strSql &" SELECT TOP 1 'wmpfashion', i.itemid, r.wfwemakeGoodNo, r.wfwemakeregdate, getdate(), r.lastErrStr" & VBCRLF
		strSql = strSql &" FROM db_item.dbo.tbl_item as i " & VBCRLF
		strSql = strSql &" JOIN db_etcmall.dbo.tbl_wfwemake_regItem as r on i.itemid = r.itemid " & VBCRLF
		strSql = strSql &" WHERE i.itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)

		strSql = ""
		strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_wfwemake_regItem " & vbcrlf
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
		iErrStr =  "OK||"&iitemid&"||����[���¼���]"
		Exit Function
	End If

	'�Ǹ������� �������� ��, ��� 0�̸� ǰ��ó��
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
				objXML.open "POST", "http://110.93.128.100:8090/fwmp/Products/start", false
			ElseIf ichgSellyn = "N" Then
				objXML.open "POST", "http://110.93.128.100:8090/fwmp/Products/stop", false
			End If
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[���º���] " & Err.Description
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
						strSql = strSql & "	SET wfwemakeSellYn = 'Y'"
						strSql = strSql & "	,wfwemakePrice = '"& mustprice &"' "
						strSql = strSql & "	,wfwemakeLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_wfwemake_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||�Ǹ�[���¼���]"
					ElseIf ichgSellyn = "N" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET wfwemakeSellYn = 'N'"
						strSql = strSql & "	,wfwemakePrice = '"& mustprice &"' "
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,wfwemakeLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_wfwemake_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||ǰ��ó��[���¼���]"
					End If
				Else
					iErrStr = "ERR||"&iitemid&"||����[���¼���] "& html2db(iMessage)
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||����[���¼���] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||����[���¼���] ��ſ���"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

'��ǰ ���� ����
Public Function fnWmpfashionPrice(iitemid, byRef iErrStr, mustprice, stockCount, iOptSellValid)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode, ichgSellyn
	If application("Svr_Info") = "Dev" Then
		istrParam = "itemid="&iitemid&"&scmid=kjy8517"
	Else
		istrParam = "itemid="&iitemid&"&scmid="&session("SSBctID")
	End If

	ichgSellyn = "Y"
	'�Ǹ������� �������� ��, ��� 0�̸� ǰ��ó��
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
				objXML.open "POST", "http://110.93.128.100:8090/fwmp/Products/start", false
			ElseIf ichgSellyn = "N" Then
				objXML.open "POST", "http://110.93.128.100:8090/fwmp/Products/stop", false
			End If
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[���ݼ���] " & Err.Description
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
						strSql = strSql & "	SET wfwemakeSellYn = 'Y'"
						strSql = strSql & "	,wfwemakePrice = '"& mustprice &"' "
						strSql = strSql & "	,wfwemakeLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_wfwemake_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||����[���ݼ���]"
					ElseIf ichgSellyn = "N" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET wfwemakeSellYn = 'N'"
						strSql = strSql & "	,wfwemakePrice = '"& mustprice &"' "
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,wfwemakeLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_wfwemake_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||ǰ��ó��[���ݼ���]"
					End If
				Else
					iErrStr = "ERR||"&iitemid&"||����[���ݼ���] "& html2db(iMessage)
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||����[���ݼ���] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||����[���ݼ���] ��ſ���"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

'��ǰ ����
Public Function fnWmpfashionItemEdit(iitemid, byRef iErrStr)
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
			objXML.open "PUT", "http://110.93.128.100:8090/fwmp/Products", false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ǰ����] " & Err.Description
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			strSql = ""
			strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_wfWemake_EditItemInfo_Upd] '"&iitemid&"' "
			dbget.Execute(strSql)
			iErrStr = "OK||"&iitemid&"||����[��ǰ����]"
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					If Instr(iMessage, "�ؽ�Ʈ�ɼǻ�뿩��") > 0 Then
						iMessage = "�ֹ������߰� �Ұ�(������ܻ�ǰ�� �߰� by.����)"
						strSql = ""
						strSql = strSql & "	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = 'wmpfashion' AND itemid = '" & iitemid & "') "
						strSql = strSql & "		BEGIN "
						strSql = strSql & "			INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid (itemid, mallgubun) VALUES('" & iitemid & "', 'wmpfashion') "
						strSql = strSql & "		END	"
						dbget.execute strSql
					End If
					iErrStr = "ERR||"&iitemid&"||����[��ǰ����] "&html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||����[��ǰ����] ��ſ���"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

'��ǰ ��ȸ
Public Function fnWmpfashionStatCheck(iitemid, byRef iErrStr, ilimityn)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode
	Dim objOption, optionValue1, optionValue2, saleStatus, stockCount, displayYn, sellerOptionCode, outmallOptName
	Dim salePrice, itemStockCount, itemsellyn, wfwemakeGoodNo, productStatusName
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
			objXML.open "GET", "http://110.93.128.100:8090/fwmp/Products?"&istrParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ǰ��ȸ] " & Err.Description
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			'rw objXML.Status
			'rw BinaryToText(objXML.ResponseBody,"utf-8")
			'response.end
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				isSuccessCode		= strObj.code
				iMessage			= strObj.message
				If isSuccessCode = "200" Then
					wfwemakeGoodNo	= strObj.outPutValue.data.productNo
					salePrice		= strObj.outPutValue.data.sale.salePrice
					itemStockCount	= strObj.outPutValue.data.sale.stockCount
					productStatusName = strObj.outPutValue.data.basic.productStatusName

					If itemStockCount > 0 Then
						itemsellyn = "Y"
					Else
						itemsellyn = "N"
					End If

					If productStatusName = "�Ǹ�����" Then
						itemsellyn = "N"
					End If

					strSQL = ""
					strSQL = strSQL & " DELETE FROM db_item.dbo.tbl_Outmall_regedoption WHERE itemid = '"&iitemid&"' and mallid = '"&CMALLNAME&"' "
					dbget.Execute strSQL

					Set objOption = strObj.outPutValue.data.option.selectOptionValueList
						For i=0 to objOption.length-1
							optionValue1		= objOption.get(i).optionValue1			'�ɼǰ�1
							optionValue2		= objOption.get(i).optionValue2			'�ɼǰ�2
							saleStatus			= objOption.get(i).saleStatus			'�ɼ� �ǸŻ���(A:�Ǹ���, S:ǰ��)
							stockCount			= objOption.get(i).stockCount			'�ɼ� ������ (0 ~ 99999)
							displayYn			= objOption.get(i).displayYn			'�ɼ� ���⿩��(Y:����, N:�����)
							sellerOptionCode	= objOption.get(i).sellerOptionCode		'��ü�ɼ��ڵ�(�ִ�50��)

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
						strSql = strSql & " FROM db_etcmall.dbo.tbl_wfwemake_regItem R " & VBCRLF
						strSql = strSql & " JOIN ( " & VBCRLF
						strSql = strSql & " 	SELECT R.itemid,count(*) as CNT " & VBCRLF
						strSql = strSql & " 	, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt " & VBCRLF
						strSql = strSql & "     FROM db_etcmall.dbo.tbl_wfwemake_regItem R " & VBCRLF
						strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid=Ro.itemid " & VBCRLF
						strSql = strSql & "			and Ro.mallid='"&CMALLNAME&"' " & VBCRLF
						strSql = strSql & "			and Ro.itemid= '"& iitemid &"' " & VBCRLF
						strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
						strSql = strSql & " ) T on R.itemid=T.itemid "
						dbget.Execute strSql

						strSql = ""
						strSql = strSql & " UPDATE R" & VbCRLF
						strSql = strSql & " SET wfwemakePrice = " & salePrice & VbCRLF
						strSql = strSql & " ,wfwemakeSellyn='"&itemsellyn&"'" & VbCRLF
						strSql = strSql & " ,regitemname = i.itemname " & VbCRLF
						strSql = strSql & " ,regImageName = i.basicImage " & VbCRLF
						strSql = strSql & " FROM db_etcmall.dbo.tbl_wfwemake_regItem R" & VbCRLF
						strSql = strSql & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid " & VbCRLF
						strSql = strSql & " WHERE R.itemid="&iitemid & VbCRLF
						strSql = strSql & " and isNULL(wfwemakeGoodNo,'') in ('','"&wfwemakeGoodNo&"')"&VbCRLF    ''�ߺ���ϵ�CaSE ���
						strSql = strSql & " and (isNULL(wfwemakePrice,0)<>"&salePrice&"" & VbCRLF
						strSql = strSql & "     or isNULL(wfwemakeSellyn,'') <> '"&itemsellyn&"'"& VbCRLF
						strSql = strSql & "     or isNULL(regitemname,'') <> i.basicImage "& VbCRLF
						strSql = strSql & "     or isNULL(wemakeGoodNo,'') <> '"&wfwemakeGoodNo&"'"& VbCRLF
						strSql = strSql & " )"
						dbget.Execute strSql
						iErrStr = "OK||"&iitemid&"||����[��ǰ��ȸ]"
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||����[��ǰ��ȸ] "&html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||����[��ǰ��ȸ] ��ſ���"
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function
%>