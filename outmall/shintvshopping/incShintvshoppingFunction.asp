<%
'############################################## ���� �����ϴ� API �Լ� ���� ##############################################
Public Function fnShintvshoppingItemReg(iitemid, strParam, byRef iErrStr, imustprice, ishintvshoppingSellYn, ilimityn, ilimitNo, ilimitSold, iitemname, iimageNm)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode, scmGoodsCode
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/pregoods-base-input/v2", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		response.write BinaryToText(objXML.ResponseBody,"utf-8")
		' response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					scmGoodsCode = strObj.scmGoodsCode

					strSql = ""
					strSql = strSql & " UPDATE R" & VbCRLF
					strSql = strSql & "	Set shintvshoppingGoodNo = '" & scmGoodsCode & "'"  & VbCRLF
					strSql = strSql & "	, shintvshoppingPrice = " &imustprice& VbCRLF
					strSql = strSql & "	, accFailCnt = 0"& VbCRLF
					strSql = strSql & "	, shintvshoppingRegdate = isNULL(shintvshoppingRegdate, getdate())" ''�߰� 2013/02/26
					strSql = strSql & "	, shintvshoppingSellyn = 'Y' "
					If (scmGoodsCode <> "") Then
						strSql = strSql & "	, shintvshoppingstatCD = '3'"& VbCRLF			'���ο�û
					Else
						strSql = strSql & "	, shintvshoppingstatCD = '1'"& VbCRLF			'���۽õ�
					End If
					strSql = strSql & " ,R.reglevel = 1 " & VbCRLF
					strSql = strSql & " ,R.regitemname = i.itemname " & VbCRLF
					strSql = strSql & "	From db_etcmall.dbo.tbl_shintvshopping_regItem R"& VbCRLF
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " Where R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||����[�ӽõ��]"
				Else
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[�ӽõ��]"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REGAddItem-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingContentReg(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/pregoods-describe-input", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCRLF
					strSql = strSql & "	Set R.reglevel = 2 " & VbCRLF
					strSql = strSql & "	From db_etcmall.dbo.tbl_shintvshopping_regItem R"& VbCRLF
					strSql = strSql & " Where R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||����[�����]"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[�����]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REGContent-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingOptReg(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/pregoodsdt-input", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode <> "200" Then
					iErrStr = "ERR"
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingImageReg(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/pregoods-image-url", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCRLF
					strSql = strSql & "	Set R.reglevel = 4 " & VbCRLF
					strSql = strSql & "	,R.regimageName = i.basicImage " & VbCRLF
					strSql = strSql & "	From db_etcmall.dbo.tbl_shintvshopping_regItem R"& VbCRLF
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " Where R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||����[�̹���URL]"
				Else
					rw "req : " & strParam
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[�̹���URL]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-RegImgUrl-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingCert(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/pregoods-cert", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					iErrStr =  "OK||"&iitemid&"||����[��������]"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[��������]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-RegCert-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingGosiReg(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/pregoods-offer", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If (returnCode <> "200") AND (iMessage <> "�̹� ��ϵ� ������ð� �Դϴ�.") Then
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR"
			rw "req : " & strParam
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			rw "-----"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingConfirm(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/pregoods-approval", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCRLF
					strSql = strSql & "	SET R.shintvshoppingstatCD = '7'"& VbCRLF
					strSql = strSql & "	, R.lastConfirmdate = getdate() "& VbCRLF
					strSql = strSql & "	FROM db_etcmall.dbo.tbl_shintvshopping_regItem R"& VbCRLF
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||����[���ο�û]"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[���ο�û]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-Confirm-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingRegChkstat(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", shintvshoppingAPIURL & "/partner/goods/pregoods-detail/v2?" & strParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		response.write BinaryToText(objXML.ResponseBody,"utf-8")
		response.end
	Set objXML= nothing
End Function

Public Function fnShintvshoppingSellyn(iitemid, ichgSellYn, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/sales-no-goods", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message

				If Instr(iMessage, "���� �Ǹ� ���¿� �����մϴ�") > 0 Then
					returnCode = "200"
				End If

				If returnCode = "200" Then
					If ichgSellyn = "Y" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	Set shintvshoppingSellyn = 'Y'"
						strSql = strSql & "	,shintvshoppingLastUpdate = getdate()"
						strSql = strSql & "	From db_etcmall.dbo.tbl_shintvshopping_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||�Ǹ�(���º���)"
					ElseIf ichgSellyn = "N" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	Set shintvshoppingSellyn = 'N'"
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,shintvshoppingLastUpdate = getdate()"
						strSql = strSql & "	From db_etcmall.dbo.tbl_shintvshopping_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||ǰ��ó��(���º���)"
					ElseIf ichgSellyn = "X" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	Set shintvshoppingSellyn = 'X'"
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,shintvshoppingLastUpdate = getdate()"
						strSql = strSql & "	From db_etcmall.dbo.tbl_shintvshopping_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						
						iErrStr =  "OK||"&iitemid&"||�Ǹ�����(���º���)"
					End If
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[���º���]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDITSELLYN-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingItemView(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
	Dim i, saleGb, shintvshoppingPrice, goodsDtList, outmallOptCode, outmallOptName, outmalllimitno, shoppingSellyn, outmallSellyn, AssignedRow
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
'		objXML.open "GET", "http://open-api.shinsegaetvshopping.com/partner/goods/detail/v2?linkCode=TENBY&entpCode=419803&entpId=E419803&entpPass=ten101010*&goodsCode=43472292" , false
		objXML.open "GET", shintvshoppingAPIURL & "/partner/goods/detail/v2?" & strParam , false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		If (session("ssBctID")="kjy8517") Then
			response.write "<textarea cols=40 rows=10>"&BinaryToText(objXML.ResponseBody,"utf-8")&"</textarea>"
		End If
'		response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					saleGb 				= strObj.goodsSelectDetail.get(0).goodsList.saleGb		'00, 11, 19
					shintvshoppingPrice	= strObj.goodsSelectDetail.get(0).goodsList.salePrice
					If saleGb = "00" Then
						shoppingSellyn = "Y"
					Else
						shoppingSellyn = "N"
					End If
					strSql = ""
					strSql =  strSql & " DELETE FROM db_item.dbo.tbl_OutMall_regedoption WHERE mallid='"&CMALLNAME&"' and itemid="&iitemid&" "
					dbget.Execute strSql

					Set goodsDtList = strObj.goodsSelectDetail.get(0).goodsDtList
						For i=0 to goodsDtList.length-1
							outmallOptCode = goodsDtList.get(i).goodsdtCode			'��ǰ�ڵ�
'							rw goodsDtList.get(i).goodsdtInfo						'��ǰ��
							outmallOptName = goodsDtList.get(i).otherText			'�ؽ�Ʈ�Է�
							outmalllimitno = goodsDtList.get(i).maxSaleQty			'�ִ��Ǹż���
							If goodsDtList.get(i).saleGb = "00" Then				'��ǰ�Ǹű����ڵ� | 00: ����  /  11:�Ǹ��ߴ�  / 19: ���
								outmallSellyn = "Y"
							Else
								outmallSellyn = "N"
							End If

							If outmalllimitno < 1 Then
								outmallSellyn = "N"
							End If

							strSql = " INSERT INTO db_item.dbo.tbl_OutMall_regedoption"
							strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outMallSellyn, outmalllimityn, outMallLimitNo)"
							strSql = strSql & " VALUES ("&iitemid
							If i = 0 AND outmallOptName = "���ϻ�ǰ" Then
								strSql = strSql & " ,'0000'"
							Else
								strSql = strSql & " ,'"& i &"'" ''�ӽ÷� �Ե� �ڵ� ���� //2013/04/01
							End If
							strSql = strSql & " ,'"&CMALLNAME&"'"
							strSql = strSql & " ,'"&outmallOptCode&"'"
							strSql = strSql & " ,'"&html2DB(outmallOptName)&"'"
							strSql = strSql & " ,'"&outmallSellyn&"'"
							strSql = strSql & " ,'Y'"
							strSql = strSql & " ,"&outmalllimitno
							strSql = strSql & ")"
							dbget.Execute strSql, AssignedRow

							If (AssignedRow > 0) Then
								strSql = ""
								strSql = strSql & "EXEC [db_etcmall].[dbo].[usp_API_shintvshopping_ItemOptionMapping_Upd] '"& iitemid &"', '"& outmallOptCode &"' "
								dbget.Execute strSql
							End If
						Next
					Set goodsDtList = nothing

					strSql = ""
					strSql = strSql & " UPDATE R " & VbCRLF
					strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VbCRLF
					strSql = strSql & " ,lastStatcheckdate = getdate()"& VbCRLF
					strSql = strSql & " ,shintvshoppingSellyn = '"& shoppingSellyn &"' "& VbCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_shintvshopping_regItem R " & VbCRLF
					strSql = strSql & " JOIN ( " & VbCRLF
					strSql = strSql & " 	SELECT R.itemid,count(*) as CNT "
					strSql = strSql & " 	, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
					strSql = strSql & "		FROM db_etcmall.dbo.tbl_shintvshopping_regItem R " & VbCRLF
					strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro " & VbCRLF
					strSql = strSql & " 		on R.itemid = Ro.itemid"   & VbCRLF
					strSql = strSql & " 		and Ro.mallid = '"&CMALLNAME&"'"   & VbCRLF
					strSql = strSql & "         and Ro.itemid = "&iitemid & VbCRLF
					strSql = strSql & " 	GROUP BY R.itemid "   & VbCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid " & VbCRLF
					dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||����[��ȸ]"
				Else
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[��ȸ]"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-CHKSTAT-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingItemEdit(iitemid, strParam, igetShipCostCode, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode, tmpshipcode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/base/v2", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		' If (session("ssBctID")="kjy8517") Then
		' 	rw "REQ EDIT : <textarea cols=40 rows=10>"&strParam&"</textarea>"
		' 	rw "RES EDIT : <textarea cols=40 rows=10>"&BinaryToText(objXML.ResponseBody,"utf-8")&"</textarea>"
		' End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

			If iitemid = "2998734" Then
				rw "req : " & strParam
				rw "res : " & iRbody
			End If

			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					If igetShipCostCode <> "B01" Then
						tmpshipcode = igetShipCostCode
					Else
						tmpshipcode = ""
					End If

					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET R.regitemname = i.itemname " & VbCRLF
					strSql = strSql & " ,shipcostcode = '"& tmpshipcode &"' " & VbCrlf
					strSql = strSql & " FROM db_etcmall.dbo.tbl_shintvshopping_regItem R" & VbCrlf
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " WHERE R.itemid = " & iitemid
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||����[��������]"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[��������]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRICE-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingEditPrice(iitemid, strParam, imustPrice, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/goodssign", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message

				If Instr(iMessage, "������ ������ ����� �� �����ϴ�") > 0 Then
					returnCode = "200"
				End If

				If returnCode = "200" Then
				    strSql = ""
	    			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_shintvshopping_regItem " & VbCRLF
	    			strSql = strSql & "	SET shintvshoppingLastUpdate=getdate() " & VbCRLF
	    			strSql = strSql & "	, shintvshoppingPrice = " & imustprice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " WHERE itemid='" & iitemid & "'"& VbCRLF
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||����[����]"
				Else
					rw "req : " & strParam
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[����]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRICE-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingEditContentReg(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/describe", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					iErrStr =  "OK||"&iitemid&"||����[�����(����)]"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[�����(����)]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDITContent-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingGosiEdit(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/offer", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If (returnCode <> "200") AND (iMessage <> "�̹� ��ϵ� ������ð� �Դϴ�.") Then
					iErrStr = "ERR"
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDITContent-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingQtyEdit(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/inplanqty-modify", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If (returnCode <> "200")  Then
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDITQty-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingOptSellyn(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/sales-no-goods", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message

				If Instr(iMessage, "���� �Ǹ� ���¿� �����մϴ�") > 0 Then
					returnCode = "200"
				End If

				If (returnCode <> "200")  Then
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDITOptSellyn-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingOptAdd(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/goodsdt/v2", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If (returnCode <> "200")  Then
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDITADDOPT-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingEditImage(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/image-url", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message

				' If Instr(iMessage, "�ε尡 �Ұ�����") > 0 Then
				' 	returnCode = "200"
				' End If

				If returnCode = "200" Then
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET R.regimageName = i.basicImage " & VbCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_shintvshopping_regItem R" & VbCrlf
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " WHERE R.itemid = " & iitemid
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||����[�̹���(����)]"
					rw "req : " & strParam
					rw "res : " & BinaryToText(objXML.ResponseBody,"utf-8")
				Else
					If InStr(iMessage, "�������� �ʰų� �ε尡 �Ұ���") Then
						strSql = ""
						strSql = strSql & " UPDATE db_etcmall.dbo.tbl_shintvshopping_regitem " & VbCrlf
						strSql = strSql & " SET shintvshoppinglastupdate = getdate()" & VbCrlf
						strSql = strSql & " ,accFailCNT=0" & VbCrlf
						strSql = strSql & " ,shintvshoppingSellYn = 'N'" & VbCRLF
						strSql = strSql & " WHERE itemid = " & iitemid
						dbget.execute strSql

						strSql = ""
						strSql = strSql & " IF NOT Exists(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE itemid='"&iitemid&"' and mallgubun = '"&CMALLNAME&"') "
						strSql = strSql & "  BEGIN "
						strSql = strSql & "  	INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid(itemid, mallgubun, bigo) VALUES('"&iitemid&"','"&CMALLNAME&"', '�������� �ʰų� �̹��� ����') "
						strSql = strSql & "  END "
						dbget.Execute strSql
						iErrStr = "ERR||"&iitemid&"||�Ǹ�����[�̹���(����)]/������ ����ó��"
					Else
						rw "req : " & strParam
						rw "res : " & BinaryToText(objXML.ResponseBody,"utf-8")
						iErrStr = "ERR||"&iitemid&"||"&iMessage&"[�̹���(����)]"
					End IF
				End If
			Set strObj = nothing
		Else
			rw "req : " & strParam
			rw "res : " & BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDITContent-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnShintvshoppingEditCert(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/goods/cert", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					iErrStr =  "OK||"&iitemid&"||����[��������(����)]"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[��������(����)]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||shintvshopping ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDITCert-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnGetCommonCodeList(iinterfaceId)
    Dim objXML, iRbody, strObj, returnCode, datalist, i, addReqUrl, addReqParam, groupList, iCode, iName
	addReqParam = "linkCode="&linkCode&"&entpCode="&entpCode&"&entpId="&entpId&"&entpPass="&entpPass
	Select Case iinterfaceId
		Case "IF_API_00_001"	addReqUrl = "/partner/code/md-list"						'MD ����Ʈ
		Case "IF_API_00_006"	addReqUrl = "/partner/code/color-group-code-list"		'��ǰ����׷� ��ȸ
		Case "IF_API_00_007"	addReqUrl = "/partner/code/size-group-code-list"		'��ǰũ��׷� ��ȸ
		Case "IF_API_00_008"	addReqUrl = "/partner/code/form-group-code-list"		'��ǰ���±׷� ��ȸ
		Case "IF_API_00_009"	addReqUrl = "/partner/code/pattern-group-code-list"		'��ǰ���̱׷� ��ȸ
		Case "IF_API_00_010"
			addReqUrl = "/partner/code/color-code-list"									'��ǰ�����ڵ�(����) ��ȸ
			addReqParam = addReqParam & "&cspfGroup="
		Case "IF_API_00_011"
			addReqUrl = "/partner/code/size-code-list"									'��ǰ�����ڵ�(ũ��) ��ȸ
			addReqParam = addReqParam & "&cspfGroup="
		Case "IF_API_00_012"
			addReqUrl = "/partner/code/form-code-list"									'��ǰ�����ڵ�(����) ��ȸ
			addReqParam = addReqParam & "&cspfGroup="
		Case "IF_API_00_013"
			addReqUrl = "/partner/code/pattern-code-list"								'��ǰ�����ڵ�(����) ��ȸ
			addReqParam = addReqParam & "&cspfGroup="
		Case "IF_API_00_014"	addReqUrl = "/partner/code/buy-method-list"				'���Թ�� ��ȸ
		Case "IF_API_00_015"	addReqUrl = "/partner/code/brand-list"					'�귣�� ��ȸ		
		Case "IF_API_00_016"	addReqUrl = "/partner/code/describe-code-list"			'������׸� ��ȸ	
		Case "IF_API_00_017"
			addReqUrl = "/partner/code/entpman-list"									'��ü ����� ��ȸ	
			addReqParam = addReqParam & "&entpManGb=30"									'���к� ����� ��� ��ȸ 10 : ��ǰ�����, 20 : ȸ�������, 30 : �������"		
		Case "IF_API_00_018"	addReqUrl = "/partner/code/origin-list"					'������ ��ȸ		
		Case "IF_API_00_019"	addReqUrl = "/partner/code/make-company-list"			'������ü ��ȸ		
		Case "IF_API_00_020"	addReqUrl = "/partner/code/order-media-list"			'�ֹ���ü ��ȸ		
		Case "IF_API_00_021"	addReqUrl = "/partner/code/nosales-reason-code-list"	'�ǸźҰ� ���� ��ȸ	
		Case "IF_API_00_022"	addReqUrl = "/partner/code/goods-offer-code-list"		'��ǰ����������� ��ǰ���� ��ȸ
		Case "IF_API_00_023"	addReqUrl = "/partner/code/goods-offer-list"			'��ǰ����������� �׸� ��ȸ
		Case "IF_API_00_024"	addReqUrl = "/partner/code/delivery-company-list"		'��ۻ� ��ȸ
		Case "IF_API_00_025"	addReqUrl = "/partner/code/shipping-policy-list"		'�� ��ۺ���å ��� ��ȸ..IF_API_00_030 �̿��ؼ� B001�� ���Է���
		Case "IF_API_00_029"	addReqUrl = "/partner/code/md-man-list"					'���MD ��ȸ
	End Select

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", shintvshoppingAPIURL & addReqUrl & "?" & addReqParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		If iinterfaceId = "IF_API_00_019" Then
			'response.write BinaryToText(objXML.ResponseBody,"utf-8")
			If objXML.Status = "200" OR objXML.Status = "201" Then
				iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
				Set strObj = JSON.parse(iRbody)
					returnCode		= strObj.code
					If returnCode = "200" Then
						Set groupList = strObj.makeCompanyList
							strSql = ""
							strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_shintvshopping_makeCompanyCode] "
							dbget.Execute(strSql)

							For i=0 to groupList.length-1
								iCode		= groupList.get(i).makeCompanyCode		'������ü �ڵ�
								iName		= groupList.get(i).makeCompanyName		'������ü ��

								strSql = ""
								strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_shintvshopping_makeCompanyCode] (makeCompanyCode, makeCompanyName) VALUES "
								strSql = strSql & " ('"&iCode&"', '"&html2db(iName)&"') "
								dbget.Execute(strSql)
								If (i mod 1000) = 0 Then
									response.flush
								End If
							Next
							rw groupList.length & " �� ���"
						Set groupList = nothing
					End If
				Set strObj = nothing
			End If
		ElseIf iinterfaceId = "IF_API_00_018" Then
			'response.write BinaryToText(objXML.ResponseBody,"utf-8")
			If objXML.Status = "200" OR objXML.Status = "201" Then
				iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
				Set strObj = JSON.parse(iRbody)
					returnCode		= strObj.code
					If returnCode = "200" Then
						Set groupList = strObj.originList
							strSql = ""
							strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_shintvshopping_originCode] "
							dbget.Execute(strSql)

							For i=0 to groupList.length-1
								iCode		= groupList.get(i).originCode		'������ �ڵ�
								iName		= groupList.get(i).originName		'������ ��

								strSql = ""
								strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_shintvshopping_originCode] (originCode, originName) VALUES "
								strSql = strSql & " ('"&iCode&"', '"&html2db(iName)&"') "
								dbget.Execute(strSql)
								If (i mod 1000) = 0 Then
									response.flush
								End If
							Next
							rw groupList.length & " �� ���"
						Set groupList = nothing
					End If
				Set strObj = nothing
			End If
		ElseIf iinterfaceId = "IF_API_00_015" Then
			'response.write BinaryToText(objXML.ResponseBody,"utf-8")
			If objXML.Status = "200" OR objXML.Status = "201" Then
				iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
				Set strObj = JSON.parse(iRbody)
					returnCode		= strObj.code
					If returnCode = "200" Then
						Set groupList = strObj.brandList
							strSql = ""
							strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_shintvshopping_brandCode] "
							dbget.Execute(strSql)

							For i=0 to groupList.length-1
								iCode		= groupList.get(i).brandCode		'�귣�� �ڵ�
								iName		= groupList.get(i).brandName		'�귣�� ��Ī

								strSql = ""
								strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_shintvshopping_brandCode] (brandCode, brandName) VALUES "
								strSql = strSql & " ('"&iCode&"', '"&html2db(iName)&"') "
								dbget.Execute(strSql)
								If (i mod 1000) = 0 Then
									response.flush
								End If
							Next
							rw groupList.length & " �� ���"
						Set groupList = nothing
					End If
				Set strObj = nothing
			End If
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
		End If
	Set objXML= nothing
End Function

'��ǰ ���з� ��ȸ
Public Function fnGetGoodsTgroupList()
    Dim objXML, iRbody, strObj, returnCode, i, strSql
	Dim groupList, lgroup,	mgroup,	sgroup,	dgroup,	tgroup,	lgroupName,mgroupName,sgroupName,dgroupName,tgroupName
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", shintvshoppingAPIURL & "/partner/code/goods-tgroup-list?linkCode="&linkCode&"&entpCode="&entpCode&"&entpId="&entpId&"&entpPass="&entpPass&"", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				If returnCode = "200" Then
					Set groupList = strObj.groupList
						strSql = ""
						strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_shintvshopping_category] "
						dbget.Execute(strSql)

						For i=0 to groupList.length-1
							lgroup		= groupList.get(i).lgroup		'CAT �ڵ�
							mgroup		= groupList.get(i).mgroup		'��з� �ڵ�
							sgroup		= groupList.get(i).sgroup		'�ߺз� �ڵ�
							dgroup		= groupList.get(i).dgroup		'�Һз� �ڵ�
							tgroup		= groupList.get(i).tgroup		'���з� �ڵ�
							lgroupName	= groupList.get(i).lgroupName	'CAT ��
							mgroupName	= groupList.get(i).mgroupName	'��з� ��
							sgroupName	= groupList.get(i).sgroupName	'�ߺз� ��
							dgroupName	= groupList.get(i).dgroupName	'�Һз� ��
							tgroupName	= groupList.get(i).tgroupName	'���з� ��

							strSql = ""
							strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Shintvshopping_Category_Ins] '"&lgroup&"', '"&mgroup&"', '"&sgroup&"', '"&dgroup&"', '"&tgroup&"' " & VBCRLF
							strSql = strSql & " ,'"&lgroupName&"' ,'"&mgroupName&"' ,'"&sgroupName&"' ,'"&dgroupName&"' ,'"&tgroupName&"'  "
							dbget.Execute(strSql)
						Next
						rw "�ż���Ȩ���� ī�װ� " & groupList.length & " �� ���"
					Set groupList = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'��ǰ����������� �׸� ��ȸ
Public Function fnGetOfferList()
    Dim objXML, iRbody, strObj, returnCode, i, strSql
	Dim offerList, offerCode, offerName, typeCode, typeName
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", shintvshoppingAPIURL & "/partner/code/goods-offer-list?linkCode="&linkCode&"&entpCode="&entpCode&"&entpId="&entpId&"&entpPass="&entpPass&"", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				If returnCode = "200" Then
					Set offerList = strObj.offerList
						strSql = ""
						strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_shintvshopping_infocd] "
						dbget.Execute(strSql)
						For i=0 to offerList.length-1
							offerCode		= offerList.get(i).offerCode		'CAT �ڵ�
							offerName		= html2db(offerList.get(i).offerName)		'��з� �ڵ�
							typeCode		= offerList.get(i).typeCode		'�ߺз� �ڵ�
							typeName		= offerList.get(i).typeName		'�Һз� �ڵ�

							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_shintvshopping_infocd]  (offerCode, offerName, typeCode, typeName) VALUES "
							strSql = strSql & " ('"& offerCode &"', '"& offerName &"', '"& typeCode &"', '"& typeName &"') "
							dbget.Execute(strSql)
						Next
						rw "��ǰ����������� �׸� " & offerList.length & " �� ���"
					Set offerList = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'�������� �׸���ȸ
Public Function fnGetCertList(ilmsdCode)
    Dim objXML, iRbody, strObj, returnCode, i, strSql
	Dim goodskindsCertList, safetyCert, safetyCertYn, safetyConfirm, safetyConfirmYn, suppSuit, suppSuitYn, radioWaveCert, radioWaveCertYn, childSafetyCert, childSafetyCertYn, childSafetyConfirm, childSafetyConfirmYn, childSuppSuit, childSuppSuitYn
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", shintvshoppingAPIURL & "/partner/code/cert-list?linkCode="&linkCode&"&entpCode="&entpCode&"&entpId="&entpId&"&entpPass="&entpPass&"&lmsdCode="&ilmsdCode, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				If returnCode = "200" Then
					Set goodskindsCertList = strObj.goodskindsCertList
						For i=0 to goodskindsCertList.length-1
							safetyCert				= goodskindsCertList.get(i).safetyCert				'�����������
							safetyCertYn			= goodskindsCertList.get(i).safetyCertYn			'����������󿩺�
							safetyConfirm			= goodskindsCertList.get(i).safetyConfirm			'����Ȯ�δ��
							safetyConfirmYn			= goodskindsCertList.get(i).safetyConfirmYn			'����Ȯ�δ�󿩺�
							suppSuit				= goodskindsCertList.get(i).suppSuit				'���������ռ�Ȯ�δ��
							suppSuitYn				= goodskindsCertList.get(i).suppSuitYn				'���������ռ�Ȯ�δ�󿩺�
							radioWaveCert			= goodskindsCertList.get(i).radioWaveCert			'�����������
							radioWaveCertYn			= goodskindsCertList.get(i).radioWaveCertYn			'����������󿩺�
							childSafetyCert			= goodskindsCertList.get(i).childSafetyCert			'��̾����������
							childSafetyCertYn		= goodskindsCertList.get(i).childSafetyCertYn		'��̾���������󿩺�
							childSafetyConfirm		= goodskindsCertList.get(i).childSafetyConfirm		'��̾���Ȯ�δ��
							childSafetyConfirmYn	= goodskindsCertList.get(i).childSafetyConfirmYn	'��̾���Ȯ�δ�󿩺�
							childSuppSuit			= goodskindsCertList.get(i).childSuppSuit			'��̰��������ռ�Ȯ�δ��
							childSuppSuitYn			= goodskindsCertList.get(i).childSuppSuitYn			'��̰��������ռ�Ȯ�δ�󿩺�

							strSql = ""
							strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Shintvshopping_Category_Upd] '"&ilmsdCode&"', '"&safetyCertYn&"', '"&safetyConfirmYn&"', '"&suppSuitYn&"', '"&radioWaveCertYn&"', '"&childSafetyCertYn&"' " & VBCRLF
							strSql = strSql & " ,'"&childSafetyConfirmYn&"' ,'"&childSuppSuitYn&"' "
							dbget.Execute(strSql)
						Next
					Set goodskindsCertList = nothing
				Else
					rw "---------"
					rw "���� : " & ilmsdCode
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					rw "---------"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'�� ��ۺ���å ���..Json Body�� ����ϴ� �� �ƴ϶� ������ param ����ؾ� �۵���;
Public Function fnInputCustShipCost()
    Dim objXML, iRbody, strObj, returnCode, i, strSql
	Dim obj, strParam, addReqParam
	
	addReqParam = "linkCode="&linkCode&"&entpCode="&entpCode&"&entpId="&entpId&"&entpPass="&entpPass
	addReqParam = addReqParam & "&shipCostCode=B002"
	addReqParam = addReqParam & "&shipCostName=6�����̻� ����"
	addReqParam = addReqParam & "&shipCostBaseAmt=60000"
	addReqParam = addReqParam & "&applyDate=202112271600"
	addReqParam = addReqParam & "&normalOrderAmt=3000"
	addReqParam = addReqParam & "&normalReturnAmt=3000"
	addReqParam = addReqParam & "&normalExchangeAmt=3000"
	addReqParam = addReqParam & "&distantOrderAmt=3000"
	addReqParam = addReqParam & "&distantReturnAmt=3000"
	addReqParam = addReqParam & "&distantExchangeAmt=3000"
	addReqParam = addReqParam & "&jejuOrderAmt=3000"
	addReqParam = addReqParam & "&jejuReturnAmt=3000"
	addReqParam = addReqParam & "&jejuExchangeAmt=3000"

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/code/cust-ship-cost-input", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(addReqParam)
		response.write BinaryToText(objXML.ResponseBody,"utf-8")
		response.end
	Set objXML= nothing
End Function

Function ArrErrStrInfo(iaction, iitemid, ierrVendorItemId)
	Dim ErrStrComma, strSql
	If iaction = "REGOpt" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||����[�ɼǵ��] " & ErrStrComma
		Else
			strSql = ""
			strSql = strSql & " UPDATE R" & VbCRLF
			strSql = strSql & "	Set R.reglevel = 3 " & VbCRLF
			strSql = strSql & "	From db_etcmall.dbo.tbl_shintvshopping_regItem R"& VbCRLF
			strSql = strSql & " Where R.itemid = '" & iitemid & "'"
			dbget.Execute(strSql)
			ArrErrStrInfo =  "OK||"&iitemid&"||����[�ɼǵ��]"
		End If
	ElseIf iaction = "REGGosi" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||����[�������] " & ErrStrComma
		Else
			strSql = ""
			strSql = strSql & " UPDATE R" & VbCRLF
			strSql = strSql & "	Set R.reglevel = 5 " & VbCRLF
			strSql = strSql & "	From db_etcmall.dbo.tbl_shintvshopping_regItem R"& VbCRLF
			strSql = strSql & " Where R.itemid = '" & iitemid & "'"
			dbget.Execute(strSql)
			ArrErrStrInfo =  "OK||"&iitemid&"||����[�������]"
		End If
	ElseIf iaction = "EDITGosi" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||����[�������] " & ErrStrComma
		Else
			ArrErrStrInfo =  "OK||"&iitemid&"||����[�������]"
		End If
	ElseIf iaction = "EDITQTY" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||����[������] " & ErrStrComma
		Else
			ArrErrStrInfo =  "OK||"&iitemid&"||����[������]"
		End If
	ElseIf iaction = "EDITSTAT" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||����[�ɼǻ���] " & ErrStrComma
		Else
			ArrErrStrInfo =  "OK||"&iitemid&"||����[�ɼǻ���]"
		End If
	ElseIf iaction = "EDITADDOPT" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||����[�ɼ��߰�] " & ErrStrComma
		Else
			ArrErrStrInfo =  "OK||"&iitemid&"||����[�ɼ��߰�]"
		End If
	End If
End Function
%>