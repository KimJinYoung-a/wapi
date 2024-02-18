<%
'############################################## 실제 수행하는 API 함수 모음 ##############################################
Public Function fnskstoaItemReg(iitemid, strParam, byRef iErrStr, imustprice, iskstoaSellYn, ilimityn, ilimitNo, ilimitSold, iitemname, iimageNm)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode, scmGoodsCode
	Dim marginRate
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/pregoods-base-input", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		' response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					scmGoodsCode = strObj.scmGoodsCode
					strSql = ""
					strSql = strSql & " SELECT TOP 1 m.margin, d.itemid "
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_ssg_marginItem_master] as m  "
					strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_ssg_marginItem_detail] as d on m.idx = d.midx  "
					strSql = strSql & " WHERE m.isusing = 'Y'  "
					strSql = strSql & " and convert(char(10), getdate(), 120) between m.startDate and m.enddate  "
					strSql = strSql & " and m.mallid = 'skstoa' "
					strSql = strSql & " and d.itemid = '"& iitemid &"' "
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If Not(rsget.EOF or rsget.BOF) Then
						marginRate = rsget("margin")
					Else
						marginRate = 12
					End If
					rsget.Close

					strSql = ""
					strSql = strSql & " UPDATE R" & VbCRLF
					strSql = strSql & "	Set skstoaTmpGoodNo = '" & scmGoodsCode & "'"  & VbCRLF
					strSql = strSql & "	, skstoaPrice = " &imustprice& VbCRLF
					strSql = strSql & "	, accFailCnt = 0"& VbCRLF
					strSql = strSql & "	, skstoaRegdate = isNULL(skstoaRegdate, getdate())" ''추가 2013/02/26
					strSql = strSql & "	, skstoaSellyn = 'Y' "
					If (scmGoodsCode <> "") Then
						strSql = strSql & "	, skstoastatCD = '3'"& VbCRLF			'승인요청
					Else
						strSql = strSql & "	, skstoastatCD = '1'"& VbCRLF			'전송시도
					End If
					strSql = strSql & " ,R.reglevel = 1 " & VbCRLF
					strSql = strSql & " ,R.regitemname = i.itemname " & VbCRLF
					strSql = strSql & " ,R.setMargin = '"& marginRate &"' " & VbCRLF
					strSql = strSql & "	From db_etcmall.dbo.tbl_skstoa_regItem R"& VbCRLF
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " Where R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||성공[임시등록]"
				Else
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[임시등록]"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-REGAddItem-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaContentReg(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/pregoods-describe-input", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
'		response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCRLF
					strSql = strSql & "	Set R.reglevel = 2 " & VbCRLF
					strSql = strSql & "	From db_etcmall.dbo.tbl_skstoa_regItem R"& VbCRLF
					strSql = strSql & " Where R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||성공[기술서]"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[기술서]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-REGContent-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaOptReg(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/pregoodsdt-input", false
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

Public Function fnSkstoaImageReg(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/pregoods-image-url", false
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
					strSql = strSql & "	From db_etcmall.dbo.tbl_skstoa_regItem R"& VbCRLF
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " Where R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||성공[이미지URL]"
				Else
					rw "req : " & strParam
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[이미지URL]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-RegImgUrl-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaGosiReg(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/pregoods-offer", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If (returnCode <> "200") AND (iMessage <> "이미 등록된 정보고시건 입니다.") Then
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

Public Function fnSkstoaCert(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/pregoods-kcinfo-input", false
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
					iErrStr =  "OK||"&iitemid&"||성공[인증정보]"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[인증정보]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-RegCert-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaConfirm(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/pregoods-approval", false
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
					strSql = strSql & "	SET R.sendConfirm = 'Y' "& VbCRLF
					strSql = strSql & "	FROM db_etcmall.dbo.tbl_skstoa_regItem R"& VbCRLF
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||성공[승인요청]"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[승인요청]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-Confirm-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnskstoaRegChkstat(iitemid, strParam, byRef iErrStr, byRef iskGoodNo)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode, skstoaGoodNo
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", skstoaAPIURL & "/partner/goods/pregoods-detail?" & strParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			If (session("ssBctID")="kjy8517") Then
				rw "승인요청 : <textarea cols=40 rows=10>"&BinaryToText(objXML.ResponseBody,"utf-8")&"</textarea>"
			End If
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					skstoaGoodNo = strObj.entpGoodsAll.get(0).entpGoodsList.confirmGoodsCode
					If skstoaGoodNo <> "" Then
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCRLF
						strSql = strSql & "	SET R.lastConfirmdate = getdate() "& VbCRLF
						strSql = strSql & "	, R.skstoastatCD = '7'"& VbCRLF
						strSql = strSql & "	, R.skstoaGoodNo = '"& skstoaGoodNo &"'"& VbCRLF
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_skstoa_regItem R"& VbCRLF
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iskGoodNo = skstoaGoodNo
						iErrStr =  "OK||"&iitemid&"||성공[승인완료]"
					Else
						iErrStr =  "OK||"&iitemid&"||성공[승인대기]"
					End If
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[승인대기]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-Confirm-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaSellyn(iitemid, ichgSellYn, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/sales-no-goods", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		response.write BinaryToText(objXML.ResponseBody,"utf-8")

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message

				If Instr(iMessage, "판매구분을 변경할 단품이 존재하지 않습니다") > 0 Then
					returnCode = "200"
				End If

				If returnCode = "200" Then
					If ichgSellyn = "Y" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET skstoaSellyn = 'Y'"
						strSql = strSql & "	,skstoaLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_skstoa_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||판매(상태변경)"
					ElseIf ichgSellyn = "N" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET skstoaSellyn = 'N'"
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,skstoaLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_skstoa_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||품절처리(상태변경)"
					ElseIf ichgSellyn = "X" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET skstoaSellyn = 'X'"
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,skstoaLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_skstoa_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						
						iErrStr =  "OK||"&iitemid&"||판매종료(상태변경)"
					End If
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[상태변경]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-EDITSELLYN-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaItemView(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
	Dim i, saleGb, skstoaPrice, goodsDtList, outmallOptCode, outmallOptName, outmalllimitno, stoaSellyn, outmallSellyn, AssignedRow
'	On Error Resume Next
'response.write strParam
'response.end
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", skstoaAPIURL & "/partner/goods/detail?" & strParam , false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		If (session("ssBctID")="kjy8517") Then
			rw "조회 : <textarea cols=40 rows=10>"&BinaryToText(objXML.ResponseBody,"utf-8")&"</textarea>"
		End If
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					saleGb 				= strObj.goodsSelectDetail.get(0).goodsList.saleGb		'00, 11, 19
					skstoaPrice			= strObj.goodsSelectDetail.get(0).goodsList.salePrice
					If saleGb = "00" Then
						stoaSellyn = "Y"
					Else
						stoaSellyn = "N"
					End If
					strSql = ""
					strSql =  strSql & " DELETE FROM db_item.dbo.tbl_OutMall_regedoption WHERE mallid='"&CMALLNAME&"' and itemid="&iitemid&" "
					dbget.Execute strSql

					Set goodsDtList = strObj.goodsSelectDetail.get(0).goodsDtList
						For i=0 to goodsDtList.length-1
							outmallOptCode = goodsDtList.get(i).goodsdtCode			'단품코드
'							rw goodsDtList.get(i).goodsdtInfo						'단품상세
							outmallOptName = goodsDtList.get(i).otherText			'텍스트입력
							outmalllimitno = goodsDtList.get(i).maxSaleQty			'최대판매수량
							If goodsDtList.get(i).saleGb = "00" Then				'단품판매구분코드 | 00: 진행  /  11:판매중단  / 19: 폐기
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
							If i = 0 AND outmallOptName = "단일상품" Then
								strSql = strSql & " ,'0000'"
							Else
								strSql = strSql & " ,'"& i &"'" ''임시로 롯데 코드 넣음 //2013/04/01
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
								strSql = strSql & "EXEC [db_etcmall].[dbo].[usp_API_skstoa_ItemOptionMapping_Upd] '"& iitemid &"', '"& outmallOptCode &"' "
								dbget.Execute strSql
							End If
						Next
					Set goodsDtList = nothing

					strSql = ""
					strSql = strSql & " UPDATE R " & VbCRLF
					strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VbCRLF
					strSql = strSql & " ,lastStatcheckdate = getdate()"& VbCRLF
					strSql = strSql & " ,skstoaSellyn = '"& stoaSellyn &"' "& VbCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_skstoa_regItem R " & VbCRLF
					strSql = strSql & " JOIN ( " & VbCRLF
					strSql = strSql & " 	SELECT R.itemid,count(*) as CNT "
					strSql = strSql & " 	, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
					strSql = strSql & "		FROM db_etcmall.dbo.tbl_skstoa_regItem R " & VbCRLF
					strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro " & VbCRLF
					strSql = strSql & " 		on R.itemid = Ro.itemid"   & VbCRLF
					strSql = strSql & " 		and Ro.mallid = '"&CMALLNAME&"'"   & VbCRLF
					strSql = strSql & "         and Ro.itemid = "&iitemid & VbCRLF
					strSql = strSql & " 	GROUP BY R.itemid "   & VbCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid " & VbCRLF
					dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||성공[조회]"
				Else
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[조회]"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-CHKSTAT-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaItemEdit(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/base", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

			If iitemid = "2736742" Then
				rw "req : " & strParam
				rw "res : " & iRbody
			End If

			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET R.regitemname = i.itemname " & VbCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_skstoa_regItem R" & VbCrlf
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " WHERE R.itemid = " & iitemid
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||성공[기초정보]"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[기초정보]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-PRICE-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaEditPrice(iitemid, strParam, imustPrice, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode, marginRate
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/goods-price-modify", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message

				If returnCode = "200" Then
					strSql = ""
					strSql = strSql & " SELECT TOP 1 m.margin, d.itemid "
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_ssg_marginItem_master] as m  "
					strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_ssg_marginItem_detail] as d on m.idx = d.midx  "
					strSql = strSql & " WHERE m.isusing = 'Y'  "
					strSql = strSql & " and convert(char(10), getdate(), 120) between m.startDate and m.enddate  "
					strSql = strSql & " and m.mallid = 'skstoa' "
					strSql = strSql & " and d.itemid = '"& iitemid &"' "
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If Not(rsget.EOF or rsget.BOF) Then
						marginRate = rsget("margin")
					Else
						marginRate = 12
					End If
					rsget.Close

				    strSql = ""
	    			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_skstoa_regItem " & VbCRLF
	    			strSql = strSql & "	SET skstoaLastUpdate = GETDATE() " & VbCRLF
	    			strSql = strSql & "	,skstoaPrice = " & imustprice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
					strSql = strSql & " ,setMargin = '"& marginRate &"' " & VbCRLF
	    			strSql = strSql & " WHERE itemid='" & iitemid & "'"& VbCRLF
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||성공[가격]"
				Else
					rw "req : " & strParam
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[가격]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-PRICE-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaEditContentReg(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/describe", false
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
					iErrStr =  "OK||"&iitemid&"||성공[기술서(수정)]"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[기술서(수정)]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-EDITContent-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaGosiEdit(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/offer", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If (returnCode <> "200") AND (iMessage <> "이미 등록된 정보고시건 입니다.") Then
					iErrStr = "ERR"
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-EDITContent-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaQtyEdit(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/inplanqty-modify", false
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
					If (session("ssBctID")="kjy8517") Then
						rw "---재고수정---"
						rw "REQ : <textarea cols=40 rows=10>"&strParam&"</textarea>"
						rw "RES : <textarea cols=40 rows=10>"&BinaryToText(objXML.ResponseBody,"utf-8")&"</textarea>"
						rw "-------------"
					End If
					iErrStr = "ERR"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-EDITQty-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaOptSellyn(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/sales-no-goods", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
'		response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message

				If Instr(iMessage, "판매구분 값이 이전과 동일합니다") > 0 Then
					returnCode = "200"
				End If

				If (returnCode <> "200")  Then
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-EDITOptSellyn-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaOptAdd(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/goodsdt", false
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
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-EDITADDOPT-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaEditImage(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/image-url", false
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
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET R.regimageName = i.basicImage " & VbCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_skstoa_regItem R" & VbCrlf
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " WHERE R.itemid = " & iitemid
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||성공[이미지(수정)]"
					rw "req : " & strParam
					rw "res : " & BinaryToText(objXML.ResponseBody,"utf-8")
				Else
					If InStr(iMessage, "존재하지 않거나 로드가 불가능") Then
						strSql = ""
						strSql = strSql & " UPDATE db_etcmall.dbo.tbl_skstoa_regItem " & VbCrlf
						strSql = strSql & " SET skstoalastupdate = getdate()" & VbCrlf
						strSql = strSql & " ,accFailCNT=0" & VbCrlf
						strSql = strSql & " ,skstoaSellYn = 'N'" & VbCRLF
						strSql = strSql & " WHERE itemid = " & iitemid
						dbget.execute strSql

						strSql = ""
						strSql = strSql & " IF NOT Exists(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE itemid='"&iitemid&"' and mallgubun = '"&CMALLNAME&"') "
						strSql = strSql & "  BEGIN "
						strSql = strSql & "  	INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid(itemid, mallgubun, bigo) VALUES('"&iitemid&"','"&CMALLNAME&"', '존재하지 않거나 이미지 오류') "
						strSql = strSql & "  END "
						dbget.Execute strSql
						iErrStr = "ERR||"&iitemid&"||판매중지[이미지(수정)]/관리자 종료처리"
					Else
						rw "req : " & strParam
						rw "res : " & BinaryToText(objXML.ResponseBody,"utf-8")
						iErrStr = "ERR||"&iitemid&"||"&iMessage&"[이미지(수정)]"
					End If
				End If
			Set strObj = nothing
		Else
			rw "req : " & strParam
			rw "res : " & BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-EDITContent-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnSkstoaEditCert(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, prdNo, strObj, returnCode
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/goods/goods-kcinfo-modify", false
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
					iErrStr =  "OK||"&iitemid&"||성공[인증정보(수정)]"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"[인증정보(수정)]"
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
			iErrStr = "ERR||"&iitemid&"||skstoa 결과 분석 중에 오류가 발생했습니다.[ERR-EDITCert-001]"
		End If
	Set objXML= nothing
End Function

Public Function fnGetCommonCodeList(iinterfaceId)
    Dim objXML, iRbody, strObj, returnCode, datalist, i, addReqUrl, addReqParam, groupList, iCode, iName
	addReqParam = "linkCode="&skstoalinkCode&"&entpCode="&skstoaentpCode&"&entpId="&skstoaentpId&"&entpPass="&skstoaentpPass
	Select Case iinterfaceId
		Case "IF_API_00_001"	addReqUrl = "/partner/code/md-list"						'MD 리스트
		' Case "IF_API_00_002"	addReqUrl = "/partner/code/goods-lgroup-list"			'상품 대분류 조회
		' Case "IF_API_00_003"	addReqUrl = "/partner/code/goods-mgroup-list"			'상품 중분류 조회
		' Case "IF_API_00_004"	addReqUrl = "/partner/code/goods-sgroup-list"			'상품 소분류 조회
		' Case "IF_API_00_005"	addReqUrl = "/partner/code/goods-dgroup-list"			'상품 세분류 조회
		Case "IF_API_00_006"	addReqUrl = "/partner/code/color-group-code-list"		'단품색상그룹 조회
		Case "IF_API_00_007"	addReqUrl = "/partner/code/size-group-code-list"		'단품크기그룹 조회
		Case "IF_API_00_008"	addReqUrl = "/partner/code/form-group-code-list"		'단품형태그룹 조회
		Case "IF_API_00_009"	addReqUrl = "/partner/code/pattern-group-code-list"		'단품무늬그룹 조회
		Case "IF_API_00_010"
			addReqUrl = "/partner/code/color-code-list"									'단품기초코드(색상) 조회
			addReqParam = addReqParam & "&cspfGroup="
		Case "IF_API_00_011"
			addReqUrl = "/partner/code/size-code-list"									'단품기초코드(크기) 조회
			addReqParam = addReqParam & "&cspfGroup="
		Case "IF_API_00_012"
			addReqUrl = "/partner/code/form-code-list"									'단품기초코드(형태) 조회
			addReqParam = addReqParam & "&cspfGroup="
		Case "IF_API_00_013"
			addReqUrl = "/partner/code/pattern-code-list"								'단품기초코드(무늬) 조회
			addReqParam = addReqParam & "&cspfGroup="
		Case "IF_API_00_014"	addReqUrl = "/partner/code/buy-method-list"				'매입방법 조회
		Case "IF_API_00_015"	addReqUrl = "/partner/code/brand-list"					'브랜드 조회		
		Case "IF_API_00_016"	addReqUrl = "/partner/code/describe-code-list"			'기술서항목 조회	
		Case "IF_API_00_017"
			addReqUrl = "/partner/code/entpman-list"									'업체 담당자 조회	
			addReqParam = addReqParam & "&entpManGb=40"									'구분별 담당자 목록 조회 10 : 상품담당자, 20 : 회수담당자, 30 : 출고담당자, 40 : 회계담당자"
		Case "IF_API_00_018"	addReqUrl = "/partner/code/origin-list"					'원산지 조회		
		Case "IF_API_00_019"	addReqUrl = "/partner/code/make-company-list"			'제조업체 조회		
		Case "IF_API_00_020"	addReqUrl = "/partner/code/order-media-list"			'주문매체 조회		
		Case "IF_API_00_021"	addReqUrl = "/partner/code/nosales-reason-code-list"	'판매불가 사유 조회	
		Case "IF_API_00_022"	addReqUrl = "/partner/code/goods-offer-code-list"		'상품정보제공고시 상품유형 조회
		Case "IF_API_00_023"	addReqUrl = "/partner/code/goods-offer-list"			'상품정보제공고시 항목 조회
		Case "IF_API_00_024"	addReqUrl = "/partner/code/delivery-company-list"		'배송사 조회
		Case "IF_API_00_025"	addReqUrl = "/partner/code/shipping-policy-list"		'고객 배송비정책 목록 조회..IF_API_00_030 이용해서 B001로 가입력함
		Case "IF_API_00_026"	addReqUrl = "/partner/code/mdkind-list"					'MD분류리스트
		Case "IF_API_00_027"	addReqUrl = "/partner/code/model-list"					'모델명 조회
	End Select

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", skstoaAPIURL & addReqUrl & "?" & addReqParam, false
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
							strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_skstoa_makeCompanyCode] "
							dbget.Execute(strSql)

							For i=0 to groupList.length-1
								iCode		= groupList.get(i).makeCompanyCode		'제조업체 코드
								iName		= groupList.get(i).makeCompanyName		'제조업체 명

								strSql = ""
								strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_skstoa_makeCompanyCode] (makeCompanyCode, makeCompanyName) VALUES "
								strSql = strSql & " ('"&iCode&"', '"&html2db(iName)&"') "
								dbget.Execute(strSql)
								If (i mod 1000) = 0 Then
									response.flush
								End If
							Next
							rw groupList.length & " 건 등록"
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
							strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_skstoa_originCode] "
							dbget.Execute(strSql)

							For i=0 to groupList.length-1
								iCode		= groupList.get(i).originCode		'원산지 코드
								iName		= groupList.get(i).originName		'원산지 명

								strSql = ""
								strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_skstoa_originCode] (originCode, originName) VALUES "
								strSql = strSql & " ('"&iCode&"', '"&html2db(iName)&"') "
								dbget.Execute(strSql)
								If (i mod 1000) = 0 Then
									response.flush
								End If
							Next
							rw groupList.length & " 건 등록"
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
							strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_skstoa_brandCode] "
							dbget.Execute(strSql)

							For i=0 to groupList.length-1
								iCode		= groupList.get(i).brandCode		'브랜드 코드
								iName		= groupList.get(i).brandName		'브랜드 명칭

								strSql = ""
								strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_skstoa_brandCode] (brandCode, brandName) VALUES "
								strSql = strSql & " ('"&iCode&"', '"&html2db(iName)&"') "
								dbget.Execute(strSql)
								If (i mod 1000) = 0 Then
									response.flush
								End If
							Next
							rw groupList.length & " 건 등록"
						Set groupList = nothing
					End If
				Set strObj = nothing
			End If
		ElseIf iinterfaceId = "IF_API_00_022" Then
			'response.write BinaryToText(objXML.ResponseBody,"utf-8")
			If objXML.Status = "200" OR objXML.Status = "201" Then
				iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
				Set strObj = JSON.parse(iRbody)
					returnCode		= strObj.code
					If returnCode = "200" Then
						Set groupList = strObj.offerTypeList
							For i=0 to groupList.length-1
								iCode		= groupList.get(i).typeCode		'상품유형코드
								iName		= groupList.get(i).typeName		'상품유형명

								strSql = ""
								strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_skstoa_infocd] "
								strSql = strSql & " SET typeCode = '"& iCode &"' "
								strSql = strSql & " WHERE typeName = '"& iName &"' "
								dbget.Execute(strSql)
							Next
							rw groupList.length & " 건 수정"
						Set groupList = nothing
					End If
				Set strObj = nothing
			End If
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
		End If
	Set objXML= nothing
End Function

'상품 세분류 조회
Public Function fnGetGoodsDgroupList()
    Dim objXML, iRbody, strObj, returnCode, i, strSql
	Dim groupList, lgroup,	mgroup,	sgroup,	dgroup,	tgroup,	lgroupName,mgroupName,sgroupName,dgroupName,tgroupName
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", skstoaAPIURL & "/partner/code/goods-dgroup-list?linkCode="&skstoalinkCode&"&entpCode="&skstoaentpCode&"&entpId="&skstoaentpId&"&entpPass="&skstoaentpPass&"", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
'		response.write BinaryToText(objXML.ResponseBody,"utf-8")
'		response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				If returnCode = "200" Then
					Set groupList = strObj.groupList
						strSql = ""
						strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_skstoa_category] "
						dbget.Execute(strSql)

						For i=0 to groupList.length-1
							lgroup		= groupList.get(i).lgroup		'대분류 코드
							mgroup		= groupList.get(i).mgroup		'중분류 코드
							sgroup		= groupList.get(i).sgroup		'소분류 코드
							dgroup		= groupList.get(i).dgroup		'세분류 코드
							lgroupName	= groupList.get(i).lgroupName	'대분류명
							mgroupName	= groupList.get(i).mgroupName	'중분류명
							sgroupName	= groupList.get(i).sgroupName	'소분류명
							dgroupName	= groupList.get(i).dgroupName	'세분류명

							strSql = ""
							strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_skstoa_Category_Ins] '"&lgroup&"', '"&mgroup&"', '"&sgroup&"', '"&dgroup&"' " & VBCRLF
							strSql = strSql & " ,'"&lgroupName&"' ,'"&mgroupName&"' ,'"&sgroupName&"' ,'"&dgroupName&"' "
							dbget.Execute(strSql)
						Next
						rw "SKSTOA 카테고리 " & groupList.length & " 건 등록"
					Set groupList = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'상품정보제공고시 항목 조회
Public Function fnGetOfferList()
    Dim objXML, iRbody, strObj, returnCode, i, strSql
	Dim offerList, offerCode, offerName, typeName
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", skstoaAPIURL & "/partner/code/goods-offer-list?linkCode="&skstoalinkCode&"&entpCode="&skstoaentpCode&"&entpId="&skstoaentpId&"&entpPass="&skstoaentpPass&"", false
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
						strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_skstoa_infocd] "
						dbget.Execute(strSql)
						For i=0 to offerList.length-1
							offerCode		= offerList.get(i).offerCode				'항목코드
							offerName		= html2db(offerList.get(i).offerName)		'항목명
							typeName		= offerList.get(i).typeName					'상품유형명

							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_skstoa_infocd]  (offerCode, offerName, typeCode, typeName) VALUES "
							strSql = strSql & " ('"& offerCode &"', '"& offerName &"', '', '"& typeName &"') "
							dbget.Execute(strSql)
						Next
						rw "상품정보제공고시 항목 " & offerList.length & " 건 등록"
					Set offerList = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

Function ArrErrStrInfo(iaction, iitemid, ierrVendorItemId)
	Dim ErrStrComma, strSql
	If iaction = "REGOpt" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||실패[옵션등록] " & ErrStrComma
		Else
			strSql = ""
			strSql = strSql & " UPDATE R" & VbCRLF
			strSql = strSql & "	Set R.reglevel = 3 " & VbCRLF
			strSql = strSql & "	From db_etcmall.dbo.tbl_skstoa_regItem R"& VbCRLF
			strSql = strSql & " Where R.itemid = '" & iitemid & "'"
			dbget.Execute(strSql)
			ArrErrStrInfo =  "OK||"&iitemid&"||성공[옵션등록]"
		End If
	ElseIf iaction = "REGGosi" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||실패[고시정보] " & ErrStrComma
		Else
			strSql = ""
			strSql = strSql & " UPDATE R" & VbCRLF
			strSql = strSql & "	Set R.reglevel = 5 " & VbCRLF
			strSql = strSql & "	From db_etcmall.dbo.tbl_skstoa_regItem R"& VbCRLF
			strSql = strSql & " Where R.itemid = '" & iitemid & "'"
			dbget.Execute(strSql)
			ArrErrStrInfo =  "OK||"&iitemid&"||성공[고시정보]"
		End If
	ElseIf iaction = "EDITGosi" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||실패[고시정보] " & ErrStrComma
		Else
			ArrErrStrInfo =  "OK||"&iitemid&"||성공[고시정보]"
		End If
	ElseIf iaction = "EDITQTY" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||실패[재고수정] " & ErrStrComma
		Else
			ArrErrStrInfo =  "OK||"&iitemid&"||성공[재고수정]"
		End If
	ElseIf iaction = "EDITSTAT" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||실패[옵션상태] " & ErrStrComma
		Else
			ArrErrStrInfo =  "OK||"&iitemid&"||성공[옵션상태]"
		End If
	ElseIf iaction = "EDITADDOPT" Then
		If Right(ierrVendorItemId,1) = "," Then
			ErrStrComma = Left(ierrVendorItemId, Len(ierrVendorItemId) - 1)
		End If
		If ierrVendorItemId <> "" Then
			ArrErrStrInfo = "ERR||"&iitemid&"||실패[옵션추가] " & ErrStrComma
		Else
			ArrErrStrInfo =  "OK||"&iitemid&"||성공[옵션추가]"
		End If
	End If
End Function
%>