<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'상품 등록
Public Function fnSsgItemReg(iitemid, istrParam, byRef iErrStr, imustprice, iimageNm, isetMargin)
    Dim objXML, xmlDOM, strSql, iResult, ssgGoodno, LagrgeNode, i, uitemId, tempUitemId, useYn, baseInvQty
    Dim iRbody, iMessage
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & ssgAPIURL&"/item/0.4/insertItem.ssg", false
		objXML.setRequestHeader "CONTENT-TYPE", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Accept", "application/xml"
		objXML.setRequestHeader "Authorization", ssgApiKey
		objXML.send(istrParam)
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[상품등록] " & Err.Description
			Exit Function
		End If
        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			iResult = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text

			If (session("ssBctID")="kjy8517") Then
				rw "REQ : <textarea cols=40 rows=10>"&istrParam&"</textarea>"
				rw "RES : <textarea cols=40 rows=10>"&objXML.responseText&"</textarea>"
			End If

			If iResult = "SUCCESS" Then
				ssgGoodno = xmlDOM.getElementsByTagName("itemId").Item(0).Text
				strSql = ""
				strSql = strSql & " UPDATE R" & VbCrlf
				strSql = strSql & " SET ssgRegdate = getdate()" & VbCrlf
				If (ssgGoodno <> "") Then
				    strSql = strSql & "	, ssgStatCd = '3'"& VbCRLF					'승인대기
				Else
					strSql = strSql & "	, ssgStatCd = '1'"& VbCRLF					'전송시도
				End If
				strSql = strSql & " ,ssgGoodNo = '" & ssgGoodno & "'" & VbCrlf
				strSql = strSql & " ,ssglastupdate = getdate()"
				strSql = strSql & " ,ssgPrice = '"&imustprice&"' " & VbCrlf
				strSql = strSql & " ,ssgsellYn = 'Y' "& VbCrlf
				strSql = strSql & " ,accFailCNT = 0" & VbCrlf                		'실패회수 초기화
				strSql = strSql & " ,regimageName = '"&iimageNm&"'"& VbCrlf
				strSql = strSql & " ,setMargin = '"& isetMargin &"'"&VbCRLF
				strSql = strSql & " FROM db_etcmall.dbo.tbl_ssg_regitem R" & VbCrlf
				strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
				strSql = strSql & " where R.itemid = " & iitemid
				dbget.execute strSql

				Set LagrgeNode = xmlDOM.SelectNodes("/result/uitems/uitem")
				If Not (LagrgeNode Is Nothing) Then
					For i = 0 To LagrgeNode.length - 1
						uitemId		= LagrgeNode(i).SelectSingleNode("uitemId").Text
						tempUitemId = LagrgeNode(i).SelectSingleNode("tempUitemId").Text
						useYn		= LagrgeNode(i).SelectSingleNode("useYn").Text
						baseInvQty	= LagrgeNode(i).SelectSingleNode("baseInvQty").Text

						strSql = ""
		                strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
			            strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate) "
			            strSql = strSql & " SELECT '"&iitemid&"', itemoption, 'ssg', '"&uitemId&"', optionname, '"&useYn&"', optlimityn, '"&baseInvQty&"', optaddprice, getdate() "
			            strSql = strSql & " FROM db_item.dbo.tbl_item_option "
			            strSql = strSql & " where itemid = '"&iitemid&"' "
			            strSql = strSql & " and itemoption = '"&tempUitemId&"' "
			            dbget.execute strSql
					Next
					strSql = ""
					strSql = strSql & " UPDATE R"   &VbCRLF
					strSql = strSql & " SET regedOptCnt = isNULL(T.CNT,0)"&VbCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_ssg_regitem R"&VbCRLF
					strSql = strSql & " JOIN ("&VbCRLF
					strSql = strSql & " 	SELECT R.itemid,count(*) as CNT "&VbCRLF
					strSql = strSql & " 	FROM db_etcmall.dbo.tbl_ssg_regitem R"&VbCRLF
					strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = 'ssg' and Ro.itemid = "&iitemid& VbCRLF
					strSql = strSql & " 	GROUP BY R.itemid"&VbCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid"
					dbget.Execute strSql
				End If
				Set LagrgeNode = nothing
				iErrStr =  "OK||"&iitemid&"||등록성공(상품등록)"
			Else
				iMessage = replaceMsg(xmlDOM.getElementsByTagName("resultDesc").item(0).text)
				iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상품등록)"
			End If
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'승인목록 조회
Public Function fnSsgStatChk(iitemid, iSsgGoodNo, iErrStr)
    Dim objXML, xmlDOM, strSql, iResult, ssgGoodno, i, ssgStatcd, tenStatcd, tenStatStr
    Dim iRbody, iMessage
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & ssgAPIURL&"/item/0.1/getItemChngDemandList.ssg?itemId="&iSsgGoodNo
		objXML.setRequestHeader "Authorization", ssgApiKey
		objXML.setRequestHeader "Accept", "application/xml"
		objXML.send()
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[승인조회] " & Err.Description
			Exit Function
		End If
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			iResult = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			If iResult = "SUCCESS" Then
				'ssgStatcd = xmlDOM.getElementsByTagName("ssgChngDemndProcStatCd").Item(0).Text		'신세계승인코드 | 00 : 해당사항없음, 10 : MD승인요청, 20 : 승인완료, 30 : MD반려, 40 : CS합의요청, 50 : CS반려
				ssgStatcd = xmlDOM.getElementsByTagName("chngDemndProcStatCd").Item(0).Text		'승인상태 코드 (commCd : I026) | 00 해당 사항 없음, 10 MD 승인 요청, 20 승인 완료, 30 MD 반려, 40 CS 합의 요청, 50 CS 반려
				If (ssgStatcd = "") OR ISNULL(ssgStatCd) Then
					ssgStatcd = "00"
				End If

				If ssgStatcd = "00" Then
					iErrStr = "OK||"&iitemid&"||[승인조회]이미 승인되었습니다. "
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ssg_regitem SET "
					strSql = strSql & " ssgStatCd = '7' "
					strSql = strSql & " ,lastConfirmdate = getdate() "
					strSql = strSql & " WHERE itemid = '"&iitemid&"' "
					dbget.Execute strSql
					Exit Function
				Else
					Select Case ssgStatcd
						Case "10"
							tenStatcd = "3"
							tenStatStr = "승인대기"
						Case "20"
							tenStatcd = "7"
							tenStatStr = "승인완료"
						Case "30"
							tenStatcd = "2"
							tenStatStr = "MD반려"
						Case "40"
							tenStatcd = "3"
							tenStatStr = "CS합의요청"
						Case "50"
							tenStatcd = "2"
							tenStatStr = "CS반려"
					End Select

					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ssg_regitem SET "
					strSql = strSql & " ssgStatCd = '"&tenStatcd&"' "
					strSql = strSql & " ,lastConfirmdate = getdate() "
					strSql = strSql & " WHERE itemid = '"&iitemid&"' "
					dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||성공("&tenStatStr&")"
				End If
			Else
				iMessage = replaceMsg(xmlDOM.getElementsByTagName("resultDesc").item(0).text)
				iErrStr = "ERR||"&iitemid&"||"&iMessage&"(승인조회)"
			End If
		Set xmlDOM = nothing
	Set objXML = nothing
End Function

'상품 수정
Public Function fnSsgItemEdit(iitemid, iSsgGoodNo, iErrStr, istrParam, imustprice, iItemName, ichgSellYn, ichgImageNm)
    Dim objXML, xmlDOM, strSql, iResult, ssgGoodno, LagrgeNode, i, uitemId, tempUitemId, useYn, baseInvQty
    Dim iRbody, iMessage
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & ssgAPIURL&"/item/0.3/updateItem.ssg", false
		objXML.setRequestHeader "CONTENT-TYPE", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Accept", "application/xml"
		objXML.setRequestHeader "Authorization", ssgApiKey
		objXML.send(istrParam)

 If (session("ssBctID")="kjy8517") Then
 	response.write replace(istrParam, "<?xml", "<aaaaaaa")
' 	response.end
 End If

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[상품수정] " & Err.Description
			Exit Function
		End If
        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			iResult = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			If iResult = "SUCCESS" Then
				strSql = ""
				strSql = strSql & " UPDATE R " & VbCrlf
				strSql = strSql & " SET lastEditDate = GETDATE()" & VbCrlf
				strSql = strSql & " ,ssgPrice = '"&imustprice&"'" & VbCrlf
				strSql = strSql & " ,accFailCNT=0" & VbCrlf
				strSql = strSql & " ,ssgSellYn = '" & ichgSellYn & "'" & VbCRLF
				strSql = strSql & " , regitemname = '" & iItemName & "' " & VbCRLF
				If (ichgImageNm <> "N") Then
					strSql = strSql & " ,regimageName='"&ichgImageNm&"'"& VbCrlf
				End If
				strSql = strSql & " from db_etcmall.dbo.tbl_ssg_regitem R" & VbCrlf
				strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
				strSql = strSql & " WHERE R.itemid = " & iitemid
				dbget.execute strSql
				iErrStr =  "OK||"&iitemid&"||성공(상품수정)"
			Else
				iMessage = replaceMsg(xmlDOM.getElementsByTagName("resultDesc").item(0).text)
				iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상품수정)"
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'상품 상태 수정
Public Function fnSsgItemEditSellyn(iitemid, iSsgGoodNo, iErrStr, istrParam, imustprice, iItemName, ichgSellYn, ichgImageNm)
    Dim objXML, xmlDOM, strSql, iResult, ssgGoodno, LagrgeNode, i, uitemId, tempUitemId, useYn, baseInvQty
    Dim iRbody, iMessage, sellStatStr
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & ssgAPIURL&"/item/0.3/updateItem.ssg", false
		objXML.setRequestHeader "CONTENT-TYPE", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Accept", "application/xml"
		objXML.setRequestHeader "Authorization", ssgApiKey
		objXML.send(istrParam)
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[상품상태] " & Err.Description
			Exit Function
		End If
        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			iResult = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			If iResult = "SUCCESS" Then
				If ichgSellYn = "Y" Then
					sellStatStr = "판매중으로 변경"
				ElseIf ichgSellYn = "X" Then
					sellStatStr = "영구중단(삭제처리)"
				Else
					sellStatStr = "품절처리"
				End If

				If ichgSellYn = "X" Then
					strSql = ""
					strSql = strSql &" INSERT INTO [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] " & VBCRLF
					strSql = strSql &" SELECT TOP 1 'ssg', i.itemid, r.ssgGoodNo, r.ssgRegdate, getdate(), r.lastErrStr" & VBCRLF
					strSql = strSql &" FROM db_item.dbo.tbl_item as i " & VBCRLF
					strSql = strSql &" JOIN db_etcmall.dbo.tbl_ssg_regitem as r on i.itemid = r.itemid " & VBCRLF
					strSql = strSql &" WHERE i.itemid = "&iitemid & VBCRLF
					dbget.Execute(strSql)

					strSql = ""
					strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_ssg_regitem " & vbcrlf
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
				Else
					strSql = ""
					strSql = strSql & " UPDATE R " & VbCrlf
					strSql = strSql & " SET ssglastupdate = getdate()" & VbCrlf
					strSql = strSql & " ,accFailCNT=0" & VbCrlf
					strSql = strSql & " ,ssgSellYn = '" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " from db_etcmall.dbo.tbl_ssg_regitem R" & VbCrlf
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " WHERE R.itemid = " & iitemid
					dbget.execute strSql
				End If
				iErrStr =  "OK||"&iitemid&"||성공("&sellStatStr&")"
			Else
				iMessage = replaceMsg(xmlDOM.getElementsByTagName("resultDesc").item(0).text)
				If InStr(iMessage, "판매금지인 상품은 판매상태 수정이 불가능합니다") Then
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ssg_regitem " & VbCrlf
					strSql = strSql & " SET ssglastupdate = getdate()" & VbCrlf
					strSql = strSql & " ,accFailCNT=0" & VbCrlf
					strSql = strSql & " ,ssgSellYn = 'N'" & VbCRLF
					strSql = strSql & " WHERE itemid = " & iitemid
					dbget.execute strSql

 					strSql = ""
					strSql = strSql & " IF NOT Exists(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE itemid='"&iitemid&"' and mallgubun = '"&CMALLNAME&"') "
					strSql = strSql & "  BEGIN "
					strSql = strSql & "  	INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid(itemid, mallgubun, bigo) VALUES('"&iitemid&"','"&CMALLNAME&"', '판매불가 처리됨') "
					strSql = strSql & "  END "
					dbget.Execute strSql
					iErrStr = "OK||"&iitemid&"||판매중지(상태)/관리자 종료처리"
				Else
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상품상태)"
				End IF
			End IF
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

Public Function fnViewItemInfo(iitemid, iSsgGoodNo, iErrStr)
    Dim objXML, xmlDOM, strSql, iResult, ssgGoodno, i, ssgStatcd, tenStatcd, tenStatStr, LargeNode2, j, setMargin, itemNm
    Dim iRbody, iMessage, ioptionTypename, ioptionname, LagrgeNode, uitemId, useYn, baseInvQty, iuitemnm, salesPrcInfoNodes, sellStatCd, optsellyn
	Dim regOptCodeCnt
	'On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & ssgAPIURL&"/item/0.3/viewItem.ssg?itemId="&iSsgGoodNo
		objXML.setRequestHeader "Authorization", ssgApiKey
		objXML.setRequestHeader "Accept", "application/xml"
		objXML.send()
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[상품조회] " & Err.Description
			Exit Function
		End If
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			iResult = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text

			If iResult = "SUCCESS" Then
				itemNm =  xmlDOM.getElementsByTagName("itemNm").Item(0).Text
'If (session("ssBctID")="kjy8517") Then
'	response.write replace(objXML.responseText, "UTF-8","euc-kr")
'	response.end
'End If
'Set LargeNode2 = xmlDOM.SelectNodes("/result/uitemPluralPrcs/uitemPrc")
'For j = 0 To LargeNode2.length - 1
'	rw LargeNode2(j).SelectSingleNode("mrgrt").Text
'Next
'Set LargeNode2 = nothing
				SET salesPrcInfoNodes = xmlDOM.SelectNodes("/result/salesPrcInfos/uitemPrc")
					For j = 0 To salesPrcInfoNodes.length - 1
						setMargin = salesPrcInfoNodes(j).SelectSingleNode("mrgrt").Text
					Next
				SET salesPrcInfoNodes = nothing

				Set LagrgeNode = xmlDOM.SelectNodes("/result/uitems/uitem")
					For i = 0 To LagrgeNode.length - 1
						iuitemnm = ""
						regOptCodeCnt = 0
						uitemId			= LagrgeNode(i).SelectSingleNode("uitemId").Text
						useYn			= LagrgeNode(i).SelectSingleNode("useYn").Text
						baseInvQty		= LagrgeNode(i).SelectSingleNode("baseInvQty").Text
						ioptionTypename = LagrgeNode(i).SelectSingleNode("uitemOptnTypeNm1").Text
						ioptionname		= LagrgeNode(i).SelectSingleNode("uitemOptnNm1").Text

						If Not (LagrgeNode(i).SelectSingleNode("uitemOptnNm2") is nothing) then
						 	ioptionname = ioptionname & "/" & LagrgeNode(i).SelectSingleNode("uitemOptnNm2").Text
						End If

						If Not (LagrgeNode(i).SelectSingleNode("uitemOptnNm3") is nothing) then
						 	ioptionname = ioptionname & "/" & LagrgeNode(i).SelectSingleNode("uitemOptnNm3").Text
						End If

						'복합명 저장
						iuitemnm		= LagrgeNode(i).SelectSingleNode("uitemNm").Text

						'옵션1번째 명과 복합명이 같다면 복합명을 옵션명 1로 교체
						If itemNm = iuitemnm Then	'2019-02-26 17:51 김진영 수정
							iuitemnm = LagrgeNode(i).SelectSingleNode("uitemOptnNm1").Text
						End If

						sellStatCd		= LagrgeNode(i).SelectSingleNode("sellStatCd").Text
						If iuitemnm <> "" Then
							ioptionname = replace(iuitemnm, "/", ",")
							ioptionname = replace(ioptionname, "#", "")
							ioptionname = replace(ioptionname, "'", "")
						End If

						If (useYn = "N") OR (sellStatCd = "80") Then
							optsellyn = "N"
						Else
							optsellyn = "Y"
						End If

						strSql = ""
						strSql = strSql & " SELECT COUNT(*) as cnt "
						strSql = strSql & " FROM db_item.dbo.tbl_OutMall_regedoption "
						strSql = strSql & " WHERE itemid = '"&iitemid&"' "
						strSql = strSql & " and outmallOptCode='"&uitemId&"'"
						strSql = strSql & " and mallid='"&CMALLNAME&"'"
						rsget.CursorLocation = adUseClient
						rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
							regOptCodeCnt = rsget("cnt")
						rsget.Close

						If regOptCodeCnt > 0 Then	'2019-05-13 16:33 김진영 If 조건 추가
							strSql = ""
							strSql = strSql & " UPDATE db_item.dbo.tbl_OutMall_regedoption SET "
							strSql = strSql & " outmallsellyn='"&optsellyn&"'"
							strSql = strSql & " , checkdate = getdate() "
							strSql = strSql & " , outmalllimitno='"&baseInvQty&"'"
							strSql = strSql & " WHERE itemid = '"&iitemid&"' and outmallOptCode='"&uitemId&"'"
							strSql = strSql & " and mallid='"&CMALLNAME&"'"
							dbget.Execute strSql
						Else
							strSql = ""
							strSql = strSql & " IF Exists(SELECT * FROM db_item.dbo.tbl_OutMall_regedoption WHERE itemid="&iitemid&" and mallid = '"&CMALLNAME&"' and replace(replace(replace(outmallOptName,'/',','),'#',''), '''', '') = '"&ioptionname&"' )"
							strSql = strSql & " BEGIN "
							strSql = strSql & " UPDATE db_item.dbo.tbl_OutMall_regedoption SET "
							strSql = strSql & " outmallsellyn='"&optsellyn&"'"
							strSql = strSql & " , outmallOptName='"&html2DB(ioptionname)&"'"
							strSql = strSql & " , checkdate = getdate() "
							strSql = strSql & " , outmallOptCode='"&uitemId&"'"
							strSql = strSql & " , outmalllimitno='"&baseInvQty&"'"
							strSql = strSql & " WHERE itemid = '"&iitemid&"' and replace(replace(replace(outmallOptName,'/',','),'#',''), '''', '') = '"&ioptionname&"' "
							strSql = strSql & " and mallid='"&CMALLNAME&"'"
							strSql = strSql & " END ELSE "
							strSql = strSql & " BEGIN "
							strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
							strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate) "
							strSql = strSql & " SELECT '"&iitemid&"', o.itemoption, 'ssg', '"&uitemId&"', o.optionname, '"&useYn&"', o.optlimityn, '"&baseInvQty&"', o.optaddprice, getdate() "
							strSql = strSql & " FROM db_item.dbo.tbl_item_option o"
							strSql = strSql & " WHERE o.itemid = '"&iitemid&"' "
							strSql = strSql & " and replace(replace(o.optionname,'/',','),'#','') = '"&ioptionname&"' "
							strSql = strSql & " and NOT EXISTS(select 1 from db_item.dbo.tbl_OutMall_regedoption ro where ro.mallid='"&CMALLNAME&"' and ro.itemid='"&iitemid&"' and ro.itemid=o.itemid and ro.itemoption=o.itemoption )"
							strSql = strSql & " END "
							dbget.Execute strSql
						End If
					Next
					strSql = ""
					strSql = strSql & " UPDATE R"   &VbCRLF
					strSql = strSql & " SET regedOptCnt = isNull(T.optSellYCNT,0) "&VbCRLF '' isNULL(T.CNT,0)"&VbCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_ssg_regitem R"&VbCRLF
					strSql = strSql & " JOIN ("&VbCRLF
					strSql = strSql & " 	SELECT R.itemid,count(*) as CNT "&VbCRLF
					strSql = strSql & " 	,sum(CASE WHEN [outmallSellyn]='Y' and (outmalllimityn='N' or (outmalllimityn='Y' and outmalllimitno>0)) THEN 1 ELSE 0 END) as optSellYCNT"&VbCRLF
					strSql = strSql & " 	FROM db_etcmall.dbo.tbl_ssg_regitem R"&VbCRLF
					strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = 'ssg' and Ro.itemid = "&iitemid& VbCRLF
					strSql = strSql & " 	GROUP BY R.itemid"&VbCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid"
					dbget.Execute strSql

					strSql = ""
					strSql = strSql & " UPDATE R"   &VbCRLF
					strSql = strSql & " SET setMargin = '"& setMargin &"'"&VbCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_ssg_regitem R"&VbCRLF
					strSql = strSql & " WHERE R.itemid = "&iitemid
					dbget.Execute strSql
				Set LagrgeNode = nothing
				iErrStr =  "OK||"&iitemid&"||성공(상품조회)"
			Else
				iMessage = replaceMsg(xmlDOM.getElementsByTagName("resultDesc").item(0).text)
				iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상품조회)"
			End If
		Set xmlDOM = nothing
	Set objXML = nothing
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################
'	1. 관리카테고리조회
'		1. http://eapi.ssgadm.com/venInfo/0.2/listStdCtgKeyPath.ssg or ?siteNo=6004
'	1. 전시카테고리 조회
'
'		1. http://eapi.ssgadm.com/common/0.2/listDispCtg.ssg?stdCtgDclsId=4000002194
'	2. 배송비 정책 조회 - 현재 무료배송인듯함.
'
'		1. http://eapi.ssgadm.com//venInfo/0.1/listShppcstPlcy.ssg
'	3. 배송비 정책 등록 - 어드민에서 해도 될듯
'
'		1. 업체정보관리>입점업체정보 관리 3만원이상 무료 2,500 등 설정
'		2. Url : /venInfo/{version}/insertShppcstPlcy.ssg
'	4. 배송(출하지) 등록 - 우선 텐배송만
'	5. 상품고시분류조회
'
'		1. /common/{version}/listItemMngPropCls.ssg
'	6. 상품고시 상세항목조회
'
'		1. /common/{version}/listItemMngProp.ssg
'	7. 브랜드코드조회 - 우선 등록된 텐바이텐으로
'
'		1. /venInfo/{version}/listBrand.ssg
'	8. 원산지조회 - 상품등록시 필요?
'
'		1. /common/{version}/listOrplc.ssg
'	9. 금칙어조회 - 일단제외
'
'		1. /venInfo/{version}/getProhibitedWordList.ssg
'	10. 배송메세지조회 - 일단제외
'
'		1. /venInfo/{version}/getVenShppMsgList.ssg
'	11. 공통콛 조회 -  postman에서 사용
'
'		1. /common/{version}/getCommCdDtlc.ssg
'
'
'
'상품조회
'
'	1. 상품목록 조회 - 상품명, 판매상태등
'
'		1.  /item/{version}/getItemList.ssg
'	2. 상품상세조회
'
'		1.  /item/{version}/viewItem.ssg
'	3. 상품등록
'
'		1. /item/{version}/insertItem.ssg
'	4.





'' for SSG ---------------------------------------------------
Public Function fnSsgGosiInfo(igosiClsId)
    Dim objXML, xmlDOM, strSql
    Dim iRbody, LagrgeNode
    Dim itemMngPropClsId, itemMngPropClsNm, itemMngPropId, itemMngPropNm, iptMthdCd, mndtyYn, dispPrioyExpsrYn
	'On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		'objXML.open "GET", "" & ssgAPIURL&"/common/0.1/listItemMngProp.ssg?itemMngPropClsId=0000000037", false
		objXML.open "GET", "" & ssgAPIURL&"/common/0.1/listItemMngProp.ssg?itemMngPropClsId="&igosiClsId, false
		objXML.setRequestHeader "Authorization", ssgApiKey
		objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
		objXML.send()
        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			Set LagrgeNode = xmlDOM.SelectNodes("/result/itemMngProps/itemMngProp")
			If Not (LagrgeNode Is Nothing) Then
				strSql = ""
				strSql = strSql & " DELETE FROM db_temp.[dbo].[tbl_ssg_infocd] WHERE itemMngPropClsId = '"& igosiClsId &"' "
				dbget.Execute(strSql)

				For i = 0 To LagrgeNode.length - 1
					itemMngPropClsId = LagrgeNode(i).SelectSingleNode("itemMngPropClsId").Text
					itemMngPropClsNm = LagrgeNode(i).SelectSingleNode("itemMngPropClsNm").Text
					itemMngPropId = LagrgeNode(i).SelectSingleNode("itemMngPropId").Text
					itemMngPropNm = LagrgeNode(i).SelectSingleNode("itemMngPropNm").Text
					iptMthdCd = LagrgeNode(i).SelectSingleNode("iptMthdCd").Text
					mndtyYn = LagrgeNode(i).SelectSingleNode("mndtyYn").Text
					dispPrioyExpsrYn = LagrgeNode(i).SelectSingleNode("dispPrioyExpsrYn").Text

					strSql = ""
					strSql = strSql & " INSERT INTO db_temp.[dbo].[tbl_ssg_infocd] (itemMngPropClsId, itemMngPropClsNm, itemMngPropId, itemMngPropNm, iptMthdCd, mndtyYn, dispPrioyExpsrYn) "
					strSql = strSql & " VALUES ('"&itemMngPropClsId&"', '"&itemMngPropClsNm&"', '"&itemMngPropId&"', '"&itemMngPropNm&"', '"&iptMthdCd&"', '"&mndtyYn&"', '"&dispPrioyExpsrYn&"') "
					dbget.Execute(strSql)
				Next
				rw "OK"
			End If
			Set LagrgeNode = nothing
			'response.write replace(objXML.responseText, "xml","aaaa")
		Set xmlDOM = nothing
	Set objXML = nothing
End Function

Public Function fnSsgAreaInfo()
    Dim objXML, xmlDOM, strSql
    Dim iRbody, LagrgeNode, areaCode, areaName, iid, iareaName
	'On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & ssgAPIURL&"/common/0.1/listOrplc", false
		objXML.setRequestHeader "Authorization", ssgApiKey
		objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
		objXML.send()
        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
'			response.write objXML.responseText

			Set LagrgeNode = xmlDOM.SelectNodes("/result/orplcs/orplc")
			If Not (LagrgeNode Is Nothing) Then
				For i = 0 To LagrgeNode.length - 1
					If LagrgeNode(i).SelectSingleNode("manufCntryYn").Text = "Y" Then
						areaCode = areaCode & LagrgeNode(i).SelectSingleNode("orplcId").Text & ","
						areaName = areaName & LagrgeNode(i).SelectSingleNode("orplcNm").Text & ","

						iid			= LagrgeNode(i).SelectSingleNode("orplcId").Text
						iareaName	= LagrgeNode(i).SelectSingleNode("orplcNm").Text
						strSql = ""
						strSql = strSql & " IF Exists(SELECT * FROM db_etcmall.dbo.tbl_ssg_sourceAreaCode WHERE id = '"&iid&"') "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	UPDATE db_etcmall.dbo.tbl_ssg_sourceAreaCode SET "
						strSql = strSql & " 	sourcearea ='"&iareaName&"'"
						strSql = strSql & " 	WHERE id = '"&iid&"' "
						strSql = strSql & " END ELSE "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_ssg_sourceAreaCode "
						strSql = strSql & " 	(id, sourcearea) VALUES ('"&iid&"', '"&iareaName&"') "
						strSql = strSql & " END "
						dbget.Execute strSql

						strSql = ""
						strSql = strSql & " IF Exists(SELECT * FROM db_etcmall.dbo.tbl_ssg_sourceAreaCodeMapping WHERE id = '"&iid&"') "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	UPDATE db_etcmall.dbo.tbl_ssg_sourceAreaCodeMapping SET "
						strSql = strSql & " 	sourcearea ='"&iareaName&"'"
						strSql = strSql & " 	WHERE id = '"&iid&"' "
						strSql = strSql & " END ELSE "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_ssg_sourceAreaCodeMapping "
						strSql = strSql & " 	(id, sourcearea) VALUES ('"&iid&"', '"&iareaName&"') "
						strSql = strSql & " END "
						dbget.Execute strSql
					End If

'rw LagrgeNode(i).SelectSingleNode("orplcId").Text
'rw LagrgeNode(i).SelectSingleNode("orplcNm").Text
'rw LagrgeNode(i).SelectSingleNode("orplcYn").Text
'rw "---------"
				Next
				If (Right(areaCode,1)=",") Then areaCode = Left(areaCode,Len(areaCode)-1)
				If (Right(areaName,1)=",") Then areaName = Left(areaName,Len(areaName)-1)
			End If

rw areaCode
rw "--------------"
rw areaName
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing
End Function

Public Function fnSsgGosiSafeInfo()
    Dim objXML, xmlDOM, strSql
    Dim iRbody, LagrgeNode, areaCode, areaName, iid, iareaName
	Dim arrRows, stdCtgDclsId, k
	Dim itemAppePropClsId, itemAppePropClsNm, itemAppePropId, itemAppePropNm, itemAppePropTypeCd, itemAppePropDtlTypeCd, refPropTypeCd, refPropCntt, mndtyYn, prcdAppePropId, prcdAppePropCntt
	'On Error Resume Next
	strSql = "EXEC [db_etcmall].[dbo].[usp_Ten_OutMall_Ssg_setSafeGosi] "   ''SSG 관리카테고리 그룹핑
	rsget.CursorLocation = adUseClient
    rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.Eof) then
        arrRows = rsget.getRows()
    end if
    rsget.Close

    if Not isArray(arrRows) then
        fnSsgGosiSafeInfo = false
        Exit function
    end if

	'' 아래 선행 한 뒤 해야될듯''
	' strSql = ""
	' strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_ssg_mmg_cate_SafeInfo] "
	' dbget.Execute(strSql)

	For k = 0 to Ubound(arrRows,2)
		stdCtgDclsId = arrRows(0,k)
		Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.open "GET", "" & ssgAPIURL&"/attribute/itemAppeCert/getItemAppeCertProps.ssg?stdCtgId="&stdCtgDclsId, false
			objXML.setRequestHeader "Authorization", ssgApiKey
			objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
			objXML.send()
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(objXML.responseText)
'				response.write objXML.responseText
'				response.end
				Set LagrgeNode = xmlDOM.SelectNodes("/result/itemAppeCerts/itemAppeCert")
					If Not (LagrgeNode Is Nothing) Then
						For i = 0 To LagrgeNode.length - 1
							itemAppePropClsId		= LagrgeNode(i).SelectSingleNode("itemAppePropClsId").Text			'상품 인증분류 ID
							itemAppePropClsNm		= LagrgeNode(i).SelectSingleNode("itemAppePropClsNm").Text			'상품 인증분류 명
							itemAppePropId			= LagrgeNode(i).SelectSingleNode("itemAppePropId").Text				'상품 인증분류 항목 ID
							itemAppePropNm			= LagrgeNode(i).SelectSingleNode("itemAppePropNm").Text				'상품 인증분류 항목 명
							itemAppePropTypeCd		= LagrgeNode(i).SelectSingleNode("itemAppePropTypeCd").Text			'상품 인증분류 항목 유형 코드 (I321)
							itemAppePropDtlTypeCd	= LagrgeNode(i).SelectSingleNode("itemAppePropDtlTypeCd").Text		'상품 인증분류 항목 상세 유형 코드 (I322)
'							refPropTypeCd			= LagrgeNode(i).SelectSingleNode("refPropTypeCd").Text				'참고항목유형코드
'							refPropCntt				= LagrgeNode(i).SelectSingleNode("refPropCntt").Text				'참고항목 값
							mndtyYn					= LagrgeNode(i).SelectSingleNode("mndtyYn").Text					'필수여부
							If LagrgeNode(i).getElementsByTagName("prcdAppePropId").length > 0 then
								prcdAppePropId			= LagrgeNode(i).SelectSingleNode("prcdAppePropId").Text				'선행항목 ID
							End If

							If LagrgeNode(i).getElementsByTagName("prcdAppePropCntt").length > 0 then
								prcdAppePropCntt		= LagrgeNode(i).SelectSingleNode("prcdAppePropCntt").Text			'선행항목 값
							End If

							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_ssg_mmg_cate_SafeInfo] "
							strSql = strSql & " ([itemAppePropClsId], [itemAppePropClsNm], [itemAppePropId], [itemAppePropNm], [itemAppePropTypeCd], [itemAppePropDtlTypeCd], [mndtyYn], [prcdAppePropId], [prcdAppePropCntt], [stdCtgDclsId] ) "
							strSql = strSql & " VALUES ( "
							strSql = strSql & " '"& itemAppePropClsId &"', '"& itemAppePropClsNm &"', '"& itemAppePropId &"', '"& itemAppePropNm &"', '"& itemAppePropTypeCd &"', '"& itemAppePropDtlTypeCd &"', '"& mndtyYn &"', '"& prcdAppePropId &"', '"& prcdAppePropCntt &"', '"& stdCtgDclsId &"' "
							strSql = strSql & " )"
							dbget.Execute(strSql)
						Next
					End If
				Set LagrgeNode = nothing
			Set xmlDOM = nothing
		Set objXML = nothing
	Next
	rw "OK"
End Function

'' 관리 카테고리 가져오기
public function getSsgMmgCateList()
    Dim objXML, xmlDOM, strSql  '', goodsCd, iResult, iComment
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim siteno, sitenm
    Dim buyFrmCd,stdCtgGrpCd,stdCtgGrpNm,stdCtgLclsId,stdCtgLclsNm,stdCtgMclsId,stdCtgMclsNm,stdCtgSclsId,stdCtgSclsNm,stdCtgDclsId,stdCtgDclsNm
	Dim itemMngPropClsId,itemMngPropClsNm,chldCertTgtYn,safeCertTgtYn,elecCertTgtYn,harmCertTgtYn,txnPermitDivCd, txnPermitDivNm
	'On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & ssgAPIURL&"/venInfo/0.2/listStdCtgKeyPath.ssg"
		objXML.setRequestHeader "Authorization", ssgApiKey
		objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
		objXML.send()
        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			' response.write replace(objXML.responseText, "UTF-8","euc-kr")
			' response.end
			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text

			''rw ssgresultCode&","&ssgresultMessage&","&ssgresultDesc
			Set LagrgeNode = xmlDOM.SelectNodes("/result/stdctgs/stdctg")
				If Not (LagrgeNode Is Nothing) Then
					strSql = "exec db_etcmall.dbo.usp_Ten_OutMall_Ssg_regmmgCateDel"
					dbget.Execute(strSql)
					For i = 0 To LagrgeNode.length - 1
						siteno = LagrgeNode(i).SelectSingleNode("siteNo").Text
						sitenm = LagrgeNode(i).SelectSingleNode("siteNm").Text
						buyFrmCd = LagrgeNode(i).SelectSingleNode("buyFrmCd").Text
						stdCtgGrpCd = LagrgeNode(i).SelectSingleNode("stdCtgGrpCd").Text        ''모바일전용 등
						stdCtgGrpNm = LagrgeNode(i).SelectSingleNode("stdCtgGrpNm").Text
						stdCtgLclsId = LagrgeNode(i).SelectSingleNode("stdCtgLclsId").Text      '' 대분류
						stdCtgLclsNm = LagrgeNode(i).SelectSingleNode("stdCtgLclsNm").Text
						stdCtgMclsId = LagrgeNode(i).SelectSingleNode("stdCtgMclsId").Text      '' 중분류
						stdCtgMclsNm = LagrgeNode(i).SelectSingleNode("stdCtgMclsNm").Text
						stdCtgSclsId = LagrgeNode(i).SelectSingleNode("stdCtgSclsId").Text      '' 소분류
						stdCtgSclsNm = LagrgeNode(i).SelectSingleNode("stdCtgSclsNm").Text
						stdCtgDclsId = LagrgeNode(i).SelectSingleNode("stdCtgDclsId").Text      '' 세분류?
						stdCtgDclsNm = LagrgeNode(i).SelectSingleNode("stdCtgDclsNm").Text

					If LagrgeNode(i).SelectSingleNode("itemMngPropClsId") is nothing Then
						itemMngPropClsId = ""
					Else
						itemMngPropClsId = LagrgeNode(i).SelectSingleNode("itemMngPropClsId").Text  ''상품 고시항목 분류 ID
					End If

					If LagrgeNode(i).SelectSingleNode("itemMngPropClsNm") is nothing Then
						itemMngPropClsNm = ""
					Else
						itemMngPropClsNm = LagrgeNode(i).SelectSingleNode("itemMngPropClsNm").Text  ''상품 고시항목 분류 명
					End if

						' chldCertTgtYn = LagrgeNode(i).SelectSingleNode("chldCertTgtYn").Text    '' 어린이안전인증 (Y/N/X)
						' safeCertTgtYn = LagrgeNode(i).SelectSingleNode("safeCertTgtYn").Text    '' 안전인증
						' elecCertTgtYn = LagrgeNode(i).SelectSingleNode("elecCertTgtYn").Text    '' 전기안전인증
						' harmCertTgtYn = LagrgeNode(i).SelectSingleNode("harmCertTgtYn").Text    '' 위해우려제품 표시대상

						chldCertTgtYn = ""	'2020-11-18 SSG에서 항목제거함
						safeCertTgtYn = ""	'2020-11-18 SSG에서 항목제거함
						elecCertTgtYn = ""	'2020-11-18 SSG에서 항목제거함
						harmCertTgtYn = ""	'2020-11-18 SSG에서 항목제거함

						txnPermitDivCd = LagrgeNode(i).SelectSingleNode("txnPermitDivCd").Text    '' 과세 허용 구분코드
						txnPermitDivNm = LagrgeNode(i).SelectSingleNode("txnPermitDivNm").Text    '' 과세 허용 구분명

						'rw siteno&","&sitenm&","&buyFrmCd&","&stdCtgGrpCd&","&stdCtgGrpNm&","&stdCtgLclsId&","
						'rw stdCtgLclsNm&","&stdCtgMclsId&","&stdCtgMclsNm&","&stdCtgSclsId&","&stdCtgSclsNm&","
						'rw stdCtgDclsId&","&stdCtgDclsNm&","&itemMngPropClsId&","&itemMngPropClsNm&","&chldCertTgtYn&","
						'rw safeCertTgtYn&","&elecCertTgtYn&","& harmCertTgtYn&","& txnPermitDivCd&","&txnPermitDivNm

                        strSql = "exec db_etcmall.dbo.usp_Ten_OutMall_Ssg_regmmgCate '"&siteno&"','"&sitenm&"','"&buyFrmCd&"','"&stdCtgGrpCd&"','"&stdCtgGrpNm&"','"&stdCtgLclsId&"','"& stdCtgLclsNm&"','"&stdCtgMclsId&"','"&stdCtgMclsNm&"','"&stdCtgSclsId&"','"&stdCtgSclsNm&"','"& stdCtgDclsId&"','"&stdCtgDclsNm&"','"&itemMngPropClsId&"','"&itemMngPropClsNm&"','"&chldCertTgtYn&"','"& safeCertTgtYn&"','"&elecCertTgtYn&"','"& harmCertTgtYn&"','"& txnPermitDivCd&"','"&txnPermitDivNm&"'"
'rw strSql
						dbget.Execute(strSql)
					Next
				End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing
end function

public function getSsgDispCateListALL()
    Dim objXML, xmlDOM, strSql ,i, k
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim stdCtgDclsId, stdCtgDclsNm, stdsiteno
    Dim siteno, dispCtgClsCd,dispCtgClsNm,dispCtgLvl,dispCtgLclsId,dispCtgLclsNm,dispCtgMclsId,dispCtgMclsNm,dispCtgSclsId,dispCtgSclsNm,dispCtgDclsId,dispCtgDclsNm,dispCtgSdclsId,dispCtgSdclsNm
	strSql = "exec db_etcmall.dbo.usp_Ten_OutMall_Ssg_getMmgDclCateList"   ''관리카테고리 세분류 목록
	rsget.CursorLocation = adUseClient
    rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.Eof) then
        arrRows = rsget.getRows()
    end if
    rsget.Close

    if Not isArray(arrRows) then
        getSsgDispCateListALL = false
        Exit function
    end if

    for k=0 to Ubound(arrRows,2)
        stdCtgDclsId = arrRows(0,k)
        stdCtgDclsNm = arrRows(1,k)
        stdsiteno      = arrRows(2,k)
	'On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & ssgAPIURL&"/common/0.2/listDispCtg.ssg?stdCtgDclsId="&stdCtgDclsId   ''&"&siteNo="&stdsiteno
		objXML.setRequestHeader "Authorization", ssgApiKey
		objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
		objXML.send()
        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text

			''rw ssgresultCode&","&ssgresultMessage&","&ssgresultDesc
			Set LagrgeNode = xmlDOM.SelectNodes("/result/dispCtgs/ctg")
				If Not (LagrgeNode Is Nothing) Then
					strSql = "exec db_etcmall.dbo.usp_Ten_OutMall_Ssg_setDispCateUsingNo '"&stdCtgDclsId&"'"
					dbget.Execute(strSql)


					For i = 0 To LagrgeNode.length - 1
						siteno = LagrgeNode(i).SelectSingleNode("siteNo").Text
						dispCtgClsCd = LagrgeNode(i).SelectSingleNode("dispCtgClsCd").Text        ''10 메인매장, 20전문매장
						dispCtgClsNm = LagrgeNode(i).SelectSingleNode("dispCtgClsNm").Text      '' 전시카테고리 분류명

						dispCtgLvl   = LagrgeNode(i).SelectSingleNode("dispCtgLvl").Text
						dispCtgLclsId = LagrgeNode(i).SelectSingleNode("dispCtgLclsId").Text      '' 대분류
						dispCtgLclsNm = LagrgeNode(i).SelectSingleNode("dispCtgLclsNm").Text
						dispCtgMclsId = LagrgeNode(i).SelectSingleNode("dispCtgMclsId").Text      '' 중분류
						dispCtgMclsNm = LagrgeNode(i).SelectSingleNode("dispCtgMclsNm").Text
						dispCtgSclsId = LagrgeNode(i).SelectSingleNode("dispCtgSclsId").Text      '' 소분류
						dispCtgSclsNm = LagrgeNode(i).SelectSingleNode("dispCtgSclsNm").Text

						dispCtgDclsId = ""
						dispCtgDclsNm = ""
						if (dispCtgLvl>3) then
						    dispCtgDclsId = LagrgeNode(i).SelectSingleNode("dispCtgDclsId").Text      '' 세분류
						    dispCtgDclsNm = LagrgeNode(i).SelectSingleNode("dispCtgDclsNm").Text
						end if

						dispCtgSdclsId = ""
						dispCtgSdclsNm = ""
						if (dispCtgLvl>4) then
						    dispCtgSdclsId = LagrgeNode(i).SelectSingleNode("dispCtgSdclsId").Text      '' 세세분류
					    	dispCtgSdclsNm = LagrgeNode(i).SelectSingleNode("dispCtgSdclsNm").Text
					    end if

						'rw siteno&","&sitenm&","&buyFrmCd&","&stdCtgGrpCd&","&stdCtgGrpNm&","&stdCtgLclsId&","
						'rw stdCtgLclsNm&","&stdCtgMclsId&","&stdCtgMclsNm&","&stdCtgSclsId&","&stdCtgSclsNm&","
						'rw stdCtgDclsId&","&stdCtgDclsNm&","&itemMngPropClsId&","&itemMngPropClsNm&","&chldCertTgtYn&","
						'rw safeCertTgtYn&","&elecCertTgtYn&","& harmCertTgtYn&","& txnPermitDivCd&","&txnPermitDivNm

	                    if (siteno="6004") or (siteno="6005") or (siteno="6001") then '' 신세계/SSG/이마트몰
                            strSql = "exec db_etcmall.dbo.usp_Ten_OutMall_Ssg_regDispCate '"&siteno&"','"&stdCtgDclsId&"','"&dispCtgLvl&"','"& dispCtgClsCd&"','"&dispCtgClsNm&"','"&dispCtgLclsId&"','"&dispCtgLclsNm&"','"&dispCtgMclsId&"','"&dispCtgMclsNm&"','"&dispCtgSclsId&"','"&dispCtgSclsNm&"','"&dispCtgDclsId&"','"&dispCtgDclsNm&"','"&dispCtgSdclsId&"','"&dispCtgSdclsNm&"'"
    						dbget.Execute(strSql)
					    end if
					Next
				End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing
	next
end function

Public Function fnSsgDispCategoryGet(isiNo)
    Dim objXML, xmlDOM, i, k, LagrgeNode, siteNo, strSql
    Dim dispCtgId, dispCtgNm, dispCtgPathNm, dispYn

	If isiNo = "" Then
		siteNo = "6005"	'6001 : 이마트몰, 6004 : 신세계몰, 6005 : SSG
	Else
		siteNo = isiNo
	End If

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & ssgAPIURL&"/common/0.1/displayCategory.ssg?siteNo="&siteNo&"&pageSize=1000000"
		objXML.setRequestHeader "Authorization", ssgApiKey
		objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
		objXML.send()
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			'rw replace(objXML.responseText, "xml", "xxxcml")
			Set LagrgeNode = xmlDOM.SelectNodes("/result/displayCategorys/category")
				If Not (LagrgeNode Is Nothing) Then
					strSql = ""
					strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_ssg_Newcategory] WHERE siteNo = '"&siteNo&"' "
					dbget.Execute(strSql)

					For i = 0 To LagrgeNode.length - 1
						If LagrgeNode(i).SelectSingleNode("dispCtgLastLvlYn").Text = "Y" Then
							dispCtgId = LagrgeNode(i).SelectSingleNode("dispCtgId").Text
							dispCtgNm = LagrgeNode(i).SelectSingleNode("dispCtgNm").Text
							'rw LagrgeNode(i).SelectSingleNode("dispCtgClsCd").Text
							'rw LagrgeNode(i).SelectSingleNode("dispCtgClsCdNm").Text
							dispCtgPathNm = LagrgeNode(i).SelectSingleNode("dispCtgPathNm").Text
							'rw LagrgeNode(i).SelectSingleNode("aplSiteNo").Text
							'rw LagrgeNode(i).SelectSingleNode("aplSiteNoNm").Text
							'rw LagrgeNode(i).SelectSingleNode("dispCtgLastLvlYn").Text
							dispYn = LagrgeNode(i).SelectSingleNode("dispYn").Text
							If dispYn = "Y" Then
								strSql = ""
								strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_ssg_Newcategory] (siteNo, dispCtgId, dispCtgNm, dispCtgPathNm) "
								strSql = strSql & " VALUES ('"&siteNo&"', '"&dispCtgId&"', '"&html2db(dispCtgNm)&"', '"&html2db(dispCtgPathNm)&"') "
								dbget.Execute(strSql)
							End If
						End If
					Next
				End If
		Set xmlDOM = nothing
	Set objXML = nothing
End Function
%>
