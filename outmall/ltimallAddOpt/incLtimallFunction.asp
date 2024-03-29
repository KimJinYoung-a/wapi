<%
'############################################## 실제 수행하는 API 함수 모음 ##############################################
''롯데아이몰 상품 등록
Function LotteiMallItemReg(iitemid, strParam, byRef iErrStr, iSellCash, iLotteSellYn, imidx)
	Dim objXML,xmlDOM,strRst, lp, resultcode, resultmsg
	Dim buf, LotteGoodNo, strSql, buf_item_list, pp, OptDesc, StockQty, AssignedRow
	Dim ArgLength, NameValueArr(), j

	On Error Resume Next
	LotteiMallItemReg = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/registApiGoodsInfo.lotte", false
'rw ltiMallAPIURL & "/openapi/registApiGoodsInfo.lotte?"&strparam
'response.end
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
		    buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			'response.write buf
			LotteGoodNo = ""
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf ''BinaryToText(objXML.ResponseBody, "euc-kr")

				resultcode	= xmlDOM.getElementsByTagName("Result").item(0).text
				resultmsg	= xmlDOM.getElementsByTagName("Message").item(0).text
				LotteGoodNo = xmlDOM.getElementsByTagName("goods_no").item(0).text

				If resultcode <> 1 Then
		            iErrStr = "ERR||"&imidx&"||"&resultmsg&"(상품등록)"
				Else
					strSql = "Select count(*) From db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] WHERE midx='" & imidx & "'"
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If rsget(0) > 0 Then
						'// 존재 -> 수정
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCRLF
						strSql = strSql & "	Set LTiMallLastUpdate = getdate() "  & VbCRLF
						strSql = strSql & "	, LTiMallTmpGoodNo = '" & LotteGoodNo & "'"  & VbCRLF
						strSql = strSql & "	, LTiMallPrice = " &iSellCash& VbCRLF
						strSql = strSql & "	, accFailCnt = 0"& VbCRLF
						strSql = strSql & "	, LTiMallRegdate = isNULL(LTiMallRegdate, getdate())" ''추가 2013/02/26
						If (LotteGoodNo <> "") Then
						    strSql = strSql & "	, LTiMallstatCD = '20'"& VbCRLF			'승인요청
						Else
							strSql = strSql & "	, LTiMallstatCD = '1'"& VbCRLF			'전송시도
						End If
						strSql = strSql & "	From db_etcmall.dbo.tbl_ltimallAddOption_regItem R"& VbCRLF
						strSql = strSql & " WHERE R.midx='" & imidx & "'"
						dbget.Execute(strSql)
					Else
						'// 없음 -> 신규등록
						strSql = ""
						strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] "
						strSql = strSql & " (midx, reguserid, LTiMallRegdate, LTiMallLastUpdate, LTiMallTmpGoodNo, LTiMallPrice, LTiMallSellYn, LTiMallStatCd) VALUES " & VbCRLF
						strSql = strSql & " ('" & imidx & "'" & VBCRLF
						strSql = strSql & " , '" & session("ssBctId") & "'" &_
						strSql = strSql & " , getdate(), getdate()" & VBCRLF
						strSql = strSql & " , '" & LotteGoodNo & "'" & VBCRLF
						strSql = strSql & " , '" & iSellCash & "'" & VBCRLF
						strSql = strSql & " , '" & iLotteSellYn & "'" & VBCRLF
						If (LotteGoodNo <> "") Then
						    strSql = strSql & ",'20'"
						Else
						    strSql = strSql & ",'10'"
						End If
						strSql = strSql & ")"
						dbget.Execute(strSql)
					End If
					rsget.Close

					strSql = ""
					strSql = strSql & " UPDATE R "
					strSql = strSql & " SET itemname = i.itemname "
					strSql = strSql & " ,optionname = o.optionname "
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] R "
					strSql = strSql & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid "
					strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on R.itemid = o.itemid and R.itemoption = o.itemoption "
					strSql = strSql & " WHERE R.idx = '"&imidx&"' "
					strSql = strSql & " and R.mallid= 'lotteimall' "
					dbget.Execute strSql
					iErrStr =  "OK||"&imidx&"||등록성공(상품등록)"
				End If
			Set xmlDOM = Nothing
			LotteiMallItemReg= true
		Else
			iErrStr = "ERR||"&imidx&"||LotteiMall 결과 분석 중에 오류가 발생했습니다.[ERR-REG-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'품절 수행 함수
Public Function fnLtiMallSellyn(imidx, ichgSellYn, istrParam, byRef iErrStr)
    Dim strParam, resultcode, resultmsg
    Dim objXML, xmlDOM
    Dim strRst, strSql, buf
    fnLtiMallSellyn = False
	on Error resume next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", ltiMallAPIURL & "/openapi/updateGoodsSaleStat.lotte" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				resultcode = xmlDOM.getElementsByTagName("Result").item(0).text
				resultmsg = xmlDOM.getElementsByTagName("Message").item(0).text

				'// 오류 검출
				If resultcode <> 1 Then
		            iErrStr = "ERR||"&imidx&"||"&resultmsg
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ltimallAddOption_regItem " & VbCRLF
					strSql = strSql & " SET LtiMallLastUpdate = getdate() " & VbCRLF
					strSql = strSql & " ,LtiMallSellYn = '" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " ,accFailCnt = 0 " & VbCRLF
					strSql = strSql & " WHERE midx='" & imidx & "'"
					dbget.Execute(strSql)
					If ichgSellYn = "N" Then
						iErrStr = "OK||"&imidx&"||품절처리"
					ElseIf ichgSellYn = "Y" Then
						iErrStr = "OK||"&imidx&"||판매중으로 변경"
					Else
						iErrStr = "OK||"&imidx&"| 영구중단으로 변경"
					End If
		        End If
			Set xmlDOM = Nothing
			fnLtiMallSellyn = True
		Else
			iErrStr = "ERR||"&imidx&"||LotteiMall 결과 분석 중에 오류가 발생했습니다.[ERR-SELLEDIT-001]"
		End If
	Set objXML = Nothing
	on Error Goto 0
End Function

'전시상품 판매가 수정
Public Function fnLtimallPrice(imidx, istrParam, imustprice, byRef iErrStr)
    Dim objXML,xmlDOM,strRst
    Dim buf, strSql
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnLtimallPrice = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/updateGoodsSalePrcOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)

		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				resultcode = xmlDOM.getElementsByTagName("Result").item(0).text
				resultmsg = xmlDOM.getElementsByTagName("Message").item(0).text

				If resultcode <> 1 Then
		            iErrStr =  "ERR||"&imidx&"||"&resultmsg&"(상품가격)"
		            fnLtimallPrice = False
				Else
				    '// 상품가격정보 수정
				    strSql = ""
	    			strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_ltimallAddOption_regItem]  " & VbCRLF
	    			strSql = strSql & "	SET ltimallLastUpdate = getdate() " & VbCRLF
	    			strSql = strSql & "	, ltimallPrice = " & imustprice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " WHERE midx='" & imidx & "'"& VbCRLF
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&imidx&"||수정성공(상품가격)"
					fnLtimallPrice = True
				End If
			Set xmlDOM = Nothing
		Else
			fnLtimallPrice = False
			iErrStr = "ERR||"&imidx&"||LotteiMall 결과 분석 중에 오류가 발생했습니다.[ERR-PRICE-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Public Function fnLtiMallChgItemname(imidx, strParam, iErrStr)
	Dim objXML, xmlDOM, strRst, resultmsg, resultcode, strSql
	On Error Resume Next
	fnLtiMallChgItemname = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/updateGoodsNmOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				resultcode = xmlDOM.getElementsByTagName("Result").item(0).text
			    resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text

				If resultcode <> 1 Then
		            iErrStr =  "ERR||"&imidx&"||"&resultmsg&"(상품명)"
		            fnLtiMallChgItemname = False
				Else
					strSql = ""
					strSql = strSql & " UPDATE R "
					strSql = strSql & " SET itemname = i.itemname "
					strSql = strSql & " ,optionname = o.optionname "
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] R "
					strSql = strSql & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid "
					strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on R.itemid = o.itemid and R.itemoption = o.itemoption "
					strSql = strSql & " WHERE R.idx = '"&imidx&"' "
					strSql = strSql & " and R.mallid= 'lotteimall' "
					dbget.Execute(strSql)
					iErrStr =  "OK||"&imidx&"||수정성공(상품명)"
					fnLtiMallChgItemname = True
			    End If
			Set xmlDOM = Nothing
		Else
			fnLtiMallChgItemname = False
			iErrStr = "ERR||"&imidx&"||LotteiMall 결과 분석 중에 오류가 발생했습니다.[ERR-NMEDIT-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Public Function fnLtiMallstatChk(imidx, iErrStr)
	Dim objXML,xmlDOM,strRst,resultmsg, iLotteGoodNo, strSql
	Dim strParam, iLotteTmpID, SaleStatCd, GoodsViewCount
	Dim iRbody, ltimallStatName
	On Error Resume Next
	fnLtiMallstatChk = False
	iLotteTmpID = getLtiMallTmpItemIdByTenItemID(imidx)

	If (iLotteTmpID = "") OR (iLotteTmpID = "전시상품") then
		iErrStr =  "ERR||"&imidx&"||이미 전시상품 입니다.(신규상품조회)"
		Exit function
	End If

	strParam = "subscriptionId=" & ltiMallAuthNo & "&search_type=3&search_value="&iLotteTmpID
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/searchNewGoodsInfoOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				'rw iRbody
				GoodsViewCount 	= Trim(xmlDOM.getElementsByTagName("GoodsCount").item(0).text)		'검색수
				resultmsg 		= xmlDOM.getElementsByTagName("Message").Item(0).Text
				iLotteGoodNo	= Trim(xmlDOM.getElementsByTagName("GoodsNo").item(0).text)			'전시상품번호
				SaleStatCd		= Trim(xmlDOM.getElementsByTagName("ConfStatCd").item(0).text)		'인증상태코드

				If GoodsViewCount <> 1 Then
					If resultmsg = "" Then
						resultmsg = "조회결과 없음"
					End If
		            iErrStr =  "ERR||"&imidx&"||"&resultmsg&"(신규상품조회)"
		            fnLtiMallstatChk = False
				Else

					Select Case SaleStatCd
						Case "10"	ltimallStatName = "임시등록"
						Case "20"	ltimallStatName = "승인요청"
						Case "30"	ltimallStatName = "승인완료"
						Case "40"	ltimallStatName = "반려"
						Case "50"	ltimallStatName = "승인불가"
						Case "51"	ltimallStatName = "재승인요청"
						Case "52"	ltimallStatName = "수정요청"
						Case "60"	ltimallStatName = "취소"
					End Select

					If SaleStatCd = "30" Then				'승인완료(LtiMallStatCd, LtiMallGoodNo, lastConfirmdate 수정)
						strSql = ""
						strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] " & VbCRLF
						strSql = strSql & " SET lastConfirmdate = getdate() "& VbCRLF
						strSql = strSql & "	,LtiMallStatCd='7' "
						strSql = strSql & " ,LtiMallGoodNo='" & iLotteGoodNo & "' "
						strSql = strSql & " WHERE midx='" & imidx & "'"& VbCRLF
						dbget.Execute(strSql)
					Else
						strSql = ""
						strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] " & VbCRLF
						strSql = strSql & " SET lastConfirmdate = getdate() "& VbCRLF
						strSql = strSql & "	,LtiMallStatCd='"&SaleStatCd&"' "& VbCRLF
						strSql = strSql & " WHERE midx='" & imidx & "'"& VbCRLF
						dbget.Execute(strSql)
					End If
					iErrStr =  "OK||"&imidx&"||성공(신규상품조회) : "&ltimallStatName
					fnLtiMallstatChk = True
			    End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "롯데아이몰과 통신중에 오류가 발생했습니다..[ERR-STATCHK-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function fnLtiMallDisView(imidx,byRef iErrStr, iLottegoodNo)
	Dim objXML, xmlDOM, strRst, resultmsg, assignedRow
	Dim strParam, iLotteItemID , SaleStatCd, GoodsViewCount
	Dim iRbody, LotteSellyn, iSalePrc, iGoodsNm, sqlstr

	fnLtiMallDisView = False
	strParam = "subscriptionId=" & ltiMallAuthNo & "&goods_no="&iLottegoodNo
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/searchGoodsListOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
'				response.write iRbody
				GoodsViewCount = xmlDOM.getElementsByTagName("GoodsCount").item(0).text
				If (GoodsViewCount = "1") Then
			    	SaleStatCd = xmlDOM.getElementsByTagName("SaleStatCd").item(0).text
			    	iSalePrc	= xmlDOM.getElementsByTagName("SalePrc").item(0).text
			    	iGoodsNm	= xmlDOM.getElementsByTagName("GoodsNm").item(0).text
			    	iGoodsNm	= replace(iGoodsNm,"@@amp@@","&")
					iGoodsNm	= Replace(iGoodsNm,"&gt;",">")
					iGoodsNm	= Replace(iGoodsNm,"&lt;","<")
					iGoodsNm	= Replace(iGoodsNm,"&nbsp;"," ")
					iGoodsNm	= Replace(iGoodsNm,"&amp;","&")

					If (SaleStatCd="10") Then
					    LotteSellyn = "Y"
					ElseIf (SaleStatCd="20") Then
					    LotteSellyn = "N"
					ElseIf (SaleStatCd="30") Then
					    LotteSellyn = "X"
					End If
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] SET " & VBCRLF
					strSql = strSql & " DispViewCnt = isnull(DispViewCnt, 0) + 1 " & VBCRLF
					strSql = strSql & " ,ltiMallPrice="&iSalePrc & VbCRLF
					If (LotteSellyn <> "") Then
					    strSql = strSql & " ,ltiMallSellyn='"&LotteSellyn&"'"
					End If
					strSql = strSql & " WHERE midx = '"&imidx&"' "
					dbget.Execute strSql, assignedRow
			    	iErrStr =  "OK||"&imidx&"||성공(전시상품조회)"
					fnLtiMallDisView = True
			    Else
			    	resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text
			    	iErrStr =  "ERR||"&imidx&"||"&resultmsg&"(전시상품조회)"
		            fnLtiMallDisView = False
			    End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "롯데아이몰과 통신중에 오류가 발생했습니다..[ERR-ItemChk-001]"
			fnLtiMallDisView = False
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Public Function fnLtiMallAddOpt(iitemid, istrParam, byRef iErrStr, iAddOptCnt)
	Dim objXML, xmlDOM, strRst, resultmsg, strSql
	Dim addOptCount, opt1Nm, opt2Nm, opt3Nm, opt4Nm, opt5Nm, opt1Tval, opt2Tval, opt3Tval, opt4Tval, opt5Tval
	On Error Resume Next
	fnLtiMallAddOpt = False
	addOptCount = 0
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/addGoodsItemInfo.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				If xmlDOM.getElementsByTagName("itemCount").item(0).text = "" Then
					addOptCount = 0
					iAddOptCnt = 0
				Else
    				addOptCount = xmlDOM.getElementsByTagName("itemCount").item(0).text
    				iAddOptCnt = addOptCount
    			End If

				If addOptCount > 0 Then
					opt1Nm 		= xmlDOM.getElementsByTagName("opt1Nm").item(0).text
					opt1Tval	= xmlDOM.getElementsByTagName("opt1Tval").item(0).text
					opt2Nm 		= xmlDOM.getElementsByTagName("opt2Nm").item(0).text
					opt2Tval	= xmlDOM.getElementsByTagName("opt2Tval").item(0).text
					opt3Nm 		= xmlDOM.getElementsByTagName("opt3Nm").item(0).text
					opt3Tval	= xmlDOM.getElementsByTagName("opt3Tval").item(0).text
					opt4Nm 		= xmlDOM.getElementsByTagName("opt4Nm").item(0).text
					opt4Tval	= xmlDOM.getElementsByTagName("opt4Tval").item(0).text
					opt5Nm 		= xmlDOM.getElementsByTagName("opt5Nm").item(0).text
					opt5Tval	= xmlDOM.getElementsByTagName("opt5Tval").item(0).text
				    resultmsg 	= xmlDOM.getElementsByTagName("Message").Item(0).Text
					If resultmsg = "" Then
						iErrStr =  "OK||"&iitemid&"||성공(옵션추가)"
						fnLtiMallAddOpt = True
				    Else
			            iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(옵션추가)"
			            fnLtiMallAddOpt = False
				    End If
			    End If
			Set xmlDOM = Nothing
		Else
			fnLtiMallAddOpt = False
			iErrStr = "롯데아이몰과 통신중에 오류가 발생했습니다..[ERR-ADDOPT-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

''롯데아이몰 상품정보 수정
Function fnLtiMallInfoEdit(imidx, strParam, byRef iErrStr, isVer2)
	Dim objXML, xmlDOM, strRst, resultmsg, resultcode
	On Error Resume Next
	fnLtiMallInfoEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	If (isVer2) Then
	    objXML.Open "POST", ltiMallAPIURL & "/openapi/upateApiNewGoodsInfo.lotte", false          ''상품수정
	Else
	    objXML.Open "POST", ltiMallAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte", false      ''전시상품수정
	End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				resultcode = xmlDOM.getElementsByTagName("Result").item(0).text
			    resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text

			    If (resultcode = "1") Then
					iErrStr =  "OK||"&imidx&"||성공(상품정보)"
					fnLtiMallInfoEdit = True
				Else
		            iErrStr =  "ERR||"&imidx&"||"&resultmsg&"(상품정보)"
		            fnLtiMallInfoEdit = False
			    End If
			Set xmlDOM = Nothing
		Else
			fnLtiMallInfoEdit = False
			iErrStr = "롯데아이몰과 통신중에 오류가 발생했습니다..[ERR-EDIT-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

''전시상품 조회
Function fnCheckLtiMallItemStat(iitemid,byRef iErrStr, iLottegoodNo)
	Dim objXML, xmlDOM, strRst, resultmsg, assignedRow
	Dim strParam, iLotteItemID , SaleStatCd, GoodsViewCount
	Dim iRbody, LotteSellyn, iSalePrc, iGoodsNm, sqlstr

	fnCheckLtiMallItemStat = False
	strParam = "subscriptionId=" & ltiMallAuthNo & "&goods_no="&iLottegoodNo
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/searchGoodsListOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody

				GoodsViewCount = xmlDOM.getElementsByTagName("GoodsCount").item(0).text
				If (GoodsViewCount = "1") Then
			    	SaleStatCd = xmlDOM.getElementsByTagName("SaleStatCd").item(0).text
			    	iSalePrc	= xmlDOM.getElementsByTagName("SalePrc").item(0).text
			    	iGoodsNm	= xmlDOM.getElementsByTagName("GoodsNm").item(0).text
			    	iGoodsNm	= replace(iGoodsNm,"@@amp@@","&")
					iGoodsNm	= Replace(iGoodsNm,"&gt;",">")
					iGoodsNm	= Replace(iGoodsNm,"&lt;","<")
					iGoodsNm	= Replace(iGoodsNm,"&nbsp;"," ")
					iGoodsNm	= Replace(iGoodsNm,"&amp;","&")

					If (SaleStatCd="10") Then
					    LotteSellyn = "Y"
					ElseIf (SaleStatCd="20") Then
					    LotteSellyn = "N"
					ElseIf (SaleStatCd="30") Then
					    LotteSellyn = "X"
					End If

					sqlstr = "Update R" & VbCRLF
					sqlstr = sqlstr & " SET regitemname='"&html2db(iGoodsNm)&"'"
'					If (LotteSellyn <> "") Then
'					    sqlstr = sqlstr & " ,ltiMallSellyn='"&LotteSellyn&"'"
'					End If
'					sqlstr = sqlstr & " ,LtiMallStatCd=(CASE WHEN isNULL(LtiMallStatCd,-9)<7 THEN 7 ELSE LtiMallStatCd END)" ''2013/09/03 추가
'					sqlstr = sqlstr & " ,lastStatCheckDate=getdate()"
					sqlstr = sqlstr & " FROM db_item.dbo.tbl_LTiMall_regItem R" & VbCRLF
					sqlstr = sqlstr & " WHERE R.itemid="&iitemid & VbCRLF
					dbget.Execute sqlstr, assignedRow
			    	iErrStr =  "OK||"&iitemid&"||성공(전시상품조회)"
					fnCheckLtiMallItemStat = True
			    Else
			    	resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text
			    	iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(전시상품조회)"
		            fnCheckLtiMallItemStat = False
			    End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "롯데아이몰과 통신중에 오류가 발생했습니다..[ERR-ItemChk-001]"
			fnCheckLtiMallItemStat = False
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################

'################################################# 각 기능 별 파라메터 정리 ###############################################
'품절 파라메타
Function getLtiMallSellynParameter(ichgSellYn, iLotteGoodNo)
'iLotteGoodNo = 1035664985
    Dim strRst
	strRst = "?subscriptionId=" & ltiMallAuthNo										'롯데아이몰 인증번호	(*)
	strRst = strRst & "&goods_no=" & iLotteGoodNo                       			'롯데아이몰 상품번호	(*)
	If ichgSellYn = "Y" Then														'판매여부(10:판매, 20:품절, 30:판매종료)
		strRst = strRst & "&sale_stat_cd=10"
	ElseIf ichgSellYn = "N" Then
		strRst = strRst & "&sale_stat_cd=20"
	ElseIf ichgSellYn = "X" Then													'''X 기능 사용안함
		strRst = strRst & "&sale_stat_cd=30"
	End If
	getLtiMallSellynParameter = strRst
End Function

'가격 수정 파라메터 생성
Function getLtiMallPriceParameter(imidx, iLotteGoodNo, byref MustPrice)
	Dim strRst, strSql
	Dim sellcash, orgprice, buycash, optaddprice
	Dim GetTenTenMargin

	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.sellcash, i.orgprice, i.buycash, o.optaddprice "
	strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M "
	strSql = strSql & " JOIN db_item.dbo.tbl_item as i on M.itemid = i.itemid "
	strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_option as o on M.itemid = o.itemid and M.itemoption = o.itemoption "
	strSql = strSql & " WHERE M.idx = '"&imidx&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		sellcash	= rsget("sellcash")
		orgprice	= rsget("orgprice")
		buycash		= rsget("buycash")
		optaddprice	= rsget("optaddprice")
	Else
		getLtiMallPriceParameter = ""
		Exit Function
		response.end
	End If
	rsget.close

	GetTenTenMargin = CLng(10000 - buycash / sellcash * 100 * 100) / 100
	If GetTenTenMargin < CMAXMARGIN Then
		MustPrice = orgprice + optaddprice
	Else
		MustPrice = sellcash + optaddprice
	End If

	strRst = "subscriptionId=" & ltiMallAuthNo
	strRst = strRst & "&strGoodsNo=" & iLotteGoodNo
	strRst = strRst & "&strReqSalePrc=" & GetRaiseValue(MustPrice/10)*10
	getLtiMallPriceParameter = strRst
End Function

''//상품명 변경 파라메터 생성(롯데닷컴과 파라매타명이 다름)
Function getLtiMallItemnameParameter(iidx, byref iitemname, iLotteGoodNo)
	Dim strSql, chgname, strRst, newitemname, itemnameChange
	strSql = ""
	strSql = strSql & " SELECT TOP 1 M.itemid, convert(varchar(30),m.itemid) + convert(varchar(30),m.itemoption) as newCode, isnull(M.newitemname, '') as newitemname, isnull(M.itemnameChange, '') as itemnameChange "
	strSql = strSql & "	FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M "
	strSql = strSql & "	JOIN db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] as R on M.idx = R.midx "
	strSql = strSql & "	WHERE M.idx = '"&iidx&"' "
	strSql = strSql & "	and M.mallid = 'lotteimall' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.Eof Then
		newitemname		= rsget("newitemname")
		itemnameChange	= rsget("itemnameChange")
	End If
	rsget.close

	If itemnameChange = "" Then
		iitemname = newitemname
	Else
		iitemname = itemnameChange
	End If

	chgname = ""
	chgname = replace(db2html(iitemname),"'","")
	chgname = replace(chgname,"<B>","")
	chgname = replace(chgname,"</B>","")
	chgname = replace(chgname,"~","-")
	chgname = replace(chgname,"<","[")
	chgname = replace(chgname,">","]")
	chgname = replace(chgname,"%","프로")
	chgname = replace(chgname,"[무료배송]","")
	chgname = replace(chgname,"[무료 배송]","")
	strRst = "subscriptionId=" & ltiMallAuthNo
	strRst = strRst & "&goods_no=" & iLotteGoodNo
	strRst = strRst & "&goods_nm=" & Trim(chgname)
	strRst = strRst & "&chg_caus_cont=api 상품명 변경"
	getLtiMallItemnameParameter = strRst
End Function

'옵션 추가 파라메타
Function getLtiMallAddOptParameter(nm, dc, iLotteGoodNo)
	Dim strRst
	strRst = "subscriptionId=" & ltiMallAuthNo											'(*)사용자인증키
	strRst = strRst & "&goods_no=" & iLotteGoodNo										'(*)롯데아이몰 상품번호
	strRst = strRst & "&opt_nm=" & nm													'(*)롯데아이몰 추가할 옵션명
	strRst = strRst & "&item_nm=" & dc													'(*)롯데아이몰 추가할 옵션종류명
	getLtiMallAddOptParameter = strRst
End Function
'################################################ 각 기능 별 파라메터 정리 끝 #############################################

'################################################ 이하는 위 기능하기 위한 함수 ############################################
Public Function GetRaiseValue(value)
    If Fix(value) < value Then
    	GetRaiseValue = Fix(value) + 1
    Else
    	GetRaiseValue = Fix(value)
    End If
End Function

Function getLtiMallTmpItemIdByTenItemID(iimidx)
	Dim sqlStr, retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT ltiMallTmpGoodNo, isnull(ltiMallGoodNo,'') as ltiMallGoodNo " & VBCRLF
	sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] " & VBCRLF
	sqlStr = sqlStr & " WHERE midx = "&iimidx & VBCRLF
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		If rsget("ltiMallGoodNo") <> "" Then
			retVal = "전시상품"
		Else
			retVal = rsget("ltiMallTmpGoodNo")
		End If
	End If
	rsget.Close

	If IsNULL(retVal) Then retVal = ""
	getLtiMallTmpItemIdByTenItemID = retVal
End Function
%>