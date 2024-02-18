<%
'############################################## 실제 수행하는 API 함수 모음 ##############################################
'롯데닷컴 상품 등록
Function fnLotteComItemReg(iitemid, strParam, byRef iErrStr, iSellCash, iLotteSellYn, imidx, ibasicImage)
	Dim objXML,xmlDOM,strRst, resultmsg, resultcode
	Dim ArgLength, NameValueArr(), j, k
	Dim buf, LotteGoodNo, strSql, buf_item_list, pp, OptDesc, StockQty, AssignedRow
	On Error Resume Next
	fnLotteComItemReg = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", lotteAPIURL & "/openapi/registApiGoodsInfo.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
		    buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			LotteGoodNo = ""
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				resultcode	= xmlDOM.getElementsByTagName("Result").item(0).text
				resultmsg	= xmlDOM.getElementsByTagName("Message").item(0).text
				LotteGoodNo = xmlDOM.getElementsByTagName("goods_no").item(0).text

				If resultcode <> 1 Then
		            iErrStr = "ERR||"&imidx&"||"&resultmsg&"(상품등록)"
				Else
					strSql = "Select count(*) From db_etcmall.[dbo].[tbl_lotteAddOption_regItem] WHERE midx='" & imidx & "'"
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If rsget(0) > 0 Then
						'// 존재 -> 수정
						strSql = "update R" & VbCRLF
						strSql = strSql & "	Set LotteLastUpdate=getdate() "  & VbCRLF
						strSql = strSql & "	, LotteTmpGoodNo='" & LotteGoodNo & "'"  & VbCRLF
						strSql = strSql & "	, LottePrice=" &iSellCash& VbCRLF
						strSql = strSql & "	, regImageName = '" & ibasicImage & "'" & VbCRLF
						strSql = strSql & "	, accFailCnt=0"& VbCRLF
						strSql = strSql & "	, lotteRegdate=isNULL(lotteRegdate,getdate())" ''추가 2013/02/26
						If (LotteGoodNo <> "") Then
							strSql = strSql & "	, lottestatCD='20'"& VbCRLF
						Else
							strSql = strSql & "	, lottestatCD='10'"& VbCRLF
						End If
						strSql = strSql & "	FROM db_etcmall.[dbo].[tbl_lotteAddOption_regItem] R"& VbCRLF
						strSql = strSql & " WHERE R.midx='" & imidx & "'"
						dbget.Execute(strSql)
					Else
						'// 없음 -> 신규등록
						strSql = "INSERT INTO db_etcmall.[dbo].[tbl_lotteAddOption_regItem] "
						strSql = strSql & " (midx, reguserid, lotteRegdate, LotteLastUpdate, LotteTmpGoodNo, LottePrice, LotteSellYn, LotteStatCd, regImageName) values " & VbCRLF
						strSql = strSql & " ('" & imidx & "'" & VbCRLF
						strSql = strSql & ", '" & session("ssBctId") & "'" &_
						strSql = strSql & ", getdate(), getdate()" & VbCRLF
						strSql = strSql & ", '" & LotteGoodNo & "'" & VbCRLF
						strSql = strSql & ", '" & iSellCash & "'" & VbCRLF
						strSql = strSql & ", '" & iLotteSellYn & "'" & VbCRLF
						If (LotteGoodNo <> "") Then
							strSql = strSql & ",'20'"
						Else
							strSql = strSql & ",'10'"
						End If
						strSql = strSql & ", '" & ibasicImage & "'" & VbCRLF
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
					strSql = strSql & " WHERE R.idx = '"&midx&"' "
					strSql = strSql & " and R.mallid= 'lotteCom' "
		       		iErrStr =  "OK||"&imidx&"||등록성공(상품등록)"
				End If
			Set xmlDOM = Nothing
			fnLotteComItemReg= true
		Else
			iErrStr = "ERR||"&imidx&"||LotteCom 결과 분석 중에 오류가 발생했습니다.[ERR-REG-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'롯데닷컴 판매상태변경
Function fnLotteComSellyn(imidx, ichgSellYn, istrParam, byRef iErrStr)
    Dim strParam
    Dim objXML, xmlDOM
    Dim strRst, strSql, resultcode, resultmsg
    fnLotteComSellyn = False
	on Error resume next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", lotteAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
'					response.write BinaryToText(objXML.ResponseBody, "euc-kr")
'					response.end
				resultcode = xmlDOM.getElementsByTagName("Result").item(0).text
				resultmsg = xmlDOM.getElementsByTagName("Message").item(0).text

				If resultcode <> 1 Then
		            iErrStr = "ERR||"&imidx&"||"&resultmsg
				Else
					'// 상품정보 수정
					strSql = "Update db_etcmall.[dbo].[tbl_lotteAddOption_regItem] " & VbCRLF
					strSql = strSql & " Set LotteLastUpdate=getdate() " & VbCRLF
					strSql = strSql & " ,LotteSellYn='" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " ,accFailCnt = 0 " & VbCRLF
					strSql = strSql & " WHERE midx='" & imidx & "'"
					dbget.Execute(strSql)

					If ichgSellYn = "N" Then
						iErrStr = "OK||"&imidx&"||품절처리(상태)"
					Else
						iErrStr = "OK||"&imidx&"||판매중으로 변경(상태)"
					End If
				End If
			Set xmlDOM = Nothing
			fnLotteComSellyn = True
		Else
			iErrStr = "ERR||"&imidx&"||롯데닷컴 결과 분석 중에 오류가 발생했습니다.[ERR-SELLEDIT-001]"
		End If
	Set objXML = Nothing
	on Error Goto 0
End Function

'롯데닷컴 전시상품 맵핑정보
Public Function fnLotteComStatChk(imidx, iErrStr)
	Dim objXML,xmlDOM,strRst,resultmsg, iLotteGoodNo, strSql
	Dim strParam, iLotteTmpID, SaleStatCd, GoodsViewCount
	Dim iRbody, resultcode, lotteStatName
	On Error Resume Next
	fnLotteComStatChk = False
	iLotteTmpID = getLotteTmpItemIdByTenItemID(imidx)

	If (iLotteTmpID = "") OR (iLotteTmpID = "전시상품") then
		iErrStr =  "ERR||"&imidx&"||이미 전시상품 입니다.(신규상품조회)"
		Exit function
	End If

	strParam = "subscriptionId=" & lotteAuthNo & "&goods_req_no=" & iLotteTmpID
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", lotteAPIURL & "/openapi/getRdToPrGoodsNoApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody

				resultcode		= xmlDOM.getElementsByTagName("Result").item(0).text
				iLotteGoodNo	= Trim(xmlDOM.getElementsByTagName("goods_no").item(0).text)		'전시상품번호
				SaleStatCd		= Trim(xmlDOM.getElementsByTagName("conf_stat_cd").item(0).text)	'인증상태코드

				If resultcode <> 1 Then
					If resultmsg = "" Then
						resultmsg = "조회결과 없음"
					End If
		            iErrStr =  "ERR||"&imidx&"||"&resultmsg&"(신규상품조회)"
		            fnLotteComStatChk = False
				Else

					Select Case SaleStatCd
						Case "10"	lotteStatName = "임시등록"
						Case "20"	lotteStatName = "승인요청"
						Case "30"	lotteStatName = "승인완료"
						Case "40"	lotteStatName = "반려"
						Case "50"	lotteStatName = "승인불가"
						Case "51"	lotteStatName = "재승인요청"
						Case "52"	lotteStatName = "수정요청"
					End Select
					If SaleStatCd = "30" Then				'승인완료(lotteStatCd, lotteGoodNo, lastConfirmdate 수정)
						strSql = ""
						strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_lotteAddOption_regItem] " & VbCRLF
						strSql = strSql & " SET lastConfirmdate = getdate() "& VbCRLF
						strSql = strSql & "	,lotteStatCd='30' "
						strSql = strSql & " ,lotteGoodNo='" & iLotteGoodNo & "' "
						strSql = strSql & " WHERE midx='" & imidx & "'"& VbCRLF
						dbget.Execute(strSql)
					Else
						strSql = ""
						strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_lotteAddOption_regItem] " & VbCRLF
						strSql = strSql & " SET lastConfirmdate = getdate() "& VbCRLF
						strSql = strSql & "	,lotteStatCd='"&SaleStatCd&"' "& VbCRLF
						strSql = strSql & " WHERE midx='" & imidx & "'"& VbCRLF
						dbget.Execute(strSql)
					End If
					iErrStr =  "OK||"&imidx&"||성공(신규상품조회) : "&lotteStatName
					fnLotteComStatChk = True
			    End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-STATCHK-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Public Function fnLotteComPrice(imidx, istrParam, imustprice, byRef iErrStr)
	Dim objXML, xmlDOM, strRst
	Dim resultcode, resultmsg, strSql
	On Error Resume Next
	fnLotteComPrice = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", lotteAPIURL & "/openapi/updateGoodsSalePrcOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				resultcode = xmlDOM.getElementsByTagName("Result").item(0).text
				resultmsg = xmlDOM.getElementsByTagName("Message").item(0).text

				If resultcode <> 1 Then
		            iErrStr =  "ERR||"&imidx&"||"&resultmsg&"(상품가격)"
		            fnLotteComPrice = False
				Else
				    '// 상품가격정보 수정
				    strSql = ""
	    			strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_lotteAddOption_regItem]  " & VbCRLF
	    			strSql = strSql & "	SET LotteLastUpdate = getdate() " & VbCRLF
	    			strSql = strSql & "	, LottePrice = " & imustprice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " WHERE midx='" & imidx & "'"& VbCRLF
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&imidx&"||수정성공(상품가격)"
					fnLotteComPrice = True
				End If
			Set xmlDOM = Nothing
		Else
			fnLotteComPrice = False
			iErrStr = "ERR||"&imidx&"||롯데닷컴과 통신중에 오류가 발생했습니다.[ERR-PRICE-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Public Function fnLotteComChgItemname(imidx, strParam, iErrStr)
	Dim objXML, xmlDOM, strRst, strSql
	Dim resultcode, resultmsg
	On Error Resume Next
	fnLotteComChgItemname = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", lotteAPIURL & "/openapi/updateGoodsNmOpenApi.lotte", false
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
		            fnLotteComChgItemname = False
				Else
					strSql = ""
					strSql = strSql & " UPDATE R "
					strSql = strSql & " SET itemname = i.itemname "
					strSql = strSql & " ,optionname = o.optionname "
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] R "
					strSql = strSql & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid "
					strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on R.itemid = o.itemid and R.itemoption = o.itemoption "
					strSql = strSql & " WHERE R.idx = '"&imidx&"' "
					strSql = strSql & " and R.mallid= 'lotteCom' "
					dbget.Execute(strSql)

					iErrStr =  "OK||"&imidx&"||수정성공(상품명)"
					fnLotteComChgItemname = True
			    End If
			Set xmlDOM = Nothing
		else
			iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-NMEDIT-002]"
		end if
	Set objXML = Nothing
	On Error Goto 0
End Function

''롯데닷컴 상품정보 수정
Function fnLotteComInfoEdit(imidx, strParam, byRef iErrStr, isVer2)
	Dim objXML, xmlDOM, strRst
	Dim resultcode, resultmsg
	On Error Resume Next
	fnLotteComInfoEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	If (isVer2) Then
		objXML.Open "POST", lotteAPIURL & "/openapi/upateApiNewGoodsInfo.lotte", false          ''상품수정
	Else
		objXML.Open "POST", lotteAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte", false      ''전시상품수정
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
					fnLotteComInfoEdit = True
				Else
		            iErrStr =  "ERR||"&imidx&"||"&resultmsg&"(상품정보)"
		            fnLotteComInfoEdit = False
			    End If
			Set xmlDOM = Nothing
		Else
			fnLotteComInfoEdit = False
			iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-EDIT-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function fnLotteComInfodivEdit(imidx, strParam, byRef iErrStr)
	Dim objXML,xmlDOM,strRst,iMessage
	Dim resultcode, resultmsg

	On Error Resume Next
	fnLotteComInfodivEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", lotteAPIURL & "/openapi/upateApiDisplayGoodsItemInfo.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				resultcode = xmlDOM.getElementsByTagName("Result").item(0).text
			    resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text

			    If (resultcode = "1") Then
					iErrStr =  "OK||"&imidx&"||성공(품목정보)"
					fnLotteComInfodivEdit = True
				Else
		            iErrStr =  "ERR||"&imidx&"||"&resultmsg&"(품목정보)"
		            fnLotteComInfodivEdit = False
			    End If
			Set xmlDOM = Nothing
		Else
			fnLotteComInfodivEdit = False
			iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-PoomEDIT-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Public Function fnLotteComImageEdit(imidx, strParam, byRef iErrStr)
	Dim objXML,xmlDOM,strRst,iMessage
	Dim resultcode, resultmsg

	On Error Resume Next
	fnLotteComImageEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", lotteAPIURL & "/openapi/registApiGoodsImageInfo.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				resultcode = xmlDOM.getElementsByTagName("Result").item(0).text
			    resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text

			    If (resultcode = "1") Then
					strSql = ""
					strSql = strSql & " UPDATE L "
					strSql = strSql & " SET L.regImageName = i.basicimage "
					strSql = strSql & " FROM db_etcmall.dbo.tbl_lotteAddOption_regItem as L "
					strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] R on R.idx = L.midx "
					strSql = strSql & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid "
					strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on R.itemid = o.itemid and R.itemoption = o.itemoption "
					strSql = strSql & " WHERE R.idx = '"&imidx&"' "
					strSql = strSql & " and R.mallid= 'lotteCom' "
					dbget.Execute(strSql)

					iErrStr =  "OK||"&imidx&"||성공(이미지수정)"
					fnLotteComImageEdit = True
				Else
		            iErrStr =  "ERR||"&imidx&"||"&resultmsg&"(이미지수정)"
		            fnLotteComImageEdit = False
			    End If
			Set xmlDOM = Nothing
		Else
			fnLotteComImageEdit = False
			iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-IMAGE-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

''전시상품 조회
Function fnCheckLotteComItemStat(imidx, byRef iErrStr, iLottegoodNo)
	Dim objXML, xmlDOM, strRst, resultmsg
	Dim strParam, SaleStatCd, GoodsViewCount, iSalePrc, iGoodsNm
	Dim iRbody, LotteSellyn, sqlStr, assignedRow

	fnCheckLotteComItemStat = false
	strParam = "subscriptionId=" & lotteAuthNo & "&strGoodsNo="&iLottegoodNo

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", lotteAPIURL & "/openapi/searchGoodsListOpenApiOther.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				iRbody = replace(iRbody,"&","@@amp@@")   '' <![CDATA[]]> 로 안 묶여옴. 상품명에 < > 있음..
				iRbody = replace(iRbody,"<GoodsNm>","<GoodsNm><![CDATA[")
				iRbody = replace(iRbody,"</GoodsNm>","]]></GoodsNm>")
				xmlDOM.LoadXML iRbody

				GoodsViewCount = xmlDOM.getElementsByTagName("GoodsViewCount").item(0).text  ''결과수

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

					sqlstr = ""
					sqlstr = sqlstr & " Update R" & VbCRLF
					sqlstr = sqlstr & " SET lastStatCheckDate=getdate()"
					IF (LotteSellyn <> "") then
						sqlstr = sqlstr & " ,LotteSellyn='"&LotteSellyn&"'"
					ENd IF
					sqlstr = sqlstr & " From db_etcmall.[dbo].[tbl_lotteAddOption_regItem] R" & VbCRLF
					sqlstr = sqlstr & " where R.midx="&imidx & VbCRLF
					dbget.Execute sqlstr,assignedRow
			    	iErrStr =  "OK||"&imidx&"||성공(전시상품조회)"
					fnCheckLotteComItemStat = True
			    Else
			    	resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text
			    	iErrStr =  "ERR||"&imidx&"||"&resultmsg&"(전시상품조회)"
		            fnCheckLotteComItemStat = False
			    End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-ItemChk-001]"
			fnCheckLotteComItemStat = False
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################

'################################################# 각 기능 별 파라메터 정리 ###############################################
'품절 파라메타
Function getLotteComSellynParameter(ichgSellYn, iLotteGoodNo)
    Dim strRst
	strRst = "?subscriptionId=" & lotteAuthNo
	strRst = strRst & "&goods_no=" & iLotteGoodNo
	If ichgSellYn = "Y" Then														'판매여부(10:판매, 20:품절, 30:판매종료)
		strRst = strRst & "&sale_stat_cd=10"
	ElseIf ichgSellYn = "N" Then
		strRst = strRst & "&sale_stat_cd=20"
	ElseIf ichgSellYn = "X" Then													'''X 기능 사용안함
		strRst = strRst & "&sale_stat_cd=20"
	End If
	getLotteComSellynParameter = strRst
End Function

Function getLotteTmpItemIdByTenItemID(iimidx)
	Dim sqlStr, retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT lotteTmpGoodNo, isnull(lotteGoodNo,'') as lotteGoodNo " & VBCRLF
	sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_lotteAddOption_regItem] " & VBCRLF
	sqlStr = sqlStr & " WHERE midx = "&iimidx & VBCRLF
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		If rsget("lotteGoodNo") <> "0" Then
			retVal = "전시상품"
		Else
			retVal = rsget("lotteTmpGoodNo")
		End If
	End If
	rsget.Close

	If IsNULL(retVal) Then retVal = ""
	getLotteTmpItemIdByTenItemID = retVal
End Function

 '// 가격 수정 파라메터 생성
Function getLotteComPriceParameter(imidx, iLotteGoodNo, byref MustPrice)
	Dim strRst, strSql
	Dim sellcash, orgprice, buycash, optaddprice
	Dim GetTenTenMargin
	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.sellcash, i.orgprice, i.buycash, o.optaddprice "
	strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M "
	strSql = strSql & " JOIN db_item.dbo.tbl_item as i on M.itemid = i.itemid "
	strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on M.itemid = o.itemid and M.itemoption = o.itemoption "
	strSql = strSql & " WHERE M.idx = '"&imidx&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		sellcash	= rsget("sellcash")
		orgprice	= rsget("orgprice")
		buycash		= rsget("buycash")
		optaddprice	= rsget("optaddprice")
	Else
		getLotteComPriceParameter = ""
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

	strRst = "subscriptionId=" & lotteAuthNo
	strRst = strRst & "&strGoodsNo=" & iLotteGoodNo
	strRst = strRst & "&strReqSalePrc=" & GetRaiseValue(MustPrice/10)*10
	getLotteComPriceParameter = strRst
End Function

''//상품명 변경 파라메터 생성(롯데아이몰과 파라매타명이 다름)
Function getLotteItemnameParameter(iidx, byref iitemname, iLotteGoodNo)
	Dim strSql, chgname, strRst, newitemname, itemnameChange
	strSql = ""
	strSql = strSql & " SELECT TOP 1 M.itemid, convert(varchar(30),m.itemid) + convert(varchar(30),m.itemoption) as newCode, isnull(M.newitemname, '') as newitemname, isnull(M.itemnameChange, '') as itemnameChange "
	strSql = strSql & "	FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M "
	strSql = strSql & "	JOIN db_etcmall.[dbo].[tbl_lotteAddoption_regitem] as R on M.idx = R.midx "
	strSql = strSql & "	WHERE M.idx = '"&iidx&"' "
	strSql = strSql & "	and M.mallid = 'lotteCom' "
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
	chgname = replace(iitemname,"'","")
	chgname = replace(chgname,"~","-")
	chgname = replace(chgname,"<","[")
	chgname = replace(chgname,">","]")
	chgname = replace(chgname,"%","프로")
	chgname = replace(chgname,"[무료배송]","")
	chgname = replace(chgname,"[무료 배송]","")

	strRst = "subscriptionId=" & lotteAuthNo
	strRst = strRst & "&strGoodsNo=" & iLotteGoodNo
	strRst = strRst & "&strGoodsNm=" & Server.URLEncode(Trim(chgname))
	strRst = strRst & "&strMblGoodsNm=" & Server.URLEncode(Trim(chgname))
	strRst = strRst & "&strChgCausCont=" & Server.URLEncode("api 상품명 변경")
	getLotteItemnameParameter = strRst
End Function
%>