<%


'############################################## 실제 수행하는 API 함수 모음 ##############################################
'롯데닷컴 상품 등록
Function fnLotteComItemReg(iitemid, strParam, byRef iErrStr, iSellCash, iLotteSellYn, ibasicImage)
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
		            iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(상품등록)"
				Else
					strSql = "Select count(itemid) From db_item.dbo.tbl_lotte_regItem Where itemid='" & iitemid & "'"
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
						strSql = strSql & "	From db_item.dbo.tbl_lotte_regItem R"& VbCRLF
						strSql = strSql & " Where R.itemid='" & iitemid & "'"
						dbget.Execute(strSql)
					Else
						'// 없음 -> 신규등록
						strSql = "INSERT INTO db_item.dbo.tbl_lotte_regItem "
						strSql = strSql & " (itemid, reguserid, lotteRegdate, LotteLastUpdate, LotteTmpGoodNo, LottePrice, LotteSellYn, LotteStatCd, regImageName) values " & VbCRLF
						strSql = strSql & " ('" & iitemid & "'" & VbCRLF
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

					ArgLength = xmlDOM.getElementsByTagName("Argument").length
					Redim NameValueArr(ArgLength, 1)			'Name과 Value 각각 하나씩 name(1) + Value(1) = 2 => 0부터므로 1
					For j=0 to ArgLength
						NameValueArr(j,0) = xmlDOM.getElementsByTagName("Argument")(j).getAttribute("name")
						NameValueArr(j,1) = xmlDOM.getElementsByTagName("Argument")(j).getAttribute("value")
						If NameValueArr(j,0) = "item_list" Then
							buf_item_list = NameValueArr(j,1)
						End If
					Next

					pp = 0
			        if (buf_item_list<>"") then
			            buf_item_list = split(buf_item_list,":")
			            For k=Lbound(buf_item_list) to Ubound(buf_item_list)
			                OptDesc = splitvalue(buf_item_list(k),",",0)
			                StockQty = splitvalue(buf_item_list(k),",",1)

			                strSql = " Insert into db_item.dbo.tbl_OutMall_regedoption"
				            strSql = strSql & " (itemid,itemoption,mallid,outmallOptCode,outmallOptName,outMallSellyn,outmalllimityn,outMallLimitNo)"
				            strSql = strSql & " values("&iitemid
				            strSql = strSql & " ,''"
				            strSql = strSql & " ,'lotteCom'"
				            strSql = strSql & " ,'"&pp&"'"
				            strSql = strSql & " ,'"&html2DB(OptDesc)&"'"
				            strSql = strSql & " ,'Y'"
					        strSql = strSql & " ,'Y'"
					        strSql = strSql & " ,"&StockQty
					        strSql = strSql & ")"
					        dbget.Execute strSql, AssignedRow

					        ''옵션 코드 매칭.
					        if (AssignedRow>0) then
					            strSql = " update oP"   &VbCRLF
					            strSql = strSql & " set itemoption=O.itemoption"&VbCRLF
					            strSql = strSql & " From db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
					            strSql = strSql & "     Join db_item.dbo.tbl_item_option o"&VbCRLF
					            strSql = strSql & "     on oP.itemid=o.itemid"&VbCRLF
					            strSql = strSql & " where oP.mallid='lotteCom'"&VbCRLF
					            strSql = strSql & " and o.itemid="&iitemid&VbCRLF
					            strSql = strSql & " and oP.itemid="&iitemid&VbCRLF
					            strSql = strSql & " and op.outmallOptCode='"&pp&"'"&VbCRLF
					            strSql = strSql & " and op.outmallOptName=o.optionname"&VbCRLF
					            dbget.Execute strSql, AssignedRow
					        end if
					        pp = pp + 1
					    Next

					    strSql = " update R"   &VbCRLF
			            strSql = strSql & " set regedOptCnt=isNULL(T.CNT,0)"   &VbCRLF
			            strSql = strSql & " from db_item.dbo.tbl_lotte_regItem R"   &VbCRLF
			            strSql = strSql & " 	Join ("   &VbCRLF
			            strSql = strSql & " 		select R.itemid,count(*) as CNT from db_item.dbo.tbl_lotte_regItem R"   &VbCRLF
			            strSql = strSql & " 			Join db_item.dbo.tbl_OutMall_regedoption Ro"   &VbCRLF
			            strSql = strSql & " 			on R.itemid=Ro.itemid"   &VbCRLF
			            strSql = strSql & " 			and Ro.mallid='lotteCom'"   &VbCRLF
			            strSql = strSql & "             and Ro.itemid="&iitemid&VbCRLF
			            strSql = strSql & " 		group by R.itemid"   &VbCRLF
			            strSql = strSql & " 	) T on R.itemid=T.itemid"   &VbCRLF
			            dbget.Execute strSql
		        	end if
		       		iErrStr =  "OK||"&iitemid&"||등록성공(상품등록)"
				End If
			Set xmlDOM = Nothing
			fnLotteComItemReg= true
		Else
			iErrStr = "ERR||"&iitemid&"||LotteCom 결과 분석 중에 오류가 발생했습니다.[ERR-REG-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'롯데닷컴 판매상태변경
Function fnLotteComSellyn(iitemid, ichgSellYn, istrParam, byRef iErrStr)
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
'                response.write BinaryToText(objXML.ResponseBody, "euc-kr")
 '               response.end
				resultcode = xmlDOM.getElementsByTagName("Result").item(0).text
				resultmsg = xmlDOM.getElementsByTagName("Message").item(0).text

				If resultcode <> 1 Then
		            iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(판매상태)"
				Else
					'// 상품정보 수정
					strSql = "Update db_item.dbo.tbl_lotte_regItem " & VbCRLF
					strSql = strSql & " Set LotteLastUpdate=getdate() " & VbCRLF
					strSql = strSql & " ,LotteSellYn='" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " ,accFailCnt = 0 " & VbCRLF
					strSql = strSql & " Where itemid='" & iitemid & "'"
					dbget.Execute(strSql)

					If ichgSellYn = "N" Then
						iErrStr = "OK||"&iitemid&"||품절처리"
					ElseIf ichgSellYn = "X" Then
						strSql = ""
						strSql = strSql &" INSERT INTO [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] " & VBCRLF
						strSql = strSql &" SELECT TOP 1 'lotteCom', i.itemid, r.lotteGoodNo, r.lotteRegdate, getdate(), r.lastErrStr" & VBCRLF
						strSql = strSql &" FROM db_item.dbo.tbl_item as i " & VBCRLF
						strSql = strSql &" JOIN db_item.dbo.tbl_lotte_regitem as r on i.itemid = r.itemid " & VBCRLF
						strSql = strSql &" WHERE i.itemid = "&iitemid & VBCRLF
						dbget.Execute(strSql)

						strSql = ""
						strSql = strSql & " DELETE FROM db_item.dbo.tbl_lotte_regitem " & vbcrlf
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
						iErrStr = "OK||"&iitemid&"||판매종료"
					Else
						iErrStr = "OK||"&iitemid&"||판매중으로 변경"
					End If
				End If
			Set xmlDOM = Nothing
			fnLotteComSellyn = True
		Else
			iErrStr = "ERR||"&iitemid&"||롯데닷컴 결과 분석 중에 오류가 발생했습니다.[ERR-SELLEDIT-001]"
		End If
	Set objXML = Nothing
	on Error Goto 0
End Function

'롯데닷컴 전시상품 맵핑정보
Public Function fnLotteComStatChk(iitemid, iErrStr)
	Dim objXML,xmlDOM,strRst,resultmsg, iLotteGoodNo, strSql
	Dim strParam, iLotteTmpID, SaleStatCd, GoodsViewCount
	Dim iRbody, resultcode, lotteStatName
	On Error Resume Next
	fnLotteComStatChk = False
	iLotteTmpID = getLotteTmpItemIdByTenItemID(iitemid)

	If (iLotteTmpID = "") OR (iLotteTmpID = "전시상품") then
		iErrStr =  "ERR||"&iitemid&"||이미 전시상품 입니다.(신규상품조회)"
		Exit function
	End If

	strParam = "subscriptionId=" & lotteAuthNo & "&goods_req_no=" & iLotteTmpID

	'rw lotteAPIURL & "/openapi/getRdToPrGoodsNoApi.lotte"
	'rw strParam

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
		            iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(신규상품조회)"
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
						strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem " & VbCRLF
						strSql = strSql & " SET lastConfirmdate = getdate() "& VbCRLF
						strSql = strSql & "	,lotteStatCd='30' "
						strSql = strSql & " ,lotteGoodNo='" & iLotteGoodNo & "' "
						strSql = strSql & " WHERE itemid='" & iitemid & "'"& VbCRLF
						dbget.Execute(strSql)
					Else
						strSql = ""
						strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem " & VbCRLF
						strSql = strSql & " SET lastConfirmdate = getdate() "& VbCRLF
						strSql = strSql & "	,lotteStatCd='"&SaleStatCd&"' "& VbCRLF
						strSql = strSql & " WHERE itemid='" & iitemid & "'"& VbCRLF
						dbget.Execute(strSql)
					End If
					iErrStr =  "OK||"&iitemid&"||성공(신규상품조회) : "&lotteStatName
					fnLotteComStatChk = True
			    End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-STATCHK-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Public Function fnLotteComPrice(iitemid, istrParam, imustprice, byRef iErrStr)
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
		            iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(상품가격)"
		            fnLotteComPrice = False
				Else
				    '// 상품가격정보 수정
				    strSql = ""
	    			strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem  " & VbCRLF
	    			strSql = strSql & "	SET LotteLastUpdate=getdate() " & VbCRLF
	    			strSql = strSql & "	, LottePrice = " & imustprice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " Where itemid='" & iitemid & "'"& VbCRLF
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||수정성공(상품가격)"
					fnLotteComPrice = True
				End If
			Set xmlDOM = Nothing
		Else
			fnLotteComPrice = False
			iErrStr = "ERR||"&iitemid&"||롯데닷컴과 통신중에 오류가 발생했습니다.[ERR-PRICE-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Public Function fnLotteComChgItemname(iitemid, strParam, iErrStr)
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
		            iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(상품명)"
		            fnLotteComChgItemname = False
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem " & VbCRLF
					strSql = strSql & " SET regitemname = B.itemname "& VbCRLF
					strSql = strSql & " FROM db_item.dbo.tbl_lotte_regItem A "& VbCRLF
					strSql = strSql & " JOIN db_item.dbo.tbl_item B on A.itemid = B.itemid "& VbCRLF
					strSql = strSql & " WHERE A.itemid='" & iitemid & "'"& VbCRLF
					dbget.Execute(strSql)

					iErrStr =  "OK||"&iitemid&"||수정성공(상품명)"
					fnLotteComChgItemname = True
			    End If
			Set xmlDOM = Nothing
		else
			iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-NMEDIT-002]"
		end if
	Set objXML = Nothing
	On Error Goto 0
End Function

Public Function fnLotteComStockChk(iitemid, iErrStr)
    Dim ilottegoods_no
    Dim objXML,xmlDOM,strRst
    Dim ProdCount, buf, AssignedRow, oneProdInfo, strParam
    Dim GoodNo,ItemNo,OptDesc,DispYn,SaleStatCd,StockQty, bufopt
    Dim strSql, actCnt, SubNodes
    On Error Resume Next
    fnLotteComStockChk = False
    ilottegoods_no = getLotteGoodno(iitemid)

    Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    strParam = "?subscriptionId=" & lotteAuthNo					'롯데닷컴 인증번호	(*)
	    strParam = strParam & "&search_gubun=goods_no"
		strParam = strParam & "&search_text=" & ilottegoods_no		'롯뎃닷컴 상품번호	(*)

		objXML.Open "POST", lotteAPIURL & "/openapi/searchStockList.lotte"&strParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML replace(buf,"&","＆")
				ProdCount   = Trim(xmlDOM.getElementsByTagName("ProdCount").item(0).text)   '' 단품 갯수

'response.write buf
'response.end
				If (ProdCount <> "") Then
					Set oneProdInfo = xmlDOM.getElementsByTagName("ProdInfo")
					' strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and itemoption='')"
					' strSql = strSql & " BEGIN"
					' strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and itemoption=''"
					' strSql = strSql & " END"
					' dbget.Execute strSql

					' strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and Len(outmalloptCode)>6)"
					' strSql = strSql & " BEGIN"
					' strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and Len(outmalloptCode)>6"
					' strSql = strSql & " END"
					' dbget.Execute strSql

					'2019/03/19 regedoption 초기화
					'strSql = "DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and (itemoption='' or Len(outmalloptCode)>6)"
					strSql = "DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" "
					dbget.Execute strSql

					For each SubNodes in oneProdInfo
						SaleStatCd = ""
						GoodNo	    = Trim(SubNodes.getElementsByTagName("GoodNo").item(0).text)
						ItemNo	    = Trim(SubNodes.getElementsByTagName("ItemNo").item(0).text)        '' 단품코드 (숫자 0,1,2,)
						OptDesc	    = Trim(SubNodes.getElementsByTagName("OptDesc").item(0).text)
						DispYn	    = Trim(SubNodes.getElementsByTagName("DispYn").item(0).text)         ''N:안함 Y:전시
						SaleStatCd	= Trim(SubNodes.getElementsByTagName("SaleStatCd").item(0).text) ''판매진행, 판매종료, 품절'
						StockQty	= Trim(SubNodes.getElementsByTagName("StockQty").item(0).text)

						If SaleStatCd <> "판매종료" Then
							OptDesc = replace(OptDesc,"＆","&")
							If (SaleStatCd<>"판매진행") Then
								DispYn="N"
							Else
								DispYn="Y"
							End If

							If (StockQty = "null") Then
								StockQty="0"
							End If

							bufopt = OptDesc

							If InStr(bufopt,",") > 0 Then
								If (splitValue(bufopt,",",0) <> "") Then
									OptDesc = splitValue(splitValue(bufopt,",",0),":",1)
								End If

								If (splitValue(bufopt,",",1) <> "") Then
									OptDesc = OptDesc+","+splitValue(splitValue(bufopt,",",1),":",1)
								End If

								If (splitValue(bufopt,",",2) <> "") Then
									OptDesc = OptDesc+","+splitValue(splitValue(bufopt,",",2),":",1)
								End If
							Else
								OptDesc = splitValue(OptDesc,":",1)
							End If

							strSql = " Insert into db_item.dbo.tbl_OutMall_regedoption"
							strSql = strSql & " (itemid,itemoption,mallid,outmallOptCode,outmallOptName,outMallSellyn,outmalllimityn,outMallLimitNo)"
							strSql = strSql & " values("&iitemid
							strSql = strSql & " ,'"&ItemNo&"'" ''임시로 롯데 코드 넣음 //2013/04/01
							strSql = strSql & " ,'lotteCom'"
							strSql = strSql & " ,'"&ItemNo&"'"
							strSql = strSql & " ,'"&html2DB(OptDesc)&"'"
							strSql = strSql & " ,'"&DispYn&"'"
							strSql = strSql & " ,'Y'"
							strSql = strSql & " ,"&StockQty
							strSql = strSql & ")"
							dbget.Execute strSql, AssignedRow

							If (AssignedRow > 0) Then
								strSql = " update oP"   &VbCRLF
								strSql = strSql & " set itemoption=O.itemoption"&VbCRLF
								strSql = strSql & " From db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
								strSql = strSql & "     Join db_item.dbo.tbl_item_option o"&VbCRLF
								strSql = strSql & "     on oP.itemid=o.itemid"&VbCRLF
								strSql = strSql & " where oP.mallid='lotteCom'"&VbCRLF
								strSql = strSql & " and o.itemid="&iitemid&VbCRLF
								strSql = strSql & " and oP.itemid="&iitemid&VbCRLF
								strSql = strSql & " and op.outmallOptCode='"&ItemNo&"'"&VbCRLF
								strSql = strSql & " and Replace(Replace(op.outmallOptName,' ',''),':','')=Replace(Replace(o.optionname,' ',''),':','')"&VbCRLF
								dbget.Execute strSql, AssignedRow
							End If
						End If
					Next

					Dim currOptCnt
					strSql = ""
					strSql = strSql & " SELECT COUNT(*) as cnt FROM db_item.dbo.tbl_item_option WHERE itemid ="&iitemid
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If Not(rsget.EOF or rsget.BOF) Then
						currOptCnt = rsget("cnt")
					End If
					rsget.Close

					If currOptCnt > 0 Then
						strSql = ""
						strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and len(itemoption) <= 2 "
						dbget.Execute strSql
					End IF

					strSql = " update R"   &VbCRLF
					strSql = strSql & " set regedOptCnt=isNULL(T.optSellYCNT,0)"   &VbCRLF  ''regedOptCnt => optSellYCNT
					strSql = strSql & " ,lastStatcheckdate=getdate()"&VbCRLF 				''추가
					strSql = strSql & " from db_item.dbo.tbl_lotte_regItem R"   &VbCRLF
					strSql = strSql & " 	Join ("   &VbCRLF
					strSql = strSql & " 		select R.itemid,count(*) as CNT "
					strSql = strSql & " 		, sum(CASE WHEN outmallOptName<>'' THEN 1 ELSE 0 END) as regedOptCnt"
					strSql = strSql & " 		, sum(CASE WHEN outmallOptName<>'' and [outmallSellyn]='Y' and (outmalllimityn='N' or (outmalllimityn='Y' and outmalllimitno>0)) THEN 1 ELSE 0 END) as optSellYCNT"
					strSql = strSql & "			from db_item.dbo.tbl_lotte_regItem R"   &VbCRLF
					strSql = strSql & " 			Join db_item.dbo.tbl_OutMall_regedoption Ro"   &VbCRLF
					strSql = strSql & " 			on R.itemid=Ro.itemid"   &VbCRLF
					strSql = strSql & " 			and Ro.mallid='lotteCom'"   &VbCRLF
					strSql = strSql & "             and Ro.itemid="&iitemid&VbCRLF
					strSql = strSql & " 		group by R.itemid"   &VbCRLF
					strSql = strSql & " 	) T on R.itemid=T.itemid"   &VbCRLF
					dbget.Execute strSql

					iErrStr =  "OK||"&iitemid&"||성공(재고조회)"
					fnLotteComStockChk = true
				End If
			Set xmlDOM = Nothing
		Else
		    iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-STOCKCHK-001]"
		    fnLotteComStockChk = false
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

''롯데닷컴 상품정보 수정
Function fnLotteComInfoEdit(iitemid, strParam, byRef iErrStr, isVer2)
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
					iErrStr =  "OK||"&iitemid&"||성공(상품정보)"
					fnLotteComInfoEdit = True
				Else
		            iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(상품정보)"
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

Function fnLotteComInfodivEdit(iitemid, strParam, byRef iErrStr)
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
					iErrStr =  "OK||"&iitemid&"||성공(품목정보)"
					fnLotteComInfodivEdit = True
				Else
		            iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(품목정보)"
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

Public Function fnLotteComImageEdit(iitemid, strParam, byRef iErrStr)
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
					strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem " & VbCRLF
					strSql = strSql & " SET regImageName = B.basicimage "& VbCRLF
					strSql = strSql & " FROM db_item.dbo.tbl_lotte_regItem A "& VbCRLF
					strSql = strSql & " JOIN db_item.dbo.tbl_item B on A.itemid = B.itemid "& VbCRLF
					strSql = strSql & " WHERE A.itemid='" & iitemid & "'"& VbCRLF
					dbget.Execute(strSql)

					iErrStr =  "OK||"&iitemid&"||수정성공(상품명)"
					fnLotteComChgItemname = True


					iErrStr =  "OK||"&iitemid&"||성공(이미지수정)"
					fnLotteComImageEdit = True
				Else
		            iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(이미지수정)"
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

Function fnLotteComCateGory(iitemid, strParam, byRef iErrStr)
	Dim objXML,xmlDOM,strRst,iMessage, strSql
	Dim resultcode, resultmsg

	On Error Resume Next
	fnLotteComCateGory = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", lotteAPIURL & "/openapi/updateGoodsCategoryOpenApi.lotte", false
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
					strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem " & VbCRLF
					strSql = strSql & " SET lastcateChgDate = getdate() "& VbCRLF
					strSql = strSql & " WHERE itemid='" & iitemid & "'"& VbCRLF
					dbget.Execute(strSql)

					iErrStr =  "OK||"&iitemid&"||성공(카테고리)"
					fnLotteComInfoEdit = True
				Else
		            iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(카테고리)"
		            fnLotteComInfoEdit = False
			    End If
			Set xmlDOM = Nothing
		Else
			fnLotteComCateGory = False
			iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-PoomEDIT-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function


''전시상품 조회
Function fnCheckLotteComItemStat(iitemid, byRef iErrStr, iLottegoodNo)
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
					sqlstr = sqlstr & " SET regitemname='"&html2db(iGoodsNm)&"'"  & VbCRLF
					IF (LotteSellyn <> "") then
						sqlstr = sqlstr & " ,LotteSellyn='"&LotteSellyn&"'"
					ENd IF
					sqlstr = sqlstr & " ,lastStatCheckDate=getdate()"
					sqlstr = sqlstr & " From db_item.dbo.tbl_lotte_regItem R" & VbCRLF
					sqlstr = sqlstr & " where R.itemid="&iitemid & VbCRLF
					dbget.Execute sqlstr,assignedRow

			    	iErrStr =  "OK||"&iitemid&"||성공(전시상품조회)"
					fnCheckLotteComItemStat = True
			    Else
			    	resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text
			    	iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(전시상품조회)"
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

'롯데닷컴 상품목록조회 등록일기준 CHKITEMLIST
Function fnLotteGoodsList(styyyymmdd,edyyyymmdd, byRef iErrStr)
    Dim objXML, xmlDOM, strRst, resultmsg
	Dim strParam, SaleStatCd, GoodsCount, iSalePrc, iGoodsNm, GoodsRegDtime, DispYn, iGoodsNo
	Dim iRbody, LotteSellyn, sqlStr, assignedRow
	Dim oneGoodsInfo, SubNodes
	Dim regDtKey : regDtKey=LEFT(NOW(),10) & " " &FormatDateTime(NOW(),4)&":"&RIGHT("0"&second(time),2)
    fnLotteGoodsList = False
'rw regDtKey
'response.end
	strParam = "?subscriptionId=" & lotteAuthNo					'롯데닷컴 인증번호	(*)
	strParam = strParam & "&strSearchStrtDtime="&styyyymmdd
	strParam = strParam & "&strSearchEndDtime="&edyyyymmdd
	''strParam = strParam & "&selDispYn=T "
	''strParam = strParam & "&selSaleStatCd="


	'on Error resume next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", lotteAPIURL & "/openapi/searchGoodsListOpenApi.lotte" & strParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				iRbody = replace(iRbody,"&","@@amp@@")   '' <![CDATA[]]> 로 안 묶여옴. 상품명에 < > 있음..
				iRbody = replace(iRbody,"<GoodsNm>","<GoodsNm><![CDATA[")
				iRbody = replace(iRbody,"</GoodsNm>","]]></GoodsNm>")
				xmlDOM.LoadXML iRbody
'rw iRbody
'response.end
				GoodsCount = xmlDOM.getElementsByTagName("GoodsCount").item(0).text  ''결과수

				If (GoodsCount >= "0") Then
					Set oneGoodsInfo = xmlDOM.getElementsByTagName("GoodsInfoList")
					For each SubNodes in oneGoodsInfo
						iGoodsNo	= SubNodes.getElementsByTagName("GoodsNo").item(0).text
						SaleStatCd  = SubNodes.getElementsByTagName("SaleStatCd").item(0).text
						iSalePrc	= SubNodes.getElementsByTagName("SalePrc").item(0).text
						iGoodsNm	= SubNodes.getElementsByTagName("GoodsNm").item(0).text
						iGoodsNm	= replace(iGoodsNm,"@@amp@@","&")
						iGoodsNm	= Replace(iGoodsNm,"&gt;",">")
						iGoodsNm	= Replace(iGoodsNm,"&lt;","<")
						iGoodsNm	= Replace(iGoodsNm,"&nbsp;"," ")
						iGoodsNm	= Replace(iGoodsNm,"&amp;","&")

						GoodsRegDtime = SubNodes.getElementsByTagName("GoodsRegDtime").item(0).text
						DispYn 		 = SubNodes.getElementsByTagName("DispYn").item(0).text

						If (SaleStatCd="10") Then
							LotteSellyn = "Y"
						ElseIf (SaleStatCd="20") Then
							LotteSellyn = "N"
						ElseIf (SaleStatCd="30") Then
							LotteSellyn = "X"
						End If

						''rw  iGoodsNo&"|"&SaleStatCd&"|"&iSalePrc&"|"&iGoodsNm&"|"&GoodsRegDtime&"|"&DispYn
						sqlstr = "exec [db_temp].[dbo].[usp_TEN_OutMall_CheckRegItemLIST] 'lotteCom','"&iGoodsNo&"','"&LotteSellyn&"',"&iSalePrc&",'"&replace(iGoodsNm,"'","''")&"','"&GoodsRegDtime&"','"&DispYn&"','"&regDtKey&"'"
						dbget.Execute sqlstr

						' sqlstr = ""
						' sqlstr = sqlstr & " Update R" & VbCRLF
						' sqlstr = sqlstr & " SET regitemname='"&html2db(iGoodsNm)&"'"  & VbCRLF
						' IF (LotteSellyn <> "") then
						' 	sqlstr = sqlstr & " ,LotteSellyn='"&LotteSellyn&"'"
						' ENd IF
						' sqlstr = sqlstr & " ,lastStatCheckDate=getdate()"
						' sqlstr = sqlstr & " From db_item.dbo.tbl_lotte_regItem R" & VbCRLF
						' sqlstr = sqlstr & " where R.itemid="&iitemid & VbCRLF
						' dbget.Execute sqlstr,assignedRow
					Next
					Set oneGoodsInfo = Nothing

					sqlstr = "exec [db_temp].[dbo].[usp_TEN_OutMall_CheckRegItemLIST_MAP] 'lotteCom','"&regDtKey&"'"
					dbget.Execute sqlstr

					iErrStr =  "OK||"&styyyymmdd&"-"&edyyyymmdd&"||성공(전시상품조회)-"&GoodsCount&"건"
					fnLotteGoodsList = True
			    Else
			    	resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text
			    	iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(전시상품조회)"
		            fnLotteGoodsList = False
			    End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-ItemChk-001]"
			fnLotteGoodsList = False
		End If
	Set objXML = Nothing
	'on Error Goto 0
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
	ElseIf ichgSellYn = "X" Then
		strRst = strRst & "&sale_stat_cd=30"
	End If
	getLotteComSellynParameter = strRst
End Function

Function getLotteTmpItemIdByTenItemID(iitemid)
	Dim sqlStr, retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT lotteTmpGoodNo, isnull(lotteGoodNo,'') as lotteGoodNo " & VBCRLF
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_lotte_regItem" & VBCRLF
	sqlStr = sqlStr & " WHERE itemid = "&iitemid & VBCRLF
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
Function getLotteComPriceParameter(iitemid, iLotteGoodNo, MustPrice)
	Dim strRst, strSql
	Dim sellcash, orgprice, buycash
	Dim GetTenTenMargin

	strRst = "subscriptionId=" & lotteAuthNo
	strRst = strRst & "&strGoodsNo=" & iLotteGoodNo
	strRst = strRst & "&strReqSalePrc=" & MustPrice
	getLotteComPriceParameter = strRst
End Function

''//상품명 변경 파라메터 생성(롯데아이몰과 파라매타명이 다름)
Function getLotteItemnameParameter(iitemid, byref iitemname, iLotteGoodNo)
	Dim strSql, chgname, strRst
	strSql = ""
	strSql = strSql & " SELECT TOP 1 r.itemid, i.ItemName "
	strSql = strSql & "	FROM db_item.dbo.tbl_lotte_regItem r "
	strSql = strSql & "	JOIN db_item.dbo.tbl_item i on r.itemid = i.itemid "
	strSql = strSql & "	WHERE i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.Eof Then
		iitemname = rsget("ItemName")
	End If
	rsget.close

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

Function getLotteCategoryParameter(iitemid, iLotteGoodNo)
	Dim strSql, chgname, strRst, i, cateRst, ogrpCode
	strSql = ""
	strSql = strSql & " SELECT top 100 c.groupCode, m.dispNo, c.disptpcd "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_item.dbo.tbl_lotte_cate_mapping as m on i.cate_large = m.tenCateLarge and i.cate_mid = m.tenCateMid and i.cate_small = m.tenCateSmall "
	strSql = strSql & " JOIN db_temp.dbo.tbl_lotte_Category c on m.DispNO = c.DispNO "
	strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
	strSql = strSql & " ORDER BY (CASE WHEN c.disptpcd='12' THEN 'ZZ' ELSE c.disptpcd END) DESC "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
	    ogrpCode = rsget("groupCode")
		i = 0
		Do until rsget.EOF
			If (rsget("disptpcd")="12") then                        ''전문몰 카테고리를 기본으로 하라고.. /2012/06/14
			    cateRst = cateRst & "&disp_no_b=" & rsget("dispNo")		'기본 전시카테고리
			Else
			    IF (ogrpCode=rsget("groupCode")) then
				    cateRst = cateRst & "&disp_no=" & rsget("dispNo") 	'추가 전시카테고리
				End IF
		    End If
			rsget.MoveNext
			i = i + 1
		Loop
	End If
	rsget.Close

	strRst = "subscriptionId=" & lotteAuthNo
	strRst = strRst & "&strGoodsNo=" & iLotteGoodNo
	strRst = strRst & cateRst
	strRst = strRst & "&strChgCausCont=" & Server.URLEncode("api 카테고리 변경")
	getLotteCategoryParameter = strRst
End Function

Function getOptCntCompare(iitemid)
	Dim strSql
	Dim oCnt, rCnt

	strSql = "SELECT COUNT(*) as oCnt FROM db_item.dbo.tbl_item_option WHERE itemid = '"&iitemid&"'"
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		oCnt = rsget("oCnt")
	End If
	rsget.Close

	strSql = "SELECT COUNT(*) as rCnt FROM db_item.dbo.tbl_outmall_regedOption WHERE itemid = '"&iitemid&"' and mallid = 'lotteCom' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		rCnt = rsget("rCnt")
	End If
	rsget.Close

	If oCnt > 0 AND rCnt = 0 Then
		getOptCntCompare = "Y"
	Else
		getOptCntCompare = "N"
	End If
End Function

Function getUseOption(iitemid)
	Dim strSql, cnt
	strSql = "SELECT COUNT(*) as cnt FROM db_item.dbo.tbl_outmall_regedOption WHERE itemid = '"&iitemid&"' and mallid = 'lotteCom' and outmallsellyn = 'Y' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		cnt = rsget("cnt")
	End If
	rsget.Close

	If cnt > 0 Then
		getUseOption = "Y"
	Else
		getUseOption = "N"
	End If
End Function
%>
