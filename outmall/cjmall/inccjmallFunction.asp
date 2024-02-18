<%
Public cjMallAPIURL
Dim isCJ_DebugMode : isCJ_DebugMode = True

IF application("Svr_Info")="Dev" THEN
	cjMallAPIURL = "http://210.122.101.154:8110/IFPAServerAction.action"	'' 테스트서버
'	cjMallAPIURL = "http://apiqa.cjmall.com:8600/IFPAServerAction.action"	'' 테스트서버
	'cjMallAPIURL = "http://210.122.101.154:8210/IFPAServerAction.action"	'' 개편될 CJ QA서버 URL
Else
	cjMallAPIURL = "http://api.cjmall.com/IFPAServerAction.action"			'' 실서버
End if
'############################################## 실제 수행하는 API 함수 모음 ##############################################
'품절 수행 함수
Public Function fnCJMallSellyn(iitemid, ichgSellYn, istrParam, byRef iErrStr)
    Dim resultcode, resultmsg
    Dim objXML, xmlDOM, strSql
    On Error Resume Next
    fnCJMallSellyn = False
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(istrParam)

		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
			'response.write objXML.ResponseText
			'response.end
			resultcode		= xmlDOM.getElementsByTagName("ns1:successYn").item(0).text
			resultmsg		= replace(xmlDOM.getElementsByTagName("ns1:errorMsg").item(0).text, "'", "")

			If Err <> 0 Then
				If (IsAutoScript) Then
					iErrStr = "ERR||"&iitemid&"||CJMall 결과 분석 중에 오류가 발생했습니다.[ERR-SELLEDIT-001]"
				Else
					iErrStr = "ERR||"&iitemid&"||CJMall 결과 분석 중에 오류가 발생했습니다.[ERR-SELLEDIT-001]"
				End If
				Set objXML = Nothing
			    Set xmlDOM = Nothing
			    On Error Goto 0
			    Exit Function
		    End If

			If resultcode = "true" Then		'성공(200)
				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_cjmall_regItem " & VbCRLF
				strSql = strSql & " SET cjmallSellYn = '"&ichgSellYn&"'" & VbCRLF
				strSql = strSql & " ,cjmallLastUpdate = getdate()" & VbCRLF
				strSql = strSql & " ,accFailCNT=0" & VbCRLF
				strSql = strSql & " WHERE itemid = "&iitemid
				dbget.Execute(strSql)
				fnCJMallSellyn = true
				Set objXML = Nothing
				Set xmlDOM = Nothing
				If ichgSellYn = "N" Then
					iErrStr = "OK||"&iitemid&"||품절처리"
				Else
					iErrStr = "OK||"&iitemid&"||판매중으로 변경"
				End If
			Else						'실패(E)
				iErrStr = "ERR||"&iitemid&"||"&resultmsg
			End If
		Else
			iErrStr = "ERR||"&iitemid&"||CJMall 결과 분석 중에 오류가 발생했습니다.[ERR-SELLEDIT-001]"
		End If
	On Error Goto 0
End Function

'등록
Function fnCJMallItemReg(iitemid, istrParam, byRef iErrStr, iSellCash, icjmallSellYn, ilimityn, ilimitno, ilimitsold, iitemname)
	Dim resultcode, resultmsg
	Dim objXML, xmlDOM, strSql, goodsCd
	On Error Resume Next
	fnCJMallItemReg = False
'response.write istrParam
'response.end
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(istrParam)
'rw objXML.Status
'response.write istrParam
'response.end
		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
			If (session("ssBctID")="kjy8517") Then
				rw "REQ : <textarea cols=40 rows=10>"&istrParam&"</textarea>"
				rw "RES : <textarea cols=40 rows=10>"&BinaryToText(objXML.ResponseBody, "euc-kr")&"</textarea>"
			End If
			goodsCd			= xmlDOM.getElementsByTagName("ns1:itemCd").item(0).text
			resultcode		= xmlDOM.getElementsByTagName("ns1:successYn").item(0).text
			resultmsg		= replace(xmlDOM.getElementsByTagName("ns1:errorMsg").item(0).text, "'", "")

			If resultcode = "true" Then		'성공(200)
				strSql = ""
				strSql = strSql & " UPDATE R"
				strSql = strSql & "	Set cjmallLastUpdate = getdate() "
				strSql = strSql & "	, cjmallPrdNo = '" & goodsCd & "'"
				strSql = strSql & "	, cjmallPrice = " &iSellCash
				strSql = strSql & "	, accFailCnt = 0"
				strSql = strSql & "	, cjmallRegdate = isNULL(cjmallRegdate, getdate())"
			    strSql = strSql & "	, cjmallStatCd=(CASE WHEN isNULL(cjmallStatCd, -1) < 3 then 3 ELSE cjmallStatCd END)"	'등록완료(임시)
				strSql = strSql & "	From db_item.dbo.tbl_cjmall_regItem R"
				strSql = strSql & " Where R.itemid = '" & iitemid & "'"
				dbget.Execute(strSql)
				fnCJMallItemReg = true
				Set objXML = Nothing
				Set xmlDOM = Nothing
				iErrStr =  "OK||"&iitemid&"||등록성공(상품등록)"
			Else						'실패(E)
				fnCJMallItemReg = false
			    iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(상품등록)"
				Set objXML = Nothing
				Set xmlDOM = Nothing
			    Exit Function
			End If
		Else
			iErrStr = "ERR||"&iitemid&"||CJMall 결과 분석 중에 오류가 발생했습니다.[ERR-REG-001]"
		End If
	On Error Goto 0
End Function

'상품 조회
Function fnCJMallStatChk(iitemid, istrParam, byRef iErrStr, ichkXML)
	Dim objXML, xmlDOM, strSql
	Dim AssignedRow, Nodes, SubNodes
	Dim OverLapNo, SelOK, AssignedItemCnt
	Dim XitemCd, Xstatus, XslCls, XHapvpn, Xvpn, XunitCd, Xitemcode
	Dim uprItemNm, itemNm, slprc,exLeadtm, zClassId, packInd, purchvat, taxYn

	SelOK = 0
	AssignedItemCnt = 0
	fnCJMallStatChk = false
	On Error Resume Next
	If ichkXML = "Y" Then
		response.write replace(istrParam, "<?xml", "<aaaaaa")
'		response.end
	End If

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(istrParam)

		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				If ichkXML = "Y" Then
					response.write replace(BinaryToText(objXML.ResponseBody, "euc-kr"), "<?xml", "<aaaaaa")
	'				response.end
				End If

				Set Nodes = xmlDOM.getElementsByTagName("ns1:unit")
				If (Not (xmlDOM is Nothing)) Then
					For each SubNodes in Nodes
						XitemCd = SubNodes.getElementsByTagName("ns1:itemCd")(0).Text		'판매코드
						Xstatus = SubNodes.getElementsByTagName("ns1:status")(0).Text		'결재상태
						XslCls 	= SubNodes.getElementsByTagName("ns1:slCls")(0).Text		'판매구분(상태)
						XHapvpn	= SubNodes.getElementsByTagName("ns1:vpn")(0).Text			'업체상품코드
						XunitCd = SubNodes.getElementsByTagName("ns1:unitCd")(0).Text		'단품코드

						uprItemNm= SubNodes.getElementsByTagName("ns1:uprItemNm")(0).Text	'판매상품명
						itemNm  = SubNodes.getElementsByTagName("ns1:itemNm")(0).Text		'단품상세
						slprc   = SubNodes.getElementsByTagName("ns1:slprc")(0).Text		'판매가
						exLeadtm= SubNodes.getElementsByTagName("ns1:exLeadtm")(0).Text		'리드타임(L/T)
						packInd = SubNodes.getElementsByTagName("ns1:packInd")(0).Text
						purchvat = SubNodes.getElementsByTagName("ns1:purchvat")(0).Text 	'매입가 vat포함?
						taxYn    = SubNodes.getElementsByTagName("ns1:taxYn")(0).Text

						Xvpn 		= Split(XHapvpn, "_")(0)
						Xitemcode 	= replace(Split(XHapvpn, "_")(1), "Q", "")

						'1.tbl_OutMall_regedoption 테이블에 있으면 업데이트 없으면 인서트 시키기
						strSql = ""
						strSql = strSql & " IF Exists(SELECT * FROM db_item.dbo.tbl_OutMall_regedoption WHERE itemid="&iitemid&" and mallid = '"&CMALLNAME&"' and itemoption = '"&Xitemcode&"' )"
						strSql = strSql & " BEGIN "
						strSql = strSql & " UPDATE db_item.dbo.tbl_OutMall_regedoption SET "
						strSql = strSql & " outmallsellyn='"&CHKIIF(XslCls="I","N","Y")&"'"
						If (Xitemcode <> "0000") Then
							strSql = strSql & " , outmallOptName='"&html2DB(itemNm)&"'"
						End If
						strSql = strSql & " , outmallAddPrice="&slprc
						strSql = strSql & " , outmallleadTime='"&exLeadtm&"'"
						strSql = strSql & " , checkdate = getdate() "
						strSql = strSql & " , outmallsuppPrc="&purchvat*1.1
						strSql = strSql & " , outmallOptCode='"&XunitCd&"'"
						strSql = strSql & " WHERE itemid = '"&Xvpn&"' and itemoption = '"&Xitemcode&"' "
						strSql = strSql & " and mallid='"&CMALLNAME&"'"
						strSql = strSql & " END ELSE "
						strSql = strSql & " BEGIN "
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
						strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice,outmallleadTime, outmallsuppPrc,lastUpdate, checkdate) "
						strSql = strSql & " VALUES "
						strSql = strSql & " ('"&Xvpn&"'"
						strSql = strSql & ",  '"&Xitemcode&"'"
						strSql = strSql & ", '"&CMALLNAME&"'"
						strSql = strSql & ", '"&XunitCd&"'"
						strSql = strSql & ", '"&html2db(CHKIIF(Xitemcode<>"0000", itemNm, "단일상품"))&"'"
						strSql = strSql & ", '"&CHKIIF(XslCls="I", "N", "Y")&"'"
						strSql = strSql & ", '"&"N"&"'"
						strSql = strSql & ", '0'"
						strSql = strSql & ", '"&slprc&"'"
						strSql = strSql & ", '"&exLeadtm&"'"
						strSql = strSql & ", "&purchvat*1.1&""
						strSql = strSql & ", getdate() "
						strSql = strSql & ", getdate()) "
						strSql = strSql & " END "
						dbget.Execute strSql, AssignedRow
						SelOK = SelOK + 1
					Next

					'2.tbl_cjmall_regitem 테이블의 cjmallStatCd, lastStatCheckDate, cjmallsellyn, cjMallPrice, regedOptCnt 등 수정하기
					'2015-01-06 김진영 cjmallprdno도 수정 => cjmallprdno가 null인것 발견!
					'2019-05-28 XitemCd <> "" 조건 추가
					If XitemCd <> "" Then
						strSql = ""
						strSql = strSql & " UPDATE R " & VBCRLF
						strSql = strSql & " SET cjmallregdate = isNULL(cjmallregdate,getdate())" & VBCRLF
						strSql = strSql & " ,cjmallStatCd = 7" & VBCRLF
						strSql = strSql & " ,lastStatCheckDate = getdate()" & VBCRLF                               ''수정
		'				strSql = strSql & " ,cjmallsellyn=(CASE WHEN T.SellCNT>0 THEN 'Y' ELSE 'N' END)"
						strSql = strSql & " ,cjMallPrice=(CASE WHEN T.mayItemPrice>0 then T.mayItemPrice ELSE R.cjMallPrice END)" & VBCRLF
						strSql = strSql & " ,regedOptCnt=isNULL(T.regedOptCnt,0)" & VBCRLF
						strSql = strSql & " ,cjmallprdno="&XitemCd & VBCRLF
						strSql = strSql & " from db_item.dbo.tbl_cjmall_regItem R" & VBCRLF
						strSql = strSql & " 	Join (" & VBCRLF
						strSql = strSql & " 		select itemid, count(*) as optCNT" & VBCRLF
						strSql = strSql & " 		, sum(CASE WHEN outmallsellyn='Y' THEN 1 ELSE 0 END) as SellCNT" & VBCRLF
						strSql = strSql & " 		, min(outmallAddPrice) as mayItemPrice" & VBCRLF
						strSql = strSql & " 		, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt" & VBCRLF
						strSql = strSql & " 		from db_item.dbo.tbl_OutMall_regedoption" & VBCRLF
						strSql = strSql & " 		where itemid="&iitemid&"" & VBCRLF
						strSql = strSql & " 		and mallid='cjmall'" & VBCRLF
						strSql = strSql & " 		group by itemid" & VBCRLF
						strSql = strSql & " 	) T on R.itemid=T.itemid" & VBCRLF
						strSql = strSql & " where R.itemid="&iitemid&""
						dbget.Execute strSql
						AssignedItemCnt = AssignedItemCnt + 1
					End If

					If SelOK = 0 Then
						If (iitemid <> "") Then
							''체크실패시 반복되지 않도록
							strSql = ""
							strSql = strSql & " update R"
							strSql = strSql & " set lastStatCheckDate = getdate()" & VBCRLF
							strSql = strSql & " from db_item.dbo.tbl_cjmall_regItem R" & VBCRLF
							strSql = strSql & " where itemid="&iitemid
							dbget.Execute strSql
						End If
						'iErrStr =  "ERR||"&iitemid&"||검색 결과 없음(상품조회)"
						iErrStr =  "OK||"&iitemid&"||검색 결과 없음(상품조회)"
						fnCJMallStatChk = false
					Else
						iErrStr =  "OK||"&iitemid&"||성공(상품조회)"
						fnCJMallStatChk = true
					End If
				End If
				Set Nodes = Nothing
			Set xmlDOM = Nothing
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'가격 수정
Function fnCJMallOptionSellPriceEdit(iitemid, byRef iErrStr, strParam)
    Dim objXML, xmlDOM, resultcode, resultmsg
    Dim strRst, sqlStr
    Dim Nodes, SubNodes
    Dim typeCD, itemCD_ZIP, newUnitRetail, newUnitCost
    Dim AssignedItemCnt : AssignedItemCnt = 0
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(strParam)
'rw "AAAAAaaaa"
'rw objXML.Status & "!!!"
'response.end
	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
			If (session("ssBctID")="kjy8517") Then
				rw "REQ : <textarea cols=40 rows=10>"&strParam&"</textarea>"
				rw "RES : <textarea cols=40 rows=10>"&BinaryToText(objXML.ResponseBody, "euc-kr")&"</textarea>"
			End If
'  If iitemid = "1027415" Then
'  	response.write replace(strParam, "<?xml", "<aaaaaaa")
'  	rw "----------END"
'  	response.write replace(BinaryToText(objXML.ResponseBody, "euc-kr"), "<?xml", "<aaaaaaa")
'  	response.end
'  End If
		Set Nodes = xmlDOM.getElementsByTagName("ns1:itemPrices")
		If (Not (Nodes is Nothing)) Then
			For each SubNodes in Nodes
				If (Not (SubNodes is Nothing)) Then
					resultcode		= SubNodes.getElementsByTagName("ns1:successYn").item(0).text
					resultmsg		= SubNodes.getElementsByTagName("ns1:errorMsg").item(0).text
					If resultcode = "true" Then		'성공(200)
						typeCD			= SubNodes.getElementsByTagName("ns1:typeCD").item(0).text
						itemCD_ZIP		= SubNodes.getElementsByTagName("ns1:itemCD_ZIP").item(0).text
						newUnitRetail	= SubNodes.getElementsByTagName("ns1:newUnitRetail").item(0).text
						newUnitCost		= SubNodes.getElementsByTagName("ns1:newUnitCost").item(0).text
						If (typeCD = "01") Then
						    sqlStr = ""
						    sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_cjmall_regItem SET cjmallLastUpdate = getdate() "
							sqlStr = sqlStr & " ,cjmallprice = '"&newUnitRetail&"'"
							sqlStr = sqlStr & " ,accFailCnt = 0"
							sqlStr = sqlStr & " ,lastpriceCheckDate = getdate()"
						    sqlStr = sqlStr & " WHERE itemid = '"&iitemid&"'"
						    dbget.Execute sqlStr
						ElseIf (typeCD = "02") Then
						    sqlStr = "UpDate R "
						    sqlStr = sqlStr & " SET outmallAddPrice="&newUnitRetail
						    sqlStr = sqlStr & " ,lastupdate=getdate()"
						    sqlStr = sqlStr & " ,checkdate=getdate()"
						    sqlStr = sqlStr & " from db_item.dbo.tbl_OutMall_regedoption R"
						    sqlStr = sqlStr & "  where mallid='"&CMALLNAME&"'"
						    sqlStr = sqlStr & "  and itemid="&iitemid
						    sqlStr = sqlStr & "  and outmallOptCode='"&itemCD_ZIP&"'"
						    dbget.Execute sqlStr
						End If
						iErrStr =  "OK||"&iitemid&"||수정성공(상품가격)"
						fnCJMallOptionSellPriceEdit = True
					Else
						If instr(resultmsg, "NullPointerException") > 0 Then
							resultmsg = "전문형식 잘못됨(by 김진영)"
						End If
						iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(상품가격)"
						fnCJMallOptionSellPriceEdit = False
					End If
					Set objXML = Nothing
					Set xmlDOM = Nothing
				End If
			Next
		Else
			iErrStr = "ERR||"&iitemid&"||CJMall과 통신중에 오류가 발생했습니다.[ERR-PRICE-001]"
			fnCJMallOptionSellPriceEdit = false
		End If
	Else
		iErrStr = "ERR||"&iitemid&"||CJMall과 통신중에 오류가 발생했습니다.[ERR-PRICE-002]"
		fnCJMallOptionSellPriceEdit = false
	End If
	On Error Goto 0
End Function

'옵션 수량 수정
Function fnCJMallOptionQTYEdit(iitemid, byRef iErrStr, strParam)
	Dim objXML, xmlDOM, resultcode, resultmsg
	Dim strRst, sqlStr
	Dim Nodes, SubNodes
	Dim unitCd, strDt, endDt, availSupQty
	Dim AssignedNotItemCnt : AssignedNotItemCnt = 0
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(strParam)
	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
			'response.write objXML.ResponseText
			'response.end

		Set Nodes = xmlDOM.getElementsByTagName("ns1:ltSupplyPlans")
		If (Not (Nodes is Nothing)) Then
			For each SubNodes in Nodes
				If (Not (SubNodes is Nothing)) Then
					resultcode		= SubNodes.getElementsByTagName("ns1:successYn").item(0).text
					resultmsg		= SubNodes.getElementsByTagName("ns1:errorMsg").item(0).text

					If resultcode = "true" Then		'성공(200)
                        unitCd          = SubNodes.getElementsByTagName("ns1:unitCd").item(0).text
                        strDt           = SubNodes.getElementsByTagName("ns1:strDt").item(0).text
                        endDt           = SubNodes.getElementsByTagName("ns1:endDt").item(0).text
                        availSupQty     = SubNodes.getElementsByTagName("ns1:availSupQty").item(0).text

                        If (strDt = endDt) Then
                            availSupQty=0
                        End If

                        sqlStr = "UpDate R"&VbCRLF
    				    sqlStr = sqlStr & " SET outmalllimitno="&availSupQty&VbCRLF
    				    If availSupQty < 0 Then
    				    	sqlStr = sqlStr & " ,outmalllimityn='N'"
    				    Else
    						sqlStr = sqlStr & " ,outmalllimityn='Y'"
    					End If
    				    sqlStr = sqlStr & " from db_item.dbo.tbl_OutMall_regedoption R"&VbCRLF
    				    sqlStr = sqlStr & "  where mallid='"&CMALLNAME&"'"&VbCRLF
    				    sqlStr = sqlStr & "  and itemid="&iitemid&VbCRLF
    				    sqlStr = sqlStr & "  and outmallOptCode='"&unitCd&"'"&VbCRLF
    				    dbget.Execute sqlStr
					Else
						AssignedNotItemCnt = AssignedNotItemCnt + 1
						resultmsg = resultmsg & "_" & resultmsg
					End If
				End If
			Next
			If AssignedNotItemCnt > 0 Then
		        fnCJMallOptionQTYEdit = false
		        iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(단품수량)"
			Else
				iErrStr =  "OK||"&iitemid&"||수정성공(단품수량)"
				fnCJMallOptionQTYEdit = False
			End If
		Else
			iErrStr = "ERR||"&iitemid&"||CJMall과 통신중에 오류가 발생했습니다.[ERR-QTY-001]"
			fnCJMallOptionQTYEdit = false
		End If
		Set objXML = Nothing
		Set xmlDOM = Nothing
		On Error Goto 0
	End If
End Function

'단품 상태 수정
Function fnCJMallOptSellEdit(iitemid, byRef iErrStr, strParam, imaySoldout)
    Dim objXML, xmlDOM, resultcode, resultmsg
    Dim strRst, sqlStr
    Dim Nodes, SubNodes, failMsg
    Dim itemCd_zip, packInd, typeCd
    Dim sellynCnt, maySellYn
    Dim AssignedNotItemCnt : AssignedNotItemCnt = 0
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(strParam)

	' If (session("ssBctID")="kjy8517") Then
	' 	response.write replace(BinaryToText(objXML.ResponseBody, "euc-kr"), "<?xml", "<aaaaaaa")
	' 	response.end
	' End If


	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

			'response.write objXML.ResponseText
			'response.end

		Set Nodes = xmlDOM.getElementsByTagName("ns1:itemStates")
		If (Not (Nodes is Nothing)) Then
			For each SubNodes in Nodes
				If (Not (SubNodes is Nothing)) Then
					resultcode		= SubNodes.getElementsByTagName("ns1:successYn").item(0).text
					resultmsg		= SubNodes.getElementsByTagName("ns1:errorMsg").item(0).text
					If resultcode = "true" Then		'성공(200)
					    typeCd      = SubNodes.getElementsByTagName("ns1:typeCd").item(0).text
						itemCd_zip 	= SubNodes.getElementsByTagName("ns1:itemCd_zip").item(0).text
						packInd		= SubNodes.getElementsByTagName("ns1:packInd").item(0).text

						If typeCd = "02" Then
							sqlStr = ""
							sqlStr = sqlStr & " UPDATE [db_item].[dbo].tbl_OutMall_regedoption  " & VBCRLF
							sqlStr = sqlStr & " SET outmallSellyn = '"&chkiif(packInd="A","Y","N")&"'" & VBCRLF
							sqlStr = sqlStr & " , lastupdate = getdate() " & VBCRLF
							sqlStr = sqlStr & " WHERE itemid = '"&iitemid&"'  " & VBCRLF
							sqlStr = sqlStr & " and outmallOptCode = '"&itemCd_zip&"' " & VBCRLF
							sqlStr = sqlStr & " and mallid='"&CMALLNAME&"'"&VbCRLF
							dbget.Execute sqlStr
						End If
					Else
						AssignedNotItemCnt = AssignedNotItemCnt + 1
						failMsg = SubNodes.getElementsByTagName("ns1:errorMsg").item(0).text
					End If
				End If
			Next

			If AssignedNotItemCnt > 0 Then
		        fnCJMallOptSellEdit = false
		        iErrStr =  "ERR||"&iitemid&"||"&failMsg&"(단품상태)"
		    Else
		        iErrStr =  "OK||"&iitemid&"||수정성공(단품상태)"
		        fnCJMallOptSellEdit = true
		    End If
		Else
			iErrStr = "ERR||"&iitemid&"||CJMall과 통신중에 오류가 발생했습니다.[ERR-OPTSELL-001]"
			fnCJMallOptSellEdit = false
		End If
		Set objXML = Nothing
		Set xmlDOM = Nothing
	End If
	On Error Goto 0
End Function

'정보 수정
Function fnCJMallOneItemEdit(iitemid, byRef iErrStr, strParam)
    Dim objXML, xmlDOM, resultcode, resultmsg
    Dim strRst, sqlStr

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

'			response.write objXML.ResponseText
'			response.end
			resultcode	= xmlDOM.getElementsByTagName("ns1:successYn").item(0).text
			resultmsg	= xmlDOM.getElementsByTagName("ns1:errorMsg").item(0).text

			If resultcode = "true" Then		'성공(200)
				sqlStr = ""
				sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_cjmall_regItem "
				sqlStr = sqlStr & " SET regitemname = B.itemname "
				sqlStr = sqlStr & " FROM db_item.dbo.tbl_cjmall_regItem A "
				sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item B on A.itemid = B.itemid "
				sqlStr = sqlStr & " WHERE A.itemid='" & iitemid & "'"
				dbget.Execute(sqlStr)
				fnCJMallOneItemEdit = true
				iErrStr =  "OK||"&iitemid&"||수정성공(상품정보)"
			Else							'실패(E)
				'lastStatCheckDate수정하는 이유 : 옵션이 새로 추가가 되면 regedoption에 저장 하는 것이 아님, 상품 조회에서만 CJ측 단품코드를 얻을 수 있음.
				'그래서 스케줄링을 lastStatCheckDate ASC로 하기 때문에 아래 작업이 필요함.
				If (Trim(resultmsg)="1번째 단품:유효하지 않은 값입니다.[단품정보-협력사상품코드(Vpn)]가 이미 존재합니다. 새로운 값을 사용하세요.") then
					sqlStr = ""
					sqlStr = sqlStr & " UPDATE R"
					sqlStr = sqlStr & " SET lastStatCheckDate=NULL"                   '''등록실패
					sqlStr = sqlStr & " FROM db_item.dbo.tbl_cjmall_regitem as R"
					sqlStr = sqlStr & " WHERE R.itemid = "&iitemid
					dbget.Execute(sqlStr)
				End If
				fnCJMallOneItemEdit = false
				iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(상품정보)"
			End If
		Else
			iErrStr = "ERR||"&iitemid&"||CJMall과 통신중에 오류가 발생했습니다.[ERR-MOD-001]"
			fnCJMallOneItemEdit = false
		End If
		Set objXML = Nothing
		Set xmlDOM = Nothing
	On Error Goto 0
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################

'################################################# 각 기능 별 파라메터 정리 ###############################################
'상품 상태 변경 XML
Public Function getCJMallSellynParameter(icjmallPrdno, ichgSellYn)
	Dim stopYN, strRst

	If ichgSellYn = "N" Then
		stopYN = "I"
	ElseIf ichgSellYn = "Y" Then
		stopYN = "A"
	End If

	strRst = ""
	strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
	strRst = strRst &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_03' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_03.xsd'>"
	strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"					'!!!협력업체코드
	strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"	'!!!인증키
	strRst = strRst &"<tns:itemStates>"
	strRst = strRst &"	<tns:typeCd>01</tns:typeCd>"								'!!!01=판매코드,02=단품코드)
	strRst = strRst &"	<tns:itemCd_zip>"&icjmallPrdno&"</tns:itemCd_zip>"
	strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
	strRst = strRst &"	<tns:packInd>"&stopYN&"</tns:packInd>"						'!!!A-진행, I-일시중단
	strRst = strRst &"</tns:itemStates>"
	strRst = strRst &"</tns:ifRequest>"
	getCJMallSellynParameter = strRst
End Function

'상품 리스트 조회 XML
Public Function getCJMallStatChkParameter(iitemid)
	Dim cjMallPrdNo : cjMallPrdNo = getCjmallPrdNo(iitemid)
	Dim firstItemoption
	Dim strParam, strRst
	Dim strSql, cjmallRegdate
	strSql = ""
	strSql = strSql & " SELECT isnull(cjmallRegdate, getdate()) as cjmallRegdate "
	strSql = strSql & " FROM db_item.dbo.tbl_cjmall_regitem "
	strSql = strSql & " WHERE itemid = " & iitemid
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		cjmallRegdate = rsget("cjmallRegdate")
	End If
	rsget.Close

	If (cjMallPrdNo <> "") Then
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_07"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_07.xsd"">"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"
		strRst = strRst &"<tns:contents>"
		' strRst = strRst &"	<tns:sinstDtFrom>"&DateAdd("m",-2,Date)&"</tns:sinstDtFrom>"
		' strRst = strRst &"	<tns:sinstDtTo>"&Left(now, 10)&"</tns:sinstDtTo>"
		strRst = strRst &"	<tns:sinstDtFrom>"&LEFT(DateAdd("d",-7, cjmallRegdate ), 10)&"</tns:sinstDtFrom>"
		strRst = strRst &"	<tns:sinstDtTo>"&LEFT(DateAdd("d",7, cjmallRegdate ), 10)&"</tns:sinstDtTo>"
		strRst = strRst &"	<tns:schnCd>30</tns:schnCd>"
		strRst = strRst &"	<tns:itemCd>"&cjMallPrdNo&"</tns:itemCd>"
		strRst = strRst &"</tns:contents>"
		strRst = strRst &"</tns:ifRequest>"
	Else
		firstItemoption = getCjMallfirstItemoption(iitemid)
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_07"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_07.xsd"">"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"
		strRst = strRst &"<tns:contents>"
		' strRst = strRst &"	<tns:sinstDtFrom>"&DateAdd("m",-2,Date)&"</tns:sinstDtFrom>"
		' strRst = strRst &"	<tns:sinstDtTo>"&Left(now, 10)&"</tns:sinstDtTo>"
		strRst = strRst &"	<tns:sinstDtFrom>"&LEFT(DateAdd("d",-7, cjmallRegdate ), 10)&"</tns:sinstDtFrom>"
		strRst = strRst &"	<tns:sinstDtTo>"&LEFT(DateAdd("d",7, cjmallRegdate ), 10)&"</tns:sinstDtTo>"
		strRst = strRst &"	<tns:schnCd>30</tns:schnCd>"
		strRst = strRst &"	<tns:vpn>"&iitemid&"_"&firstItemoption&"</tns:vpn>"
		strRst = strRst &"</tns:contents>"
		strRst = strRst &"</tns:ifRequest>"
    End If
    getCJMallStatChkParameter = strRst
End Function
%>