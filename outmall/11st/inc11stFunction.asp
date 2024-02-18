<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'카테고리 정보 얻기
Public Function fn11stGetCate(iitemid, iErrStr)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, certType, requiredYn
	On Error Resume Next
'response.write strParam
'response.end
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & APIURL&"/cateservice/category/"&iitemid
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				certType = xmlDOM.getElementsByTagName("certType").item(0).text
				requiredYn = xmlDOM.getElementsByTagName("requiredYn").item(0).text

				strSql = ""
				strSql = strSql & " update db_etcmall.dbo.tbl_11st_category "
				strSql = strSql & " set safeDiv = '"&certType&"' "
				strSql = strSql & " ,isNeed = '"&requiredYn&"' "
				strSql = strSql & " where depthCode in ( "
				strSql = strSql & " '"&iitemid&"' "
				strSql = strSql & " ) "
				dbget.execute strSql
				iErrStr =  "OK||"&iitemid&"||성공(카테정보)"

			Set xmlDOM = nothing
		End If
	SET objXML = nothing
End Function

'상품 기본정보 등록
Public Function fn11stItemReg(iitemid, strParam, byRef iErrStr, imustprice, ist11sellyn, ilimityn, ilimitno, ilimiysold, iitemname, iimageNm)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode, prdNo
	On Error Resume Next
'response.write strParam
'response.end
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & APIURL&"/prodservices/product"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				retCode  = xmlDOM.getElementsByTagName("resultCode").item(0).text
				If retCode = "200" Then
					prdNo = xmlDOM.getElementsByTagName("productNo").item(0).text

					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET st11regdate = getdate()" & VbCrlf
					If (prdNo <> "") Then
					    strSql = strSql & "	, st11StatCd = '7'"& VbCRLF					'등록완료
					Else
						strSql = strSql & "	, st11StatCd = '1'"& VbCRLF					'전송시도
					End If
					strSql = strSql & " ,st11GoodNo = '" & prdNo & "'" & VbCrlf
					strSql = strSql & " ,st11lastupdate = getdate()"
					strSql = strSql & " ,st11Price = '"&imustprice&"' " & VbCrlf
					strSql = strSql & " ,st11sellYn = 'Y' "& VbCrlf
					strSql = strSql & " ,accFailCNT = 0" & VbCrlf                 ''실패회수 초기화
					strSql = strSql & " ,saleregdate = getdate()"
					strSql = strSql & " ,regimageName = '"&iimageNm&"'"& VbCrlf
					strSql = strSql & " FROM db_etcmall.dbo.tbl_11st_regitem R" & VbCrlf
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " where R.itemid = " & iitemid
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||등록성공(상품등록)"
				Else
					iMessage = replaceMsg(xmlDOM.getElementsByTagName("message").item(0).text)
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상품등록)"
				End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "ERR||"&iitemid&"||11st 결과 분석 중에 오류가 발생했습니다.[ERR-REG]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'상품 상태 변경
Public Function fn11stSellyn(iitemid, ichgSellyn, i11stGoodno, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, iMessage, goodsCd
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If ichgSellyn = "N" Then
			objXML.open "PUT", "" & APIURL&"/prodstatservice/stat/stopdisplay/"&i11stGoodno
		ElseIf ichgSellyn = "Y" Then
			objXML.open "PUT", "" & APIURL&"/prodstatservice/stat/restartdisplay/"&i11stGoodno
		End If
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody

				retCode  = xmlDOM.getElementsByTagName("resultCode").item(0).text
				If retCode = "200" Then
					If ichgSellyn = "Y" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	Set st11SellYn = 'Y'"
						strSql = strSql & "	,st11LastUpdate = getdate()"
						strSql = strSql & "	From db_etcmall.dbo.tbl_11st_regitem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||판매(상태변경)"
					ElseIf ichgSellyn = "N" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	Set st11SellYn = 'N'"
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,st11LastUpdate = getdate()"
						strSql = strSql & "	From db_etcmall.dbo.tbl_11st_regitem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||품절처리(상태변경)"
					End If
				Else
					iMessage = replaceMsg(xmlDOM.getElementsByTagName("message").item(0).text)
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상태변경)"
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||11st 결과 분석 중에 오류가 발생했습니다.[ERR-SOLDOUT]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'상품 가격 변경
Public Function fn11stPrice(iitemid, i11stGoodno, imustPrice, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, iMessage, goodsCd
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & APIURL&"/prodservices/product/price/"&i11stGoodno&"/"&imustPrice
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody

				retCode  = xmlDOM.getElementsByTagName("resultCode").item(0).text
				If retCode = "200" Then
				    strSql = ""
	    			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_11st_regitem  " & VbCRLF
	    			strSql = strSql & "	SET st11LastUpdate = getdate() " & VbCRLF
	    			strSql = strSql & "	, st11Price = " & imustprice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " Where itemid='" & iitemid & "'"& VbCRLF
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||수정성공(상품가격)"
					fn11stPrice = True
				Else
					iMessage = replaceMsg(xmlDOM.getElementsByTagName("message").item(0).text)
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상품가격)"
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||11st 결과 분석 중에 오류가 발생했습니다.[ERR-PRICE]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'재고 조회
Public Function fn11stStockChk(iitemid, i11stGoodno, iOptCnt, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, iMessage, prdNo, iRbody, Nodes, SubNodes
	Dim addPrc, mixDtlOptNm, mixOptNm, mixOptNo, optWght, vprdNo, prdStckNo, prdStckStatCd, selQty, sellerStockCd, stckQty
	Dim actCnt, AssignedRow
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & APIURL&"/prodmarketservice/prodmarket/stck/"&i11stGoodno
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				prdNo = xmlDOM.getElementsByTagName("ns2:prdNo").item(0).text
				If prdNo = i11stGoodno Then
					strSQL = ""
					strSQL = strSQL & " DELETE FROM db_item.dbo.tbl_Outmall_regedoption WHERE itemid = '"&iitemid&"' and mallid = '"&CMALLNAME&"' "
					dbget.Execute strSQL
					Set Nodes = xmlDOM.getElementsByTagName("ns2:ProductStock")
						For each SubNodes in Nodes
							addPrc			= SubNodes.getElementsByTagName("addPrc")(0).Text			'옵션가격
							mixDtlOptNm		= SubNodes.getElementsByTagName("mixDtlOptNm")(0).Text		'상세옵션명
							mixOptNm		= SubNodes.getElementsByTagName("mixOptNm")(0).Text			'옵션명
							mixOptNo		= SubNodes.getElementsByTagName("mixOptNo")(0).Text			'옵션번호
							optWght			= SubNodes.getElementsByTagName("optWght")(0).Text			'추가무게
							vprdNo			= SubNodes.getElementsByTagName("prdNo")(0).Text			'상품번호
							prdStckNo		= SubNodes.getElementsByTagName("prdStckNo")(0).Text		'재고번호 | 재고번호로 재고수량 변경가능합니다.(추가구성상품에 대한 조회/변경은 API에서 지원하지 않습니다.)
							prdStckStatCd	= SubNodes.getElementsByTagName("prdStckStatCd")(0).Text	'재고상태 | 01 : 사용, 02 : 품절
							selQty			= SubNodes.getElementsByTagName("selQty")(0).Text			'판매수량
							sellerStockCd	= SubNodes.getElementsByTagName("sellerStockCd")(0).Text	'셀러재고번호
							If Instr(sellerStockCd, iitemid&"_") > 0 Then
								sellerStockCd = replace(sellerStockCd, iitemid&"_", "")
							End If
'rw sellerStockCd
							stckQty			= SubNodes.getElementsByTagName("stckQty")(0).Text			'재고수량

							If iOptCnt = 0 Then
								sellerStockCd	= "0000"
								mixDtlOptNm		= "단일상품"
							End If

							' strSQL = ""
							' strSQL = strSQL & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
							' strSQL = strSQL & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastUpdate, checkdate) "
							' strSQL = strSQL & " VALUES "
							' strSQL = strSQL & " ('"&itemid&"'"
							' strSQL = strSQL & ",  '"&sellerStockCd&"'"
							' strSQL = strSQL & ", '"&CMALLNAME&"'"
							' strSQL = strSQL & ", '"&prdStckNo&"'"
							' strSQL = strSQL & ", '"&html2db(Trim(mixDtlOptNm))&"'"
							' strSQL = strSQL & ", '"&Chkiif(prdStckStatCd="01","Y","N")&"'"
							' strSQL = strSQL & ", '"&"Y"&"'"
							' strSQL = strSQL & ", '"&stckQty&"'"
							' strSQL = strSQL & ", getdate() "
							' strSQL = strSQL & ", getdate()) "

							'2019-05-02 14:20 김진영 하단 쿼리로 변경
							If sellerStockCd = "0000" Then
								strSQL = ""
								strSQL = strSQL & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
								strSQL = strSQL & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastUpdate, checkdate) "
								strSQL = strSQL & " VALUES "
								strSQL = strSQL & " ('"&itemid&"'"
								strSQL = strSQL & ",  '"&sellerStockCd&"'"
								strSQL = strSQL & ", '"&CMALLNAME&"'"
								strSQL = strSQL & ", '"&prdStckNo&"'"
								strSQL = strSQL & ", '"&html2db(Trim(mixDtlOptNm))&"'"
								strSQL = strSQL & ", '"&Chkiif(prdStckStatCd="01","Y","N")&"'"
								strSQL = strSQL & ", '"&"Y"&"'"
								strSQL = strSQL & ", '"&stckQty&"'"
								strSQL = strSQL & ", getdate() "
								strSQL = strSQL & ", getdate()) "
								dbget.Execute strSQL, AssignedRow
							Else
								strSQL = ""
								strSQL = strSQL & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
								strSQL = strSQL & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastUpdate, checkdate) "
								strSQL = strSQL & " SELECT itemid, itemoption, '"&CMALLNAME&"', '"&prdStckNo&"', optionname, '"&Chkiif(prdStckStatCd="01","Y","N")&"', 'Y', '"&stckQty&"', getdate(), getdate() "
								strSQL = strSQL & " FROM db_item.dbo.tbl_item_option "
								strSQL = strSQL & " WHERE itemid = '"& itemid &"' "
								strSQL = strSQL & " and itemoption = '"& sellerStockCd &"' "
								dbget.Execute strSQL, AssignedRow
							End If
							actCnt = actCnt + AssignedRow
						Next

						If (actCnt > 0) Then
							strSql = " update R"   &VbCRLF
							strSql = strSql & " set regedOptCnt=isNULL(T.regedOptCnt,0)"   &VbCRLF
							strSql = strSql & " from db_etcmall.dbo.tbl_11st_regItem R"   &VbCRLF
							strSql = strSql & " 	Join ("   &VbCRLF
							strSql = strSql & " 		select R.itemid,count(*) as CNT "
							strSql = strSql & " 		, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
							strSql = strSql & "        from db_etcmall.dbo.tbl_11st_regItem R"   &VbCRLF
							strSql = strSql & " 			Join db_item.dbo.tbl_OutMall_regedoption Ro"   &VbCRLF
							strSql = strSql & " 			on R.itemid=Ro.itemid"   &VbCRLF
							strSql = strSql & " 			and Ro.mallid='"&CMALLNAME&"'"   &VbCRLF
							strSql = strSql & "             and Ro.itemid="&iitemid&VbCRLF
							strSql = strSql & " 		group by R.itemid"   &VbCRLF
							strSql = strSql & " 	) T on R.itemid=T.itemid"   &VbCRLF
							dbget.Execute strSql
						End If
					Set Nodes = nothing
					iErrStr =  "OK||"&iitemid&"||성공(재고조회)"
				Else
					iMessage = replaceMsg(xmlDOM.getElementsByTagName("message").item(0).text)
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(재고조회)"
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||11st 결과 분석 중에 오류가 발생했습니다.[ERR-CHKSTOCK]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'10x10 상품코드로 11번가 상품 조회
Public Function fn11stStatChk(iitemid, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, iMessage, prdNo, iRbody, Nodes, SubNodes
	Dim addPrc, mixDtlOptNm, mixOptNm, mixOptNo, optWght, vprdNo, prdStckNo, prdStckStatCd, selQty, sellerStockCd, stckQty
	Dim actCnt, AssignedRow
'	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & APIURL&"/prodmarketservice/sellerprodcode/"&iitemid
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				If xmlDOM.getElementsByTagName("prdNo").length > 0 AND xmlDOM.getElementsByTagName("selPrc").length > 0 Then
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_11st_regitem "
					strSql = strSql & " SET st11Goodno = '"& xmlDOM.getElementsByTagName("prdNo").item(0).text &"' "
					strSql = strSql & " , st11Price = '"& xmlDOM.getElementsByTagName("selPrc").item(0).text &"' "
					strSql = strSql & " , st11StatCd = '7' "
					strSql = strSql & " WHERE itemid = '"& iitemid &"' "
					dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||성공(상품조회)"
				Else
					iMessage = "데이터없음"
					iErrStr = "OK||"&iitemid&"||"&iMessage&"(재고조회)"
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||11st 결과 분석 중에 오류가 발생했습니다.[ERR-CHKSTOCK]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'기본정보 수정
Public Function fn11stItemEdit(iitemid, strParam, byRef iErrStr, ichgImageNm, imustprice, i11stGoodno)
	Dim objXML, xmlDOM, strSql, goodsCd, iMessage, iRbody, retCode
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "PUT", "" & APIURL&"/prodservices/product/"&i11stGoodno
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				If (session("ssBctID")="kjy8517") Then
					rw "REQ : <textarea cols=40 rows=10>"&strParam&"</textarea>"
					rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
				End If
				retCode  = xmlDOM.getElementsByTagName("resultCode").item(0).text
				If (retCode = "200") OR (retCode = "210")  Then
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET st11GoodNo = '" & i11stGoodno & "'" & VbCrlf
					strSql = strSql & " ,st11lastupdate = getdate()"
					strSql = strSql & " ,saleregdate = getdate()"
'					strSql = strSql & " ,st11Price = '"&imustprice&"' " & VbCrlf
					strSql = strSql & " ,regitemname = '"&html2db(o11st.FOneItem.FItemname)&"' " & VbCRLF
					strSql = strSql & " ,regimageName = '"&ichgImageNm&"'"& VbCrlf
					strSql = strSql & " ,returnShippingFee = '3000'"& VbCrlf
					strSql = strSql & " FROM db_etcmall.dbo.tbl_11st_regitem R" & VbCrlf
					strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " where R.itemid = " & iitemid
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||성공(상품수정)"
				Else
					iMessage = replaceMsg(xmlDOM.getElementsByTagName("message").item(0).text)
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상품수정)"
				End If
			Set xmlDOM = nothing
		End If
	Set objXML = nothing
	iErrStr = replace(iErrStr, "'", "′")
	On Error Goto 0
End Function

'카테고리 코드
Public Function fn11stCommonCode(iccd, strParam)
	Dim objXML, xmlDOM, SubNodes, strSql
	Dim retCode, iMessage, Nodes
	Dim AssignedRow, iRbody, depth, dispNm, dispNo, parentDispNo
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & APIURL&"/cateservice/category"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.send()
		If iccd = "category" Then
			If objXML.Status = "200" Then
				Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
					xmlDOM.async = False
					iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
					xmlDOM.LoadXML iRbody
					Set Nodes = xmlDOM.getElementsByTagName("ns2:category")
						strSql = ""
						strSql = strSql & " DELETE FROM db_temp.[dbo].[tbl_11st_tmpCategory] "
						dbget.Execute(strSql)

						strSql = ""
						strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_11st_Category] "
						dbget.Execute(strSql)

						For each SubNodes in Nodes
							depth			= SubNodes.getElementsByTagName("depth")(0).Text		'트리 구조의 깊이. 11번가 전체 카테고리를 조회하기에 1 : 대카테고리. 2 : 중카테고리. 3 : 소카테고리. 4 : 세카테고리 로 보셔도 무방합니다. 단, 하위카테고리 조회는 예외입니다.
							dispNm			= SubNodes.getElementsByTagName("dispNm")(0).Text		'카테고리 이름
							dispNo			= SubNodes.getElementsByTagName("dispNo")(0).Text		'카테고리 번호
							parentDispNo	= SubNodes.getElementsByTagName("parentDispNo")(0).Text	'상위 카테고리 번호. 0 번은 트리 구조상 최상위로 대카테고리를 의미합니다.

							strSql = ""
							strSql = strSql & " INSERT INTO db_temp.[dbo].[tbl_11st_tmpCategory] (dispNo, depth, dispNm, parentdispNo) VALUES ('"&dispNo&"', '"&depth&"', '"&dispNm&"', '"&parentdispNo&"') "
							dbget.Execute(strSql)
						Next
						strSql = ""
						strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_11st_Category] (depthCode, depth1Code, Depth1Nm, depth2Code, Depth2Nm, depth3Code, Depth3Nm, depth4Code, Depth4Nm) "
						strSql = strSql & " SELECT "
						strSql = strSql & " (Case When isnull(T4.dispNo, '') <> '' THEN T4.dispNo Else T3.dispNo End) as depthCode "
						strSql = strSql & " ,T1.dispNo, T1.dispNm, T2.dispNo, T2.dispNm, T3.dispNo, T3.dispNm, isnull(T4.dispNo, ''), isnull(T4.dispNm, '') "
						strSql = strSql & " FROM db_temp.[dbo].[tbl_11st_tmpCategory] as T1 "
						strSql = strSql & " JOIN db_temp.[dbo].[tbl_11st_tmpCategory] as T2 on T1.dispno = T2.parentdispNo and t2.depth = '2' "
						strSql = strSql & " JOIN db_temp.[dbo].[tbl_11st_tmpCategory] as T3 on T2.dispno = T3.parentdispNo and t3.depth = '3' "
						strSql = strSql & " LEFT JOIN db_temp.[dbo].[tbl_11st_tmpCategory] as T4 on T3.dispno = T4.parentdispNo and t4.depth = '4' "
						strSql = strSql & " GROUP BY T1.dispNo, T1.dispNm, T2.dispNo, T2.dispNm, T3.dispNo, T3.dispNm, T4.dispNo, T4.dispNm "
						strSql = strSql & " ORDER BY T1.dispNo, T2.dispNo, T3.dispNo, T4.dispNo "
						dbget.Execute(strSql)
						iErrStr = "OK||카테고리||[category]성공 "
					Set Nodes = nothing
				Set xmlDOM = nothing
			End If
		End If
	Set objXML = nothing
End Function

Public Function fn11stoutinboundCode(iccd, strParam)
	Dim objXML, xmlDOM, SubNodes, strSql
	Dim retCode, iMessage, Nodes
	Dim AssignedRow, iRbody, depth, dispNm, dispNo, parentDispNo
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	If iccd= "outboundCode" Then
		objXML.open "GET", "" & APIURL&"/areaservice/outboundarea"
	Else
		objXML.open "GET", "" & APIURL&"/areaservice/inboundarea"
	End If
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				response.write iRbody

			Set xmlDOM = nothing
		End If
	Set objXML = nothing
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################
%>
