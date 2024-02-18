<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'등록 함수
Public Function fnInterparkItemReg(iitemid, strParam, idataUrl, byRef iErrStr, imustprice, ichgImageNm)
    Dim retCode, iMessage
    Dim objXML, xmlDOM, prdNo
    Dim strRst, strSql, iRbody, errorNodes, Nodes

	Dim fso,tFile
	Dim opath : opath = "/outmall/interpark/interparkXML/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	Dim fileName : fileName = "REG" &"_"& getCurrDateTimeFormat&"_"&iitemid&".xml"
	CALL CheckFolderCreate(defaultPath)
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & FileName )
			tFile.WriteLine idataUrl
		Set tFile = nothing
	Set fso = nothing
	If application("Svr_Info")="Dev" Then
		strParam = strParam & "&dataUrl=http://wapi.10x10.co.kr/outmall/interpark/interparkXML/2015-10-14/REG_2015-10-14_53700.29_889080.xml"
	Else
		strParam = strParam & "&dataUrl="&wapiURL&opath&FileName
	End If

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & interparkAPIURL&"/openapi/product/ProductAPIService.do", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				If (session("ssBctID")="kjy8517") Then
					rw "REQ : <textarea cols=40 rows=10>"&idataUrl&"</textarea>"
					rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
				End If
				Set errorNodes = xmlDOM.getElementsByTagName("error")
				If Not (errorNodes(0) is Nothing) Then
					retCode		= errorNodes(0).getElementsByTagName("code")(0).Text
					iMessage	= errorNodes(0).getElementsByTagName("explanation")(0).Text
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상품등록)"
				Else
					prdNo = xmlDOM.getElementsByTagName("prdNo").item(0).text
					If prdNo <> "" Then
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCrlf
						strSql = strSql & " SET interparkregdate = getdate()" & VbCrlf
						strSql = strSql & " ,interParkPrdNo = '" & prdNo & "'" & VbCrlf
						strSql = strSql & " ,interparklastupdate = getdate()"
						strSql = strSql & " ,mayiParkPrice = '"&imustprice&"' " & VbCrlf
						strSql = strSql & " ,mayiParkSellYn = 'Y' "& VbCrlf
						strSql = strSql & " ,R.saleregdate = getdate()"
						strSql = strSql & " ,accFailCNT = 0" & VbCrlf                 ''실패회수 초기화
						strSql = strSql & " ,regimageName = '"&ichgImageNm&"'"& VbCrlf
						strSql = strSql & " ,R.regitemname = i.itemname " & VbCRLF			''2020-11-23 김진영 추가
						strSql = strSql & " FROM [db_item].[dbo].tbl_interpark_reg_item R" & VbCrlf
						strSql = strSql & " JOIN  db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
						strSql = strSql & " where R.itemid = " & iitemid
						dbget.execute strSql

						' strSql = ""
						' strSql = strSql & " UPDATE R" & VbCrlf
						' strSql = strSql & " SET interparkSupplyCtrtSeq = 2"                   '''무조건 2번만 사용..
						' strSql = strSql & " ,interparkStoreCategory = D.interparkStoreCategory"
						' strSql = strSql & " ,Pinterparkdispcategory = D.interparkdispcategory"
						' strSql = strSql & " FROM [db_item].[dbo].tbl_interpark_reg_item R "
						' strSql = strSql & " JOIN [db_item].[dbo].tbl_item i on R.itemid=i.itemid"
						' strSql = strSql & " JOIN [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
						' strSql = strSql & " LEFT JOIN [db_item].[dbo].tbl_interpark_dspcategory_mapping D on D.tencdl=i.cate_large and D.tencdm=i.cate_mid and D.tencdn=i.cate_small "
						' strSql = strSql & " WHERE D.SupplyCtrtSeq is Not NULL"
						' strSql = strSql & " and i.itemid = "& iitemid & VbCrlf
						' strSql = strSql & " and R.interParkPrdNo is Not NULL"
						' dbget.execute strSql
						iErrStr =  "OK||"&iitemid&"||등록성공(상품등록)"
					Else
						iErrStr = "ERR||"&iitemid&"||인터파크 결과 분석 중에 오류가 발생했습니다.[ERR-REG-001]"
					End If
				End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "ERR||"&iitemid&"||인터파크 결과 분석 중에 오류가 발생했습니다.[ERR-REG-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
	Call DelAPITMPFile(wapiURL&opath&FileName)
End Function

'수정 함수
Public Function fnInterparkInfoEdit(iitemid, strParam, idataUrl, byRef iErrStr, ichgImageNm, imustprice, getiparkTp)
    Dim retCode, iMessage
    Dim objXML, xmlDOM, prdNo, editstat
    Dim strRst, strSql, iRbody, errorNodes, Nodes

	Dim fso,tFile
	Dim opath : opath = "/outmall/interpark/interparkXML/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	Dim fileName : fileName = "EDIT" &"_"& getCurrDateTimeFormat&"_"&iitemid&".xml"
	CALL CheckFolderCreate(defaultPath)
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & FileName )
			tFile.WriteLine idataUrl
		Set tFile = nothing
	Set fso = nothing

	If session("ssBctID")="kjy8517" Then
		response.write "<textarea cols=100 rows=30>"&idataUrl&"</textarea>"
	End If

	If application("Svr_Info")="Dev" Then
		strParam = strParam & "&dataUrl=http://wapi.10x10.co.kr/outmall/interpark/interparkXML/2015-10-06/EDIT_2015-10-12_52718.34_867805.xml"
	Else
		strParam = strParam & "&dataUrl="&wapiURL&opath&FileName
	End If
	On Error Resume Next

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & interparkAPIURL&"/openapi/product/ProductAPIService.do", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody

				Set errorNodes = xmlDOM.getElementsByTagName("error")
				If Not (errorNodes(0) is Nothing) Then
					retCode		= errorNodes(0).getElementsByTagName("code")(0).Text
					iMessage	= errorNodes(0).getElementsByTagName("explanation")(0).Text
					If Len(iMessage, "상세 이미지 업로드에 실패하였습니다") = "상세 이미지 업로드에 실패하였습니다" Then
						iMessage = "상세 이미지 업로드에 실패하였습니다.. 이미지 확인 부탁드립니다. detailImg"
					End If
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상품수정)"
				Else
					'response.write oInterpark.FOneItem.GetInterParkSaleStatTp		판매중:01, 품절:02, 판매중지:03, 일시품절:05, 예약판매:09, 상품삭제:98
					Select Case getiparkTp
						Case "01"		editstat = "Y"
						Case "02"		editstat = "N"
						Case "03"		editstat = "N"
						Case "05"		editstat = "N"
					End Select

					strSql = ""
					strSql = strSql & " UPDATE R " & VbCrlf
					strSql = strSql & " SET interparklastupdate = getdate()" & VbCrlf
					strSql = strSql & " ,mayiParkPrice='"&imustprice&"'"
					strSql = strSql & " ,accFailCNT=0" & VbCrlf
					strSql = strSql & " ,mayiParkSellYn = '" & editstat & "'" & VbCRLF
					strSql = strSql & " ,R.saleregdate = getdate()"
					strSql = strSql & " ,R.regitemname = i.itemname " & VbCRLF			''2020-11-23 김진영 추가
					If (ichgImageNm <> "N") Then
						strSql = strSql & " ,regimageName='"&ichgImageNm&"'"& VbCrlf
					End If
					strSql = strSql & " from [db_item].[dbo].tbl_interpark_reg_item R" & VbCrlf
					strSql = strSql & " Join db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
					strSql = strSql & " where R.itemid = " & iitemid
					dbget.execute strSql

					'''카테고리 매핑정보 때려 넣음. :: 카테고리가 바뀌어도 변경되지 않게.. // 매장은 바뀌믄안됨.
					' strSql = ""
					' strSql = strSql & " UPDATE R " & VbCrlf
					' strSql = strSql & " set interparkStoreCategory=D.interparkStoreCategory"
					' strSql = strSql & " , Pinterparkdispcategory=D.interparkdispcategory"
					' strSql = strSql & " from [db_item].[dbo].tbl_interpark_reg_item R"
					' strSql = strSql & " 	Join [db_item].[dbo].tbl_item i"
					' strSql = strSql & " 	on R.itemid=i.itemid"
					' strSql = strSql & " 	left join [db_item].[dbo].tbl_interpark_dspcategory_mapping D"
					' strSql = strSql & " 	on D.tencdl=i.cate_large"
					' strSql = strSql & " 	and D.tencdm=i.cate_mid"
					' strSql = strSql & " 	and D.tencdn=i.cate_small"
					' strSql = strSql & " where D.SupplyCtrtSeq is Not NULL"
					' strSql = strSql & " and i.itemid="& iitemid & VbCrlf
					' strSql = strSql & " and R.interParkPrdNo is Not NULL"
					'rw strSql
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||성공(상품수정)"
				End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "ERR||"&iitemid&"||인터파크 결과 분석 중에 오류가 발생했습니다.[ERR-EDIT-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
	iErrStr = replace(iErrStr, "'", "")
	Call DelAPITMPFile(wapiURL&opath&FileName)
End Function

'품절 수행 함수
Public Function fnInterparkSellyn(iitemid, ichgSellYn, strParam, idataUrl, byRef iErrStr)
    Dim retCode, iMessage
    Dim objXML, xmlDOM, prdNo
    Dim strRst, strSql, iRbody, errorNodes, Nodes

	Dim fso,tFile
	Dim opath : opath = "/outmall/interpark/interparkXML/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	Dim fileName : fileName = "EditSellYn" &"_"& getCurrDateTimeFormat&"_"&iitemid&".xml"
	CALL CheckFolderCreate(defaultPath)
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & fileName )
			tFile.WriteLine idataUrl
		Set tFile = nothing
	Set fso = nothing
	If application("Svr_Info")="Dev" Then
		strParam = strParam & "&dataUrl=http://wapi.10x10.co.kr/outmall/interpark/interparkXML/2015-09-24/EditSellYn_2015-09-24_65702.19_867805.xml"	'판매중
		'strParam = strParam & "&dataUrl=http://wapi.10x10.co.kr/outmall/interpark/interparkXML/2015-09-24/EditSellYn_2015-09-24_65177.89_867805.xml"	'품절
	Else
		strParam = strParam & "&dataUrl="&wapiURL&opath&FileName
	End If
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & interparkAPIURL&"/openapi/product/ProductAPIService.do", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				Set errorNodes = xmlDOM.getElementsByTagName("error")
				If Not (errorNodes(0) is Nothing) Then
					retCode		= errorNodes(0).getElementsByTagName("code")(0).Text
					iMessage	= errorNodes(0).getElementsByTagName("explanation")(0).Text
					iErrStr = "ERR||"&iitemid&"||"&iMessage
				Else
					prdNo = xmlDOM.getElementsByTagName("prdNo").item(0).text
					If prdNo <> "" Then
						If ichgSellYn = "X" Then
							strSql = ""
							strSql = strSql &" INSERT INTO [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] " & VBCRLF
							strSql = strSql &" SELECT TOP 1 'interpark', i.itemid, r.interParkPrdNo, r.interparkregdate, getdate(), r.lastErrStr" & VBCRLF
							strSql = strSql &" FROM db_item.dbo.tbl_item as i " & VBCRLF
							strSql = strSql &" JOIN db_item.dbo.tbl_interpark_reg_item as r on i.itemid = r.itemid " & VBCRLF
							strSql = strSql &" WHERE i.itemid = "&iitemid & VBCRLF
							dbget.Execute(strSql)

							strSql = ""
							strSql = strSql & " DELETE FROM db_item.dbo.tbl_interpark_reg_item " & vbcrlf
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
							strSql = strSql & " UPDATE [db_item].[dbo].tbl_interpark_reg_item " & VbCRLF
							strSql = strSql & " SET interparklastupdate = getdate() " & VbCRLF
							strSql = strSql & " ,mayiParkSellYn = '" & ichgSellYn & "'" & VbCRLF
							strSql = strSql & " ,accFailCnt = 0 " & VbCRLF
							strSql = strSql & " WHERE itemid='" & iitemid & "'"
							dbget.Execute(strSql)
						End If
						If ichgSellYn = "N" Then
							iErrStr = "OK||"&iitemid&"||품절처리"
						ElseIf ichgSellYn = "Y" Then
							iErrStr = "OK||"&iitemid&"||판매중으로 변경"
						ElseIf ichgSellYn = "X" Then
							iErrStr = "OK||"&iitemid&"||삭제"
						End If
					End If
				End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "ERR||"&iitemid&"||인터파크 결과 분석 중에 오류가 발생했습니다.[ERR-SELLEDIT-001]"
		End If
	Set objXML = Nothing
	on Error Goto 0
	Call DelAPITMPFile(wapiURL&opath&FileName)
	iErrStr = replace(iErrStr, "'", "")
End Function

'선택상품 판매상태 조회
Public Function fnInterparkstatChk(strParam, iitemid, iiparkprdno, iErrStr)
	Dim objXML, xmlDOM, Nodes, SubNodes, sqlStr, errorNodes
	Dim retCode, iMessage, optlist, iRbody, MasterPrice
	Dim prdNm,saleUnitcost,saleStatTp, optStkMgtYn, externalPrdNo, saleLmtQty, salePossRestQty
	Dim isOption, dispNo, prdOrOptNo, iparkRegitemname, iParkSellyn, orgOptcnt
	On Error Resume Next
	dbget.beginTrans

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & interparkAPIURL&"/openapi/product/ProductAPIService.do", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody

				If session("ssBctID")="kjy8517" Then
					response.write "<textarea cols=100 rows=30>"&iRbody&"</textarea>"
				End If

				Set errorNodes = xmlDOM.getElementsByTagName("error")
				If Not (errorNodes(0) is Nothing) Then
					retCode		= errorNodes(0).getElementsByTagName("code")(0).Text
					iMessage	= errorNodes(0).getElementsByTagName("explanation")(0).Text
					iErrStr =  "ERR||"&iitemid&"||"&iMessage&"(판매상태조회)"
				Else
					strSQL = ""
					strSQL = strSQL & " DELETE FROM db_item.dbo.tbl_Outmall_regedoption WHERE itemid = '"&iitemid&"' and mallid = '"&CMALLNAME&"' "
					dbget.Execute strSQL
					Set Nodes = xmlDOM.getElementsByTagName("item")
						For each SubNodes in Nodes
							MasterPrice = 0
							externalPrdNo	= SubNodes.getElementsByTagName("externalPrdNo")(0).Text	''TEN 상품번호 또는 옵션번호
							saleLmtQty		= SubNodes.getElementsByTagName("saleLmtQty")(0).Text       ''판매(예정)수량, 특판일 경우는 특판 한정수량
							salePossRestQty	= SubNodes.getElementsByTagName("salePossRestQty")(0).Text  ''남은수량, 특판일 경우는 특판 남은수량
							'optStkMgtYn		= SubNodes.getElementsByTagName("optStkMgtYn")(0).Text		''옵션재고수량 사용여부 - Y:사용함, N:사용안함
							saleUnitcost	= SubNodes.getElementsByTagName("saleUnitcost")(0).Text		''판매가, 옵션인경우 옵션추가금액을 합한금액
							prdOrOptNo		= SubNodes.getElementsByTagName("prdNo")(0).Text			''옵션인경우 인터파크 단품코드

							If (Trim(externalPrdNo) = Trim(iitemid)) Then
								MasterPrice			= saleUnitcost
								saleStatTp			= SubNodes.getElementsByTagName("saleStatTp")(0).Text       ''판매상태 - 판매중:01, 품절:02, 판매중지:03, 일시품절:05
								iparkRegitemname	= SubNodes.getElementsByTagName("prdNm")(0).Text
								iparkRegitemname	= Trim(replace(iparkRegitemname,"[텐바이텐]",""))
								iparkRegitemname	= replace(replace(replace(replace(iparkRegitemname,Chr(34),""),"<",""),">",""),"^","")
								isOption = False

								Select Case saleStatTp
									Case "01"				iParkSellyn = "Y"		''판매중
									Case "02"				iParkSellyn = "N"		''품절
									Case "05"				iParkSellyn = "S"		''일시품절
									Case "03", "10", "98"	iParkSellyn = "X"		''03 : 판매중지
								End Select

								strSQL = ""
								strSQL = strSQL & " UPDATE R" & VbCRLF
								strSQL = strSQL & " SET mayiparkPrice = " & MasterPrice & VbCRLF
								strSQL = strSQL & " ,mayiparkSellyn='"&iParkSellyn&"'" & VbCRLF
								''strSQL = strSQL & " ,regitemname='"&html2db(iparkRegitemname)&"'" & VbCRLF		''2020-11-23 김진영 주석처리
								''strSQL = strSQL & " ,lastStatCheckDate = getdate()" & VbCRLF
								strSQL = strSQL & " ,interparkprdno = isNULL(R.interparkprdno,'"&iiparkprdno&"')"&VbCRLF
								strSQL = strSQL & " From db_item.dbo.tbl_interpark_reg_Item R" & VbCRLF
								strSQL = strSQL & " where R.itemid="&iitemid & VbCRLF
								strSQL = strSQL & " and isNULL(interparkprdno,'') in ('','"&iiparkprdno&"')"&VbCRLF    ''중복등록된CaSE 대비
								strSQL = strSQL & " and (isNULL(mayiparkPrice,0)<>"&MasterPrice&"" & VbCRLF
								strSQL = strSQL & "     or isNULL(mayiparkSellyn,'')<>'"&iParkSellyn&"'"& VbCRLF
								strSQL = strSQL & "     or isNULL(regitemname,'')<>'"&html2db(iparkRegitemname)&"'"& VbCRLF
								strSQL = strSQL & "     or isNULL(interparkprdno,'')<>'"&iiparkprdno&"'"& VbCRLF
								strSQL = strSQL & " )"
								dbget.Execute strSQL
							Else
								isOption = True
							End If

							If (isOption) Then
								prdOrOptNo		= SubNodes.getElementsByTagName("prdNo")(0).Text			''옵션인경우 인터파크 단품코드
								prdNm			= SubNodes.getElementsByTagName("prdNm")(0).Text			''상품명 또는 옵션
								externalPrdNo	= SubNodes.getElementsByTagName("externalPrdNo")(0).Text	''TEN 상품번호 또는 옵션번호
								saleUnitcost	= SubNodes.getElementsByTagName("saleUnitcost")(0).Text		''판매가, 옵션인경우 옵션추가금액을 합한금액
								saleLmtQty		= SubNodes.getElementsByTagName("saleLmtQty")(0).Text       ''판매(예정)수량, 특판일 경우는 특판 한정수량
								If iitemid <> 824439 Then
									If InStr(prdNm,"|") > 0 Then
										prdNm = html2db(Trim(SplitValue(prdNm,"|",1)))
									End If
								End If

								If prdNm <> "" Then
									' strSQL = ""
									' strSQL = strSql & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_OutMall_regedoption WHERE itemid="&itemid&" and mallid = '"&CMALLNAME&"' and itemoption = '"&externalPrdNo&"' )"
									' strSQL = strSql & " BEGIN "
									' strSQL = strSQL & " 	INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
									' strSQL = strSQL & " 	(itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastUpdate, checkdate) "
									' strSQL = strSQL & " 	VALUES "
									' strSQL = strSQL & " 	('"&itemid&"'"
									' strSQL = strSQL & "		, '"&externalPrdNo&"'"
									' strSQL = strSQL & "		, '"&CMALLNAME&"'"
									' strSQL = strSQL & "		, '"&prdOrOptNo&"'"
									' strSQL = strSQL & "		, '"&html2db(Trim(SplitValue(prdNm,"/",1)))&"'"
									' strSQL = strSQL & "		, '"&Chkiif(saleStatTp="01","Y","N")&"'"
									' strSQL = strSQL & "		, '"&"Y"&"'"
									' strSQL = strSQL & "		, '"&saleLmtQty&"'"
									' strSQL = strSQL & "		, getdate() "
									' strSQL = strSQL & "		, getdate()) "
									' strSQL = strSql & " END "
									' strSQL = strSql & " ELSE "
									' strSQL = strSql & " BEGIN "
									' strSQL = strSQL & " 	UPDATE db_item.dbo.tbl_OutMall_regedoption SET "
									' strSQL = strSQL & "		outmalllimitno = '"&saleLmtQty&"' "
									' strSQL = strSQL & "		WHERE itemid="&itemid&" and mallid = '"&CMALLNAME&"' and itemoption = '"&externalPrdNo&"' "
									' strSQL = strSql & " END "
									'2019-04-30 15:51 김진영 하단 쿼리로 변경

'1. item_option테이블에서 검색해본다.
'2. 만약 위에서 있으면 기존대로 하면 된다
'3. 만약 위에서 있으면 다른 방법 모색
									orgOptcnt = 0
									strSQL = ""
									strSQL = strSQL & " SELECT COUNT(*) cnt "
									strSQL = strSQL & " FROM db_item.dbo.tbl_item_option "
									strSQL = strSQL & " WHERE itemid = '"& itemid &"' "
									strSQL = strSQL & " and itemoption = '"& externalPrdNo &"' "
									rsget.CursorLocation = adUseClient
									rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
										orgOptcnt = rsget("cnt")
									rsget.Close

 									strSQL = ""
									strSQL = strSQL & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_OutMall_regedoption WHERE itemid="&itemid&" and mallid = '"&CMALLNAME&"' and itemoption = '"&externalPrdNo&"' )"
									strSQL = strSQL & " BEGIN "
									If orgOptcnt > 0 Then
										strSQL = strSQL & " 	INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
										strSQL = strSQL & " 	(itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastUpdate, checkdate) "
										strSQL = strSQL & " 	SELECT itemid, itemoption, '"&CMALLNAME&"', '"&prdOrOptNo&"', optionname, '"&Chkiif(saleStatTp="01","Y","N")&"', 'Y', '"&saleLmtQty&"', getdate(), getdate() "
										strSQL = strSQL & " 	FROM db_item.dbo.tbl_item_option "
										strSQL = strSQL & " 	WHERE itemid = '"& itemid &"' "
										strSQL = strSQL & " 	and itemoption = '"& externalPrdNo &"' "
									Else
										strSQL = strSQL & " 	INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
										strSQL = strSQL & " 	(itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastUpdate, checkdate) "
										strSQL = strSQL & " 	VALUES "
										strSQL = strSQL & " 	('"&itemid&"'"
										strSQL = strSQL & "		, '"&externalPrdNo&"'"
										strSQL = strSQL & "		, '"&CMALLNAME&"'"
										strSQL = strSQL & "		, '"&prdOrOptNo&"'"
										strSQL = strSQL & "		, '"&html2db(Trim(SplitValue(prdNm,"/",1)))&"'"
										strSQL = strSQL & "		, '"&Chkiif(saleStatTp="01","Y","N")&"'"
										strSQL = strSQL & "		, 'Y'"
										strSQL = strSQL & "		, '"&saleLmtQty&"'"
										strSQL = strSQL & "		, getdate() "
										strSQL = strSQL & "		, getdate()) "
									End If
									strSQL = strSQL & " END "
									strSQL = strSQL & " ELSE "
									strSQL = strSQL & " BEGIN "
									strSQL = strSQL & " 	UPDATE db_item.dbo.tbl_OutMall_regedoption SET "
									strSQL = strSQL & "		outmalllimitno = '"&saleLmtQty&"' "
									strSQL = strSQL & "		WHERE itemid="&itemid&" and mallid = '"&CMALLNAME&"' and itemoption = '"&externalPrdNo&"' "
									strSQL = strSQL & " END "
									dbget.Execute strSQL
								End If
							End If
						Next
					Set Nodes = nothing
					strSQL = ""
					strSQL = strSQL & " UPDATE R"   &VbCRLF
					strSQL = strSQL & " SET regedOptCnt=isNULL(T.regedOptCnt,0)"   &VbCRLF
					strSQL = strSQL & " FROM db_item.dbo.tbl_interpark_reg_Item R"   &VbCRLF
					strSQL = strSQL & " JOIN ("   &VbCRLF
					strSQL = strSQL & " 	SELECT R.itemid,count(*) as CNT "
					strSQL = strSQL & " 	, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
					strSQL = strSQL & "     FROM db_item.dbo.tbl_interpark_reg_Item R"   &VbCRLF
					strSQL = strSQL & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = '"&CMALLNAME&"' and Ro.itemid = "&itemid &VbCRLF
					strSQL = strSQL & " 	GROUP BY R.itemid"   &VbCRLF
					strSQL = strSQL & " ) T on R.itemid=T.itemid"   &VbCRLF
					dbget.Execute strSQL


				    ''2017/12/21 by eastone  lastStatCheckDate 따로 뺌
				    strSQL = strSQL & " UPDATE R" & VbCRLF
				    strSQL = strSQL & " SET lastStatCheckDate = getdate()" & VbCRLF
				    strSQL = strSQL & " From db_item.dbo.tbl_interpark_reg_Item R" & VbCRLF
				    strSQL = strSQL & " where R.itemid="&iitemid & VbCRLF
				    strSQL = strSQL & " and isNULL(interparkprdno,'') in ('','"&iiparkprdno&"')"&VbCRLF    ''중복등록된CaSE 대비
				    dbget.Execute strSQL
					iErrStr =  "OK||"&iitemid&"||성공(판매상태조회)"
				End If
				Set errorNodes = nothing
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||인터파크 결과 분석 중에 오류가 발생했습니다.[ERR-STAT-001]"
		End If
	Set objXML = nothing

'rw Err.Number
'rw Err.Description
'rw Err.Source

	If Err.Number = 0 Then
		dbget.CommitTrans
	Else
		iErrStr = "ERR||"&iitemid&"||인터파크 결과 분석 중에 오류가 발생했습니다.[ERR-STAT-002]"
		dbget.RollBackTrans
	End If
	On Error Goto 0
	iErrStr = replace(iErrStr, "'", "")
End Function

'카테고리 땡겨오기
Public Function fnInterparkCategory(strParam)
	Dim objXML, xmlDOM, SubNodes, sqlStr
	Dim retCode, iMessage, optlist, buf, AssignedRow
	Dim Nodes, dispNo, dispNm, dispYn, regDts, modDts
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & interparkAPIURL&"/openapi/product/ProductAPIService.do", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)

		If objXML.Status = "200" Then
		    buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf
				Set Nodes = xmlDOM.getElementsByTagName("item")
					For each SubNodes in Nodes
						dispNo = SubNodes.getElementsByTagName("dispNo")(0).Text
						dispNm = SubNodes.getElementsByTagName("dispNm")(0).Text
						dispYn = SubNodes.getElementsByTagName("dispYn")(0).Text
						regDts = SubNodes.getElementsByTagName("regDts")(0).Text
						modDts = SubNodes.getElementsByTagName("modDts")(0).Text

						sqlStr = "update db_temp.dbo.tbl_interpark_Tmp_DispCategory"
						sqlStr = sqlStr & " set DispCateName=convert(varchar(255),'"&html2db(dispNm)&"')"
						sqlStr = sqlStr & " ,dispYn='"&dispYn&"'"
						sqlStr = sqlStr & " ,iParkregDts='"&regDts&"'"
						sqlStr = sqlStr & " ,iParkmodDts='"&modDts&"'"
						sqlStr = sqlStr & " where DispcateCode='"&dispNo&"'"
						dbget.Execute sqlStr, AssignedRow
						If (AssignedRow<1) and (dispYn<>"N") then  ''사용중인거만 입력
							sqlStr = "Insert Into db_temp.dbo.tbl_interpark_Tmp_DispCategory"
							sqlStr = sqlStr & " (DispcateCode,DispCateName,dispYn,lastRegdate,iParkregDts,iParkmodDts)"
							sqlStr = sqlStr & " values('"&dispNo&"'"
							sqlStr = sqlStr & " ,convert(varchar(255),'"&html2db(dispNm)&"')"
							sqlStr = sqlStr & " ,'"&dispYn&"'"
							sqlStr = sqlStr & " ,getdate()"
							sqlStr = sqlStr & " ,'"&regDts&"'"
							sqlStr = sqlStr & " ,'"&modDts&"'"
							sqlStr = sqlStr & " )"
							dbget.Execute sqlStr, AssignedRow
						End If
					Next

					iErrStr =  "OK||1111||성공(카테고리)"
				Set Nodes = nothing

			Set xmlDOM = nothing
		End If
	Set objXML = nothing
End Function

'카테고리 조회 개편
Public Function fnInterparkCategoryView(strParam)
	Dim objXML, xmlDOM, SubNodes, sqlStr
	Dim retCode, iMessage, optlist, buf, AssignedRow
	Dim Nodes, shopNo, dispNo, dispNm, dispYn, infoGroupNm, infoGroupNo, industrial, electric, child, medical, health, food, regDts, modDts

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & interparkAPIURL&"/openapi/product/ProductAPIService.do", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)

		If objXML.Status = "200" Then
		    buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf
'rw buf
				Set Nodes = xmlDOM.getElementsByTagName("item")
					For each SubNodes in Nodes
						shopNo		= SubNodes.getElementsByTagName("shopNo")(0).Text
						dispNo		= SubNodes.getElementsByTagName("dispNo")(0).Text
						dispNm		= SubNodes.getElementsByTagName("dispNm")(0).Text
						dispYn		= SubNodes.getElementsByTagName("dispYn")(0).Text
						infoGroupNm	= SubNodes.getElementsByTagName("infoGroupNm")(0).Text
						infoGroupNo	= SubNodes.getElementsByTagName("infoGroupNo")(0).Text
						industrial	= SubNodes.getElementsByTagName("industrial")(0).Text
						electric	= SubNodes.getElementsByTagName("electric")(0).Text
						child		= SubNodes.getElementsByTagName("child")(0).Text
						medical		= SubNodes.getElementsByTagName("medical")(0).Text
						health		= SubNodes.getElementsByTagName("health")(0).Text
						food		= SubNodes.getElementsByTagName("food")(0).Text
						regDts		= SubNodes.getElementsByTagName("regDts")(0).Text
						modDts		= SubNodes.getElementsByTagName("modDts")(0).Text

						sqlStr = ""
						sqlStr = sqlStr & " IF EXISTS(SELECT * FROM db_etcmall.dbo.tbl_interpark_category WHERE dispNo='"&dispNo&"') "
						sqlStr = sqlStr & " BEGIN"& VbCRLF
						sqlStr = sqlStr & " 	UPDATE db_etcmall.dbo.tbl_interpark_category " & VbCRLF
						sqlStr = sqlStr & " 	SET dispNm = '"& dispNm &"' " & VbCRLF
						sqlStr = sqlStr & " 	,dispYn = '"& dispYn &"' " & VbCRLF
						sqlStr = sqlStr & " 	,infoGroupNm = '"& infoGroupNm &"' " & VbCRLF
						sqlStr = sqlStr & " 	,infoGroupNo = '"& infoGroupNo &"' " & VbCRLF
						sqlStr = sqlStr & " 	,industrial = '"& industrial &"' " & VbCRLF
						sqlStr = sqlStr & " 	,electric = '"& electric &"' " & VbCRLF
						sqlStr = sqlStr & " 	,child = '"& child &"' " & VbCRLF
						sqlStr = sqlStr & " 	,medical = '"& medical &"' " & VbCRLF
						sqlStr = sqlStr & " 	,health = '"& health &"' " & VbCRLF
						sqlStr = sqlStr & " 	,food = '"& food &"' " & VbCRLF
						sqlStr = sqlStr & " 	WHERE dispNo='"&dispNo&"' " & VbCRLF
						sqlStr = sqlStr & " END ELSE " & VbCRLF
						sqlStr = sqlStr & " BEGIN"& VbCRLF
						sqlStr = sqlStr & " 	INSERT INTO db_etcmall.dbo.tbl_interpark_category " & VBCRLF
						sqlStr = sqlStr & " 	(shopNo, dispNo, dispNm, dispYn, infoGroupNm, infoGroupNo, industrial, electric, child, medical, health, food, regDts, modDts, regdate)  " & VBCRLF
						sqlStr = sqlStr & "		VALUES ('"& shopNo &"', '"& dispNo &"', '"& dispNm &"', '"& dispYn &"', '"& infoGroupNm &"', '"& infoGroupNo &"', '"& industrial &"', '"& electric &"', '"& child &"', '"& medical &"', '"& health &"', '"& food &"', '"& regDts &"', '"& modDts &"',  getdate()) " & VbCRLF
						sqlStr = sqlStr & " END "
						dbget.Execute sqlStr
					Next
rw "---------------끝"
response.end
					iErrStr =  "OK||1111||성공(카테고리)"
				Set Nodes = nothing
			Set xmlDOM = nothing
		End If
	Set objXML = nothing
	On Error Goto 0
End Function


'배송비정책 조회
Public Function fnInterparkDeliveryView(strParam)
	Dim objXML, xmlDOM, SubNodes, sqlStr
	Dim retCode, iMessage, optlist, buf, AssignedRow
	Dim Nodes, delvCostPlcNo, defaultYn, distCostTp, distCost, distCostCd, localCostYn, maxbuyAmt
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & interparkAPIURL&"/openapi/enterprise/EntrDelvAPIService.do", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
		    buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf
				Set Nodes = xmlDOM.getElementsByTagName("delvCostPlc")
					For each SubNodes in Nodes
						delvCostPlcNo	= SubNodes.getElementsByTagName("delvCostPlcNo")(0).Text	'묶음배송비번호
						defaultYn		= SubNodes.getElementsByTagName("defaultYn")(0).Text		'기본배송비여부 (Y : 예, N : 아니오)
						distCostTp		= SubNodes.getElementsByTagName("distCostTp")(0).Text		'배송비종류(00:무료, 98:판매자 조건부 무료, 99:판매자 정액)
						distCost		= SubNodes.getElementsByTagName("distCost")(0).Text			'배송비(배송비 종류가 무료일 경우 0)
						distCostCd		= SubNodes.getElementsByTagName("distCostCd")(0).Text		'결제방법(01:착불, 02:선불, 03:선/착불)
						localCostYn		= SubNodes.getElementsByTagName("localCostYn")(0).Text		'일부지역유료배송여부(무료배송의 경우 유효. Y:예, N:아니오)
						maxbuyAmt		= SubNodes.getElementsByTagName("maxbuyAmt")(0).Text		'무료배송 최소금액(배송비 종류가 조건부 무료일 때만)

						rw "묶음배송비코드 : " & delvCostPlcNo
						rw "기본배송비여부 : " &defaultYn
						rw "배송비종류 : " &distCostTp
						rw "배송비 : " &distCost
						rw "결제방법 : " &distCostCd
						rw "일부지역유료배송여부 : " &localCostYn
						rw "무료배송 최소금액 : " &maxbuyAmt
						rw "---------------------------------------"
					Next
					response.end
				Set Nodes = nothing
			Set xmlDOM = nothing
		End If
	Set objXML = nothing
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################

'################################################# 각 기능 별 파라메터 정리시작 ###############################################
'품절 파라메타
Function getInterparkSellynParameter(ichgSellYn, iParkPrdNo)
    Dim strRst, stopYN
	If ichgSellYn = "N" Then
		stopYN = "02"
	ElseIf ichgSellYn = "Y" Then
		stopYN = "01"
	ElseIf ichgSellYn = "X" Then
		'stopYN = "03"		'판매중지
		stopYN = "98"		'상품삭제
	End If

    strRst = ""
    strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr"" ?>"
    strRst = strRst & "<result>"
    strRst = strRst & "	<title>Interpark Product API</title>"
    strRst = strRst & "	<description>상품 재고 수정</description>"
    strRst = strRst & "	<item>"
    strRst = strRst & "		<prdNo>"&iParkPrdNo&"</prdNo>"
    strRst = strRst & "		<saleStatTp>"&stopYN&"</saleStatTp>"
    strRst = strRst & "</item>"
    strRst = strRst & "</result>"
	getInterparkSellynParameter = strRst
End Function

Function getCurrDateTimeFormat()
	Dim nowtimer : nowtimer= timer()
	getCurrDateTimeFormat = left(now(),10)&"_"&nowtimer
End Function

Sub CheckFolderCreate(sFolderPath)
	Dim objfile
	Set objfile = Server.CreateObject("Scripting.FileSystemObject")
	If NOT objfile.FolderExists(sFolderPath) Then
		objfile.CreateFolder sFolderPath
	End If
	Set objfile = Nothing
End Sub

''xml 파일 삭제
Function DelAPITMPFile(iFileURI)
	Dim iFullPath
	iFullPath = server.mappath(replace(iFileURI,"http://wapi.10x10.co.kr",""))

	Dim FSO, iFile
	Set FSO = CreateObject("Scripting.FileSystemObject")
		Set iFile = FSO.GetFile(iFullPath)
			If (iFile <> "") Then iFile.Delete
		Set iFile = Nothing
	Set FSO = Nothing
End Function
'################################################# 각 기능 별 파라메터 정리 끝 ###############################################
%>