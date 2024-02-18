<%
Function getApiUrl(mallid)
	Select Case mallid
		Case "lotteon"
			If application("Svr_Info") = "Dev" Then
				getApiUrl = "https://dev-openapi.lotteon.com"
			Else
				getApiUrl = "https://openapi.lotteon.com"
			End If
	End Select
End Function

Function getApiKey(mallid)
	Select Case mallid
		Case "lotteon"
			If application("Svr_Info") = "Dev" Then
				getApiKey = "5d5b2cb498f3d20001665f4e5451c4d923ac4e2c95df619996f35476"
			Else
				getApiKey = "5d5b2cb498f3d20001665f4e18a41621005d4c1ba262804ec7a10732"
			End If
	End Select
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

Function GetCSQnA_sabangnet(selldate)
	Dim istrParam
	istrParam = ""
	istrParam = istrParam & "<?xml version=""1.0"" encoding=""euc-kr""?>"
	istrParam = istrParam & "<SABANG_CS_LIST>"
	istrParam = istrParam & "	<HEADER>"
	istrParam = istrParam & "		<SEND_COMPAYNY_ID>tenbyten</SEND_COMPAYNY_ID>"
	istrParam = istrParam & "		<SEND_AUTH_KEY>PTxNV3d9CXPXBNu60X72EbSNYTJd5955b</SEND_AUTH_KEY>"
	istrParam = istrParam & "		<SEND_DATE>"&Replace(Date(), "-", "")&"</SEND_DATE>"
	istrParam = istrParam & "		<SEND_GOODS_CD_RT>Y</SEND_GOODS_CD_RT>"
	istrParam = istrParam & "	</HEADER>"
	istrParam = istrParam & "	<DATA>"
	istrParam = istrParam & "		<CS_ST_DATE>"& Replace(selldate, "-", "") &"</CS_ST_DATE>"
	istrParam = istrParam & "		<CS_ED_DATE>"& Replace(selldate, "-", "") &"</CS_ED_DATE>"
	istrParam = istrParam & "		<CS_STATUS></CS_STATUS>"
	istrParam = istrParam & "	</DATA>"
	istrParam = istrParam & "</SABANG_CS_LIST>"

	Dim dataURL, objXML, xmlDOM, strRst, iRbody, iMessage, maySabangnetGoodno, tmpGoodNo
	Dim fso,tFile, tenOptcd
	Dim Toptionname, Tlimitno, Tlimitsold, Tlimityn, Titemsu
	Dim opath : opath = "/outmall/sabangnet/sabangnetXML/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	Dim fileName : fileName = "reqQna" &"_"& getCurrDateTimeFormat&".xml"
	CALL CheckFolderCreate(defaultPath)
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & FileName )
			tFile.WriteLine istrParam
		Set tFile = nothing
	Set fso = nothing

	dataURL = "?xml_url=http://wapi.10x10.co.kr"&opath&FileName

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "https://r.sabangnet.co.kr/RTL_API/xml_cs_info.html" & dataURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				response.write iRbody
				response.end
			Set xmlDOM = Nothing
		End If
	Set objXML= nothing
End Function

Function GetCSQnA_lotteCom(selldate)
	Dim sellsite : sellsite = "lotteCom"
	Dim xmlURL, xmlSelldate, strParam
	Dim objXML, xmlDOM, iRbody, authNo, strSql
	Dim masterCnt, objMasterListXML, objMasterOneXML, questionPrefix, AssignedRow, iInputCnt, Cval
	Dim SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM
	GetCSQnA_lotteCom = False
	MALL_ID = "롯데닷컴"
	xmlSelldate = Replace(selldate, "-", "")

	If application("Svr_Info")="Dev" Then
		authNo = "c7ed8e97a9f5657ecde6d29455094ff2bb046cf7c864f0790615bd3298fb049931d500ed1614db2c3e1e17fc528b3269449545325572051804110afe770820aa"
	Else
		authNo = GetLotteAuthNo()
	End If

	xmlURL = "https://openapi.lotte.com"
	strParam = "subscriptionId=" & authNo & "&strSearchStrtDtime=" & xmlSelldate & "&strSearchEndDtime=" & xmlSelldate
	'// =======================================================================
	'// 데이타 가져오기
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", xmlURL & "/openapi/searchQnAListOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				If xmlDOM.getElementsByTagName("Response/Result/GoodsQuestInfo").length < 1 then
					response.write "내역없음(0)<br />"

					GetCSQnA_lotteCom = True
					Set xmlDOM = Nothing
					Set objXML = Nothing
					Exit Function
				End If
				masterCnt = xmlDOM.getElementsByTagName("Response/Result/GoodsQuestInfo").length
				rw "건수(" & masterCnt & ") "
				iInputCnt = 0
				set objMasterListXML = xmlDOM.selectNodes("/Response/Result/GoodsQuestInfo")
					For Each objMasterOneXML In objMasterListXML
						If Trim(objMasterOneXML.getElementsByTagName("ReceiptType").item(0).text) = "C"	Then	'타입 | C :고객센터(핫라인), Q:상품Q
							questionPrefix = "고객센터(핫라인)"
						Else
							questionPrefix ="상품Q&A"
						End If
						SabanetNum = Trim(objMasterOneXML.getElementsByTagName("ReceiptNo").item(0).text)		'순번
						'rw Trim(objMasterOneXML.getElementsByTagName("Gubun").item(0).text)					'구분
						'rw Trim(objMasterOneXML.getElementsByTagName("AskNm").item(0).text)					'답변자
						MALL_USER_ID = Trim(objMasterOneXML.getElementsByTagName("QuestNm").item(0).text)		'문의자
						'rw Trim(objMasterOneXML.getElementsByTagName("CellHpNo").item(0).text)					'연락처
						Select Case Trim(objMasterOneXML.getElementsByTagName("Cval").item(0).text)				'문의유형 | 고객핫라인 일경우 12:일반콜, 21:1:1E메일, 22:일반E메일, 상품Q 일경우 1:사이즈, 2:구성품, 3:디자인/색상, 4:사용설명, 9기타
							Case "12"	Cval = "일반콜"
							Case "21"	Cval = "1:1E메일"
							Case "22"	Cval = "일반E메일"
							Case "1"	Cval = "사이즈"
							Case "2"	Cval = "구성품"
							Case "3"	Cval = "디자인/색상"
							Case "4"	Cval = "사용설명"
							Case "9"	Cval = "기타"
							Case Else	Cval = "기타"
						End Select

						CS_GUBUN = "DIRECT"		'c#으로 작업한 것외에 제휴API이용은 전부 DIRECT처리

						SUBJECT = "[" & questionPrefix & "_" & Cval & "] " & Trim(objMasterOneXML.getElementsByTagName("Subject").item(0).text)	'제목
						CNTS = Trim(objMasterOneXML.getElementsByTagName("Content").item(0).text)				'문의내용
						'rw Trim(objMasterOneXML.getElementsByTagName("ReplyTitle").item(0).text)				'답변제목
						REG_DM = Trim(objMasterOneXML.getElementsByTagName("ReceiptDate").item(0).text)			'등록일
						REG_DM = Left(REG_DM,4) & "-" & Mid(REG_DM,5,2) & "-" & Mid(REG_DM,7,2) & " " & Mid(REG_DM,9,2) & ":" & Mid(REG_DM,11,2) & ":" & Mid(REG_DM,13,2)

						If Trim(objMasterOneXML.getElementsByTagName("Result").item(0).text) = "02" Then		'처리여부 | 02: 처리완료, 02를 제외한 값은 미처리
							CS_STATUS = "003"
						Else
							CS_STATUS = "001"
						End If

						'rw Trim(objMasterOneXML.getElementsByTagName("ResultMsg").item(0).text)				'처리내용
						PRODUCT_ID = Trim(objMasterOneXML.getElementsByTagName("GoodsNo").item(0).text)			'상품번호
						'rw Trim(objMasterOneXML.getElementsByTagName("MsgType").item(0).text)					'문의구분(문의유형과 다름)
						ProductNM = Trim(objMasterOneXML.getElementsByTagName("GoodsNm").item(0).text)			'상품명
						'rw Trim(objMasterOneXML.getElementsByTagName("ResultDate").item(0).text)				'처리일

						' rw "SabanetNum : " & SabanetNum
						' rw "MALL_ID : " & MALL_ID
						' rw "MALL_USER_ID : " & MALL_USER_ID
						' rw "CS_STATUS : " & CS_STATUS
						' rw "REG_DM : " & REG_DM
						' rw "PRODUCT_ID : " & PRODUCT_ID
						' rw "SUBJECT : " & SUBJECT
						' rw "CNTS : " & CNTS
						' rw "INS_NM : " & INS_NM
						' rw "INS_DM : " & INS_DM
						' rw "RPLY_CNTS : " & RPLY_CNTS
						' rw "UPD_NM : " & UPD_NM
						' rw "UPD_DM : " & UPD_DM
						' rw "CS_GUBUN : " & CS_GUBUN
						' rw "COMPAYNY_GOODS_CD : " & COMPAYNY_GOODS_CD
						' rw "OrderID : " & OrderID
						' rw "SEND_DM : " & SEND_DM
						' rw "ProductNM : " & ProductNM
						' rw "-----------------------------------------------"

						strSql = ""
						strSql = strSql & " IF NOT Exists(SELECT * FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"' )"
						strSql = strSql & " 	BEGIN "
						strSql = strSql & " 		INSERT INTO db_temp.dbo.tbl_Sabannet_Detail "
						strSql = strSql & " 		(SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM, MALL_PROD_ID) VALUES "
						strSql = strSql & " 		('"&SabanetNum&"', '"&MALL_ID&"', '"&html2db(MALL_USER_ID)&"', '"&CS_STATUS&"', '"&REG_DM&"', '"&PRODUCT_ID&"', '"&html2db(SUBJECT)&"', '"&html2db(CNTS)&"', '"&INS_NM&"', null, '"&RPLY_CNTS&"', '"&UPD_NM&"', null, '"&CS_GUBUN&"', '"&PRODUCT_ID&"', '"&OrderID&"', '"&SEND_DM&"', '"&html2db(ProductNM)&"', '"&PRODUCT_ID&"') "
						strSql = strSql & "		END "
					If CS_STATUS = "003" Then
						strSql = strSql & " ELSE"
						strSql = strSql & " 	BEGIN "
						strSql = strSql & " 		UPDATE db_temp.dbo.tbl_Sabannet_Detail "
						strSql = strSql & " 		SET CS_STATUS = '003', TenStatus = 'C' "
						strSql = strSql & " 		WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"'  "
						strSql = strSql & "		END "
					End If
						dbget.Execute strSql, AssignedRow

						If (AssignedRow > 0) Then
							iInputCnt = iInputCnt + AssignedRow
						End If
					Next
					rw "상품Q&A입력및수정건수 : " & iInputCnt
					GetCSQnA_lotteCom = True
				set objMasterListXML = nothing
			Set xmlDOM = nothing
		Else
			response.write "ERROR : 통신오류" & objXML.Status
			dbget.close : response.end
		End If
	Set objXML= nothing

	If Err Then
		response.write "ERROR : 통신오류 - " & Err.Description
	End If
End Function

Function resCSQnA_lotteCom(iNum, iRply)
	Dim xmlURL, strParam
	Dim objXML, xmlDOM, iRbody, authNo, strSql
	resCSQnA_lotteCom = False

	If application("Svr_Info")="Dev" Then
		authNo = "c7ed8e97a9f5657ecde6d29455094ff2bb046cf7c864f0790615bd3298fb049931d500ed1614db2c3e1e17fc528b3269449545325572051804110afe770820aa"
'		iNum = "AAAAAAAAA"
	Else
		authNo = GetLotteAuthNo()
	End If

	xmlURL = "https://openapi.lotte.com"
	strParam = "subscriptionId=" & authNo & "&strInqNo=" & iNum & "&selAnsContType=5&strInqAnsCont=" & Server.URLEncode(iRply) & "&strAnsDispYn=Y&selProcType=1&strMemo="
	'// =======================================================================
	'// 데이타 가져오기
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", xmlURL & "/openapi/updateQnaAnswerOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				If (xmlDOM.getElementsByTagName("Result").length > 0) then
					If xmlDOM.getElementsByTagName("Result").item(0).text = "1" Then
						strSql = ""
						strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
						strSql = strSql & " SET CS_STATUS = '003' "
						strSql = strSql & " ,TenStatus = 'C' "
						strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
						strSql = strSql & " and SellSite = 'lotteCom' "
						dbget.Execute strSql
						resCSQnA_lotteCom = True
						rw iNum & " 답변완료"
					End If
				Else
					If (xmlDOM.getElementsByTagName("Response/Errors/Error/Code").length > 0) then
						Select Case xmlDOM.getElementsByTagName("Response/Errors/Error/Code").item(0).text
							Case "4001"	'이미 고객센터에 전달하였습니다.
								strSql = ""
								strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
								strSql = strSql & " SET CS_STATUS = '003' "
								strSql = strSql & " ,TenStatus = 'C' "
								strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
								strSql = strSql & " and SellSite = 'lotteCom' "
								dbget.Execute strSql
								resCSQnA_lotteCom = True
								rw iNum & " 이미처리완료"
							Case Else
								response.write "ERROR : " & xmlDOM.getElementsByTagName("Response/Errors/Error/Message").item(0).text
						End Select
					End If
				End If
			Set xmlDOM = nothing
		Else
			response.write "ERROR : 통신오류" & objXML.Status
		End If
	Set objXML= nothing

	If Err Then
		response.write "ERROR : 통신오류 - " & Err.Description
	End If
End Function

Function GetCSQnA_lotteimall(selldate)
	Dim sellsite : sellsite = "lotteimall"
	Dim xmlURL, xmlSelldate, strParam
	Dim objXML, xmlDOM, iRbody, authNo, strSql
	Dim masterCnt, objMasterListXML, objMasterOneXML, questionPrefix, AssignedRow, iInputCnt, Cval
	Dim SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM
	GetCSQnA_lotteimall = False
	MALL_ID = "롯데아이몰"
	xmlSelldate = Replace(selldate, "-", "")

	If application("Svr_Info")="Dev" Then
		authNo = "4ef8af11a83ffe9129c2aeb3d799760d2c95aa3f7c29a4d683d82a0015f92d6b4e1f9da0f5fac137419c787f711dacf42dd065a35158a7b8657aeb6eb48e2cc3"
	Else
		authNo = GetLotteimallAuthNo()
	End If

	xmlURL = "https://openapi.lotteimall.com"
	strParam = "subscriptionId=" & authNo & "&req_start_dtime=" & xmlSelldate & "&req_end_dtime=" & xmlSelldate
	'// =======================================================================
	'// 데이타 가져오기
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", xmlURL & "/openapi/searchQnAListOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				If xmlDOM.getElementsByTagName("Response/Result/GoodsQuestInfo").length < 1 then
					response.write "내역없음(0)<br />"

					GetCSQnA_lotteimall = True
					Set xmlDOM = Nothing
					Set objXML = Nothing
					Exit Function
				End If

				masterCnt = xmlDOM.getElementsByTagName("Response/Result/GoodsQuestInfo").length
				rw "건수(" & masterCnt & ") "
				iInputCnt = 0

				set objMasterListXML = xmlDOM.selectNodes("/Response/Result/GoodsQuestInfo")
					For Each objMasterOneXML In objMasterListXML
						If Trim(objMasterOneXML.getElementsByTagName("ReceiptType").item(0).text) = "C"	Then	'타입 | C :고객센터(핫라인), Q:상품Q
							questionPrefix = "고객센터(핫라인)"
						Else
							questionPrefix ="상품Q&A"
						End If
						SabanetNum = Trim(objMasterOneXML.getElementsByTagName("ReceiptNo").item(0).text)		'순번
						'rw Trim(objMasterOneXML.getElementsByTagName("Gubun").item(0).text)					'구분
						'rw Trim(objMasterOneXML.getElementsByTagName("AskNm").item(0).text)					'답변자
						MALL_USER_ID = Trim(objMasterOneXML.getElementsByTagName("QuestNm").item(0).text)		'문의자
						'rw Trim(objMasterOneXML.getElementsByTagName("CellHpNo").item(0).text)					'연락처

						Select Case Trim(objMasterOneXML.getElementsByTagName("Cval").item(0).text)				'문의유형 | 고객핫라인 일경우 12:일반콜, 21:1:1E메일, 22:일반E메일, 상품Q 일경우 1:사이즈, 2:구성품, 3:디자인/색상, 4:사용설명, 9기타
							Case "12"	Cval = "일반콜"
							Case "21"	Cval = "1:1E메일"
							Case "22"	Cval = "일반E메일"
							Case "1"	Cval = "사이즈"
							Case "2"	Cval = "구성품"
							Case "3"	Cval = "디자인/색상"
							Case "4"	Cval = "사용설명"
							Case "5"	Cval = "기타"
							Case Else	Cval = "기타"
						End Select

						CS_GUBUN = "DIRECT"		'c#으로 작업한 것외에 제휴API이용은 전부 DIRECT처리

						SUBJECT = "[" & questionPrefix & "_" & Cval & "] " & Trim(objMasterOneXML.getElementsByTagName("Subject").item(0).text)	'제목
						CNTS = Trim(objMasterOneXML.getElementsByTagName("Content").item(0).text)				'문의내용
						'rw Trim(objMasterOneXML.getElementsByTagName("ReplyTitle").item(0).text)				'답변제목
						REG_DM = Trim(objMasterOneXML.getElementsByTagName("ReceiptDate").item(0).text)			'등록일
						REG_DM = Left(REG_DM,4) & "-" & Mid(REG_DM,5,2) & "-" & Mid(REG_DM,7,2) & " " & Mid(REG_DM,9,2) & ":" & Mid(REG_DM,11,2) & ":" & Mid(REG_DM,13,2)

						If Trim(objMasterOneXML.getElementsByTagName("Result").item(0).text) = "02" Then		'처리여부 | 02: 처리완료, 02를 제외한 값은 미처리
							CS_STATUS = "003"
						Else
							CS_STATUS = "001"
						End If

						'rw Trim(objMasterOneXML.getElementsByTagName("ResultMsg").item(0).text)				'처리내용
						PRODUCT_ID = Trim(objMasterOneXML.getElementsByTagName("GoodsNo").item(0).text)			'상품번호
						'rw Trim(objMasterOneXML.getElementsByTagName("MsgType").item(0).text)					'문의구분(문의유형과 다름)
						ProductNM = Trim(objMasterOneXML.getElementsByTagName("GoodsNm").item(0).text)			'상품명
						'rw Trim(objMasterOneXML.getElementsByTagName("ResultDate").item(0).text)				'처리일

						' rw "SabanetNum : " & SabanetNum
						' rw "MALL_ID : " & MALL_ID
						' rw "MALL_USER_ID : " & MALL_USER_ID
						' rw "CS_STATUS : " & CS_STATUS
						' rw "REG_DM : " & REG_DM
						' rw "PRODUCT_ID : " & PRODUCT_ID
						' rw "SUBJECT : " & SUBJECT
						' rw "CNTS : " & CNTS
						' rw "INS_NM : " & INS_NM
						' rw "INS_DM : " & INS_DM
						' rw "RPLY_CNTS : " & RPLY_CNTS
						' rw "UPD_NM : " & UPD_NM
						' rw "UPD_DM : " & UPD_DM
						' rw "CS_GUBUN : " & CS_GUBUN
						' rw "COMPAYNY_GOODS_CD : " & COMPAYNY_GOODS_CD
						' rw "OrderID : " & OrderID
						' rw "SEND_DM : " & SEND_DM
						' rw "ProductNM : " & ProductNM
						' rw "-----------------------------------------------"

						strSql = ""
						strSql = strSql & " IF NOT Exists(SELECT * FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"' )"
						strSql = strSql & " 	BEGIN "
						strSql = strSql & " 		INSERT INTO db_temp.dbo.tbl_Sabannet_Detail "
						strSql = strSql & " 		(SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM, MALL_PROD_ID) VALUES "
						strSql = strSql & " 		('"&SabanetNum&"', '"&MALL_ID&"', '"&html2db(MALL_USER_ID)&"', '"&CS_STATUS&"', '"&REG_DM&"', '"&PRODUCT_ID&"', '"&html2db(SUBJECT)&"', '"&html2db(CNTS)&"', '"&INS_NM&"', null, '"&RPLY_CNTS&"', '"&UPD_NM&"', null, '"&CS_GUBUN&"', '"&PRODUCT_ID&"', '"&OrderID&"', '"&SEND_DM&"', '"&html2db(ProductNM)&"', '"&PRODUCT_ID&"') "
						strSql = strSql & "		END "
					If CS_STATUS = "003" Then
						strSql = strSql & " ELSE"
						strSql = strSql & " 	BEGIN "
						strSql = strSql & " 		UPDATE db_temp.dbo.tbl_Sabannet_Detail "
						strSql = strSql & " 		SET CS_STATUS = '003', TenStatus = 'C' "
						strSql = strSql & " 		WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"'  "
						strSql = strSql & "		END "
					End If
						dbget.Execute strSql, AssignedRow

						If (AssignedRow > 0) Then
							iInputCnt = iInputCnt + AssignedRow
						End If
					Next
					rw "상품Q&A입력및수정건수 : " & iInputCnt
					GetCSQnA_lotteimall = True
				set objMasterListXML = nothing
			Set xmlDOM = nothing
		Else
			response.write "ERROR : 통신오류" & objXML.Status
			dbget.close : response.end
		End If
	Set objXML= nothing

	If Err Then
		response.write "ERROR : 통신오류 - " & Err.Description
	End If
End Function

Function resCSQnA_lotteimall(iNum, iRply)
	Dim xmlURL, strParam
	Dim objXML, xmlDOM, iRbody, authNo, strSql
	resCSQnA_lotteimall = False

	If application("Svr_Info")="Dev" Then
		authNo = "4ef8af11a83ffe9129c2aeb3d799760d2c95aa3f7c29a4d683d82a0015f92d6b4e1f9da0f5fac137419c787f711dacf42dd065a35158a7b8657aeb6eb48e2cc3"
'		iNum = "AAAAAAAAA"
	Else
		authNo = GetLotteimallAuthNo()
	End If
	iRply = replace(replace(replace(iRply,"&","%26"), "+", "%2B"), "%", "프로")

	xmlURL = "https://openapi.lotteimall.com"
	strParam = "subscriptionId=" & authNo & "&inq_no=" & iNum & "&ans_cont_type=5&inq_ans_cont=" & iRply & "&ans_disp_yn=Y&proc_type=1&memo="
	'// =======================================================================
	'// 데이타 가져오기
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", xmlURL & "/openapi/updateQnaAnswerOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				If (xmlDOM.getElementsByTagName("Result").length > 0) then
					If xmlDOM.getElementsByTagName("Result").item(0).text = "1" Then
						strSql = ""
						strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
						strSql = strSql & " SET CS_STATUS = '003' "
						strSql = strSql & " ,TenStatus = 'C' "
						strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
						strSql = strSql & " and SellSite = 'lotteimall' "
						dbget.Execute strSql
						resCSQnA_lotteimall = True
						rw iNum & " 답변완료"
					End If
				Else
					If (xmlDOM.getElementsByTagName("Response/Errors/Error/Code").length > 0) then
						Select Case xmlDOM.getElementsByTagName("Response/Errors/Error/Code").item(0).text
							Case "4001"	'이미 고객센터에 전달하였습니다.
								strSql = ""
								strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
								strSql = strSql & " SET CS_STATUS = '003' "
								strSql = strSql & " ,TenStatus = 'C' "
								strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
								strSql = strSql & " and SellSite = 'lotteimall' "
								dbget.Execute strSql
								resCSQnA_lotteimall = True
								rw iNum & " 이미처리완료"
							Case Else
								response.write "ERROR : " & xmlDOM.getElementsByTagName("Response/Errors/Error/Message").item(0).text
						End Select
					End If
				End If
			Set xmlDOM = nothing
		Else
			response.write "ERROR : 통신오류" & objXML.Status
		End If
	Set objXML= nothing

	If Err Then
		response.write "ERROR : 통신오류 - " & Err.Description
	End If
End Function

Function GetCSQnA_lotteon(selldate)
	Dim sellsite : sellsite = "lotteon"
	Dim apiUrl, xmlSelldate, strParam, obj
	Dim objXML, xmlDOM, iRbody, apiKey, strSql, jParam, trCd, objData
	Dim masterCnt, objMasterListXML, objMasterOneXML, questionPrefix, AssignedRow, iInputCnt, Cval
	Dim dataList, strObj, returnCode
	Dim SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM
	GetCSQnA_lotteon = False
	MALL_ID = "롯데On(item)"
	xmlSelldate = Replace(selldate, "-", "")

	apiUrl = getApiUrl("lotteon")
	apiKey = getApiKey("lotteon")
	apiUrl = apiUrl & "/v1/openapi/product/v1/product/qna/list"

	If application("Svr_Info") = "Dev" Then
		trCd = "LO10001101"
	Else
		trCd = "LD304013"
	End If
	'// =======================================================================
	Set obj = jsObject()
		obj("trGrpCd") = "SR"			'#거래처그룹코드 | 일반셀러 : SR
		obj("trNo") = trCd				'#거래처번호
		obj("lrtrNo") = ""				'하위거래처번호
		obj("spdNo") = ""				'판매자상품번호
		obj("sitmNo") = ""				'판매자단품번호
		obj("regStrDttm") = xmlSelldate & "000000"	'#등록일시 조회시작일시 [YYYYMMDDHH24MISS ex) 20190801150010]
		obj("regEndDttm") = xmlSelldate & "235959"	'#등록일시 조회종료일시 [YYYYMMDDHH24MISS ex) 20190810150010]
		obj("qstTypCd") = null			'QnA 유형 [공통코드 : QST_TYP_CD] | NULL인 경우에는 전체유형을 조회한다. SZ_CAPA : 사이즈/용량, DSGN_CLR : 디자인/색상, DP_INFO : 상품정보, USE_EPN : 사용설명, ETC : 기타
		obj("qnaStatCd") = null			'QnA처리상태코드 [공통코드 : QNA_STAT_CD] NULL인 경우에는 전체 조회한다. NPROC : 미처리, PROC : 처리완료, CC_TCTL : 고객센터이관
		obj("pageNo") = 1				'#페이지
		obj("rowsPerPage") = 100		'#페이당건수 (MAX 100)
		jParam = obj.jsString
	Set obj = nothing

	'// 데이타 가져오기
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(jParam)

		If objXML.Status <> "200" Then
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
			dbget.close : response.end
		Else
			objData = BinaryToText(objXML.ResponseBody,"utf-8")
		End If

	'// =======================================================================
	'// Json 파싱
	Set strObj = JSON.parse(objData)
'rw objData
		returnCode		= strObj.returnCode
		If returnCode = "0000" Then
			iInputCnt = 0
			Set dataList = strObj.data
				If dataList.length > 0 Then
					For i=0 to dataList.length-1
						ProductNM = ""
						questionPrefix = "상품Q&A"
						SabanetNum = dataList.get(i).pdQnaNo		'#상품QnA번호
						MALL_USER_ID = ""
						CS_GUBUN = "DIRECT"							'c#으로 작업한 것외에 제휴API이용은 전부 DIRECT처리
						Select Case Trim(dataList.get(i).qstTypCd)	'QnA 유형 [공통코드 : QST_TYP_CD] | SZ_CAPA : 사이즈/용량, DSGN_CLR : 디자인/색상, DP_INFO : 상품정보, USE_EPN : 사용설명, ETC : 기타
							Case "SZ_CAPA"	Cval = "사이즈/용량"
							Case "DSGN_CLR"	Cval = "디자인/색상"
							Case "DP_INFO"	Cval = "상품정보"
							Case "USE_EPN"	Cval = "사용설명"
							Case "ETC"		Cval = "기타"
							Case Else		Cval = "기타"
						End Select
						SUBJECT = "[" & questionPrefix & "_" & Cval & "] " & LEFT(Trim(dataList.get(i).qstCnts), 20)	'제목
						CNTS = Trim(dataList.get(i).qstCnts)		'질문내용
						REG_DM = dataList.get(i).regDttm			'등록일시
						REG_DM = Left(REG_DM,4) & "-" & Mid(REG_DM,5,2) & "-" & Mid(REG_DM,7,2) & " " & Mid(REG_DM,9,2) & ":" & Mid(REG_DM,11,2) & ":" & Mid(REG_DM,13,2)
						If Trim(dataList.get(i).qnaStatCd) = "PROC" Then		''#QnA처리상태코드 [공통코드 : QNA_STAT_CD] NULL인 경우에는 전체 조회한다. | NPROC : 미처리, PROC : 처리완료, CC_TCTL : 고객센터이관
							CS_STATUS = "003"
						Else
							CS_STATUS = "001"
						End If

						PRODUCT_ID = Trim(dataList.get(i).spdNo)			'판매자상품번호
						COMPAYNY_GOODS_CD = Trim(dataList.get(i).sitmNo)	'판매자단품번호

						strSql = ""
						strSql = strSql & " SELECT TOP 1 regitemname "
						strSql = strSql & " FROM db_etcmall.dbo.tbl_lotteon_regItem "
						strSql = strSql & " WHERE lotteonGoodNo = '"&PRODUCT_ID&"' "
						rsget.CursorLocation = adUseClient
						rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
						If Not rsget.Eof Then
							ProductNM 		= rsget("regitemname")
						End If
						rsget.Close

						strSql = ""
						strSql = strSql & " IF NOT Exists(SELECT * FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"' )"
						strSql = strSql & " 	BEGIN "
						strSql = strSql & " 		INSERT INTO db_temp.dbo.tbl_Sabannet_Detail "
						strSql = strSql & " 		(SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM, MALL_PROD_ID) VALUES "
						strSql = strSql & " 		('"&SabanetNum&"', '"&MALL_ID&"', '"&html2db(MALL_USER_ID)&"', '"&CS_STATUS&"', '"&REG_DM&"', '"&PRODUCT_ID&"', '"&html2db(SUBJECT)&"', '"&html2db(CNTS)&"', '"&INS_NM&"', null, '"&RPLY_CNTS&"', '"&UPD_NM&"', null, '"&CS_GUBUN&"', '"&COMPAYNY_GOODS_CD&"', '"&OrderID&"', '"&SEND_DM&"', '"&html2db(ProductNM)&"', '"&PRODUCT_ID&"') "
						strSql = strSql & "		END "
						If CS_STATUS = "003" Then
							strSql = strSql & " ELSE"
							strSql = strSql & " 	BEGIN "
							strSql = strSql & " 		UPDATE db_temp.dbo.tbl_Sabannet_Detail "
							strSql = strSql & " 		SET CS_STATUS = '003', TenStatus = 'C' "
							strSql = strSql & " 		WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"'  "
							strSql = strSql & "		END "
						End If
						dbget.Execute strSql, AssignedRow

						If (AssignedRow > 0) Then
							iInputCnt = iInputCnt + AssignedRow
						End If
					Next
					rw "상품Q&A입력및수정건수 : " & iInputCnt
					GetCSQnA_lotteon = True
				Else
					response.write "내역없음(0)<br />"

					GetCSQnA_lotteon = True
					Set dataList = Nothing
					Set strObj = Nothing
					exit function
				End If
			Set dataList = nothing
		Else
			rw strObj.message
		End If
	Set strObj = nothing
End Function

Function GetCSSellerQnA_lotteon(selldate)
	Dim sellsite : sellsite = "lotteon"
	Dim apiUrl, xmlSelldate, strParam, obj
	Dim objXML, xmlDOM, iRbody, apiKey, strSql, jParam, trCd, objData
	Dim masterCnt, objMasterListXML, objMasterOneXML, questionPrefix, AssignedRow, iInputCnt, Cval
	Dim rsltList, strObj, returnCode
	Dim SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM
	GetCSSellerQnA_lotteon = False
	MALL_ID = "롯데On(seller)"
	xmlSelldate = Replace(selldate, "-", "")

	apiUrl = getApiUrl("lotteon")
	apiKey = getApiKey("lotteon")
	apiUrl = apiUrl & "/v1/openapi/customer/v1/getSellerInquiryList"
	'// =======================================================================
	Set obj = jsObject()
		obj("scStrtDt") = xmlSelldate			'#조회기간 시작일 [yyyymmdd : 20190801]
		obj("scEndDt") = xmlSelldate			'#조회기간 종료일 [yyyymmdd : 20190801]
'		obj("vocLcsfCd") = ""					'문의유형코드
'		obj("slrInqProcStatCd") = ""			'판매자문의처리상태코드 | 전체:공란, 답변:ANS, 미답변:UNANS
'		obj("scKwd") = ""						'검색어
'		obj("spdNo") = ""						'판매자 상품번호
'		obj("spdNm") = ""						'판매자 상품명
'		obj("sitmNo") = ""						'판매자 단품번호
'		obj("sitmNm") = ""						'판매자 단품명
'		obj("lrtrNo") = ""						'거래처번호
'		obj("pageNo") = ""						'1
'		obj("rowsPerPage") = ""					'50
		jParam = obj.jsString
	Set obj = nothing

	'// 데이타 가져오기
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(jParam)

		If objXML.Status <> "200" Then
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
			dbget.close : response.end
		Else
			objData = BinaryToText(objXML.ResponseBody,"utf-8")
		End If
	'// =======================================================================
	'// Json 파싱
	Set strObj = JSON.parse(objData)
		returnCode		= strObj.rsltCd
		If returnCode = "0000" Then
			Set rsltList = strObj.rsltList
				iInputCnt = 0
				If rsltList.length > 0 Then
					For i=0 to rsltList.length-1
'						rw "aa : " & rsltList.length
						questionPrefix = "판매자문의Q&A"
						SabanetNum		= rsltList.get(i).slrInqNo				'판매자문의번호
						MALL_USER_ID	= ""									'롯데On API로 알 수가 없다;
						CS_GUBUN		= "DIRECT"								'c#으로 작업한 것외에 제휴API이용은 전부 DIRECT처리
						Cval			= rsltList.get(i).vocTypNm				'문의유형명
						SUBJECT			= "[" & questionPrefix & "_" & Cval & "] " & LEFT(Trim(rsltList.get(i).inqTtl), 20)	'문의제목
						CNTS 			= Trim(rsltList.get(i).inqCnts)			'문의내용
						If SabanetNum = "2902954" Then
							CNTS 		= "사이즈 285는 없나요?"
						End If
						REG_DM			= Trim(rsltList.get(i).accpDttm)		'접수일시
						REG_DM = Left(REG_DM,4) & "-" & Mid(REG_DM,5,2) & "-" & Mid(REG_DM,7,2) & " " & Mid(REG_DM,9,2) & ":" & Mid(REG_DM,11,2) & ":" & Mid(REG_DM,13,2)
						If Trim(rsltList.get(i).slrInqProcStatCd) = "ANS" Then	'판매자문의처리상태코드(ANS, UNANS)
							CS_STATUS = "003"
						Else
							CS_STATUS = "001"
						End If
						PRODUCT_ID		= Trim(rsltList.get(i).pdNo)			'상품번호
						ProductNM		= Trim(rsltList.get(i).pdNm)			'상품명
						OrderID			= rsltList.get(i).odNo					'주문번호

						' rsltList.get(i).slrInqNo			'판매자 문의번호
						' rsltList.get(i).vocLcsfCd			'문의유형코드
						' rsltList.get(i).vocTypNm			'문의유형명
						' rsltList.get(i).slrInqProcStatCd	'판매자문의처리상태코드(ANS, UNANS)
						' rsltList.get(i).slrInqProcStatNm	'판매자문의처리상태명(답변, 미답변)
						' rsltList.get(i).inqTtl				'문의제목
						' rsltList.get(i).inqCnts				'문의내용
						' rsltList.get(i).odNo				'주문번호
						' rsltList.get(i).pdNo				'상품번호
						' rsltList.get(i).pdNm				'상품명
						' rsltList.get(i).spdNo				'판매자 상품번호
						' rsltList.get(i).spdNm				'판매자 상품명
						' rsltList.get(i).sitmNo				'판매자 단품번호
						' rsltList.get(i).sitmNm				'판매자 단품명
						' rsltList.get(i).trNo				'상위거래처번호
						' rsltList.get(i).trNm				'상위거래처명
						' rsltList.get(i).lrtrNo				'하위거래처번호
						' rsltList.get(i).lrtrNm				'하위거래처명
						' rsltList.get(i).slrNo				'판매자번호
						' rsltList.get(i).ansCnts				'답변내용
						' rsltList.get(i).ansRqPrd			'답변소요기간
						' rsltList.get(i).accpDttm			'접수일시
						' rsltList.get(i).procDttm			'처리일시

						strSql = ""
						strSql = strSql & " IF NOT Exists(SELECT * FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"' )"
						strSql = strSql & " 	BEGIN "
						strSql = strSql & " 		INSERT INTO db_temp.dbo.tbl_Sabannet_Detail "
						strSql = strSql & " 		(SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM, MALL_PROD_ID) VALUES "
						strSql = strSql & " 		('"&SabanetNum&"', '"&MALL_ID&"', '"&html2db(MALL_USER_ID)&"', '"&CS_STATUS&"', '"&REG_DM&"', '"&PRODUCT_ID&"', '"&html2db(SUBJECT)&"', '"&html2db(CNTS)&"', '"&INS_NM&"', null, '"&RPLY_CNTS&"', '"&UPD_NM&"', null, '"&CS_GUBUN&"', '"&PRODUCT_ID&"', '"&OrderID&"', '"&SEND_DM&"', '"&html2db(ProductNM)&"', '"&PRODUCT_ID&"') "
						strSql = strSql & "		END "
						If CS_STATUS = "003" Then
							strSql = strSql & " ELSE"
							strSql = strSql & " 	BEGIN "
							strSql = strSql & " 		UPDATE db_temp.dbo.tbl_Sabannet_Detail "
							strSql = strSql & " 		SET CS_STATUS = '003', TenStatus = 'C' "
							strSql = strSql & " 		WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"'  "
							strSql = strSql & "		END "
						End If
						dbget.Execute strSql, AssignedRow

						If (AssignedRow > 0) Then
							iInputCnt = iInputCnt + AssignedRow
						End If
					Next
					rw "판매자Q&A입력및수정건수 : " & iInputCnt
					GetCSSellerQnA_lotteon = True
				Else
					response.write "내역없음(0)<br />"

					GetCSSellerQnA_lotteon = True
					Set rsltList = Nothing
					Set strObj = Nothing
					exit function
				End If
			Set rsltList = nothing
		Else
			rw strObj.rsltMsg
		End If
	Set strObj = nothing
End Function

Function GetCSQnA_11st1010(selldate, sugi)
	Dim sellsite : sellsite = "11st1010"
	Dim apiUrl, xmlSelldate, strParam, Nodes, SubNodes
	Dim objXML, xmlDOM, iRbody, strSql, jParam, trCd, objData, addDateParam
	Dim masterCnt, objMasterListXML, objMasterOneXML, questionPrefix, AssignedRow, iInputCnt, Cval
	Dim dataList, strObj, returnCode, iMessage
	Dim SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM
	Dim counselLgroup, counselMgroup, counselTitle, qaMethod
	GetCSQnA_11st1010 = False
	MALL_ID = "11번가"
	xmlSelldate = Replace(selldate, "-", "")

	qaMethod = "00"
	If sugi = "Y" then
		qaMethod = "02"
	End If
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://api.11st.co.kr/rest/prodqnaservices/prodqnalist/"&xmlSelldate&"/"&xmlSelldate&"/" & qaMethod		'처리여부 | 00 : 전체조회, 01 : 답변완료조회, 02 : 미답변조회
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey","a2319e071dbc304243ee60abd07e9664"
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				Set Nodes = xmlDOM.getElementsByTagName("ns2:productQna")
					For each SubNodes in Nodes
						ProductNM = ""
						questionPrefix = "["& SubNodes.getElementsByTagName("qnaDtlsCdNm")(0).Text &"] "		'문의유형
						CS_STATUS = "001"
						SabanetNum = SubNodes.getElementsByTagName("brdInfoNo")(0).Text			'QnA 글번호 | 답변을 update 하실 경우 필요합니다.
						MALL_USER_ID = SubNodes.getElementsByTagName("memID")(0).Text			'고객ID
						CS_GUBUN = "DIRECT"														'c#으로 작업한 것외에 제휴API이용은 전부 DIRECT처리
						SUBJECT = questionPrefix & SubNodes.getElementsByTagName("brdInfoSbjct")(0).Text			'제목
						CNTS = SubNodes.getElementsByTagName("brdInfoCont")(0).Text				'질문내용
						REG_DM = SubNodes.getElementsByTagName("createDt")(0).Text				'문의일자
						If SubNodes.getElementsByTagName("answerYn")(0).Text = "Y" Then			'처리상태 | Y : 답변을 완료한 상태입니다, N : 미답변 상태입니다.
							CS_STATUS = "003"
						Else
							CS_STATUS = "001"
						End If
						PRODUCT_ID = SubNodes.getElementsByTagName("brdInfoClfNo")(0).Text			'상품번호
						ProductNM = SubNodes.getElementsByTagName("prdNm")(0).Text					'상품명
						If (SubNodes.getElementsByTagName("ordNoDe").length > 0) then
							OrderID			= SubNodes.getElementsByTagName("ordNoDe")(0).Text	'주문번호 | 구매여부가 'Y'인 경우 노출
						End If

						' rw SubNodes.getElementsByTagName("answerCont")(0).Text				'답변내용
						' rw SubNodes.getElementsByTagName("answerDt")(0).Text				'처리일자 | 답변을 update 한 날짜입니다.
						' rw SubNodes.getElementsByTagName("buyYn")(0).Text					'구매여부 | Y : 질문자가 상품을 구매한 상태, N : 구매자가 상품을 구매안한 상태
						' rw SubNodes.getElementsByTagName("dispYn")(0).Text					'전시상태 | Y : 전시, N : 전시안함
						' rw SubNodes.getElementsByTagName("memNM")(0).Text					'고객이름
						' rw SubNodes.getElementsByTagName("qnaDtlsCd")(0).Text				'문의유형코드 |  01 : 상품 ,02 : 배송 ,03 : 반품/환불/취소 ,04 : 교환/변경 ,05 : 기타
						' If (SubNodes.getElementsByTagName("ordStlEndDt").length > 0) then
						' 	rw "BB : " & SubNodes.getElementsByTagName("ordStlEndDt")(0).Text	'결제일시 | 구매여부가 'Y'인 경우 노출
						' End If

						strSql = ""
						strSql = strSql & " IF NOT Exists(SELECT * FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"' )"
						strSql = strSql & " 	BEGIN "
						strSql = strSql & " 		INSERT INTO db_temp.dbo.tbl_Sabannet_Detail "
						strSql = strSql & " 		(SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM, MALL_PROD_ID) VALUES "
						strSql = strSql & " 		('"&SabanetNum&"', '"&MALL_ID&"', '"&html2db(MALL_USER_ID)&"', '"&CS_STATUS&"', '"&REG_DM&"', '"&PRODUCT_ID&"', '"&html2db(SUBJECT)&"', '"&html2db(CNTS)&"', '"&INS_NM&"', null, '"&RPLY_CNTS&"', '"&UPD_NM&"', null, '"&CS_GUBUN&"', '"&PRODUCT_ID&"', '"&OrderID&"', '"&SEND_DM&"', '"&html2db(ProductNM)&"', '"&PRODUCT_ID&"') "
						strSql = strSql & "		END "
						If CS_STATUS = "003" Then
							strSql = strSql & " ELSE"
							strSql = strSql & " 	BEGIN "
							strSql = strSql & " 		UPDATE db_temp.dbo.tbl_Sabannet_Detail "
							strSql = strSql & " 		SET CS_STATUS = '003', TenStatus = 'C' "
							strSql = strSql & " 		WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"'  "
							strSql = strSql & "		END "
						End If
						dbget.Execute strSql, AssignedRow
						If (AssignedRow > 0) Then
							iInputCnt = iInputCnt + AssignedRow
						End If
					Next
					rw "Q&A입력및수정건수 : " & iInputCnt
					GetCSQnA_11st1010 = True
				Set Nodes = nothing
			Set xmlDOM = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If
	Set objXML = nothing
End Function

Function resCSQnA_11st1010(iNum, iRply, iPrdno)
	Dim objXML, strSql, iMessage, iRbody, strObj, returnCode, strParam, xmlDOM
	Dim sellsite : sellsite = "11st1010"
	strParam = ""
	strParam = strParam & "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
	strParam = strParam & "<ProductQna>"
	strParam = strParam & "	<answerCont><![CDATA["&iRply&"]]></answerCont>"				'답변내용
	strParam = strParam & "</ProductQna>"
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "PUT", "http://api.11st.co.kr/rest/prodqnaservices/prodqnaanswer/"&iNum&"/"&iPrdno
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey","a2319e071dbc304243ee60abd07e9664"
		objXML.send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				returnCode  = xmlDOM.getElementsByTagName("resultCode").item(0).text
				If (returnCode = "200") OR (returnCode = "210")  Then
					strSql = ""
					strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
					strSql = strSql & " SET CS_STATUS = '003' "
					strSql = strSql & " ,TenStatus = 'C' "
					strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
					strSql = strSql & " and SellSite = '"& sellsite &"' "
					dbget.Execute strSql
					resCSQnA_11st1010 = True
					rw iNum & " 답변완료"
				Else
					response.write BinaryToText(objXML.ResponseBody, "euc-kr")
				End If
			Set strObj = nothing
		Else
			rw "req : " & strParam
			rw iNum & "오류 : 종료"
			rw "message : " & BinaryToText(objXML.ResponseBody, "euc-kr")
			Set objXML = Nothing
'			dbget.close : response.end
		End If
	Set objXML= nothing
End Function

Function GetCSQnA_skstoa(selldate)
	Dim sellsite : sellsite = "skstoa"
	Dim apiUrl, xmlSelldate, strParam, obj
	Dim objXML, xmlDOM, iRbody, strSql, jParam, trCd, objData, addDateParam
	Dim masterCnt, objMasterListXML, objMasterOneXML, questionPrefix, AssignedRow, iInputCnt, Cval
	Dim dataList, strObj, returnCode, iMessage, i, j, counselDtList
	Dim SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM
	Dim counselLgroup, counselMgroup, counselTitle
	GetCSQnA_skstoa = False
	MALL_ID = "SKSTOA"
	xmlSelldate = Replace(selldate, "-", "")
	addDateParam = ""
	addDateParam = addDateParam & "&bDate="&xmlSelldate		'조회시작일자
	addDateParam = addDateParam & "&eDate="&xmlSelldate		'조회종료일자

 	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", skstoaAPIURL & "/partner/counsel/list?linkCode="&skstoalinkCode&"&entpCode="&skstoaentpCode&"&entpId="&skstoaentpId&"&entpPass="&skstoaentpPass & addDateParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
'		response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
		 			iInputCnt = 0
					CS_STATUS = ""
		 			Set dataList = strObj.counselAll
		 				For i=0 to dataList.length-1
							ProductNM = ""
							questionPrefix = "["& dataList.get(i).counselList.lgroupName & "_" & dataList.get(i).counselList.mgroupName &"] "		'상담대분류명_상담중분류명
							CS_STATUS = "001"
							SabanetNum = dataList.get(i).counselList.counselSeq			'접수번호
							MALL_USER_ID = ""											'고객ID
							CS_GUBUN = "DIRECT"											'c#으로 작업한 것외에 제휴API이용은 전부 DIRECT처리
							SUBJECT = dataList.get(i).counselList.sgroupName			'상담소분류명
							PRODUCT_ID = dataList.get(i).counselList.goodsCode			'상품코드
							ProductNM = dataList.get(i).counselList.goodsName			'상품명
							OrderID = dataList.get(i).counselList.orderNo				'주문번호

							Set counselDtList = dataList.get(i).counselDtList
								For j=0 to counselDtList.length-1
									CNTS	= ""
									REG_DM	= ""
									CNTS	= Trim(counselDtList.get(j).procNote)			'처리내역
									REG_DM	= Replace(LEFT(counselDtList.get(j).procDate, 19), "/", "-")	'처리시간

									If counselDtList.get(j).dtDoFlagCode = "25" Then			'처리단계코드
										CS_STATUS = "001"
									Else
										CS_STATUS = "003"
									End If

									strSql = ""
									strSql = strSql & " IF NOT Exists(SELECT * FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"' )"
									strSql = strSql & " 	BEGIN "
									strSql = strSql & " 		INSERT INTO db_temp.dbo.tbl_Sabannet_Detail "
									strSql = strSql & " 		(SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM, MALL_PROD_ID) VALUES "
									strSql = strSql & " 		('"&SabanetNum&"', '"&MALL_ID&"', '"&html2db(MALL_USER_ID)&"', '"&CS_STATUS&"', '"&REG_DM&"', '"&PRODUCT_ID&"', '"&html2db(SUBJECT)&"', '"&html2db(CNTS)&"', '"&INS_NM&"', null, '"&RPLY_CNTS&"', '"&UPD_NM&"', null, '"&CS_GUBUN&"', '"&PRODUCT_ID&"', '"&OrderID&"', '"&SEND_DM&"', '"&html2db(ProductNM)&"', '"&PRODUCT_ID&"') "
									strSql = strSql & "		END "
									If CS_STATUS = "003" Then
										strSql = strSql & " ELSE"
										strSql = strSql & " 	BEGIN "
										strSql = strSql & " 		UPDATE db_temp.dbo.tbl_Sabannet_Detail "
										strSql = strSql & " 		SET CS_STATUS = '003', TenStatus = 'C' "
										strSql = strSql & " 		WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"'  "
										strSql = strSql & "		END "
									End If
									dbget.Execute strSql, AssignedRow

									If (AssignedRow > 0) Then
										iInputCnt = iInputCnt + AssignedRow
									End If
								Next
							Set counselDtList = nothing

							If dataList.get(i).counselList.doFlag = "27" Then
								strSql = ""
								strSql = strSql & "	UPDATE db_temp.dbo.tbl_Sabannet_Detail "
								strSql = strSql & "	SET CS_STATUS = '003', TenStatus = 'C' "
								strSql = strSql & "	WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"'  "
								dbget.Execute strSql
							End If
		 				Next
		 				rw "Q&A입력및수정건수 : " & iInputCnt
		 				GetCSQnA_skstoa = True
		 			Set dataList = nothing
		 		Else
		 			response.write "내역없음(0)<br />"

		 			GetCSQnA_skstoa = True
		 			Set strObj = Nothing
		 			Set objXML = Nothing
		 			exit function
		 		End If
		 	Set strObj = nothing
		Else
		 	response.write BinaryToText(objXML.ResponseBody,"utf-8")
		End If
	Set objXML = nothing
End Function

Function resCSQnA_skstoa(iNum, iRply)
	Dim objXML, strSql, iMessage, iRbody, strObj, returnCode, strParam
	Dim sellsite : sellsite = "skstoa"
	strParam = ""
	strParam = strParam & "linkCode=" & skstoalinkCode				'#연결코드 | SKB에서 부여한 연결코드
	strParam = strParam & "&entpCode=" & skstoaentpCode				'#업체코드 | SKB에서 부여한 업체코드 6자리
	strParam = strParam & "&entpId=" & skstoaentpId					'#업체사용자ID | SKB에서 부여한 업체사용자 ID
	strParam = strParam & "&entpPass=" & skstoaentpPass				'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
	strParam = strParam & "&counselSeq=" & iNum						'#접수번호
	strParam = strParam & "&doFlag=27" 								'#처리단계코드 | 26:업체처리, 27:업체완료 ||| 업체에서 주문/클레임 처리로 상담이 완료된 경우 27:업체완료, 스토아와 추가 확인이 필요한 경우 26:업체처리로 연동해주시면 됩니다
	strParam = strParam & "&procNote=" & iRply						'#처리내역

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/counsel/proc", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
'		response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					strSql = ""
					strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
					strSql = strSql & " SET CS_STATUS = '003' "
					strSql = strSql & " ,TenStatus = 'C' "
					strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
					strSql = strSql & " and SellSite = '"& sellsite &"' "
					dbget.Execute strSql
					resCSQnA_skstoa = True
					rw iNum & " 답변완료"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
				End If
			Set strObj = nothing
		Else
			rw "req : " & strParam
			rw iNum & "오류 : 종료"
			rw "message : " & BinaryToText(objXML.ResponseBody,"utf-8")
			Set objXML = Nothing
'			dbget.close : response.end
		End If
	Set objXML= nothing
End Function

Function GetCSQnA_shintvshopping(selldate)
	Dim sellsite : sellsite = "shintvshopping"
	Dim apiUrl, xmlSelldate, strParam, obj
	Dim objXML, xmlDOM, iRbody, strSql, jParam, trCd, objData, addDateParam
	Dim masterCnt, objMasterListXML, objMasterOneXML, questionPrefix, AssignedRow, iInputCnt, Cval
	Dim dataList, strObj, returnCode, iMessage
	Dim SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM
	Dim counselLgroup, counselMgroup, counselTitle
	GetCSQnA_shintvshopping = False
	MALL_ID = "신세계TV쇼핑"
	xmlSelldate = Replace(selldate, "-", "")
	addDateParam = ""
	addDateParam = addDateParam & "&fromDate="&xmlSelldate		'FROM날짜
	addDateParam = addDateParam & "&toDate="&xmlSelldate		'TO날짜
	addDateParam = addDateParam & "&counselListGb=02"			'상담조회구분 | 00 : 전체조회, 01 : 답변완료조회, 02 : 미답변완료 조회(Default : 00)

 	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", shintvshoppingAPIURL & "/partner/counsel/cust-counsel-list?linkCode="&linkCode&"&entpCode="&entpCode&"&entpId="&entpId&"&entpPass="&entpPass & addDateParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					iInputCnt = 0
					Set dataList = strObj.custCounselList
						For i=0 to dataList.length-1
							ProductNM = ""
							questionPrefix = "주문Q&A"
							CS_STATUS = "001"
							SabanetNum = dataList.get(i).counselSeq		'고객상담순번
							counselLgroup = dataList.get(i).counselLgroup	'상담대분류
							counselMgroup = dataList.get(i).counselMgroup	'상담중분류

							MALL_USER_ID = ""
							CS_GUBUN	= "DIRECT"								'c#으로 작업한 것외에 제휴API이용은 전부 DIRECT처리
							On Error Resume Next
								counselTitle = LEFT(Trim(dataList.get(i).title), 80)
								If Err.number <> 0 Then
									counselTitle = ""
								End If
							On Error Goto 0

							If counselTitle = "" Then
								SUBJECT		= "[" & questionPrefix & "] " & counselLgroup & " " & counselMgroup
							Else
								SUBJECT		= "[" & questionPrefix & "_" & counselLgroup & "_" & counselMgroup & "] " & counselTitle
							End If
							CNTS		= Trim(dataList.get(i).procNote)		'상담내용
							REG_DM		= LEFT(dataList.get(i).procDate, 19)	'처리일시

							On Error Resume Next
								OrderID = dataList.get(i).orderNo			'주문번호
								If Err.number <> 0 Then
									OrderID = ""
								End If
							On Error Goto 0

							If OrderID <> "" Then
								strSql = ""
								strSql = strSql & " IF NOT Exists(SELECT * FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"' )"
								strSql = strSql & " 	BEGIN "
								strSql = strSql & " 		INSERT INTO db_temp.dbo.tbl_Sabannet_Detail "
								strSql = strSql & " 		(SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM, MALL_PROD_ID) VALUES "
								strSql = strSql & " 		('"&SabanetNum&"', '"&MALL_ID&"', '"&html2db(MALL_USER_ID)&"', '"&CS_STATUS&"', '"&REG_DM&"', '"&PRODUCT_ID&"', '"&html2db(SUBJECT)&"', '"&html2db(CNTS)&"', '"&INS_NM&"', null, '"&RPLY_CNTS&"', '"&UPD_NM&"', null, '"&CS_GUBUN&"', '"&PRODUCT_ID&"', '"&OrderID&"', '"&SEND_DM&"', '"&html2db(ProductNM)&"', '"&PRODUCT_ID&"') "
								strSql = strSql & "		END "
								dbget.Execute strSql, AssignedRow

								If (AssignedRow > 0) Then
									iInputCnt = iInputCnt + AssignedRow
								End If
							End If
						Next
						rw "주문Q&A입력및수정건수 : " & iInputCnt
						GetCSQnA_shintvshopping = True
					Set dataList = nothing
				Else
					response.write "내역없음(0)<br />"

					GetCSQnA_shintvshopping = True
					Set strObj = Nothing
					Set objXML = Nothing
					exit function
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
		End If
	Set objXML = nothing
End Function

Function GetCSQnA_shintvshopping_complete(selldate)
	Dim sellsite : sellsite = "shintvshopping"
	Dim apiUrl, xmlSelldate, strParam, obj
	Dim objXML, xmlDOM, iRbody, strSql, jParam, trCd, objData, addDateParam
	Dim masterCnt, objMasterListXML, objMasterOneXML, questionPrefix, AssignedRow, iInputCnt, Cval
	Dim dataList, strObj, returnCode, iMessage
	Dim SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM
	Dim counselLgroup, counselMgroup, counselTitle
	GetCSQnA_shintvshopping_complete = False
	MALL_ID = "신세계TV쇼핑"
	xmlSelldate = Replace(selldate, "-", "")
	addDateParam = ""
	addDateParam = addDateParam & "&fromDate="&xmlSelldate		'FROM날짜
	addDateParam = addDateParam & "&toDate="&xmlSelldate		'TO날짜
	addDateParam = addDateParam & "&counselListGb=01"			'상담조회구분 | 00 : 전체조회, 01 : 답변완료조회, 02 : 미답변완료 조회(Default : 00)

 	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", shintvshoppingAPIURL & "/partner/counsel/cust-counsel-list?linkCode="&linkCode&"&entpCode="&entpCode&"&entpId="&entpId&"&entpPass="&entpPass & addDateParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					iInputCnt = 0
					Set dataList = strObj.custCounselList
						For i=0 to dataList.length-1
							SabanetNum = dataList.get(i).counselSeq		'고객상담순번
							strSql =  ""
							strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
							strSql = strSql & " SET CS_STATUS = '003', TenStatus = 'C' "
							strSql = strSql & " WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"'  "
							dbget.Execute strSql
							GetCSQnA_shintvshopping_complete = True
						Next
						rw "주문Q&A수정건수(Complete) : " & iInputCnt
						GetCSQnA_shintvshopping_complete = True
					Set dataList = nothing
				Else
					response.write "내역없음(0)<br />"

					GetCSQnA_shintvshopping_complete = True
					Set strObj = Nothing
					Set objXML = Nothing
					exit function
				End If
			Set strObj = nothing
		Else
			response.write BinaryToText(objXML.ResponseBody,"utf-8")
		End If
	Set objXML = nothing
End Function

Function resCSQnA_shintvshopping(iNum, iRply)
	Dim objXML, strSql, iMessage, iRbody, strObj, returnCode, strParam
	Dim sellsite : sellsite = "shintvshopping"
	strParam = ""
	strParam = strParam & "linkCode=" & linkCode					'#연결코드 | 샵링커: SLINK [ TCODE.LGROUP : xxx ]
	strParam = strParam & "&entpCode=" & entpCode					'#업체코드 | 신세계TV쇼핑에서 부여한 업체코드 6자리
	strParam = strParam & "&entpId=" & entpId						'#업체사용자ID | 신세계TV쇼핑에서 부여한 업체사용자 ID
	strParam = strParam & "&entpPass=" & entpPass					'#업체PASSWORD | 신세계TV쇼핑에서 등록한 업체사용자 비밀번호
	strParam = strParam & "&counselSeq=" & iNum						'#고객상담순번
	strParam = strParam & "&procNote=" & iRply						'#처리내역

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/counsel/cust-counsel-proc", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		'response.write BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					strSql = ""
					strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
					strSql = strSql & " SET CS_STATUS = '003' "
					strSql = strSql & " ,TenStatus = 'C' "
					strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
					strSql = strSql & " and SellSite = '"& sellsite &"' "
					dbget.Execute strSql
					resCSQnA_shintvshopping = True
					rw iNum & " 답변완료"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
				End If
			Set strObj = nothing
		Else
			rw "req : " & strParam
			rw iNum & "오류 : 종료"
			rw "message : " & BinaryToText(objXML.ResponseBody,"utf-8")
			Set objXML = Nothing
'			dbget.close : response.end
		End If
	Set objXML= nothing
End Function

function GetCSQnA_nvstorefarm(selldate, isellsite)
	Dim sellsite
	If isellsite = "nvstorefarm" Then
		sellsite = "nvstorefarm"
	ElseIf isellsite = "nvstoregift" Then
		sellsite = "nvstoregift"
	ElseIf isellsite = "Mylittlewhoopee" Then
		sellsite = "Mylittlewhoopee"
	Else
		sellsite = "nvstoremoonbangu"
	End If
	dim xmlURL, strRst, reqID
	dim objXML, xmlDOM, masterCnt
	dim i, j, k
	dim startdate, enddate, questionPrefix
	dim AssignedRow, iInputCnt, objMasterListXML, objMasterOneXML
	dim strSql
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd
	dim ResponseType, QuestionType
	Dim SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM
	GetCSQnA_nvstorefarm = False

	If isellsite = "nvstorefarm" Then
		MALL_ID = "스토어팜(item)"
	ElseIf isellsite = "Mylittlewhoopee" Then
		MALL_ID = "스토어팜캣앤독(item)"
	ElseIf isellsite = "nvstoregift" Then
		MALL_ID = "스토어팜선물하기(item)"
	Else
		MALL_ID = "스토어팜문방구(item)"
	End If
	iServ	= "QuestionAnswerService"
	iCcd	= "GetQuestionAnswerList"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If sellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf sellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
	strRst = strRst & "	<soapenv:Header/>"
	strRst = strRst & "	<soapenv:Body>"
	strRst = strRst & "		<shop:GetQuestionAnswerListRequest>"
	strRst = strRst & "			<shop:RequestID>"&reqID&"</shop:RequestID>"
	strRst = strRst & "			<shop:AccessCredentials>"
	strRst = strRst & "				<shop:AccessLicense>"&iaccessLicense&"</shop:AccessLicense>"
	strRst = strRst & "				<shop:Timestamp>"&iTimestamp&"</shop:Timestamp>"
	strRst = strRst & "				<shop:Signature>"&isignature&"</shop:Signature>"
	strRst = strRst & "			</shop:AccessCredentials>"
	strRst = strRst & "			<shop:Version>2.0</shop:Version>"
	strRst = strRst & "			<SellerId>"&reqID&"</SellerId>"
	strRst = strRst & "			<FromDate>"&selldate&"</FromDate>"
	strRst = strRst & "			<ToDate>"&selldate&"</ToDate>"
'	strRst = strRst & "			<Answered></Answered>"		'선택 | 답변여부 Y or N or 비움
'	strRst = strRst & "			<Page></Page>"				'선택 | 빈값인 경우 기본값으로 1page조회
	strRst = strRst & "		</shop:GetQuestionAnswerListRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", xmlURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
		objXML.send(strRst)

		If objXML.Status <> "200" then
			response.write "ERROR : 통신오류" & objXML.Status
			dbget.close : response.end
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
			If ResponseType <> "SUCCESS" Then
				rw "오류 : 종료"
				rw "message : " & xmlDOM.getElementsByTagName("n:Message")(0).Text
				rw "Detail : " & xmlDOM.getElementsByTagName("n:Detail")(0).Text
				Set xmlDOM = Nothing
				Set objXML = Nothing
				dbget.close : response.end
			Else
				If xmlDOM.getElementsByTagName("n:QuestionAnswerList").length < 1 Then
					response.write "내역없음(0)<br />"

					GetCSQnA_nvstorefarm = True
					Set xmlDOM = Nothing
					Set objXML = Nothing
					Exit Function
				Else
					masterCnt = xmlDOM.getElementsByTagName("n:QuestionAnswerList").length
					rw "건수(" & masterCnt & ") "
					iInputCnt = 0
					set objMasterListXML = xmlDOM.getElementsByTagName("n:QuestionAnswerList")
						For Each objMasterOneXML In objMasterListXML
							SabanetNum = objMasterOneXML.getElementsByTagName("n:QuestionAnswerId").item(0).text	'상품 Q&A ID
							MALL_USER_ID = Trim(objMasterOneXML.getElementsByTagName("n:WriterId").item(0).text)	'마스킹된 질문자 ID

							Select Case Trim(objMasterOneXML.getElementsByTagName("n:QuestionType").item(0).text)	'PROD : 상품, DLV : 배송, RTN : 반품, EXCG : 교환, RFND : 환불, ETC : 기타
								Case "PROD"	QuestionType = "상품"
								Case "DLV"	QuestionType = "배송"
								Case "RTN"	QuestionType = "반품"
								Case "EXCG"	QuestionType = "교환"
								Case "RFND"	QuestionType = "환불"
								Case "ETC"	QuestionType = "기타"
								Case Else	QuestionType = "기타"
							End Select

							CS_GUBUN = "DIRECT"		'c#으로 작업한 것외에 제휴API이용은 전부 DIRECT처리
							questionPrefix ="상품Q&A"
							SUBJECT = "[" & questionPrefix & "_" & QuestionType & "] " & LEFT(Trim(objMasterOneXML.getElementsByTagName("n:Question").item(0).text), 100)
							CNTS = Trim(objMasterOneXML.getElementsByTagName("n:Question").item(0).text)			'문의 내용
							REG_DM = objMasterOneXML.getElementsByTagName("n:CreateDate").item(0).text				'질문 등록일 YYYY-MM-DD 형식

							If Trim(objMasterOneXML.getElementsByTagName("n:Answered").item(0).text) = "Y" Then
								CS_STATUS = "003"
							Else
								CS_STATUS = "001"
							End If

							PRODUCT_ID = Trim(objMasterOneXML.getElementsByTagName("n:ProductId").item(0).text)		'상품 ID
							ProductNM = Trim(objMasterOneXML.getElementsByTagName("n:ProductName").item(0).text)	'상품명

							' rw objMasterOneXML.getElementsByTagName("n:QuestionAnswerId").item(0).text	'상품 Q&A ID
							' rw objMasterOneXML.getElementsByTagName("n:CreateDate").item(0).text			'질문 등록일 YYYY-MM-DD 형식
							' rw objMasterOneXML.getElementsByTagName("n:QuestionType").item(0).text		'PROD : 상품, DLV : 배송, RTN : 반품, EXCG : 교환, RFND : 환불, ETC : 기타
							' rw objMasterOneXML.getElementsByTagName("n:Subject").item(0).text				'제목
							' rw objMasterOneXML.getElementsByTagName("n:Question").item(0).text			'문의 내용
							' rw objMasterOneXML.getElementsByTagName("n:Answer").item(0).text				'답변 내용
							' rw objMasterOneXML.getElementsByTagName("n:Purchased").item(0).text			'구매한 YN
							' rw objMasterOneXML.getElementsByTagName("n:ProductId").item(0).text			'상품 ID
							' rw objMasterOneXML.getElementsByTagName("n:ProductName").item(0).text			'상품명
							' rw objMasterOneXML.getElementsByTagName("n:WriterId").item(0).text			'마스킹된 질문자 ID
							strSql = ""
							strSql = strSql & " IF NOT Exists(SELECT * FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"' )"
							strSql = strSql & " 	BEGIN "
							strSql = strSql & " 		INSERT INTO db_temp.dbo.tbl_Sabannet_Detail "
							strSql = strSql & " 		(SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM, MALL_PROD_ID) VALUES "
							strSql = strSql & " 		('"&SabanetNum&"', '"&MALL_ID&"', '"&html2db(MALL_USER_ID)&"', '"&CS_STATUS&"', '"&REG_DM&"', '"&PRODUCT_ID&"', '"&html2db(SUBJECT)&"', '"&html2db(CNTS)&"', '"&INS_NM&"', null, '"&RPLY_CNTS&"', '"&UPD_NM&"', null, '"&CS_GUBUN&"', '"&PRODUCT_ID&"', '"&OrderID&"', '"&SEND_DM&"', '"&html2db(ProductNM)&"', '"&PRODUCT_ID&"') "
							strSql = strSql & "		END "
						If CS_STATUS = "003" Then
							strSql = strSql & " ELSE"
							strSql = strSql & " 	BEGIN "
							strSql = strSql & " 		UPDATE db_temp.dbo.tbl_Sabannet_Detail "
							strSql = strSql & " 		SET CS_STATUS = '003', TenStatus = 'C' "
							strSql = strSql & " 		WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"'  "
							strSql = strSql & "		END "
						End If
							dbget.Execute strSql, AssignedRow

							If (AssignedRow > 0) Then
								iInputCnt = iInputCnt + AssignedRow
							End If
						Next
						rw "상품Q&A입력및수정건수 : " & iInputCnt
						GetCSQnA_nvstorefarm = True
					set objMasterListXML = nothing
				End If
			End If
		Set xmlDOM = nothing
	Set objXML = nothing
End Function

function GetCSOrderQnA_nvstorefarm(selldate, isellsite)
	Dim sellsite
	If isellsite = "nvstorefarm" Then
		sellsite = "nvstorefarm"
	ElseIf isellsite = "nvstoregift" Then
		sellsite = "nvstoregift"
	ElseIf isellsite = "Mylittlewhoopee" Then
		sellsite = "Mylittlewhoopee"
	Else
		sellsite = "nvstoremoonbangu"
	End If
	dim xmlURL, strRst, reqID
	dim objXML, xmlDOM, masterCnt
	dim i, j, k
	dim startdate, enddate, questionPrefix
	dim AssignedRow, iInputCnt, objMasterListXML, objMasterOneXML
	dim strSql
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd
	dim ResponseType, QuestionType
	Dim SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM
	GetCSOrderQnA_nvstorefarm = False
	If isellsite = "nvstorefarm" Then
		MALL_ID = "스토어팜(order)"
	ElseIf isellsite = "nvstoregift" Then
		MALL_ID = "스토어팜선물하기(order)"
	ElseIf isellsite = "Mylittlewhoopee" Then
		MALL_ID = "스토어팜캣앤독(order)"
	Else
		MALL_ID = "스토어팜문방구(order)"
	End If

	iServ	= "CustomerInquiryService"
	iCcd	= "GetCustomerInquiryList"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If sellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf sellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:cus=""http://customerinquiry.shopn.platform.nhncorp.com/"">"
	strRst = strRst & "	<soapenv:Header/>"
	strRst = strRst & "	<soapenv:Body>"
	strRst = strRst & "		<cus:GetCustomerInquiryListRequest>"
	strRst = strRst & "			<cus:RequestID>"&reqID&"</cus:RequestID>"
	strRst = strRst & "			<cus:AccessCredentials>"
	strRst = strRst & "				<cus:AccessLicense>"&iaccessLicense&"</cus:AccessLicense>"
	strRst = strRst & "				<cus:Timestamp>"&iTimestamp&"</cus:Timestamp>"
	strRst = strRst & "				<cus:Signature>"&isignature&"</cus:Signature>"
	strRst = strRst & "			</cus:AccessCredentials>"
	strRst = strRst & "			<cus:DetailLevel>Full</cus:DetailLevel>"	'돌려 받는 데이터의 상세 정도(Compact/Full). 기본값은 "Full"이다.
	strRst = strRst & "			<cus:Version>1.0</cus:Version>"
	strRst = strRst & "			<ServiceType>SHOPN</ServiceType>"			'서비스 타입 코드. 스마트스토어 판매자는 SHOPN을 입력한다("A.3.2 서비스 타입 코드" 참고). 네이버페이 가맹점은 CHECKOUT 입력
	strRst = strRst & "			<MallID>"&reqID&"</MallID>"					'판매자 ID
	strRst = strRst & "			<InquiryTimeFrom>"&selldate&"T00:00:00</InquiryTimeFrom>"		'조회 시작 일시
	strRst = strRst & "			<InquiryTimeTo>"& Left(DateAdd("d", 1, CDate(selldate)), 10)&"T00:00:00</InquiryTimeTo>"			'조회 종료 일시
'	strRst = strRst & "			<IsAnswered></IsAnswered>"					'답변여부 Y or N
	strRst = strRst & "		</cus:GetCustomerInquiryListRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", xmlURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
		objXML.send(strRst)

		If objXML.Status <> "200" then
			response.write "ERROR : 통신오류" & objXML.Status
			dbget.close : response.end
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
			If ResponseType <> "SUCCESS" Then
				rw "오류 : 종료"
				rw "message : " & xmlDOM.getElementsByTagName("n:Message")(0).Text
				rw "Detail : " & xmlDOM.getElementsByTagName("n:Detail")(0).Text
				Set xmlDOM = Nothing
				Set objXML = Nothing
				dbget.close : response.end
			Else
				If xmlDOM.getElementsByTagName("CustomerInquiry").length < 1 Then
					response.write "내역없음(0)<br />"

					GetCSOrderQnA_nvstorefarm = True
					Set xmlDOM = Nothing
					Set objXML = Nothing
					Exit Function
				Else
					masterCnt = xmlDOM.getElementsByTagName("CustomerInquiry").length
					rw "건수(" & masterCnt & ") "
					iInputCnt = 0
					set objMasterListXML = xmlDOM.getElementsByTagName("CustomerInquiry")
						For Each objMasterOneXML In objMasterListXML
							SabanetNum = objMasterOneXML.getElementsByTagName("InquiryID").item(0).text				'쇼핑 문의 번호
							MALL_USER_ID = Trim(objMasterOneXML.getElementsByTagName("CustomerName").item(0).text)	'고객명
							QuestionType = objMasterOneXML.getElementsByTagName("Category").item(0).text			'문의 카테고리
							CS_GUBUN = "DIRECT"		'c#으로 작업한 것외에 제휴API이용은 전부 DIRECT처리
							questionPrefix ="주문Q&A"
							SUBJECT = "[" & questionPrefix & "_" & QuestionType & "] " & Trim(objMasterOneXML.getElementsByTagName("Title").item(0).text)	'쇼핑 문의 제목
							CNTS = Trim(objMasterOneXML.getElementsByTagName("InquiryContent").item(0).text)		'문의 내용
							REG_DM = LEFT(objMasterOneXML.getElementsByTagName("InquiryDateTime").item(0).text, 10)

							If Trim(objMasterOneXML.getElementsByTagName("IsAnswered").item(0).text) = "true" Then		'답변 여부
								CS_STATUS = "003"
							Else
								CS_STATUS = "001"
							End If
							PRODUCT_ID = Trim(objMasterOneXML.getElementsByTagName("ProductID").item(0).text)		'상품 번호
							ProductNM = Trim(objMasterOneXML.getElementsByTagName("ProductName").item(0).text)		'상품명
'							OrderID = objMasterOneXML.getElementsByTagName("OrderID").item(0).text					'주문번호
							If Instr(objMasterOneXML.getElementsByTagName("ProductOrderID").item(0).text, ",") > 0 Then
								OrderID = Split(objMasterOneXML.getElementsByTagName("ProductOrderID").item(0).text, ",")(0)
							Else
								OrderID = objMasterOneXML.getElementsByTagName("ProductOrderID").item(0).text			'상품 주문 번호
							End If

							' rw objMasterOneXML.getElementsByTagName("InquiryID").item(0).text				'쇼핑 문의 번호
							' rw objMasterOneXML.getElementsByTagName("OrderID").item(0).text				'주문번호
							' rw objMasterOneXML.getElementsByTagName("ProductOrderID").item(0).text		'상품 주문 번호
							' rw objMasterOneXML.getElementsByTagName("ProductName").item(0).text			'상품명
							' rw objMasterOneXML.getElementsByTagName("ProductID").item(0).text				'상품 번호
							' rw objMasterOneXML.getElementsByTagName("ProductOrderOption").item(0).text	'상품 옵션
							' rw objMasterOneXML.getElementsByTagName("CustomerID").item(0).text			'고객ID. 끝 세 자리는 별표(*)로 처리한다.
							' rw objMasterOneXML.getElementsByTagName("Title").item(0).text					'쇼핑 문의 제목
							' rw objMasterOneXML.getElementsByTagName("Category").item(0).text				'문의 카테고리
							' rw objMasterOneXML.getElementsByTagName("InquiryDateTime").item(0).text		'문의 일시
							' rw objMasterOneXML.getElementsByTagName("InquiryContent").item(0).text		'문의 내용
							' rw objMasterOneXML.getElementsByTagName("AnswerContent").item(0).text			'답변 내용
							' rw objMasterOneXML.getElementsByTagName("IsAnswered").item(0).text			'답변 여부
							' rw objMasterOneXML.getElementsByTagName("CustomerName").item(0).text			'고객명
							strSql = ""
							strSql = strSql & " IF NOT Exists(SELECT * FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"' )"
							strSql = strSql & " 	BEGIN "
							strSql = strSql & " 		INSERT INTO db_temp.dbo.tbl_Sabannet_Detail "
							strSql = strSql & " 		(SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM, MALL_PROD_ID) VALUES "
							strSql = strSql & " 		('"&SabanetNum&"', '"&MALL_ID&"', '"&html2db(MALL_USER_ID)&"', '"&CS_STATUS&"', '"&REG_DM&"', '"&PRODUCT_ID&"', '"&html2db(SUBJECT)&"', '"&html2db(CNTS)&"', '"&INS_NM&"', null, '"&RPLY_CNTS&"', '"&UPD_NM&"', null, '"&CS_GUBUN&"', '"&PRODUCT_ID&"', '"&OrderID&"', '"&SEND_DM&"', '"&html2db(ProductNM)&"', '"&PRODUCT_ID&"') "
							strSql = strSql & "		END "
						If CS_STATUS = "003" Then
							strSql = strSql & " ELSE"
							strSql = strSql & " 	BEGIN "
							strSql = strSql & " 		UPDATE db_temp.dbo.tbl_Sabannet_Detail "
							strSql = strSql & " 		SET CS_STATUS = '003', TenStatus = 'C' "
							strSql = strSql & " 		WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"'  "
							strSql = strSql & "		END "
						End If
							dbget.Execute strSql, AssignedRow

							If (AssignedRow > 0) Then
								iInputCnt = iInputCnt + AssignedRow
							End If
						Next
						rw "주문Q&A입력및수정건수 : " & iInputCnt
						GetCSOrderQnA_nvstorefarm = True
					set objMasterListXML = nothing
				End If
			End If
		Set xmlDOM = nothing
	Set objXML = nothing
End Function

Function resCSQnA_nvstorefarm(iNum, iRply, isellsite)
	Dim xmlURL, strRst, reqID
	Dim objXML, xmlDOM, strSql
	Dim iaccessLicense, iTimestamp, isignature, iServ, iCcd
	Dim ResponseType
	resCSQnA_nvstorefarm = False
	iServ	= "QuestionAnswerService"
	iCcd	= "ManageQuestionAnswer"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If isellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf isellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
	strRst = strRst & "	<soap:Header/>"
	strRst = strRst & "	<soap:Body>"
	strRst = strRst & "		<shop:ManageQuestionAnswerRequest>"
	strRst = strRst & "			<shop:RequestID>"&reqID&"</shop:RequestID>"
	strRst = strRst & "			<shop:AccessCredentials>"
	strRst = strRst & "				<shop:AccessLicense>"&iaccessLicense&"</shop:AccessLicense>"
	strRst = strRst & "				<shop:Timestamp>"&iTimestamp&"</shop:Timestamp>"
	strRst = strRst & "				<shop:Signature>"&isignature&"</shop:Signature>"
	strRst = strRst & "			</shop:AccessCredentials>"
	strRst = strRst & "			<shop:Version>2.0</shop:Version>"
	strRst = strRst & "			<SellerId>"&reqID&"</SellerId>"
	strRst = strRst & "			<QuestionAnswerId>"&iNum&"</QuestionAnswerId>"
	strRst = strRst & "			<Answer><![CDATA["&iRply&"]]></Answer>"
	strRst = strRst & "		</shop:ManageQuestionAnswerRequest>"
	strRst = strRst & "	</soap:Body>"
	strRst = strRst & "</soap:Envelope>"

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", xmlURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
		objXML.send(strRst)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(objXML.responseText)
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text

				If ResponseType = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
					strSql = strSql & " SET CS_STATUS = '003' "
					strSql = strSql & " ,TenStatus = 'C' "
					strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
					strSql = strSql & " and SellSite = '"& isellsite &"' "
					dbget.Execute strSql
					resCSQnA_nvstorefarm = True
					rw iNum & " 답변완료(item)"
				Else
					rw iNum & "오류 : 종료"
					rw "message : " & xmlDOM.getElementsByTagName("n:Message")(0).Text
					rw "Detail : " & xmlDOM.getElementsByTagName("n:Detail")(0).Text
					Set xmlDOM = Nothing
					Set objXML = Nothing
'					dbget.close : response.end
				End If
			Set xmlDOM = nothing
	 	Else
	 		response.write "ERROR : 통신오류" & objXML.Status
		End If
	Set objXML= nothing
End Function

Function resCSOrderQnA_nvstorefarm(iNum, iRply, isellsite)
	Dim xmlURL, strRst, reqID
	Dim objXML, xmlDOM, strSql
	Dim iaccessLicense, iTimestamp, isignature, iServ, iCcd
	Dim ResponseType
	resCSOrderQnA_nvstorefarm = False
	iServ	= "CustomerInquiryService"
	iCcd	= "AnswerCustomerInquiry"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If isellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf isellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:cus=""http://customerinquiry.shopn.platform.nhncorp.com/"">"
	strRst = strRst & "	<soapenv:Header/>"
	strRst = strRst & "	<soapenv:Body>"
	strRst = strRst & "		<cus:AnswerCustomerInquiryRequest>"
	strRst = strRst & "			<cus:RequestID>"&reqID&"</cus:RequestID>"
	strRst = strRst & "			<cus:AccessCredentials>"
	strRst = strRst & "				<cus:AccessLicense>"&iaccessLicense&"</cus:AccessLicense>"
	strRst = strRst & "				<cus:Timestamp>"&iTimestamp&"</cus:Timestamp>"
	strRst = strRst & "				<cus:Signature>"&isignature&"</cus:Signature>"
	strRst = strRst & "			</cus:AccessCredentials>"
	strRst = strRst & "			<cus:DetailLevel>Full</cus:DetailLevel>"	'돌려 받는 데이터의 상세 정도(Compact/Full). 기본값은 "Full"이다.
	strRst = strRst & "			<cus:Version>1.0</cus:Version>"
	strRst = strRst & "			<MallID>"&reqID&"</MallID>"					'판매자 ID
	strRst = strRst & "			<ServiceType>SHOPN</ServiceType>"			'서비스 타입 코드. 스마트스토어 판매자는 SHOPN을 입력한다("A.3.2 서비스 타입 코드" 참고). 네이버페이 가맹점은 CHECKOUT 입력
	strRst = strRst & "			<InquiryID>"&iNum&"</InquiryID>"			'쇼핑 문의 번호
	strRst = strRst & "			<AnswerContent><![CDATA["&iRply&"]]></AnswerContent>"	'답변 내용
'	strRst = strRst & "			<AnswerContentID>?</AnswerContentID>"		'답변 번호 (답변 메시지를 수정하는 경우에는 필수)
	strRst = strRst & "			<ActionType>INSERT</ActionType>"			'명령어타입코드 (INSERT : 답변등록, UPDATE : 답변수정 )
'	strRst = strRst & "			<AnswerTempleteID>?</AnswerTempleteID>"		'문의 답변 템플릿 일련번호
	strRst = strRst & "		</cus:AnswerCustomerInquiryRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", xmlURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
		objXML.send(strRst)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(objXML.responseText)
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text

				If ResponseType = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
					strSql = strSql & " SET CS_STATUS = '003' "
					strSql = strSql & " ,TenStatus = 'C' "
					strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
					strSql = strSql & " and SellSite = '"& isellsite &"' "
					dbget.Execute strSql
					resCSOrderQnA_nvstorefarm = True
					rw iNum & " 답변완료(order)"
				Else
					rw iNum & "오류 : 종료"
					rw "message : " & xmlDOM.getElementsByTagName("n:Message")(0).Text
					rw "Detail : " & xmlDOM.getElementsByTagName("n:Detail")(0).Text
					Set xmlDOM = Nothing
					Set objXML = Nothing
'					dbget.close : response.end
				End If
			Set xmlDOM = nothing
	 	Else
	 		response.write "ERROR : 통신오류" & objXML.Status
		End If
	Set objXML= nothing
End Function

Function resCSQnA_lotteon(iNum, iRply)
	Dim sellsite : sellsite = "lotteon"
	Dim apiUrl, xmlSelldate, strParam, obj
	Dim objXML, xmlDOM, iRbody, apiKey, strSql, jParam, trCd, objData
	Dim masterCnt, objMasterListXML, objMasterOneXML, questionPrefix, AssignedRow, iInputCnt, Cval
	Dim rsltList, strObj, returnCode, spdNo, sitmNo
	resCSQnA_lotteon = False

	apiUrl = getApiUrl("lotteon")
	apiKey = getApiKey("lotteon")
	apiUrl = apiUrl & "/v1/openapi/product/v1/product/qna/reply"

	If application("Svr_Info") = "Dev" Then
		trCd = "LO10001101"
	Else
		trCd = "LD304013"
	End If

	strSql = ""
	strSql = strSql & " SELECT TOP 1 PRODUCT_ID, COMPAYNY_GOODS_CD "
	strSql = strSql & " FROM db_temp.dbo.tbl_Sabannet_Detail "
	strSql = strSql & " WHERE SabanetNum = '"& iNum &"'"
	strSql = strSql & " and SellSite = 'lotteon'"
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not rsget.Eof Then
		spdNo	= rsget("PRODUCT_ID")
		sitmNo	= rsget("COMPAYNY_GOODS_CD")
	End If
	rsget.Close

	'// =======================================================================
	Set obj = jsObject()
		Set obj("ansInfo")= jsArray()
			Set obj("ansInfo")(null) = jsObject()
				obj("ansInfo")(null)("trGrpCd") = "SR"			'#거래처그룹코드
				obj("ansInfo")(null)("trNo") = trCd				'#거래처번호
				obj("ansInfo")(null)("lrtrNo") = ""				'하위거래처번호
				obj("ansInfo")(null)("pdQnaNo") = iNum			'#상품QnA번호
				obj("ansInfo")(null)("spdNo") = spdNo			'#판매자상품코드
				obj("ansInfo")(null)("sitmNo") = sitmNo			'#판매자단품코드
				obj("ansInfo")(null)("ansCnts") = iRply			'#답변내용
				jParam = obj.jsString
	Set obj = nothing

	'// 데이타 가져오기
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(jParam)
		If objXML.Status <> "200" Then
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
			dbget.close : response.end
		Else
			objData = BinaryToText(objXML.ResponseBody,"utf-8")
		End If

	'// =======================================================================
	'// Json 파싱
	Set strObj = JSON.parse(objData)
		returnCode		= strObj.returnCode
		If returnCode = "0000" Then
			strSql = ""
			strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
			strSql = strSql & " SET CS_STATUS = '003' "
			strSql = strSql & " ,TenStatus = 'C' "
			strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
			strSql = strSql & " and SellSite = '"& sellsite &"' "
			dbget.Execute strSql
			resCSQnA_lotteon = True
			rw iNum & " 답변완료(item)"
		Else
			rw iNum & "오류 : 종료"
			rw "message : " & strObj.message
			Set xmlDOM = Nothing
			Set objXML = Nothing
			dbget.close : response.end
		End If
	Set strObj = nothing
End Function

Function resCSSellerQnA_lotteon(iNum, iRply)
	Dim sellsite : sellsite = "lotteon"
	Dim apiUrl, xmlSelldate, strParam, obj
	Dim objXML, xmlDOM, iRbody, apiKey, strSql, jParam, trCd, objData
	Dim masterCnt, objMasterListXML, objMasterOneXML, questionPrefix, AssignedRow, iInputCnt, Cval
	Dim rsltList, strObj, returnCode
	resCSSellerQnA_lotteon = False

	apiUrl = getApiUrl("lotteon")
	apiKey = getApiKey("lotteon")
	apiUrl = apiUrl & "/v1/openapi/customer/v1/updateSellerInquiry"
	'// =======================================================================
	Set obj = jsObject()
		obj("slrInqNo") = iNum			'#판매자문의번호
		obj("ansCnts") = iRply			'#답변내용
		jParam = obj.jsString
	Set obj = nothing

	'// 데이타 가져오기
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(jParam)

		If objXML.Status <> "200" Then
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
			dbget.close : response.end
		Else
			objData = BinaryToText(objXML.ResponseBody,"utf-8")
		End If
	'// =======================================================================
	'// Json 파싱
	Set strObj = JSON.parse(objData)
		returnCode		= strObj.rsltCd
		If returnCode = "0000" Then
			strSql = ""
			strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
			strSql = strSql & " SET CS_STATUS = '003' "
			strSql = strSql & " ,TenStatus = 'C' "
			strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
			strSql = strSql & " and SellSite = '"& sellsite &"' "
			dbget.Execute strSql
			resCSSellerQnA_lotteon = True
			rw iNum & " 답변완료(seller)"
		Else
			rw iNum & "오류 : 종료"
			rw "message : " & strObj.rsltMsg
			Set xmlDOM = Nothing
			Set objXML = Nothing
			dbget.close : response.end
		End If
	Set strObj = nothing
End Function

Public Function getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iserv, ioper)
	Dim cryptoLib, oLicense, osecretKey, otimeStamp, osignature
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
		If (application("Svr_Info") = "Dev") Then
			iaccessLicense = "01000100004b035a25d67f991849cad1c7042b8da528d13e9ddce6878f2e43ac88080e0a5e" 'AccessLicense Key 입력, PDF파일참조
			osecretKey = "AQABAAAWPWagCrPjFQnFEtxs5j+oyZFwuzomdNq0XZSricPuMw=="  'SecreKey 입력, PDF파일참조
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		Else
			iaccessLicense = "010001000019133c715650b9c85b820961612f2b90b431ddd8654b42c097c4df1a43d0be09" 'AccessLicense Key 입력, PDF파일참조
			osecretKey = "AQABAADX6Hz/wORFJS5pSIy4KQXkH83gC9G1aXChxBjcnUMqWw=="  'SecreKey 입력, PDF파일참조
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		End If
	Set cryptoLib = nothing
End Function

function GetCSCheckStatus(byVal sellsite, byVal csGubun, byRef LastCheckDate, byRef isSuccess)
	dim strSql

    strSql = " IF NOT Exists("
    strSql = strSql + " 	select LastcheckDate"
    strSql = strSql + " 	from db_temp.[dbo].[tbl_xSite_TMPCS_timestamp]"
    strSql = strSql + " 	where sellsite='" + CStr(sellsite) + "' and csGubun='" + CStr(csGubun) + "'"
	strSql = strSql + " )"
	strSql = strSql + " BEGIN"
	strSql = strSql + "		insert into db_temp.[dbo].[tbl_xSite_TMPCS_timestamp](sellsite, csGubun, lastcheckdate, issuccess, LastUpdate) "
	strSql = strSql + "		values('" & sellsite & "', '" & csGubun & "', '" & Left(DateAdd("d", -1, Now()), 10) & "', 'N', getdate()) "
	strSql = strSql + " END"
	dbget.Execute strSql

	strSql = " select convert(varchar(10), LastCheckDate, 121) as LastCheckDate, isSuccess from db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	strSql = strSql + " 	where sellsite='" + CStr(sellsite) + "' and csGubun='" + CStr(csGubun) + "'"

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		LastCheckDate = rsget("LastCheckDate")
		isSuccess = rsget("isSuccess")
	rsget.Close
end function

function SetCSCheckStatus(sellsite, csGubun, LastCheckDate, isSuccess)
	dim strSql

	strSql = " update db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	strSql = strSql + " set lastcheckdate = '" & LastCheckDate & "', issuccess = '" & isSuccess & "', LastUpdate = getdate() "
	strSql = strSql + " where sellsite = '" + CStr(sellsite) + "' and csGubun='" + CStr(csGubun) + "' "
	''response.write strSql
	dbget.Execute strSql
end function

function GetLotteAuthNo()
	dim lotteAPIURL, lotteAuthNo, lottenTenID, tenBrandCd, tenDlvCd, CertPasswd, iisql

	GetLotteAuthNo = ""

	IF application("Svr_Info")="Dev" THEN
		'lotteAPIURL = "http://openapidev.lotte.com"	'' 테스트서버
		lotteAPIURL = "http://openapitest.lotte.com"	'' 테스트서버
		tenBrandCd = "14846"	'텐바(임시)
		tenDlvCd = "513484"		'배송정책코드
		CertPasswd = "1234"		'Dev는 비번 : 1234
	Else
		lotteAPIURL = "https://openapi.lotte.com"		'' 실서버
		tenBrandCd = "155112"	'텐바이텐
		tenDlvCd = "513484"
		CertPasswd = "store101010*"
	End if
	lottenTenID = "124072"					'텐바이텐ID

	Dim updateAuth, dbAuthNo
	iisql = "select top 1 isnull(iniVal, '') as iniVal, lastupdate "&VbCRLF
	iisql = iisql & " from db_etcmall.dbo.tbl_outmall_ini"&VbCRLF
	iisql = iisql & " where mallid='lotteCom'"&VbCRLF
	iisql = iisql & " and inikey='auth'"
	rsget.CursorLocation = adUseClient
	rsget.Open iisql, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.Eof then
	    dbAuthNo	= rsget("iniVal")
	    updateAuth	= rsget("lastupdate")
	end if
	rsget.close

	If DateDiff("h", updateAuth, now()) > 12 OR dbAuthNo = "" then
		dim objXML, xmlDOM
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", lotteAPIURL & "/openapi/createCertification.lotte?strUserId=" & lottenTenID & "&strPassWd="&CertPasswd&"", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			'XML을 담을 DOM 객체 생성
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

			''on Error Resume Next
				GetLotteAuthNo = xmlDOM.getElementsByTagName("SubscriptionId").item(0).text
				if Err<>0 then
					Response.Write "통신오류(XML)"
					Response.End
				end if
			''on Error Goto 0
				iisql = "update db_etcmall.dbo.tbl_outmall_ini "&VbCRLF
				iisql = iisql & " set iniVal='"&GetLotteAuthNo&"'"&VbCRLF
				iisql = iisql & " ,lastupdate=getdate()"&VbCRLF
				iisql = iisql & " where mallid='lotteCom'"&VbCRLF
				iisql = iisql & " and inikey='auth'"
				dbget.Execute iisql

			Set xmlDOM = Nothing
		else
			Response.Write "통신오류"
			Response.End
		end if
		Set objXML = Nothing
	Else
		GetLotteAuthNo = dbAuthNo
	End If
end function

Function GetLotteimallAuthNo()
	GetLotteimallAuthNo = ""
	'// 롯데아이몰API 연동서버 URL
	Dim ltiMallAPIURL, ltiMallAuthNo, ltiMallTenID, tenBrandCd, tenDlvCd, tenDlvFreeCd
	Dim iisql
	IF application("Svr_Info") = "Dev" THEN
		'ltiMallAPIURL = "http://openapidev.lotteimall.com"	'' 테스트서버
		ltiMallAPIURL = "http://openapitst.lotteimall.com"	'' 테스트서버
		tenDlvCd = "23725"
		tenDlvFreeCd = "577045"
	Else
		ltiMallAPIURL = "https://openapi.lotteimall.com"		'' 실서버
		tenDlvCd = "23725"
		tenDlvFreeCd = "577045"
	End if
	ltiMallTenID = "011799LT"


	'// 롯데아이몰 인증코드 확인(매일 업데이트; 어플리케이션변수에 저장)
	Dim updateAuth, dbAuthNo
	iisql = "select top 1 isnull(iniVal, '') as iniVal, lastupdate "&VbCRLF
	iisql = iisql & " from db_etcmall.dbo.tbl_outmall_ini"&VbCRLF
	iisql = iisql & " where mallid='lotteimall'"&VbCRLF
	iisql = iisql & " and inikey='auth'"
	rsget.Open iisql, dbget, 1
	if not rsget.Eof then
		dbAuthNo	= rsget("iniVal")
		updateAuth	= rsget("lastupdate")
	end if
	rsget.close

	If DateDiff("h", updateAuth, now()) > 12 OR dbAuthNo = "" then
		Dim objXML, xmlDOM
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.Open "GET", ltiMallAPIURL & "/openapi/createCertification.lotte?strUserId=" & ltiMallTenID & "&strPassWd=store101010*", False
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.Send()
			If objXML.Status = "200" Then
				Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				'On Error Resume Next
					GetLotteimallAuthNo = xmlDOM.getElementsByTagName("SubscriptionId").item(0).text		'인증번호 저장
					if Err<>0 then
						Response.Write "통신오류(XML)"
						Response.End
					end if

					iisql = "update db_etcmall.dbo.tbl_outmall_ini "&VbCRLF
					iisql = iisql & " set iniVal='"&GetLotteimallAuthNo&"'"&VbCRLF
					iisql = iisql & " ,lastupdate=getdate()"&VbCRLF
					iisql = iisql & " where mallid='lotteimall'"&VbCRLF
					iisql = iisql & " and inikey='auth'"
					dbget.Execute iisql
				'On Error Goto 0
				Set xmlDOM = Nothing
			else
				Response.Write "통신오류"
				Response.End
			end if
		Set objXML = Nothing
	Else
		GetLotteimallAuthNo = dbAuthNo
	End If
end function

Function getCSAnswerComplete(imallid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT SabanetNum, isNull(RPLY_CNTS, '') as RPLY_CNTS "
	strSql = strSql & " FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) "
	strSql = strSql & " WHERE SellSite = '"& imallid &"' "
	strSql = strSql & " AND TenStatus = 'S' "
	strSql = strSql & " and CS_STATUS = '001' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		getCSAnswerComplete = rsget.getRows
	End If
	rsget.close
End Function

Function get11stCSAnswerComplete(imallid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT SabanetNum, isNull(RPLY_CNTS, '') as RPLY_CNTS, PRODUCT_ID "
	strSql = strSql & " FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) "
	strSql = strSql & " WHERE SellSite = '"& imallid &"' "
	strSql = strSql & " AND TenStatus = 'S' "
	strSql = strSql & " and CS_STATUS = '001' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		get11stCSAnswerComplete = rsget.getRows
	End If
	rsget.close
End Function

Function getNvstorefarmCSAnswerComplete(imallid, icsGubun)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT SabanetNum, isNull(RPLY_CNTS, '') as RPLY_CNTS "
	strSql = strSql & " FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) "
	strSql = strSql & " WHERE SellSite = '"& imallid &"' "

	If imallid = "nvstorefarm" Then
		If icsGubun = "item" Then
			strSql = strSql & " and MALL_ID = '스토어팜(item)' "
		Else
			strSql = strSql & " and MALL_ID = '스토어팜(order)' "
		End If
	ElseIf imallid = "nvstoregift" Then
		If icsGubun = "item" Then
			strSql = strSql & " and MALL_ID = '스토어팜선물하기(item)' "
		Else
			strSql = strSql & " and MALL_ID = '스토어팜선물하기(order)' "
		End If
	ElseIf imallid = "Mylittlewhoopee" Then
		If icsGubun = "item" Then
			strSql = strSql & " and MALL_ID = '스토어팜캣앤독(item)' "
		Else
			strSql = strSql & " and MALL_ID = '스토어팜캣앤독(order)' "
		End If
	Else
		If icsGubun = "item" Then
			strSql = strSql & " and MALL_ID = '스토어팜문방구(item)' "
		Else
			strSql = strSql & " and MALL_ID = '스토어팜문방구(order)' "
		End If
	End If
	strSql = strSql & " AND TenStatus = 'S' "
	strSql = strSql & " and CS_STATUS = '001' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		getNvstorefarmCSAnswerComplete = rsget.getRows
	End If
	rsget.close
End Function

Function getLotteonCSAnswerComplete(imallid, icsGubun)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT SabanetNum, isNull(RPLY_CNTS, '') as RPLY_CNTS "
	strSql = strSql & " FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) "
	strSql = strSql & " WHERE SellSite = '"& imallid &"' "
	If icsGubun = "item" Then
		strSql = strSql & " and MALL_ID = '롯데On(item)' "
	Else
		strSql = strSql & " and MALL_ID = '롯데On(seller)' "
	End If
	strSql = strSql & " AND TenStatus = 'S' "
	strSql = strSql & " and CS_STATUS = '001' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		getLotteonCSAnswerComplete = rsget.getRows
	End If
	rsget.close
End Function

Function getShintvshoppingCSAnswerComplete(imallid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT SabanetNum, isNull(RPLY_CNTS, '') as RPLY_CNTS "
	strSql = strSql & " FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) "
	strSql = strSql & " WHERE SellSite = '"& imallid &"' "
	strSql = strSql & " and MALL_ID = '신세계TV쇼핑' "
	strSql = strSql & " AND TenStatus = 'S' "
	strSql = strSql & " and CS_STATUS = '001' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		getShintvshoppingCSAnswerComplete = rsget.getRows
	End If
	rsget.close
End Function

Function getSkstoaCSAnswerComplete(imallid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT SabanetNum, isNull(RPLY_CNTS, '') as RPLY_CNTS "
	strSql = strSql & " FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) "
	strSql = strSql & " WHERE SellSite = '"& imallid &"' "
	strSql = strSql & " and MALL_ID = 'SKSTOA' "
	strSql = strSql & " AND TenStatus = 'S' "
	strSql = strSql & " and CS_STATUS = '001' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		getSkstoaCSAnswerComplete = rsget.getRows
	End If
	rsget.close
End Function
%>
