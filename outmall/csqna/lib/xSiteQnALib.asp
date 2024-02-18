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

''xml ���� ����
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
	MALL_ID = "�Ե�����"
	xmlSelldate = Replace(selldate, "-", "")

	If application("Svr_Info")="Dev" Then
		authNo = "c7ed8e97a9f5657ecde6d29455094ff2bb046cf7c864f0790615bd3298fb049931d500ed1614db2c3e1e17fc528b3269449545325572051804110afe770820aa"
	Else
		authNo = GetLotteAuthNo()
	End If

	xmlURL = "https://openapi.lotte.com"
	strParam = "subscriptionId=" & authNo & "&strSearchStrtDtime=" & xmlSelldate & "&strSearchEndDtime=" & xmlSelldate
	'// =======================================================================
	'// ����Ÿ ��������
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
					response.write "��������(0)<br />"

					GetCSQnA_lotteCom = True
					Set xmlDOM = Nothing
					Set objXML = Nothing
					Exit Function
				End If
				masterCnt = xmlDOM.getElementsByTagName("Response/Result/GoodsQuestInfo").length
				rw "�Ǽ�(" & masterCnt & ") "
				iInputCnt = 0
				set objMasterListXML = xmlDOM.selectNodes("/Response/Result/GoodsQuestInfo")
					For Each objMasterOneXML In objMasterListXML
						If Trim(objMasterOneXML.getElementsByTagName("ReceiptType").item(0).text) = "C"	Then	'Ÿ�� | C :������(�ֶ���), Q:��ǰQ
							questionPrefix = "������(�ֶ���)"
						Else
							questionPrefix ="��ǰQ&A"
						End If
						SabanetNum = Trim(objMasterOneXML.getElementsByTagName("ReceiptNo").item(0).text)		'����
						'rw Trim(objMasterOneXML.getElementsByTagName("Gubun").item(0).text)					'����
						'rw Trim(objMasterOneXML.getElementsByTagName("AskNm").item(0).text)					'�亯��
						MALL_USER_ID = Trim(objMasterOneXML.getElementsByTagName("QuestNm").item(0).text)		'������
						'rw Trim(objMasterOneXML.getElementsByTagName("CellHpNo").item(0).text)					'����ó
						Select Case Trim(objMasterOneXML.getElementsByTagName("Cval").item(0).text)				'�������� | ���ֶ��� �ϰ�� 12:�Ϲ���, 21:1:1E����, 22:�Ϲ�E����, ��ǰQ �ϰ�� 1:������, 2:����ǰ, 3:������/����, 4:��뼳��, 9��Ÿ
							Case "12"	Cval = "�Ϲ���"
							Case "21"	Cval = "1:1E����"
							Case "22"	Cval = "�Ϲ�E����"
							Case "1"	Cval = "������"
							Case "2"	Cval = "����ǰ"
							Case "3"	Cval = "������/����"
							Case "4"	Cval = "��뼳��"
							Case "9"	Cval = "��Ÿ"
							Case Else	Cval = "��Ÿ"
						End Select

						CS_GUBUN = "DIRECT"		'c#���� �۾��� �Ϳܿ� ����API�̿��� ���� DIRECTó��

						SUBJECT = "[" & questionPrefix & "_" & Cval & "] " & Trim(objMasterOneXML.getElementsByTagName("Subject").item(0).text)	'����
						CNTS = Trim(objMasterOneXML.getElementsByTagName("Content").item(0).text)				'���ǳ���
						'rw Trim(objMasterOneXML.getElementsByTagName("ReplyTitle").item(0).text)				'�亯����
						REG_DM = Trim(objMasterOneXML.getElementsByTagName("ReceiptDate").item(0).text)			'�����
						REG_DM = Left(REG_DM,4) & "-" & Mid(REG_DM,5,2) & "-" & Mid(REG_DM,7,2) & " " & Mid(REG_DM,9,2) & ":" & Mid(REG_DM,11,2) & ":" & Mid(REG_DM,13,2)

						If Trim(objMasterOneXML.getElementsByTagName("Result").item(0).text) = "02" Then		'ó������ | 02: ó���Ϸ�, 02�� ������ ���� ��ó��
							CS_STATUS = "003"
						Else
							CS_STATUS = "001"
						End If

						'rw Trim(objMasterOneXML.getElementsByTagName("ResultMsg").item(0).text)				'ó������
						PRODUCT_ID = Trim(objMasterOneXML.getElementsByTagName("GoodsNo").item(0).text)			'��ǰ��ȣ
						'rw Trim(objMasterOneXML.getElementsByTagName("MsgType").item(0).text)					'���Ǳ���(���������� �ٸ�)
						ProductNM = Trim(objMasterOneXML.getElementsByTagName("GoodsNm").item(0).text)			'��ǰ��
						'rw Trim(objMasterOneXML.getElementsByTagName("ResultDate").item(0).text)				'ó����

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
					rw "��ǰQ&A�Է¹׼����Ǽ� : " & iInputCnt
					GetCSQnA_lotteCom = True
				set objMasterListXML = nothing
			Set xmlDOM = nothing
		Else
			response.write "ERROR : ��ſ���" & objXML.Status
			dbget.close : response.end
		End If
	Set objXML= nothing

	If Err Then
		response.write "ERROR : ��ſ��� - " & Err.Description
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
	'// ����Ÿ ��������
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
						rw iNum & " �亯�Ϸ�"
					End If
				Else
					If (xmlDOM.getElementsByTagName("Response/Errors/Error/Code").length > 0) then
						Select Case xmlDOM.getElementsByTagName("Response/Errors/Error/Code").item(0).text
							Case "4001"	'�̹� �����Ϳ� �����Ͽ����ϴ�.
								strSql = ""
								strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
								strSql = strSql & " SET CS_STATUS = '003' "
								strSql = strSql & " ,TenStatus = 'C' "
								strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
								strSql = strSql & " and SellSite = 'lotteCom' "
								dbget.Execute strSql
								resCSQnA_lotteCom = True
								rw iNum & " �̹�ó���Ϸ�"
							Case Else
								response.write "ERROR : " & xmlDOM.getElementsByTagName("Response/Errors/Error/Message").item(0).text
						End Select
					End If
				End If
			Set xmlDOM = nothing
		Else
			response.write "ERROR : ��ſ���" & objXML.Status
		End If
	Set objXML= nothing

	If Err Then
		response.write "ERROR : ��ſ��� - " & Err.Description
	End If
End Function

Function GetCSQnA_lotteimall(selldate)
	Dim sellsite : sellsite = "lotteimall"
	Dim xmlURL, xmlSelldate, strParam
	Dim objXML, xmlDOM, iRbody, authNo, strSql
	Dim masterCnt, objMasterListXML, objMasterOneXML, questionPrefix, AssignedRow, iInputCnt, Cval
	Dim SabanetNum, MALL_ID, MALL_USER_ID, CS_STATUS, REG_DM, PRODUCT_ID, SUBJECT, CNTS, INS_NM, INS_DM, RPLY_CNTS, UPD_NM, UPD_DM, CS_GUBUN, COMPAYNY_GOODS_CD, OrderID, SEND_DM, ProductNM
	GetCSQnA_lotteimall = False
	MALL_ID = "�Ե����̸�"
	xmlSelldate = Replace(selldate, "-", "")

	If application("Svr_Info")="Dev" Then
		authNo = "4ef8af11a83ffe9129c2aeb3d799760d2c95aa3f7c29a4d683d82a0015f92d6b4e1f9da0f5fac137419c787f711dacf42dd065a35158a7b8657aeb6eb48e2cc3"
	Else
		authNo = GetLotteimallAuthNo()
	End If

	xmlURL = "https://openapi.lotteimall.com"
	strParam = "subscriptionId=" & authNo & "&req_start_dtime=" & xmlSelldate & "&req_end_dtime=" & xmlSelldate
	'// =======================================================================
	'// ����Ÿ ��������
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
					response.write "��������(0)<br />"

					GetCSQnA_lotteimall = True
					Set xmlDOM = Nothing
					Set objXML = Nothing
					Exit Function
				End If

				masterCnt = xmlDOM.getElementsByTagName("Response/Result/GoodsQuestInfo").length
				rw "�Ǽ�(" & masterCnt & ") "
				iInputCnt = 0

				set objMasterListXML = xmlDOM.selectNodes("/Response/Result/GoodsQuestInfo")
					For Each objMasterOneXML In objMasterListXML
						If Trim(objMasterOneXML.getElementsByTagName("ReceiptType").item(0).text) = "C"	Then	'Ÿ�� | C :������(�ֶ���), Q:��ǰQ
							questionPrefix = "������(�ֶ���)"
						Else
							questionPrefix ="��ǰQ&A"
						End If
						SabanetNum = Trim(objMasterOneXML.getElementsByTagName("ReceiptNo").item(0).text)		'����
						'rw Trim(objMasterOneXML.getElementsByTagName("Gubun").item(0).text)					'����
						'rw Trim(objMasterOneXML.getElementsByTagName("AskNm").item(0).text)					'�亯��
						MALL_USER_ID = Trim(objMasterOneXML.getElementsByTagName("QuestNm").item(0).text)		'������
						'rw Trim(objMasterOneXML.getElementsByTagName("CellHpNo").item(0).text)					'����ó

						Select Case Trim(objMasterOneXML.getElementsByTagName("Cval").item(0).text)				'�������� | ���ֶ��� �ϰ�� 12:�Ϲ���, 21:1:1E����, 22:�Ϲ�E����, ��ǰQ �ϰ�� 1:������, 2:����ǰ, 3:������/����, 4:��뼳��, 9��Ÿ
							Case "12"	Cval = "�Ϲ���"
							Case "21"	Cval = "1:1E����"
							Case "22"	Cval = "�Ϲ�E����"
							Case "1"	Cval = "������"
							Case "2"	Cval = "����ǰ"
							Case "3"	Cval = "������/����"
							Case "4"	Cval = "��뼳��"
							Case "5"	Cval = "��Ÿ"
							Case Else	Cval = "��Ÿ"
						End Select

						CS_GUBUN = "DIRECT"		'c#���� �۾��� �Ϳܿ� ����API�̿��� ���� DIRECTó��

						SUBJECT = "[" & questionPrefix & "_" & Cval & "] " & Trim(objMasterOneXML.getElementsByTagName("Subject").item(0).text)	'����
						CNTS = Trim(objMasterOneXML.getElementsByTagName("Content").item(0).text)				'���ǳ���
						'rw Trim(objMasterOneXML.getElementsByTagName("ReplyTitle").item(0).text)				'�亯����
						REG_DM = Trim(objMasterOneXML.getElementsByTagName("ReceiptDate").item(0).text)			'�����
						REG_DM = Left(REG_DM,4) & "-" & Mid(REG_DM,5,2) & "-" & Mid(REG_DM,7,2) & " " & Mid(REG_DM,9,2) & ":" & Mid(REG_DM,11,2) & ":" & Mid(REG_DM,13,2)

						If Trim(objMasterOneXML.getElementsByTagName("Result").item(0).text) = "02" Then		'ó������ | 02: ó���Ϸ�, 02�� ������ ���� ��ó��
							CS_STATUS = "003"
						Else
							CS_STATUS = "001"
						End If

						'rw Trim(objMasterOneXML.getElementsByTagName("ResultMsg").item(0).text)				'ó������
						PRODUCT_ID = Trim(objMasterOneXML.getElementsByTagName("GoodsNo").item(0).text)			'��ǰ��ȣ
						'rw Trim(objMasterOneXML.getElementsByTagName("MsgType").item(0).text)					'���Ǳ���(���������� �ٸ�)
						ProductNM = Trim(objMasterOneXML.getElementsByTagName("GoodsNm").item(0).text)			'��ǰ��
						'rw Trim(objMasterOneXML.getElementsByTagName("ResultDate").item(0).text)				'ó����

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
					rw "��ǰQ&A�Է¹׼����Ǽ� : " & iInputCnt
					GetCSQnA_lotteimall = True
				set objMasterListXML = nothing
			Set xmlDOM = nothing
		Else
			response.write "ERROR : ��ſ���" & objXML.Status
			dbget.close : response.end
		End If
	Set objXML= nothing

	If Err Then
		response.write "ERROR : ��ſ��� - " & Err.Description
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
	iRply = replace(replace(replace(iRply,"&","%26"), "+", "%2B"), "%", "����")

	xmlURL = "https://openapi.lotteimall.com"
	strParam = "subscriptionId=" & authNo & "&inq_no=" & iNum & "&ans_cont_type=5&inq_ans_cont=" & iRply & "&ans_disp_yn=Y&proc_type=1&memo="
	'// =======================================================================
	'// ����Ÿ ��������
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
						rw iNum & " �亯�Ϸ�"
					End If
				Else
					If (xmlDOM.getElementsByTagName("Response/Errors/Error/Code").length > 0) then
						Select Case xmlDOM.getElementsByTagName("Response/Errors/Error/Code").item(0).text
							Case "4001"	'�̹� �����Ϳ� �����Ͽ����ϴ�.
								strSql = ""
								strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
								strSql = strSql & " SET CS_STATUS = '003' "
								strSql = strSql & " ,TenStatus = 'C' "
								strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
								strSql = strSql & " and SellSite = 'lotteimall' "
								dbget.Execute strSql
								resCSQnA_lotteimall = True
								rw iNum & " �̹�ó���Ϸ�"
							Case Else
								response.write "ERROR : " & xmlDOM.getElementsByTagName("Response/Errors/Error/Message").item(0).text
						End Select
					End If
				End If
			Set xmlDOM = nothing
		Else
			response.write "ERROR : ��ſ���" & objXML.Status
		End If
	Set objXML= nothing

	If Err Then
		response.write "ERROR : ��ſ��� - " & Err.Description
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
	MALL_ID = "�Ե�On(item)"
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
		obj("trGrpCd") = "SR"			'#�ŷ�ó�׷��ڵ� | �Ϲݼ��� : SR
		obj("trNo") = trCd				'#�ŷ�ó��ȣ
		obj("lrtrNo") = ""				'�����ŷ�ó��ȣ
		obj("spdNo") = ""				'�Ǹ��ڻ�ǰ��ȣ
		obj("sitmNo") = ""				'�Ǹ��ڴ�ǰ��ȣ
		obj("regStrDttm") = xmlSelldate & "000000"	'#����Ͻ� ��ȸ�����Ͻ� [YYYYMMDDHH24MISS ex) 20190801150010]
		obj("regEndDttm") = xmlSelldate & "235959"	'#����Ͻ� ��ȸ�����Ͻ� [YYYYMMDDHH24MISS ex) 20190810150010]
		obj("qstTypCd") = null			'QnA ���� [�����ڵ� : QST_TYP_CD] | NULL�� ��쿡�� ��ü������ ��ȸ�Ѵ�. SZ_CAPA : ������/�뷮, DSGN_CLR : ������/����, DP_INFO : ��ǰ����, USE_EPN : ��뼳��, ETC : ��Ÿ
		obj("qnaStatCd") = null			'QnAó�������ڵ� [�����ڵ� : QNA_STAT_CD] NULL�� ��쿡�� ��ü ��ȸ�Ѵ�. NPROC : ��ó��, PROC : ó���Ϸ�, CC_TCTL : �������̰�
		obj("pageNo") = 1				'#������
		obj("rowsPerPage") = 100		'#���̴�Ǽ� (MAX 100)
		jParam = obj.jsString
	Set obj = nothing

	'// ����Ÿ ��������
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(jParam)

		If objXML.Status <> "200" Then
			response.write "ERROR : ��ſ���" & objXML.Status
			response.write "<script>alert('ERROR : ��ſ���.');</script>"
			dbget.close : response.end
		Else
			objData = BinaryToText(objXML.ResponseBody,"utf-8")
		End If

	'// =======================================================================
	'// Json �Ľ�
	Set strObj = JSON.parse(objData)
'rw objData
		returnCode		= strObj.returnCode
		If returnCode = "0000" Then
			iInputCnt = 0
			Set dataList = strObj.data
				If dataList.length > 0 Then
					For i=0 to dataList.length-1
						ProductNM = ""
						questionPrefix = "��ǰQ&A"
						SabanetNum = dataList.get(i).pdQnaNo		'#��ǰQnA��ȣ
						MALL_USER_ID = ""
						CS_GUBUN = "DIRECT"							'c#���� �۾��� �Ϳܿ� ����API�̿��� ���� DIRECTó��
						Select Case Trim(dataList.get(i).qstTypCd)	'QnA ���� [�����ڵ� : QST_TYP_CD] | SZ_CAPA : ������/�뷮, DSGN_CLR : ������/����, DP_INFO : ��ǰ����, USE_EPN : ��뼳��, ETC : ��Ÿ
							Case "SZ_CAPA"	Cval = "������/�뷮"
							Case "DSGN_CLR"	Cval = "������/����"
							Case "DP_INFO"	Cval = "��ǰ����"
							Case "USE_EPN"	Cval = "��뼳��"
							Case "ETC"		Cval = "��Ÿ"
							Case Else		Cval = "��Ÿ"
						End Select
						SUBJECT = "[" & questionPrefix & "_" & Cval & "] " & LEFT(Trim(dataList.get(i).qstCnts), 20)	'����
						CNTS = Trim(dataList.get(i).qstCnts)		'��������
						REG_DM = dataList.get(i).regDttm			'����Ͻ�
						REG_DM = Left(REG_DM,4) & "-" & Mid(REG_DM,5,2) & "-" & Mid(REG_DM,7,2) & " " & Mid(REG_DM,9,2) & ":" & Mid(REG_DM,11,2) & ":" & Mid(REG_DM,13,2)
						If Trim(dataList.get(i).qnaStatCd) = "PROC" Then		''#QnAó�������ڵ� [�����ڵ� : QNA_STAT_CD] NULL�� ��쿡�� ��ü ��ȸ�Ѵ�. | NPROC : ��ó��, PROC : ó���Ϸ�, CC_TCTL : �������̰�
							CS_STATUS = "003"
						Else
							CS_STATUS = "001"
						End If

						PRODUCT_ID = Trim(dataList.get(i).spdNo)			'�Ǹ��ڻ�ǰ��ȣ
						COMPAYNY_GOODS_CD = Trim(dataList.get(i).sitmNo)	'�Ǹ��ڴ�ǰ��ȣ

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
					rw "��ǰQ&A�Է¹׼����Ǽ� : " & iInputCnt
					GetCSQnA_lotteon = True
				Else
					response.write "��������(0)<br />"

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
	MALL_ID = "�Ե�On(seller)"
	xmlSelldate = Replace(selldate, "-", "")

	apiUrl = getApiUrl("lotteon")
	apiKey = getApiKey("lotteon")
	apiUrl = apiUrl & "/v1/openapi/customer/v1/getSellerInquiryList"
	'// =======================================================================
	Set obj = jsObject()
		obj("scStrtDt") = xmlSelldate			'#��ȸ�Ⱓ ������ [yyyymmdd : 20190801]
		obj("scEndDt") = xmlSelldate			'#��ȸ�Ⱓ ������ [yyyymmdd : 20190801]
'		obj("vocLcsfCd") = ""					'���������ڵ�
'		obj("slrInqProcStatCd") = ""			'�Ǹ��ڹ���ó�������ڵ� | ��ü:����, �亯:ANS, �̴亯:UNANS
'		obj("scKwd") = ""						'�˻���
'		obj("spdNo") = ""						'�Ǹ��� ��ǰ��ȣ
'		obj("spdNm") = ""						'�Ǹ��� ��ǰ��
'		obj("sitmNo") = ""						'�Ǹ��� ��ǰ��ȣ
'		obj("sitmNm") = ""						'�Ǹ��� ��ǰ��
'		obj("lrtrNo") = ""						'�ŷ�ó��ȣ
'		obj("pageNo") = ""						'1
'		obj("rowsPerPage") = ""					'50
		jParam = obj.jsString
	Set obj = nothing

	'// ����Ÿ ��������
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(jParam)

		If objXML.Status <> "200" Then
			response.write "ERROR : ��ſ���" & objXML.Status
			response.write "<script>alert('ERROR : ��ſ���.');</script>"
			dbget.close : response.end
		Else
			objData = BinaryToText(objXML.ResponseBody,"utf-8")
		End If
	'// =======================================================================
	'// Json �Ľ�
	Set strObj = JSON.parse(objData)
		returnCode		= strObj.rsltCd
		If returnCode = "0000" Then
			Set rsltList = strObj.rsltList
				iInputCnt = 0
				If rsltList.length > 0 Then
					For i=0 to rsltList.length-1
'						rw "aa : " & rsltList.length
						questionPrefix = "�Ǹ��ڹ���Q&A"
						SabanetNum		= rsltList.get(i).slrInqNo				'�Ǹ��ڹ��ǹ�ȣ
						MALL_USER_ID	= ""									'�Ե�On API�� �� ���� ����;
						CS_GUBUN		= "DIRECT"								'c#���� �۾��� �Ϳܿ� ����API�̿��� ���� DIRECTó��
						Cval			= rsltList.get(i).vocTypNm				'����������
						SUBJECT			= "[" & questionPrefix & "_" & Cval & "] " & LEFT(Trim(rsltList.get(i).inqTtl), 20)	'��������
						CNTS 			= Trim(rsltList.get(i).inqCnts)			'���ǳ���
						If SabanetNum = "2902954" Then
							CNTS 		= "������ 285�� ������?"
						End If
						REG_DM			= Trim(rsltList.get(i).accpDttm)		'�����Ͻ�
						REG_DM = Left(REG_DM,4) & "-" & Mid(REG_DM,5,2) & "-" & Mid(REG_DM,7,2) & " " & Mid(REG_DM,9,2) & ":" & Mid(REG_DM,11,2) & ":" & Mid(REG_DM,13,2)
						If Trim(rsltList.get(i).slrInqProcStatCd) = "ANS" Then	'�Ǹ��ڹ���ó�������ڵ�(ANS, UNANS)
							CS_STATUS = "003"
						Else
							CS_STATUS = "001"
						End If
						PRODUCT_ID		= Trim(rsltList.get(i).pdNo)			'��ǰ��ȣ
						ProductNM		= Trim(rsltList.get(i).pdNm)			'��ǰ��
						OrderID			= rsltList.get(i).odNo					'�ֹ���ȣ

						' rsltList.get(i).slrInqNo			'�Ǹ��� ���ǹ�ȣ
						' rsltList.get(i).vocLcsfCd			'���������ڵ�
						' rsltList.get(i).vocTypNm			'����������
						' rsltList.get(i).slrInqProcStatCd	'�Ǹ��ڹ���ó�������ڵ�(ANS, UNANS)
						' rsltList.get(i).slrInqProcStatNm	'�Ǹ��ڹ���ó�����¸�(�亯, �̴亯)
						' rsltList.get(i).inqTtl				'��������
						' rsltList.get(i).inqCnts				'���ǳ���
						' rsltList.get(i).odNo				'�ֹ���ȣ
						' rsltList.get(i).pdNo				'��ǰ��ȣ
						' rsltList.get(i).pdNm				'��ǰ��
						' rsltList.get(i).spdNo				'�Ǹ��� ��ǰ��ȣ
						' rsltList.get(i).spdNm				'�Ǹ��� ��ǰ��
						' rsltList.get(i).sitmNo				'�Ǹ��� ��ǰ��ȣ
						' rsltList.get(i).sitmNm				'�Ǹ��� ��ǰ��
						' rsltList.get(i).trNo				'�����ŷ�ó��ȣ
						' rsltList.get(i).trNm				'�����ŷ�ó��
						' rsltList.get(i).lrtrNo				'�����ŷ�ó��ȣ
						' rsltList.get(i).lrtrNm				'�����ŷ�ó��
						' rsltList.get(i).slrNo				'�Ǹ��ڹ�ȣ
						' rsltList.get(i).ansCnts				'�亯����
						' rsltList.get(i).ansRqPrd			'�亯�ҿ�Ⱓ
						' rsltList.get(i).accpDttm			'�����Ͻ�
						' rsltList.get(i).procDttm			'ó���Ͻ�

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
					rw "�Ǹ���Q&A�Է¹׼����Ǽ� : " & iInputCnt
					GetCSSellerQnA_lotteon = True
				Else
					response.write "��������(0)<br />"

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
	MALL_ID = "11����"
	xmlSelldate = Replace(selldate, "-", "")

	qaMethod = "00"
	If sugi = "Y" then
		qaMethod = "02"
	End If
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://api.11st.co.kr/rest/prodqnaservices/prodqnalist/"&xmlSelldate&"/"&xmlSelldate&"/" & qaMethod		'ó������ | 00 : ��ü��ȸ, 01 : �亯�Ϸ���ȸ, 02 : �̴亯��ȸ
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
						questionPrefix = "["& SubNodes.getElementsByTagName("qnaDtlsCdNm")(0).Text &"] "		'��������
						CS_STATUS = "001"
						SabanetNum = SubNodes.getElementsByTagName("brdInfoNo")(0).Text			'QnA �۹�ȣ | �亯�� update �Ͻ� ��� �ʿ��մϴ�.
						MALL_USER_ID = SubNodes.getElementsByTagName("memID")(0).Text			'��ID
						CS_GUBUN = "DIRECT"														'c#���� �۾��� �Ϳܿ� ����API�̿��� ���� DIRECTó��
						SUBJECT = questionPrefix & SubNodes.getElementsByTagName("brdInfoSbjct")(0).Text			'����
						CNTS = SubNodes.getElementsByTagName("brdInfoCont")(0).Text				'��������
						REG_DM = SubNodes.getElementsByTagName("createDt")(0).Text				'��������
						If SubNodes.getElementsByTagName("answerYn")(0).Text = "Y" Then			'ó������ | Y : �亯�� �Ϸ��� �����Դϴ�, N : �̴亯 �����Դϴ�.
							CS_STATUS = "003"
						Else
							CS_STATUS = "001"
						End If
						PRODUCT_ID = SubNodes.getElementsByTagName("brdInfoClfNo")(0).Text			'��ǰ��ȣ
						ProductNM = SubNodes.getElementsByTagName("prdNm")(0).Text					'��ǰ��
						If (SubNodes.getElementsByTagName("ordNoDe").length > 0) then
							OrderID			= SubNodes.getElementsByTagName("ordNoDe")(0).Text	'�ֹ���ȣ | ���ſ��ΰ� 'Y'�� ��� ����
						End If

						' rw SubNodes.getElementsByTagName("answerCont")(0).Text				'�亯����
						' rw SubNodes.getElementsByTagName("answerDt")(0).Text				'ó������ | �亯�� update �� ��¥�Դϴ�.
						' rw SubNodes.getElementsByTagName("buyYn")(0).Text					'���ſ��� | Y : �����ڰ� ��ǰ�� ������ ����, N : �����ڰ� ��ǰ�� ���ž��� ����
						' rw SubNodes.getElementsByTagName("dispYn")(0).Text					'���û��� | Y : ����, N : ���þ���
						' rw SubNodes.getElementsByTagName("memNM")(0).Text					'���̸�
						' rw SubNodes.getElementsByTagName("qnaDtlsCd")(0).Text				'���������ڵ� |  01 : ��ǰ ,02 : ��� ,03 : ��ǰ/ȯ��/��� ,04 : ��ȯ/���� ,05 : ��Ÿ
						' If (SubNodes.getElementsByTagName("ordStlEndDt").length > 0) then
						' 	rw "BB : " & SubNodes.getElementsByTagName("ordStlEndDt")(0).Text	'�����Ͻ� | ���ſ��ΰ� 'Y'�� ��� ����
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
					rw "Q&A�Է¹׼����Ǽ� : " & iInputCnt
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
	strParam = strParam & "	<answerCont><![CDATA["&iRply&"]]></answerCont>"				'�亯����
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
					rw iNum & " �亯�Ϸ�"
				Else
					response.write BinaryToText(objXML.ResponseBody, "euc-kr")
				End If
			Set strObj = nothing
		Else
			rw "req : " & strParam
			rw iNum & "���� : ����"
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
	addDateParam = addDateParam & "&bDate="&xmlSelldate		'��ȸ��������
	addDateParam = addDateParam & "&eDate="&xmlSelldate		'��ȸ��������

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
							questionPrefix = "["& dataList.get(i).counselList.lgroupName & "_" & dataList.get(i).counselList.mgroupName &"] "		'����з���_����ߺз���
							CS_STATUS = "001"
							SabanetNum = dataList.get(i).counselList.counselSeq			'������ȣ
							MALL_USER_ID = ""											'��ID
							CS_GUBUN = "DIRECT"											'c#���� �۾��� �Ϳܿ� ����API�̿��� ���� DIRECTó��
							SUBJECT = dataList.get(i).counselList.sgroupName			'���Һз���
							PRODUCT_ID = dataList.get(i).counselList.goodsCode			'��ǰ�ڵ�
							ProductNM = dataList.get(i).counselList.goodsName			'��ǰ��
							OrderID = dataList.get(i).counselList.orderNo				'�ֹ���ȣ

							Set counselDtList = dataList.get(i).counselDtList
								For j=0 to counselDtList.length-1
									CNTS	= ""
									REG_DM	= ""
									CNTS	= Trim(counselDtList.get(j).procNote)			'ó������
									REG_DM	= Replace(LEFT(counselDtList.get(j).procDate, 19), "/", "-")	'ó���ð�

									If counselDtList.get(j).dtDoFlagCode = "25" Then			'ó���ܰ��ڵ�
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
		 				rw "Q&A�Է¹׼����Ǽ� : " & iInputCnt
		 				GetCSQnA_skstoa = True
		 			Set dataList = nothing
		 		Else
		 			response.write "��������(0)<br />"

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
	strParam = strParam & "linkCode=" & skstoalinkCode				'#�����ڵ� | SKB���� �ο��� �����ڵ�
	strParam = strParam & "&entpCode=" & skstoaentpCode				'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
	strParam = strParam & "&entpId=" & skstoaentpId					'#��ü�����ID | SKB���� �ο��� ��ü����� ID
	strParam = strParam & "&entpPass=" & skstoaentpPass				'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
	strParam = strParam & "&counselSeq=" & iNum						'#������ȣ
	strParam = strParam & "&doFlag=27" 								'#ó���ܰ��ڵ� | 26:��üó��, 27:��ü�Ϸ� ||| ��ü���� �ֹ�/Ŭ���� ó���� ����� �Ϸ�� ��� 27:��ü�Ϸ�, ����ƿ� �߰� Ȯ���� �ʿ��� ��� 26:��üó���� �������ֽø� �˴ϴ�
	strParam = strParam & "&procNote=" & iRply						'#ó������

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
					rw iNum & " �亯�Ϸ�"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
				End If
			Set strObj = nothing
		Else
			rw "req : " & strParam
			rw iNum & "���� : ����"
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
	MALL_ID = "�ż���TV����"
	xmlSelldate = Replace(selldate, "-", "")
	addDateParam = ""
	addDateParam = addDateParam & "&fromDate="&xmlSelldate		'FROM��¥
	addDateParam = addDateParam & "&toDate="&xmlSelldate		'TO��¥
	addDateParam = addDateParam & "&counselListGb=02"			'�����ȸ���� | 00 : ��ü��ȸ, 01 : �亯�Ϸ���ȸ, 02 : �̴亯�Ϸ� ��ȸ(Default : 00)

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
							questionPrefix = "�ֹ�Q&A"
							CS_STATUS = "001"
							SabanetNum = dataList.get(i).counselSeq		'��������
							counselLgroup = dataList.get(i).counselLgroup	'����з�
							counselMgroup = dataList.get(i).counselMgroup	'����ߺз�

							MALL_USER_ID = ""
							CS_GUBUN	= "DIRECT"								'c#���� �۾��� �Ϳܿ� ����API�̿��� ���� DIRECTó��
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
							CNTS		= Trim(dataList.get(i).procNote)		'��㳻��
							REG_DM		= LEFT(dataList.get(i).procDate, 19)	'ó���Ͻ�

							On Error Resume Next
								OrderID = dataList.get(i).orderNo			'�ֹ���ȣ
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
						rw "�ֹ�Q&A�Է¹׼����Ǽ� : " & iInputCnt
						GetCSQnA_shintvshopping = True
					Set dataList = nothing
				Else
					response.write "��������(0)<br />"

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
	MALL_ID = "�ż���TV����"
	xmlSelldate = Replace(selldate, "-", "")
	addDateParam = ""
	addDateParam = addDateParam & "&fromDate="&xmlSelldate		'FROM��¥
	addDateParam = addDateParam & "&toDate="&xmlSelldate		'TO��¥
	addDateParam = addDateParam & "&counselListGb=01"			'�����ȸ���� | 00 : ��ü��ȸ, 01 : �亯�Ϸ���ȸ, 02 : �̴亯�Ϸ� ��ȸ(Default : 00)

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
							SabanetNum = dataList.get(i).counselSeq		'��������
							strSql =  ""
							strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
							strSql = strSql & " SET CS_STATUS = '003', TenStatus = 'C' "
							strSql = strSql & " WHERE SabanetNum = '"&SabanetNum&"' and MALL_ID = '"& MALL_ID &"'  "
							dbget.Execute strSql
							GetCSQnA_shintvshopping_complete = True
						Next
						rw "�ֹ�Q&A�����Ǽ�(Complete) : " & iInputCnt
						GetCSQnA_shintvshopping_complete = True
					Set dataList = nothing
				Else
					response.write "��������(0)<br />"

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
	strParam = strParam & "linkCode=" & linkCode					'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
	strParam = strParam & "&entpCode=" & entpCode					'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
	strParam = strParam & "&entpId=" & entpId						'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
	strParam = strParam & "&entpPass=" & entpPass					'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
	strParam = strParam & "&counselSeq=" & iNum						'#��������
	strParam = strParam & "&procNote=" & iRply						'#ó������

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
					rw iNum & " �亯�Ϸ�"
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
				End If
			Set strObj = nothing
		Else
			rw "req : " & strParam
			rw iNum & "���� : ����"
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
		MALL_ID = "�������(item)"
	ElseIf isellsite = "Mylittlewhoopee" Then
		MALL_ID = "�������Ĺ�ص�(item)"
	ElseIf isellsite = "nvstoregift" Then
		MALL_ID = "������ʼ����ϱ�(item)"
	Else
		MALL_ID = "������ʹ��汸(item)"
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
'	strRst = strRst & "			<Answered></Answered>"		'���� | �亯���� Y or N or ���
'	strRst = strRst & "			<Page></Page>"				'���� | ���� ��� �⺻������ 1page��ȸ
	strRst = strRst & "		</shop:GetQuestionAnswerListRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", xmlURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
		objXML.send(strRst)

		If objXML.Status <> "200" then
			response.write "ERROR : ��ſ���" & objXML.Status
			dbget.close : response.end
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
			If ResponseType <> "SUCCESS" Then
				rw "���� : ����"
				rw "message : " & xmlDOM.getElementsByTagName("n:Message")(0).Text
				rw "Detail : " & xmlDOM.getElementsByTagName("n:Detail")(0).Text
				Set xmlDOM = Nothing
				Set objXML = Nothing
				dbget.close : response.end
			Else
				If xmlDOM.getElementsByTagName("n:QuestionAnswerList").length < 1 Then
					response.write "��������(0)<br />"

					GetCSQnA_nvstorefarm = True
					Set xmlDOM = Nothing
					Set objXML = Nothing
					Exit Function
				Else
					masterCnt = xmlDOM.getElementsByTagName("n:QuestionAnswerList").length
					rw "�Ǽ�(" & masterCnt & ") "
					iInputCnt = 0
					set objMasterListXML = xmlDOM.getElementsByTagName("n:QuestionAnswerList")
						For Each objMasterOneXML In objMasterListXML
							SabanetNum = objMasterOneXML.getElementsByTagName("n:QuestionAnswerId").item(0).text	'��ǰ Q&A ID
							MALL_USER_ID = Trim(objMasterOneXML.getElementsByTagName("n:WriterId").item(0).text)	'����ŷ�� ������ ID

							Select Case Trim(objMasterOneXML.getElementsByTagName("n:QuestionType").item(0).text)	'PROD : ��ǰ, DLV : ���, RTN : ��ǰ, EXCG : ��ȯ, RFND : ȯ��, ETC : ��Ÿ
								Case "PROD"	QuestionType = "��ǰ"
								Case "DLV"	QuestionType = "���"
								Case "RTN"	QuestionType = "��ǰ"
								Case "EXCG"	QuestionType = "��ȯ"
								Case "RFND"	QuestionType = "ȯ��"
								Case "ETC"	QuestionType = "��Ÿ"
								Case Else	QuestionType = "��Ÿ"
							End Select

							CS_GUBUN = "DIRECT"		'c#���� �۾��� �Ϳܿ� ����API�̿��� ���� DIRECTó��
							questionPrefix ="��ǰQ&A"
							SUBJECT = "[" & questionPrefix & "_" & QuestionType & "] " & LEFT(Trim(objMasterOneXML.getElementsByTagName("n:Question").item(0).text), 100)
							CNTS = Trim(objMasterOneXML.getElementsByTagName("n:Question").item(0).text)			'���� ����
							REG_DM = objMasterOneXML.getElementsByTagName("n:CreateDate").item(0).text				'���� ����� YYYY-MM-DD ����

							If Trim(objMasterOneXML.getElementsByTagName("n:Answered").item(0).text) = "Y" Then
								CS_STATUS = "003"
							Else
								CS_STATUS = "001"
							End If

							PRODUCT_ID = Trim(objMasterOneXML.getElementsByTagName("n:ProductId").item(0).text)		'��ǰ ID
							ProductNM = Trim(objMasterOneXML.getElementsByTagName("n:ProductName").item(0).text)	'��ǰ��

							' rw objMasterOneXML.getElementsByTagName("n:QuestionAnswerId").item(0).text	'��ǰ Q&A ID
							' rw objMasterOneXML.getElementsByTagName("n:CreateDate").item(0).text			'���� ����� YYYY-MM-DD ����
							' rw objMasterOneXML.getElementsByTagName("n:QuestionType").item(0).text		'PROD : ��ǰ, DLV : ���, RTN : ��ǰ, EXCG : ��ȯ, RFND : ȯ��, ETC : ��Ÿ
							' rw objMasterOneXML.getElementsByTagName("n:Subject").item(0).text				'����
							' rw objMasterOneXML.getElementsByTagName("n:Question").item(0).text			'���� ����
							' rw objMasterOneXML.getElementsByTagName("n:Answer").item(0).text				'�亯 ����
							' rw objMasterOneXML.getElementsByTagName("n:Purchased").item(0).text			'������ YN
							' rw objMasterOneXML.getElementsByTagName("n:ProductId").item(0).text			'��ǰ ID
							' rw objMasterOneXML.getElementsByTagName("n:ProductName").item(0).text			'��ǰ��
							' rw objMasterOneXML.getElementsByTagName("n:WriterId").item(0).text			'����ŷ�� ������ ID
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
						rw "��ǰQ&A�Է¹׼����Ǽ� : " & iInputCnt
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
		MALL_ID = "�������(order)"
	ElseIf isellsite = "nvstoregift" Then
		MALL_ID = "������ʼ����ϱ�(order)"
	ElseIf isellsite = "Mylittlewhoopee" Then
		MALL_ID = "�������Ĺ�ص�(order)"
	Else
		MALL_ID = "������ʹ��汸(order)"
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
	strRst = strRst & "			<cus:DetailLevel>Full</cus:DetailLevel>"	'���� �޴� �������� �� ����(Compact/Full). �⺻���� "Full"�̴�.
	strRst = strRst & "			<cus:Version>1.0</cus:Version>"
	strRst = strRst & "			<ServiceType>SHOPN</ServiceType>"			'���� Ÿ�� �ڵ�. ����Ʈ����� �Ǹ��ڴ� SHOPN�� �Է��Ѵ�("A.3.2 ���� Ÿ�� �ڵ�" ����). ���̹����� �������� CHECKOUT �Է�
	strRst = strRst & "			<MallID>"&reqID&"</MallID>"					'�Ǹ��� ID
	strRst = strRst & "			<InquiryTimeFrom>"&selldate&"T00:00:00</InquiryTimeFrom>"		'��ȸ ���� �Ͻ�
	strRst = strRst & "			<InquiryTimeTo>"& Left(DateAdd("d", 1, CDate(selldate)), 10)&"T00:00:00</InquiryTimeTo>"			'��ȸ ���� �Ͻ�
'	strRst = strRst & "			<IsAnswered></IsAnswered>"					'�亯���� Y or N
	strRst = strRst & "		</cus:GetCustomerInquiryListRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", xmlURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
		objXML.send(strRst)

		If objXML.Status <> "200" then
			response.write "ERROR : ��ſ���" & objXML.Status
			dbget.close : response.end
		End If

		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
			If ResponseType <> "SUCCESS" Then
				rw "���� : ����"
				rw "message : " & xmlDOM.getElementsByTagName("n:Message")(0).Text
				rw "Detail : " & xmlDOM.getElementsByTagName("n:Detail")(0).Text
				Set xmlDOM = Nothing
				Set objXML = Nothing
				dbget.close : response.end
			Else
				If xmlDOM.getElementsByTagName("CustomerInquiry").length < 1 Then
					response.write "��������(0)<br />"

					GetCSOrderQnA_nvstorefarm = True
					Set xmlDOM = Nothing
					Set objXML = Nothing
					Exit Function
				Else
					masterCnt = xmlDOM.getElementsByTagName("CustomerInquiry").length
					rw "�Ǽ�(" & masterCnt & ") "
					iInputCnt = 0
					set objMasterListXML = xmlDOM.getElementsByTagName("CustomerInquiry")
						For Each objMasterOneXML In objMasterListXML
							SabanetNum = objMasterOneXML.getElementsByTagName("InquiryID").item(0).text				'���� ���� ��ȣ
							MALL_USER_ID = Trim(objMasterOneXML.getElementsByTagName("CustomerName").item(0).text)	'����
							QuestionType = objMasterOneXML.getElementsByTagName("Category").item(0).text			'���� ī�װ�
							CS_GUBUN = "DIRECT"		'c#���� �۾��� �Ϳܿ� ����API�̿��� ���� DIRECTó��
							questionPrefix ="�ֹ�Q&A"
							SUBJECT = "[" & questionPrefix & "_" & QuestionType & "] " & Trim(objMasterOneXML.getElementsByTagName("Title").item(0).text)	'���� ���� ����
							CNTS = Trim(objMasterOneXML.getElementsByTagName("InquiryContent").item(0).text)		'���� ����
							REG_DM = LEFT(objMasterOneXML.getElementsByTagName("InquiryDateTime").item(0).text, 10)

							If Trim(objMasterOneXML.getElementsByTagName("IsAnswered").item(0).text) = "true" Then		'�亯 ����
								CS_STATUS = "003"
							Else
								CS_STATUS = "001"
							End If
							PRODUCT_ID = Trim(objMasterOneXML.getElementsByTagName("ProductID").item(0).text)		'��ǰ ��ȣ
							ProductNM = Trim(objMasterOneXML.getElementsByTagName("ProductName").item(0).text)		'��ǰ��
'							OrderID = objMasterOneXML.getElementsByTagName("OrderID").item(0).text					'�ֹ���ȣ
							If Instr(objMasterOneXML.getElementsByTagName("ProductOrderID").item(0).text, ",") > 0 Then
								OrderID = Split(objMasterOneXML.getElementsByTagName("ProductOrderID").item(0).text, ",")(0)
							Else
								OrderID = objMasterOneXML.getElementsByTagName("ProductOrderID").item(0).text			'��ǰ �ֹ� ��ȣ
							End If

							' rw objMasterOneXML.getElementsByTagName("InquiryID").item(0).text				'���� ���� ��ȣ
							' rw objMasterOneXML.getElementsByTagName("OrderID").item(0).text				'�ֹ���ȣ
							' rw objMasterOneXML.getElementsByTagName("ProductOrderID").item(0).text		'��ǰ �ֹ� ��ȣ
							' rw objMasterOneXML.getElementsByTagName("ProductName").item(0).text			'��ǰ��
							' rw objMasterOneXML.getElementsByTagName("ProductID").item(0).text				'��ǰ ��ȣ
							' rw objMasterOneXML.getElementsByTagName("ProductOrderOption").item(0).text	'��ǰ �ɼ�
							' rw objMasterOneXML.getElementsByTagName("CustomerID").item(0).text			'��ID. �� �� �ڸ��� ��ǥ(*)�� ó���Ѵ�.
							' rw objMasterOneXML.getElementsByTagName("Title").item(0).text					'���� ���� ����
							' rw objMasterOneXML.getElementsByTagName("Category").item(0).text				'���� ī�װ�
							' rw objMasterOneXML.getElementsByTagName("InquiryDateTime").item(0).text		'���� �Ͻ�
							' rw objMasterOneXML.getElementsByTagName("InquiryContent").item(0).text		'���� ����
							' rw objMasterOneXML.getElementsByTagName("AnswerContent").item(0).text			'�亯 ����
							' rw objMasterOneXML.getElementsByTagName("IsAnswered").item(0).text			'�亯 ����
							' rw objMasterOneXML.getElementsByTagName("CustomerName").item(0).text			'����
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
						rw "�ֹ�Q&A�Է¹׼����Ǽ� : " & iInputCnt
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
					rw iNum & " �亯�Ϸ�(item)"
				Else
					rw iNum & "���� : ����"
					rw "message : " & xmlDOM.getElementsByTagName("n:Message")(0).Text
					rw "Detail : " & xmlDOM.getElementsByTagName("n:Detail")(0).Text
					Set xmlDOM = Nothing
					Set objXML = Nothing
'					dbget.close : response.end
				End If
			Set xmlDOM = nothing
	 	Else
	 		response.write "ERROR : ��ſ���" & objXML.Status
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
	strRst = strRst & "			<cus:DetailLevel>Full</cus:DetailLevel>"	'���� �޴� �������� �� ����(Compact/Full). �⺻���� "Full"�̴�.
	strRst = strRst & "			<cus:Version>1.0</cus:Version>"
	strRst = strRst & "			<MallID>"&reqID&"</MallID>"					'�Ǹ��� ID
	strRst = strRst & "			<ServiceType>SHOPN</ServiceType>"			'���� Ÿ�� �ڵ�. ����Ʈ����� �Ǹ��ڴ� SHOPN�� �Է��Ѵ�("A.3.2 ���� Ÿ�� �ڵ�" ����). ���̹����� �������� CHECKOUT �Է�
	strRst = strRst & "			<InquiryID>"&iNum&"</InquiryID>"			'���� ���� ��ȣ
	strRst = strRst & "			<AnswerContent><![CDATA["&iRply&"]]></AnswerContent>"	'�亯 ����
'	strRst = strRst & "			<AnswerContentID>?</AnswerContentID>"		'�亯 ��ȣ (�亯 �޽����� �����ϴ� ��쿡�� �ʼ�)
	strRst = strRst & "			<ActionType>INSERT</ActionType>"			'��ɾ�Ÿ���ڵ� (INSERT : �亯���, UPDATE : �亯���� )
'	strRst = strRst & "			<AnswerTempleteID>?</AnswerTempleteID>"		'���� �亯 ���ø� �Ϸù�ȣ
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
					rw iNum & " �亯�Ϸ�(order)"
				Else
					rw iNum & "���� : ����"
					rw "message : " & xmlDOM.getElementsByTagName("n:Message")(0).Text
					rw "Detail : " & xmlDOM.getElementsByTagName("n:Detail")(0).Text
					Set xmlDOM = Nothing
					Set objXML = Nothing
'					dbget.close : response.end
				End If
			Set xmlDOM = nothing
	 	Else
	 		response.write "ERROR : ��ſ���" & objXML.Status
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
				obj("ansInfo")(null)("trGrpCd") = "SR"			'#�ŷ�ó�׷��ڵ�
				obj("ansInfo")(null)("trNo") = trCd				'#�ŷ�ó��ȣ
				obj("ansInfo")(null)("lrtrNo") = ""				'�����ŷ�ó��ȣ
				obj("ansInfo")(null)("pdQnaNo") = iNum			'#��ǰQnA��ȣ
				obj("ansInfo")(null)("spdNo") = spdNo			'#�Ǹ��ڻ�ǰ�ڵ�
				obj("ansInfo")(null)("sitmNo") = sitmNo			'#�Ǹ��ڴ�ǰ�ڵ�
				obj("ansInfo")(null)("ansCnts") = iRply			'#�亯����
				jParam = obj.jsString
	Set obj = nothing

	'// ����Ÿ ��������
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(jParam)
		If objXML.Status <> "200" Then
			response.write "ERROR : ��ſ���" & objXML.Status
			response.write "<script>alert('ERROR : ��ſ���.');</script>"
			dbget.close : response.end
		Else
			objData = BinaryToText(objXML.ResponseBody,"utf-8")
		End If

	'// =======================================================================
	'// Json �Ľ�
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
			rw iNum & " �亯�Ϸ�(item)"
		Else
			rw iNum & "���� : ����"
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
		obj("slrInqNo") = iNum			'#�Ǹ��ڹ��ǹ�ȣ
		obj("ansCnts") = iRply			'#�亯����
		jParam = obj.jsString
	Set obj = nothing

	'// ����Ÿ ��������
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(jParam)

		If objXML.Status <> "200" Then
			response.write "ERROR : ��ſ���" & objXML.Status
			response.write "<script>alert('ERROR : ��ſ���.');</script>"
			dbget.close : response.end
		Else
			objData = BinaryToText(objXML.ResponseBody,"utf-8")
		End If
	'// =======================================================================
	'// Json �Ľ�
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
			rw iNum & " �亯�Ϸ�(seller)"
		Else
			rw iNum & "���� : ����"
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
			iaccessLicense = "01000100004b035a25d67f991849cad1c7042b8da528d13e9ddce6878f2e43ac88080e0a5e" 'AccessLicense Key �Է�, PDF��������
			osecretKey = "AQABAAAWPWagCrPjFQnFEtxs5j+oyZFwuzomdNq0XZSricPuMw=="  'SecreKey �Է�, PDF��������
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		Else
			iaccessLicense = "010001000019133c715650b9c85b820961612f2b90b431ddd8654b42c097c4df1a43d0be09" 'AccessLicense Key �Է�, PDF��������
			osecretKey = "AQABAADX6Hz/wORFJS5pSIy4KQXkH83gC9G1aXChxBjcnUMqWw=="  'SecreKey �Է�, PDF��������
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
		'lotteAPIURL = "http://openapidev.lotte.com"	'' �׽�Ʈ����
		lotteAPIURL = "http://openapitest.lotte.com"	'' �׽�Ʈ����
		tenBrandCd = "14846"	'�ٹ�(�ӽ�)
		tenDlvCd = "513484"		'�����å�ڵ�
		CertPasswd = "1234"		'Dev�� ��� : 1234
	Else
		lotteAPIURL = "https://openapi.lotte.com"		'' �Ǽ���
		tenBrandCd = "155112"	'�ٹ�����
		tenDlvCd = "513484"
		CertPasswd = "store101010*"
	End if
	lottenTenID = "124072"					'�ٹ�����ID

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
			'XML�� ���� DOM ��ü ����
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

			''on Error Resume Next
				GetLotteAuthNo = xmlDOM.getElementsByTagName("SubscriptionId").item(0).text
				if Err<>0 then
					Response.Write "��ſ���(XML)"
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
			Response.Write "��ſ���"
			Response.End
		end if
		Set objXML = Nothing
	Else
		GetLotteAuthNo = dbAuthNo
	End If
end function

Function GetLotteimallAuthNo()
	GetLotteimallAuthNo = ""
	'// �Ե����̸�API �������� URL
	Dim ltiMallAPIURL, ltiMallAuthNo, ltiMallTenID, tenBrandCd, tenDlvCd, tenDlvFreeCd
	Dim iisql
	IF application("Svr_Info") = "Dev" THEN
		'ltiMallAPIURL = "http://openapidev.lotteimall.com"	'' �׽�Ʈ����
		ltiMallAPIURL = "http://openapitst.lotteimall.com"	'' �׽�Ʈ����
		tenDlvCd = "23725"
		tenDlvFreeCd = "577045"
	Else
		ltiMallAPIURL = "https://openapi.lotteimall.com"		'' �Ǽ���
		tenDlvCd = "23725"
		tenDlvFreeCd = "577045"
	End if
	ltiMallTenID = "011799LT"


	'// �Ե����̸� �����ڵ� Ȯ��(���� ������Ʈ; ���ø����̼Ǻ����� ����)
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
					GetLotteimallAuthNo = xmlDOM.getElementsByTagName("SubscriptionId").item(0).text		'������ȣ ����
					if Err<>0 then
						Response.Write "��ſ���(XML)"
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
				Response.Write "��ſ���"
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
			strSql = strSql & " and MALL_ID = '�������(item)' "
		Else
			strSql = strSql & " and MALL_ID = '�������(order)' "
		End If
	ElseIf imallid = "nvstoregift" Then
		If icsGubun = "item" Then
			strSql = strSql & " and MALL_ID = '������ʼ����ϱ�(item)' "
		Else
			strSql = strSql & " and MALL_ID = '������ʼ����ϱ�(order)' "
		End If
	ElseIf imallid = "Mylittlewhoopee" Then
		If icsGubun = "item" Then
			strSql = strSql & " and MALL_ID = '�������Ĺ�ص�(item)' "
		Else
			strSql = strSql & " and MALL_ID = '�������Ĺ�ص�(order)' "
		End If
	Else
		If icsGubun = "item" Then
			strSql = strSql & " and MALL_ID = '������ʹ��汸(item)' "
		Else
			strSql = strSql & " and MALL_ID = '������ʹ��汸(order)' "
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
		strSql = strSql & " and MALL_ID = '�Ե�On(item)' "
	Else
		strSql = strSql & " and MALL_ID = '�Ե�On(seller)' "
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
	strSql = strSql & " and MALL_ID = '�ż���TV����' "
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
