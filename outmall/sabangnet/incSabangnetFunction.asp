<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'상품 등록
Public Function fnSabangnetItemReg(iitemid, istrParam, byRef iErrStr, imustprice, iimageNm, ilimityn, ilimitno, ilimitsold)
	Dim dataURL, objXML, xmlDOM, strRst, iRbody, iMessage, maySabangnetGoodno, tmpGoodNo
	Dim fso,tFile, tenOptcd
	Dim Toptionname, Tlimitno, Tlimitsold, Tlimityn, Titemsu
	Dim opath : opath = "/outmall/sabangnet/sabangnetXML/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	Dim fileName : fileName = "REG" &"_"& getCurrDateTimeFormat&".xml"
	CALL CheckFolderCreate(defaultPath)
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & FileName )
			tFile.WriteLine istrParam
		Set tFile = nothing
	Set fso = nothing

	dataURL = "?xml_url="&wapiURL&opath&FileName

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & sabangnetAPIURL&"/RTL_API/xml_goods_info.html" & dataURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
'				response.write iRbody
				iMessage = Split(iRbody,"[1] ")(1)
				iMessage = GetFirstMatch("<font[^>]*>([^<]+)</font>",iMessage)

				'API 호출시 응답값이 XML이 아닌 string으로 넘어옴; 전화 문의 해보니 원래 그렇다는데...
				If InStr(iMessage,"성공 : ") > 0 Then
					'성공 : 100001 [1814737]
					tmpGoodNo = Split(iMessage, "성공 : ")(1)			'100001 [1814737]
					If InStr(tmpGoodNo,"[") > 0 Then
						maySabangnetGoodno = Trim(Split(tmpGoodNo, "[")(0))
					End If

					If maySabangnetGoodno <> "" Then
						strSql = ""
						strSql = strSql & " UPDATE R"& VbCrlf
						strSql = strSql & "	Set sabangnetLastUpdate = getdate() "& VbCrlf
						strSql = strSql & "	, sabangnetGoodNo = '" & maySabangnetGoodno & "'"& VbCrlf
						strSql = strSql & "	, sabangnetPrice = " &imustprice & VbCrlf
						strSql = strSql & "	, regImageName = '"&iimageNm&"' " & VbCrlf
						strSql = strSql & "	, accFailCnt = 0" & VbCrlf
						strSql = strSql & "	, sabangnetRegdate = isNULL(sabangnetRegdate, getdate())" & VbCrlf
					    strSql = strSql & "	, sabangnetStatCd=(CASE WHEN isNULL(sabangnetStatCd, -1) < 7 then 7 ELSE sabangnetStatCd END)" & VbCrlf
					    strSql = strSql & " , sabangnetSellYn = 'Y' "& VbCrlf
						strSql = strSql & "	From db_etcmall.dbo.tbl_sabangnet_regItem  R"& VbCrlf
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)

						strSql = ""
						strSql = strSql & " SELECT COUNT(*) FROM db_item.dbo.tbl_item_option WHERE itemid = '"&iitemid&"' "
						rsget.CursorLocation = adUseClient
						rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
						If rsget(0) = 0 Then
							tenOptcd	= "0000"
						End If
						rsget.Close

						If tenOptcd = "0000"  Then	'단일 상품이라면
							Toptionname		= "단일상품"
							Tlimitno		= ilimitno
							Tlimitsold		= ilimitsold
							Tlimityn		= ilimityn
							If (Tlimityn="Y") then
								If (Tlimitno - Tlimitsold - 5) < 1 Then
									Titemsu = 0
								Else
									Titemsu = Tlimitno - Tlimitsold - 5
								End If
							Else
								Titemsu = 999
							End If
							strSql = ""
							strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
							strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
							strSql = strSql & " VALUES " & VBCRLF
							strSql = strSql & " ('"&iitemid&"',  '"&tenOptcd&"', 'sabangnet', '', '"&html2db(Toptionname)&"', 'Y', '"&Tlimityn&"', '"&Titemsu&"', '0', getdate()) "
							dbget.Execute strSql
						Else
							strSql = ""
							strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
							strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
							strSql = strSql & " SELECT itemid, itemoption, 'sabangnet', '', optionname "
							strSql = strSql & " ,Case WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold <= 5) THEN 'N' " & VBCRLF
							strSql = strSql & " 	 WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold > 5) THEN optsellyn " & VBCRLF
							strSql = strSql & "	Else optsellyn End, optlimityn, " & VBCRLF
							strSql = strSql & " Case WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold <= 5) THEN '0' " & VBCRLF
							strSql = strSql & " 	 WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold > 5) THEN optlimitno - optlimitsold - 5 " & VBCRLF
							strSql = strSql & " 	 WHEN (optlimityn = 'N') THEN '999' End " & VBCRLF
							strSql = strSql & " , optaddprice, getdate() " & VBCRLF
							strSql = strSql & " FROM db_item.dbo.tbl_item_option " & VBCRLF
							strSql = strSql & " WHERE isUsing='Y' and optsellyn='Y' and itemid= '"&iitemid&"' " & VBCRLF
							dbget.Execute strSql
						End If
						strSql = ""
						strSql = strSql & " UPDATE R " & VBCRLF
						strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VBCRLF
						strSql = strSql & " FROM db_etcmall.dbo.tbl_sabangnet_regItem R " & VBCRLF
						strSql = strSql & " JOIN ( " & VBCRLF
						strSql = strSql & " 	SELECT R.itemid, count(*) as CNT, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt " & VBCRLF
						strSql = strSql & " 	FROM db_etcmall.dbo.tbl_sabangnet_regItem R " & VBCRLF
						strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = 'sabangnet' and Ro.itemid = " &iitemid & VBCRLF
						strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
						strSql = strSql & " ) T on R.itemid = T.itemid " & VBCRLF
						dbget.Execute strSql
						iErrStr =  "OK||"&iitemid&"||[등록]성공"
					Else
						iErrStr = "ERR||"&iitemid&"||[등록] "& iMessage
					End If
				Else
					iErrStr = "ERR||"&iitemid&"||[등록]"&iMessage
				End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "ERR||사방넷 결과 분석 중에 오류가 발생했습니다.[ERR-REG-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
	Call DelAPITMPFile(wapiURL&opath&FileName)
End Function

'상품 가격 상태 요약 수정
Public Function fnSabangnetSimpleEdit(iitemid, ichgSellyn, imustprice, iitemname, istrParam, byRef iErrStr, gubun)
	Dim dataURL, objXML, xmlDOM, strRst, iRbody, iMessage, maySabangnetGoodno, tmpGoodNo
	Dim fso,tFile, gubunStr
	Dim opath : opath = "/outmall/sabangnet/sabangnetXML/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	Dim fileName : fileName = "sEDIT" &"_"& getCurrDateTimeFormat&".xml"
	CALL CheckFolderCreate(defaultPath)
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & FileName )
			tFile.WriteLine istrParam
		Set tFile = nothing
	Set fso = nothing

	dataURL = "?xml_url="&wapiURL&opath&FileName

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & sabangnetAPIURL&"/RTL_API/xml_goods_info2.html" & dataURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
'				response.write iRbody
				iMessage = Split(iRbody,"[1] ")(1)
				iMessage = GetFirstMatch("<font[^>]*>([^<]+)</font>",iMessage)

				If gubun = "sellyn" Then
					gubunStr = "요약수정상태"
				Else
					gubunStr = "요약수정가격"
				End If

				'API 호출시 응답값이 XML이 아닌 string으로 넘어옴; 전화 문의 해보니 원래 그렇다는데...
				If InStr(iMessage,"성공 : ") > 0 Then
					If ichgSellyn = "Y" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET sabangnetPrice = '"&imustPrice&"' "
						strSql = strSql & "	,sabangnetSellYn = 'Y'"
						strSql = strSql & "	,regitemname = '"&iitemname&"'"
						strSql = strSql & "	,sabangnetLastUpdate = getdate()"
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_sabangnet_regItem R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||["&gubunStr&"]판매"
					Else
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET sabangnetPrice = '"&imustPrice&"' "
						strSql = strSql & "	,sabangnetSellYn = 'N'"
						strSql = strSql & "	,regitemname = '"&iitemname&"'"
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,sabangnetLastUpdate = getdate()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_sabangnet_regItem R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||["&gubunStr&"]품절"
					End If
				Else
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"("&gubunStr&")"
				End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "ERR||사방넷 결과 분석 중에 오류가 발생했습니다.[ERR-sEDIT-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
	Call DelAPITMPFile(wapiURL&opath&FileName)
End Function

'상품 쇼핑몰별 DATA 수정
Public Function fnShoppingDataSabangnet(iitemid, istrParam, byRef iErrStr)
	Dim dataURL, objXML, xmlDOM, strRst, iRbody, iMessage, maySabangnetGoodno, tmpGoodNo
	Dim fso,tFile, gubunStr
	Dim opath : opath = "/outmall/sabangnet/sabangnetXML/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	Dim fileName : fileName = "sDATA" &"_"& getCurrDateTimeFormat&".xml"
	CALL CheckFolderCreate(defaultPath)
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & FileName )
			tFile.WriteLine istrParam
		Set tFile = nothing
	Set fso = nothing

	dataURL = "?xml_url="&wapiURL&opath&FileName

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & sabangnetAPIURL&"/RTL_API/xml_goods_info3.html" & dataURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
rw "---------!!"
response.write  iRbody
response.end
rw "---------@@"
			Set xmlDOM = Nothing
		Else
			iErrStr = "ERR||사방넷 결과 분석 중에 오류가 발생했습니다.[ERR-sEDIT-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
	Call DelAPITMPFile(wapiURL&opath&FileName)
End Function

'상품 기본정보 수정
Public Function fnSabangnetItemEdit(iitemid, istrParam, byRef iErrStr, imustprice, iimageNm, ilimityn, ilimitno, ilimitsold, ichgSellyn)
	Dim dataURL, objXML, xmlDOM, strRst, iRbody, iMessage, maySabangnetGoodno, tmpGoodNo
	Dim fso,tFile, tenOptcd
	Dim Toptionname, Tlimitno, Tlimitsold, Tlimityn, Titemsu
	Dim opath : opath = "/outmall/sabangnet/sabangnetXML/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	Dim fileName : fileName = "EDIT" &"_"& getCurrDateTimeFormat&".xml"
	CALL CheckFolderCreate(defaultPath)
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & FileName )
			tFile.WriteLine istrParam
		Set tFile = nothing
	Set fso = nothing

	dataURL = "?xml_url="&wapiURL&opath&FileName

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & sabangnetAPIURL&"/RTL_API/xml_goods_info.html" & dataURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
'				response.write iRbody
				iMessage = Split(iRbody,"[1] ")(1)
				iMessage = GetFirstMatch("<font[^>]*>([^<]+)</font>",iMessage)

				'API 호출시 응답값이 XML이 아닌 string으로 넘어옴; 전화 문의 해보니 원래 그렇다는데...
				If InStr(iMessage,"성공 : ") > 0 Then
					strSql = ""
					strSql = strSql & " UPDATE R"
					strSql = strSql & "	SET sabangnetPrice = '"&imustPrice&"' "
					strSql = strSql & "	, regImageName = '"&iimageNm&"' " & VbCrlf
					strSql = strSql & "	,sabangnetSellYn = '"&ichgSellyn&"'"
					strSql = strSql & "	,regitemname = '"&iitemname&"'"
					strSql = strSql & "	,sabangnetLastUpdate = getdate()"
					strSql = strSql & "	,accFailCnt = 0"
					strSql = strSql & "	FROM db_etcmall.dbo.tbl_sabangnet_regItem R"
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)

					strSql = ""
					strSql = strSql & " DELETE FROM db_item.dbo.tbl_OutMall_regedoption WHERE itemid = '"&iitemid&"' and mallid = 'sabangnet' "
					dbget.Execute strSql

					strSql = ""
					strSql = strSql & " SELECT COUNT(*) FROM db_item.dbo.tbl_item_option WHERE itemid = '"&iitemid&"' "
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If rsget(0) = 0 Then
						tenOptcd	= "0000"
					End If
					rsget.Close

					If tenOptcd = "0000"  Then	'단일 상품이라면
						Toptionname		= "단일상품"
						Tlimitno		= ilimitno
						Tlimitsold		= ilimitsold
						Tlimityn		= ilimityn
						If (Tlimityn="Y") then
							If (Tlimitno - Tlimitsold - 5) < 1 Then
								Titemsu = 0
							Else
								Titemsu = Tlimitno - Tlimitsold - 5
							End If
						Else
							Titemsu = 999
						End If
						strSql = ""
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
						strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						strSql = strSql & " VALUES " & VBCRLF
						strSql = strSql & " ('"&iitemid&"',  '"&tenOptcd&"', 'sabangnet', '', '"&html2db(Toptionname)&"', 'Y', '"&Tlimityn&"', '"&Titemsu&"', '0', getdate()) "
						dbget.Execute strSql
					Else
						strSql = ""
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
						strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						strSql = strSql & " SELECT itemid, itemoption, 'sabangnet', '', optionname "
						strSql = strSql & " ,Case WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold <= 5) THEN 'N' " & VBCRLF
						strSql = strSql & " 	 WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold > 5) THEN optsellyn " & VBCRLF
						strSql = strSql & "	Else optsellyn End, optlimityn, " & VBCRLF
						strSql = strSql & " Case WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold <= 5) THEN '0' " & VBCRLF
						strSql = strSql & " 	 WHEN (optlimityn = 'Y') AND (optlimitno - optlimitsold > 5) THEN optlimitno - optlimitsold - 5 " & VBCRLF
						strSql = strSql & " 	 WHEN (optlimityn = 'N') THEN '999' End " & VBCRLF
						strSql = strSql & " , optaddprice, getdate() " & VBCRLF
						strSql = strSql & " FROM db_item.dbo.tbl_item_option " & VBCRLF
						strSql = strSql & " WHERE isUsing='Y' and optsellyn='Y' and itemid= '"&iitemid&"' " & VBCRLF
						dbget.Execute strSql
					End If
					strSql = ""
					strSql = strSql & " UPDATE R " & VBCRLF
					strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VBCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_sabangnet_regItem R " & VBCRLF
					strSql = strSql & " JOIN ( " & VBCRLF
					strSql = strSql & " 	SELECT R.itemid, count(*) as CNT, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt " & VBCRLF
					strSql = strSql & " 	FROM db_etcmall.dbo.tbl_sabangnet_regItem R " & VBCRLF
					strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = 'sabangnet' and Ro.itemid = " &iitemid & VBCRLF
					strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid " & VBCRLF
					dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||[전체수정]성공"
				Else
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(전체수정)"
				End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "ERR||사방넷 결과 분석 중에 오류가 발생했습니다.[ERR-EDIT-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
	Call DelAPITMPFile(wapiURL&opath&FileName)
End Function

'카테고리 정보 얻기
Public Function fnRegSabangnetCategory(istrParam)
	Dim dataURL, objXML, xmlDOM, strRst, iRbody
	Dim fso,tFile
	Dim opath : opath = "/outmall/sabangnet/sabangnetXML/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	Dim fileName : fileName = "CATE" &"_"& getCurrDateTimeFormat&".xml"
	CALL CheckFolderCreate(defaultPath)
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & FileName )
			tFile.WriteLine istrParam
		Set tFile = nothing
	Set fso = nothing

	dataURL = "?xml_url="&wapiURL&opath&FileName

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & sabangnetAPIURL&"/RTL_API/xml_category_info2.html" & dataURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				'response.write iRbody
			Set xmlDOM = Nothing
		Else
			iErrStr = "ERR||사방넷 결과 분석 중에 오류가 발생했습니다.[ERR-CATE-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
	Call DelAPITMPFile(wapiURL&opath&FileName)
End Function

'상품정보고시 정보 얻기
Public Function fnGosiInfoSabangnet(vReqDiv)
	Dim dataURL, objXML, xmlDOM, strRst, iRbody, istrParam
	Dim fso,tFile
	Dim opath : opath = "/outmall/sabangnet/sabangnetXML/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	Dim fileName : fileName = "GosiInfo" &"_"& getCurrDateTimeFormat&".xml"
	CALL CheckFolderCreate(defaultPath)
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & FileName )
			istrParam = ""
			istrParam = ""
			istrParam = istrParam & "<?xml version=""1.0"" encoding=""EUC-KR""?>"
			istrParam = istrParam & "<SABANG_GOODS_PROP_CODE_INFO_LIST>"
			istrParam = istrParam & "	<HEADER>"
			istrParam = istrParam & "		<SEND_COMPAYNY_ID>"&sabangnetID&"</SEND_COMPAYNY_ID>"
			istrParam = istrParam & "		<SEND_AUTH_KEY>"&sabangnetAPIKEY&"</SEND_AUTH_KEY>"
			istrParam = istrParam & "		<SEND_DATE>"&Replace(Date(), "-", "")&"</SEND_DATE>"
			istrParam = istrParam & "	</HEADER>"
			istrParam = istrParam & "	<DATA>"
			istrParam = istrParam & "		<PROP1_CD>"&vReqDiv&"</PROP1_CD>"
			istrParam = istrParam & "	</DATA>"
			istrParam = istrParam & "</SABANG_GOODS_PROP_CODE_INFO_LIST>"
			tFile.WriteLine istrParam
		Set tFile = nothing
	Set fso = nothing

	dataURL = "?xml_url="&wapiURL&opath&FileName

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & sabangnetAPIURL&"/RTL_API/xml_goods_prop_code_info.html" & dataURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody

				response.write iRbody
			Set xmlDOM = Nothing
		Else
			iErrStr = "ERR||사방넷 결과 분석 중에 오류가 발생했습니다.[ERR-GOSI-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
	Call DelAPITMPFile(wapiURL&opath&FileName)
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################

'################################################# 각 기능 별 파라메터 정리시작 ###############################################
'관리카테고리 XML
Public Function get10x10CategoryParameter()
	Dim strSQL, strRst1, strRst2, strRst3
	Dim cdlarge, cdmid, cdsmall, nmlarge, nmmid, nmsmall
	strRst1 = ""
	strRst1 = strRst1 & "<?xml version=""1.0"" encoding=""EUC-KR""?>"
	strRst1 = strRst1 & "<SABANG_CATEGORY_REGI>"
	strRst1 = strRst1 & "	<HEADER>"
	strRst1 = strRst1 & "		<SEND_COMPAYNY_ID>"&sabangnetID&"</SEND_COMPAYNY_ID>"
	strRst1 = strRst1 & "		<SEND_AUTH_KEY>"&sabangnetAPIKEY&"</SEND_AUTH_KEY>"
	strRst1 = strRst1 & "		<SEND_DATE>"&Replace(Date(), "-", "")&"</SEND_DATE>"
	strRst1 = strRst1 & "	</HEADER>"
	'--------------------------------  쿼리부분 시작 --------------------------------
	strSQL = ""
	strSQL = strSQL & " SELECT l.code_large AS cdlarge, m.code_mid AS cdmid, s.code_small AS cdsmall, "
	strSQL = strSQL & " l.code_nm AS nmlarge, m.code_nm AS nmmid, s.code_nm AS nmsmall "
	strSQL = strSQL & " FROM [dbo].tbl_Cate_large l  "
	strSQL = strSQL & " INNER JOIN [dbo].tbl_Cate_mid m ON l.code_large = m.code_large and m.display_yn = 'Y' "
	strSQL = strSQL & " INNER JOIN [dbo].tbl_Cate_small s ON l.code_large = s.code_large AND m.code_mid = s.code_mid and s.display_yn = 'Y' "
	strSQL = strSQL & " where l.display_yn = 'Y' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) then
		Do until rsget.EOF
		    cdlarge	= rsget("cdlarge")
		    cdmid	= rsget("cdmid")
		    cdsmall	= rsget("cdsmall")
		    nmlarge	= rsget("nmlarge")
		    nmmid	= rsget("nmmid")
		    nmsmall	= rsget("nmsmall")

			strRst2 = strRst2 & "	<DATA>"
			strRst2 = strRst2 & "		<CLASS_CD1>"
			strRst2 = strRst2 & "			<CODE>"&cdlarge&"</CODE>"
			strRst2 = strRst2 & "			<CATEGORY><![CDATA["&nmlarge&"]]></CATEGORY>"
			strRst2 = strRst2 & "			<USE_YN>Y</USE_YN>"
			strRst2 = strRst2 & "		</CLASS_CD1>"
			strRst2 = strRst2 & "		<CLASS_CD2>"
			strRst2 = strRst2 & "			<CODE>"&cdmid&"</CODE>"
			strRst2 = strRst2 & "			<CATEGORY><![CDATA["&nmmid&"]]></CATEGORY>"
			strRst2 = strRst2 & "			<USE_YN>Y</USE_YN>"
			strRst2 = strRst2 & "		</CLASS_CD2>"
			strRst2 = strRst2 & "		<CLASS_CD3>"
			strRst2 = strRst2 & "			<CODE>"&cdsmall&"</CODE>"
			strRst2 = strRst2 & "			<CATEGORY><![CDATA["&nmsmall&"]]></CATEGORY>"
			strRst2 = strRst2 & "			<USE_YN>Y</USE_YN>"
			strRst2 = strRst2 & "		</CLASS_CD3>"
			strRst2 = strRst2 & "	</DATA>"
			rsget.MoveNext
		Loop
	End If
	rsget.Close
	'--------------------------------  쿼리부분 끝 --------------------------------
	strRst3 = ""
	strRst3 = strRst3 & "</SABANG_CATEGORY_REGI>"
	get10x10CategoryParameter = strRst1 & strRst2 & strRst3
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

Function GetFirstMatch(PatternToMatch, StringToSearch)
	Dim regEx, CurrentMatch, CurrentMatches

	Set regEx = New RegExp
	regEx.Pattern = PatternToMatch
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.MultiLine = True
	Set CurrentMatches = regEx.Execute(StringToSearch)

	GetFirstMatch = ""
	If CurrentMatches.Count >= 1 Then
		Set CurrentMatch = CurrentMatches(0)
		If CurrentMatch.SubMatches.Count >= 1 Then
			GetFirstMatch = CurrentMatch.SubMatches(0)
		End If
	End If
	Set regEx = Nothing
End Function
'################################################# 각 기능 별 파라메터 정리 끝 ###############################################
%>