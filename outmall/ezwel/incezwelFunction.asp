<%
Public ezwelAPIURL
Public postParam
Public ezwelNewAPIURL

Const dataHead = "<?xml version=""1.0"" encoding=""euc-kr"" standalone=""yes""?>"
IF application("Svr_Info") = "Dev" THEN
	ezwelAPIURL = "http://api.dev.ezwel.com/if/api/goodsInfoAPI.ez"
	ezwelNewAPIURL = "https://openapi.ezwel.com"
	'ezwelNewAPIURL = "https://apitest.ezwel.com"
Else
	ezwelAPIURL = "http://api.ezwel.com/if/api/goodsInfoAPI.ez"
	ezwelNewAPIURL = "https://openapi.ezwel.com"
End if
postParam	= "cspCd="&cspCd&"&crtCd="&crtCd&"&dataSet="
'############################################## 실제 수행하는 API 함수 모음 ##############################################
''이지웰 상품 등록
Function EzwelItemReg(iitemid, strParam, byRef iErrStr, iSellCash, iezwelSellYn, ilimityn, ilimitno, ilimitsold, iitemname, iimageNm)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, strSql, tenOptCnt
	Dim retCode, goodsCd, iMessage, AssignedRow, Toptionname, Tlimitno, Tlimitsold, Tlimityn
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
		objXML.send(postParam&xmlStr)
'response.write xmlStr
'response.end
'response.write objXML.ResponseText
	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

'			response.write objXML.ResponseText
'			response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
			'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If

		goodsCd		= xmlDOM.getElementsByTagName("goodsCd").item(0).text
		retCode		= xmlDOM.getElementsByTagName("resultCode").item(0).text
		iMessage	= xmlDOM.getElementsByTagName("resultMsg").item(0).text

		If retCode = "200" Then		'성공(200)
			strSql = "SELECT COUNT(itemid) FROM db_etcmall.dbo.tbl_ezwel_regItem WHERE itemid='" & iitemid & "' and ezwelgoodno = '"&goodsCd&"'"
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If rsget(0) = 0 Then
				strSql = ""
				strSql = strSql & " UPDATE R" & VbCRLF
				strSql = strSql & "	Set ezwelLastUpdate = getdate() "  & VbCRLF
				strSql = strSql & "	, ezwelGoodNo = '" & goodsCd & "'"  & VbCRLF
				strSql = strSql & "	, ezwelPrice = " &iSellCash& VbCRLF
				strSql = strSql & "	, accFailCnt = 0"& VbCRLF
				strSql = strSql & "	, ezwelRegdate = isNULL(ezwelRegdate, getdate())"
				If (goodsCd <> "") Then
				    strSql = strSql & "	, ezwelstatCD = '3'"& VbCRLF					'등록완료(임시)
				Else
					strSql = strSql & "	, ezwelstatCD = '1'"& VbCRLF					'전송시도
				End If
				strSql = strSql & "	From db_etcmall.dbo.tbl_ezwel_regItem R"& VbCRLF
				strSql = strSql & " Where R.itemid = '" & iitemid & "'"
				dbget.Execute(strSql)
			Else
				'// 없음 -> 신규등록
				strSql = ""
				strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_ezwel_regItem "
				strSql = strSql & " (itemid, regitemname, reguserid, ezwelRegdate, ezwelLastUpdate, ezwelGoodNo, ezwelPrice, ezwelSellYn, ezwelStatCd, regImageName) VALUES " & VbCRLF
				strSql = strSql & " ('" & iitemid & "'" & VBCRLF
				strSql = strSql & " , '" & iitemname & "'" &_
				strSql = strSql & " , '" & session("ssBctId") & "'" &_
				strSql = strSql & " , getdate(), getdate()" & VBCRLF
				strSql = strSql & " , '" & goodsCd & "'" & VBCRLF
				strSql = strSql & " , '" & iSellCash & "'" & VBCRLF
				strSql = strSql & " , '" & iezwelSellYn & "'" & VBCRLF
				If (goodsCd <> "") Then
				    strSql = strSql & ",'3'"											'등록완료(임시)
				Else
				    strSql = strSql & ",'1'"											'전송시도
				End If
				strSql = strSql & " , '" & iimageNm & "'" & VBCRLF
				strSql = strSql & ")"
				dbget.Execute(strSql)
			End If
			rsget.Close

			strSql = ""
			strSql = strSql &  "SELECT count(*) as cnt "
			strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
			strSql = strSql & " WHERE itemid=" & iitemid
			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				tenOptCnt = rsget("cnt")
			rsget.Close

			If tenOptCnt = 0 Then
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
				strSql = strSql & " ('"&iitemid&"',  '0000', 'ezwel', '', '"&html2db(Toptionname)&"', 'Y', '"&Tlimityn&"', '"&Titemsu&"', '0', getdate()) "
				dbget.Execute strSql
			Else
				strSql = ""
				strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
				strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
				strSql = strSql & " SELECT itemid, itemoption, 'ezwel', '', optionname "
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
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ezwel_regItem SET "
			strSql = strSql & " regedOptCnt = " & tenOptCnt
			strSql = strSql & " WHERE itemid = " & iitemid
			dbget.Execute strSql
			EzwelOneItemReg = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			iErrStr =  "OK||"&iitemid&"||등록성공(상품등록)"
		Else						'실패(E)
		    iErrStr = "ERR||"&iitemid&"||ezwel 결과 분석 중에 오류가 발생했습니다.[ERR-REG-001]"
		End If
		On Error Goto 0
	Else
		iErrStr = "ERR||"&iitemid&"||ezwel 통신 오류가 발생했습니다.[ERR-REG-002]"
	End If
End Function

Function EzwelOneItemEdit(iitemid, iEzwelGoodNo, byRef iErrStr, strParam, imustprice, ichgSellYn, optMust, ilimityn, ilimitno, ilimitsold, ichkXML, iezwelsellyn)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, strSql, tenOptCnt
	Dim retCode, goodsCd, iMessage, AssignedRow, oMsg, ocount, Toptionname, Tlimitno, Tlimitsold, Tlimityn, Titemsu, mayReSell
	On Error Resume Next

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
		objXML.send(postParam&xmlStr)

		If (session("ssBctID")="kjy8517") Then
			response.write "수정Resquest : "
			response.write "<textarea cols=80 rows=2>"&xmlStr&"</textarea><br />"
		End If

		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				If (session("ssBctID")="kjy8517") Then
					response.write "수정Response : "
					response.write "<textarea cols=80 rows=2>"&BinaryToText(objXML.ResponseBody, "euc-kr")&"</textarea><br />"
				End If

				goodsCd		= xmlDOM.getElementsByTagName("goodsCd").item(0).text
				retCode		= xmlDOM.getElementsByTagName("resultCode").item(0).text
				iMessage	= xmlDOM.getElementsByTagName("resultMsg").item(0).text

				If retCode = "200" Then		'성공(200)
					mayReSell = ""
					If iezwelsellyn = "N" and ichgSellYn = "Y" Then
						mayReSell = "Y"
					End If

					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ezwel_regItem " & VbCRLF
					strSql = strSql & " SET accFailCnt = 0 " & VbCRLF
					strSql = strSql & " ,ezwelLastUpdate = getdate() " & VbCRLF
					strSql = strSql & " ,ezwelprice = '"&imustprice&"' " & VbCRLF
					strSql = strSql & " ,regitemname = '"&html2db(oEzwel.FOneItem.FItemname)&"' " & VbCRLF
					' If ichgSellYn = "N" Then
					' 	strSql = strSql & " ,ezwelsellyn = '"&ichgSellYn&"' " & VbCRLF
					' ElseIf ichgSellYn = "AdminOK" OR ichgSellYn = "Y" Then
					' 	strSql = strSql & " ,ezwelsellyn = 'Y' " & VbCRLF
					' End If

					'ezwel 재판매될 값을 ezwelstatcd 값을 4로 정의
					If mayReSell = "Y" Then
						strSql = strSql & " ,ezwelStatcd = '4' " & VbCRLF
					End If

					If oEzwel.FOneItem.isImageChanged Then
						strSql = strSql & " ,regImageName = '"&oEzwel.FOneItem.getBasicImage&"' " & VbCRLF
					End If
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute(strSql)

					If optMust = "all" Then
						strSql = ""
						strSql = strSql & " DELETE FROM db_item.dbo.tbl_OutMall_regedoption WHERE itemid = '"&iitemid&"' and mallid = 'ezwel' "
						dbget.Execute strSql

						strSql = ""
						strSql = strSql &  "SELECT count(*) as cnt "
						strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
						strSql = strSql & " WHERE itemid=" & iitemid
						strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
						rsget.CursorLocation = adUseClient
						rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
							ocount = rsget("cnt")
						rsget.Close

						If ocount = 0 Then
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
							strSql = strSql & " ('"&iitemid&"',  '0000', 'ezwel', '', '"&html2db(Toptionname)&"', 'Y', '"&Tlimityn&"', '"&Titemsu&"', '0', getdate()) "
							dbget.Execute strSql
						Else
							strSql = ""
							strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
							strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
							strSql = strSql & " SELECT itemid, itemoption, 'ezwel', '', optionname "
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
						strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ezwel_regItem SET "
						strSql = strSql & " regedOptCnt = " & ocount
						strSql = strSql & " WHERE itemid = " & iitemid
						dbget.Execute strSql
					End If
					EzwelOneItemEdit = true
					Set objXML = Nothing
					Set xmlDOM = Nothing
					If mayReSell = "Y" Then
						iErrStr = "OK||"&iitemid&"||성공(수정[재판매])"
					Else
						iErrStr = "OK||"&iitemid&"||성공(수정)"
					End If
				Else						'실패(E)
					iErrStr = "ERR||"&iitemid&"||"&iMessage
				End If
			On Error Goto 0
		Else
			iErrStr = "ERR||"&iitemid&"||ezwel 통신 오류가 발생했습니다.[ERR-SELLEDIT-002]"
		End If
End Function

Function EzwelOneItemEditSellyn(iitemid, iEzwelGoodNo, byRef iErrStr, strParam, imustprice, ichgSellYn, optMust, ilimityn, ilimitno, ilimitsold, ichkXML)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, strSql, tenOptCnt
	Dim retCode, goodsCd, iMessage, AssignedRow, oMsg, ocount, Toptionname, Tlimitno, Tlimitsold, Tlimityn, Titemsu
	On Error Resume Next
'rw oEzwel.FOneItem.isImageChanged
'rw oEzwel.FOneItem.getBasicImage
'rw "=-=-="
'response.end
	If ichkXML = "Y" Then
		response.write replace(xmlStr, "?xml", "?AAAAA")
	End If

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
		objXML.send(postParam&xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		If ichkXML = "Y" Then
			response.write replace(BinaryToText(objXML.ResponseBody, "euc-kr"), "?xml", "?AAAAA")
			response.end
		End If

		goodsCd		= xmlDOM.getElementsByTagName("goodsCd").item(0).text
		retCode		= xmlDOM.getElementsByTagName("resultCode").item(0).text
		iMessage	= xmlDOM.getElementsByTagName("resultMsg").item(0).text

		If retCode = "200" Then		'성공(200)
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ezwel_regItem " & VbCRLF
			strSql = strSql & " SET accFailCnt = 0 " & VbCRLF
			strSql = strSql & " ,ezwelLastUpdate = getdate() " & VbCRLF
			strSql = strSql & " ,ezwelprice = '"&imustprice&"' " & VbCRLF
			strSql = strSql & " ,regitemname = '"&html2db(oEzwel.FOneItem.FItemname)&"' " & VbCRLF
			If ichgSellYn = "N" Then
				strSql = strSql & " ,ezwelsellyn = '"&ichgSellYn&"' " & VbCRLF
			ElseIf ichgSellYn = "AdminOK" OR ichgSellYn = "Y" Then
				strSql = strSql & " ,ezwelsellyn = 'Y' " & VbCRLF
			End If

			If oEzwel.FOneItem.isImageChanged Then
				strSql = strSql & " ,regImageName = '"&oEzwel.FOneItem.getBasicImage&"' " & VbCRLF
			End If
			strSql = strSql & " WHERE itemid='" & iitemid & "'"
			dbget.Execute(strSql)

			If optMust = "all" Then
				strSql = ""
				strSql = strSql & " DELETE FROM db_item.dbo.tbl_OutMall_regedoption WHERE itemid = '"&iitemid&"' and mallid = 'ezwel' "
				dbget.Execute strSql

				strSql = ""
				strSql = strSql &  "SELECT count(*) as cnt "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & iitemid
				strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					ocount = rsget("cnt")
				rsget.Close

				If ocount = 0 Then
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
					strSql = strSql & " ('"&iitemid&"',  '0000', 'ezwel', '', '"&html2db(Toptionname)&"', 'Y', '"&Tlimityn&"', '"&Titemsu&"', '0', getdate()) "
					dbget.Execute strSql
				Else
					strSql = ""
					strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
					strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
					strSql = strSql & " SELECT itemid, itemoption, 'ezwel', '', optionname "
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
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ezwel_regItem SET "
				strSql = strSql & " regedOptCnt = " & ocount
				strSql = strSql & " WHERE itemid = " & iitemid
				dbget.Execute strSql
			End If
			EzwelOneItemEditSellyn = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			If ichgSellYn = "N" Then
				iErrStr = "OK||"&iitemid&"||품절처리"
			ElseIf ichgSellYn = "AdminOK" Then
				iErrStr = "OK||"&iitemid&"||이지웰페어 담당자 승인 후 판매중 처리"
			Else
				iErrStr = "OK||"&iitemid&"||성공(수정)"
			End If
		Else						'실패(E)
			iErrStr = "ERR||"&iitemid&"||"&iMessage
		End If
		On Error Goto 0
	Else
		iErrStr = "ERR||"&iitemid&"||ezwel 통신 오류가 발생했습니다.[ERR-SELLEDIT-002]"
	End If
End Function

Function EzwelItemChkstat(iitemid, byRef iErrStr, iEzwelGoodNo)
	Dim objXML, xmlDOM, strSql, tenOptCnt
	Dim retCode, goodsCd, iMessage, AssignedRow, goodsStatus, ezwelsellyn
	Dim getParam
	getParam = "cspCd=10040413&crtCd=8e5a6dbdd27efb49fc600c293884ef47"
	getParam = getParam & "&goodsCd=" & iEzwelGoodNo

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://api.ezwel.com/if/api/goodsListAPI.ez?" & getParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				If (session("ssBctID")="kjy8517") Then
					response.write "조회Response : "
					response.write "<textarea cols=80 rows=2>"&BinaryToText(objXML.ResponseBody, "euc-kr")&"</textarea><br />"
				End If

				retCode		= xmlDOM.getElementsByTagName("resultCode").item(0).text
				iMessage	= xmlDOM.getElementsByTagName("resultMsg").item(0).text

				If retCode = "200" Then		'성공(200)
					' rw xmlDOM.getElementsByTagName("goodsCd").item(0).text
					' rw xmlDOM.getElementsByTagName("cspGoodsCd").item(0).text
					' rw xmlDOM.getElementsByTagName("goodsNm").item(0).text
					' rw xmlDOM.getElementsByTagName("regDt").item(0).text
					goodsStatus = xmlDOM.getElementsByTagName("goodsStatus").item(0).text
					Select Case goodsStatus
						Case "판매중"
							ezwelsellyn = "Y"
						Case Else
							ezwelsellyn = "N"
					End Select

					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ezwel_regItem SET "
					If goodsStatus = "등록대기" Then
						strSql = strSql & " ezwelStatcd = '3' "
						strSql = strSql & " , lastStatCheckDate = getdate() "
					Else
						strSql = strSql & " ezwelStatcd = '7' "
						strSql = strSql & " , ezwelsellyn = '"& ezwelsellyn &"' "
						strSql = strSql & " , lastStatCheckDate = getdate() "
					End If
					strSql = strSql & " WHERE itemid = '"& iitemid &"' "
					dbget.Execute strSql

					iErrStr = "OK||"&iitemid&"||성공("&goodsStatus&")"
				Else						'실패(E)
					iErrStr = "ERR||"&iitemid&"||"&iMessage
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||ezwel 통신 오류가 발생했습니다.[ERR-CHKSTAT-002]"
		End If
	Set objXML = nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function getEzwelToken()
	Dim objXML, xmlDOM, iRbody, strSql, returnCode, resultMsg, accessToken, refreshToken
	Dim obj, strParam, strObj
	Dim updateAuth, dbAuthNo
	strSql = ""
	strSql = strSql & " SELECT TOP 1 isnull(accessToken, '') as accessToken, lastupdate "&VbCRLF
	strSql = strSql & " FROM db_etcmall.dbo.tbl_outmall_ini"&VbCRLF
	strSql = strSql & " WHERE mallid='"& CMALLNAME &"'"&VbCRLF
	strSql = strSql & " and inikey = 'auth'"
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.Eof then
		dbAuthNo	= rsget("accessToken")
		updateAuth	= rsget("lastupdate")
	end if
	rsget.close

	If DateDiff("h", updateAuth, now()) > 6 OR dbAuthNo = "" then
		Set obj = jsObject()
			obj("cspCd") = cspCd
			obj("crtCd") = crtCd
			strParam = obj.jsString
		Set obj = nothing

		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.open "POST", ezwelNewAPIURL & "/api/auth/request-token", false
			objXML.setRequestHeader "Content-Type", "application/json"
			objXML.Send(strParam)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
				Set strObj = JSON.parse(iRbody)
					returnCode		= strObj.resultCode
					resultMsg		= strObj.resultMsg
					If returnCode = "200" Then
						accessToken = strObj.accessToken
						refreshToken = strObj.refreshToken

						strSql = ""
						strSql = strSql & " UPDATE db_etcmall.dbo.tbl_outmall_ini "&VbCRLF
						strSql = strSql & " SET accessToken = '"& accessToken &"'"&VbCRLF
						strSql = strSql & " ,refreshToken = '"& refreshToken &"'"&VbCRLF
						strSql = strSql & " ,lastupdate = GETDATE()"&VbCRLF
						strSql = strSql & " WHERE mallid = '"& CMALLNAME &"'"&VbCRLF
						strSql = strSql & " and inikey='auth'"
						dbget.Execute strSql
					Else
						rw resultMsg
					End If
				Set strObj = nothing
			End If
		Set objXML = nothing
	End If
End Function

Function EzwelItemNewReg(iitemid, strParam, byRef iErrStr, iSellCash, iezwelSellYn, ilimityn, ilimitno, ilimitsold, iitemname, iimageNm)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, strSql, tenOptCnt, resultMsg, iRbody, strObj
	Dim returnCode, goodsCd, iMessage, AssignedRow, Toptionname, Tlimitno, Tlimitsold, Tlimityn

	Call getEzwelToken()

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelNewAPIURL & "/if/api/goodsSaveAPI", false
		objXML.setRequestHeader "X-Ezwel-Token", getAccessToken
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If (session("ssBctID")="kjy8517") Then
			response.write "등록 Request : "
			response.write "<textarea cols=80 rows=2>"&strParam&"</textarea><br />"
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.resultCode
				resultMsg		= strObj.resultMsg
				If returnCode = "200" Then
					goodsCd		= strObj.goodsCd
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCRLF
					strSql = strSql & "	SET ezwelLastUpdate = GETDATE() "  & VbCRLF
					strSql = strSql & "	, ezwelGoodNo = '" & goodsCd & "'"  & VbCRLF
					strSql = strSql & "	, ezwelPrice = " &iSellCash& VbCRLF
					strSql = strSql & "	, accFailCnt = 0"& VbCRLF
					strSql = strSql & "	, ezwelRegdate = isNULL(ezwelRegdate, GETDATE())"
					If (goodsCd <> "") Then
						strSql = strSql & "	, ezwelstatCD = '3'"& VbCRLF					'등록완료(임시)
					Else
						strSql = strSql & "	, ezwelstatCD = '1'"& VbCRLF					'전송시도
					End If
					strSql = strSql & " ,regimageName = '"&iimageNm&"'"& VbCrlf
					strSql = strSql & "	FROM db_etcmall.dbo.tbl_ezwel_regItem R"& VbCRLF
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||성공(등록)"
				Else						'실패(E)
					iErrStr = "ERR||"&iitemid&"||"& resultMsg &"(등록)"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||ezwel 통신 오류가 발생했습니다.[ERR-REG-002]"
		End If
	Set objXML = nothing
End Function

Function EzwelItemNewChkstat(iitemid, ezwelGoodNo, ilimityn, ilimitno, ilimitsold, iErrStr)
	Dim objXML, xmlDOM, iRbody, strSql, returnCode, resultMsg
	Dim obj, strParam, strObj, i
	Dim goodsStatus, realSalePrice, ezwelStatCd, ezwelsellyn, regOptCnt
	Dim optionFullContentList, optionContent1, optionAddAmt, optionAddPrice, cspOptionFullNum
	Dim tlimityn, tlimitsu
	regOptCnt = 0
	Call getEzwelToken()

	Set obj = jsObject()
		obj("goodsCd") = ezwelGoodno
		strParam = obj.jsString
	Set obj = nothing

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelNewAPIURL & "/if/api/goodsViewAPI", false
		objXML.setRequestHeader "X-Ezwel-Token", getAccessToken
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

'		If (session("ssBctID")="kjy8517") Then
			response.write "조회 Response : "
			response.write "<textarea cols=80 rows=2>"&BinaryToText(objXML.ResponseBody,"utf-8")&"</textarea><br />"
'		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.resultCode
				resultMsg		= strObj.resultMsg
				If returnCode = "200" Then
					goodsStatus		= strObj.goodsInfo.goodsStatus						'등록(1001), 판매중(1002), 판매중지(1005), 삭제(1006)
					realSalePrice	= strObj.goodsInfo.realSalePrice

					Select Case goodsStatus
						Case "1001"
							ezwelStatCd	= 3
							ezwelsellyn = "N"
						Case "1002"
							ezwelStatCd	= 7
							ezwelsellyn = "Y"
						Case "1005"
							ezwelStatCd	= 7
							ezwelsellyn = "N"
						Case "1006"
							ezwelStatCd	= 7
							ezwelsellyn = "N"
					End Select

					Set optionFullContentList = strObj.goodsInfo.optionFullContentList
						strSql = ""
						strSql = strSql & " DELETE FROM db_item.dbo.tbl_OutMall_regedoption WHERE itemid = '"&iitemid&"' and mallid = 'ezwel' "
						dbget.Execute strSql
						If optionFullContentList.length > 0 Then
							For i=0 to optionFullContentList.length-1
'								optionFullContentList.get(i).ImgPath					'옵션상세이미지
'								optionFullContentList.get(i).optionCdNm					'옵션명
								optionAddAmt		= optionFullContentList.get(i).optionAddAmt				'옵션수량
'								optionFullContentList.get(i).optionAddBuyPrice			'옵션매입가
								cspOptionFullNum	= optionFullContentList.get(i).cspOptionFullNum			'업체옵션상세코드
'								optionFullContentList.get(i).optionSortNo				'옵션상세정렬순번
'								optionFullContentList.get(i).optionFullUseYn			'옵션상세사용여부
								optionAddPrice		= optionFullContentList.get(i).optionAddPrice				'옵션추가가격
								optionContent1		= optionFullContentList.get(i).optionContent1				'옵션내용1

								strSql = ""
								strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
								strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
								strSql = strSql & " SELECT itemid, itemoption, 'ezwel', '', optionname "
								strSql = strSql & " ,Case WHEN ("& optionAddAmt &" < 1) THEN 'N' Else 'Y' End, optlimityn " & VBCRLF
								strSql = strSql & " ,'"& optionAddAmt &"', '"& optionAddPrice &"', GETDATE() " & VBCRLF
								strSql = strSql & " FROM db_item.dbo.tbl_item_option " & VBCRLF
								strSql = strSql & " WHERE itemid= '"&iitemid&"' " & VBCRLF
								If cspOptionFullNum <> "" Then
									strSql = strSql & " AND itemoption = '"& cspOptionFullNum &"' " & VBCRLF
								Else
									strSql = strSql & " AND optionname = '"& html2db(optionContent1) &"' " & VBCRLF
								End If
								dbget.Execute strSql

								If (i mod 100) = 0 Then
									response.flush
								End If
								regOptCnt = optionFullContentList.length
							Next
						Else
							If ilimityn = "Y" Then
								tlimityn = "Y"
								If ilimitno - ilimitsold - 5 < 0 Then
									tlimitsu = 0
								Else
									tlimitsu = ilimitno - ilimitsold - 5
								End If
							Else
								tlimityn = "N"
								tlimitsu = 10000
							End If

							strSql = ""
							strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
							strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
							strSql = strSql & " VALUES " & VBCRLF
							strSql = strSql & " ('"&iitemid&"',  '0000', 'ezwel', '', '단일상품', 'Y', '"&ilimityn&"', '"&tlimitsu&"', '0', getdate()) "
							dbget.Execute strSql
							regOptCnt = 0
						End If
					Set optionFullContentList = nothing

					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ezwel_regitem "
					strSql = strSql & " SET ezwelStatCd = '"& ezwelStatCd &"' "
					strSql = strSql & " , ezwelSellYn = '"& ezwelsellyn &"' "
					strSql = strSql & " , ezwelPrice = '"& realSalePrice &"' "
					strSql = strSql & " , regedOptCnt = '"& regOptCnt &"' "
					strSql = strSql & " WHERE itemid = '"& iitemid &"' "
					dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||성공(조회)"
				Else						'실패(E)
					iErrStr = "ERR||"&iitemid&"||"& resultMsg &"(조회)"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||ezwel 통신 오류가 발생했습니다.[ERR-CHKSTAT-002]"
		End If
	Set objXML = nothing
End Function

Function EzwelItemPrice(iitemid, strParam, imustprice, iErrStr)
	Dim objXML, xmlDOM, iRbody, strSql, returnCode, resultMsg
	Dim obj, strObj, i

	Call getEzwelToken()

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelNewAPIURL & "/if/api/goodsPriceAPI", false
		objXML.setRequestHeader "X-Ezwel-Token", getAccessToken
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.resultCode
				resultMsg		= strObj.resultMsg
				If returnCode = "200" Then
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ezwel_regitem "
					strSql = strSql & " SET ezwelPrice = '"& imustprice &"' "
					strSql = strSql & " ,ezwelLastUpdate = GETDATE() "
					strSql = strSql & "	,accFailCnt = 0"
					strSql = strSql & " WHERE itemid = '"& iitemid &"' "
					dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||성공(가격)"
				Else						'실패(E)
					iErrStr = "ERR||"&iitemid&"||"& resultMsg &"(가격)"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||ezwel 통신 오류가 발생했습니다.[ERR-PRICE-002]"
		End If
	Set objXML = nothing
End Function

Function EzwelItemOption(iitemid, strParam, iErrStr)
	Dim objXML, xmlDOM, iRbody, strSql, returnCode, resultMsg
	Dim obj, strObj, i

	Call getEzwelToken()

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelNewAPIURL & "/if/api/goodsOptionUpdateAPI", false
		objXML.setRequestHeader "X-Ezwel-Token", getAccessToken
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
rw iRbody
response.end
		Else
			iErrStr = "ERR||"&iitemid&"||ezwel 통신 오류가 발생했습니다.[ERR-PRICE-002]"
		End If
	Set objXML = nothing
End Function

Function EzwelNewEditSellyn(iitemid, ichgSellyn, ezwelGoodno, iErrStr)
	Dim objXML, xmlDOM, iRbody, strSql, returnCode, resultMsg
	Dim obj, strParam, strObj
	Dim EzwelStatus
	
	Select Case ichgSellyn
		Case "Y"		EzwelStatus = "1002"
		Case "N"		EzwelStatus = "1005"
	End Select

	Call getEzwelToken()
	Set obj = jsObject()
		obj("goodsCd") = ezwelGoodno
		obj("goodsStatus") = EzwelStatus		'등록(1001), 판매중(1002),판매중지(1005), 삭제(1006)
		strParam = obj.jsString
	Set obj = nothing

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelNewAPIURL & "/if/api/goodsStatusAPI", false
		objXML.setRequestHeader "X-Ezwel-Token", getAccessToken
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.resultCode
				resultMsg		= strObj.resultMsg
				If returnCode = "200" Then
					If ichgSellyn = "Y" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET ezwelSellyn = 'Y'"
						strSql = strSql & "	,ezwelLastUpdate = GETDATE()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_ezwel_regitem R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||판매(상태)"
					ElseIf ichgSellyn = "N" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET ezwelSellyn = 'N'"
						strSql = strSql & "	,accFailCnt = 0"
						strSql = strSql & "	,ezwelLastUpdate = GETDATE()"
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_ezwel_regitem R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||품절처리(상태)"
					End If
				Else
					iErrStr = "ERR||"&iitemid&"||"& resultMsg &"(상태)"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||ezwel 통신 오류가 발생했습니다.[ERR-SELLYN-002]"
		End If
	Set objXML = nothing
End Function

Function EzwelItemNewEdit(iitemid, strParam, iErrStr, iitemname, iimageNm)
	Dim objXML, xmlDOM, iRbody, strSql, returnCode, resultMsg
	Dim obj, strObj, i

	Call getEzwelToken()

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelNewAPIURL & "/if/api/goodsSaveAPI", false
		objXML.setRequestHeader "X-Ezwel-Token", getAccessToken
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)
'		If (session("ssBctID")="kjy8517") Then
			response.write "수정 Request : "
			response.write "<textarea cols=80 rows=2>"&strParam&"</textarea><br />"
'		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.resultCode
				resultMsg		= strObj.resultMsg
				If returnCode = "200" Then
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ezwel_regitem "
					strSql = strSql & " SET ezwelLastUpdate = GETDATE() "
					strSql = strSql & "	,accFailCnt = 0"
					strSql = strSql & " ,regitemname = '"& iitemname &"'"
					strSql = strSql & " ,regimageName = '"& iimageNm &"'"
					strSql = strSql & " WHERE itemid = '"& iitemid &"' "
					dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||성공(수정)"
				Else						'실패(E)
					iErrStr = "ERR||"&iitemid&"||"& resultMsg &"(수정)"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||ezwel 통신 오류가 발생했습니다.[ERR-EDIT-002]"
		End If
	Set objXML = nothing
End Function

Function fnEzwelMafcList()
	Dim objXML, xmlDOM, iRbody, strSql, returnCode, resultMsg
	Dim obj, strParam, strObj, i
	Dim mafcList, mafcCd, mafcNm
	Dim SubNodes, Nodes

	Call getEzwelToken()
	strParam = ""
	strParam = strParam & "<?xml version=""1.0"" encoding=""UTF-8""?>"
	strParam = strParam & "<dataSet>"
	strParam = strParam & "   	<mafcCd></mafcCd>"
	strParam = strParam & "  	<mafcNm></mafcNm>"
	strParam = strParam & "</dataSet>"

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelNewAPIURL & "/if/api/getMafcListAPI", false
		objXML.setRequestHeader "Accept", "application/xml"
		objXML.setRequestHeader "Content-Type", "application/xml;charset=UTF-8"
		objXML.setRequestHeader "X-Ezwel-Token", getAccessToken
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.send(strParam)

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
				xmlDOM.LoadXML iRbody
				returnCode	= xmlDOM.getElementsByTagName("resultCode").item(0).text
				resultMsg	= xmlDOM.getElementsByTagName("resultMsg").item(0).text
				If returnCode = "200" Then
					Set mafcList = xmlDOM.getElementsByTagName("mafcList")
						strSql = ""
						strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_ezwel_mafcList] "
						dbget.Execute(strSql)
						For each SubNodes in mafcList
							mafcCd		= SubNodes.SelectSingleNode("mafcCd").Text		'제조사코드
							mafcNm		= SubNodes.SelectSingleNode("mafcNm").Text		'제조사명
							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_ezwel_mafcList] (mafcCd, mafcNm, regdate) VALUES "
							strSql = strSql & " ('"&mafcCd&"', '"&html2db(mafcNm)&"', GETDATE()) "
							dbget.Execute(strSql)
							If (i mod 1000) = 0 Then
								response.flush
							End If
						Next
					Set mafcList = nothing
				Else
					rw resultMsg
				End If
'rw "A : " & iRbody
			Set xmlDOM = nothing
		End If
	Set objXML = nothing
End Function

Function fnEzwelBrandList()
	Dim objXML, xmlDOM, iRbody, strSql, returnCode, resultMsg
	Dim obj, strParam, strObj, i
	Dim brandList, brandCd, brandNm
	
	Call getEzwelToken()
	Set obj = jsObject()
		obj("brandCd") = ""
		obj("brandNm") = ""
		strParam = obj.jsString
	Set obj = nothing

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelNewAPIURL & "/if/api/getBrandListAPI", false
		objXML.setRequestHeader "X-Ezwel-Token", getAccessToken
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.resultCode
				resultMsg		= strObj.resultMsg
				If returnCode = "200" Then
					Set brandList = strObj.brandList
						If brandList.length > 0 Then
							strSql = ""
							strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_ezwel_brandList] "
							dbget.Execute(strSql)

							For i=0 to brandList.length-1
								brandCd		= brandList.get(i).brandCd		'브랜드코드
								brandNm		= brandList.get(i).brandNm		'브랜드명

								strSql = ""
								strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_ezwel_brandList] (brandCd, brandNm, regdate) VALUES "
								strSql = strSql & " ('"&brandCd&"', '"&html2db(brandNm)&"', GETDATE()) "
								dbget.Execute(strSql)

								If (i mod 1000) = 0 Then
									response.flush
								End If
							Next
							rw brandList.length & " 건 등록"
						End If
					Set brandList = nothing
				Else
					rw resultMsg
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing

End Function

%>