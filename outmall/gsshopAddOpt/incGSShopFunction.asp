<%
Public gsshopAPIURL
Public gsshopNewAPIURL
IF application("Svr_Info") = "Dev" THEN
	gsshopAPIURL = "http://test1.gsshop.com/alia/aliaCommonPrd.gs"	'테스트서버
	gsshopNewAPIURL = "http://realapi.gsshop.com/alia/aliaCommonPrd.gs"
Else
	gsshopAPIURL = "http://ecb2b.gsshop.com/alia/aliaCommonPrd.gs"	'실서버
	gsshopNewAPIURL = "http://realapi.gsshop.com/alia/aliaCommonPrd.gs"
End If
'############################################## 실제 수행하는 API 함수 모음 ##############################################
'New 상품 등록 함수
Function fnGSShopNewItemReg(iitemid, strParam, byRef iErrStr, iRealSellprice, iGSShopSellYn, ilimityn, ilimitno, ilimiysold, iitemname, iitemoption, imidx, ioptionname)
	Dim objXML, xmlDOM, strRst
	Dim buf, strSql, AssignedRow
	Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	Dim attrPrdlist, lp, tenOptcd, gsOptcd, strObj, iRbody
	Dim Tlimitno, Tlimitsold, Titemoption, Toptionname, Toptlimitno, Toptlimitsold, Toptsellyn, Toptlimityn, Toptaddprice, Tlimityn, Tsellyn, Titemsu, Tsellcash

'	On Error Resume Next
	fnGSShopNewItemReg = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
				response.write iRbody
			End If
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					If Err.number <> 0 Then
						iErrStr = "ERR||"&imidx&"||"&Err.Description&"(ERR.상품등록)"
					Else
						iErrStr = "ERR||"&imidx&"||"&resultmsg&"(상품등록)"
					End If
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					'상품존재여부 확인
					strSql = "Select count(*) From db_etcmall.dbo.tbl_gsshopAddoption_regitem Where midx='" & imidx & "'"
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If rsget(0) > 0 Then
						'// 존재 -> 수정
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCRLF
						strSql = strSql & "	Set GSShopLastUpdate = getdate() "  & VbCRLF
						strSql = strSql & "	, GSShopGoodNo = '" & prdCd & "'"  & VbCRLF
						strSql = strSql & "	, GSShopPrice = " &iRealSellprice& VbCRLF
						strSql = strSql & "	, accFailCnt = 0"& VbCRLF
						strSql = strSql & "	, GSShopRegdate = isNULL(GSShopRegdate, getdate())"& VbCRLF
						strSql = strSql & "	, GSShopSellYn = '" & iGSShopSellYn & "'"& VbCRLF
						If (prdCd <> "") Then
						    strSql = strSql & "	, GSShopstatCD = '3'"& VbCRLF					'등록완료(임시)
						Else
							strSql = strSql & "	, GSShopstatCD = '1'"& VbCRLF					'전송시도
						End If
						strSql = strSql & "	From db_etcmall.dbo.tbl_gsshopAddoption_regitem R"& VbCRLF
						strSql = strSql & " Where R.midx = '" & imidx & "'"
						dbget.Execute(strSql)
					Else
						'// 없음 -> 신규등록
						strSql = ""
						strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_gsshopAddoption_regitem "
						strSql = strSql & " (regedOptCnt, reguserid, GSShopRegdate, GSShopLastUpdate, GSShopGoodNo, GSShopPrice, GSShopSellYn, GSShopStatCd, accFailCnt) VALUES " & VbCRLF
						strSql = strSql & " (0" &_
						strSql = strSql & " , '" & session("ssBctId") & "'" &_
						strSql = strSql & " , getdate(), getdate()" & VBCRLF
						strSql = strSql & " , '" & prdCd & "'" & VBCRLF
						strSql = strSql & " , '" & iRealSellprice & "'" & VBCRLF
						strSql = strSql & " , '" & iGSShopSellYn & "'" & VBCRLF
						If (prdCd <> "") Then
						    strSql = strSql & ",'3'"											'등록완료(임시)
						Else
						    strSql = strSql & ",'1'"											'전송시도
						End If
						strSql = strSql & " , 0" & VBCRLF
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
					strSql = strSql & " and R.mallid= 'gsshop' "
					iErrStr =  "OK||"&imidx&"||"&resultmsg&"(상품등록)"
					fnGSShopNewItemReg = True
		        End If
			Set strObj = nothing
		Else
			fnGSShopNewItemReg = False
			iErrStr = "ERR||"&imidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-REG-002]"
		End If
	Set objXML= nothing
End Function

'상품 등록 함수
Function fnGSShopItemReg(iitemid, strParam, byRef iErrStr, iRealSellprice, iGSShopSellYn, ilimityn, ilimitno, ilimiysold, iitemname, iitemoption, imidx, ioptionname)
	Dim objXML, xmlDOM, strRst
	Dim buf, strSql, AssignedRow
	Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	Dim attrPrdlist, lp, tenOptcd, gsOptcd
	Dim Tlimitno, Tlimitsold, Titemoption, Toptionname, Toptlimitno, Toptlimitsold, Toptsellyn, Toptlimityn, Toptaddprice, Tlimityn, Tsellyn, Titemsu, Tsellcash

	On Error Resume Next
	fnGSShopItemReg = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
'rw gsshopAPIURL&"?"&strparam
'response.end
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
		    buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf
				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				resultcode	= Split(buf, "|")(0)	'등록결과코드
				resultmsg	= Split(buf, "|")(1)	'등록결과메시지
				supPrdCd	= Split(buf, "|")(2)	'협력사상품코드
				supCd		= Split(buf, "|")(3)	'협력사코드
				prdCd		= Split(buf, "|")(4)	'상품코드
				attrPrdCd	= Split(buf, "|")(5)	'이샵속성상품코드^협력사속성상품코드,이샵속성상품코드^협력사속성상품코드	'속성파라메타 전송전에 에러가 나면 못 받음

				If resultcode = "S" Then	'성공(S)
					'상품존재여부 확인
					strSql = "Select count(itemid) From db_etcmall.dbo.tbl_gsshopAddoption_regitem Where midx='" & imidx & "'"
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If rsget(0) > 0 Then
						'// 존재 -> 수정
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCRLF
						strSql = strSql & "	Set GSShopLastUpdate = getdate() "  & VbCRLF
						strSql = strSql & "	, GSShopGoodNo = '" & prdCd & "'"  & VbCRLF
						strSql = strSql & "	, GSShopPrice = " &iRealSellprice& VbCRLF
						strSql = strSql & "	, accFailCnt = 0"& VbCRLF
						strSql = strSql & "	, GSShopRegdate = isNULL(GSShopRegdate, getdate())"& VbCRLF
						strSql = strSql & "	, GSShopSellYn = '" & iGSShopSellYn & "'"& VbCRLF
						If (prdCd <> "") Then
						    strSql = strSql & "	, GSShopstatCD = '3'"& VbCRLF					'등록완료(임시)
						Else
							strSql = strSql & "	, GSShopstatCD = '1'"& VbCRLF					'전송시도
						End If
						strSql = strSql & "	From db_etcmall.dbo.tbl_gsshopAddoption_regitem R"& VbCRLF
						strSql = strSql & " Where R.midx = '" & imidx & "'"
						dbget.Execute(strSql)
					Else
						'// 없음 -> 신규등록
						strSql = ""
						strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_gsshopAddoption_regitem "
						strSql = strSql & " (regedOptCnt, reguserid, GSShopRegdate, GSShopLastUpdate, GSShopGoodNo, GSShopPrice, GSShopSellYn, GSShopStatCd, accFailCnt) VALUES " & VbCRLF
						strSql = strSql & " (0" &_
						strSql = strSql & " , '" & session("ssBctId") & "'" &_
						strSql = strSql & " , getdate(), getdate()" & VBCRLF
						strSql = strSql & " , '" & prdCd & "'" & VBCRLF
						strSql = strSql & " , '" & iRealSellprice & "'" & VBCRLF
						strSql = strSql & " , '" & iGSShopSellYn & "'" & VBCRLF
						If (prdCd <> "") Then
						    strSql = strSql & ",'3'"											'등록완료(임시)
						Else
						    strSql = strSql & ",'1'"											'전송시도
						End If
						strSql = strSql & " , 0" & VBCRLF
						strSql = strSql & ")"
						dbget.Execute(strSql)
					End If
					rsget.Close

					attrPrdlist = split(attrPrdCd,",")
					gsOptcd			= split(attrPrdCd,"^")(0)
					Toptionname		= ioptionname
					Tlimitno		= ilimitno
					Tlimitsold		= ilimiysold
					Tlimityn		= ilimityn
					If (Tlimityn="Y") then
						If (Tlimitno - Tlimitsold) < 5 Then
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
					strSql = strSql & " SELECT TOP 1 itemid, itemoption, 'gsshop', '"&gsOptcd&"', '"&Toptionname&"', 'Y', '"&Tlimityn&"', '"&Titemsu&"', optaddprice, getdate() " & VBCRLF
					strSql = strSql & " FROM db_item.dbo.tbl_item_option " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&iitemid&"' " & VBCRLF
					strSql = strSql & " and itemoption = '"&iitemoption&"' " & VBCRLF
					dbget.Execute strSql

					strSql = ""
					strSql = strSql & " UPDATE R "
					strSql = strSql & " SET itemname = i.itemname "
					strSql = strSql & " ,optionname = o.optionname "
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] R "
					strSql = strSql & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid "
					strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on R.itemid = o.itemid and R.itemoption = o.itemoption "
					strSql = strSql & " WHERE R.idx = '"&midx&"' "
					strSql = strSql & " and R.mallid= 'gsshop' "
					iErrStr =  "OK||"&imidx&"||"&resultmsg&"(상품등록)"
				Else						'실패(E)
	                iErrStr =  "ERR||"&imidx&"||"&resultmsg&"(상품등록)"
				End If
			Set xmlDOM = Nothing
			fnGSShopItemReg= true
		Else
			iErrStr = "ERR||"&imidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-REG-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 품절 수행 함수
Public Function fnGSShopNewSellyn(iidx, ichgSellYn, istrParam, byRef iErrStr)
    Dim strParam, resultcode, resultmsg, supPrdCd, supCd, prdCd
    Dim objXML, xmlDOM, strObj
    Dim strRst, strSql, iRbody
    On Error Resume Next
    fnGSShopNewSellyn = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(istrParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(상태변경)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_gsshopAddoption_regitem] " & VbCRLF
					strSql = strSql & " SET GSShopLastUpdate = getdate() " & VbCRLF
					strSql = strSql & " ,lastStatCheckDate = getdate() " & VbCRLF
					strSql = strSql & " ,GSShopSellYn = '" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " ,accFailCnt = 0 " & VbCRLF
					strSql = strSql & " WHERE midx = '" & iidx & "'"
					dbget.Execute(strSql)
					If ichgSellYn = "N" Then
						iErrStr = "OK||"&iidx&"||품절처리"
					Else
						iErrStr = "OK||"&iidx&"||판매중으로 변경"
					End If
		        End If
			Set strObj = nothing
			fnGSShopNewSellyn = true
		Else
			iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-SELLEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'품절 수행 함수
Public Function fnGSShopSellyn(iidx, ichgSellYn, istrParam, byRef iErrStr)
    Dim strParam, resultcode, resultmsg, supPrdCd, supCd, prdCd
    Dim objXML, xmlDOM
    Dim strRst, strSql, buf
    On Error Resume Next
    fnGSShopSellyn = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(istrParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
'					rw buf
				End If

				'결과 코드
				resultcode	= Split(buf, "|")(0)	'등록결과코드
				resultmsg	= Split(buf, "|")(1)	'등록결과메시지
				supPrdCd	= Split(buf, "|")(2)	'협력사상품코드
				supCd		= Split(buf, "|")(3)	'협력사코드
				prdCd		= Split(buf, "|")(4)	'상품코드

				If Err <> 0 Then
					If (IsAutoScript) Then
						iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-SELLEDIT-001]"
					Else
						iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-SELLEDIT-001]"
					End If
					Set objXML = Nothing
				    Set xmlDOM = Nothing
				    On Error Goto 0
				    Exit Function
			    End If

				If resultcode <> "S" Then
					iErrStr = "ERR||"&iidx&"||"&resultmsg
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_gsshopAddoption_regitem] " & VbCRLF
					strSql = strSql & " SET GSShopLastUpdate = getdate() " & VbCRLF
					strSql = strSql & " ,lastStatCheckDate = getdate() " & VbCRLF
					strSql = strSql & " ,GSShopSellYn = '" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " ,accFailCnt = 0 " & VbCRLF
					strSql = strSql & " WHERE midx = '" & iidx & "'"
					dbget.Execute(strSql)
					If ichgSellYn = "N" Then
						iErrStr = "OK||"&iidx&"||품절처리"
					Else
						iErrStr = "OK||"&iidx&"||판매중으로 변경"
					End If
		        End If
			Set xmlDOM = Nothing
			fnGSShopSellyn = True
		Else
			iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-SELLEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 전시상품 판매가 수정
Public Function fnGSShopNewPrice(iidx, istrParam, imustprice, byRef iErrStr)
    Dim objXML,xmlDOM,strRst
    Dim buf, strSql, strObj, iRbody
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewPrice = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(istrParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)

				If Instr(resultmsg, "협력사지급액 금액과 동일합니다") > 0 Then
					resultcode = "True"
				End If

				If resultcode <> "True" Then
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(상품가격)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
				    '// 상품가격정보 수정
				    strSql = ""
	    			strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_gsshopAddoption_regitem] " & VbCRLF
	    			strSql = strSql & "	SET GSShopLastUpdate=getdate() " & VbCRLF
	    			strSql = strSql & "	, GSShopPrice = " & imustprice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " Where midx='" & iidx & "'"& VbCRLF
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(상품가격)"
					fnGSShopNewPrice = True
		        End If
			Set strObj = nothing
		Else
			fnGSShopNewPrice = False
			iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-PRICE-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'전시상품 판매가 수정
Public Function fnGSShopPrice(iidx, istrParam, imustprice, byRef iErrStr)
    Dim objXML,xmlDOM,strRst
    Dim buf, strSql
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopPrice = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(istrParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				'결과 코드
				resultcode	= Split(buf, "|")(0)	'등록결과코드
				resultmsg	= Split(buf, "|")(1)	'등록결과메시지
				supPrdCd	= Split(buf, "|")(2)	'협력사상품코드
				supCd		= Split(buf, "|")(3)	'협력사코드
				prdCd		= Split(buf, "|")(4)	'상품코드

				If Err <> 0 Then
					If (IsAutoScript) Then
						iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-PRICE-001]"
					Else
						iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-PRICE-001]"
					End If
					Set objXML = Nothing
				    Set xmlDOM = Nothing
				    On Error Goto 0
				    Exit Function
			    End If

				If resultcode = "S" Then
				    '// 상품가격정보 수정
				    strSql = ""
	    			strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_gsshopAddoption_regitem] " & VbCRLF
	    			strSql = strSql & "	SET GSShopLastUpdate=getdate() " & VbCRLF
	    			strSql = strSql & "	, GSShopPrice = " & imustprice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " Where midx='" & iidx & "'"& VbCRLF
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&idx&"||"&resultmsg&"(상품가격)"
					fnGSShopPrice = True
				Else
	                iErrStr =  "ERR||"&idx&"||"&resultmsg&"(상품가격)"
					fnGSShopPrice = False
				End If
			Set xmlDOM = Nothing
		Else
			fnGSShopPrice = False
			iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-PRICE-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 상품 옵션 수량 수정
Function fnGSShopNewOPTSuEdit(iitemid, strParam, iidx, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage, i
    Dim buf, tenOptcd, lp, gsOptcd, sqlStr, strObj, iRbody
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd, attrPrdlist, Assignedrow
	On Error Resume Next
	fnGSShopNewOPTSuEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)

				If Instr(resultmsg, "주문가능수량과 동일합니다") > 0 Then
					resultcode = "True"
				End If

				If resultcode <> "True" Then
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(옵션 수량 수정)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					Set attrPrdlist = strObj.attr
						tenOptcd = attrPrdlist.get(0).supAttrPrdCd
						gsOptcd = attrPrdlist.get(0).attrPrdCd
						sqlStr = ""
						sqlStr = sqlStr & "UPDATE db_item.dbo.tbl_OutMall_regedoption SET "
						sqlStr = sqlStr & "outmalllimitno =  "
						sqlStr = sqlStr & "Case WHEN i.limityn = 'Y' and o.optlimitno - o.optlimitsold <= 5 THEN '0' "
						sqlStr = sqlStr & "	 WHEN i.limityn = 'Y' and o.optlimitno - o.optlimitsold > 5 THEN o.optlimitno - o.optlimitsold - 5 "
						sqlStr = sqlStr & "	 WHEN i.limityn = 'N' THEN '999' END "
						sqlStr = sqlStr & "FROM db_item.dbo.tbl_OutMall_regedoption R  "
						sqlStr = sqlStr & "JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid "
						sqlStr = sqlStr & "JOIN db_item.dbo.tbl_item_option o on i.itemid = o.itemid and o.itemoption = '"&tenOptcd&"' "
						sqlStr = sqlStr & "WHERE R.itemid = '"&iitemid&"' and R.itemoption = '"&tenOptcd&"' and R.mallid = 'gsshop' "
						dbget.Execute sqlStr
					Set attrPrdlist = nothing
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(옵션 수량 수정)"
					fnGSShopNewOPTSuEdit = True
		        End If
			Set strObj = nothing
		Else
			fnGSShopNewOPTSuEdit = False
			iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-OPTSuEdit-002]"
		End If
	Set objXML= nothing
	On Error Goto 0
End Function


'상품 옵션 수량 수정
Function fnGSShopOPTSuEdit(iitemid, strParam, iidx, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, tenOptcd, lp, gsOptcd, sqlStr
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd, attrPrdlist, Assignedrow
	On Error Resume Next
	fnGSShopOPTSuEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				If Err <> 0 Then
					iErrStr =  "ERR||"&iidx&"||GSShop과 통신중에 오류가 발생했습니다..[ERR-OPTSuEdit-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If

				'결과 코드
				resultcode	= Split(buf, "|")(0)	'등록결과코드
				resultmsg	= Split(buf, "|")(1)	'등록결과메시지
				supPrdCd	= Split(buf, "|")(2)	'협력사상품코드
				supCd		= Split(buf, "|")(3)	'협력사코드
				prdCd		= Split(buf, "|")(4)	'상품코드
				attrPrdCd	= Split(buf, "|")(5)	'이샵속성상품코드^협력사속성상품코드,이샵속성상품코드^협력사속성상품코드	'속성파라메타 전송전에 에러가 나면 못 받음

				If resultcode = "S" Then	'성공(S)
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(옵션 수량 수정)"
					attrPrdlist = split(attrPrdCd,",")
					gsOptcd		= split(attrPrdlist(0),"^")(0)
	                tenOptcd	= split(attrPrdlist(0),"^")(1)
					sqlStr = ""
					sqlStr = sqlStr & "UPDATE db_item.dbo.tbl_OutMall_regedoption SET "
					sqlStr = sqlStr & "outmalllimitno =  "
					sqlStr = sqlStr & "Case WHEN i.limityn = 'Y' and o.optlimitno - o.optlimitsold <= 5 THEN '0' "
					sqlStr = sqlStr & "	 WHEN i.limityn = 'Y' and o.optlimitno - o.optlimitsold > 5 THEN o.optlimitno - o.optlimitsold - 5 "
					sqlStr = sqlStr & "	 WHEN i.limityn = 'N' THEN '999' END "
					sqlStr = sqlStr & "FROM db_item.dbo.tbl_OutMall_regedoption R  "
					sqlStr = sqlStr & "JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid "
					sqlStr = sqlStr & "JOIN db_item.dbo.tbl_item_option o on i.itemid = o.itemid and o.itemoption = '"&tenOptcd&"' "
					sqlStr = sqlStr & "WHERE R.itemid = '"&iitemid&"' and R.itemoption = '"&tenOptcd&"' and R.mallid = 'gsshop' "
					dbget.Execute sqlStr
				Else						'실패(E)
				    iErrStr =  "ERR||"&iidx&"||"&resultmsg&"(옵션 수량 수정)"
			        Set objXML = Nothing
			        Set xmlDOM = Nothing
			        On Error Goto 0
				    Exit Function
				End If
			Set xmlDOM = Nothing
			fnGSShopOPTSuEdit = True
		Else
			iErrStr =  "ERR||"&iidx&"||GSShop과 통신중에 오류가 발생했습니다..[ERR-OPTSuEdit-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 상품 옵션 상태변경
Function fnGSShopNewOPTSellEdit(iitemid, strParam, iidx, iitemoption, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, tenOptcd, lp, gsOptcd, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd, attrPrdlist, sqlStr
	On Error Resume Next
    fnGSShopNewOPTSellEdit = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(옵션상태변경)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

	                sqlStr = ""
					sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_OutMall_regedoption SET " & VbCRLF
					sqlStr = sqlStr & " outmallsellyn = " & VbCRLF
					sqlStr = sqlStr & " Case WHEN (o.isusing <> 'Y' OR o.optsellyn <> 'Y') THEN 'N'  " & VbCRLF
					sqlStr = sqlStr & " 	 WHEN (i.limityn = 'Y' AND o.optlimitno - o.optlimitsold <= 5) THEN 'N' " & VbCRLF
					sqlStr = sqlStr & " 	 WHEN (R.outmallOptName <> o.optionname) THEN 'N' " & VbCRLF
					sqlStr = sqlStr & " ELSE 'Y' END " & VbCRLF
					sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption R  " & VbCRLF
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid  " & VbCRLF
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_option o on i.itemid = o.itemid and o.itemoption = '"&iitemoption&"'  " & VbCRLF
					sqlStr = sqlStr & " WHERE R.itemid = '"&iitemid&"' and R.itemoption = '"&iitemoption&"' and R.mallid = 'gsshop' "
				    dbget.Execute sqlStr
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(옵션상태변경)"
				End If
			Set strObj = nothing
			fnGSShopNewOPTSellEdit = True
		Else
			iErrStr = "ERR||"&iidx&"||GSShop과 통신중에 오류가 발생했습니다..[ERR-OPTSellEdit-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'상품 옵션 상태변경
Function fnGSShopOPTSellEdit(iitemid, strParam, iidx, iitemoption, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, tenOptcd, lp, gsOptcd
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd, attrPrdlist, sqlStr
	On Error Resume Next
	fnGSShopOPTSellEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				If Err <> 0 Then
					iErrStr = "GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-OPTSellEdit-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If

				'결과 코드
				resultcode	= Split(buf, "|")(0)	'등록결과코드
				resultmsg	= Split(buf, "|")(1)	'등록결과메시지
				supPrdCd	= Split(buf, "|")(2)	'협력사상품코드
				supCd		= Split(buf, "|")(3)	'협력사코드
				prdCd		= Split(buf, "|")(4)	'상품코드
				''봐야될 것 S->P로 변할 수가 있다.(옵션상품에서 단품상품으로 변하는 경우에는 수정 보내지 말고 수정해야되는 데 반드시 처리해야함..

				If resultcode = "S" Then	'성공(S)
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(옵션상태변경)"
	                sqlStr = ""
					sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_OutMall_regedoption SET " & VbCRLF
					sqlStr = sqlStr & " outmallsellyn = " & VbCRLF
					sqlStr = sqlStr & " Case WHEN (o.isusing <> 'Y' OR o.optsellyn <> 'Y') THEN 'N'  " & VbCRLF
					sqlStr = sqlStr & " 	 WHEN (i.limityn = 'Y' AND o.optlimitno - o.optlimitsold <= 5) THEN 'N' " & VbCRLF
					sqlStr = sqlStr & " 	 WHEN (R.outmallOptName <> o.optionname) THEN 'N' " & VbCRLF
					sqlStr = sqlStr & " ELSE 'Y' END " & VbCRLF
					sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption R  " & VbCRLF
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid  " & VbCRLF
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_option o on i.itemid = o.itemid and o.itemoption = '"&iitemoption&"'  " & VbCRLF
					sqlStr = sqlStr & " WHERE R.itemid = '"&iitemid&"' and R.itemoption = '"&iitemoption&"' and R.mallid = 'gsshop' "
				    dbget.Execute sqlStr
				Else						'실패(E)
				    iErrStr =  "ERR||"&iidx&"||"&resultmsg&"(옵션상태변경)"
			        Set objXML = Nothing
			        Set xmlDOM = Nothing
				    Exit Function
				End If
			Set xmlDOM = Nothing
			fnGSShopOPTSellEdit = True
		Else
			iErrStr =  "ERR||"&iidx&"||GSShop과 통신중에 오류가 발생했습니다..[ERR-OPTSellEdit-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 상품명 변경 수행 함수
Public Function fnGSShopChgNewItemname(iidx, strParam, iitemname, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, strSql, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
    On Error Resume Next
    fnGSShopChgNewItemname = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(상품명 변경)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					strSql = ""
					strSql = strSql & " UPDATE R "
					strSql = strSql & " SET itemname = i.itemname "
					strSql = strSql & " ,optionname = o.optionname "
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] R "
					strSql = strSql & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid "
					strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on R.itemid = o.itemid and R.itemoption = o.itemoption "
					strSql = strSql & " WHERE R.idx = '"&iidx&"' "
					strSql = strSql & " and R.mallid= 'gsshop' "
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(상품명 변경)"
		        End If
			Set strObj = nothing
			fnGSShopChgNewItemname = true
		Else
			iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-NMEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function fnGSShopChgItemname(iidx, strParam, iitemname, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, strSql
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopChgItemname = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				If Err <> 0 Then
					iErrStr =  "ERR||"&iidx&"||GSShop과 통신중에 오류가 발생했습니다..[ERR-NMEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'결과 코드
				resultcode	= Split(buf, "|")(0)	'등록결과코드
				resultmsg	= Split(buf, "|")(1)	'등록결과메시지
				supPrdCd	= Split(buf, "|")(2)	'협력사상품코드
				supCd		= Split(buf, "|")(3)	'협력사코드
				prdCd		= Split(buf, "|")(4)	'상품코드

				If resultcode = "S" Then
					strSql = ""
					strSql = strSql & " UPDATE R "
					strSql = strSql & " SET itemname = i.itemname "
					strSql = strSql & " ,optionname = o.optionname "
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] R "
					strSql = strSql & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid "
					strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on R.itemid = o.itemid and R.itemoption = o.itemoption "
					strSql = strSql & " WHERE R.idx = '"&iidx&"' "
					strSql = strSql & " and R.mallid= 'gsshop' "
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(상품명 변경)"
				Else
					iErrStr =  "ERR||"&iidx&"||"&resultmsg
				End If

			Set xmlDOM = Nothing
			fnGSShopChgItemname = True
		Else
			iErrStr =  "ERR||"&iidx&"||GSShop과 통신중에 오류가 발생했습니다..[ERR-NMEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 이미지 변경 수행 함수
Function fnGSShopNewImageEdit(iidx, strParam, iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewImageEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(이미지수정)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(이미지수정)"
		        End If
			Set strObj = nothing
			fnGSShopNewImageEdit = true
		Else
			iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-IMGEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 상품정보 변경 수행 함수
Function fnGSShopNewItemInfoEdit(iidx, strParam, iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewItemInfoEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(상품정보)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(상품정보)"
		        End If
			Set strObj = nothing
			fnGSShopNewItemInfoEdit = true
		Else
			iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-INFOEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function fnGSShopImageEdit(iidx, strParam, iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopImageEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				If Err <> 0 Then
					iErrStr =  "ERR||"&iidx&"||GSShop과 통신중에 오류가 발생했습니다..[ERR-IMGEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'결과 코드
				resultcode	= Split(buf, "|")(0)	'등록결과코드
				resultmsg	= Split(buf, "|")(1)	'등록결과메시지
				supPrdCd	= Split(buf, "|")(2)	'협력사상품코드
				supCd		= Split(buf, "|")(3)	'협력사코드
				prdCd		= Split(buf, "|")(4)	'상품코드

				If resultcode = "S" Then
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(이미지 수정)"
				Else
					iErrStr =  "ERR||"&iidx&"||"&resultmsg
				End If

			Set xmlDOM = Nothing
			fnGSShopImageEdit = True
		Else
			iErrStr =  "ERR||"&iidx&"||GSShop과 통신중에 오류가 발생했습니다..[ERR-IMGEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 전시상품 설명 수정
Function fnGSShopNewContentsEdit(iidx, strParam, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewContentsEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(상품설명수정)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(상품설명수정)"
		        End If
			Set strObj = nothing
			fnGSShopNewContentsEdit = true
		Else
			iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-CONTEDIT-002]"
		End If
End Function

'전시상품 설명 수정
Function fnGSShopContentsEdit(iidx, strParam, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopContentsEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				If Err <> 0 Then
					iErrStr =  "ERR||"&iidx&"||GSShop과 통신중에 오류가 발생했습니다..[ERR-CONTEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'결과 코드
				resultcode	= Split(buf, "|")(0)	'등록결과코드
				resultmsg	= Split(buf, "|")(1)	'등록결과메시지
				supPrdCd	= Split(buf, "|")(2)	'협력사상품코드
				supCd		= Split(buf, "|")(3)	'협력사코드
				prdCd		= Split(buf, "|")(4)	'상품코드

				If resultcode = "S" Then
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(상품설명 수정)"
				Else
					iErrStr =  "ERR||"&iidx&"||"&resultmsg
				End If

			Set xmlDOM = Nothing
			fnGSShopContentsEdit = True
		Else
			iErrStr =  "ERR||"&iidx&"||GSShop과 통신중에 오류가 발생했습니다..[ERR-CONTEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 정부고시항목 수정
Function fnGSShopNewInfodivEdit(iidx, strParam, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewInfodivEdit = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(정부고시항목 수정)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(정부고시항목 수정)"
		        End If
			Set strObj = nothing
			fnGSShopNewInfodivEdit = true
		Else
			iErrStr = "ERR||"&iidx&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-DIVEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'정부고시항목 수정
Function fnGSShopInfodivEdit(iidx, strParam, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopInfodivEdit = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				If Err <> 0 Then
					iErrStr =  "ERR||"&iidx&"||GSShop과 통신중에 오류가 발생했습니다..[ERR-DIVEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'결과 코드
				resultcode	= Split(buf, "|")(0)	'등록결과코드
				resultmsg	= Split(buf, "|")(1)	'등록결과메시지
				supPrdCd	= Split(buf, "|")(2)	'협력사상품코드
				supCd		= Split(buf, "|")(3)	'협력사코드
				prdCd		= Split(buf, "|")(4)	'상품코드

				If resultcode = "S" Then
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(정부고시항목 수정)"
				Else
					iErrStr =  "ERR||"&iidx&"||"&resultmsg
				End If

			Set xmlDOM = Nothing
			fnGSShopInfodivEdit = True
		Else
			iErrStr =  "ERR||"&iidx&"||GSShop과 통신중에 오류가 발생했습니다..[ERR-DIVEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################

'################################################# 각 기능 별 파라메터 정리 ###############################################
'품절 파라메타
Function getGSShopSellynParameter(iidx, ichgSellYn)
	Dim strRst, strSql, newCode

	strSql = ""
	strSql = strSql & " SELECT TOP 1 convert(varchar(30),itemid)+convert(varchar(30),itemoption) as newCode "
	strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] "
	strSql = strSql & " where idx = '"&iidx&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		newCode	= rsget("newCode")
	End If
	rsget.close

	strRst = ""
	strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
	strRst = strRst & "&modGbn=S"														'(*)수정구분 S : 판매상태 수정
	strRst = strRst & "&regId="&COurRedId												'(*)등록자
	'상품기본(prdBaseInfo)
	strRst = strRst & "&supPrdCd="&newCode												'(*)협력사상품코드
	strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
	'상품가격(prdPrc)

	If ichgSellYn = "Y" Then
		strRst = strRst & "&saleEndDtm=29991231235959"									'(*)판매종료일시 | 상품을 중단(판매종료)하려면 중단시점의 판매종료일시를 입력합니다.
	ElseIf (ichgSellYn = "N") Then
		strRst = strRst & "&saleEndDtm="&FormatDate(now(), "00000000000000")			'(*)판매종료일시 | 상품을 중단(판매종료)하려면 중단시점의 판매종료일시를 입력합니다.
	End If
	strRst = strRst & "&attrSaleEndStModYn=N"											'(*)속성판매종료상태수정설정 | 속성구분(S) 상품판매상태를 변경할 때 사용하는 항목으로, 상품마스터 종료 및 해제 시 속성상품의 상태도 함께 종료 및 해제하려면 Y, 상품마스터와 속성 별도로 상태변경 동작 시엔 N

	getGSShopSellynParameter = strRst
End Function

Public Function getGSShopPriceParameter(iidx, byref mustprice)
	Dim strRst, strSql
	Dim sellcash, orgprice, buycash, optaddprice, newCode
	Dim GetTenTenMargin

	strSql = ""
	strSql = strSql & " SELECT TOP 1 sellcash, buycash, orgprice, o.optaddprice, convert(varchar(30),m.itemid) + convert(varchar(30),m.itemoption) as newCode "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
	strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as m on i.itemid = m.itemid and o.itemoption = m.itemoption "
	strSql = strSql & " WHERE m.idx = '"&iidx&"' "
	strSql = strSql & " and m.mallid = 'gsshop' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		sellcash	= rsget("sellcash")
		orgprice	= rsget("orgprice")
		buycash		= rsget("buycash")
		optaddprice	= rsget("optaddprice")
		newCode		= rsget("newCode")
	Else
		getGSShopPriceParameter = ""
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

	'전송 구분 및 반복리스트 건수
	strRst = ""
	strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
	strRst = strRst & "&modGbn=P"														'(*)수정구분 P : 가격 수정
	strRst = strRst & "&regId="&COurRedId												'(*)등록자
	strRst = strRst & "&regSubjCd=SUP"													'(*)등록주체코드 | 엠디가 수정한 경우 : MD, 협력사가 수정한 경우 : SUP
	'상품기본(prdBaseInfo)
	strRst = strRst & "&supPrdCd="&newCode												'(*)협력사상품코드
	strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
	strRst = strRst & "&subSupCd="&COurCompanyCode										'(*)하위협력사코드 | 하위협력사가 없으면 supCd와 같은값 입력
	'상품가격(prdPrc)
	strRst = strRst & "&prdPrcValidStrDtm="&FormatDate(now(), "00000000000000")			'(*)유효시작일시
	strRst = strRst & "&prdPrcValidEndDtm=29991231235959"								'(*)유효종료일시
	strRst = strRst & "&prdPrcSalePrc="&Clng(GetRaiseValue(MustPrice/10)*10)			'(*)판매가격
	'strRst = strRst & "&prdPrcPrchPrc="													'(SYS)매입가격 | (SYS는 저희쪽에서 자동으로 생성해주는 코드 및 값을 말합니다. Null로 보내주시면 됩니다.)
	strRst = strRst & "&prdPrcSupGivRtamtCd=01"											'(*)협력사지급율/액코드 | 01 : 액
	strRst = strRst & "&prdPrcSupGivRtamt="&getGSShopSuplyPrice_update(MustPrice)		'(*)협력사지급율/액 | 기본값 : 판매가*(1-0.12)
	getGSShopPriceParameter = strRst
End Function

Public Function getGSShopItemnameParameter(iidx, byref iitemname)
	Dim strRst, chgname, strSql, newitemname, itemnameChange, newCode
	strSql = ""
	strSql = strSql & " SELECT TOP 1 M.itemid, convert(varchar(30),m.itemid) + convert(varchar(30),m.itemoption) as newCode, isnull(M.newitemname, '') as newitemname, isnull(M.itemnameChange, '') as itemnameChange "
	strSql = strSql & "	FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M "
	strSql = strSql & "	JOIN db_etcmall.[dbo].[tbl_gsshopAddoption_regitem] as R on M.idx = R.midx "
	strSql = strSql & "	WHERE M.idx = '"&iidx&"' "
	strSql = strSql & "	and M.mallid = 'gsshop' "
	strSql = strSql & "	and (R.GSShopStatCd=3 OR R.GSShopStatCd=7) "
	strSql = strSql & " and R.GSShopGoodNo is Not Null "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.Eof Then
		newCode			= rsget("newCode")
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
	chgname = "[텐바이텐]"&replace(iitemname,"'","")		'최초 상품명 앞에 [텐바이텐] 이라고 붙임
	chgname = replace(chgname,"&#8211;","-")
	chgname = replace(chgname,"~","-")
	chgname = replace(chgname,"&","＆")
	chgname = replace(chgname,"<","[")
	chgname = replace(chgname,">","]")
	chgname = replace(chgname,"%","프로")
	chgname = replace(chgname,"+","%2B")
	chgname = replace(chgname,"[무료배송]","")
	chgname = replace(chgname,"[무료 배송]","")

	strRst = ""
	strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
	strRst = strRst & "&modGbn=N"														'(*)수정구분 N : 노출상품명 수정
	strRst = strRst & "&regId="&COurRedId												'(*)등록자
	'상품기본(prdBaseInfo)
	strRst = strRst & "&supPrdCd="&newCode												'(*)협력사상품코드
	strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
	'노출상품명(prdNmChg)
	strRst = strRst & "&prdNmChgValidStrDtm="&FormatDate(now(), "00000000000000")		'(*)유효시작일시
	strRst = strRst & "&prdNmChgValidEndDtm=29991231235959"								'(*)유효종료일시
	strRst = strRst & "&prdNmChgExposPrdNm=" & Trim(chrbyte(chgname,56,"Y"))							'(*)노출상품명 | GSShop노출상품명
	getGSShopItemnameParameter = strRst
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

Public Function getGSShopSuplyPrice_update(iMustPrice)
	getGSShopSuplyPrice_update = CLNG(iMustPrice * (100-CGSSHOPMARGIN) / 100)
End Function
%>