<%
Public gsshopAPIURL
Public gsshopNewAPIURL
IF application("Svr_Info") = "Dev" THEN
	'gsshopAPIURL = "http://test1.gsshop.com/alia/aliaCommonPrd.gs"	'테스트서버
	'gsshopNewAPIURL = "http://testapi.gsshop.com/alia/aliaCommonPrd.gs"
	gsshopAPIURL = "http://ecb2b.gsshop.com/alia/aliaCommonPrd.gs"	'실서버
	gsshopNewAPIURL = "http://realapi.gsshop.com/alia/aliaCommonPrd.gs"
Else
	gsshopAPIURL = "http://ecb2b.gsshop.com/alia/aliaCommonPrd.gs"	'실서버
	gsshopNewAPIURL = "http://realapi.gsshop.com/alia/aliaCommonPrd.gs"
End If
'############################################## 실제 수행하는 API 함수 모음 ##############################################
'New 상품 등록 함수
Function fnGSShopNewItemReg(iitemid, strParam, byRef iErrStr, iSellCash, iGSShopSellYn, ilimityn, ilimitno, ilimiysold, iitemname, iimagename)
	Dim objXML, xmlDOM, strRst
	Dim buf, strSql, AssignedRow
	Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	Dim attrPrdlist, lp, tenOptcd, gsOptcd, strObj, iRbody
	Dim Tlimitno, Tlimitsold, Titemoption, Toptionname, Toptlimitno, Toptlimitsold, Toptsellyn, Toptlimityn, Toptaddprice, Tlimityn, Tsellyn, Titemsu, Tsellcash
	Dim isAttrYn

	On Error Resume Next
	fnGSShopNewItemReg = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			If (session("ssBctID")="kjy8517") Then
				rw "REQ : <textarea cols=40 rows=10>"&strParam&"</textarea>"
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If

			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					If Err.number <> 0 Then
						iErrStr = "ERR||"&iitemid&"||"&Err.Description&"(ERR.상품등록)"
					Else
						iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(상품등록)"
					End If
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					'상품존재여부 확인
					strSql = "Select count(itemid) From db_item.dbo.tbl_gsshop_regitem Where itemid='" & iitemid & "'"
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If rsget(0) > 0 Then
						'// 존재 -> 수정
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCRLF
						strSql = strSql & "	Set GSShopLastUpdate = getdate() "  & VbCRLF
						strSql = strSql & "	, GSShopGoodNo = '" & prdCd & "'"  & VbCRLF
						strSql = strSql & "	, GSShopPrice = " &iSellCash& VbCRLF
						strSql = strSql & "	, accFailCnt = 0"& VbCRLF
						strSql = strSql & "	, GSShopRegdate = isNULL(GSShopRegdate, getdate())"& VbCRLF
						strSql = strSql & "	, GSShopSellYn = '" & iGSShopSellYn & "'"& VbCRLF
						If (prdCd <> "") Then
						    strSql = strSql & "	, GSShopstatCD = '3'"& VbCRLF					'등록완료(임시)
						Else
							strSql = strSql & "	, GSShopstatCD = '1'"& VbCRLF					'전송시도
						End If
						strSql = strSql & "	, regImageName = '" & iimagename & "'"& VbCRLF
						strSql = strSql & "	From db_item.dbo.tbl_gsshop_regItem R"& VbCRLF
						strSql = strSql & " Where R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
					Else
						'// 없음 -> 신규등록
						strSql = ""
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_gsshop_regItem "
						strSql = strSql & " (itemid, regitemname, reguserid, GSShopRegdate, GSShopLastUpdate, GSShopGoodNo, GSShopPrice, GSShopSellYn, GSShopStatCd, regImageName) VALUES " & VbCRLF
						strSql = strSql & " ('" & iitemid & "'" & VBCRLF
						strSql = strSql & " , '" & iitemname & "'" &_
						strSql = strSql & " , '" & session("ssBctId") & "'" &_
						strSql = strSql & " , getdate(), getdate()" & VBCRLF
						strSql = strSql & " , '" & prdCd & "'" & VBCRLF
						strSql = strSql & " , '" & iSellCash & "'" & VBCRLF
						strSql = strSql & " , '" & iGSShopSellYn & "'" & VBCRLF
						If (prdCd <> "") Then
						    strSql = strSql & ",'3'"											'등록완료(임시)
						Else
						    strSql = strSql & ",'1'"											'전송시도
						End If
						strSql = strSql & " , '" & iimagename & "'" & VBCRLF
						strSql = strSql & ")"
						dbget.Execute(strSql)
					End If
					rsget.Close

					' On Error Resume Next
					' If isobject(strObj.attr) Then
					' 	If Err = 0 then
					' 		isAttrYn = "Y"
					' 	Else
					' 		isAttrYn = "N"
					' 	End If
					' End If
					' On Error Goto 0

					' If isAttrYn = "Y" Then				'옵션이라면
					' 	Set attrPrdlist = strObj.attr
					' 		For lp=0 to attrPrdlist.length-1
					' 			tenOptcd = attrPrdlist.get(lp).supAttrPrdCd
					' 			gsOptcd = attrPrdlist.get(lp).attrPrdCd
					' 			strSql = ""
					' 			strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
					' 			strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
					' 			strSql = strSql & " SELECT itemid, itemoption, 'gsshop', '"&gsOptcd&"', optionname, optsellyn, optlimityn, " & VBCRLF
					' 			strSql = strSql & " Case WHEN optlimityn = 'Y' AND optlimitno - optlimitsold <= 5 THEN '0' " & VBCRLF
					' 			strSql = strSql & " 	 WHEN optlimityn = 'Y' AND optlimitno - optlimitsold > 5 THEN optlimitno - optlimitsold - 5 " & VBCRLF
					' 			strSql = strSql & " 	 WHEN optlimityn = 'N' THEN '999' End " & VBCRLF
					' 			strSql = strSql & " , '0', getdate() " & VBCRLF
					' 			strSql = strSql & " FROM db_item.dbo.tbl_item_option " & VBCRLF
					' 			strSql = strSql & " WHERE itemid= '"&iitemid&"' " & VBCRLF
					' 			strSql = strSql & " and itemoption = '"& tenOptcd &"' "
					' 			dbget.Execute strSql
					' 		Next

					' 	Set attrPrdlist = nothing
					' Else								'단품이라면
					' 	strSql = ""
					' 	strSql = strSql & " SELECT COUNT(*) FROM db_item.dbo.tbl_item_option WHERE itemid = '"&iitemid&"' "
					' 	rsget.CursorLocation = adUseClient
					' 	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					' 	If rsget(0) = 0 Then
					' 		tenOptcd	= "0000"
					' 	End If
					' 	rsget.Close

					' 	If (tenOptcd = "0000")  Then	'단일 상품이라면
					' 		'gsOptcd			= split(attrPrdCd,"^")(0)
					' 		gsOptcd			= ""
					' 		Toptionname		= "공통"
					' 		Tlimitno		= ilimitno
					' 		Tlimitsold		= ilimiysold
					' 		Tlimityn		= ilimityn
					' 		If (Tlimityn="Y") then
					' 			If (Tlimitno - Tlimitsold) < 5 Then
					' 				Titemsu = 0
					' 			Else
					' 				Titemsu = Tlimitno - Tlimitsold - 5
					' 			End If
					' 		Else
					' 			Titemsu = 999
					' 		End If
					' 		strSql = ""
					' 		strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
					' 		strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
					' 		strSql = strSql & " VALUES " & VBCRLF
					' 		strSql = strSql & " ('"&iitemid&"',  '"&tenOptcd&"', 'gsshop', '"&gsOptcd&"', '"&html2db(Toptionname)&"', 'Y', '"&Tlimityn&"', '"&Titemsu&"', '0', getdate()) "
					' 		dbget.Execute strSql
					' 	End If
					' End If
					' strSql = ""
					' strSql = strSql & " UPDATE R " & VBCRLF
					' strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VBCRLF
					' strSql = strSql & " FROM db_item.dbo.tbl_gsshop_regItem R " & VBCRLF
					' strSql = strSql & " Join ( " & VBCRLF
					' strSql = strSql & " 	SELECT R.itemid, count(*) as CNT, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt "
					' strSql = strSql & " 	FROM db_item.dbo.tbl_gsshop_regItem R " & VBCRLF
					' strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = 'gsshop' and Ro.itemid = " &iitemid & VBCRLF
					' strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
					' strSql = strSql & " ) T on R.itemid = T.itemid " & VBCRLF
					' dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(상품등록)"
					fnGSShopNewItemReg = True
		        End If
			Set strObj = nothing
		Else
			fnGSShopNewItemReg = False
			iErrStr = "ERR||"&iitemid&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-REG-002]"
		End If
	Set objXML= nothing
End Function

'New 품절 수행 함수
Public Function fnGSShopNewSellyn(iitemid, ichgSellYn, istrParam, byRef iErrStr)
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
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(상태변경)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem " & VbCRLF
					strSql = strSql & " SET GSShopLastUpdate = getdate() " & VbCRLF
					strSql = strSql & " ,lastStatCheckDate = getdate() " & VbCRLF
					strSql = strSql & " ,GSShopSellYn = '" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " ,accFailCnt = 0 " & VbCRLF
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute(strSql)
					If ichgSellYn = "N" Then
						iErrStr = "OK||"&iitemid&"||품절처리"
					Else
						iErrStr = "OK||"&iitemid&"||판매중으로 변경"
					End If
		        End If
			Set strObj = nothing
			fnGSShopNewSellyn = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-SELLEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'조회 함수
Public Function fnGSShopItemView(iitemid, istrParam, byRef iErrStr, iVal)
    Dim strParam, resultcode, resultmsg, supPrdCd, supCd, prdCd
    Dim objXML, xmlDOM, strObj, i, AssignedRow
    Dim strRst, strSql, iRbody

	Dim gsshopsellYn, GSShopGoodNo, gsshopPrice, prdPrcList, prdAttrInfoList, outmallSellyn, outmallOptCode, outmallOptName, tenOptcd, outmalllimitno
    On Error Resume Next
    fnGSShopItemView = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", "http://realapi.gsshop.com/api/v3/getPrdInfo.gs", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(istrParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.result
				resultmsg	= replaceMsg(strObj.message)
				If resultcode <> "success" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(조회)"
				Else
					strSql = ""
					strSql =  strSql & " DELETE FROM db_item.dbo.tbl_OutMall_regedoption WHERE mallid='"&CMALLNAME&"' and itemid="&iitemid&" "
					dbget.Execute strSql

					GSShopGoodNo		= strObj.data.prdBaseInfo.prdCd
					gsshopsellYn = ""
					If strObj.data.prdBaseInfo.prdStCd = "Y" Then	'판매상태 | "판매대기 : N (MD승인 후 최종 승인나기 전), 판매중 : Y, 판매종료 : E, 일시품절 : T, 완전판매종료 : D (판종사유코드가 31)"
						gsshopsellYn = "Y"
					ElseIf strObj.data.prdBaseInfo.prdStCd = "N" Then
						gsshopsellYn = "E"
					Else
						gsshopsellYn = "N"
					End If
					Set prdPrcList = strObj.data.prdPrcList		'단품목록
						For i=0 to prdPrcList.length-1
							gsshopPrice = prdPrcList.get(i).prdPrcSalePrc	'판매가격
						Next
					Set prdPrcList = nothing

					Set prdAttrInfoList = strObj.data.prdAttrInfoList		'상품 속성 정보 리스트
						For i=0 to prdAttrInfoList.length-1
							outmallSellyn = ""
							outmallOptCode	= prdAttrInfoList.get(i).attrPrdCd			'GS속성상품코드
							outmallOptName	= prdAttrInfoList.get(i).prdAttrVal1		'속성값1
							tenOptcd		= prdAttrInfoList.get(i).supAttrPrdCd						'협력사속성상품코드
							If prdAttrInfoList.get(i).attrStCd = "Y" Then				'속성판매상태 | 판매대기 : N (MD승인 후 최종 승인나기 전), 판매중 : Y, 판매종료 : E, 일시품절 : T, 완전판매종료 : D (판종사유코드가 31)
								outmallSellyn = "Y"
							Else
								outmallSellyn = "N"
							End If
							outmalllimitno	= prdAttrInfoList.get(i).attrOrdPsblQty		'주문가능수량

							strSql = " INSERT INTO db_item.dbo.tbl_OutMall_regedoption"
							strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outMallSellyn, outmalllimityn, outMallLimitNo)"
							strSql = strSql & " VALUES ("&iitemid
							If i = 0 AND outmallOptName = "공통" Then
								strSql = strSql & " ,'0000'"
							Else
								strSql = strSql & " ,'"& tenOptcd &"'"
							End If
							strSql = strSql & " ,'"&CMALLNAME&"'"
							strSql = strSql & " ,'"&outmallOptCode&"'"
							strSql = strSql & " ,'"&html2DB(outmallOptName)&"'"
							strSql = strSql & " ,'"&outmallSellyn&"'"
							strSql = strSql & " ,'Y'"
							strSql = strSql & " ,"&outmalllimitno
							strSql = strSql & ")"
							dbget.Execute strSql, AssignedRow
						Next
					Set prdAttrInfoList = nothing
					strSql = ""
					strSql = strSql & " UPDATE R " & VbCRLF
					strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VbCRLF
					strSql = strSql & " ,lastconfirmdate = getdate()"& VbCRLF
					strSql = strSql & " ,GSShopSellYn = '"& gsshopsellYn &"' "& VbCRLF
					strSql = strSql & " ,GSShopPrice = '"& gsshopPrice &"' "& VbCRLF
					strSql = strSql & " ,GSShopGoodNo = CASE WHEN isNull(GSShopGoodNo, '') = '' THEN '"& GSShopGoodNo &"' ELSE GSShopGoodNo END "& VbCRLF
					If GSShopGoodNo <> "" Then
						strSql = strSql & " ,GSShopstatCD = 3 "& VbCRLF
					End If

					If iVal = "REG" Then
						strSql = strSql & " ,GSShopLastUpdate = GETDATE() "& VbCRLF
						strSql = strSql & " ,accFailCnt = 0 "& VbCRLF
						strSql = strSql & " ,GSShopRegdate = GETDATE()  "& VbCRLF
					End If
					strSql = strSql & " FROM db_item.dbo.tbl_gsshop_regItem R " & VbCRLF
					strSql = strSql & " JOIN ( " & VbCRLF
					strSql = strSql & " 	SELECT R.itemid,count(*) as CNT "
					strSql = strSql & " 	, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
					strSql = strSql & "		FROM db_item.dbo.tbl_gsshop_regItem R " & VbCRLF
					strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro " & VbCRLF
					strSql = strSql & " 		on R.itemid = Ro.itemid"   & VbCRLF
					strSql = strSql & " 		and Ro.mallid = '"&CMALLNAME&"'"   & VbCRLF
					strSql = strSql & "         and Ro.itemid = "&iitemid & VbCRLF
					strSql = strSql & " 	GROUP BY R.itemid "   & VbCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid " & VbCRLF
					dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||성공(조회)"
		        End If
			Set strObj = nothing
			fnGSShopItemView = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-CHKSTAT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 전시상품 판매가 수정
Public Function fnGSShopNewPrice(iitemid, istrParam, imustprice, byRef iErrStr)
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
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(상품가격)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

				    strSql = ""
	    			strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem  " & VbCRLF
	    			strSql = strSql & "	SET GSShopLastUpdate=getdate() " & VbCRLF
	    			strSql = strSql & "	, GSShopPrice = " & imustprice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " Where itemid='" & iitemid & "'"& VbCRLF
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(상품가격)"
					fnGSShopNewPrice = True
		        End If
			Set strObj = nothing
		Else
			fnGSShopNewPrice = False
			iErrStr = "ERR||"&iitemid&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-PRICE-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New상품 옵션 추가 및 수량 수정
Function fnGSShopNewOPTSuEdit(iitemid, strParam, byRef iErrStr)
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
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(옵션 추가 및 수량 수정)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					Set attrPrdlist = strObj.attr
						For i=0 to attrPrdlist.length-1
							tenOptcd = attrPrdlist.get(i).supAttrPrdCd
							gsOptcd = attrPrdlist.get(i).attrPrdCd
							If attrPrdlist.length-1 = 0 AND tenOptcd = "0000" Then	'단품이라면
								sqlStr = ""
								sqlStr = sqlStr & "UPDATE db_item.dbo.tbl_OutMall_regedoption SET "
								sqlStr = sqlStr & "outmalllimitno =  "
								sqlStr = sqlStr & "Case WHEN B.limityn = 'Y' and B.limitno - B.limitsold <= 5 THEN '0'  "
								sqlStr = sqlStr & "	 WHEN B.limityn = 'Y' and B.limitno - B.limitsold > 5 THEN B.limitno - B.limitsold - 5 "
								sqlStr = sqlStr & "	 WHEN B.limityn = 'N' THEN '999' END "
								sqlStr = sqlStr & "FROM db_item.dbo.tbl_OutMall_regedoption A  "
								sqlStr = sqlStr & "JOIN db_item.dbo.tbl_item B on A.itemid = B.itemid "
								sqlStr = sqlStr & "WHERE A.itemid = '"&iitemid&"' and A.itemoption = '"&tenOptcd&"' and A.mallid = 'gsshop' "
								dbget.Execute sqlStr
							Else
								sqlStr = ""
								sqlStr = sqlStr & " IF Exists(SELECT * FROM db_item.dbo.tbl_OutMall_regedoption where itemid='"&iitemid&"' and itemoption = '"&tenOptcd&"' and mallid = 'gsshop') "
								sqlStr = sqlStr & " BEGIN"& VbCRLF
								sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_OutMall_regedoption " & VbCRLF
								sqlStr = sqlStr & " SET outmalllimitno = " & VbCRLF
								sqlStr = sqlStr & " Case WHEN optlimityn = 'Y' AND optlimitno - optlimitsold <= 5 THEN '0' " & VbCRLF
								sqlStr = sqlStr & " 	 WHEN optlimityn = 'Y' AND optlimitno - optlimitsold > 5 THEN optlimitno - optlimitsold - 5" & VbCRLF
								sqlStr = sqlStr & " 	 WHEN optlimityn = 'N' THEN '999' End" & VbCRLF
								sqlStr = sqlStr & " ,outmalllimityn = B.optlimityn " & VbCRLF
								sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption A  " & VbCRLF
								sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_option B on A.itemid = B.itemid and A.itemoption = B.itemoption " & VbCRLF
								sqlStr = sqlStr & " WHERE B.itemid = '"&iitemid&"' and B.itemoption = '"&tenOptcd&"' and A.mallid = 'gsshop' "
								sqlStr = sqlStr & " END ELSE "
								sqlStr = sqlStr & " BEGIN"& VbCRLF
								sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
								sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
								sqlStr = sqlStr & " SELECT itemid, itemoption, 'gsshop', '"&gsOptcd&"', optionname, optsellyn, optlimityn, " & VBCRLF
								sqlStr = sqlStr & " Case WHEN optlimityn = 'Y' AND optlimitno - optlimitsold <= 5 THEN '0' " & VBCRLF
								sqlStr = sqlStr & " 	 WHEN optlimityn = 'Y' AND optlimitno - optlimitsold > 5 THEN optlimitno - optlimitsold - 5 " & VBCRLF
								sqlStr = sqlStr & " 	 WHEN optlimityn = 'N' THEN '999' End " & VBCRLF
								sqlStr = sqlStr & " , '0', getdate() " & VBCRLF
								sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option " & VBCRLF
								sqlStr = sqlStr & " WHERE itemid= '"&iitemid&"' " & VBCRLF
								sqlStr = sqlStr & " and itemoption = '"& tenOptcd &"' "
								sqlStr = sqlStr & " END "
								dbget.Execute sqlStr
							End If
						Next
					Set attrPrdlist = nothing
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(옵션 추가 및 수량 수정)"
					fnGSShopNewOPTSuEdit = True
		        End If
			Set strObj = nothing
		Else
			fnGSShopNewOPTSuEdit = False
			iErrStr = "ERR||"&iitemid&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-OPTSuEdit-002]"
		End If
	Set objXML= nothing
End Function

'New상품 옵션 상태변경
Function fnGSShopNewOPTSellEdit(iitemid,strParam,byRef iErrStr)
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
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(옵션상태변경)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

	                ' sqlStr = ""
					' sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_OutMall_regedoption " & VbCRLF
					' sqlStr = sqlStr & " SET outmallsellyn = " & VbCRLF
					' sqlStr = sqlStr & " Case WHEN (B.isusing = 'N' OR B.optsellyn = 'N') THEN 'N' " & VbCRLF
					' sqlStr = sqlStr & " 	 WHEN (B.optlimityn = 'Y' AND B.optlimitno - B.optlimitsold <= 5) THEN 'N'  " & VbCRLF
					' sqlStr = sqlStr & " 	 WHEN (A.outmallOptName <> B.optionname) THEN 'N'  " & VbCRLF
					' sqlStr = sqlStr & " 	 WHEN (isNull(B.itemoption, '') = '') THEN 'N'  " & VbCRLF
					' sqlStr = sqlStr & " ELSE 'Y' END " & VbCRLF
					' sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption A  " & VbCRLF
					' sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item_option B on A.itemid = B.itemid and A.itemoption = B.itemoption " & VbCRLF
					' sqlStr = sqlStr & " WHERE A.itemid = '"&iitemid&"' and A.mallid = 'gsshop' "
				    ' dbget.Execute sqlStr
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(옵션상태변경)"
		        End If
			Set strObj = nothing
			fnGSShopNewOPTSellEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-OPTSellEdit-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 상품명 변경 수행 함수
Public Function fnGSShopChgNewItemname(iitemid,strParam,byRef iErrStr)
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

				If Instr(resultmsg, "등록된 정보와 동일한 요청입니다") > 0 Then
					resultcode = "True"
				End If

				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(상품명 변경)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem " & VbCRLF
					strSql = strSql & " SET regitemname = B.itemname "& VbCRLF
					strSql = strSql & " FROM db_item.dbo.tbl_gsshop_regItem A "& VbCRLF
					strSql = strSql & " JOIN db_item.dbo.tbl_item B on A.itemid = B.itemid "& VbCRLF
					strSql = strSql & " WHERE A.itemid='" & iitemid & "'"& VbCRLF
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(상품명 변경)"
		        End If
			Set strObj = nothing
			fnGSShopChgNewItemname = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-NMEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 이미지 변경 수행 함수
Function fnGSShopNewImageEdit(iitemid, strParam, iErrStr, ichgImageNm)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj, strSql
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
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(이미지수정)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regitem "
					strSql = strSql & " SET regimageName='"&ichgImageNm&"'"
					strSql = strSql & " WHERE itemid = '"& iitemid &"' "
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(이미지수정)"
		        End If
			Set strObj = nothing
			fnGSShopNewImageEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-IMGEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 안전인증정보 수정 함수
Function fnGSShopNewSafeCertEdit(iitemid, strParam, iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewSafeCertEdit = False

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
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(전안법수정)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(전안법수정)"
		        End If
			Set strObj = nothing
			fnGSShopNewSafeCertEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-SAFEEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 상품정보 변경 수행 함수
Function fnGSShopNewItemInfoEdit(iitemid, strParam, iErrStr)
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
				prdCd		= strObj.prdCd
				supPrdCd	= strObj.supPrdCd
				resultmsg	= replaceMsg(strObj.msg)
				resultcode	= strObj.success
				supCd 		= strObj.supCd

				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(상품정보)"
				Else
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET GSShopLastUpdate = getdate()"
					strSql = strSql & " FROM db_item.dbo.tbl_gsshop_regitem R" & VbCrlf
					strSql = strSql & " where R.itemid = " & iitemid
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(상품정보)"
		        End If
			Set strObj = nothing
			fnGSShopNewItemInfoEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-INFOEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 전시상품 설명 수정
Function fnGSShopNewContentsEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewContentsEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=euc-kr"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(상품설명수정)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(상품설명수정)"

					strSql = ""
					strSql = strSql & " UPDATE db_item.[dbo].[tbl_gsshop_regitem] "
					strSql = strSql & " SET isRegHtmlErr = NULL "
					strSql = strSql & " WHERE itemid = '"& itemid &"' "
					dbget.Execute strSql
		        End If
			Set strObj = nothing
			fnGSShopNewContentsEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-CONTEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New 정부고시항목 수정
Function fnGSShopNewInfodivEdit(iitemid,strParam,byRef iErrStr)
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
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(정부고시항목수정)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(정부고시항목 수정)"
		        End If
			Set strObj = nothing
			fnGSShopNewInfodivEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-DIVEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'매장정보 수정
Function fnGSShopCateEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopCateEdit = False
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
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(매장정보수정)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(매장정보수정)"
		        End If
			Set strObj = nothing
			fnGSShopCateEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop 결과 분석 중에 오류가 발생했습니다.[ERR-CATEEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function getGSShopDivCodeView()
    Dim objXML, xmlDOM, strRst, resultList, strObj
    Dim strSql, iRbody, i, j
	Dim lrgCd, lrgNm, midCd, midNm, smCd, smNm, dtlCd, dtlNm, isusing, infoDivArr, infoDivNameArr, infoDiv, infoDivName, safeGbnCd
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", "http://realapi.gsshop.com/SupSendPrdClsInfo.gs", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				Set resultList = strObj.resultList
					strSql = ""
					strSql = " DELETE FROM db_temp.[dbo].[tbl_gsshopMng_metaInfo] "
					dbget.Execute(strSql)
					For i=0 to resultList.length-1
						lrgCd = resultList.get(i).lrgClsCd
						lrgNm = resultList.get(i).lrgClsNm
						midCd = resultList.get(i).midClsCd
						midNm = resultList.get(i).midClsNm
						smCd = resultList.get(i).smlClsCd
						smNm = resultList.get(i).smlClsNm
						dtlCd = resultList.get(i).dtlClsCd
						dtlNm = resultList.get(i).dtlClsNm
						isusing = resultList.get(i).useYn
						safeGbnCd = resultList.get(i).safeCertTgtGbnCd
						infoDivArr = Split(resultList.get(i).govPublsPrdGrpCd, "$")
						infoDivNameArr = Split(resultList.get(i).govPublsPrdGrpNm, "$")

						strSql = ""
						strSql = strSql & " IF EXISTS (SELECT TOP 1 dtlCd FROM db_temp.[dbo].[tbl_gsshopMng_category] WHERE lrgCd = '"& lrgCd &"' and midCd = '"& midCd &"' and smCd = '"& smCd &"' and dtlCd = '"& dtlCd &"' ) "
						strSql = strSql & " 	BEGIN "
						strSql = strSql & " 		UPDATE db_temp.[dbo].[tbl_gsshopMng_category] "
						strSql = strSql & " 		SET isusing = '"& isusing &"'"
						strSql = strSql & " 		,safeGbnCd = '"& safeGbnCd &"'"
						strSql = strSql & " 		WHERE lrgCd = '"& lrgCd &"' and midCd = '"& midCd &"' and smCd = '"& smCd &"' and dtlCd = '"& dtlCd &"' "
						strSql = strSql & " 	END "
						strSql = strSql & " ELSE "
						strSql = strSql & " 	BEGIN "
						strSql = strSql & " 		INSERT INTO db_temp.[dbo].[tbl_gsshopMng_category] "
						strSql = strSql & " 		(lrgCd, lrgNm, midCd, midNm, smCd, smNm, dtlCd, dtlNm, isusing) VALUES "
						strSql = strSql & " 		('"&lrgCd&"', '"&lrgNm&"', '"&midCd&"', '"&midNm&"', '"&smCd&"', '"&smNm&"', '"&dtlCd&"', '"&dtlNm&"', '"&isusing&"') "
						strSql = strSql & " 	END "
						dbget.Execute(strSql)
						 For j = 0 to Ubound(infoDivArr)
						 	strSql = ""
						 	strSql = strSql & " INSERT INTO db_temp.[dbo].[tbl_gsshopMng_metaInfo]  "
						 	strSql = strSql & " (dtlCd, infoDiv, infoDivName) VALUES "
						 	strSql = strSql & " ('"&dtlCd&"', '"&infoDivArr(j)&"', '"&infoDivNameArr(j)&"') "
						 	dbget.Execute(strSql)
						 Next

						' rw "대분류코드 : " & resultList.get(i).lrgClsCd
						' rw "대분류명 : " & resultList.get(i).lrgClsNm
						' rw "중분류코드 : " & resultList.get(i).midClsCd
						' rw "중분류명 : " & resultList.get(i).midClsNm
						' rw "소분류코드 : " & resultList.get(i).smlClsCd
						' rw "소분류명 : " & resultList.get(i).smlClsNm
						' rw "세분류코드 : " & resultList.get(i).dtlClsCd
						' rw "세분류명 : " & resultList.get(i).dtlClsNm
						' rw "사용여부 : " & resultList.get(i).useYn
						' rw "안전인증대상구분코드 : " & resultList.get(i).safeCertTgtGbnCd
						' rw "유효기관관리대상여부 : " & resultList.get(i).validTermMngYn
						' rw "단위가격표시대상여부 : " & resultList.get(i).unitPrcYn
						' rw "세금유형코드 : " & resultList.get(i).taxTypSelCd
						' rw "정보고시그룹코드 : " & resultList.get(i).govPublsPrdGrpCd
						' rw "정보고시그룹명 : " & resultList.get(i).govPublsPrdGrpNm
						' rw "--------------"
					Next
					rw "완료"
				Set resultList = nothing
			Set strObj = nothing
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function getGSShopCateCodeView(iDate)
    Dim objXML, xmlDOM, strRst, resultList, strObj
    Dim strSql, iRbody, i, j, strParam
	Dim fromDtm, toDtm, sectId, sectNm, sectLrgId, sectLrgNm, sectMidId, sectMidNm, sectDtlId, sectDtlNm, prdDispYn, shopAttrCd, shopAttrNm
	fromDtm = replace(iDate, "-", "") & "000000"
	toDtm	= replace(dateadd("d", 6, iDate), "-", "") & "235959"
	strParam = "?fromDtm=" & fromDtm & "&toDtm=" & toDtm
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", "http://realapi.gsshop.com/DispSectInfo.gs" & strParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		Dim v : v= 0
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
rw iRbody
			Set strObj = JSON.parse(iRbody)
				Set resultList = strObj.resultList
					rw iDate & " ~ " & dateadd("d", 6, iDate)
					For i=0 to resultList.length-1
						sectId		= resultList.get(i).sectId				'진열매장아이디
						sectNm		= resultList.get(i).sectNm				'진열매장명
						sectLrgId	= resultList.get(i).sectLrgId			'대분류매장아이디
						sectLrgNm	= resultList.get(i).sectLrgNm			'대분류매장명
						sectMidId	= resultList.get(i).sectMidId			'중분류매장아이디
						sectMidNm	= resultList.get(i).sectMidNm			'중분류매장명
						sectDtlId	= resultList.get(i).sectDtlId			'소분류매장아이디
						sectDtlNm	= resultList.get(i).sectDtlNm			'소분류매장명
						prdDispYn	= resultList.get(i).prdDispYn			'상품진열가능여부
						shopAttrCd	= resultList.get(i).shopAttrCd			'매장속성코드
						shopAttrNm	= resultList.get(i).shopAttrNm			'매장속성명

						If prdDispYn = "Y" Then
							' If shopAttrNm = "일반매장" and (Trim(sectId) <> Trim(sectDtlId)) Then
							' 	rw "진열매장아이디 : " & sectId
							' 	rw "진열매장명 : " & sectNm
							' 	rw "대분류매장아이디 : " & sectLrgId
							' 	rw "대분류매장명 : " & sectLrgNm
							' 	rw "중분류매장아이디 : " & sectMidId
							' 	rw "중분류매장명 : " & sectMidNm
							' 	rw "소분류매장아이디 : " & sectDtlId
							' 	rw "소분류매장명 : " & sectDtlNm
							' 	rw "상품진열가능여부 : " & prdDispYn
							' 	rw "매장속성코드 : " & shopAttrCd
							' 	rw "매장속성명 : " & shopAttrNm
							' 	rw "-------------"
							' End If

							If sectLrgNm = "10x10" Then
								v = v + 1
								strSql = ""
								strSql = strSql & " IF NOT EXISTS (SELECT * FROM db_temp.dbo.tbl_gsshop_category WHERE catekey = '"& sectId &"'  ) "
								strSql = strSql & " 	BEGIN "
								strSql = strSql & " 		INSERT INTO db_temp.dbo.tbl_gsshop_category (CateKey, categbn, L_NAME, L_CODE, M_NAME, M_CODE, S_NAME, S_CODE, D_NAME, D_CODE, lastupdate, isusing) VALUES "
								strSql = strSql & " 		('"& sectId &"', 'B', '"&sectLrgNm&"', '"&sectLrgId&"', '"&sectMidNm&"', '"&sectMidId&"', '"&sectDtlNm&"', '"&sectDtlId&"', NULL, NULL, GETDATE(), 'Y') "
								strSql = strSql & " 	END "
								dbget.Execute strSql

								rw "진열매장아이디 : " & sectId
								rw "진열매장명 : " & sectNm
								rw "대분류매장아이디 : " & sectLrgId
								rw "대분류매장명 : " & sectLrgNm
								rw "중분류매장아이디 : " & sectMidId
								rw "중분류매장명 : " & sectMidNm
								rw "소분류매장아이디 : " & sectDtlId
								rw "소분류매장명 : " & sectDtlNm
								rw "상품진열가능여부 : " & prdDispYn
								rw "매장속성코드 : " & shopAttrCd
								rw "매장속성명 : " & shopAttrNm
								rw "-------------" & v
							End IF
						End IF
					Next
					response.write "<input type='button' value='Go' onclick=location.replace('/outmall/gsshop/gsshopActproc.asp?act=CateCodeView&sDate="&dateadd("d", 7, iDate)&"');>"
				Set resultList = nothing
			Set strObj = nothing
		End If
	Set objXML = Nothing
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################

'################################################# 각 기능 별 파라메터 정리 ###############################################
'품절 파라메타
Function getGSShopSellynParameter(iitemid, ichgSellYn)
	Dim strRst, strSql
	strRst = ""
	strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
	strRst = strRst & "&modGbn=S"														'(*)수정구분 S : 판매상태 수정
	strRst = strRst & "&regId="&COurRedId												'(*)등록자
	'상품기본(prdBaseInfo)
	strRst = strRst & "&supPrdCd="&iitemid												'(*)협력사상품코드
	strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
	'상품가격(prdPrc)

	If ichgSellYn = "Y" Then
		strRst = strRst & "&saleEndDtm=29991231235959"									'(*)판매종료일시 | 상품을 중단(판매종료)하려면 중단시점의 판매종료일시를 입력합니다.
	ElseIf (ichgSellYn = "N") Then
		strRst = strRst & "&saleEndDtm="&FormatDate(now(), "00000000000000")			'(*)판매종료일시 | 상품을 중단(판매종료)하려면 중단시점의 판매종료일시를 입력합니다.
	End If
	'strRst = strRst & "&attrSaleEndStModYn=N"											'(*)속성판매종료상태수정설정 | 속성구분(S) 상품판매상태를 변경할 때 사용하는 항목으로, 상품마스터 종료 및 해제 시 속성상품의 상태도 함께 종료 및 해제하려면 Y, 상품마스터와 속성 별도로 상태변경 동작 시엔 N
	strRst = strRst & "&attrSaleEndStModYn=Y"											'(*)속성판매종료상태수정설정 | 속성구분(S) 상품판매상태를 변경할 때 사용하는 항목으로, 상품마스터 종료 및 해제 시 속성상품의 상태도 함께 종료 및 해제하려면 Y, 상품마스터와 속성 별도로 상태변경 동작 시엔 N

	getGSShopSellynParameter = strRst
End Function

'조회 파라메타
Function getGSShopItemViewParameter(iitemid)
	Dim strRst, strSql
	strRst = ""
	strRst = strRst & "supPrdCd="&iitemid												'(*)협력사상품(속성)코드
	strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
	strRst = strRst & "&searchItmCd=PRC,ATTR"											'조회항목코드 설정 | 선택 항목으로 입력값 없을 때는 상품기본정보만 조회 '원하는 항목을 추가적으로 조회시 , (콤마)를 사용하여 전달 (전체항목조회:ALL, 상품명:NM, 가격:PRC, 속성:ATTR, 배송:DLV, 확장정보:ADD, 구성정보:CMP, 매장:SECT, 안전인증:SAFE, 정보고시:GOV, 사양정보:SPEC, 기술서:HTML) ex) 전체항목조회 ALL 기본정보, 가격, 속성, 매장 조회시 : PRC,ATTR,SECT 속성, 상품명, 정보고시, 기술서 : ATTR,NM,GOV,HTML
	getGSShopItemViewParameter = strRst
End Function

Public Function getGSShopPriceParameter(iitemid, mustprice)
	Dim strRst, strSql
	Dim sellcash, orgprice, buycash
	Dim GetTenTenMargin

	'전송 구분 및 반복리스트 건수
	strRst = ""
	strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
	strRst = strRst & "&modGbn=P"														'(*)수정구분 P : 가격 수정
	strRst = strRst & "&regId="&COurRedId												'(*)등록자
	strRst = strRst & "&regSubjCd=SUP"													'(*)등록주체코드 | 엠디가 수정한 경우 : MD, 협력사가 수정한 경우 : SUP
	'상품기본(prdBaseInfo)
	strRst = strRst & "&supPrdCd="&iitemid												'(*)협력사상품코드
	strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
	strRst = strRst & "&subSupCd="&COurCompanyCode										'(*)하위협력사코드 | 하위협력사가 없으면 supCd와 같은값 입력
	'상품가격(prdPrc)
	strRst = strRst & "&prdPrcValidStrDtm="&FormatDate(now(), "00000000000000")			'(*)유효시작일시
	strRst = strRst & "&prdPrcValidEndDtm=29991231235959"								'(*)유효종료일시
	strRst = strRst & "&prdPrcSalePrc="&mustprice										'(*)판매가격
	'strRst = strRst & "&prdPrcPrchPrc="												'(SYS)매입가격 | (SYS는 저희쪽에서 자동으로 생성해주는 코드 및 값을 말합니다. Null로 보내주시면 됩니다.)
	strRst = strRst & "&prdPrcSupGivRtamtCd=01"											'(*)협력사지급율/액코드 | 01 : 액
	strRst = strRst & "&prdPrcSupGivRtamt="&getGSShopSuplyPrice_update(MustPrice)		'(*)협력사지급율/액 | 기본값 : 판매가*(1-0.12)
	getGSShopPriceParameter = strRst
End Function

Public Function getGSShopItemnameParameter(iitemid, byref iitemname)
	Dim strRst, chgname, strSql, brandName
	strSql = ""
	strSql = strSql & " SELECT TOP 1 r.itemid, r.GSShopGoodNo, i.ItemName, c.socname_kor "
	strSql = strSql & "	FROM db_item.dbo.tbl_gsshop_regItem r "
	strSql = strSql & "	JOIN db_item.dbo.tbl_item i on r.itemid = i.itemid "
	strSql = strSql & "	JOIN db_user.dbo.tbl_user_c as c on i.makerid = c.userid "
	strSql = strSql & "	WHERE r.regitemname is Not NULL "
	strSql = strSql & "	and (r.GSShopStatCd=3 OR r.GSShopStatCd=7) "
	strSql = strSql & "	and r.GSShopGoodNo is Not Null "
	strSql = strSql & " and	i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.Eof Then
		iitemname = rsget("ItemName")
		brandName = Trim(rsget("socname_kor"))
	End If
	rsget.close

	chgname = ""
'	chgname = "[텐바이텐]"&replace(iitemname,"'","")			'최초 상품명 앞에 [텐바이텐] 이라고 붙임
	chgname = replace(iitemname,"'","")							'최초 상품명 앞에 [텐바이텐] 삭제

	If Left(iitemname, Len(Trim(brandName)) + 2) = "[" & brandName & "]" Then
	ElseIf (Left(iitemname, len(brandName)) <> brandName) Then
		chgname = brandName & " " & Replace(iitemname,"'","")		'[텐바이텐] 문구 삭제 / 브랜드한글명 붙임 / 2020-07-30 위로 원복
	End If

	chgname = replace(chgname,"&#8211;","-")
	chgname = replace(chgname,"~","-")
	chgname = replace(chgname,"&","＆")
	chgname = replace(chgname,"<","[")
	chgname = replace(chgname,">","]")
	chgname = replace(chgname,"%","프로")
	chgname = replace(chgname,"+","%2B")
	chgname = replace(chgname,":","%3A")
	chgname = replace(chgname,"[무료배송]","")
	chgname = replace(chgname,"[무료 배송]","")

	strRst = ""
	strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
	strRst = strRst & "&modGbn=N"														'(*)수정구분 N : 노출상품명 수정
	strRst = strRst & "&regId="&COurRedId												'(*)등록자
	'상품기본(prdBaseInfo)
	strRst = strRst & "&supPrdCd="&iitemid												'(*)협력사상품코드
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
