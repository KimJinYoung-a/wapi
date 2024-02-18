<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/aspJSON1.17.asp"-->
<!-- #include virtual="/outmall/ezwel/ezwelItemcls.asp"-->
<!-- #include virtual="/outmall/ezwel/incezwelFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, oEzwel, failCnt, chgSellYn, arrRows, skipItem, sellgubun, getMustprice, chkXML, ezwelGoodno, isItemIdChk
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, chkparam, optReset, optString
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
chkXML			= request("chkXML")
failCnt			= 0

Select Case action
	Case "mafcList", "brandList"	isItemIdChk = "N"
	Case Else				isItemIdChk = "Y"
End Select

If isItemIdChk = "Y" Then
	If itemid="" or itemid="0" Then
		response.write "<script>alert('상품번호가 없습니다.')</script>"
		response.end
	ElseIf Not(isNumeric(itemid)) Then
		response.write "<script>alert('잘못된 상품번호입니다.')</script>"
		response.end
	Else
		'정수형태로 변환
		itemid=CLng(getNumeric(itemid))
	End If
End If
'######################################################## Ezwel API ########################################################
If action = "EditSellYn" Then								'상태변경
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem

		If chgSellYn = "N" Then
			sellgubun = "SellN"
		Else
			chgSellYn = "AdminOK"
			sellgubun = "SellY"
		End If

		strParam = ""
		strParam = oEzwel.FOneItem.getEzwelItemRegXML(sellgubun, chkXML)
		getMustprice = ""
		getMustprice = oEzwel.FOneItem.fngetMustPrice()
		Call EzwelOneItemEditSellyn(itemid, oEzwel.FOneItem.FEzwelGoodNo, iErrStr, strParam, getMustprice, chgSellYn, "all", oEzwel.FOneItem.FLimityn, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitsold, chkXML)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oEzwel = nothing
ElseIf action = "EDIT" Then									'상품 수정
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정가능한 상품이 아닙니다."
		Else
			strParam = ""
			chkparam = ""
			iErrStr = ""
			optReset = "N"
			optString = "all"
			'*********************************************************************************************************************************************************
			'2014-11-06 김진영 | dev_Comment
			'API가 전송되는 족족 상품옵션을 인식하지 않음 | 등록된 옵션카운트가 크다면 10x10에서 옵션 삭제한 것이 살아있음
			'결국 이지웰의 옵션사용안함으로 돌리면 옵션이 초기화 됨을 발견
			'추가 : 두번 API전송시 많은 확률로 에러가 뜸 | 아마 이지웰페어 DB쪽 상품가격 수정하는 데 뭔가 걸려있는 듯 함..
			'		따라서 우선 이런 상품은 품절로
			strSql = ""
			strSql = strSql &  "SELECT top 1 r.itemid, i.optioncnt, r.regedoptcnt "
			strSql = strSql & " FROM db_item.dbo.tbl_item as i "
			strSql = strSql & " join db_etcmall.dbo.tbl_ezwel_regitem as r on i.itemid=r.itemid "
			strSql = strSql & " WHERE i.itemid=" & itemid
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If not rsget.EOF Then
				If CInt(rsget("optioncnt")) > 0 Then
					If CInt(rsget("optioncnt")) <> CInt(rsget("regedoptcnt")) Then
						optReset = "Y"
						optString = "optMustN"
					End If
				End If
			End If
			rsget.Close

			If (oEzwel.FOneItem.FmaySoldOut = "Y") OR (oEzwel.FOneItem.IsSoldOutLimit5Sell) OR (optReset = "Y") OR (oEzwel.FOneItem.IsMayLimitSoldout = "Y") Then
				If optReset = "Y" Then
					strParam = oEzwel.FOneItem.getEzwelItemRegXML("MustNotOpt", chkXML)
				Else
					strParam = oEzwel.FOneItem.getEzwelItemRegXML("SellN", chkXML)
				End If
				chgSellYn = "N"
			Else
				strParam = oEzwel.FOneItem.getEzwelItemRegXML("SellY", chkXML)
				chgSellYn = "Y"
			End If

			getMustprice = ""
			getMustprice = oEzwel.FOneItem.fngetMustPrice()
			Call EzwelOneItemEdit(itemid, oEzwel.FOneItem.FEzwelGoodNo, iErrStr, strParam, getMustprice, chgSellYn, optString, oEzwel.FOneItem.FLimityn, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitsold, chkXML, oEzwel.FOneItem.FezwelSellYn)
		End If

		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If

		If InStr(iErrStr, "[재판매]") = 0 Then
			Call EzwelItemChkstat(itemid, iErrStr, oEzwel.FOneItem.FEzwelGoodNo)
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("ezwel", itemid, SumErrStr)
			Call SugiQueLogInsert("ezwel", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("ezwel", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oEzwel = nothing
ElseIf action = "REG" Then									'상품 등록
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelNotRegOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf oEzwel.FOneItem.FdepthCode = "0" Then
			iErrStr = "ERR||"&itemid&"||카테고리 매칭 여부 확인하세요."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_ezwel_regItem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_ezwel_regItem "
	        strSql = strSql & " 	(itemid, regdate, reguserid, ezwelstatCD, regitemname)"
	        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oEzwel.FOneItem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oEzwel.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oEzwel.FOneItem.getEzwelItemRegXML("Reg", chkXML)
				Call EzwelItemReg(itemid, strParam, iErrStr, oEzwel.FOneItem.FSellCash, oEzwel.FOneItem.getEzwelSellYn, oEzwel.FOneItem.FLimityn, oEzwel.FOneItem.FLimitNo, oEzwel.FOneItem.FLimitSold, html2db(oEzwel.FOneItem.FItemName), oEzwel.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||옵션검사 실패"
			End If
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oEzwel = nothing
ElseIf action = "CHKSTAT" Then									'상태 조회
	ezwelGoodno = getEzwelGoodno(itemid)
	If (ezwelGoodno = "") Then
		iErrStr = "ERR||"&itemid&"||조회 가능한 상품이 아닙니다."
	Else
		Call EzwelItemChkstat(itemid, iErrStr, ezwelGoodno)
	End If

	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ElseIf action = "REG2" Then										'상품등록
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelNotRegOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf oEzwel.FOneItem.FdepthCode = "0" Then
			iErrStr = "ERR||"&itemid&"||카테고리 매칭 여부 확인하세요."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Outmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"', '"&CMALLNAME&"' "
			dbget.execute strSql
			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oEzwel.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oEzwel.FOneItem.getEzwelItemRegJson("N")
				Call EzwelItemNewReg(itemid, strParam, iErrStr, oEzwel.FOneItem.FSellCash, oEzwel.FOneItem.getEzwelSellYn, oEzwel.FOneItem.FLimityn, oEzwel.FOneItem.FLimitNo, oEzwel.FOneItem.FLimitSold, html2db(oEzwel.FOneItem.FItemName), oEzwel.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||옵션검사 실패"
			End If
		End If

		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	SET oEzwel = nothing

	If failCnt = 0 Then
		SET oEzwel = new cEzwel
			oEzwel.FRectItemID	= itemid
			oEzwel.getEzwelEditOneItem
			If (oEzwel.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||조회 상품이 아닙니다."
			Else
				strParam = ""
				Call EzwelItemNewChkstat(itemid, oEzwel.FOneItem.FEzwelGoodNo, oEzwel.FOneItem.FLimitYN, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitSold, iErrStr)
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		SET oEzwel = nothing
	End If

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("ezwel", itemid, SumErrStr)
		Call SugiQueLogInsert("ezwel", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("ezwel", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=REG2&itemid=2343355
ElseIf action = "CHKSTAT2" Then									'상세조회
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||조회 상품이 아닙니다."
		Else
			strParam = ""
			Call EzwelItemNewChkstat(itemid, oEzwel.FOneItem.FEzwelGoodNo, oEzwel.FOneItem.FLimitYN, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitSold, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oEzwel = nothing
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=CHKSTAT2&itemid=2930817
ElseIf action = "PRICE" Then									'가격변경
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		Else
			strParam = oEzwel.FOneItem.getEzwelItemPriceJson()
			Call EzwelItemPrice(itemid, strParam, oEzwel.FOneItem.MustPrice, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oEzwel = nothing
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=PRICE&itemid=2930817
ElseIf action = "EditSellYn2" Then								'상태변경
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		Else
			Call EzwelNewEditSellyn(itemid, chgSellYn, oEzwel.FOneItem.FEzwelGoodNo, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oEzwel = nothing
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=EditSellYn2&itemid=2930817&chgSellYn=Y
ElseIf action = "EDITOPT" Then									'옵션변경
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		Else
			strParam = oEzwel.FOneItem.getEzwelItemOptionJson()
			Call EzwelItemOption(itemid, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oEzwel = nothing
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=EDITOPT&itemid=2930817
ElseIf action = "EDIT2" Then									'상품수정
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem

		If oEzwel.FResultCount = 0 Then
	    	failCnt = failCnt + 1
			iErrStr = "ERR||"&itemid&"||수정가능한 상품이 아닙니다."
		Else
			If (oEzwel.FOneItem.FmaySoldOut = "Y") OR (oEzwel.FOneItem.IsMayLimitSoldout = "Y") OR (oEzwel.FOneItem.FLimityn = "Y" AND (oEzwel.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) Then
				Call EzwelNewEditSellyn(itemid, "N", oEzwel.FOneItem.FEzwelGoodNo, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
				rw "판매상태 수정"
				response.flush
				response.clear
			Else
			'##################################### 판매상품 조회 시작 #######################################
				Call EzwelItemNewChkstat(itemid, oEzwel.FOneItem.FEzwelGoodNo, oEzwel.FOneItem.FLimitYN, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitSold, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
				rw "판매상품 조회"
				response.flush
				response.clear
			'##################################### 판매상품 조회 끝 #########################################

			'##################################### 가격 수정 시작 ###########################################
				If failCnt = 0 Then
					strParam = ""
					strParam = oEzwel.FOneItem.getEzwelItemPriceJson()
					Call EzwelItemPrice(itemid, strParam, oEzwel.FOneItem.MustPrice, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "가격 수정"
					response.flush
					response.clear
				End If
			'##################################### 가격 수정 끝 #############################################

			'##################################### 상품 수정 시작 ###########################################
				If failCnt = 0 Then
					strParam = ""
					strParam = oEzwel.FOneItem.getEzwelItemRegJson("EDIT")
					Call EzwelItemNewEdit(itemid, strParam, iErrStr, html2db(oEzwel.FOneItem.FItemName), oEzwel.FOneItem.FbasicimageNm)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "상품 수정"
					response.flush
					response.clear
				End If
			'####################################### 상품 수정 끝 ###########################################

			'##################################### 판매 상태 수정 시작 ######################################
				If failCnt = 0 Then
					Call EzwelNewEditSellyn(itemid, "Y", oEzwel.FOneItem.FEzwelGoodNo, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "판매상태 수정"
					response.flush
					response.clear
				End If
			'##################################### 판매 상태 수정 끝 ########################################

			'##################################### 판매상품 조회 시작 #######################################
				If failCnt = 0 Then
					Call EzwelItemNewChkstat(itemid, oEzwel.FOneItem.FEzwelGoodNo, oEzwel.FOneItem.FLimitYN, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitSold, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "판매상품 조회"
					response.flush
					response.clear
				End If
			'##################################### 판매상품 조회 끝 #########################################
			End If
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("ezwel", itemid, SumErrStr)
			Call SugiQueLogInsert("ezwel", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("ezwel", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oEzwel = nothing
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=EDIT2&itemid=2648990
ElseIf action = "brandList" Then								'브랜드조회
	Call fnEzwelBrandList()
	response.end
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=brandList
ElseIf action = "mafcList" Then									'제조사조회
	Call fnEzwelMafcList()
	response.end
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=mafcList
End If

If iErrStr <> "" Then
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str = '"&iErrStr&"<br>' + str " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
					"</script>"
End If
'###################################################### ezwel API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
