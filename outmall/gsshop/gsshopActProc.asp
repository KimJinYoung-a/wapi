<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/gsshop/gsshopItemcls.asp"-->
<!-- #include virtual="/outmall/gsshop/incGSShopFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, oGSShop, failCnt, chgSellYn, chkXml, prdDescErrYN, chgImageNm
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, isItemIdChk
Dim sDate
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
chkXml			= request("chkXml")
sDate			= request("sDate")
failCnt			= 0

Select Case action
	Case "DivCodeView", "CateCodeView"	isItemIdChk = "N"
	Case Else							isItemIdChk = "Y"
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
'######################################################## GSShop API ########################################################
If action = "EditSellYn" Then								'상태변경
	strParam = ""
	strParam = getGSShopSellynParameter(itemid, chgSellYn)
	Call fnGSShopNewSellyn(itemid, chgSellYn, strParam, iErrStr)
	'response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("gsshop", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKSTAT" Then								'상품 조회
	strParam = ""
	strParam = getGSShopItemViewParameter(itemid)
	Call fnGSShopItemView(itemid, strParam, iErrStr, "")
	'response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("gsshop", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "REG" Then									'상품등록
	SET oGSShop = new CGSShop
		oGSShop.FRectItemID	= itemid
		oGSShop.getGSShopNotRegOneItem
		strSql = ""
		strSql = "EXEC [db_etcmall].[dbo].[usp_API_Outmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"', '"&CMALLNAME&"' "
		dbget.Execute strSql

	    If (oGSShop.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			If oGSShop.FOneItem.FDivcode = "" Then		'만약 상품분류 매칭을 안 한 카테고리 상품이라면..
				iErrStr = "ERR||"&itemid&"||상품분류 매칭을 하지 않은 상품번호"
			Else
				'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
				If oGSShop.FOneItem.checkTenItemOptionValid Then
					strParam = ""
					strParam = oGSShop.FOneItem.getGSShopItemNewRegParameter(1)
					If chkXml = "Y" Then
						response.write strParam
					End If
					CALL fnGSShopNewItemReg(itemid, strParam, iErrStr, oGSShop.FOneItem.FSellCash, oGSShop.FOneItem.getGSShopSellYn, oGSShop.FOneItem.FLimityn, oGSShop.FOneItem.FLimitNo, oGSShop.FOneItem.FLimitSold, html2db(oGSShop.FOneItem.FItemName), oGSShop.FOneItem.FbasicimageNm)
				Else
					iErrStr = "ERR||"&itemid&"||옵션검사 실패"
				End If
			End If
		End If

		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If

		strParam = ""
		strParam = getGSShopItemViewParameter(itemid)
		Call fnGSShopItemView(itemid, strParam, iErrStr, "REG")
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("gsshop", itemid, SumErrStr)
			Call SugiQueLogInsert("gsshop", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("gsshop", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oGSShop = nothing
ElseIf action = "REG2" Then									'상품등록(필수 값을 확인 해주세요. : prdDescdHtmlDescdExplnCntnt(상품등록))..오류 관련
	strSql = ""
	strSql = strSql & " SELECT COUNT(*) as cnt "
	strSql = strSql & " FROM db_item.dbo.tbl_gsshop_regitem "
	strSql = strSql & " WHERE itemid = " & itemid
	strSql = strSql & " and GSShopStatCd = 1 "
	strSql = strSql & " and lastErrStr like '%prdDescdHtmlDescdExplnCntnt%' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("cnt") > 0 Then
		prdDescErrYN = "Y"
	Else
		prdDescErrYN = "N"
	End If
	rsget.Close

	If prdDescErrYN = "Y" Then
		SET oGSShop = new CGSShop
			oGSShop.FRectItemID	= itemid
			oGSShop.getGSShopNotRegOneItem
			strSql = ""
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Outmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"', '"&CMALLNAME&"' "
			dbget.Execute strSql

			If (oGSShop.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
			Else
				If oGSShop.FOneItem.FDivcode = "" Then		'만약 상품분류 매칭을 안 한 카테고리 상품이라면..
					iErrStr = "ERR||"&itemid&"||상품분류 매칭을 하지 않은 상품번호"
				Else
					'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
					If oGSShop.FOneItem.checkTenItemOptionValid Then
						strParam = ""
						strParam = oGSShop.FOneItem.getGSShopItemNewRegParameter(2)
						If chkXml = "Y" Then
							response.write strParam
						End If
						CALL fnGSShopNewItemReg(itemid, strParam, iErrStr, oGSShop.FOneItem.FSellCash, oGSShop.FOneItem.getGSShopSellYn, oGSShop.FOneItem.FLimityn, oGSShop.FOneItem.FLimitNo, oGSShop.FOneItem.FLimitSold, html2db(oGSShop.FOneItem.FItemName), oGSShop.FOneItem.FbasicimageNm)
					Else
						iErrStr = "ERR||"&itemid&"||옵션검사 실패"
					End If
				End If
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			strParam = ""
			strParam = getGSShopItemViewParameter(itemid)
			Call fnGSShopItemView(itemid, strParam, iErrStr, "REG")
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("gsshop", itemid, SumErrStr)
				Call SugiQueLogInsert("gsshop", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				Call SugiQueLogInsert("gsshop", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
				iErrStr = "OK||"&itemid&"||"&SumOKStr
			End If
		SET oGSShop = nothing
	Else
		iErrStr = "ERR||"&itemid&"||오류 상품 등록가능한 상품이 아닙니다."
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("gsshop", "REG", itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "PRICE" Then
	SET oGSShop = new CGSShop
		oGSShop.FRectItemID	= itemid
		oGSShop.getGSShopEditOneItem
		If oGSShop.FResultCount > 0 Then
			strParam = ""
			mustPrice = ""
			mustPrice = oGSShop.FOneItem.MustPrice()
			strParam = getGSShopPriceParameter(itemid, mustPrice)
			If strParam = "" Then
				response.write "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다."
			Else
				Call fnGSShopNewPrice(itemid, strParam, mustPrice, iErrStr)
				'response.write iErrStr
				If LEFT(iErrStr, 2) <> "OK" Then
					CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
				End If
				Call SugiQueLogInsert("gsshop", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
			end if
		else
			response.write "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다.[1]"
		end if
ElseIf action = "IMAGE" Then
	SET oGSShop = new CGSShop
		oGSShop.FRectItemID	= itemid
		oGSShop.getGSShopEditOneItem
		If oGSShop.FResultCount > 0 Then
			strParam = ""
			strParam = oGSShop.FOneItem.getGSShopImageEditParameter()
			chgImageNm = oGSShop.FOneItem.getBasicImage
			Call fnGSShopNewImageEdit(itemid, strParam, iErrStr, chgImageNm)
			'response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("gsshop", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oGSShop = nothing
ElseIf action = "SAFECERT" Then
	SET oGSShop = new CGSShop
		oGSShop.FRectItemID	= itemid
		oGSShop.getGSShopEditOneItem
		If oGSShop.FResultCount > 0 Then
			strParam = ""
			strParam = oGSShop.FOneItem.getGSShopSafeCertEditParameter()
			Call fnGSShopNewSafeCertEdit(itemid, strParam, iErrStr)
			'response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("gsshop", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oGSShop = nothing
ElseIf action = "EDITINFO" Then
	SET oGSShop = new CGSShop
		oGSShop.FRectItemID	= itemid
		oGSShop.getGSShopEditOneItem
		If oGSShop.FResultCount > 0 Then
			strParam = ""
			strParam = oGSShop.FOneItem.getGSShopItemEditParameter()
			Call fnGSShopNewItemInfoEdit(itemid, strParam, iErrStr)
			'response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("gsshop", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oGSShop = nothing
ElseIf action = "OPTSU" Then
	SET oGSShop = new CGSShop
		oGSShop.FRectItemID	= itemid
		oGSShop.getGSShopEditOneItem
		If oGSShop.FResultCount > 0 Then
			strParam = ""
			strParam = oGSShop.FOneItem.getGSShopOptParameter()
			Call fnGSShopNewOPTSuEdit(itemid, strParam, iErrStr)
			'response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("gsshop", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oGSShop = nothing
ElseIf action = "CONTENT" Then
	SET oGSShop = new CGSShop
		oGSShop.FRectItemID	= itemid
		oGSShop.getGSShopEditOneItem
		If oGSShop.FResultCount > 0 Then
			strParam = ""
			strParam = oGSShop.FOneItem.getGSShopContentsEditParameter()
			Call fnGSShopNewContentsEdit(itemid, strParam, iErrStr)
			'response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("gsshop", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oGSShop = nothing
ElseIf action = "ITEMNAME" Then
	strParam = ""
	strParam = getGSShopItemnameParameter(itemid, iitemname)
	Call fnGSShopChgNewItemname(itemid, strParam, iErrStr)
	'response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("gsshop", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "INFODIV" Then								'정부고시항목
	SET oGSShop = new CGSShop
		oGSShop.FRectItemID	= itemid
		oGSShop.getGSShopEditOneItem
		If oGSShop.FResultCount > 0 Then
			strParam = ""
			strParam = oGSShop.FOneItem.getGSShopInfodivEditParameter()
			Call fnGSShopNewInfodivEdit(itemid, strParam, iErrStr)
'			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("gsshop", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oGSShop = nothing
ElseIf action = "EDITCATE" Then								'매장정보
	SET oGSShop = new CGSShop
		oGSShop.FRectItemID	= itemid
		oGSShop.getGSShopEditOneItem
		If oGSShop.FResultCount > 0 Then
			strParam = ""
			strParam = oGSShop.FOneItem.getGSShopCategoryEditParameter()
			Call fnGSShopCateEdit(itemid, strParam, iErrStr)
'			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("gsshop", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oGSShop = nothing
ElseIf action = "EDIT" Then
	SET oGSShop = new CGSShop
		oGSShop.FRectItemID	= itemid
		oGSShop.getGSShopEditOneItem
		If oGSShop.FResultCount > 0 Then
			'2023-07-13 김진영 하단 쿼리 제거
			' strSql = ""
			' strSql = strSql & " SELECT COUNT(*) as cnt FROM db_item.dbo.tbl_item_option WHERE itemid = '"& itemid &"' and isusing='Y' and optAddPrice > 0 "
			' rsget.CursorLocation = adUseClient
			' rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			' If rsget("cnt") > 0 Then
			' 	oGSShop.FOneItem.FmaySoldOut = "Y"
			' ElseIf oGSShop.FOneItem.FOptionCnt = 0 and oGSShop.FOneItem.FregedOptCnt > 0 Then
			' 	oGSShop.FOneItem.FmaySoldOut = "Y"
			' End If
			' rsget.Close
			If oGSShop.FOneItem.FOptionCnt = 0 and oGSShop.FOneItem.FregedOptCnt > 0 Then
				oGSShop.FOneItem.FmaySoldOut = "Y"
			End If

			'2014-12-02 18:49:00 김진영 추가 끝
			If (oGSShop.FOneItem.FmaySoldOut = "Y") OR (oGSShop.FOneItem.IsSoldOutLimit5Sell) OR (oGSShop.FOneItem.IsMayLimitSoldout = "Y") Then
				strParam = ""
				strParam = getGSShopSellynParameter(itemid, "N")
				Call fnGSShopNewSellyn(itemid, "N", strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				strParam = ""
				strParam = getGSShopItemViewParameter(itemid)
				Call fnGSShopItemView(itemid, strParam, iErrStr, "")
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				If failCnt = 0 Then
					strParam = ""
					mustPrice = ""
					mustPrice = oGSShop.FOneItem.MustPrice()
					If (mustPrice <> oGSShop.FOneItem.FGSShopPrice) Then
						strParam = getGSShopPriceParameter(itemid, mustPrice)
						If strParam = "" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다."
						Else
							Call fnGSShopNewPrice(itemid, strParam, mustPrice, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						End If
					End If
				End If
				'타임아웃 등으로 단품상품의 regedoption테이블에 입력이 안 되었을 경우
				' If oGSShop.FOneItem.FoptionCnt = 0 AND oGSShop.FOneItem.FIsNulltoTimeout = "Y" Then
				' 	If oGSShop.FOneItem.FLimitYn = "Y" Then
				' 		strSql = ""
				' 		strSql = strSql & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_outmall_regedoption where itemid='"&itemid&"' and itemoption = '0000' and mallid = 'gsshop') "
				' 		strSql = strSql & " BEGIN"& VbCRLF
				' 		strSql = strSql & " insert into db_item.dbo.tbl_outmall_regedoption (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate) values " & VBCRLF
				' 		strSql = strSql & " ('"&itemid&"', '0000', 'gsshop', '"&oGSShop.FOneItem.FGsshopGoodNo&"001', '단일상품', 'Y', 'Y', '220', '0', getdate()) " & VBCRLF
				' 		strSql = strSql & " END "
				' 	Else
				' 		strSql = ""
				' 		strSql = strSql & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_outmall_regedoption where itemid='"&itemid&"' and itemoption = '0000' and mallid = 'gsshop') "
				' 		strSql = strSql & " BEGIN"& VbCRLF
				' 		strSql = strSql & " insert into db_item.dbo.tbl_outmall_regedoption (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate) values " & VBCRLF
				' 		strSql = strSql & " ('"&itemid&"', '0000', 'gsshop', '"&oGSShop.FOneItem.FGsshopGoodNo&"001', '단일상품', 'Y', 'N', '999', '0', getdate()) " & VBCRLF
				' 		strSql = strSql & " END "
				' 	End If
				' 	dbget.Execute strSql
				' End If

				If failCnt = 0 Then
					'기본 정보 수정
					strParam = ""
					strParam = oGSShop.FOneItem.getGSShopItemEditParameter()
					Call fnGSShopNewItemInfoEdit(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				If failCnt = 0 Then
					'상품설명 수정
					strParam = ""
					strParam = oGSShop.FOneItem.getGSShopContentsEditParameter()
					Call fnGSShopNewContentsEdit(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				If failCnt = 0 Then
					If oGSShop.FOneItem.isImageChanged Then
						chgImageNm = oGSShop.FOneItem.getBasicImage
						Call fnGSShopNewImageEdit(itemid, strParam, iErrStr, chgImageNm)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				End If

				If failCnt = 0 Then
					'옵션 추가 및 재고 수정
					strParam = ""
					strParam = oGSShop.FOneItem.getGSShopOptParameter()
					Call fnGSShopNewOPTSuEdit(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				If failCnt = 0 Then
					'옵션 판매상태 수정
					strParam = ""
					strParam = oGSShop.FOneItem.getGSShopOptSellParameter()
					Call fnGSShopNewOPTSellEdit(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				If (failCnt = 0) AND (oGSShop.FOneItem.FGsshopSellYn = "N" AND oGSShop.FOneItem.IsSoldOut = False) AND (oGSShop.FOneItem.isOptNotMatch <> "Y") Then
					iErrStr = ""
					strParam = ""
					strParam = getGSShopSellynParameter(itemid, "Y")
					Call fnGSShopNewSellyn(itemid, "Y", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				If oGSShop.FOneItem.isOptNotMatch = "Y" Then
					strParam = ""
					strParam = getGSShopSellynParameter(itemid, "N")
					Call fnGSShopNewSellyn(itemid, "N", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
						strSql = "	DECLARE @Temp CHAR(1) " & _
									"	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = 'gsshop' AND itemid = '"& itemid &"') " & _
									"		BEGIN " & _
									"			INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid(itemid,mallgubun) VALUES('"& itemid &"','gsshop') " & _
									"		END	"
						dbget.execute strSql
					End If
				End If

				strParam = ""
				strParam = getGSShopItemViewParameter(itemid)
				Call fnGSShopItemView(itemid, strParam, iErrStr, "")
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'OK던 ERR이던 editQuecnt에 + 1을 시킴..
				'스케줄링에서 editQuecnt ASC, i.lastupdate DESC로 중복을 막자
				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regitem SET " & VBCRLF
				strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				'response.write "ERR||"&itemid&"||"&SumErrStr
				CALL Fn_AcctFailTouch("gsshop", itemid, SumErrStr)
				Call SugiQueLogInsert("gsshop", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regitem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				'response.write "OK||"&itemid&"||"&SumOKStr
				Call SugiQueLogInsert("gsshop", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))

				iErrStr = "OK||"&itemid&"||"&SumOKStr
			End If
		Else
			iErrstr = "ERR||"&itemid&"||전체 수정 가능한 상품이 아닙니다."
		End If
		'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=EDIT
	SET oGSShop = nothing
ElseIf action = "DivCodeView" Then
	call getGSShopDivCodeView()
ElseIf action = "CateCodeView" Then
	call getGSShopCateCodeView(sDate)
	response.end
End If

If iErrStr <> "" Then
	if (IsAutoScript) then
		response.write iErrStr
	else
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str += '"&iErrStr&"<br>' " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
					"</script>"
	end if
End If
'###################################################### GSShop API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
