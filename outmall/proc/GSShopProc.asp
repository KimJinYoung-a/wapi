<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/gsshop/gsshopItemcls.asp"-->
<!-- #include virtual="/outmall/gsshop/incGSShopFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, mallid, action, oGSShop, failCnt
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, chgImageNm
Dim jenkinsBatchYn, idx, lastErrStr
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
jenkinsBatchYn	= request("jenkinsBatchYn")
idx				= request("idx")
lastErrStr		= ""
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
'######################################################## GSShop API ########################################################
If mallid = "gsshop" Then
	If action = "REG" Then										'상품등록
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
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oGSShop = nothing
		'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=REG
	ElseIf action = "REG2" Then
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
				strSql = ""
				strSql = strSql & " UPDATE db_item.[dbo].[tbl_gsshop_regitem] "
				strSql = strSql & " SET isRegHtmlErr = 'Y' "
				strSql = strSql & " WHERE itemid = '"& itemid &"' "
				dbget.Execute strSql
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
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oGSShop = nothing
		'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=REG2
	ElseIf action = "SOLDOUT" Then								'품절처리
		strParam = ""
		strParam = getGSShopSellynParameter(itemid, "N")
		Call fnGSShopNewSellyn(itemid, "N", strParam, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
		End If
		'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=SOLDOUT
	ElseIf action = "CHKSTAT" Then								'상품 조회
		strParam = ""
		strParam = getGSShopItemViewParameter(itemid)
		Call fnGSShopItemView(itemid, strParam, iErrStr, "")
		'response.write iErrStr
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
		End If
		'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=CHKSTAT
	ElseIf action = "EDITINFO" Then								'기본수정
		SET oGSShop = new CGSShop
			oGSShop.FRectItemID	= itemid
			oGSShop.getGSShopEditOneItem
			If oGSShop.FResultCount > 0 Then
				strParam = ""
				strParam = oGSShop.FOneItem.getGSShopItemEditParameter()
				Call fnGSShopNewItemInfoEdit(itemid, strParam, iErrStr)
				lastErrStr = iErrStr
				response.write iErrStr
				If LEFT(iErrStr, 2) <> "OK" Then
					CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
				End If
			End If
		SET oGSShop = nothing
		'testURL : http://testwapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=EDITINFO
	ElseIf action = "PRICE" Then								'가격수정
		strParam = ""
		SET oGSShop = new CGSShop
			oGSShop.FRectItemID	= itemid
			oGSShop.getGSShopEditOneItem
			If oGSShop.FResultCount > 0 Then
				mustPrice = ""
				mustPrice = oGSShop.FOneItem.MustPrice()
				strParam = getGSShopPriceParameter(itemid, mustPrice)
				If strParam = "" Then
					lastErrStr = "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다."
					response.write "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다."
				Else
					Call fnGSShopNewPrice(itemid, strParam, mustPrice, iErrStr)
					lastErrStr = iErrStr
					response.write iErrStr
					If LEFT(iErrStr, 2) <> "OK" Then
						CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
					End If
				End If
			Else
				response.write "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다.[1]"
			end if
			'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=PRICE
	ElseIf action = "ITEMNAME" Then								'상품명
		strParam = ""
		strParam = getGSShopItemnameParameter(itemid, iitemname)
		Call fnGSShopChgNewItemname(itemid, strParam, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
		End If
		'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=ITEMNAME
	ElseIf action = "IMAGE" Then								'이미지(대표 및 썸네일)
		SET oGSShop = new CGSShop
			oGSShop.FRectItemID	= itemid
			oGSShop.getGSShopEditOneItem
			If oGSShop.FResultCount > 0 Then
				strParam = ""
				strParam = oGSShop.FOneItem.getGSShopImageEditParameter()
				Call fnGSShopImageEdit(itemid, strParam, iErrStr)
				lastErrStr = iErrStr
				response.write iErrStr
				If LEFT(iErrStr, 2) <> "OK" Then
					CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
				End If
			End If
		SET oGSShop = nothing
		'testURL : http://testwapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=IMAGE
	ElseIf action = "SAFECERT" Then								'전안법 수정
		SET oGSShop = new CGSShop
			oGSShop.FRectItemID	= itemid
			oGSShop.getGSShopEditOneItem
			If oGSShop.FResultCount > 0 Then
				strParam = ""
				strParam = oGSShop.FOneItem.getGSShopSafeCertEditParameter()
				Call fnGSShopNewSafeCertEdit(itemid, strParam, iErrStr)
				lastErrStr = iErrStr
				response.write iErrStr
				If LEFT(iErrStr, 2) <> "OK" Then
					CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
				End If
			End If
		SET oGSShop = nothing
		'testURL : http://testwapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=SAFECERT
	ElseIf action = "CONTENT" Then								'상품설명
		SET oGSShop = new CGSShop
			oGSShop.FRectItemID	= itemid
			oGSShop.getGSShopEditOneItem
			If oGSShop.FResultCount > 0 Then
				strParam = ""
				strParam = oGSShop.FOneItem.getGSShopContentsEditParameter()
				Call fnGSShopNewContentsEdit(itemid, strParam, iErrStr)
				lastErrStr = iErrStr
				response.write iErrStr
				If LEFT(iErrStr, 2) <> "OK" Then
					CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
				End If
			End If
		SET oGSShop = nothing
		'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=CONTENT
	ElseIf action = "INFODIV" Then								'정부고시항목
		SET oGSShop = new CGSShop
			oGSShop.FRectItemID	= itemid
			oGSShop.getGSShopEditOneItem
			If oGSShop.FResultCount > 0 Then
				strParam = ""
				strParam = oGSShop.FOneItem.getGSShopInfodivEditParameter()
				Call fnGSShopNewInfodivEdit(itemid, strParam, iErrStr)
				lastErrStr = iErrStr
				response.write iErrStr
				If LEFT(iErrStr, 2) <> "OK" Then
					CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
				End If
			End If
		SET oGSShop = nothing
		'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=INFODIV
	ElseIf action = "EDITCATE" Then								'전시매장
		SET oGSShop = new CGSShop
			oGSShop.FRectItemID	= itemid
			oGSShop.getGSShopEditOneItem
			If oGSShop.FResultCount > 0 Then
				strParam = ""
				strParam = oGSShop.FOneItem.getGSShopCategoryEditParameter()
				Call fnGSShopCateEdit(itemid, strParam, iErrStr)
				lastErrStr = iErrStr
				response.write iErrStr
				If LEFT(iErrStr, 2) <> "OK" Then
					CALL Fn_AcctFailTouch("gsshop", itemid, iErrStr)
				End If
			End If
		SET oGSShop = nothing
		'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=EDITCATE
	ElseIf action = "EDIT" Then									'가격&재고&옵션&상태 수정 | 상태 -> 가격 -> 옵션 순 수정
		SET oGSShop = new CGSShop
			oGSShop.FRectItemID	= itemid
			oGSShop.getGSShopEditOneItem
			If oGSShop.FResultCount > 0 Then
				strSql = ""
				strSql = strSql & " SELECT COUNT(*) as cnt FROM db_item.dbo.tbl_item_option WHERE itemid = '"& itemid &"' and isusing='Y' and optAddPrice > 0 "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If rsget("cnt") > 0 Then
					oGSShop.FOneItem.FmaySoldOut = "Y"
				ElseIf oGSShop.FOneItem.FOptionCnt = 0 and oGSShop.FOneItem.FregedOptCnt > 0 Then
					oGSShop.FOneItem.FmaySoldOut = "Y"
				End If
				rsget.Close
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

					If (failCnt = 0) AND (oGSShop.FOneItem.FGsshopSellYn = "N" AND oGSShop.FOneItem.IsSoldOut = False) AND (oGSShop.FOneItem.FOptNotMatch <> "Y") Then
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

					If oGSShop.FOneItem.FOptNotMatch = "Y" Then
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
					lastErrStr = "ERR||"&itemid&"||"&SumErrStr
					response.write "ERR||"&itemid&"||"&SumErrStr
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regitem SET " & VBCRLF
					strSql = strSql & " accFailcnt = 0  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql
					SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
					lastErrStr = "OK||"&itemid&"||"&SumOKStr
					response.write "OK||"&itemid&"||"&SumOKStr
				End If
			Else
				iErrstr = "ERR||"&itemid&"||전체 수정 가능한 상품이 아닙니다."
			End If
			'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=EDIT
		SET oGSShop = nothing
	End If
End If
'###################################################### GSShop API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
