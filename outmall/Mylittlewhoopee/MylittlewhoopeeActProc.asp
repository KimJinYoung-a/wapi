<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/Mylittlewhoopee/MylittlewhoopeeItemcls.asp"-->
<!-- #include virtual="/outmall/Mylittlewhoopee/incMylittlewhoopeeFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, oMylittlewhoopee, failCnt, chgSellYn, arrRows, skipItem, getMustprice, oService, oOperation, chkXML
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, ccd, isItemIdChk, getfarmGoodno
Dim i, chgImageNm, mayOptSoldOut, endItemErrMsgReplace
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
chkXML			= request("chkXML")
ccd				= request("ccd")
failCnt			= 0
Select Case action
	Case "nvstorefarmCommonCode", "CATE", "CATEDETAIL"	isItemIdChk = "N"
	Case Else						isItemIdChk = "Y"
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
'######################################################## 스토어팜 API ########################################################
If action = "REG" Then											'상품 기본 정보 + 옵션 등록
	SET oMylittlewhoopee = new CMylittlewhoopee
		oMylittlewhoopee.FRectItemID	= itemid
		oMylittlewhoopee.getMylittlewhoopeeNotRegOneItem
		If (oMylittlewhoopee.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf oMylittlewhoopee.FOneItem.FAPIaddImg <> "Y" Then
			iErrStr = "ERR||"&itemid&"||이미지 부터 업로드 하세요."
		ElseIf oMylittlewhoopee.FOneItem.FNvstorefarmid <> "0" Then
			iErrStr = "ERR||"&itemid&"||스토어팜 상품과 중복입니다."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_Mylittlewhoopee_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_Mylittlewhoopee_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, MylittlewhoopeestatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oMylittlewhoopee.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oMylittlewhoopee.FOneitem.checkTenItemOptionValid Then
				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oMylittlewhoopee.FOneitem.getMylittlewhoopeeItemRegXML(oService, oOperation, "")

				getMustprice = ""
				getMustprice = oMylittlewhoopee.FOneItem.MustPrice()
				Call fnMylittlewhoopeeItemReg(itemid, strParam, iErrStr, getMustprice, oMylittlewhoopee.FOneItem.getMylittlewhoopeeSellYn, oMylittlewhoopee.FOneItem.FLimityn, oMylittlewhoopee.FOneItem.FLimitNo, oMylittlewhoopee.FOneItem.FLimitSold, html2db(oMylittlewhoopee.FOneItem.FItemName), oMylittlewhoopee.FOneItem.FbasicimageNm, oService, oOperation, chkXML)
			Else
				iErrStr = "ERR||"&itemid&"||옵션검사 실패"
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
'------------------------------------------------------------ 상품 기본 정보 등록 ------------------------------------------------------------
			If failCnt = 0  Then
				If oMylittlewhoopee.FOneitem.FOptioncnt > 0 Then				'옵션수가 0이면 단품이므로 옵션 등록 X
					oService		= "ProductService"
					oOperation		= "ManageOption"

					getfarmGoodno = getMylittlewhoopeeGoodNo(itemid)
					If getfarmGoodno <> "" Then
						strParam = ""
						strParam = getMylittlewhoopeeOptionRegXML(itemid, getfarmGoodno, oService, oOperation)
						If strParam <> "X" Then
							Call fnMylittlewhoopeeOptionReg(itemid, strParam, iErrStr, oService, oOperation)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						End If
					Else
						failCnt = failCnt + 1
						iErrStr = "ERR||"&itemid&"||미등록상품입니다."
						SumErrStr = SumErrStr & iErrStr
					End If

					If failCnt > 0 Then
						'옵션 등록 중 오류시 삭제 API 이용
						oService		= "ProductService"
						oOperation		= "DeleteProduct"

						strParam = ""
						strParam = getMylittlewhoopeeDeleteParameter(getfarmGoodno, oService, oOperation)
						Call fnMylittlewhoopeeDelete(itemid, strParam, iErrStr, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							SumErrStr = SumErrStr & iErrStr
							SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
							SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
							CALL Fn_AcctFailTouch("Mylittlewhoopee", itemid, SumErrStr)
							Call SugiQueLogInsert("Mylittlewhoopee", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
							iErrStr = "ERR||"&itemid&"||"&SumErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
							SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
							CALL Fn_AcctFailTouch("Mylittlewhoopee", itemid, SumErrStr)
							iErrStr = "ERR||"&itemid&"||옵션API 오류, 삭제처리"
						End If
					Else
						SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
						Call SugiQueLogInsert("Mylittlewhoopee", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
						iErrStr = "OK||"&itemid&"||"&SumOKStr
					End If
				Else
					SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
					Call SugiQueLogInsert("Mylittlewhoopee", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
					iErrStr = "OK||"&itemid&"||"&SumOKStr
				End If
			Else
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("Mylittlewhoopee", itemid, SumErrStr)
				Call SugiQueLogInsert("Mylittlewhoopee", action, itemid, Split(iErrStr,"||")(0), SumErrStr, session("ssBctID"))
				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			End If
		End If
	SET oMylittlewhoopee = nothing
ElseIf action = "REGITEM" Then									'상품 기본 정보 등록
	SET oMylittlewhoopee = new CMylittlewhoopee
		oMylittlewhoopee.FRectItemID	= itemid
		oMylittlewhoopee.getMylittlewhoopeeNotRegOneItem
		If (oMylittlewhoopee.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf oMylittlewhoopee.FOneItem.FAPIaddImg <> "Y" Then
			iErrStr = "ERR||"&itemid&"||이미지 부터 업로드 하세요."
		ElseIf oMylittlewhoopee.FOneItem.FNvstorefarmid <> "0" Then
			iErrStr = "ERR||"&itemid&"||스토어팜 상품과 중복입니다."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_Mylittlewhoopee_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_Mylittlewhoopee_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, MylittlewhoopeestatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oMylittlewhoopee.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oMylittlewhoopee.FOneitem.checkTenItemOptionValid Then
				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oMylittlewhoopee.FOneitem.getMylittlewhoopeeItemRegXML(oService, oOperation, "")

				getMustprice = ""
				getMustprice = oMylittlewhoopee.FOneItem.MustPrice()
				Call fnMylittlewhoopeeItemReg(itemid, strParam, iErrStr, getMustprice, oMylittlewhoopee.FOneItem.getMylittlewhoopeeSellYn, oMylittlewhoopee.FOneItem.FLimityn, oMylittlewhoopee.FOneItem.FLimitNo, oMylittlewhoopee.FOneItem.FLimitSold, html2db(oMylittlewhoopee.FOneItem.FItemName), oMylittlewhoopee.FOneItem.FbasicimageNm, oService, oOperation, chkXML)
			Else
				iErrStr = "ERR||"&itemid&"||옵션검사 실패"
			End If
		End If
	SET oMylittlewhoopee = nothing
ElseIf action = "REGOPT" Then									'옵션 등록
	oService		= "ProductService"
	oOperation		= "ManageOption"

	getfarmGoodno = getMylittlewhoopeeGoodNo(itemid)
	If getfarmGoodno <> "" Then
		strParam = ""
		strParam = getMylittlewhoopeeOptionRegXML(itemid, getfarmGoodno, oService, oOperation)
		If strParam <> "X" Then
			Call fnMylittlewhoopeeOptionReg(itemid, strParam, iErrStr, oService, oOperation)
		End If
	Else
		iErrStr = "ERR||"&itemid&"||상품이 등록되지 않았습니다."
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("Mylittlewhoopee", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("Mylittlewhoopee", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "Image" Then									'이미지 등록
	SET oMylittlewhoopee = new CMylittlewhoopee
		oMylittlewhoopee.FRectItemID	= itemid
		oMylittlewhoopee.FRectGubun		= "IMG"
		oMylittlewhoopee.getMylittlewhoopeeNotRegOneItem
		If (oMylittlewhoopee.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf oMylittlewhoopee.FOneItem.FNvstorefarmid <> "0" Then
			iErrStr = "ERR||"&itemid&"||스토어팜 상품과 중복입니다."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_Mylittlewhoopee_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_Mylittlewhoopee_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, MylittlewhoopeestatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oMylittlewhoopee.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oMylittlewhoopee.FOneitem.checkTenItemOptionValid Then
				oService		= "ImageService"
				oOperation		= "UploadImage"

				strParam = ""
				strParam = oMylittlewhoopee.FOneitem.getMylittlewhoopeeImageRegXML(oService, oOperation)
				chgImageNm = oMylittlewhoopee.FOneItem.getBasicImage
				Call fnMylittlewhoopeeImageReg(itemid, strParam, iErrStr, chgImageNm, oService, oOperation)
			Else
				iErrStr = "ERR||"&itemid&"||옵션검사 실패"
			End If
		End If
	SET oMylittlewhoopee = nothing
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("Mylittlewhoopee", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("Mylittlewhoopee", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKOPT" Then									'옵션 조회
	oService		= "ProductService"
	oOperation		= "GetOption"

	strParam = ""
	strParam = getMylittlewhoopeeOptionSearchParameter(getMylittlewhoopeeGoodNo(itemid), oService, oOperation)
	Call fnMylittlewhoopeeOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("Mylittlewhoopee", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("Mylittlewhoopee", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKSTAT" Then									'상품 조회
	oService		= "ProductService"
	oOperation		= "GetProduct"

	strParam = ""
	strParam = getMylittlewhoopeeItemSearchParameter(getMylittlewhoopeeGoodNo(itemid), oService, oOperation)
	Call fnMylittlewhoopeeItemSearch(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("Mylittlewhoopee", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("Mylittlewhoopee", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDIT" Then										'상품조회 -> 옵션조회 -> 상품수정 -> 옵션수정 순
	SET oMylittlewhoopee = new CMylittlewhoopee
		oMylittlewhoopee.FRectItemID	= itemid
		oMylittlewhoopee.getMylittlewhoopeeEditOneItem

		If (oMylittlewhoopee.FResultCount < 1) OR (oMylittlewhoopee.FOneItem.FMylittlewhoopeeGoodNo = "") Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
			failCnt = failCnt + 1
		Else
			If oMylittlewhoopee.FOneItem.FOptioncnt > 0 Then
				mayOptSoldOut = oMylittlewhoopee.FOneItem.IsMayLimitSoldout
			End If

			If (oMylittlewhoopee.FOneItem.FMaySoldOut = "Y") OR (oMylittlewhoopee.FOneItem.IsSoldOutLimit5Sell) OR (mayOptSoldOut = "Y") OR (oMylittlewhoopee.FOneItem.FLimitYn = "Y" AND oMylittlewhoopee.FOneItem.getiszeroWonSoldOut(itemid) = "Y") Then
				oService		= "ProductService"
				oOperation		= "ChangeProductSaleStatus"

				strParam = ""
				strParam = getMylittlewhoopeeSellynParameter(oMylittlewhoopee.FOneItem.FMylittlewhoopeeGoodNo, "N", oService, oOperation)
				Call fnMylittlewhoopeeSellyn(itemid, "N", strParam, iErrStr, oService, oOperation)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (oMylittlewhoopee.FOneItem.FMylittlewhoopeeSellYn = "N" AND oMylittlewhoopee.FOneItem.IsSoldOutLimit5Sell = False) Then
					oService		= "ProductService"
					oOperation		= "ChangeProductSaleStatus"

					strParam = ""
					strParam = getMylittlewhoopeeSellynParameter(oMylittlewhoopee.FOneItem.FMylittlewhoopeeGoodNo, "Y", oService, oOperation)
					Call fnMylittlewhoopeeSellyn(itemid, "Y", strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
	'################################################ 0.상품 수정부터(ReturnCostReason 신규 필드 때문..) ####################
				If oMylittlewhoopee.FOneItem.isImageChanged = True Then
					chgImageNm = oMylittlewhoopee.FOneItem.getBasicImage
				Else
					chgImageNm = "N"
				End If

				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oMylittlewhoopee.FOneitem.getMylittlewhoopeeItemRegXML(oService, oOperation, "Y")
				getMustprice = ""
				getMustprice = oMylittlewhoopee.FOneItem.MustPrice()
				Call fnMylittlewhoopeeItemEDIT(itemid, strParam, iErrStr, getMustprice, oMylittlewhoopee.FOneItem.getMylittlewhoopeeSellYn, oMylittlewhoopee.FOneItem.FLimityn, oMylittlewhoopee.FOneItem.FLimitNo, oMylittlewhoopee.FOneItem.FLimitSold, oMylittlewhoopee.FOneItem.FItemName, chgImageNm, oService, oOperation)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
	'################################################ 1.옵션 가져오기(regedoption때문) #######################################
				oService		= "ProductService"
				oOperation		= "GetOption"

				strParam = ""
				strParam = getMylittlewhoopeeOptionSearchParameter(oMylittlewhoopee.FOneItem.FMylittlewhoopeeGoodNo, oService, oOperation)
				Call fnMylittlewhoopeeOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
	'##########################################################################################################################
	'################################################ 2.이미지 변경시 이미지 재업로드 #########################################
				If chgImageNm <> "N" Then
					oService		= "ImageService"
					oOperation		= "UploadImage"

					strParam = ""
					strParam = oMylittlewhoopee.FOneitem.getMylittlewhoopeeImageRegXML(oService, oOperation)
					chgImageNm = oMylittlewhoopee.FOneItem.getBasicImage
					Call fnMylittlewhoopeeImageReg(itemid, strParam, iErrStr, chgImageNm, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
	'##########################################################################################################################
	'############################################## 3.실패횟수가 0일때 상품수정 ###############################################
				If failCnt = "0" Then
					oService		= "ProductService"
					oOperation		= "ManageProduct"

					strParam = ""
					strParam = oMylittlewhoopee.FOneitem.getMylittlewhoopeeItemRegXML(oService, oOperation, "Y")
					getMustprice = ""
					getMustprice = oMylittlewhoopee.FOneItem.MustPrice()
					Call fnMylittlewhoopeeItemEDIT(itemid, strParam, iErrStr, getMustprice, oMylittlewhoopee.FOneItem.getMylittlewhoopeeSellYn, oMylittlewhoopee.FOneItem.FLimityn, oMylittlewhoopee.FOneItem.FLimitNo, oMylittlewhoopee.FOneItem.FLimitSold, (oMylittlewhoopee.FOneItem.FItemName), chgImageNm, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
	'##########################################################################################################################
	'############################################## 4.옵션수정 ################################################################
					oService		= "ProductService"
					oOperation		= "ManageOption"

					strParam = ""
					strParam = getMylittlewhoopeeOptionRegXML(itemid, oMylittlewhoopee.FOneItem.FMylittlewhoopeeGoodno, oService, oOperation)
					If strParam <> "X" Then
						Call fnMylittlewhoopeeOptionReg(itemid, strParam, iErrStr, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
	'##########################################################################################################################
	'################################################ 5.옵션 가져오기 #######################################
					oService		= "ProductService"
					oOperation		= "GetOption"

					strParam = ""
					strParam = getMylittlewhoopeeOptionSearchParameter(oMylittlewhoopee.FOneItem.FMylittlewhoopeeGoodNo, oService, oOperation)
					Call fnMylittlewhoopeeOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

	'##########################################################################################################################
				endItemErrMsgReplace = replace(SumErrStr, "OK||"&itemid&"||", "")
				endItemErrMsgReplace = replace(SumErrStr, "ERR||"&itemid&"||", "")

				If (Instr(endItemErrMsgReplace, "대분류는 변경할 수 없습니다") > 0) OR (Instr(endItemErrMsgReplace, "대분류는변경할수없습니다") > 0) OR (Instr(endItemErrMsgReplace, "옵션의옵션가/사용여부항목을") > 0) OR (Instr(endItemErrMsgReplace, "옵션의 옵션가/사용여부 항목을") > 0) OR (Instr(endItemErrMsgReplace, "옵션값항목에콤마(,)는") > 0) OR (Instr(endItemErrMsgReplace, "옵션값 항목에 콤마(,)는") > 0) Then
					oService		= "ProductService"
					oOperation		= "DeleteProduct"

					strParam = ""
					strParam = getMylittlewhoopeeDeleteParameter(oMylittlewhoopee.FOneItem.FMylittlewhoopeeGoodNo, oService, oOperation)
					Call fnMylittlewhoopeeDelete(itemid, strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						failCnt = 0
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			End If
		End If
		'OK던 ERR이던 editQuecnt에 + 1을 시킴..
		'스케줄링에서 editQuecnt ASC, i.lastupdate DESC로 중복을 막자
		strSql = ""
		strSql = strSql & " UPDATE [db_etcmall].[dbo].tbl_Mylittlewhoopee_regitem SET " & VBCRLF
		strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
		strSql = strSql & " ,Mylittlewhoopeelastupdate = getdate()  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("Mylittlewhoopee", itemid, SumErrStr)
			Call SugiQueLogInsert("Mylittlewhoopee", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_Mylittlewhoopee_regitem SET " & VBCRLF
			strSql = strSql & " accFailcnt = 0  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("Mylittlewhoopee", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oMylittlewhoopee = nothing
ElseIf action = "EditSellYn" Then								'상태 변경
	oService		= "ProductService"
	oOperation		= "ChangeProductSaleStatus"

	strParam = ""
	strParam = getMylittlewhoopeeSellynParameter(getMylittlewhoopeeGoodNo(itemid), chgSellYn, oService, oOperation)
	Call fnMylittlewhoopeeSellyn(itemid, chgSellYn, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("Mylittlewhoopee", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("Mylittlewhoopee", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "DEL" Then										'상품 삭제
	oService		= "ProductService"
	oOperation		= "DeleteProduct"

	strParam = ""
	strParam = getMylittlewhoopeeDeleteParameter(getMylittlewhoopeeGoodNo(itemid), oService, oOperation)
	Call fnMylittlewhoopeeDelete(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("Mylittlewhoopee", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("Mylittlewhoopee", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "nvstorefarmCommonCode" Then					'공통코드 검색
	If ccd = "GetAddressBookList" Then
		strParam = ""
		strParam = getAddressBookList(ccd)
	End If
'	Call fnAuctionCommonCode(ccd, strParam)
ElseIf action = "CATE" Then										'카테고리 검색
	ccd = "GetAllCategoryList"
	strParam = getAllCategoryList(ccd)
ElseIf action = "CATEDETAIL" Then								'카테고리 상세조회
	ccd = "GetCategoryInfo"
rw "아래 catekey를 1000단위로 끊어서 실행..sp에서 주석삭제해야함..response.end 처리.."
'response.end
'	strSql = ""
'	strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_nvstorefarm_certInfo] "
'	dbget.Execute strSql

'	strSql = ""
'	strSql = strSql & " SELECT CateKey "
'	strSql = strSql & " FROM db_etcmall.dbo.tbl_nvstorefarm_category "
'	'strSql = strSql & " WHERE CateKey <= 50001777 "
'	'strSql = strSql & " WHERE CateKey > 50001777 AND CateKey <= 50002805 "
'	'strSql = strSql & " WHERE CateKey > 50002805 AND CateKey <= 50003911 "
'	'strSql = strSql & " WHERE CateKey > 50003911 AND CateKey <= 50005605 "
'	strSql = strSql & " WHERE CateKey > 50005605 "
'	strSql = strSql & " GROUP BY CateKey "
'	strSql = strSql & " ORDER BY CateKey "
'	rsget.Open strSql,dbget,1
'	If not rsget.Eof Then
'		arrRows = rsget.getRows()
'	End If
'	rsget.Close
'response.end

	strSql = "exec [db_etcmall].[dbo].[usp_API_Nvstorefarm_CatekeyList_Get]"
	rsget.CursorLocation = adUseClient
	rsget.CursorType = adOpenStatic
	rsget.LockType = adLockOptimistic
	rsget.Open strSql, dbget
	If Not(rsget.EOF or rsget.BOF) Then
		arrRows = rsget.getRows()
	End If
	rsget.Close

	For i = 0 To Ubound(arrRows, 2)
		strParam = getCategoryInfo(ccd, arrRows(0, i))
	Next
	rw "OK"
End If

If iErrStr <> "" Then
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str += '"&iErrStr&"<br>' " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
					"</script>"
End If
'###################################################### 스토어팜 API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
