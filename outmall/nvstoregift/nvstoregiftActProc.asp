<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/nvstoregift/nvstoregiftItemcls.asp"-->
<!-- #include virtual="/outmall/nvstoregift/incnvstoregiftFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, oNvstoregift, failCnt, chgSellYn, arrRows, skipItem, getMustprice, oService, oOperation, chkXML
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, ccd, isItemIdChk, getfarmGoodno
Dim i, chgImageNm, mayOptSoldOut, endItemErrMsgReplace
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
chkXML			= request("chkXML")
ccd				= request("ccd")
failCnt			= 0
Select Case action
	Case "nvstorefarmCommonCode"	isItemIdChk = "N"
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
	SET oNvstoregift = new CNvstoregift
		oNvstoregift.FRectItemID	= itemid
		oNvstoregift.getNvstoregiftNotRegOneItem
		If (oNvstoregift.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf oNvstoregift.FOneItem.FAPIaddImg <> "Y" Then
			iErrStr = "ERR||"&itemid&"||이미지 부터 업로드 하세요."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_nvstoregift_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstoregift_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, nvstoregiftstatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oNvstoregift.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oNvstoregift.FOneitem.checkTenItemOptionValid Then
				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oNvstoregift.FOneitem.getNvstoregiftItemRegXML(oService, oOperation, "")

				getMustprice = ""
				getMustprice = oNvstoregift.FOneItem.MustPrice()
				Call fnNvstoregiftItemReg(itemid, strParam, iErrStr, getMustprice, oNvstoregift.FOneItem.getNvstoregiftSellYn, oNvstoregift.FOneItem.FLimityn, oNvstoregift.FOneItem.FLimitNo, oNvstoregift.FOneItem.FLimitSold, html2db(oNvstoregift.FOneItem.FItemName), oNvstoregift.FOneItem.FbasicimageNm, oService, oOperation, chkXML)
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
				If oNvstoregift.FOneitem.FOptioncnt > 0 Then				'옵션수가 0이면 단품이므로 옵션 등록 X
					oService		= "ProductService"
					oOperation		= "ManageOption"

					getfarmGoodno = getNvstoregiftGoodNo(itemid)
					If getfarmGoodno <> "" Then
						strParam = ""
						strParam = getNvstoregiftOptionRegXML(itemid, getfarmGoodno, oService, oOperation)
						If strParam <> "X" Then
							Call fnNvstoregiftOptionReg(itemid, strParam, iErrStr, oService, oOperation)
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
						strParam = getNvstoregiftDeleteParameter(getfarmGoodno, oService, oOperation)
						Call fnNvstoregiftDelete(itemid, strParam, iErrStr, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							SumErrStr = SumErrStr & iErrStr
							SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
							SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
							CALL Fn_AcctFailTouch("nvstoregift", itemid, SumErrStr)
							Call SugiQueLogInsert("nvstoregift", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
							iErrStr = "ERR||"&itemid&"||"&SumErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
							SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
							CALL Fn_AcctFailTouch("nvstoregift", itemid, SumErrStr)
							iErrStr = "ERR||"&itemid&"||옵션API 오류, 삭제처리"
						End If
					Else
						SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
						Call SugiQueLogInsert("nvstoregift", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
						iErrStr = "OK||"&itemid&"||"&SumOKStr
					End If
				Else
					SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
					Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
					iErrStr = "OK||"&itemid&"||"&SumOKStr
				End If
			Else
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("nvstoregift", itemid, SumErrStr)
				Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), SumErrStr, session("ssBctID"))
				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			End If
		End If
	SET oNvstoregift = nothing
ElseIf action = "REGITEM" Then									'상품 기본 정보 등록
	SET oNvstoregift = new CNvstoregift
		oNvstoregift.FRectItemID	= itemid
		oNvstoregift.getNvstoregiftNotRegOneItem
		If (oNvstoregift.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf oNvstoregift.FOneItem.FAPIaddImg <> "Y" Then
			iErrStr = "ERR||"&itemid&"||이미지 부터 업로드 하세요."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_nvstoregift_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstoregift_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, nvstoregiftstatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oNvstoregift.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oNvstoregift.FOneitem.checkTenItemOptionValid Then
				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oNvstoregift.FOneitem.getNvstoregiftItemRegXML(oService, oOperation, "")

				getMustprice = ""
				getMustprice = oNvstoregift.FOneItem.MustPrice()
				Call fnNvstoregiftItemReg(itemid, strParam, iErrStr, getMustprice, oNvstoregift.FOneItem.getNvstoregiftSellYn, oNvstoregift.FOneItem.FLimityn, oNvstoregift.FOneItem.FLimitNo, oNvstoregift.FOneItem.FLimitSold, html2db(oNvstoregift.FOneItem.FItemName), oNvstoregift.FOneItem.FbasicimageNm, oService, oOperation, chkXML)
			Else
				iErrStr = "ERR||"&itemid&"||옵션검사 실패"
			End If
		End If
	SET oNvstoregift = nothing
ElseIf action = "REGOPT" Then									'옵션 등록
	oService		= "ProductService"
	oOperation		= "ManageOption"

	getfarmGoodno = getNvstoregiftGoodNo(itemid)
	If getfarmGoodno <> "" Then
		strParam = ""
		strParam = getNvstoregiftOptionRegXML(itemid, getfarmGoodno, oService, oOperation)
		If strParam <> "X" Then
			Call fnNvstoregiftOptionReg(itemid, strParam, iErrStr, oService, oOperation)
		Else
			iErrStr = "OK||"&itemid&"||단품으로 API연동 불필요..by김진영"
		End If
	Else
		iErrStr = "ERR||"&itemid&"||상품이 등록되지 않았습니다."
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "Image" Then									'이미지 등록
	SET oNvstoregift = new CNvstoregift
		oNvstoregift.FRectItemID	= itemid
		oNvstoregift.FRectGubun		= "IMG"
		oNvstoregift.getNvstoregiftNotRegOneItem
		If (oNvstoregift.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_nvstoregift_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstoregift_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, nvstoregiftstatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oNvstoregift.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oNvstoregift.FOneitem.checkTenItemOptionValid Then
				oService		= "ImageService"
				oOperation		= "UploadImage"

				strParam = ""
				strParam = oNvstoregift.FOneitem.getNvstoregiftImageRegXML(oService, oOperation)
				chgImageNm = oNvstoregift.FOneItem.getBasicImage
				Call fnNvstoregiftImageReg(itemid, strParam, iErrStr, chgImageNm, oService, oOperation)
			Else
				iErrStr = "ERR||"&itemid&"||옵션검사 실패"
			End If
		End If
	SET oNvstoregift = nothing
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKOPT" Then									'옵션 조회
	oService		= "ProductService"
	oOperation		= "GetOption"

	strParam = ""
	strParam = getNvstoregiftOptionSearchParameter(getNvstoregiftGoodNo(itemid), oService, oOperation)
	Call fnNvstoregiftOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKSTAT" Then									'상품 조회
	oService		= "ProductService"
	oOperation		= "GetProduct"

	strParam = ""
	strParam = getNvstoregiftItemSearchParameter(getNvstoregiftGoodNo(itemid), oService, oOperation)
	Call fnNvstoregiftItemSearch(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDIT" Then										'상품조회 -> 옵션조회 -> 상품수정 -> 옵션수정 순
	SET oNvstoregift = new CNvstoregift
		oNvstoregift.FRectItemID	= itemid
		oNvstoregift.getNvstoregiftEditOneItem

		If (oNvstoregift.FResultCount < 1) OR (oNvstoregift.FOneItem.FNvstoregiftGoodNo = "") Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
			failCnt = failCnt + 1
		Else
			If oNvstoregift.FOneItem.FOptioncnt > 0 Then
				mayOptSoldOut = oNvstoregift.FOneItem.IsMayLimitSoldout
			End If

			If (oNvstoregift.FOneItem.FMaySoldOut = "Y") OR (oNvstoregift.FOneItem.IsSoldOutLimit5Sell) OR (mayOptSoldOut = "Y") OR (oNvstoregift.FOneItem.FLimitYn = "Y" AND oNvstoregift.FOneItem.getiszeroWonSoldOut(itemid) = "Y") Then
				oService		= "ProductService"
				oOperation		= "ChangeProductSaleStatus"

				strParam = ""
				strParam = getNvstoregiftSellynParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, "N", oService, oOperation)
				Call fnNvstoregiftSellyn(itemid, "N", strParam, iErrStr, oService, oOperation)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (oNvstoregift.FOneItem.FNvstoregiftSellYn = "N" AND oNvstoregift.FOneItem.IsSoldOutLimit5Sell = False) Then
					oService		= "ProductService"
					oOperation		= "ChangeProductSaleStatus"

					strParam = ""
					strParam = getNvstoregiftSellynParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, "Y", oService, oOperation)
					Call fnNvstoregiftSellyn(itemid, "Y", strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
	'################################################ 0.상품 수정부터(ReturnCostReason 신규 필드 때문..) ####################
				If oNvstoregift.FOneItem.isImageChanged = True Then
					chgImageNm = oNvstoregift.FOneItem.getBasicImage
				Else
					chgImageNm = "N"
				End If

				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oNvstoregift.FOneitem.getNvstoregiftItemRegXML(oService, oOperation, "Y")
				getMustprice = ""
				getMustprice = oNvstoregift.FOneItem.MustPrice()
				Call fnNvstoregiftItemEDIT(itemid, strParam, iErrStr, getMustprice, oNvstoregift.FOneItem.getNvstoregiftSellYn, oNvstoregift.FOneItem.FLimityn, oNvstoregift.FOneItem.FLimitNo, oNvstoregift.FOneItem.FLimitSold, oNvstoregift.FOneItem.FItemName, chgImageNm, oService, oOperation)
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
				strParam = getNvstoregiftOptionSearchParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, oService, oOperation)
				Call fnNvstoregiftOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
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
					strParam = oNvstoregift.FOneitem.getNvstoregiftImageRegXML(oService, oOperation)
					chgImageNm = oNvstoregift.FOneItem.getBasicImage
					Call fnNvstoregiftImageReg(itemid, strParam, iErrStr, chgImageNm, oService, oOperation)
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
					strParam = oNvstoregift.FOneitem.getNvstoregiftItemRegXML(oService, oOperation, "Y")
					getMustprice = ""
					getMustprice = oNvstoregift.FOneItem.MustPrice()
					Call fnNvstoregiftItemEDIT(itemid, strParam, iErrStr, getMustprice, oNvstoregift.FOneItem.getNvstoregiftSellYn, oNvstoregift.FOneItem.FLimityn, oNvstoregift.FOneItem.FLimitNo, oNvstoregift.FOneItem.FLimitSold, (oNvstoregift.FOneItem.FItemName), chgImageNm, oService, oOperation)
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
					strParam = getNvstoregiftOptionRegXML(itemid, oNvstoregift.FOneItem.FNvstoregiftGoodno, oService, oOperation)
					If strParam <> "X" Then
						Call fnNvstoregiftOptionReg(itemid, strParam, iErrStr, oService, oOperation)
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
					strParam = getNvstoregiftOptionSearchParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, oService, oOperation)
					Call fnNvstoregiftOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
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
					strParam = getNvstoregiftDeleteParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, oService, oOperation)
					Call fnNvstoregiftDelete(itemid, strParam, iErrStr, oService, oOperation)
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
		strSql = strSql & " UPDATE [db_etcmall].[dbo].tbl_nvstoregift_regItem SET " & VBCRLF
		strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
		strSql = strSql & " ,nvstoregiftlastupdate = getdate()  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("nvstoregift", itemid, SumErrStr)
			Call SugiQueLogInsert("nvstoregift", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_nvstoregift_regItem SET " & VBCRLF
			strSql = strSql & " accFailcnt = 0  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("nvstoregift", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oNvstoregift = nothing
ElseIf action = "EditSellYn" Then								'상태 변경
	oService		= "ProductService"
	oOperation		= "ChangeProductSaleStatus"

	strParam = ""
	strParam = getNvstoregiftSellynParameter(getNvstoregiftGoodNo(itemid), chgSellYn, oService, oOperation)
	Call fnNvstoregiftSellyn(itemid, chgSellYn, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "DEL" Then										'상품 삭제
	oService		= "ProductService"
	oOperation		= "DeleteProduct"

	strParam = ""
	strParam = getNvstoregiftDeleteParameter(getNvstoregiftGoodNo(itemid), oService, oOperation)
	Call fnNvstoregiftDelete(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "nvstorefarmCommonCode" Then					'공통코드 검색
	If ccd = "GetAddressBookList" Then
		strParam = ""
		strParam = getAddressBookList(ccd)
	End If
'	Call fnAuctionCommonCode(ccd, strParam)
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
