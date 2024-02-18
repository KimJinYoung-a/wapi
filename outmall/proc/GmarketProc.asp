<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/gmarket/gmarketItemcls.asp"-->
<!-- #include virtual="/outmall/gmarket/incGmarketFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, failCnt, arrRows, skipItem, oGmarket, getMustprice
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname
Dim tGmarketGoodno, tOptionCnt, tLimityn, isAllRegYn, displayDate, isFiftyUpDown, isiframe
Dim isChild, isLife, isElec
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
'######################################################## Gmarket API ########################################################
If mallid = "gmarket1010" Then
	If action = "REG" Then					'상품등록
	'##################################### 기본 정보 등록 시작 #####################################
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketNotRegOneItem
		    If (oGmarket.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
			ElseIf (oGmarket.FOneItem.FDepthCode = "0") Then
				iErrStr = "ERR||"&itemid&"||카테고리 매칭이 필요합니다."
			' ElseIf (oGmarket.FOneItem.FBrandCode = "0") Then
			' 	iErrStr = "ERR||"&itemid&"||브랜드 매칭이 필요합니다."
			Else
				strSql = ""
				strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_gmarket_regitem where itemid="&itemid&")"
				strSql = strSql & " BEGIN"& VbCRLF
				strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_gmarket_regitem "
		        strSql = strSql & " 	(itemid, regdate, reguserid, gmarketstatCD, regitemname, gmarketSellYn)"
		        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oGmarket.FOneItem.FItemName)&"', 'N')"
				strSql = strSql & " END "
				dbget.Execute strSql

				'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
				If oGmarket.FOneItem.checkTenItemOptionValid Then
					strParam = ""
					strParam = oGmarket.FOneItem.getGmarketItemRegParameter(FALSE)
					Call fnGmarketItemReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FbasicimageNm)
				Else
					iErrStr = "ERR||"&itemid&"||[AddItem] 옵션검사 실패"
				End If
			End If
		SET oGmarket = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
		'##################################### 기본 정보 등록 끝 #####################################

		'#################################### 고시 정보 등록 시작 ####################################
		If failCnt = 0 Then
			tGmarketGoodno = getGmarketGoodno(itemid)
			If tGmarketGoodno = "" Then
				iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
			Else
				strParam = ""
				strParam = getGmarketInfoCdParameter(itemid, tGmarketGoodno)
				Call fnGmarketItemInfoCd(itemid, strParam, iErrStr)
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
		'#################################### 고시 정보 등록 끝 ####################################

		'#################################### 어린이 인증 등록 시작 ####################################
		If failCnt = 0 Then
			Call getGmarketChildrenCate(itemid, isChild, isLife, isElec)
			If isChild = "Y" OR isLife = "Y" OR isElec = "Y" Then
				strParam = ""
				strParam = getGmarketChildrenParameter(itemid, tGmarketGoodno, isChild, isLife, isElec)
				Call fnGmarketItemChildren(itemid, strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		End If
		'#################################### 어린이 인증 등록 끝 ####################################

		'#################################### 반품비 등록 시작 ####################################
		If failCnt = 0 Then
			strParam = ""
			strParam = getGmarketReturnFeeParameter(itemid, tGmarketGoodno, CRETURNFEE)
			Call fnGmarketReturnFee(itemid, strParam, CRETURNFEE, iErrStr)

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
		'#################################### 반품비 등록 끝 ####################################

		'#################################### 옵션 정보 등록 시작 ####################################
		If failCnt = 0 Then
			SET oGmarket = new CGmarket
				oGmarket.FRectItemID	= itemid
				oGmarket.getGmarketNotOptOneItem
			    If (oGmarket.FResultCount < 1) Then
					iErrStr = "ERR||"&itemid&"||옵션 등록 가능한 상품이 아닙니다."
				ElseIf (oGmarket.FOneItem.FGmarketGoodNo = "") Then
					iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
				ElseIf (oGmarket.FOneItem.FAPIadditem = "N") Then
					iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
				ElseIf (oGmarket.FOneItem.getFiftyUpDown = "Y") Then
					iErrStr = "ERR||"&itemid&"||옵션가격이 50%를 초과합니다."
				Else
					'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
					If oGmarket.FOneItem.checkTenItemOptionValid Then
						strParam = ""
						strParam = oGmarket.FOneItem.getGmarketItemOptRegParameter()
						Call fnGmarketOPTReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FLimityn, oGmarket.FOneItem.FLimitno, oGmarket.FOneItem.FLimitsold)
					Else
						iErrStr = "ERR||"&itemid&"||[AddOPT] 옵션검사 실패"
					End If
				End If
			SET oGmarket = nothing
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
		'#################################### 옵션 정보 등록 끝 ####################################
		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/gmarketProc.asp?itemid=699617&mallid=gmarket1010&action=REG
	ElseIf action = "REGOnSale" Then						'신규등록 상품 판매중으로 변경
		isAllRegYn = getAllRegChk(itemid, "X")
		If isAllRegYn <> "Y" Then
			iErrStr = "ERR||"&itemid&"||기본정보, 옵션정보, 상품고시 입력을 확인하세요"
		Else
			tGmarketGoodno = getGmarketGoodno(itemid)
			strParam = ""
			strParam = getGmarketAddPriceParameter(itemid, tGmarketGoodno, "Y", mustPrice, displayDate)
			Call fnGmarketItemAddPrice(itemid, strParam, mustPrice, displayDate, "Y", iErrStr)
		End If
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
	ElseIf action = "SOLDOUT" Then			'상태변경
		isAllRegYn = getAllRegChk2(itemid, tGmarketGoodno, tOptionCnt, tLimityn, "Y")
		If tGmarketGoodno = "" Then
			iErrStr = "ERR||"&itemid&"||기본정보, 옵션정보, 상품고시 입력을 확인하세요"
		Else
			strParam = ""
			strParam = getGmarketAddPriceParameter(itemid, tGmarketGoodno, "N", mustPrice, displayDate)
			Call fnGmarketItemAddPrice(itemid, strParam, mustPrice, displayDate, "N", iErrStr)
		End If
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=SOLDOUT
	ElseIf action = "EDITIMG" Then		'이미지 수정
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketEditImageOneItem
		    If (oGmarket.FOneItem.FGmarketGoodNo = "") Then
				iErrStr = "ERR||"&itemid&"||이미지 수정 가능한 상품이 아닙니다."
			Else
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketItemEditImgParameter()
				Call fnGmarketEditImg(itemid, strParam, iErrStr, oGmarket.FOneItem.FbasicimageNm)
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
			End If
		SET oGmarket = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=EDITIMG
	ElseIf action = "EDITINFO" Then		'기본정보 수정
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketEditOneItem

			If (oGmarket.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
			ElseIf oGmarket.FOneItem.checkItemContent = "Y" Then
				iErrStr = "ERR||"&itemid&"||iframe이 속한 상품은 수정 할 수 없습니다."
			Else
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketItemRegParameter(TRUE)
				Call fnGmarketIteminfoEdit(itemid, oGmarket.FOneItem.FGmarketGoodNo, oGmarket.FOneItem.FItemName, iErrStr, strParam)
			End If

			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
			End If
		SET oGmarket = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=EDITINFO
	ElseIf action = "EDITPOLICY" Then	'기본정보 수정 + 반품비 수정
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketEditOneItem

			If (oGmarket.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
			ElseIf oGmarket.FOneItem.checkItemContent = "Y" Then
				iErrStr = "ERR||"&itemid&"||iframe이 속한 상품은 수정 할 수 없습니다."
			Else
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketItemRegParameter(TRUE)
				Call fnGmarketIteminfoEdit(itemid, oGmarket.FOneItem.FGmarketGoodNo, oGmarket.FOneItem.FItemName, iErrStr, strParam)
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			If failCnt = 0 Then
				strParam = ""
				strParam = getGmarketReturnFeeParameter(itemid, oGmarket.FOneItem.FGmarketGoodNo, CRETURNFEE)
				Call fnGmarketReturnFee(itemid, strParam, CRETURNFEE, iErrStr)

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
				CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oGmarket = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=EDITPOLICY
	ElseIf action = "KEEPSELL" Then		'상품 판매 유지
		isAllRegYn = getAllRegChk2(itemid, tGmarketGoodno, tOptionCnt, tLimityn, "Y")
		If tGmarketGoodno = "" Then
			iErrStr = "ERR||"&itemid&"||기본정보, 옵션정보, 상품고시 입력을 확인하세요"
		Else
			strParam = ""
			strParam = getGmarketAddPriceParameter(itemid, tGmarketGoodno, "Y", mustPrice, displayDate)
			Call fnGmarketItemAddPrice(itemid, strParam, mustPrice, displayDate, "Y", iErrStr)
		End If
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=KEEPSELL
	ElseIf action = "PRICE" Then		'가격수정
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketEditPriceOptOneItem
			If oGmarket.FResultCount > 0 Then
				'옵션추가금액이 상품금액의 50%초과 검사
				isFiftyUpDown = oGmarket.FOneItem.getFiftyUpDown

				getMustprice = ""
				getMustprice = oGmarket.FOneItem.MustPrice()
				'만약 품절에 해당하거나 50%초과하거나 0원옵션이 모두 품절일 때..(한정상품경우 재고 5개이하도 포함함)
				If (oGmarket.FOneItem.FmaySoldOut = "Y") OR (isFiftyUpDown = "Y") OR (oGmarket.FOneItem.FLimityn = "Y" AND (oGmarket.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) OR (oGmarket.FOneItem.IsMayLimitSoldout = "Y") Then
					strParam = ""
					strParam = oGmarket.FOneItem.getGmarketAddPriceParameter("N", getMustprice, displayDate)
					Call fnGmarketItemAddPrice(itemid, strParam, getMustprice, displayDate, "N", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					SET oGmarket = nothing
				Else
				'위 조건에 해당하지 않으면 무조건 판매처리
					iErrStr = ""
					strParam = ""
					strParam = oGmarket.FOneItem.getGmarketAddPriceParameter("Y", getMustprice, displayDate)
					Call fnGmarketItemAddPrice(itemid, strParam, getMustprice, displayDate, "Y", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					SET oGmarket = nothing

					SET oGmarket = new CGmarket
						oGmarket.FRectItemID	= itemid
						oGmarket.getGmarketNotOptOneItem
					    If (oGmarket.FResultCount < 1) Then
							iErrStr = "ERR||"&itemid&"||옵션 등록 가능한 상품이 아닙니다."
						ElseIf (oGmarket.FOneItem.FGmarketGoodNo = "") Then
							iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
						ElseIf (oGmarket.FOneItem.FAPIadditem = "N") Then
							iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
						ElseIf (oGmarket.FOneItem.getFiftyUpDown = "Y") Then
							iErrStr = "ERR||"&itemid&"||옵션가격이 50%를 초과합니다."
						Else
							'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
							If oGmarket.FOneItem.checkTenItemOptionValid Then
								strParam = ""
								strParam = oGmarket.FOneItem.getGmarketItemOptRegParameter()
								Call fnGmarketOPTReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FLimityn, oGmarket.FOneItem.FLimitno, oGmarket.FOneItem.FLimitsold)
							Else
								iErrStr = "ERR||"&itemid&"||[AddOPT] 옵션검사 실패"
							End If
						End If
					SET oGmarket = nothing
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			Else
				iErrStr = "ERR||"&itemid&"||수정할 데이터 없음"
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
				CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_gmarket_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=PRICE
	ElseIf action = "EDIT" Then			'재고조회 + 상품정보 + 가격 + 필요에 따라 상품판매상태수정
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketEditOneItem
			'#################################### 기본 정보 수정 시작 ####################################
		    If (oGmarket.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
			ElseIf oGmarket.FOneItem.checkItemContent = "Y" Then
				iErrStr = "ERR||"&itemid&"||iframe이 속한 상품은 수정 할 수 없습니다."
				isiframe = "Y"
			Else
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketItemRegParameter(TRUE)
				Call fnGmarketIteminfoEdit(itemid, oGmarket.FOneItem.FGmarketGoodNo, oGmarket.FOneItem.FItemName, iErrStr, strParam)
			End If
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			'#################################### 반품비 정보 수정 시작 ####################################
			If (oGmarket.FResultCount > 0) AND (oGmarket.FOneItem.FReturnShippingFee < 100) Then
				strParam = ""
				strParam = getGmarketReturnFeeParameter(itemid, oGmarket.FOneItem.FGmarketGoodNo, CRETURNFEE)
				Call fnGmarketReturnFee(itemid, strParam, CRETURNFEE, iErrStr)

				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If

			'#################################### 이미지 수정 시작 ####################################
			If (oGmarket.FResultCount > 0) AND (oGmarket.FOneItem.isImageChanged) Then
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketItemEditImgParameter()
				Call fnGmarketEditImg(itemid, strParam, iErrStr, oGmarket.FOneItem.FbasicimageNm)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		SET oGmarket = nothing

		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketEditPriceOptOneItem
			'#################################### 상품 가격 수정 시작 ####################################
			If oGmarket.FResultCount > 0 Then
				'옵션추가금액이 상품금액의 50%초과 검사
				isFiftyUpDown = oGmarket.FOneItem.getFiftyUpDown

				getMustprice = ""
				getMustprice = oGmarket.FOneItem.MustPrice()
				'만약 품절에 해당하거나 아이프레임이 있거나 50%초과하거나 0원옵션이 모두 품절일 때..(한정상품경우 재고 5개이하도 포함함)
				If (oGmarket.FOneItem.FmaySoldOut = "Y") OR (isFiftyUpDown = "Y") OR (isiframe = "Y") OR (oGmarket.FOneItem.FLimityn = "Y" AND (oGmarket.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) OR (oGmarket.FOneItem.IsMayLimitSoldout = "Y") Then
					strParam = ""
					strParam = oGmarket.FOneItem.getGmarketAddPriceParameter("N", getMustprice, displayDate)
					Call fnGmarketItemAddPrice(itemid, strParam, getMustprice, displayDate, "N", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						failCnt = 0
						SumOKStr = SumOKStr & iErrStr
					End If
					SET oGmarket = nothing
				Else
				'위 조건에 해당하지 않으면 무조건 판매처리
					iErrStr = ""
					strParam = ""
					strParam = oGmarket.FOneItem.getGmarketAddPriceParameter("Y", getMustprice, displayDate)
					Call fnGmarketItemAddPrice(itemid, strParam, getMustprice, displayDate, "Y", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					SET oGmarket = nothing
			'#################################### 상품 옵션 수정 시작 ####################################
					SET oGmarket = new CGmarket
						oGmarket.FRectItemID	= itemid
						oGmarket.getGmarketNotOptOneItem
					    If (oGmarket.FResultCount < 1) Then
							iErrStr = "ERR||"&itemid&"||옵션 등록 가능한 상품이 아닙니다."
						ElseIf (oGmarket.FOneItem.FGmarketGoodNo = "") Then
							iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
						ElseIf (oGmarket.FOneItem.FAPIadditem = "N") Then
							iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
						ElseIf (oGmarket.FOneItem.getFiftyUpDown = "Y") Then
							iErrStr = "ERR||"&itemid&"||옵션가격이 50%를 초과합니다."
						Else
							'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
							If oGmarket.FOneItem.checkTenItemOptionValid Then
								strParam = ""
								strParam = oGmarket.FOneItem.getGmarketItemOptRegParameter()
								Call fnGmarketOPTReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FLimityn, oGmarket.FOneItem.FLimitno, oGmarket.FOneItem.FLimitsold)
							Else
								iErrStr = "ERR||"&itemid&"||[AddOPT] 옵션검사 실패"
							End If
						End If
					SET oGmarket = nothing
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			Else
				iErrStr = "ERR||"&itemid&"||수정할 데이터 없음"
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If

			'OK던 ERR이던 editQuecnt에 + 1을 시킴..
			'스케줄링에서 editQuecnt ASC, i.lastupdate DESC로 중복을 막자
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_gmarket_regItem SET " & VBCRLF
			strSql = strSql & " editQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
			strSql = strSql & " ,gmarketlastupdate = getdate()  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_gmarket_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=EDIT
	End If
End If
'###################################################### Gmarket API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
