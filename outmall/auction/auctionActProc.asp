<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/auction/auctionItemcls.asp"-->
<!-- #include virtual="/outmall/auction/incAuctionFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, oAuction, oAuctionOpt, failCnt, chgSellYn, arrRows, skipItem, tAuctionGoodno, isAllRegYn, getMustprice
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, ccd, isItemIdChk
Dim isoptionyn, isText, i, isiframe
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
ccd				= request("ccd")
failCnt			= 0

Select Case action
	Case "auctionCommonCode"	isItemIdChk = "N"
	Case Else					isItemIdChk = "Y"
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
'######################################################## Auction API ########################################################
If action = "REGAddItem" Then							'상품 기본 정보 등록
	SET oAuction = new CAuction
		oAuction.FRectItemID	= itemid
		oAuction.getAuctionNotRegOneItem

	    If (oAuction.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf (oAuction.FOneItem.FNotinCate = "Y") Then
			iErrStr = "ERR||"&itemid&"||상품 등록 제외 카테고리입니다."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_auction_regitem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_auction_regitem "
	        strSql = strSql & " 	(itemid, regdate, reguserid, auctionstatCD, regitemname, auctionSellYn)"
	        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oAuction.FOneItem.FItemName)&"', 'N')"
			strSql = strSql & " END "
			dbget.Execute strSql

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oAuction.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oAuction.FOneItem.getAuctionItemRegParameter()
'response.write strParam
'response.end
				getMustprice = ""
				getMustprice = oAuction.FOneItem.MustPrice()
				Call fnAuctionItemReg(itemid, strParam, iErrStr, getMustprice, oAuction.FOneItem.getAuctionSellYn, oAuction.FOneItem.FLimityn, oAuction.FOneItem.FLimitNo, oAuction.FOneItem.FLimitSold, html2db(oAuction.FOneItem.FItemName), oAuction.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[AddItem] 옵션검사 실패"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("auction1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oAuction = nothing
ElseIf action = "REGAddOPT" Then						'상품 옵션 정보 등록
	SET oAuction = new CAuction
		oAuction.FRectItemID	= itemid
		oAuction.getAuctionNotOptOneItem

	    If (oAuction.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||옵션 등록 가능한 상품이 아닙니다."
		ElseIf (oAuction.FOneItem.FAuctionGoodNo = "") Then
			iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
		ElseIf (oAuction.FOneItem.FAPIadditem = "N") Then
			iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
		ElseIf (oAuction.FOneItem.FAPIaddopt = "Y") Then
			iErrStr = "ERR||"&itemid&"||이미 옵션정보를 등록하셨습니다."
		Else
			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oAuction.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
				Call fnAuctionOPTReg(itemid, strParam, iErrStr)
			Else
				iErrStr = "ERR||"&itemid&"||[AddOPT] 옵션검사 실패"
			End If

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("auction1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oAuction = nothing
ElseIf action = "REGInfoCd" Then						'상품고시 등록
	tAuctionGoodno = getAuctionGoodno(itemid)
	If tAuctionGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
	Else
		strParam = ""
		strParam = getAuctionInfoCdParameter(itemid, tAuctionGoodno)
		Call fnAuctionItemInfoCd(itemid, strParam, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("auction1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "REG" Then								'기본정보 + 옵션정보 + 고시정보 등록
	'##################################### 기본 정보 등록 시작 #####################################
	SET oAuction = new CAuction
		oAuction.FRectItemID	= itemid
		oAuction.getAuctionNotRegOneItem
	    If (oAuction.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf (oAuction.FOneItem.FNotinCate = "Y") Then
			iErrStr = "ERR||"&itemid&"||상품 등록 제외 카테고리입니다."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_auction_regitem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_auction_regitem "
	        strSql = strSql & " 	(itemid, regdate, reguserid, auctionstatCD, regitemname, auctionSellYn)"
	        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oAuction.FOneItem.FItemName)&"', 'N')"
			strSql = strSql & " END "
			dbget.Execute strSql

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oAuction.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oAuction.FOneItem.getAuctionItemRegParameter()
				getMustprice = ""
				getMustprice = oAuction.FOneItem.MustPrice()
				Call fnAuctionItemReg(itemid, strParam, iErrStr, getMustprice, oAuction.FOneItem.getAuctionSellYn, oAuction.FOneItem.FLimityn, oAuction.FOneItem.FLimitNo, oAuction.FOneItem.FLimitSold, html2db(oAuction.FOneItem.FItemName), oAuction.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[AddItem] 옵션검사 실패"
			End If
		End If
	SET oAuction = nothing
	If Left(iErrStr, 2) <> "OK" Then
		failCnt = failCnt + 1
		SumErrStr = SumErrStr & iErrStr
	Else
		SumOKStr = SumOKStr & iErrStr
	End If
	'##################################### 기본 정보 등록 끝 #####################################

	'#################################### 옵션 정보 등록 시작 ####################################
	If failCnt = 0 Then
		SET oAuctionOpt = new CAuction
			oAuctionOpt.FRectItemID	= itemid
			oAuctionOpt.getAuctionNotOptOneItem
		    If (oAuctionOpt.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||옵션 등록 가능한 상품이 아닙니다."
			ElseIf (oAuctionOpt.FOneItem.FAuctionGoodNo = "") Then
				iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
			ElseIf (oAuctionOpt.FOneItem.FAPIadditem = "N") Then
				iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
			ElseIf (oAuctionOpt.FOneItem.FAPIaddopt = "Y") Then
				iErrStr = "ERR||"&itemid&"||이미 옵션정보를 등록하셨습니다."
			Else
				strParam = ""
				strParam = oAuctionOpt.FOneItem.getAuctionOPTRegParameter()
				Call fnAuctionOPTReg(itemid, strParam, iErrStr)
			End If
			tAuctionGoodno = oAuctionOpt.FOneItem.FAuctionGoodNo
		SET oAuctionOpt = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'#################################### 옵션 정보 등록 끝 ####################################

	'################################# 상품고시 정보 등록 시작 #################################
	If failCnt = 0 Then
		If tAuctionGoodno = "" Then
			iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
		Else
			strParam = ""
			strParam = getAuctionInfoCdParameter(itemid, tAuctionGoodno)
			Call fnAuctionItemInfoCd(itemid, strParam, iErrStr)
		End If
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'################################## 상품고시 정보 등록 끝 ##################################
	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("auction1010", itemid, SumErrStr)
		Call SugiQueLogInsert("auction1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("auction1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "REGOnSale" Then						'옵션 조회 후  신규등록 상품 판매중으로 변경
	isAllRegYn = getAllRegChk(itemid)
	If isAllRegYn <> "Y" Then
		iErrStr = "ERR||"&itemid&"||기본정보, 옵션정보, 상품고시 입력을 확인하세요"
	Else
		tAuctionGoodno = getAuctionGoodno(itemid)
		strParam = ""
		strParam = getAuctionOptSellModParameter(tAuctionGoodno)
		Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If

		If failCnt = 0 Then
			strParam = ""
			strParam = getAuctionSellYnParameter("Y", itemid, tAuctionGoodno)
			Call fnAuctionSellyn(itemid, "Y", strParam, iErrStr)
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
			CALL Fn_AcctFailTouch("auction1010", itemid, SumErrStr)
			Call SugiQueLogInsert("auction1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("auction1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	End If
ElseIf action = "EditSellYn" Then						'상품 상태 변경
	isAllRegYn = getAllRegChk2(itemid)
	If isAllRegYn <> "Y" Then
		iErrStr = "ERR||"&itemid&"||기본정보, 옵션정보, 상품고시 입력을 확인하세요"
	Else
		tAuctionGoodno = getAuctionGoodno(itemid)
		strParam = ""
		strParam = getAuctionSellYnParameter(chgSellYn, itemid, tAuctionGoodno)
		Call fnAuctionSellyn(itemid, chgSellYn, strParam, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("auction1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "EditInfo" Then							'기본정보수정(상품명, 가격, 이미지, 상품상세등등)
	SET oAuction = new CAuction
		oAuction.FRectItemID	= itemid
		oAuction.getAuctionEditOneItem
	    If (oAuction.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		ElseIf getAllRegChk2(itemid) <> "Y" Then
			iErrStr = "ERR||"&itemid&"||OnSale변경 확인하세요"
		Else
			strParam = ""
			strParam = oAuction.FOneItem.getAuctionItemInfoEditParameter()

			getMustprice = ""
			getMustprice = oAuction.FOneItem.MustPrice()
			Call fnAuctionIteminfoEdit(itemid, oAuction.FOneItem.FAuctionGoodNo, iErrStr, strParam, getMustprice)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
			End If

			If (Left(iErrStr,2)) <> "OK" and (Left(iErrStr,2)) <> "ER" Then
				iErrStr = "ERR||"&itemid&"||잘못된 호출"
			End If

			Call SugiQueLogInsert("auction1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oAuction = nothing
ElseIf action = "OPTSTAT" Then							'상품 조회(옵션 정보 가져오기)
	tAuctionGoodno = getAuctionGoodno(itemid)
	If tAuctionGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||등록된 상품이 아닙니다."
	Else
		strParam = ""
		strParam = getAuctionOptSellModParameter(tAuctionGoodno)
		Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("auction1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "OPTEDIT" Then							'옵션 정보 수정
	SET oAuction = new CAuction
		oAuction.FRectItemID	= itemid
		oAuction.getAuctionEditOneItem
	    If (oAuction.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		ElseIf getAllRegChk2(itemid) <> "Y" Then
			iErrStr = "ERR||"&itemid&"||OnSale변경 확인하세요"
		Else
			If (oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt > 0) OR (oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt = 0) Then			'텐바이텐 옵션있고, 옥션의 옵션도 등록되어있다면..즉 둘다 옵션상태

				'## 총 3번의 API를 돌려야 될 것
				'1.옵션을 완전 초기화
				strParam = ""
				strParam = oAuction.FOneItem.getAuctionOPTDeleteParameter()
				Call fnAuctionOPTDel(itemid, strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'2.초기화 후 재 세팅
				strParam = ""
				strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
				Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'3.옵션 조회로 가져오기
				strParam = ""
				strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
				Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If oAuction.FOneItem.FOptioncnt = 0 AND oAuction.FOneItem.FRegedoptcnt = 0 Then		'둘 다 단품 상태
					If (oAuction.FOneItem.FaccFailCNT > 0 AND InStr(oAuction.FOneItem.FlastErrStr, "텍스트형은 최소 1건 이상 노출되어야 합니다") > 0) Then
						strParam = ""
						strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
						Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

					strParam = ""
					strParam = oAuction.FOneItem.getAuctionDanPoomModParameter()
					Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					If failCnt = 0 Then
						strSql = ""
						strSql = " DELETE FROM db_item.dbo.tbl_outmall_regedoption WHERE mallid = '"&CMALLNAME&"' and itemid = '"&itemid&"' "
						dbget.Execute(strSql)

						strSql = ""
						strSql = "UPDATE db_etcmall.dbo.tbl_auction_regitem SET regedoptcnt = null WHERE itemid = '"&itemid&"'"
						dbget.Execute(strSql)
					End If

					'2. 옵션 조회로 가져오기
					strParam = ""
					strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
					Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				ElseIf oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt = 0 Then	'텐바이텐상품이 옵션있음 으로 변경되고, 등록된 옵션은 없는 상태
					'1. 재 세팅
					strParam = ""
					strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
					Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					If failCnt = 0 Then
						strSql = ""
						strSql = " DELETE FROM db_item.dbo.tbl_outmall_regedoption WHERE mallid = '"&CMALLNAME&"' and itemid = '"&itemid&"' "
						dbget.Execute(strSql)

						strSql = ""
						strSql = "UPDATE db_etcmall.dbo.tbl_auction_regitem SET regedoptcnt = null WHERE itemid = '"&itemid&"'"
						dbget.Execute(strSql)
					End If

					'2. 옵션 조회로 가져오기
					strParam = ""
					strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
					Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

				ElseIf oAuction.FOneItem.FOptioncnt = 0 AND oAuction.FOneItem.FRegedoptcnt > 0 Then	'텐바이텐상품이 옵션있음에서 단품으로 변경되고, 등록된 옵션이 있는 상태
					'1.옵션을 완전 초기화
					strParam = ""
					strParam = oAuction.FOneItem.getAuctionOPTDeleteParameter()
					Call fnAuctionOPTDel(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'2.옵션 조회로 가져오기
					strParam = ""
					strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
					Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("auction1010", itemid, SumErrStr)
				Call SugiQueLogInsert("auction1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				Call SugiQueLogInsert("auction1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
				iErrStr = "OK||"&itemid&"||"&SumOKStr
			End If
		End If
	SET oAuction = nothing
ElseIf action = "EDIT" Then
	If getAllRegChk2(itemid) <> "Y" Then
		iErrStr = "ERR||"&itemid&"||OnSale변경 확인하세요"
	Else
		SET oAuction = new CAuction
			oAuction.FRectItemID	= itemid
			oAuction.getAuctionEditOneItem
			If oAuction.FResultCount > 0 Then

				If oAuction.FOneItem.checkItemContent = "Y" Then
					isiframe = "Y"
				End If
'rw Instr(oAuction.FOneItem.FItemcontent, "<IFRAME")
'response.end
'response.write isiframe & "<br />"
				If (oAuction.FOneItem.FmaySoldOut = "Y") OR (isiframe = "Y") OR (oAuction.FOneItem.IsMayLimitSoldout = "Y") Then
					strParam = ""
					strParam = getAuctionSellYnParameter("N", itemid, oAuction.FOneItem.FAuctionGoodNo)
					Call fnAuctionSellyn(itemid, "N", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					If (oAuction.FOneItem.FAuctionSellYn = "N" AND oAuction.FOneItem.IsSoldOut = False) Then
						iErrStr = ""
						strParam = ""
						strParam = getAuctionSellYnParameter("Y", itemid, oAuction.FOneItem.FAuctionGoodNo)
						Call fnAuctionSellyn(itemid, "Y", strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

					strParam = ""
					strParam = oAuction.FOneItem.getAuctionItemInfoEditParameter()
					getMustprice = ""
					getMustprice = oAuction.FOneItem.MustPrice()
					Call fnAuctionIteminfoEdit(itemid, oAuction.FOneItem.FAuctionGoodNo, iErrStr, strParam, getMustprice)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					If (oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt > 0) OR (oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt = 0) Then			'텐바이텐 옵션있고, 옥션의 옵션도 등록되어있다면..즉 둘다 옵션상태
						'## 총 3번의 API를 돌려야 될 것
						'1.옵션을 완전 초기화
						strParam = ""
						strParam = oAuction.FOneItem.getAuctionOPTDeleteParameter()
						Call fnAuctionOPTDel(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If

						'2.초기화 후 재 세팅
						strParam = ""
						strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
						Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If

						'3.옵션 조회로 가져오기
						strParam = ""
						strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
						Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					Else
						If oAuction.FOneItem.FOptioncnt = 0 AND oAuction.FOneItem.FRegedoptcnt = 0 Then		'둘 다 단품 상태
							If (oAuction.FOneItem.FaccFailCNT > 0 AND InStr(oAuction.FOneItem.FlastErrStr, "텍스트형은 최소 1건 이상 노출되어야 합니다") > 0) Then
								strParam = ""
								strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
								Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
								If Left(iErrStr, 2) <> "OK" Then
									failCnt = failCnt + 1
									SumErrStr = SumErrStr & iErrStr
								Else
									SumOKStr = SumOKStr & iErrStr
								End If
							End If

							strParam = ""
							strParam = oAuction.FOneItem.getAuctionDanPoomModParameter()
							Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If

							If failCnt = 0 Then
								strSql = ""
								strSql = " DELETE FROM db_item.dbo.tbl_outmall_regedoption WHERE mallid = '"&CMALLNAME&"' and itemid = '"&itemid&"' "
								dbget.Execute(strSql)

								strSql = ""
								strSql = "UPDATE db_etcmall.dbo.tbl_auction_regitem SET regedoptcnt = null WHERE itemid = '"&itemid&"'"
								dbget.Execute(strSql)
							End If

							'2. 옵션 조회로 가져오기
							strParam = ""
							strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
							Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						ElseIf oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt = 0 Then	'텐바이텐상품이 옵션있음 으로 변경되고, 등록된 옵션은 없는 상태
							'1. 재 세팅
							strParam = ""
							strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
							Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If

							If failCnt = 0 Then
								strSql = ""
								strSql = " DELETE FROM db_item.dbo.tbl_outmall_regedoption WHERE mallid = '"&CMALLNAME&"' and itemid = '"&itemid&"' "
								dbget.Execute(strSql)

								strSql = ""
								strSql = "UPDATE db_etcmall.dbo.tbl_auction_regitem SET regedoptcnt = null WHERE itemid = '"&itemid&"'"
								dbget.Execute(strSql)
							End If

							'2. 옵션 조회로 가져오기
							strParam = ""
							strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
							Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If

						ElseIf oAuction.FOneItem.FOptioncnt = 0 AND oAuction.FOneItem.FRegedoptcnt > 0 Then	'텐바이텐상품이 옵션있음에서 단품으로 변경되고, 등록된 옵션이 있는 상태
							'1.옵션을 완전 초기화
							strParam = ""
							strParam = oAuction.FOneItem.getAuctionOPTDeleteParameter()
							Call fnAuctionOPTDel(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If

							'2.옵션 조회로 가져오기
							strParam = ""
							strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
							Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						End If
					End If

					strParam = ""
					strParam = getAuctionInfoCdParameter(itemid, oAuction.FOneItem.FAuctionGoodNo)
					Call fnAuctionItemInfoCd(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'OK던 ERR이던 editQuecnt에 + 1을 시킴..
					'스케줄링에서 editQuecnt ASC, i.lastupdate DESC로 중복을 막자
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_auction_regItem SET " & VBCRLF
					strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
					strSql = strSql & " ,AuctionLastUpdate = getdate()  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql
				End If

				If failCnt > 0 Then
					SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
					SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
					CALL Fn_AcctFailTouch("auction1010", itemid, SumErrStr)
					Call SugiQueLogInsert("auction1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
					iErrStr = "ERR||"&itemid&"||"&SumErrStr
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_auction_regItem SET " & VBCRLF
					strSql = strSql & " accFailcnt = 0  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql

					SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
					Call SugiQueLogInsert("auction1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
					iErrStr = "OK||"&itemid&"||"&SumOKStr
				End If
			End If
		SET oAuction = nothing
	End If
ElseIf action = "auctionCommonCode" Then
	Dim isday
	If ccd = "GetNationCode" Then
		strParam = ""
		strParam = getAuctionCommonCodeList(ccd)
	ElseIf ccd = "GetShippingPlaceCode" Then
		strParam = ""
		strParam = getAuctionCommonCodeShipPlace(ccd)
	ElseIf ccd = "GetPaidOrderList" Then
		strSql = ""
	    strSql = strSql&"select top 1 convert(varchar(10),selldate,21) as lastOrdInputDt"
	    strSql = strSql&" from db_temp.dbo.tbl_XSite_TMpOrder"
	    strSql = strSql&" where sellsite='auction1010'"
	    strSql = strSql&" order by selldate desc"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		if (Not rsget.Eof) then
			isday = rsget("lastOrdInputDt")
		end if
		rsget.Close
	ElseIf ccd = "GetDeliveryList" Then
		strParam = ""
		strParam = getAuctionCommonCodeGetDeliveryList(ccd)
	ElseIf ccd = "GetDeliveryPrepareList" Then
		strParam = ""
		strParam = getAuctionOrderList2(ccd,isday)
	End If
	Call fnAuctionCommonCode(ccd, strParam)
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
'###################################################### LotteiMall API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->