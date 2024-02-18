<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/lotteCom/lotteItemcls.asp"-->
<!-- #include virtual="/outmall/lotteCom/incLotteFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/lotteCom/inc_dailyAuthCheck.asp"-->
<%

Dim itemid, action, oLotteitem, failCnt, chgSellYn, arrRows, skipItem, isMayEndItem, isMayEndItem2
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, endItemErrMsgReplace
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
failCnt			= 0
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
'######################################################## LotteCom API ########################################################
If action = "EditSellYn" Then								'상태변경
	strParam = ""
	strParam = getLotteComSellynParameter(chgSellYn, getLotteGoodno(itemid))
	Call fnLotteComSellyn(itemid, chgSellYn, strParam, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("lotteCom", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "PRICE" Then								'가격수정
	SET oLotteitem = new CLotte
		oLotteitem.FRectItemID	= itemid
		oLotteitem.getLotteEditOneItem
		If oLotteitem.FResultCount > 0 Then
			strParam = ""
			mustPrice = ""
			mustPrice = oLotteitem.FOneItem.MustPrice()
			strParam = getLotteComPriceParameter(itemid, getLotteGoodno(itemid), mustPrice)
			If strParam = "" Then
				iErrStr = "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다."
			Else
				Call fnLotteComPrice(itemid, strParam, mustPrice, iErrStr)
				'response.write iErrStr
				If LEFT(iErrStr, 2) <> "OK" Then
					CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
				End If
				Call SugiQueLogInsert("lotteCom", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
			End If
		else
			iErrStr = "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다.[1]"
		end if
	SET oLotteitem = nothing
ElseIf action = "ITEMNAME" Then								'상품명수정
	strParam = ""
	strParam = getLotteItemnameParameter(itemid, iitemname, getLotteGoodno(itemid))
	Call fnLotteComChgItemname(itemid, strParam, iErrStr)
	'response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("lotteCom", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKSTAT" Then								'신규상품조회
	Call fnLotteComStatChk(itemid, iErrStr)
	'response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("lotteCom", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKSTOCK" Then								'재고조회
	Call fnLotteComStockChk(itemid, iErrStr)
	'response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("lotteCom", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CATEGORY" Then
	strParam = ""
	strParam = getLotteCategoryParameter(itemid, getLotteGoodno(itemid))
	'response.write iErrStr
	If strParam = "" Then
		iErrStr = "ERR||"&itemid&"||카테고리 수정 할 상품이 등록되어 있지 않습니다."
	Else
		Call fnLotteComCateGory(itemid, strParam, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteCom", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "EDIT" Then									'재고조회 + 상품정보 + 가격 + 필요에 따라 상품판매상태수정
	SET oLotteitem = new CLotte
		oLotteitem.FRectItemID	= itemid
		oLotteitem.getLotteEditOneItem
		If oLotteitem.FResultCount > 0 Then
			'1. 품절에 해당하면 품절처리
			If (oLotteitem.FOneItem.FmaySoldOut = "Y")  OR (oLotteitem.FOneItem.IsSoldOutLimit5Sell) OR (oLotteitem.FOneItem.IsMayLimitSoldout = "Y") Then
				strParam = ""
				strParam = getLotteComSellynParameter("N", getLotteGoodno(itemid))
				Call fnLotteComSellyn(itemid, "N", strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				'2. 재고 조회
				Call fnLotteComStockChk(itemid, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'3. 재고 조회시 판매가 불가능하면 판매종료처리
				isMayEndItem = getOptCntCompare(itemid)
				If isMayEndItem = "Y" Then
					strParam = ""
					strParam = getLotteComSellynParameter("X", getLotteGoodno(itemid))

					Call fnLotteComSellyn(itemid, "X", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					'3-1 판매가 가능한 상품이면 상품 수정
					strParam = ""
					strParam = oLotteitem.FOneItem.getLotteComItemEditParameter()
					Call fnLotteComInfoEdit(itemid, strParam, iErrStr, FALSE)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'3-2 판매로 변경
					strParam = ""
					strParam = getLotteComSellynParameter("Y", getLotteGoodno(itemid))
					Call fnLotteComSellyn(itemid, "Y", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'4. 판매가 수정
					strParam = ""
					mustPrice = ""
					mustPrice = oLotteitem.FOneItem.MustPrice()
					strParam = getLotteComPriceParameter(itemid, getLotteGoodno(itemid), mustPrice)
					If strParam = "" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다."
					Else
						Call fnLotteComPrice(itemid, strParam, mustPrice, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

					'5. 이미지 수정
					If oLotteitem.FOneItem.isImageChanged Then
						strParam = ""
						strParam = oLotteitem.FOneItem.getLotteItemImageEdit()
						Call fnLotteComImageEdit(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

					'6. 재고 조회
					Call fnLotteComStockChk(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					isMayEndItem2 = getUseOption(itemid)
					If isMayEndItem2 = "N" Then
						strParam = ""
						strParam = getLotteComSellynParameter("X", getLotteGoodno(itemid))

						Call fnLotteComSellyn(itemid, "X", strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					Else
						'전시상품 조회해서 현재 상품상태를 가져오기
						Call fnCheckLotteComItemStat(itemid, iErrStr, getLotteGoodno(itemid))
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
					strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem SET " & VBCRLF
					strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
					strSql = strSql & " ,LotteLastUpdate = getdate()  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql
				End If
			End If

			If failCnt > 0 Then
				endItemErrMsgReplace = replace(SumErrStr, "OK||"&itemid&"||", "")
				endItemErrMsgReplace = replace(SumErrStr, "ERR||"&itemid&"||", "")

				If Instr(endItemErrMsgReplace, "단품이 모두 품절 또는 판매종료인 경우 상품의 판매상태를 판매중으로 변경할 수 없습니다.") > 0 Then
					strParam = ""
					strParam = getLotteComSellynParameter("X", getLotteGoodno(itemid))

					Call fnLotteComSellyn(itemid, "X", strParam, iErrStr)
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
				CALL Fn_AcctFailTouch("lotteCom", itemid, SumErrStr)
				Call SugiQueLogInsert("lotteCom", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				Call SugiQueLogInsert("lotteCom", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))

				iErrStr = "OK||"&itemid&"||"&SumOKStr
			End If
		End If
	SET oLotteitem = nothing
ElseIf action = "EDIT2" Then								'승인예정상품 수정
	SET oLotteitem = new CLotte
		oLotteitem.FRectItemID	= itemid
		oLotteitem.getLotteEditOneItem

		strParam = ""
		strParam = oLotteitem.FOneItem.getLotteComItemEditParameter2()
		Call fnLotteComInfoEdit(itemid, strParam, iErrStr, TRUE)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteCom", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oLotteitem = nothing
ElseIf action = "REG" Then									'상품 등록
	SET oLotteitem = new CLotte
		oLotteitem.FRectItemID	= itemid
		oLotteitem.getLotteNotRegOneItem
	    If (oLotteitem.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_lotte_regItem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_item.dbo.tbl_lotte_regItem "
	        strSql = strSql & " 	(itemid, regdate, reguserid, LotteStatCd)"
	        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '10')"
			strSql = strSql & " END "
		    dbget.Execute strSql
			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oLotteitem.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oLotteitem.FOneItem.getLotteComItemRegParameter(FALSE)
				Call fnLotteComItemReg(itemid, strParam, iErrStr, oLotteitem.FOneItem.FSellCash, oLotteitem.FOneItem.getLotteSellYn, oLotteitem.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||옵션검사 실패"
			End If

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("lotteCom", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oLotteitem = nothing
ElseIf action = "INFODIV" Then
	SET oLotteitem = new CLotte
		oLotteitem.FRectItemID	= itemid
		oLotteitem.getLotteEditOneItem

		strParam = ""
		strParam = oLotteitem.FOneItem.getLotteItemInfoCdToEdt()
		Call fnLotteComInfodivEdit(itemid, strParam, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteCom", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oLotteitem = nothing
ElseIf action = "IMAGE" Then
	SET oLotteitem = new CLotte
		oLotteitem.FRectItemID	= itemid
		oLotteitem.getLotteEditOneItem
		If oLotteitem.FResultCount > 0 Then
			strParam = ""
			strParam = oLotteitem.FOneItem.getLotteItemImageEdit()
			Call fnLotteComImageEdit(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("lotteCom", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oLotteitem = nothing
ELSEIF action = "CHKITEMLIST" Then
	Call fnLotteGoodsList(replace(request("yyyymmdd"),"-",""),replace(request("yyyymmdd"),"-",""), iErrStr)

	rw iErrStr
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
