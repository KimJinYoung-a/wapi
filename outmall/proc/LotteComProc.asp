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
Dim itemid, mallid, action, oLotteitem, failCnt, chgSellYn, arrRows, skipItem, assin, isMayEndItem, isMayEndItem2
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, mode, tLotteGoodno
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
mode			= request("mode")

If mode = "updateSendState" Then
	strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendState='"&request("updateSendState")&"'"
	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
	strSql = strSql & "	where OutMallOrderSerial='"&request("ORG_ord_no")&"'"
	strSql = strSql & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	dbget.Execute strSql,assin
	response.write "<script>alert('"&assin&"건 완료 처리.');opener.close();window.close()</script>"
	response.end
ElseIf mode = "etcSongjangFin" Then
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendState=7"
	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
    strSql = strSql & "	where OutMallOrderSerial='"&request("ORG_ord_no")&"'"
    strSql = strSql & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
    dbget.Execute strSql,assin
    response.write "<script>alert('"&assin&"건 완료 처리.');opener.close();window.close()</script>"
Else
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
'######################################################## LotteCom API ########################################################

If action = "SOLDOUT" Then								'상태변경
	strParam = ""
	strParam = getLotteComSellynParameter("N", getLotteGoodno(itemid))
	Call fnLotteComSellyn(itemid, "N", strParam, iErrStr)
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/LotteComProc.asp?itemid=279397&mallid=lotteCom&action=SOLDOUT
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
				response.write "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다."
			Else
				Call fnLotteComPrice(itemid, strParam, mustPrice, iErrStr)
				response.write iErrStr
			End If
		else
			response.write "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다.[1]"
		end if
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
		End If
	SET oLotteitem = nothing
	'http://wapi.10x10.co.kr/outmall/proc/LotteComProc.asp?itemid=279397&mallid=lotteCom&action=PRICE
ElseIf action = "ITEMNAME" Then								'상품명수정
	strParam = ""
	strParam = getLotteItemnameParameter(itemid, iitemname, getLotteGoodno(itemid))
	Call fnLotteComChgItemname(itemid, strParam, iErrStr)
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/LotteComProc.asp?itemid=279397&mallid=lotteCom&action=ITEMNAME
ElseIf action = "CHKSTAT" Then								'신규상품조회
	Call fnLotteComStatChk(itemid, iErrStr)
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/LotteComProc.asp?itemid=279397&mallid=lotteCom&action=CHKSTAT
ElseIf action = "CHKSTOCK" Then								'재고조회
	Call fnLotteComStockChk(itemid, iErrStr)
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/LotteComProc.asp?itemid=279397&mallid=lotteCom&action=CHKSTOCK
ElseIf action = "EDIT" Then									'재고조회 + 상품정보 + 가격 + 필요에 따라 상품판매상태수정
	SET oLotteitem = new CLotte
		oLotteitem.FRectItemID	= itemid
		oLotteitem.getLotteEditOneItem
		If oLotteitem.FResultCount > 0 Then
			'1. 품절에 해당하면 품절처리
			If (oLotteitem.FOneItem.FmaySoldOut = "Y") OR (oLotteitem.FOneItem.IsSoldOutLimit5Sell) OR (oLotteitem.FOneItem.IsMayLimitSoldout = "Y") Then
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
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("lotteCom", itemid, SumErrStr)
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				response.write "OK||"&itemid&"||"&SumOKStr
			End If

		End If
	SET oLotteitem = nothing
ElseIf action = "REG" Then									'상품 등록
	SET oLotteitem = new CLotte
		oLotteitem.FRectItemID	= itemid
		oLotteitem.getLotteNotRegOneItem

		tLotteGoodno = getLotteGoodno(itemid)
		If tLotteGoodno <> "" Then
			iErrStr = "ERR||"&itemid&"||이미 등록된 상품 입니다."
	    ElseIf (oLotteitem.FResultCount < 1) Then
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
		End If
	SET oLotteitem = nothing
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/LotteComProc.asp?itemid=1860480&mallid=lotteCom&action=REG
End If
'###################################################### LotteCom API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
