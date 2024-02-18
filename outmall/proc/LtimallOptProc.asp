<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/ltimallAddOpt/LtimallItemcls.asp"-->
<!-- #include virtual="/outmall/ltimallAddOpt/incLtimallFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/ltimallAddOpt/inc_dailyAuthCheck.asp"-->
<%
Dim idx, mallid, action, oiMall, failCnt, chgSellYn, arrRows
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, mode
idx			= requestCheckVar(request("idx"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
mode			= request("mode")

If Not(isNumeric(idx)) Then
	response.write "<script>alert('잘못된 상품번호입니다.')</script>"
	response.end
End If
'######################################################## LotteCom API ########################################################

If action = "SOLDOUT" Then								'상태변경
	strParam = ""
	strParam = getLtiMallSellynParameter("N", getLtimallGoodno(idx))
	Call fnLtiMallSellyn(idx, "N", strParam, iErrStr)
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchOption("lotteimall", idx, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/ltimallOptProc.asp?idx=7312&mallid=lotteimall&action=SOLDOUT
ElseIf action = "PRICE" Then								'가격수정
	strParam = ""
	strParam = getLtiMallPriceParameter(idx, getLtimallGoodno(idx), mustPrice)
	If strParam = "" Then
		response.write "ERR||"&idx&"||가격수정 할 상품이 등록되어 있지 않습니다."
	Else
		Call fnLtimallPrice(idx, strParam, mustPrice, iErrStr)
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouchOption("lotteimall", idx, iErrStr)
		End If
	End If
	'http://wapi.10x10.co.kr/outmall/proc/ltimallOptProc.asp?idx=7312&mallid=lotteimall&action=PRICE
ElseIf action = "CHKSTAT" Then								'신규상품조회
	Call fnLtiMallstatChk(idx, iErrStr)
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchOption("lotteimall", idx, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/ltimallOptProc.asp?idx=7312&mallid=lotteimall&action=CHKSTAT
ElseIf action = "EDIT" Then									'재고조회 + 상품정보 + 가격 + 필요에 따라 상품판매상태수정
	SET oiMall = new CLotteiMall
		oiMall.FRectIdx = idx
		oiMall.getLtimallEditOneItem
		If oiMall.FResultCount > 0 Then
			If (oiMall.FOneItem.FmaySoldOut = "Y") OR (oiMall.FOneItem.IsOptionSoldOut) OR (oiMall.FOneItem.isDiffName) Then
				strParam = ""
				strParam = getLtiMallSellynParameter("N", getLtimallGoodno(idx))
				Call fnLtiMallSellyn(idx, "N", strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (oiMall.FOneItem.FLtimallSellYn = "N" AND oiMall.FOneItem.FmaySoldOut = "N" AND oiMall.FOneItem.IsOptionSoldOut = False) Then
					iErrStr = ""
					strParam = ""
					strParam = getLtiMallSellynParameter("Y", getLtimallGoodno(idx))
					Call fnLtiMallSellyn(idx, "Y", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strParam = ""
				strParam = getLtiMallPriceParameter(idx, getLtimallGoodno(idx), mustPrice)
				If strParam = "" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & "ERR||"&idx&"||가격수정 할 상품이 등록되어 있지 않습니다."
				Else
					Call fnLtimallPrice(idx, strParam, mustPrice, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strParam = ""
				strParam = oiMall.FOneItem.getLotteiMallItemEditParameter()
				Call fnLtiMallInfoEdit(idx, strParam, iErrStr, FALSE)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'전시상품 조회해서 현재 상품상태를 가져오기
				Call fnLtiMallDisView(idx, iErrStr, getLtimallGoodno(idx))
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If


				'OK던 ERR이던 editQuecnt에 + 1을 시킴..
				'스케줄링에서 editQuecnt ASC, i.lastupdate DESC로 중복을 막자
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] SET " & VBCRLF
				strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
				strSql = strSql & " ,LTiMallLastupdate = getdate()  " & VBCRLF
				strSql = strSql & " WHERE midx = '"&idx&"' " & VBCRLF
				dbget.Execute strSql
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&idx&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&idx&"||", "")
				CALL Fn_AcctFailTouch("lotteimall", idx, SumErrStr)
				response.write "ERR||"&idx&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE midx = '"&idx&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&idx&"||", "")
				response.write "OK||"&idx&"||"&SumOKStr
			End If
		End If
	SET oiMall = nothing
	'http://wapi.10x10.co.kr/outmall/proc/ltimallOptProc.asp?idx=7312&mallid=lotteimall&action=EDIT
End If
'###################################################### LotteiMall API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->