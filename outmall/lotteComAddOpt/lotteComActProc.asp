<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/lotteComAddOpt/lotteItemcls.asp"-->
<!-- #include virtual="/outmall/lotteComAddOpt/incLotteFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/lotteComAddOpt/inc_dailyAuthCheck.asp"-->
<%
Dim idx, action, oLotteitem, failCnt, chgSellYn, arrRows, skipItem
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname
idx				= requestCheckVar(request("idx"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
failCnt			= 0

If Not(isNumeric(idx)) Then
	response.write "<script>alert('잘못된 상품번호입니다.')</script>"
	response.end
End If
'######################################################## LotteCom API ########################################################
If action = "EditSellYn" Then								'상태변경
	strParam = ""
	strParam = getLotteComSellynParameter(chgSellYn, getLotteGoodno(idx))
	Call fnLotteComSellyn(idx, chgSellYn, strParam, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchOption("lotteCom", idx, iErrStr)
	End If
	Call SugiOptionQueLogInsert("lotteCom", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "PRICE" Then								'가격수정
	strParam = ""
	strParam = getLotteComPriceParameter(idx, getLotteGoodno(idx), mustPrice)
	If strParam = "" Then
		iErrStr = "ERR||"&idx&"||가격수정 할 상품이 등록되어 있지 않습니다."
	Else
		Call fnLotteComPrice(idx, strParam, mustPrice, iErrStr)
		'response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouchOption("lotteCom", idx, iErrStr)
		End If
		Call SugiOptionQueLogInsert("lotteCom", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "ITEMNAME" Then								'상품명수정
	strParam = ""
	strParam = getLotteItemnameParameter(idx, iitemname, getLotteGoodno(idx))
	Call fnLotteComChgItemname(idx, strParam, iErrStr)
	'response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchOption("lotteCom", idx, iErrStr)
	End If
	Call SugiOptionQueLogInsert("lotteCom", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKSTAT" Then								'신규상품조회
	Call fnLotteComStatChk(idx, iErrStr)
	'response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchOption("lotteCom", idx, iErrStr)
	End If
	Call SugiOptionQueLogInsert("lotteCom", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDIT" Then									'재고조회 + 상품정보 + 가격 + 필요에 따라 상품판매상태수정
	SET oLotteitem = new CLotte
		oLotteitem.FRectIdx = idx
		oLotteitem.getLotteEditOneItem
		If oLotteitem.FResultCount > 0 Then
			If (oLotteitem.FOneItem.FmaySoldOut = "Y") OR (oLotteitem.FOneItem.IsOptionSoldOut) OR (oLotteitem.FOneItem.isDiffName) Then
				strParam = ""
				strParam = getLotteComSellynParameter("N", getLotteGoodno(idx))

				Call fnLotteComSellyn(idx, "N", strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (oLotteitem.FOneItem.FLotteSellYn = "N" AND oLotteitem.FOneItem.FmaySoldOut = "N" AND oLotteitem.FOneItem.IsOptionSoldOut = False) Then
					iErrStr = ""
					strParam = ""
					strParam = getLotteComSellynParameter("Y", getLotteGoodno(idx))
					Call fnLotteComSellyn(idx, "Y", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strParam = ""
				strParam = getLotteComPriceParameter(idx, getLotteGoodno(idx), mustPrice)
				If strParam = "" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & "ERR||"&idx&"||가격수정 할 상품이 등록되어 있지 않습니다."
				Else
					Call fnLotteComPrice(idx, strParam, mustPrice, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strParam = ""
				strParam = oLotteitem.FOneItem.getLotteComItemEditParameter()
				Call fnLotteComInfoEdit(idx, strParam, iErrStr, FALSE)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				If oLotteitem.FOneItem.isImageChanged Then
					strParam = ""
					strParam = oLotteitem.FOneItem.getLotteItemImageEdit()
					Call fnLotteComImageEdit(idx, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				'전시상품 조회해서 현재 상품상태를 가져오기
				Call fnCheckLotteComItemStat(idx, iErrStr, getLotteGoodno(idx))
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'OK던 ERR이던 editQuecnt에 + 1을 시킴..
				'스케줄링에서 editQuecnt ASC, i.lastupdate DESC로 중복을 막자
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_lotteAddOption_regItem] SET " & VBCRLF
				strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
				strSql = strSql & " ,LotteLastUpdate = getdate()  " & VBCRLF
				strSql = strSql & " WHERE midx = '"&idx&"' " & VBCRLF
				dbget.Execute strSql
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&idx&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&idx&"||", "")
				CALL Fn_AcctFailTouchOption("lotteCom", idx, SumErrStr)
				Call SugiOptionQueLogInsert("lotteCom", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
				iErrStr = "ERR||"&idx&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_lotteAddOption_regItem] SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE midx = '"&idx&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&idx&"||", "")
				Call SugiOptionQueLogInsert("lotteCom", action, idx, "OK", "OK||"&idx&"||"&SumOKStr, session("ssBctID"))
				iErrStr = "OK||"&idx&"||"&SumOKStr
			End If
		End If
	SET oLotteitem = nothing
ElseIf action = "EDIT2" Then								'승인예정상품 수정
	SET oLotteitem = new CLotte
		oLotteitem.FRectIdx = idx
		oLotteitem.getLotteEditOneItem

		strParam = ""
		strParam = oLotteitem.FOneItem.getLotteComItemEditParameter2()
		Call fnLotteComInfoEdit(idx, strParam, iErrStr, TRUE)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouchOption("lotteCom", idx, iErrStr)
		End If
		Call SugiOptionQueLogInsert("lotteCom", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oLotteitem = nothing
ElseIf action = "REG" Then									'상품 등록
	SET oLotteitem = new CLotte
		oLotteitem.FRectIdx = idx
		oLotteitem.getLotteNotRegOneItem

	    If (oLotteitem.FResultCount < 1) Then
			iErrStr = "ERR||"&idx&"||등록가능한 상품이 아닙니다."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT TOP 1 * FROM db_etcmall.[dbo].[tbl_lotteAddOption_regItem] where midx="&idx&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.[dbo].[tbl_lotteAddOption_regItem] "
	        strSql = strSql & " 	(midx, regdate, reguserid, LotteStatCd)"
	        strSql = strSql & " 	VALUES ("&idx&", getdate(), '"&session("SSBctID")&"', '10')"
			strSql = strSql & " END "
		    dbget.Execute strSql

			strParam = ""
			strParam = oLotteitem.FOneItem.getLotteComItemRegParameter(FALSE)
			Call fnLotteComItemReg(oLotteitem.FOneItem.Fitemid, strParam, iErrStr, oLotteitem.FOneItem.FRealSellprice, oLotteitem.FOneItem.getLotteSellYn, idx, oLotteitem.FOneItem.FbasicimageNm)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouchOption("lotteCom", idx, iErrStr)
			End If
			Call SugiOptionQueLogInsert("lotteCom", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oLotteitem = nothing
ElseIf action = "INFODIV" Then
	SET oLotteitem = new CLotte
		oLotteitem.FRectIdx	= idx
		oLotteitem.getLotteEditOneItem

		strParam = ""
		strParam = oLotteitem.FOneItem.getLotteItemInfoCdToEdt()
		Call fnLotteComInfodivEdit(idx, strParam, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouchOption("lotteCom", idx, iErrStr)
		End If
		Call SugiOptionQueLogInsert("lotteCom", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oLotteitem = nothing
ElseIf action = "IMAGE" Then
	SET oLotteitem = new CLotte
		oLotteitem.FRectIdx	= idx
		oLotteitem.getLotteEditOneItem
		If oLotteitem.FResultCount > 0 Then
			strParam = ""
			strParam = oLotteitem.FOneItem.getLotteItemImageEdit()
			Call fnLotteComImageEdit(idx, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouchOption("lotteCom", idx, iErrStr)
			End If
			Call SugiOptionQueLogInsert("lotteCom", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oLotteitem = nothing
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