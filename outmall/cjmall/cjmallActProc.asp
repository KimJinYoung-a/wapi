<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/cjmall/cjmallItemcls.asp"-->
<!-- #include virtual="/outmall/cjmall/inccjmallFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, oCJMall, failCnt, chgSellYn, arrRows, skipItem, sellgubun, getMustprice, sellmoney, chkXML
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, chkparam, optReset, optString
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
chkXML			= request("chkXML")
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
'######################################################## Cjmall API ########################################################
If action = "EditSellYn" Then								'상태변경
	strParam = ""
	strParam = getCJMallSellynParameter(getCjmallPrdNo(itemid), chgSellYn)
	Call fnCJMallSellyn(itemid, chgSellYn, strParam, iErrStr)
'	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("cjmall", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("cjmall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "PRICE" Then								'단품 가격 수정
	SET oCJMall = new CCJMall
		oCJMall.FRectItemID	= itemid
		oCJMall.getCJMallNotEditOneItem
		If oCJMall.FResultCount > 0 Then
			strParam = ""
			strParam = oCJMall.FOneItem.getCJMallPriceParameter()
'			If itemid = 1303019 Then
'				response.write strParam
'				response.end
'			End If
			Call fnCJMallOptionSellPriceEdit(itemid, iErrStr, strParam)
		Else
			iErrstr = "ERR||"&itemid&"||가격 수정 가능한 상품이 아닙니다."
		End If

		If (LEFT(iErrStr, 2) <> "OK") AND (LEFT(iErrStr, 3) <> "ERR") Then
			iErrstr = "ERR||"&itemid&"||전문형식 잘못됨..kjy(가격수정)"
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("cjmall", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("cjmall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oCJMall = nothing
ElseIf action = "QTY" Then									'단품 수량 수정
	SET oCJMall = new CCJMall
		oCJMall.FRectItemID	= itemid
		oCJMall.getCJMallNotEditOneItem
		If oCJMall.FResultCount > 0 Then
			strParam = ""
			strParam = oCJMall.FOneItem.getCJMallQTYParameter()
			Call fnCJMallOptionQTYEdit(itemid, iErrStr, strParam)
		Else
			iErrstr = "ERR||"&itemid&"||수량 수정 가능한 상품이 아닙니다."
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("cjmall", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("cjmall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oCJMall = nothing
ElseIf action = "OPTSTAT" Then								'단품 상태 수정
	SET oCJMall = new CCJMall
		oCJMall.FRectItemID	= itemid
		oCJMall.getCJMallNotEditOneItem
		If oCJMall.FResultCount > 0 Then
			strParam = ""
			strParam = oCJMall.FOneItem.getcjmallOptSellModParameter()
			Call fnCJMallOptSellEdit(itemid, iErrStr, strParam, oCJMall.FOneItem.FMaySoldout)
		Else
			iErrstr = "ERR||"&itemid&"||상태 수정 가능한 상품이 아닙니다."
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("cjmall", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("cjmall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oCJMall = nothing
ElseIf action = "MOD" Then									'정보 수정
	SET oCJMall = new CCJMall
		oCJMall.FRectItemID	= itemid
		oCJMall.getCJMallNotEditOneItem
		If oCJMall.FResultCount > 0 Then
			strParam = ""
			strParam = oCJMall.FOneItem.getcjmallItemModXML()
response.write replace(strParam, "<?xml", "<aaaaaaa")
response.end
			Call fnCJMallOneItemEdit(itemid, iErrStr, strParam)
		Else
			iErrstr = "ERR||"&itemid&"||정보 수정 가능한 상품이 아닙니다."
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("cjmall", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("cjmall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oCJMall = nothing
ElseIf action = "REG" Then									'상품 등록
	SET oCJMall = new CCJMall
		oCJMall.FRectItemID	= itemid
		oCJMall.getCJMallNotRegOneItem

	    If (oCJMall.FResultCount < 1) Then
			iErrstr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			If (oCJMall.FOneItem.FCddKey = "") Then
				iErrstr = "ERR||"&itemid&"||상품분류 미매칭"
			End If

			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_cjmall_regItem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_item.dbo.tbl_cjmall_regItem "
	        strSql = strSql & " (itemid, regdate, reguserid, cjmallstatCD, regitemname)"
	        strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oCJMall.FOneItem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oCJMall.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oCJMall.FOneItem.getCjmallItemRegXML()

				If CLng(10000 - oCJMall.FOneItem.Fbuycash / oCJMall.FOneItem.Fsellcash * 100 * 100) / 100 < 15 Then
					sellmoney = oCJMall.FOneItem.Forgprice
				Else
					sellmoney = oCJMall.FOneItem.Fsellcash
				End If

				If chkXML = "Y" Then
					response.write replace(strParam, "<?xml", "<aaaaaaa")
					response.end
				End If

				Call fnCJMallItemReg(itemid, strParam, iErrStr, sellmoney, oCJMall.FOneItem.getCjmallSellYn, oCJMall.FOneItem.FLimityn, oCJMall.FOneItem.FLimitNo, oCJMall.FOneItem.FLimitSold, html2db(oCJMall.FOneItem.FItemName))
			Else
				iErrStr = "ERR||"&itemid&"||옵션검사 실패"
			End If

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("cjmall", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("cjmall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oCJMall = nothing
ElseIf action = "EDIT" Then
	SET oCJMall = new CCJMall
		oCJMall.FRectItemID	= itemid
		oCJMall.getCJMallNotEditOneItem
		If oCJMall.FResultCount > 0 Then
			If (oCJMall.FOneItem.FmaySoldOut = "Y")  OR (oCJMall.FOneItem.IsSoldOutLimit5Sell) OR (oCJMall.FOneItem.IsMayLimitSoldout = "Y") Then
				strParam = ""
				strParam = getCJMallSellynParameter(getCjmallPrdNo(itemid), "N")
				Call fnCJMallSellyn(itemid, "N", strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (oCJMall.FOneItem.FcjmallSellYn = "N" AND oCJMall.FOneItem.IsSoldOut = False) Then
					iErrStr = ""
					strParam = ""
					strParam = getCJMallSellynParameter(getCjmallPrdNo(itemid), "Y")
					Call fnCJMallSellyn(itemid, "Y", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strParam = ""
				strParam = oCJMall.FOneItem.getcjmallItemModXML()
				Call fnCJMallOneItemEdit(itemid, iErrStr, strParam)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				strParam = ""
				strParam = getCJMallStatChkParameter(itemid)
				Call fnCJMallStatChk(itemid, strParam, iErrStr, chkXML)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

'		        If (oCJMall.FOneItem.FSellcash <> oCJMall.FOneItem.Fcjmallprice) Then			'2016-03-31 김진영 주석처리..위에 상품조회에서 가져왔을 떄 엄한 가격을 가져왔다면 이 조건이 통과될 수 있음..
					strParam = ""
					strParam = oCJMall.FOneItem.getCJMallPriceParameter()
					If strParam = "" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다."
					Else
						Call fnCJMallOptionSellPriceEdit(itemid, iErrStr, strParam)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
'				End If

				strParam = ""
				strParam = oCJMall.FOneItem.getcjmallOptSellModParameter()
				Call fnCJMallOptSellEdit(itemid, iErrStr, strParam, oCJMall.FOneItem.FMaySoldout)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				If oCJMall.FOneItem.IsRegedOptionSellyn = "N" OR oCJMall.FOneItem.FmaySoldOut = "Y" Then
					strParam = ""
					strParam = getCJMallSellynParameter(getCjmallPrdNo(itemid), "N")
					Call fnCJMallSellyn(itemid, "N", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					strParam = ""
					strParam = oCJMall.FOneItem.getCJMallQTYParameter()
					Call fnCJMallOptionQTYEdit(itemid, iErrStr, strParam)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_cjmall_regItem SET " & VBCRLF
				strSql = strSql & " cjmallLastUpdate = getdate()  " & VBCRLF
				strSql = strSql & " ,lastStatCheckDate = getdate() " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("cjmall", itemid, SumErrStr)
				Call SugiQueLogInsert("cjmall", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_cjmall_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				Call SugiQueLogInsert("cjmall", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))

				iErrStr = "OK||"&itemid&"||"&SumOKStr
			End If
		Else
			iErrstr = "ERR||"&itemid&"||전체 수정 가능한 상품이 아닙니다."
		End If
	SET oCJMall = nothing
ElseIf action = "NOTREGEDIT" Then
	SET oCJMall = new CCJMall
		oCJMall.FRectItemID	= itemid
		oCJMall.getCJMallNotRegEditOneItem
		If oCJMall.FResultCount > 0 Then
			If (oCJMall.FOneItem.FmaySoldOut = "Y")  OR (oCJMall.FOneItem.IsSoldOutLimit5Sell) OR (oCJMall.FOneItem.IsMayLimitSoldout = "Y") Then
				strParam = ""
				strParam = getCJMallSellynParameter(getCjmallPrdNo(itemid), "N")
				Call fnCJMallSellyn(itemid, "N", strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (oCJMall.FOneItem.FcjmallSellYn = "N" AND oCJMall.FOneItem.IsSoldOut = False) Then
					iErrStr = ""
					strParam = ""
					strParam = getCJMallSellynParameter(getCjmallPrdNo(itemid), "Y")
					Call fnCJMallSellyn(itemid, "Y", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strParam = ""
				strParam = oCJMall.FOneItem.getcjmallItemModXML()
				Call fnCJMallOneItemEdit(itemid, iErrStr, strParam)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				strParam = ""
				strParam = getCJMallStatChkParameter(itemid)
				Call fnCJMallStatChk(itemid, strParam, iErrStr, chkXML)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

'		        If (oCJMall.FOneItem.FSellcash <> oCJMall.FOneItem.Fcjmallprice) Then			'2016-03-31 김진영 주석처리..위에 상품조회에서 가져왔을 떄 엄한 가격을 가져왔다면 이 조건이 통과될 수 있음..
					strParam = ""
					strParam = oCJMall.FOneItem.getCJMallPriceParameter()
					If strParam = "" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다."
					Else
						Call fnCJMallOptionSellPriceEdit(itemid, iErrStr, strParam)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
'				End If

				strParam = ""
				strParam = oCJMall.FOneItem.getcjmallOptSellModParameter()
				Call fnCJMallOptSellEdit(itemid, iErrStr, strParam, oCJMall.FOneItem.FMaySoldout)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				If oCJMall.FOneItem.IsRegedOptionSellyn = "N" OR oCJMall.FOneItem.FmaySoldOut = "Y" Then
					strParam = ""
					strParam = getCJMallSellynParameter(getCjmallPrdNo(itemid), "N")
					Call fnCJMallSellyn(itemid, "N", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					strParam = ""
					strParam = oCJMall.FOneItem.getCJMallQTYParameter()
					Call fnCJMallOptionQTYEdit(itemid, iErrStr, strParam)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_cjmall_regItem SET " & VBCRLF
				strSql = strSql & " cjmallLastUpdate = getdate()  " & VBCRLF
				strSql = strSql & " ,lastStatCheckDate = getdate() " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("cjmall", itemid, SumErrStr)
				Call SugiQueLogInsert("cjmall", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_cjmall_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				Call SugiQueLogInsert("cjmall", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))

				iErrStr = "OK||"&itemid&"||"&SumOKStr
			End If
		Else
			iErrstr = "ERR||"&itemid&"||전체 수정 가능한 상품이 아닙니다."
		End If
	SET oCJMall = nothing


ElseIf action = "CHKSTAT" Then							'상품 조회
	strParam = ""
	strParam = getCJMallStatChkParameter(itemid)
	Call fnCJMallStatChk(itemid, strParam, iErrStr, chkXML)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("cjmall", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("cjmall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
End If

If iErrStr <> "" Then
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str += '"&replace(iErrStr, "'", "")&"<br>' " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
					"</script>"
End If
'###################################################### ezwel API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->