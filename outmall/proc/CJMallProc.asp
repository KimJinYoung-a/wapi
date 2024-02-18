<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/cjmall/cjmallItemcls.asp"-->
<!-- #include virtual="/outmall/cjmall/inccjmallFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, oCJMall, failCnt, chgSellYn, arrRows, skipItem, sellgubun, getMustprice, chkXML
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, optReset, optString, sellmoney
Dim jenkinsBatchYn, idx, lastErrStr
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
chkXML			= request("chkXML")
failCnt			= 0
jenkinsBatchYn	= request("jenkinsBatchYn")
idx				= request("idx")
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
'######################################################## CJMall API ########################################################
If mallid = "cjmall" Then
	If action = "SOLDOUT" Then			'품절처리
		strParam = ""
		strParam = getCJMallSellynParameter(getCjmallPrdNo(itemid), "N")
		Call fnCJMallSellyn(itemid, "N", strParam, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("cjmall", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/CJMallProc.asp?itemid=1235519&mallid=cjmall&action=SOLDOUT
	ElseIf action = "EDIT" Then			'상품수정
		SET oCJMall = new CCJMall
			oCJMall.FRectItemID	= itemid
			oCJMall.getCJMallNotEditOneItem
		    If (oCJMall.FResultCount < 1) Then
				iErrstr = "ERR||"&itemid&"||수정가능한 상품이 아닙니다."
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
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

'			        If (oCJMall.FOneItem.FSellcash <> oCJMall.FOneItem.Fcjmallprice) Then
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
'					End If

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
			End If
			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("cjmall", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_cjmall_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oCJMall = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/CJMallProc.asp?itemid=1235519&mallid=cjmall&action=EDIT
	ElseIf action = "PRICE" Then		'가격수정
		SET oCJMall = new CCJMall
			oCJMall.FRectItemID	= itemid
			oCJMall.getCJMallNotEditOneItem
			If oCJMall.FResultCount > 0 Then
				strParam = ""
				strParam = oCJMall.FOneItem.getCJMallPriceParameter()
				Call fnCJMallOptionSellPriceEdit(itemid, iErrStr, strParam)
				lastErrStr = iErrStr
				response.write iErrStr
				If LEFT(iErrStr, 2) <> "OK" Then
					CALL Fn_AcctFailTouch("cjmall", itemid, iErrStr)
				End If
			End If
		SET oCJMall = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/CJMallProc.asp?itemid=1235519&mallid=cjmall&action=PRICE
	ElseIf action = "REG" Then			'상품등록
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

					Call fnCJMallItemReg(itemid, strParam, iErrStr, sellmoney, oCJMall.FOneItem.getCjmallSellYn, oCJMall.FOneItem.FLimityn, oCJMall.FOneItem.FLimitNo, oCJMall.FOneItem.FLimitSold, html2db(oCJMall.FOneItem.FItemName))
				Else
					iErrStr = "ERR||"&itemid&"||옵션검사 실패"
				End If
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("cjmall", itemid, iErrStr)
			End If
		SET oCJMall = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/CJMallProc.asp?itemid=1235519&mallid=cjmall&action=REG
	ElseIf (action = "CHKSTAT") or (action = "CONFIRM") Then		'승인조회
		strParam = ""
		strParam = getCJMallStatChkParameter(itemid)

		Call fnCJMallStatChk(itemid, strParam, iErrStr, chkXML)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("cjmall", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/CJMallProc.asp?itemid=1235519&mallid=cjmall&action=CHKSTAT
	End If
End If
'###################################################### CJMall API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->