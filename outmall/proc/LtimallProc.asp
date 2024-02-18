<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/ltimall/LtimallItemcls.asp"-->
<!-- #include virtual="/outmall/ltimall/incLtimallFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<!-- #include virtual="/outmall/ltimall/inc_dailyAuthCheck.asp"-->
<%
Dim itemid, mallid, action, oiMall, failCnt, arrRows, skipItem
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, iAddOptCnt, mode, assin
Dim jenkinsBatchYn, idx, lastErrStr
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
mode			= request("mode")
jenkinsBatchYn	= request("jenkinsBatchYn")
idx				= request("idx")
lastErrStr		= ""
If mode = "updateSendState" Then
	strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendState='"&request("updateSendState")&"'"
	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
	strSql = strSql & "	where OutMallOrderSerial='"&request("ord_no")&"'"
	strSql = strSql & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	dbget.Execute strSql,assin
	response.write "<script>alert('"&assin&"건 완료 처리.');opener.close();window.close()</script>"
	response.end
ElseIf mode = "etcSongjangFin" Then
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendState=7"
	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
    strSql = strSql & "	where OutMallOrderSerial='"&request("ord_no")&"'"
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
'######################################################## LotteiMall API ########################################################
If mallid = "lotteimall" Then
	If action = "SOLDOUT" Then									'상태변경
		strParam = ""
		strParam = getLtiMallSellynParameter("N", getLtimallGoodno(itemid))
		Call fnLtiMallSellyn(itemid, "N", strParam, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=SOLDOUT
	ElseIf action = "PRICE" Then								'가격수정
		SET oiMall = new CLotteiMall
			oiMall.FRectItemID	= itemid
			oiMall.getLtimallEditOneItem
			If oiMall.FResultCount > 0 Then
				mustPrice = ""
				mustPrice = oiMall.FOneItem.MustPrice()
				strParam = ""
				strParam = getLtiMallPriceParameter(itemid, getLtimallGoodno(itemid), mustPrice)
				If strParam = "" Then
					lastErrStr = "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다."
					response.write "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다."
				Else
					Call fnLtimallPrice(itemid, strParam, mustPrice, iErrStr)
					lastErrStr = iErrStr
					response.write iErrStr
				End If
			else
				lastErrStr = "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다.[1]"
				response.write "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다.[1]"
			end if

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
			End If
		SET oiMall = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=PRICE
	ElseIf action = "ITEMNAME" Then								'상품명수정
		strParam = ""
		strParam = getLtiMallItemnameParameter(itemid, iitemname, getLtimallGoodno(itemid))
		Call fnLtiMallChgItemname(itemid, strParam, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=ITEMNAME
	ElseIf action = "CHKSTAT" Then								'신규상품조회
		Call fnLtiMallstatChk(itemid, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=CHKSTAT
	ElseIf action = "CHKSTOCK" Then								'재고조회
		Call fnLtiMallStockChk(itemid, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=CHKSTOCK
	ElseIf action = "DISPVIEW" Then
		Call fnLtiMallDisView(itemid, iErrStr, getLtimallGoodno(itemid))
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=DISPVIEW
	ElseIf action = "EDIT" Then									'재고조회 + 상품정보 + 가격 + 필요에 따라 상품판매상태수정(2015-10-06 김진영 전시상태조회 주석처리)
		SET oiMall = new CLotteiMall
			oiMall.FRectItemID	= itemid
			oiMall.getLtimallEditOneItem
			If oiMall.FResultCount > 0 Then
				If (oiMall.FOneItem.FmaySoldOut = "Y") OR (oiMall.FOneItem.IsSoldOutLimit5Sell) OR (oiMall.FOneItem.IsMayLimitSoldout = "Y") Then
					strParam = ""
					strParam = getLtiMallSellynParameter("N", getLtimallGoodno(itemid))
					Call fnLtiMallSellyn(itemid, "N", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				ElseIf oiMall.FOneItem.isDuppOptionItemYn = "Y" Then
					strParam = ""
					strParam = getLtiMallSellynParameter("X", getLtimallGoodno(itemid))

					Call fnLtiMallSellyn(itemid, "X", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					If (oiMall.FOneItem.FLtimallSellYn = "N" AND oiMall.FOneItem.IsSoldOut = False) Then
						iErrStr = ""
						strParam = ""
						strParam = getLtiMallSellynParameter("Y", getLtimallGoodno(itemid))
						Call fnLtiMallSellyn(itemid, "Y", strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

			        If (oiMall.FOneItem.FSellcash <> oiMall.FOneItem.FLtiMallPrice) Then
						mustPrice = ""
						mustPrice = oiMall.FOneItem.MustPrice()
						strParam = ""
						strParam = getLtiMallPriceParameter(itemid, getLtimallGoodno(itemid), mustPrice)
						If strParam = "" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다."
						Else
							Call fnLtimallPrice(itemid, strParam, mustPrice, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						End If
					End If

					'2013-07-01 전시상품 단품추가 될 경우 추가
					Dim dp, aoptNm, aoptDc
					strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_ltimall '"&CMALLNAME&"'," & itemid
					rsget.CursorLocation = adUseClient
					rsget.CursorType = adOpenStatic
					rsget.LockType = adLockOptimistic
					rsget.Open strSql, dbget
					If Not(rsget.EOF or rsget.BOF) Then
					    arrRows = rsget.getRows
					End If
					rsget.close

					''추가된 옵션 등록
					If isArray(arrRows) Then
						For dp = 0 To UBound(ArrRows, 2)
							If (ArrRows(11,dp)=0) and ArrRows(12,dp) = "1" AND ArrRows(15,dp) = "" Then		'옵션명이 다르고 옵션코드값이 없을 때 ==> 단품추가 의미// preged 0
								aoptNm = Replace(db2Html(ArrRows(2,dp)),":","")
								If aoptNm = "" Then
									aoptNm = "옵션"
								End If
								aoptDc = aoptDc & Replace(Replace(db2Html(ArrRows(3,dp)),":",""),"'","")&","
							End If
						Next

						If aoptDc <> "" Then
	'						rw "단품추가:"&aoptDc
							aoptDc = Left(aoptDc, Len(aoptDc) - 1)
							strParam = ""
							strParam = getLtiMallAddOptParameter(aoptNm, aoptDc, getLtimallGoodno(itemid))
							CALL fnLtiMallAddOpt(itemid, strParam, iErrStr, iAddOptCnt)
							If iAddOptCnt > 0 Then
								If Left(iErrStr, 2) <> "OK" Then
									failCnt = failCnt + 1
									SumErrStr = SumErrStr & iErrStr
								Else
									SumOKStr = SumOKStr & iErrStr
								End If
							End If
						End If
					End If

					'위에서 단품추가 경우가 있기 때문에 재고확인
					Call fnLtiMallStockChk(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					strParam = ""
					strParam = oiMall.FOneItem.getLotteiMallItemEditParameter()
					Call fnLtiMallInfoEdit(itemid, strParam, iErrStr, FALSE)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'전시상품 조회해서 현재 상품상태를 가져오기
					Call fnCheckLtiMallItemStat(itemid, iErrStr, getLtimallGoodno(itemid))
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					''수정 후 단품 재고조회를 함더하자 2018/12/17 추가
					Call fnLtiMallStockChk(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'OK던 ERR이던 editQuecnt에 + 1을 시킴..
					'스케줄링에서 editQuecnt ASC, i.lastupdate DESC로 중복을 막자
					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_ltimall_regitem SET " & VBCRLF
					strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
					strSql = strSql & " ,LTiMallLastupdate = getdate()  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql
				End If
		    Else
				iErrstr = "ERR||"&itemid&"||수정가능한 상품이 아닙니다."
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			End If
			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("lotteimall", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oiMall = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=EDIT
	ElseIf action = "EDIT2" Then								'승인예정상품 수정
		SET oiMall = new CLotteiMall
			oiMall.FRectItemID	= itemid
			oiMall.getLtimallEditOneItem

			strParam = ""
			strParam = oiMall.FOneItem.getLotteiMallItemEditParameter2()
			Call fnLtiMallInfoEdit(itemid, strParam, iErrStr, TRUE)
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
			End If
		SET oiMall = nothing
	ElseIf action = "REG" Then									'상품 등록
		SET oiMall = new CLotteiMall
			oiMall.FRectItemID	= itemid
			oiMall.getLtimallNotRegOneItem
		    If (oiMall.FResultCount < 1) Then
				lastErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
				response.write "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
			Else
				strSql = ""
				strSql = strSql & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_LTiMall_regitem where itemid="&itemid&")"
				strSql = strSql & " BEGIN"& VbCRLF
				strSql = strSql & " 	INSERT INTO db_item.dbo.tbl_LTiMall_regitem "
		        strSql = strSql & " 	(itemid, regdate, reguserid, LTiMallstatCD)"
		        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1')"
				strSql = strSql & " END "
			    dbget.Execute strSql

				If oiMall.FOneItem.checkNotRegWords = "N" Then
					iErrStr = "ERR||"&itemid&"||등록불가 단어 포함(세일, 1+1, 증정, 제공)"
				Else
					'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
					If oiMall.FOneItem.checkTenItemOptionValid Then
						strParam = ""
						strParam = oiMall.FOneItem.getLotteiMallItemRegParameter(FALSE)
						Call LotteiMallItemReg(itemid, strParam, iErrStr, oiMall.FOneItem.FSellCash, oiMall.FOneItem.getLotteiMallSellYn)
					Else
						iErrStr = "ERR||"&itemid&"||옵션검사 실패"
					End If
				End If
				lastErrStr = iErrStr
				response.write iErrStr
				If LEFT(iErrStr, 2) <> "OK" Then
					CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
				End If
			End If
		SET oiMall = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=REG
	End If
End If
'###################################################### LotteiMall API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
