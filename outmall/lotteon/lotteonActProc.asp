<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/lotteon/lotteonItemcls.asp"-->
<!-- #include virtual="/outmall/lotteon/inclotteonFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, olotteon, failCnt, chgSellYn, arrRows, getMustprice, addOptErrItem
Dim iErrStr, strParam, strSql, SumErrStr, SumOKStr, isItemIdChk, grpVal, rSkip, rLimit, i, outmallorderserial
Dim requestJson, responseJson, callComplete, hasnext
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
grpVal			= request("grpVal")
rSkip			= request("rSkip")
rLimit			= request("rLimit")
requestJson		= request("requestJson")
responseJson	= request("responseJson")
failCnt			= 0
outmallorderserial = request("outmallorderserial")
addOptErrItem	= "N"
callComplete = "N"

''카테고리 끌어올시 하단 실행해야 함..
' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_StdCategory]
' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_StdCategory_Disp]
' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_StdCategory_Attr] 
' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_tmpStdCategory] 

' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_DispCategory]
' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_tmpDispCategory]

' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_Attribute] 
' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_Attribute_Values]
Select Case action
	Case "DVPVIEW", "GRPCD", "GRPDTLCD", "ATTRVIEW", "DISPCATE", "STDCATE", "BRANDVIEW", "ORDVIEW"
		isItemIdChk = "N"
	Case Else
		isItemIdChk = "Y"
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
		itemid = CLng(getNumeric(itemid))
	End If
End If
'######################################################## LotteOn API ########################################################
'http://localhost:11117/outmall/lotteon/lotteonActProc.asp?itemid=214560&act=PRICE&requestJson=Y&responseJson=Y
If action = "REG" Then									'상품 등록
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotRegOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Outmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"', '"&CMALLNAME&"' "
			dbget.execute strSql
			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If olotteon.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = olotteon.FOneItem.getLotteonItemRegParameter()
				If requestJson = "Y" Then
					response.write strParam
				End If
				getMustprice = ""
				getMustprice = olotteon.FOneItem.MustPrice()
				CALL fnLotteonItemReg(itemid, strParam, iErrStr, getMustprice, olotteon.FOneItem.getLotteonSellYn, olotteon.FOneItem.FLimityn, olotteon.FOneItem.FLimitNo, olotteon.FOneItem.FLimitSold, html2db(olotteon.FOneItem.FItemName), olotteon.FOneItem.FbasicimageNm, responseJson)
			Else
				iErrStr = "ERR||"&itemid&"||[상품등록] 옵션검사 실패"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteon", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET olotteon = nothing
ElseIf action = "CHKSTAT" Then							'상품 상세 조회
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||조회 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = olotteon.FOneItem.getLotteonItemViewParameter()
			If requestJson = "Y" Then
				response.write strParam
			End If
			CALL fnLotteonItemView(itemid, strParam, iErrStr, responseJson)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteon", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET olotteon = nothing
ElseIf action = "EDITINFO" Then							'상품만 수정
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = olotteon.FOneItem.getLotteonItemEditParameter()			'승인 상품 수정
			If requestJson = "Y" Then
				response.write strParam
			End If
			CALL fnLotteonItemEdit(itemid, olotteon.FOneItem.FItemName, strParam, iErrStr, responseJson)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteon", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET olotteon = nothing
ElseIf action = "EDIT" Then								'상품 수정
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		Else
			strSql = "exec db_etcmall.dbo.usp_Ten_OutMall_optEditParamList_lotteon '"&CMallName&"'," & itemid
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows
			End If
			rsget.close

			If UBound(arrRows,2) = 0 AND arrRows(0,0) = "Z1" Then
				addOptErrItem = "Y"
			End If

			If (oLotteon.FOneItem.FmaySoldOut = "Y") OR (oLotteon.FOneItem.IsMayLimitSoldout = "Y") OR (oLotteon.FOneItem.IsSoldOut) OR (oLotteon.FOneItem.FOptionCnt = 0 AND oLotteon.FOneItem.getRegedOptionCnt > 0)  OR (oLotteon.FOneItem.FLimityn = "Y" AND (oLotteon.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) OR (addOptErrItem = "Y") Then
				chgSellYn = "N"
			Else
				chgSellYn = "Y"
			End If

            If chgSellYn = "N" Then
				strParam = ""
				strParam = olotteon.FOneItem.getLotteonSellynParameter(chgSellYn)
				Call fnLotteOnSellyn(itemid, chgSellYn, strParam, iErrStr, responseJson)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
            Else
				If (olotteon.FOneItem.FLotteonSellYn <> "Y" AND chgSellYn = "Y") Then
					strParam = ""
					strParam = olotteon.FOneItem.getLotteonSellynParameter(chgSellYn)
					Call fnLotteOnSellyn(itemid, chgSellYn, strParam, iErrStr, responseJson)
					response.flush
					response.clear
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strParam = ""
				strParam = olotteon.FOneItem.getLotteonItemViewParameter()				'상품 상세 조회
				CALL fnLotteonItemView(itemid, strParam, iErrStr, responseJson)
			    rw "상품 상세 조회"
				response.flush
				response.clear
                If Left(iErrStr, 2) <> "OK" Then
                    failCnt = failCnt + 1
                    SumErrStr = SumErrStr & iErrStr
                Else
                    SumOKStr = SumOKStr & iErrStr
                End If

				'승인 상품 수정, 상품 가격 변경, 상품 재고 변경까지 주석처리
				' If failCnt = 0 Then
				' 	strParam = ""
				' 	strParam = olotteon.FOneItem.getLotteonItemEditParameter()			'승인 상품 수정
				' 	CALL fnLotteonItemEdit(itemid, olotteon.FOneItem.FItemName, strParam, iErrStr, responseJson)
				' 	rw "승인 상품 수정"
				' 	response.flush
				' 	response.clear
				' 	If Left(iErrStr, 2) <> "OK" Then
				' 		failCnt = failCnt + 1
				' 		SumErrStr = SumErrStr & iErrStr
				' 	Else
				' 		SumOKStr = SumOKStr & iErrStr
				' 	End If
				' End If

				' If failCnt = 0 Then
				' 	getMustprice = ""
				' 	getMustprice = olotteon.FOneItem.MustPrice()
				' 	strParam = ""
				' 	strParam = olotteon.FOneItem.getLotteonPriceParameter()				'상품 가격 변경
				' 	Call fnLotteOnPrice(itemid, strParam, getMustprice, iErrStr, responseJson)
				' 	rw "상품 가격 변경"
				' 	response.flush
				' 	response.clear
				' 	If Left(iErrStr, 2) <> "OK" Then
				' 		failCnt = failCnt + 1
				' 		SumErrStr = SumErrStr & iErrStr
				' 	Else
				' 		SumOKStr = SumOKStr & iErrStr
				' 	End If
				' End If

				' If failCnt = 0 Then
				' 	strParam = ""
				' 	strParam = olotteon.FOneItem.getLotteonQuantityParameter()			'상품 재고 변경
				' 	Call fnLotteOnQuantity(itemid, strParam, iErrStr, responseJson)
				' 	rw "상품 재고 변경"
				' 	response.flush
				' 	response.clear
				' 	If Left(iErrStr, 2) <> "OK" Then
				' 		failCnt = failCnt + 1
				' 		SumErrStr = SumErrStr & iErrStr
				' 	Else
				' 		SumOKStr = SumOKStr & iErrStr
				' 	End If
				' End If

				'################## 위 주석을 아래로 개선 버전..2020-05-07 수정 ######################
				If failCnt = 0 Then
				 	getMustprice = ""
				 	getMustprice = olotteon.FOneItem.MustPrice()

					strParam = ""
					strParam = olotteon.FOneItem.getLotteonItemEditParameter()			'승인 상품 수정
					CALL fnLotteonItemEdit2(itemid, olotteon.FOneItem.FItemName, getMustprice, strParam, iErrStr, responseJson)
					rw "승인 상품 수정"
					response.flush
					response.clear
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
				'################## 개선 버전..2020-05-07 수정 ######################

				If failCnt = 0 Then
					strParam = ""
					strParam = olotteon.FOneItem.getLotteonOptStatusParameter()			'단품 판매상태 변경
					Call fnLotteOnOptStat(itemid, strParam, iErrStr, responseJson)
					rw "단품 판매상태 변경"
					response.flush
					response.clear
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				If failCnt = 0 Then
					strParam = ""
					strParam = olotteon.FOneItem.getLotteonItemViewParameter()			'상품 상세 조회
					CALL fnLotteonItemView(itemid, strParam, iErrStr, responseJson)
					rw "상품 상세 조회"
					response.flush
					response.clear
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
                strSql = strSql & " UPDATE db_etcmall.dbo.tbl_lotteon_regitem SET " & VBCRLF
                strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
                strSql = strSql & " ,lotteonlastupdate = getdate()  " & VBCRLF
                strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
                dbget.Execute strSql
			End If
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("lotteon", itemid, SumErrStr)
			Call SugiQueLogInsert("lotteon", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_lotteon_regitem SET " & VBCRLF
			strSql = strSql & " accFailcnt = 0  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("lotteon", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET olotteon = nothing
ElseIf action = "QTY" Then								'상품 재고 변경
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = olotteon.FOneItem.getLotteonQuantityParameter()
			If requestJson = "Y" Then
				response.write strParam
			End If
			Call fnLotteOnQuantity(itemid, strParam, iErrStr, responseJson)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteon", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET olotteon = nothing
ElseIf action = "PRICE" Then							'상품 가격 변경
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		Else
			If LEFT(olotteon.FOneItem.FLastStatCheckDate, 10) = "1900-01-01" Then
				strParam = ""
				strParam = olotteon.FOneItem.getLotteonItemViewParameter()			'상품 상세 조회
				CALL fnLotteonItemView(itemid, strParam, iErrStr, responseJson)
				rw "상품 상세 조회"
				response.flush
				response.clear
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If

			If failCnt = 0 Then
				getMustprice = ""
				getMustprice = olotteon.FOneItem.MustPrice()
				strParam = ""
				strParam = olotteon.FOneItem.getLotteonPriceParameter()				'상품 가격 변경
				Call fnLotteOnPrice(itemid, strParam, getMustprice, iErrStr, responseJson)
				rw "상품 가격 변경"
				response.flush
				response.clear
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
			CALL Fn_AcctFailTouch("lotteon", itemid, SumErrStr)
			Call SugiQueLogInsert("lotteon", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("lotteon", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET olotteon = nothing
ElseIf action = "EditSellYn" Then						'상품 상태 변경
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = olotteon.FOneItem.getLotteonSellynParameter(chgSellYn)
			If requestJson = "Y" Then
				response.write strParam
			End If
			Call fnLotteOnSellyn(itemid, chgSellYn, strParam, iErrStr, responseJson)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteon", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET olotteon = nothing
ElseIf action = "OPTSTAT" Then							'단품 판매상태 변경
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = olotteon.FOneItem.getLotteonOptStatusParameter()
			If requestJson = "Y" Then
				response.write strParam
			End If
			Call fnLotteOnOptStat(itemid, strParam, iErrStr, responseJson)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteon", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET olotteon = nothing
ElseIf action = "DVPVIEW" Then							'판매자 출고지/반품지 리스트 조회
	Call fnlotteonDVPView()
ElseIf action = "ATTRVIEW" Then							'속성 기본 조회
	Do Until callComplete = "Y"
		Call fnlotteonAttrView(rSkip, hasnext)
		If hasnext = "N" Then
			callComplete = "Y"
			rw "완료"
		Else
			rw "API 호출 중 입니다. "
			rw "-------------------------"
		End If
		response.flush
	Loop
ElseIf action = "DISPCATE" Then							'전시카테고리 조회
	Do Until callComplete = "Y"
		Call fnlotteonDispCateView(rSkip, hasnext)
		If hasnext = "N" Then
			callComplete = "Y"
			rw "완료"
		Else
			rw "API 호출 중 입니다. "
			rw "-------------------------"
		End If
		response.flush
	Loop
ElseIf action = "STDCATE" Then							'표준카테고리 조회
	Do Until callComplete = "Y"
		Call fnlotteonStdCateView(rSkip, hasnext)
		If hasnext = "N" Then
			callComplete = "Y"
			rw "완료"
		Else
			rw "API 호출 중 입니다. "
			rw "-------------------------"
		End If
		response.flush
	Loop
ElseIf action = "BRANDVIEW" Then						'브랜드 조회
	Call fnlotteonBrandView(rSkip, rLimit)
ElseIf action = "GRPCD" Then							'공통코드 조회
	Call fnlotteonGetGroupCode()
ElseIf action = "GRPDTLCD" Then							'공통코드 상세 조회
	Call fnlotteonGetGroupCodeDetail(grpVal)
ElseIf action = "ORDVIEW" Then							'배송상태 조회 / JSON 출력만 했음
	Call fnlotteonViewOrder(outmallorderserial)
End If

If iErrStr <> "" Then
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str = '"&iErrStr&"<br>' + str " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
					"</script>"
End If
'###################################################### LotteOn API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
