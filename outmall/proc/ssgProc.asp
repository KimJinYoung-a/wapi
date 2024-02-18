<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/ssg/ssgItemcls.asp"-->
<!-- #include virtual="/outmall/ssg/incssgFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, failCnt, oSsg, getMustprice, tSsgGoodNo, vOptCnt, chgImageNm, chgSellYn
Dim iErrStr, strParam, mustPrice, strSql, SumErrStr, SumOKStr, endItemErrMsgReplace
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
'######################################################## ssg API ########################################################
If mallid = "ssg" Then
	If action = "REG" Then					'상품등록
		SET oSsg = new CSsg
			oSsg.FRectItemID	= itemid
			oSsg.getSsgNotRegOneItem
			If (oSsg.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
			ElseIf (oSsg.FOneItem.FMapcnt = 0) Then
				iErrStr = "ERR||"&itemid&"||카테고리 매칭이 필요합니다."
			Else
				strSql = ""
				strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_ssg_regitem where itemid="&itemid&")"
				strSql = strSql & " BEGIN"& VbCRLF
				strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_ssg_regitem "
				strSql = strSql & " 	(itemid, regdate, reguserid, ssgstatCD, regitemname, ssgSellYn)"
				strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oSsg.FOneItem.FItemName)&"', 'N')"
				strSql = strSql & " END "
				dbget.Execute strSql
				'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
				If oSsg.FOneItem.checkTenItemOptionValid Then
					getMustprice = ""
					getMustprice = oSsg.FOneItem.MustPrice()

					strParam = ""
					strParam = oSsg.FOneItem.getssgItemRegParameter(getMustprice)
					Call fnSsgItemReg(itemid, strParam, iErrStr, getMustprice, oSsg.FOneItem.FbasicimageNm, oSsg.FOneItem.getSSGMargin)
				Else
					iErrStr = "ERR||"&itemid&"||[상품등록] 옵션검사 실패"
				End If
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("ssg", itemid, iErrStr)
			End If
		SET oSsg = nothing
	ElseIf action = "SOLDOUT" Then			'상태변경
		SET oSsg = new Cssg
			oSsg.FRectItemID	= itemid
			oSsg.FRectMustSellyn= "Y"
			oSsg.getSsgEditOneItem

		    If (oSsg.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||[상품수정] 수정 가능한 상품이 아닙니다."
			Else
				strParam = ""
				strParam = oSsg.FOneItem.getssgItemEditSellynParameter("N")
				getMustprice = ""
				getMustprice = oSsg.FOneItem.MustPrice()
				If oSsg.FOneItem.isImageChanged Then
					chgImageNm = oSsg.FOneItem.getBasicImage
				Else
					chgImageNm = "N"
				End If
				Call fnSsgItemEditSellyn(itemid, oSsg.FOneItem.FSsgGoodNo, iErrStr, strParam, getMustprice, html2db(oSsg.FOneItem.FItemName), "N", chgImageNm)
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("ssg", itemid, iErrStr)
			End If
		SET oSsg = nothing
		'http://wapi.10x10.co.kr/outmall/proc/ssgProc.asp?itemid=325046&mallid=ssg&action=SOLDOUT
	ElseIf action = "CHKSTAT" Then			'승인확인
		tSsgGoodNo = getSsgGoodNo(itemid)
		Call fnSsgStatChk(itemid, tSsgGoodNo, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ssg", itemid, iErrStr)
		End If
		'http://wapi.10x10.co.kr/outmall/proc/ssgProc.asp?itemid=325046&mallid=ssg&action=CHKSTAT
	ElseIf (action = "EDIT") OR (action = "PRICE") Then		'가격 및 상품수정
		SET oSsg = new Cssg
			oSsg.FRectItemID	= itemid
			oSsg.getSsgEditOneItem
		    If (oSsg.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||[상품수정] 수정 가능한 상품이 아닙니다."
			Else
				If (oSsg.FOneItem.getiszeroWonSoldOut(itemid) = "Y") OR (oSsg.FOneItem.FmaySoldOut = "Y") OR (oSsg.FOneItem.IsMayLimitSoldout = "Y") OR (oSsg.FOneItem.IsSoldOut) OR (oSsg.FOneItem.FOptionCnt = 0 AND oSsg.FOneItem.getRegedOptionCnt > 0) Then
					chgSellYn = "N"
				Else
					chgSellYn = "Y"
				End If

				If chgSellYn = "N" Then
					strParam = ""
					strParam = oSsg.FOneItem.getssgItemEditSellynParameter(chgSellYn)
					getMustprice = ""
					getMustprice = oSsg.FOneItem.MustPrice()
					If oSsg.FOneItem.isImageChanged Then
						chgImageNm = oSsg.FOneItem.getBasicImage
					Else
						chgImageNm = "N"
					End If
					Call fnSsgItemEditSellyn(itemid, oSsg.FOneItem.FSsgGoodNo, iErrStr, strParam, getMustprice, html2db(oSsg.FOneItem.FItemName), chgSellYn, chgImageNm)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					strParam = ""
					strParam = oSsg.FOneItem.getssgItemEditParameter(chgSellYn)
					getMustprice = ""
					getMustprice = oSsg.FOneItem.MustPrice()
					If oSsg.FOneItem.isImageChanged Then
						chgImageNm = oSsg.FOneItem.getBasicImage
					Else
						chgImageNm = "N"
					End If
					Call fnSsgItemEdit(itemid, oSsg.FOneItem.FSsgGoodNo, iErrStr, strParam, getMustprice, html2db(oSsg.FOneItem.FItemName), chgSellYn, chgImageNm)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					If failCnt > 0 Then
						endItemErrMsgReplace = replace(SumErrStr, "OK||"&itemid&"||", "")
						endItemErrMsgReplace = replace(SumErrStr, "ERR||"&itemid&"||", "")

						If (Instr(endItemErrMsgReplace, "중복된옵션이존재합니다") > 0) OR (Instr(endItemErrMsgReplace, "중복 된 옵션이") > 0) Then
							strParam = ""
							strParam = oSsg.FOneItem.getssgItemEditSellynParameter("X")
							Call fnSsgItemEditSellyn(itemid, oSsg.FOneItem.FSsgGoodNo, iErrStr, strParam, getMustprice, html2db(oSsg.FOneItem.FItemName), "X", chgImageNm)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						End If
					Else
						Call fnViewItemInfo(itemid, oSsg.FOneItem.FSsgGoodNo, iErrStr)
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
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ssg_regitem SET " & VBCRLF
					strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
					strSql = strSql & " ,ssglastupdate = getdate()  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql
				End If
			End If
			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("ssg", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oSsg = nothing
		'http://wapi.10x10.co.kr/outmall/proc/ssgProc.asp?itemid=325046&mallid=ssg&action=EDIT
	End If
End If
'###################################################### ssg API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->