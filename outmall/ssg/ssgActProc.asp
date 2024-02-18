<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/ssg/ssgItemcls.asp"-->
<!-- #include virtual="/outmall/ssg/incssgFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim testkey, siNo
testkey = request("testkey")
siNo = request("siNo")


'If (session("ssBctID")="kjy8517") and (testkey <> "") Then
If (testkey <> "") Then
	If testkey = "1" Then
		rw ssgAPIURL
		call getSsgMmgCateList()		''관리카테고리 가져오기
		dbget.close() : response.end
	ElseIf testkey = "2" Then
		rw ssgAPIURL
		'call getSsgDispCateListALL()	'' 관리카테고리 세분류 코드로 전시카테고리 가져오기
		Call fnSsgDispCategoryGet(siNo)
		dbget.close() : response.end
	ElseIf testkey = "3" Then
		rw ssgAPIURL
		Call fnSsgGosiSafeInfo()		'' 표준분류별 인증정보 조회
		dbget.close() : response.end
	End If
End If
' http://wapi.10x10.co.kr/outmall/ssg/ssgActproc.asp?testkey=1		[관리카테고리]
' http://wapi.10x10.co.kr/outmall/ssg/ssgActproc.asp?testkey=2&siNo=6001	[전시카테고리]	6001 : 이마트몰
' http://wapi.10x10.co.kr/outmall/ssg/ssgActproc.asp?testkey=2&siNo=6004	[전시카테고리]	6004 : 신세계몰
' http://wapi.10x10.co.kr/outmall/ssg/ssgActproc.asp?testkey=2&siNo=6005	[전시카테고리]	6005 : SSG
' 2020-11-18 하단 추가 / 표준분류별 인증정보 조회
' http://wapi.10x10.co.kr/outmall/ssg/ssgActproc.asp?testkey=3


Dim itemid, action, oSsg, oSsgOpt, failCnt, chgSellYn, arrRows, skipItem, tssgGoodno, tOptionCnt, tLimityn, isAllRegYn, getMustprice, tIsChildrenCate
Dim iErrStr, strParam, mustPrice, displayDate, ret1, strSql, SumErrStr, SumOKStr, iitemname, isItemIdChk, isFiftyUpDown, isiframe, chgImageNm
Dim isChild, isLife, isElec, endItemErrMsgReplace
Dim isoptionyn, isText, i
Dim failCnt2, ckParam, gosiClsId
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
ckParam			= request("ckParam")
gosiClsId		= request("gosiClsId")
failCnt			= 0
failCnt2		= 0

Select Case action
	Case "GOSI", "AREA", "DISPCATE", "GOSISAFE"		isItemIdChk = "N"
	Case Else										isItemIdChk = "Y"
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
'######################################################## ssg API ########################################################
If action = "REG" Then								'등록
	SET oSsg = new CSsg
		oSsg.FRectItemID	= itemid
		oSsg.getSsgNotRegOneItem
	    If (oSsg.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf (oSsg.FOneItem.FMapcnt = 0) Then
			iErrStr = "ERR||"&itemid&"||카테고리 매칭이 필요합니다."
		Else
'rw oSsg.FOneItem.getSSGMargin
'response.end

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
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ssg", itemid, iErrStr)
		End If

		if (NOT IsAutoScript) then '' SSG 상품등록 자동으로 할경우 로그를 넣을필요가 없음..
			Call SugiQueLogInsert("ssg", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		end if
	SET oSsg = nothing
ElseIf action = "CHKSTAT" Then						'승인목록 조회
	tSsgGoodNo = getSsgGoodNo(itemid)
	Call fnSsgStatChk(itemid, tSsgGoodNo, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("ssg", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("ssg", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EditSellYn" Then
	SET oSsg = new Cssg
		oSsg.FRectItemID	= itemid
		oSsg.FRectMustSellyn= "Y"
		oSsg.getSsgEditOneItem

	    If (oSsg.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||[상품수정] 수정 가능한 상품이 아닙니다."
		Else
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
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ssg", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ssg", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSsg = nothing
ElseIf action = "EDITINFO" Then
	SET oSsg = new Cssg
		oSsg.FRectItemID	= itemid
		oSsg.getSsgEditOneItem

		strParam = ""
		strParam = oSsg.FOneItem.getssgItemEditParameter("Y")
		getMustprice = ""
		getMustprice = oSsg.FOneItem.MustPrice()
		If oSsg.FOneItem.isImageChanged Then
			chgImageNm = oSsg.FOneItem.getBasicImage
		Else
			chgImageNm = "N"
		End If

		Call fnSsgItemEdit(itemid, oSsg.FOneItem.FSsgGoodNo, iErrStr, strParam, getMustprice, html2db(oSsg.FOneItem.FItemName), chgSellYn, chgImageNm)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ssg", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ssg", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSsg = nothing
ElseIf action = "EDIT" Then
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
                Call fnViewItemInfo(itemid, oSsg.FOneItem.FSsgGoodNo, iErrStr)
                If Left(iErrStr, 2) <> "OK" Then
                    failCnt = failCnt + 1
                    SumErrStr = SumErrStr & iErrStr
                Else
                    SumOKStr = SumOKStr & iErrStr
                End If

                strParam = ""
                strParam = oSsg.FOneItem.getssgItemEditParameter(chgSellYn)
				If ckParam = "Y" Then
					response.write strParam
				End If

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
			Call SugiQueLogInsert("ssg", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("ssg", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oSsg = nothing
ElseIf action = "VIEW" Then
	tSsgGoodNo = getSsgGoodNo(itemid)
	Call fnViewItemInfo(itemid, tSsgGoodNo, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("ssg", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("ssg", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "GOSI" Then
	strParam = ""
	Call fnSsgGosiInfo(gosiClsId)
ElseIf action = "AREA" Then
	strParam = ""
	Call fnSsgAreaInfo()
' ElseIf action = "DISPCATE" Then
' 	Call fnSsgDispCategoryGet()
ElseIf action = "GOSISAFE" Then
	strParam = ""
	Call fnSsgGosiSafeInfo()
End If

If iErrStr <> "" Then
	if (IsAutoScript) then
		response.write iErrStr
	else
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str = '"&iErrStr&"<br>' + str " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
					"</script>"
	end if
End If
'###################################################### ssg API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
