<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/interpark/interparkItemcls.asp"-->
<!-- #include virtual="/outmall/interpark/incInterparkFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, oInterpark, oAuctionOpt, failCnt, chgSellYn, arrRows, skipItem, tAuctionGoodno, isAllRegYn, getMustprice
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, ccd, isItemIdChk
Dim isoptionyn, isText, i, interparkPrdno, dataUrl, chgImageNm, chkRegeditem, getLimityn, sDate, eDate
Dim param1, param2
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
ccd				= request("ccd")
sDate			= request("sDate")
eDate			= request("eDate")
param1			= request("param1")
param2			= request("param2")
failCnt			= 0

Select Case action
	Case "cateRcv", "getDelivery", "CATE"	isItemIdChk = "N"
	Case Else								isItemIdChk = "Y"
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
'######################################################## interpark API ########################################################
If action = "REG" Then								'상품등록
	SET oInterpark = new CInterpark
		oInterpark.FRectItemID	= itemid
		oInterpark.getInterparkNotRegOneItem

		If (oInterpark.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf getInterparkPrdno(itemid) <> "" Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oInterpark.FOneItem.checkTenItemOptionValid Then
				dataUrl = ""
				dataUrl = oInterpark.FOneItem.getInterparkItemRegParameter()
				chgImageNm = oInterpark.FOneItem.getBasicImage

				strParam = ""
				strParam = "_method=InsertProductAPIData&citeKey=Cxyso3Izaa7VNiHAauqT3ocgYfDqdiqpO6Z02j63U4w=&secretKey=u6r9q5YmW9nOnAuo6w6kDJF1/43iVb42"
				Call fnInterparkItemReg(itemid, strParam, dataUrl, iErrStr, oInterpark.FOneItem.MustPrice, chgImageNm)
			Else
				iErrStr = "ERR||"&itemid&"||옵션검사 실패"
			End If
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("interpark", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("interpark", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oInterpark = nothing
ElseIf action = "EDIT" Then							'상품수정
	SET oInterpark = new CInterpark
		oInterpark.FRectItemID	= itemid
		oInterpark.getInterparkEditOneItem
		If oInterpark.FResultCount = 0 Then
			failCnt = failCnt + 1
			iErrStr = "ERR||"&itemid&"||수정가능한 상품이 아닙니다."
		Else
			getLimityn = oInterpark.FOneItem.Flimityn
			If (oInterpark.FOneItem.FMaySoldOut = "Y") OR (oInterpark.FOneItem.IsSoldOutLimit5Sell) OR (oInterpark.FOneItem.getiszeroWonSoldOut(itemid, getLimityn) = "Y") Then
				dataUrl = ""
				dataUrl = getInterparkSellynParameter("N", oInterpark.FOneItem.FInterparkPrdNo)
				strParam = ""
				strParam = "_method=UpdateProductAPIStatTpQty&citeKey=h5Szn1XPevegFsnSYKfGgEOI6E1mQnQqeEjffn5p5Zo=&secretKey=6CxkBwV1Bk^CiWEqdQ5cV^XiFcrjBaTn"
				Call fnInterparkSellyn(itemid, "N", strParam, dataUrl, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				'1.판매상태 가져오기(regedoption때문)
				interparkPrdno = ""
				interparkPrdno = oInterpark.FOneItem.FInterparkPrdNo
				strParam = ""
				strParam = "_method=GetPrdSaleQtyForAPI&citeKey=HmMTYbcJDv7aeUsOEUJ5gDCGH7eaEqrg&secretKey=dzpAObpfn37MkqwHIXXm7aFJchN0b9Yw&prdNo="&interparkPrdno
				Call fnInterparkstatChk(strParam, itemid, interparkPrdno, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'rw oInterpark.FOneItem.FLimityn
				'rw oInterpark.FOneItem.FOptionCnt
				If failCnt = "0" Then
					If (oInterpark.FOneItem.FmayiParkSellYn = "N") AND ((oInterpark.FOneItem.FMaySoldOut <> "Y") AND (oInterpark.FOneItem.IsSoldOutLimit5Sell = False )) Then
						dataUrl = ""
						dataUrl = getInterparkSellynParameter("Y", oInterpark.FOneItem.FInterparkPrdNo)
						strParam = ""
						strParam = "_method=UpdateProductAPIStatTpQty&citeKey=h5Szn1XPevegFsnSYKfGgEOI6E1mQnQqeEjffn5p5Zo=&secretKey=6CxkBwV1Bk^CiWEqdQ5cV^XiFcrjBaTn"
						Call fnInterparkSellyn(itemid, "Y", strParam, dataUrl, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

					'2.실패횟수가 0일때 상품수정, 1이상이라면 품절처리
					dataUrl = ""
					dataUrl = oInterpark.FOneItem.getInterparkItemEditParameter()
					If oInterpark.FOneItem.FMayLimitSoldout = "Y" Then
						failCnt = "0"
						SumErrStr = ""
						SumOKStr = ""
						dataUrl = ""
						dataUrl = getInterparkSellynParameter("N", oInterpark.FOneItem.FInterparkPrdNo)
						strParam = ""
						strParam = "_method=UpdateProductAPIStatTpQty&citeKey=h5Szn1XPevegFsnSYKfGgEOI6E1mQnQqeEjffn5p5Zo=&secretKey=6CxkBwV1Bk^CiWEqdQ5cV^XiFcrjBaTn"
						Call fnInterparkSellyn(itemid, "N", strParam, dataUrl, iErrStr)
					Else
						If oInterpark.FOneItem.isImageChanged Then
							chgImageNm = oInterpark.FOneItem.getBasicImage
						Else
							chgImageNm = "N"
						End If

						strParam = ""
						strParam = "_method=UpdateProductAPIData&citeKey=9CIgE/zSo2ZlDnPaviyqoKmRUPF6ZRea&secretKey=MaMpPg2WSWUE1NiGGmgTm7Ax63xqcqgJ"
						Call fnInterparkInfoEdit(itemid, strParam, dataUrl, iErrStr, chgImageNm, oInterpark.FOneItem.MustPrice,oInterpark.FOneItem.GetInterParkSaleStatTp)
					End If

					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr

						''상품수정후 재고조회를 함 하자 2018/12/13----------
						strParam = ""
						strParam = "_method=GetPrdSaleQtyForAPI&citeKey=HmMTYbcJDv7aeUsOEUJ5gDCGH7eaEqrg&secretKey=dzpAObpfn37MkqwHIXXm7aFJchN0b9Yw&prdNo="&interparkPrdno
						Call fnInterparkstatChk(strParam, itemid, interparkPrdno, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
						''------------------------------------------------
					End If
				Else
					If right(SumErrStr,4) = "002]" Then
						failCnt = "0"
						SumErrStr = ""
						SumOKStr = ""
						dataUrl = ""

						dataUrl = getInterparkSellynParameter("N", oInterpark.FOneItem.FInterparkPrdNo)
						strParam = ""
						strParam = "_method=UpdateProductAPIStatTpQty&citeKey=h5Szn1XPevegFsnSYKfGgEOI6E1mQnQqeEjffn5p5Zo=&secretKey=6CxkBwV1Bk^CiWEqdQ5cV^XiFcrjBaTn"
						Call fnInterparkSellyn(itemid, "N", strParam, dataUrl, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr & "_옵션명 오류"
						Else
							SumOKStr = SumOKStr & iErrStr & "_옵션명 오류"
						End If
					End If
				End If
			End If
		End If

		'OK던 ERR이던 editQuecnt에 + 1을 시킴..
		'스케줄링에서 editQuecnt ASC, i.lastupdate DESC로 중복을 막자
		strSql = ""
		strSql = strSql & " UPDATE [db_item].[dbo].tbl_interpark_reg_item SET " & VBCRLF
		strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
		strSql = strSql & " ,interparklastupdate = getdate()  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("interpark", itemid, SumErrStr)
			Call SugiQueLogInsert("interpark", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("interpark", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oInterpark = nothing
ElseIf action = "EditSellYn" Then					'판매상태변경
	dataUrl = getInterparkSellynParameter(chgSellYn, getInterparkPrdno(itemid))
	strParam = ""
	strParam = "_method=UpdateProductAPIStatTpQty&citeKey=h5Szn1XPevegFsnSYKfGgEOI6E1mQnQqeEjffn5p5Zo=&secretKey=6CxkBwV1Bk^CiWEqdQ5cV^XiFcrjBaTn"
	Call fnInterparkSellyn(itemid, chgSellYn, strParam, dataUrl, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("interpark", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("interpark", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKSTAT" Then						'판매상태조회
	interparkPrdno = ""
	interparkPrdno = getInterparkPrdno(itemid)
	strParam = ""
	strParam = "_method=GetPrdSaleQtyForAPI&citeKey=HmMTYbcJDv7aeUsOEUJ5gDCGH7eaEqrg&secretKey=dzpAObpfn37MkqwHIXXm7aFJchN0b9Yw&prdNo="&interparkPrdno
	Call fnInterparkstatChk(strParam, itemid, interparkPrdno, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("interpark", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("interpark", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "cateRcv" Then						'카테고리 떙겨오기
	If sDate <> "" Then
		sDate = Left(Replace(CStr(DateAdd("d",-365,date())),"-",""),10)
	End If

	If eDate <> "" Then
		eDate = "20181029"
	End If

	strParam = ""
	strParam = "_method=GetBasicCategoryForAPI&citeKey=KIQpKWSzGVladyAxxM4vAz3nCetGjAmmAXKkQotL8KQ=&secretKey=2FfOmboyJ6EG17kcxUnIcZF1/43iVb42"
	strParam = strParam & "&strDt=" & sDate ''[기간시작] YYYYMMDD
	strParam = strParam & "&endDt=" & eDate
'	strParam = strParam & "&endDt=20180321"
'	strParam = strParam & "&strDt=20110601"
'	strParam = strParam & "&endDt=20120601"

	Call fnInterparkCategory(strParam)
ElseIf action = "getDelivery" Then					'배송비정책 조회
	strParam = ""
	strParam = "_method=getDelvCostPlcAPIData&citeKey=o0y^YpvNFa3iHOjFBEwwehL9BRjiI0e9&secretKey=usJwLKiJPSpMWsfqHdt4WiZgdpEZ5DYr"
	Call fnInterparkDeliveryView(strParam)
ElseIf action = "CATE" Then							'카테고리 조회
	If param1 = "" Then
		response.write "<script>alert('시작일이 없습니다.')</script>"
		response.end
	End If

	If param2 = "" Then
		response.write "<script>alert('종료일이 없습니다.')</script>"
		response.end
	End If

	param1 = replace(param1, "-", "")
	param2 = replace(param2, "-", "")

	strParam = ""
	strParam = "_method=GetBasicCategoryForAPI&citeKey=KIQpKWSzGVladyAxxM4vAz3nCetGjAmmAXKkQotL8KQ=&secretKey=2FfOmboyJ6EG17kcxUnIcZF1/43iVb42"
	' strParam = strParam & "&strDt=" & param1
	' strParam = strParam & "&endDt=" & param2
	strParam = strParam & "&dispYn=Y"
	Call fnInterparkCategoryView(strParam)
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
'###################################################### interpark API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
