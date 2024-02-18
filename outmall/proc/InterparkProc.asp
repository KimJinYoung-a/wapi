<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/interpark/interparkItemcls.asp"-->
<!-- #include virtual="/outmall/interpark/incInterparkFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, oInterpark, failCnt, arrRows, skipItem, isAllRegYn, getMustprice
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname
Dim isoptionyn, isText, i, interparkPrdno, dataUrl, chgImageNm, getLimityn
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
'######################################################## Interpark API ########################################################
If mallid = "interpark" Then
	If action = "REG" Then													'상품등록
		SET oInterpark = new CInterpark
			oInterpark.FRectItemID	= itemid
			oInterpark.getInterparkNotRegScheduleOneItem
		If (oInterpark.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf getInterparkPrdno(itemid) <> "" Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM [db_item].[dbo].tbl_interpark_reg_item where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO [db_item].[dbo].tbl_interpark_reg_item  "
	        strSql = strSql & " 	(itemid,reguserid) "
	        strSql = strSql & " 	VALUES ("&itemid&", '"&session("SSBctID")&"')"
			strSql = strSql & " END "
			dbget.Execute strSql
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
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("interpark", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/InterparkProc.asp?itemid=279397&mallid=interpark&action=REG
	ElseIf (action = "EDIT") OR (action = "PRICE") Then						'상품수정
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

					'2.실패횟수가 0일때만 상품수정
					If failCnt = "0" Then
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
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oInterpark = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/InterparkProc.asp?itemid=997155&mallid=interpark&action=EDIT
	ElseIf action = "SOLDOUT" Then						'판매상태변경
		dataUrl = getInterparkSellynParameter("N", getInterparkPrdno(itemid))
		strParam = ""
		strParam = "_method=UpdateProductAPIStatTpQty&citeKey=h5Szn1XPevegFsnSYKfGgEOI6E1mQnQqeEjffn5p5Zo=&secretKey=6CxkBwV1Bk^CiWEqdQ5cV^XiFcrjBaTn"
		Call fnInterparkSellyn(itemid, "N", strParam, dataUrl, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("interpark", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/InterparkProc.asp?itemid=279397&mallid=interpark&action=SOLDOUT
	ElseIf action = "DELETE" Then						'삭제
		dataUrl = getInterparkSellynParameter("X", getInterparkPrdno(itemid))
		strParam = ""
		strParam = "_method=UpdateProductAPIStatTpQty&citeKey=h5Szn1XPevegFsnSYKfGgEOI6E1mQnQqeEjffn5p5Zo=&secretKey=6CxkBwV1Bk^CiWEqdQ5cV^XiFcrjBaTn"
		Call fnInterparkSellyn(itemid, "X", strParam, dataUrl, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("interpark", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/InterparkProc.asp?itemid=279397&mallid=interpark&action=SOLDOUT
	ElseIf action = "CHKSTAT" Then						'판매상태조회
		interparkPrdno = ""
		interparkPrdno = getInterparkPrdno(itemid)
		strParam = ""
		strParam = "_method=GetPrdSaleQtyForAPI&citeKey=HmMTYbcJDv7aeUsOEUJ5gDCGH7eaEqrg&secretKey=dzpAObpfn37MkqwHIXXm7aFJchN0b9Yw&prdNo="&interparkPrdno
		Call fnInterparkstatChk(strParam, itemid, interparkPrdno, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("interpark", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/InterparkProc.asp?itemid=279397&mallid=interpark&action=CHKSTAT
	End If
End If
'###################################################### Interpark API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->