<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/nvstoregift/nvstoregiftItemcls.asp"-->
<!-- #include virtual="/outmall/nvstoregift/incNvstoregiftFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, oNvstoregift, failCnt, chgSellYn, arrRows, skipItem, sellgubun, getMustprice, chkXML
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, getfarmGoodno
Dim oService, oOperation, mayOptSoldOut, chgImageNm, endItemErrMsgReplace
Dim jenkinsBatchYn, idx, lastErrStr
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
chkXML			= request("chkXML")
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
'######################################################## 스토어팜 API ########################################################
If mallid = "nvstoregift" Then
	If action = "CHKOPT" Then									'옵션 조회
		oService		= "ProductService"
		oOperation		= "GetOption"

		strParam = ""
		strParam = getNvstoregiftOptionSearchParameter(getNvstoregiftGoodNo(itemid), oService, oOperation)
		Call fnNvstoregiftOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
		End If
	ElseIf action = "SOLDOUT" Then												'상태변경
		oService		= "ProductService"
		oOperation		= "ChangeProductSaleStatus"

		strParam = ""
		strParam = getNvstoregiftSellynParameter(getNvstoregiftGoodNo(itemid), "N", oService, oOperation)
		Call fnNvstoregiftSellyn(itemid, "N", strParam, iErrStr, oService, oOperation)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
		End If
		'http://wapi.10x10.co.kr/outmall/proc/NvstoregiftProc.asp?itemid=3144076&mallid=nvstoregift&action=SOLDOUT
	ElseIf action = "EDIT" OR action = "ITEMNAME" OR action = "PRICE" Then		'상품수정
		SET oNvstoregift = new CNvstoregift
			oNvstoregift.FRectItemID	= itemid
			oNvstoregift.getNvstoregiftEditOneItem

			If (oNvstoregift.FResultCount < 1) OR (oNvstoregift.FOneItem.FNvstoregiftGoodNo = "") Then
				iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
				failCnt = failCnt + 1
			Else
				If oNvstoregift.FOneItem.FOptioncnt > 0 Then
					mayOptSoldOut = oNvstoregift.FOneItem.IsMayLimitSoldout
				End If

				If (oNvstoregift.FOneItem.FMaySoldOut = "Y") OR (oNvstoregift.FOneItem.IsSoldOutLimit5Sell) OR (mayOptSoldOut = "Y") Then
					oService		= "ProductService"
					oOperation		= "ChangeProductSaleStatus"

					strParam = ""
					strParam = getNvstoregiftSellynParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, "N", oService, oOperation)
					Call fnNvstoregiftSellyn(itemid, "N", strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					If (oNvstoregift.FOneItem.FNvstoregiftSellYn = "N" AND oNvstoregift.FOneItem.IsSoldOutLimit5Sell = False) Then
						oService		= "ProductService"
						oOperation		= "ChangeProductSaleStatus"

						strParam = ""
						strParam = getNvstoregiftSellynParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, "Y", oService, oOperation)
						Call fnNvstoregiftSellyn(itemid, "Y", strParam, iErrStr, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
	'################################################ 0.상품 수정부터(ReturnCostReason 신규 필드 때문..) ####################
					If oNvstoregift.FOneItem.isImageChanged = True Then
						chgImageNm = oNvstoregift.FOneItem.getBasicImage
					Else
						chgImageNm = "N"
					End If

					oService		= "ProductService"
					oOperation		= "ManageProduct"

					strParam = ""
					strParam = oNvstoregift.FOneitem.getNvstoregiftItemRegXML(oService, oOperation, "Y")
					getMustprice = ""
					getMustprice = oNvstoregift.FOneItem.MustPrice()
					Call fnNvstoregiftItemEDIT(itemid, strParam, iErrStr, getMustprice, oNvstoregift.FOneItem.getNvstoregiftSellYn, oNvstoregift.FOneItem.FLimityn, oNvstoregift.FOneItem.FLimitNo, oNvstoregift.FOneItem.FLimitSold, oNvstoregift.FOneItem.FItemName, chgImageNm, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
	'################################################ 1.옵션 가져오기(regedoption때문) #######################################
					oService		= "ProductService"
					oOperation		= "GetOption"

					strParam = ""
					strParam = getNvstoregiftOptionSearchParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, oService, oOperation)
					Call fnNvstoregiftOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
	'##########################################################################################################################
	'################################################ 2.이미지 변경시 이미지 재업로드 #########################################
					If chgImageNm <> "N" Then
						oService		= "ImageService"
						oOperation		= "UploadImage"

						strParam = ""
						strParam = oNvstoregift.FOneitem.getNvstoregiftImageRegXML(oService, oOperation)
						chgImageNm = oNvstoregift.FOneItem.getBasicImage
						Call fnNvstoregiftImageReg(itemid, strParam, iErrStr, chgImageNm, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
	'##########################################################################################################################
	'############################################## 3.실패횟수가 0일때 상품수정 ###############################################
					If failCnt = "0" Then
						oService		= "ProductService"
						oOperation		= "ManageProduct"

						strParam = ""
						strParam = oNvstoregift.FOneitem.getNvstoregiftItemRegXML(oService, oOperation, "Y")
						getMustprice = ""
						getMustprice = oNvstoregift.FOneItem.MustPrice()
						Call fnNvstoregiftItemEDIT(itemid, strParam, iErrStr, getMustprice, oNvstoregift.FOneItem.getNvstoregiftSellYn, oNvstoregift.FOneItem.FLimityn, oNvstoregift.FOneItem.FLimitNo, oNvstoregift.FOneItem.FLimitSold, (oNvstoregift.FOneItem.FItemName), chgImageNm, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
	'##########################################################################################################################
	'############################################## 4.옵션수정 ################################################################
						oService		= "ProductService"
						oOperation		= "ManageOption"

						strParam = ""
						strParam = getNvstoregiftOptionRegXML(itemid, oNvstoregift.FOneItem.FNvstoregiftGoodno, oService, oOperation)
						If strParam <> "X" Then
							Call fnNvstoregiftOptionReg(itemid, strParam, iErrStr, oService, oOperation)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						End If
	'##########################################################################################################################
	'################################################ 5.옵션 가져오기 #######################################
						oService		= "ProductService"
						oOperation		= "GetOption"

						strParam = ""
						strParam = getNvstoregiftOptionSearchParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, oService, oOperation)
						Call fnNvstoregiftOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
	'##########################################################################################################################
					endItemErrMsgReplace = replace(SumErrStr, "OK||"&itemid&"||", "")
					endItemErrMsgReplace = replace(SumErrStr, "ERR||"&itemid&"||", "")

					If (Instr(endItemErrMsgReplace, "대분류는 변경할 수 없습니다") > 0) OR (Instr(endItemErrMsgReplace, "대분류는변경할수없습니다") > 0) OR (Instr(endItemErrMsgReplace, "옵션의옵션가/사용여부항목을") > 0) OR (Instr(endItemErrMsgReplace, "옵션의 옵션가/사용여부 항목을") > 0) OR (Instr(endItemErrMsgReplace, "옵션값항목에콤마(,)는") > 0) OR (Instr(endItemErrMsgReplace, "옵션값 항목에 콤마(,)는") > 0) Then
						oService		= "ProductService"
						oOperation		= "DeleteProduct"

						strParam = ""
						strParam = getNvstoregiftDeleteParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, oService, oOperation)
						Call fnNvstoregiftDelete(itemid, strParam, iErrStr, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							failCnt = 0
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				End If
			End If

			'OK던 ERR이던 editQuecnt에 + 1을 시킴..
			'스케줄링에서 editQuecnt ASC, i.lastupdate DESC로 중복을 막자
			strSql = ""
			strSql = strSql & " UPDATE [db_etcmall].[dbo].tbl_nvstoregift_regItem SET " & VBCRLF
			strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
			strSql = strSql & " ,nvstoregiftlastupdate = getdate()  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("nvstoregift", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_nvstoregift_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oNvstoregift = nothing
	End If
End If
'###################################################### 스토어팜 API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
