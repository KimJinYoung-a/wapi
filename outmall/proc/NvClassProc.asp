<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/nvstorefarmClass/nvClassItemcls.asp"-->
<!-- #include virtual="/outmall/nvstorefarmClass/incNvClassFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, mallid, action, oNvclass, failCnt, chgSellYn, arrRows, skipItem, sellgubun, getMustprice
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, getfarmGoodno
Dim oService, oOperation, mayOptSoldOut, chgImageNm
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
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
'######################################################## 스토어팜 API ########################################################
If mallid = "nvstorefarmclass" Then
	If action = "SOLDOUT" Then												'상태변경
		oService		= "ProductService"
		oOperation		= "ChangeProductSaleStatus"

		strParam = ""
		strParam = getNvClassSellynParameter(getNvClassGoodNo(itemid), "N", oService, oOperation)
		Call fnNvClassSellyn(itemid, "N", strParam, iErrStr, oService, oOperation)
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("nvstorefarmclass", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/NvClassProc.asp?itemid=699617&mallid=nvstorefarmclass&action=SOLDOUT
	ElseIf action = "EDIT" OR action = "ITEMNAME" OR action = "PRICE" Then		'상품수정
		SET oNvclass = new CNvClass
			oNvclass.FRectItemID	= itemid
			oNvclass.getNvClassEditOneItem
			If (oNvclass.FResultCount < 1) OR (oNvclass.FOneItem.FNvClassGoodNo = "") Then
				iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
				failCnt = failCnt + 1
			Else

				If oNvclass.FOneItem.FOptioncnt > 0 Then
					mayOptSoldOut = oNvclass.FOneItem.IsMayLimitSoldout
				End If

				If (oNvclass.FOneItem.FMaySoldOut = "Y") OR (mayOptSoldOut = "Y") OR (oNvclass.FOneItem.FLimitYn = "Y" AND oNvclass.FOneItem.getiszeroWonSoldOut(itemid) = "Y") Then
					oService		= "ProductService"
					oOperation		= "ChangeProductSaleStatus"

					strParam = ""
					strParam = getNvClassSellynParameter(oNvclass.FOneItem.FNvClassGoodNo, "N", oService, oOperation)
					Call fnNvClassSellyn(itemid, "N", strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					If (oNvclass.FOneItem.FNvClassSellyn = "N") Then
						oService		= "ProductService"
						oOperation		= "ChangeProductSaleStatus"

						strParam = ""
						strParam = getNvClassSellynParameter(oNvclass.FOneItem.FNvClassGoodNo, "Y", oService, oOperation)
						Call fnNvClassSellyn(itemid, "Y", strParam, iErrStr, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
		'################################################ 0.상품 수정부터(ReturnCostReason 신규 필드 때문..) ####################
					oService		= "ProductService"
					oOperation		= "ManageProduct"

					strParam = ""
					strParam = oNvclass.FOneitem.getNvClassItemRegXML(oService, oOperation, "Y")
					getMustprice = ""
					getMustprice = oNvclass.FOneItem.MustPrice()
					Call fnNvClassItemEDIT(itemid, strParam, iErrStr, getMustprice, oNvclass.FOneItem.FItemName, chgImageNm, oService, oOperation)

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
					strParam = getNvClassOptionSearchParameter(oNvclass.FOneItem.FNvClassGoodNo, oService, oOperation)
					Call fnNvClassOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
		'##########################################################################################################################
		'################################################ 2.이미지 변경시 이미지 재업로드 #########################################
					If oNvclass.FOneItem.isImageChanged = True Then
						chgImageNm = oNvclass.FOneItem.getBasicImage
					Else
						chgImageNm = "N"
					End If

					If chgImageNm <> "N" Then
						oService		= "ImageService"
						oOperation		= "UploadImage"

						strParam = ""
						strParam = oNvclass.FOneitem.getNvClassImageRegXML(oService, oOperation)
						chgImageNm = oNvclass.FOneItem.getBasicImage
						Call fnNvClassImageReg(itemid, strParam, iErrStr, chgImageNm, oService, oOperation)
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
						strParam = oNvclass.FOneitem.getNvClassItemRegXML(oService, oOperation, "Y")
						getMustprice = ""
						getMustprice = oNvclass.FOneItem.MustPrice()
						Call fnNvClassItemEDIT(itemid, strParam, iErrStr, getMustprice, oNvclass.FOneItem.FItemName, chgImageNm, oService, oOperation)
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
						strParam = getNvClassOptionRegXML(itemid, oNvclass.FOneItem.FNvClassGoodNo, oService, oOperation)
						If strParam <> "X" Then
							Call fnNvClassOptionReg(itemid, strParam, iErrStr, oService, oOperation)
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
						strParam = getNvClassOptionSearchParameter(oNvclass.FOneItem.FNvClassGoodNo, oService, oOperation)
						Call fnNvClassOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
		'##########################################################################################################################
				End If
			End If

			'OK던 ERR이던 editQuecnt에 + 1을 시킴..
			'스케줄링에서 editQuecnt ASC, i.lastupdate DESC로 중복을 막자
			strSql = ""
			strSql = strSql & " UPDATE [db_etcmall].[dbo].tbl_nvstorefarmclass_regItem SET " & VBCRLF
			strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
			strSql = strSql & " ,nvClasslastupdate = getdate()  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("nvstorefarmclass", itemid, SumErrStr)
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_nvstorefarmclass_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oNvclass = nothing
	End If
End If
'###################################################### 스토어팜 API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
