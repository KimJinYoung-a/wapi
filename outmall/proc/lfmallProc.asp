﻿<% option explicit %>
<%
Response.CharSet="utf-8"
Session.codepage="65001"
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/lfmall/lfmallItemcls.asp"-->
<!-- #include virtual="/outmall/lfmall/inclfmallFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, mallid, action, oLfmall, failCnt, chgSellYn, arrRows, isItemIdChk, mustPrice, getMustprice
Dim iErrStr, strSql, SumErrStr, SumOKStr, i, strparam, mrgnRate, endItemErrMsgReplace
Dim jenkinsBatchYn, idx, lastErrStr
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
chgSellYn		= request("chgSellYn")
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
'######################################################## LFmall API ########################################################
If mallid = "lfmall" Then
	If action = "SOLDOUT" Then							'상태변경
		SET oLfmall = new CLfmall
			oLfmall.FRectItemID	= itemid
			oLfmall.getLfmallEditOneItem
			If oLfmall.FResultCount = 0 Then
				iErrStr = "ERR||"&itemid&"||상태수정 할 상품이 등록되어 있지 않습니다."
			Else
				strParam = ""
				strParam = oLfmall.FOneItem.getLFmallSellynParameter("N")
				Call fnLfmallSellYN(itemid, strParam, "N", iErrStr)
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("lfmall", itemid, iErrStr)
			End If
		SET oLfmall = nothing
	ElseIf action = "REG" Then							'등록
		SET oLfmall = new CLfmall
			oLfmall.FRectItemID	= itemid
			oLfmall.getLfmallNotRegOneItem
			If (oLfmall.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
			Else
				strSql = "EXEC [db_etcmall].[dbo].[usp_API_Lfmall_RegItem_Add] '"&itemid&"', 'system'"
				dbget.execute strSql

				'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
				If oLfmall.FOneItem.checkTenItemOptionValid Then
					strParam = ""
					strParam = oLfmall.FOneItem.getlfmallItemRegParameter("REG")
					getMustprice = ""
					getMustprice = oLfmall.FOneItem.MustPrice()
					Call fnlfmallItemReg(itemid, strParam, iErrStr, getMustprice, oLfmall.FOneItem.getLfmallSellYn, oLfmall.FOneItem.FLimityn, oLfmall.FOneItem.FLimitNo, oLfmall.FOneItem.FLimitSold, html2db(oLfmall.FOneItem.FItemName), oLfmall.FOneItem.FbasicimageNm)
				Else
					iErrStr = "ERR||"&itemid&"||[등록] 옵션검사 실패"
				End If
			End If
			lastErrStr = iErrStr
			response.write iErrStr

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("lfmall", itemid, iErrStr)
			End If
		SET oLfmall = nothing
	ElseIf action = "EDIT" Then
		SET oLfmall = new CLfmall
			oLfmall.FRectItemID	= itemid
			oLfmall.getLfmallEditOneItem
			If oLfmall.FResultCount = 0 Then
				failCnt = failCnt + 1
				iErrStr = "ERR||"&itemid&"||수정가능한 상품이 아닙니다."
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (oLfmall.FOneItem.FmaySoldOut = "Y") OR (oLfmall.FOneItem.IsMayLimitSoldout = "Y") OR (oLfmall.FOneItem.FOptionCnt = 0 AND oLfmall.FOneItem.getRegedOptionCnt > 1) Then
					strParam = ""
					strParam = oLfmall.FOneItem.getLFmallSellynParameter("N")
					Call fnLfmallSellYN(itemid, strParam, "N", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
				'############## Lfmall 상세 조회 #################
					strParam = ""
					strParam = oLfmall.FOneItem.getLfmallItemViewParameter()
					CALL fnLfmallItemView(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

				'############## Lfmall 재고 0으로 수정 #################
					If failCnt = 0 Then
						strParam = ""
						strParam = oLfmall.FOneItem.getlfmallQuantityParameter("Z")
						Call fnLfmallQuantity(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				'############## Lfmall 상품 수정 #################
					If failCnt = 0 Then
						strParam = ""
						strParam = oLfmall.FOneItem.getlfmallItemRegParameter("EDIT")
						getMustprice = ""
						getMustprice = oLfmall.FOneItem.MustPrice()
						Call fnLfmallItemEdit(itemid, strParam, iErrStr, oLfmall.FOneItem.FbasicimageNm, getMustprice, oLfmall.FOneItem.FLfmallGoodNo, html2db(oLfmall.FOneItem.FItemName))
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				'############## Lfmall 상태 수정 #################
					If failCnt = 0 Then
						strParam = ""
						strParam = oLfmall.FOneItem.getLFmallSellynParameter("Y")
						Call fnLfmallSellYN(itemid, strParam, "Y", iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				'############## Lfmall 상세 조회 #################
					If failCnt = 0 Then
						strParam = ""
						strParam = oLfmall.FOneItem.getLfmallItemViewParameter()
						CALL fnLfmallItemView(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				End If
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("lfmall", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_lfmall_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oLfmall = nothing
	ElseIf action = "CHKSTAT" Then						'상품정보조회_NEW
		SET oLfmall = new CLfmall
			oLfmall.FRectItemID	= itemid
			oLfmall.getLfmallEditOneItem
			If (oLfmall.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||조회 가능한 상품이 아닙니다."
			Else
				strParam = ""
				strParam = oLfmall.FOneItem.getLfmallItemViewParameter()
				CALL fnLfmallItemView(itemid, strParam, iErrStr)
			End If
			lastErrStr = iErrStr
			response.write iErrStr

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("lfmall", itemid, iErrStr)
			End If
		SET oLfmall = nothing
	End If
End If

If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
session.codePage = 949
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->