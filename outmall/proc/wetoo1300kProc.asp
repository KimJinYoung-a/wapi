<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/wetoo1300k/wetoo1300kItemcls.asp"-->
<!-- #include virtual="/outmall/wetoo1300k/incwetoo1300kFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, mallid, action, o1300k, failCnt, chgSellYn, getMustprice
Dim iErrStr, strParam, strSql, SumErrStr, SumOKStr, isItemIdChk, i
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
'######################################################## 1300k API ########################################################
If mallid = "wetoo1300k" Then
	If action = "SOLDOUT" Then
		SET o1300k = new CWetoo1300k
			o1300k.FRectItemID	= itemid
			o1300k.getWetoo1300kNotEditOneItem
			If (o1300k.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
			Else
				getMustprice = ""
				getMustprice = o1300k.FOneItem.MustPrice()

				strParam = ""
				strParam = o1300k.FOneItem.getWetoo1300kPriceSellynParameter("N")
				Call fnWetoo1300kPriceSellyn(itemid, "N", strParam, getMustprice, iErrStr)
			End If
			lastErrStr = iErrStr
			response.write iErrStr

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("wetoo1300k", itemid, iErrStr)
			End If
		SET o1300k = nothing
		'http://localhost:11117/outmall/proc/wetoo1300kProc.asp?mallid=wetoo1300k&act=SOLDOUT&itemid=1906619
	ElseIf action = "PRICE" Then
		SET o1300k = new CWetoo1300k
			o1300k.FRectItemID	= itemid
			o1300k.getWetoo1300kNotEditOneItem
			If (o1300k.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
			Else
				getMustprice = ""
				getMustprice = o1300k.FOneItem.MustPrice()
				If (o1300k.FOneItem.FmaySoldOut = "Y") OR (o1300k.FOneItem.IsMayLimitSoldout = "Y") OR (o1300k.FOneItem.FLimityn = "Y" AND (o1300k.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) Then
					chgSellYn = "N"
				Else
					chgSellYn = "Y"
				End IF

				strParam = ""
				strParam = o1300k.FOneItem.getWetoo1300kPriceSellynParameter(chgSellYn)
				Call fnWetoo1300kPriceSellyn(itemid, chgSellYn, strParam, getMustprice, iErrStr)
			End If
			lastErrStr = iErrStr
			response.write iErrStr

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("wetoo1300k", itemid, iErrStr)
			End If
		SET o1300k = nothing
		'http://localhost:11117/outmall/proc/wetoo1300kProc.asp?mallid=wetoo1300k&act=PRICE&itemid=1906619
	ElseIf action = "EDIT" Then
		SET o1300k = new CWetoo1300k
			o1300k.FRectItemID	= itemid
			o1300k.getWetoo1300kNotEditOneItem
			If (o1300k.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
			Else
				If (o1300k.FOneItem.FmaySoldOut = "Y") OR (o1300k.FOneItem.IsMayLimitSoldout = "Y") OR (o1300k.FOneItem.FLimityn = "Y" AND (o1300k.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) Then
					getMustprice = ""
					getMustprice = o1300k.FOneItem.MustPrice()

					strParam = ""
					strParam = o1300k.FOneItem.getWetoo1300kPriceSellynParameter("N")
					Call fnWetoo1300kPriceSellyn(itemid, "N", strParam, getMustprice, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					strParam = ""
					strParam = o1300k.FOneItem.getWetoo1300kItemEditParameter()
					getMustprice = ""
					getMustprice = o1300k.FOneItem.MustPrice()
					Call fnWetoo1300kItemEdit(itemid, strParam, iErrStr, getMustprice, o1300k.FOneItem.FOptionCnt, o1300k.FOneItem.FLimityn, o1300k.FOneItem.FLimitNo, o1300k.FOneItem.FLimitSold, o1300k.FOneItem.FbasicimageNm)
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
				CALL Fn_AcctFailTouch("wetoo1300k", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET o1300k = nothing
		'http://localhost:11117/outmall/proc/wetoo1300kProc.asp?mallid=wetoo1300k&act=PRICE&itemid=3471386
	End If
End If
'###################################################### 1300k API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
