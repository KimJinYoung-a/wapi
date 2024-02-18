<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/homeplus/homepluscls.asp"-->
<!-- #include virtual="/outmall/homeplus/incHomeplusFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, oHomeplus, failCnt, chgSellYn, arrRows, skipItem, sellgubun, getMustprice, sellmoney, mallid
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, chkparam, optReset, optString
itemid			= requestCheckVar(request("itemid"),9)
action			= request("action")
chgSellYn		= request("chgSellYn")
failCnt			= 0
mallid			= request("mallid")
If action <> "CategoryView" Then
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
'######################################################## Homeplus API ########################################################
If mallid = "homeplus" Then
	If action = "SOLDOUT" Then					'상태변경
		strParam = ""
		strParam = getHomplusSellynParameter(getHomplusGoodNo(itemid), "N")
		Call fnHomeplusSellyn(itemid, "N", strParam, iErrStr, "setProductStatus")
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("homeplus", itemid, iErrStr)
		End If
		'http://wapi.10x10.co.kr/outmall/proc/HomeplusProc.asp?itemid=699617&mallid=homeplus&action=SOLDOUT
	ElseIf action = "ITEMNAME" Then				'정보수정
		SET oHomeplus = new CHomeplus
			oHomeplus.FRectItemID	= itemid
			oHomeplus.getHomeplusEditOneItem
			If oHomeplus.FResultCount > 0 Then
				strParam = ""
				strParam = oHomeplus.FOneItem.getHomeplusItemEditXML()
				Call fnHomeplusOneItemEdit(itemid, oHomeplus.FOneItem.FHomeplusGoodNo, iErrStr, strParam, "updateProduct")
			Else
				iErrstr = "ERR||"&itemid&"||정보 수정 가능한 상품이 아닙니다."
			End If
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("homeplus", itemid, iErrStr)
			End If
			'http://wapi.10x10.co.kr/outmall/proc/HomeplusProc.asp?itemid=699617&mallid=homeplus&action=ITEMNAME
		SET oHomeplus = nothing
'	ElseIf action = "PRICE" Then				'가격 수정
'		SET oHomeplus = new CHomeplus
'			oHomeplus.FRectItemID	= itemid
'			oHomeplus.getHomeplusEditOneItem
'			If oHomeplus.FResultCount > 0 Then
' 				strParam = ""
'				strParam = oHomeplus.FOneItem.getHomeplusItemEditOPTXML()
'
'				getMustprice = ""
'				getMustprice = oHomeplus.FOneItem.fngetMustPrice()
'				Call fnHomeplusOneItemOPTEdit(itemid, oHomeplus.FOneItem.FHomeplusGoodNo, iErrStr, strParam, getMustprice, "updateProduct")
'				response.write iErrStr
'				If Left(iErrStr, 2) <> "OK" Then
'					CALL Fn_AcctFailTouch("homeplus", itemid, iErrStr)
'				End If
'				'http://wapi.10x10.co.kr/outmall/proc/HomeplusProc.asp?itemid=699617&mallid=homeplus&action=PRICE
'			End If
'		SET oHomeplus = nothing
	ElseIf action = "EDIT" OR action = "PRICE" Then					'정보외 수정
		SET oHomeplus = new CHomeplus
			oHomeplus.FRectItemID	= itemid
			oHomeplus.getHomeplusEditOneItem
			If oHomeplus.FResultCount > 0 Then
				strParam = ""
				iErrStr = ""
				If (oHomeplus.FOneItem.FmaySoldOut = "Y") OR (oHomeplus.FOneItem.IsSoldOutLimit5Sell) Then
					strParam = ""
					strParam = getHomplusSellynParameter(oHomeplus.FOneItem.FHomeplusGoodno, "N")
					Call fnHomeplusSellyn(itemid, "N", strParam, iErrStr, "setProductStatus")
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					If (oHomeplus.FOneItem.FHomeplusSellYn = "N" AND oHomeplus.FOneItem.IsSoldOut = False) Then
						chgSellYn = "Y"
						strParam = getHomplusSellynParameter(oHomeplus.FOneItem.FHomeplusGoodno, "Y")
						Call fnHomeplusSellyn(itemid, "Y", strParam, iErrStr, "setProductStatus")
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

					strParam = ""
					strParam = getHomplusStatChkParameter(itemid)
					Call fnHomeplusOneItemView(itemid, oHomeplus.FOneItem.FHomeplusGoodno, iErrStr, strParam, "searchProduct")
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
	
	 				strParam = ""
					strParam = oHomeplus.FOneItem.getHomeplusItemEditOPTXML()
	
					getMustprice = ""
					getMustprice = oHomeplus.FOneItem.fngetMustPrice()
					Call fnHomeplusOneItemOPTEdit(itemid, oHomeplus.FOneItem.FHomeplusGoodNo, iErrStr, strParam, getMustprice, "updateProduct")
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
	
					strParam = ""
					strParam = getHomplusStatChkParameter(itemid)
					Call fnHomeplusOneItemView(itemid, oHomeplus.FOneItem.FHomeplusGoodno, iErrStr, strParam, "searchProduct")
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
				If failCnt > 0 Then
					SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
					SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
					CALL Fn_AcctFailTouch("homeplus", itemid, SumErrStr)
					response.write "ERR||"&itemid&"||"&SumErrStr
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_homeplus_regItem SET " & VBCRLF
					strSql = strSql & " accFailcnt = 0  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql
	
					SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
					response.write "OK||"&itemid&"||"&SumOKStr
				End If
			End If
			'http://wapi.10x10.co.kr/outmall/proc/HomeplusProc.asp?itemid=1336178&mallid=homeplus&action=EDIT
		SET oHomeplus = nothing
	End If
End If
'###################################################### Homeplus API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->