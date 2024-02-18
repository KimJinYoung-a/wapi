<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/wetoo1300k/wetoo1300kItemcls.asp"-->
<!-- #include virtual="/outmall/wetoo1300k/incwetoo1300kFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, o1300k, failCnt, chgSellYn, getMustprice
Dim iErrStr, strParam, strSql, SumErrStr, SumOKStr, isItemIdChk, i
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
failCnt			= 0

''카테고리 끌어올시 하단 실행해야 함..
Select Case action
	Case "cateList"
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
'######################################################## 1300k API ########################################################
If action = "REG" Then								'상품등록
	SET o1300k = new CWetoo1300k
		o1300k.FRectItemID	= itemid
		o1300k.getWetoo1300kNotRegOneItem
	    If (o1300k.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Outmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"', '"&CMALLNAME&"' "
			dbget.execute strSql

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If o1300k.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = o1300k.FOneItem.getWetoo1300kItemRegParameter()
				getMustprice = ""
				getMustprice = o1300k.FOneItem.MustPrice()
				Call fnWetoo1300kItemReg(itemid, strParam, iErrStr, getMustprice, o1300k.FOneItem.FOptionCnt, o1300k.FOneItem.FLimityn, o1300k.FOneItem.FLimitNo, o1300k.FOneItem.FLimitSold, o1300k.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[상품등록] 옵션검사 실패"
			End If
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("wetoo1300k", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("wetoo1300k", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET o1300k = nothing
ElseIf action = "EditSellYn" Then
	SET o1300k = new CWetoo1300k
		o1300k.FRectItemID	= itemid
		o1300k.getWetoo1300kNotEditOneItem
	    If (o1300k.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		Else
			getMustprice = ""
			getMustprice = o1300k.FOneItem.MustPrice()

			strParam = ""
			strParam = o1300k.FOneItem.getWetoo1300kPriceSellynParameter(chgSellYn)
			Call fnWetoo1300kPriceSellyn(itemid, chgSellYn, strParam, getMustprice, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("wetoo1300k", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("wetoo1300k", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET o1300k = nothing
	'http://localhost:11117/outmall/wetoo1300k/wetoo1300kActProc.asp?act=EditSellYn&itemid=1906619&chgSellYn=N
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

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("wetoo1300k", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("wetoo1300k", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET o1300k = nothing
	'http://localhost:11117/outmall/wetoo1300k/wetoo1300kActProc.asp?act=PRICE&itemid=1906619
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
			Call SugiQueLogInsert("wetoo1300k", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("wetoo1300k", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET o1300k = nothing
	'http://localhost:11117/outmall/wetoo1300k/wetoo1300kActProc.asp?act=EDIT&itemid=1906619
ElseIf action = "cateList" Then						'카테고리조회
	Call fnGetCateList()
	response.end
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
'###################################################### 1300k API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
