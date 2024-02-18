<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/zilingo/zilingoItemcls.asp"-->
<!-- #include virtual="/outmall/zilingo/incZilingoFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, itemoption, action, oZilingo, failCnt, chgSellYn, arrRows, skipItem, tGmarketGoodno, tLimityn, getMustprice
Dim iErrStr, strParam, mustPrice, displayDate, ret1, strSql, SumErrStr, SumOKStr, iitemname, isItemIdChk
Dim i, newItemid, newItemname, strTmpGoodNo, quantity, strSKUGoodNo, maylimitEa
Dim failCnt2
itemid			= requestCheckVar(Split(request("itemid"), "_")(0),9)
itemoption		= requestCheckVar(Split(request("itemid"), "_")(1),4)

If itemoption = "" Then
	itemoption = "0000"
End If

newItemid		= itemid&"_"&itemoption
action			= request("act")
chgSellYn		= request("chgSellYn")
failCnt			= 0
failCnt2		= 0

Select Case action
	Case "SubCategory"	isItemIdChk = "N"
	Case Else			isItemIdChk = "Y"
End Select

If isItemIdChk = "Y" Then
	If itemid="" or itemid="0" Then
		response.write "<script>alert('상품번호가 없습니다.')</script>"
		response.end
	Else
		'정수형태로 변환
		itemid = CLng(getNumeric(itemid))
	End If
End If

'######################################################## Zilingo API ########################################################
If action = "REG" Then
	SET oZilingo = new CZilingo
		oZilingo.FRectItemID		= itemid
		oZilingo.FRectitemOption	= itemoption
		If itemoption <> "0000" Then
			If oZilingo.fnMaySoldout(itemid, itemoption) = "Y" Then
				iErrStr = "ERR||"&itemid&"||"&itemoption&"||등록가능한 상품이 아닙니다."
			End If
		End If
		oZilingo.getZilingoNotRegOneItem
		If (oZilingo.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||"&itemoption&"||등록가능한 상품이 없습니다."
		Else
			newItemname = oZilingo.fnZilingoItemname(itemid, itemoption, oZilingo.FOneitem.FChgItemName)
			strSql = ""
			strSql = strSql & " IF NOT EXISTS(SELECT * FROM db_etcmall.[dbo].[tbl_zilingo_regItem] WHERE itemid="&itemid&" and itemoption = '"&itemoption&"')"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_zilingo_regItem] "
			strSql = strSql & " (itemid, itemoption, regdate, reguserid, zilingostatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", '"&itemoption&"', getdate(), '"&session("SSBctID")&"', '1', '"&html2db(newItemname)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			strParam = ""
			strParam = oZilingo.FOneItem.getZilingoItemRegJSON(newItemname, itemoption, quantity)
			'response.write strParam
			Call fnZilingoItemReg(itemid, itemoption, strParam, oZilingo.FOneItem.FOrgprice, oZilingo.FOneItem.FWonprice, oZilingo.FOneItem.FMultiplerate, oZilingo.FOneItem.FExchangeRate, quantity, iErrStr)
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouchEtcOption("zilingo", itemid, itemoption, iErrStr)
		End If
		Call SugiQueLogInsertByOption("zilingo", action, itemid, itemoption, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oZilingo = nothing
ElseIf action = "EditSellYn" Then
	strSKUGoodNo = getSKUZilingoGoodNo2(itemid, itemoption, quantity)
	If strSKUGoodNo = "" Then
		iErrStr = "ERR||"&itemid&"||"&itemoption&"||등록하지 않은 상품입니다."
	Else
		strParam = ""
		strParam = fnZilingoQuantitySoldOutJSON(itemid, itemoption, quantity, maylimitEa, strSKUGoodNo)
		Call fnZilingoEditQuantityZero(itemid, itemoption, maylimitEa, strParam, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchEtcOption("zilingo", itemid, itemoption, iErrStr)
	End If
	Call SugiQueLogInsertByOption("zilingo", action, itemid, itemoption, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "PRICE" Then
	SET oZilingo = new CZilingo
		oZilingo.FRectItemID		= itemid
		oZilingo.FRectitemOption	= itemoption
		oZilingo.getZilingoPriceOneItem
		If (oZilingo.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||"&itemoption&"||가격 수정 가능한 상품이 없습니다."
		Else
'			strParam = ""
'			strParam = oZilingo.FOneItem.getZilingoPriceBySkuNoJSON()
'			Call fnZilingoItemPriceBySkuNo(itemid, itemoption, strParam, oZilingo.FOneItem.FOrgprice, oZilingo.FOneItem.FWonprice, oZilingo.FOneItem.FMultiplerate, oZilingo.FOneItem.FExchangeRate, iErrStr)
			strParam = ""
			strParam = oZilingo.FOneItem.getZilingoPriceJSON()
			Call fnZilingoItemPrice(itemid, itemoption, strParam, oZilingo.FOneItem.FOrgprice, oZilingo.FOneItem.FWonprice, oZilingo.FOneItem.FMultiplerate, oZilingo.FOneItem.FExchangeRate, iErrStr)
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouchEtcOption("zilingo", itemid, itemoption, iErrStr)
		End If
		Call SugiQueLogInsertByOption("zilingo", action, itemid, itemoption, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oZilingo = nothing
ElseIf action = "QTY" Then
	strSKUGoodNo = getSKUZilingoGoodNo2(itemid, itemoption, quantity)
	If strSKUGoodNo = "" Then
		iErrStr = "ERR||"&itemid&"||"&itemoption&"||등록하지 않은 상품입니다."
	Else
		strParam = ""
		strParam = fnZilingoQuantityEditJSON(itemid, itemoption, quantity, maylimitEa, strSKUGoodNo)
		Call fnZilingoEditQuantity(itemid, itemoption, maylimitEa, strParam, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchEtcOption("zilingo", itemid, itemoption, iErrStr)
	End If
	Call SugiQueLogInsertByOption("zilingo", action, itemid, itemoption, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDITQTY" Then
	strSKUGoodNo = getSKUZilingoGoodNo2(itemid, itemoption, quantity)
	If strSKUGoodNo = "" Then
		iErrStr = "ERR||"&itemid&"||"&itemoption&"||등록하지 않은 상품입니다."
	Else
		strParam = ""
		strParam = fnZilingoQuantitySearchJSON(strSKUGoodNo)
		Call fnZilingoSKUGoodNo(itemid, itemoption, strParam, iErrStr)
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If

		If failCnt = 0 Then
			strParam = ""
			strParam = fnZilingoQuantityEditJSON(itemid, itemoption, quantity, maylimitEa, strSKUGoodNo)
			Call fnZilingoEditQuantity(itemid, itemoption, maylimitEa, strParam, iErrStr)
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
'OK||1867487||0000||[CHKQTY]성공
		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||"&itemoption&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||"&itemoption&"||", "")
			CALL Fn_AcctFailTouchEtcOption("zilingo", itemid, itemoption, SumErrStr)
			Call SugiQueLogInsertByOption("zilingo", action, itemid, itemoption, "ERR", "ERR||"&itemid&"||"&itemoption&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&itemoption&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||"&itemoption&"||", "")
			Call SugiQueLogInsertByOption("zilingo", action, itemid, itemoption, "OK", "OK||"&itemid&"||"&itemoption&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&itemoption&"||"&SumOKStr
		End If
	End If
ElseIf action = "CHKSTAT" Then
	strTmpGoodNo = getTmpZilingoGoodNo(itemid, itemoption)
	If strTmpGoodNo = "" Then
		iErrStr = "ERR||"&itemid&"||"&itemoption&"||등록하지 않은 상품입니다."
	Else
		Call fnZilingoTmpGoodNo(itemid, itemoption, strTmpGoodNo, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchEtcOption("zilingo", itemid, itemoption, iErrStr)
	End If
	Call SugiQueLogInsertByOption("zilingo", action, itemid, itemoption, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKQUANTITY" Then
	strSKUGoodNo = getSKUZilingoGoodNo(itemid, itemoption)
	If strSKUGoodNo = "" Then
		iErrStr = "ERR||"&itemid&"||"&itemoption&"||등록하지 않은 상품입니다."
	Else
		strParam = ""
		strParam = fnZilingoQuantitySearchJSON(strSKUGoodNo)
		Call fnZilingoSKUGoodNo(itemid, itemoption, strParam, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchEtcOption("zilingo", itemid, itemoption, iErrStr)
	End If
	Call SugiQueLogInsertByOption("zilingo", action, itemid, itemoption, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "SubCategory" Then
	Call fnZilingoSubCategory()
End If

If iErrStr <> "" Then
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str += '"&iErrStr&"<br>' " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
					"</script>"
End If
'###################################################### Zilingo API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->