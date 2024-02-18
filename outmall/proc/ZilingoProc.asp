<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/zilingo/zilingoItemcls.asp"-->
<!-- #include virtual="/outmall/zilingo/inczilingoFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, itemoption, mallid, action, oZilingo, failCnt, chgSellYn, arrRows
Dim iErrStr, strParam, strSql, SumErrStr, SumOKStr, iitemname
Dim vMy11stGoodno, strSKUGoodNo, quantity, maylimitEa, strTmpGoodNo
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
itemid			= requestCheckVar(request("itemid"),9)
itemoption		= requestCheckVar(request("itemoption"),4)

If itemoption = "" Then
	itemoption = "0000"
End If

If itemid="" or itemid="0" Then
	response.write "<script>alert('상품번호가 없습니다.')</script>"
	response.end
Else
	'정수형태로 변환
	itemid=CLng(getNumeric(itemid))
End If
'######################################################## ZILINGO API ########################################################
If mallid = "zilingo" Then
	If action = "SOLDOUT" Then													'상태 변경
		strSKUGoodNo = getSKUZilingoGoodNo2(itemid, itemoption, quantity)
		If strSKUGoodNo = "" Then
			iErrStr = "ERR||"&itemid&"||"&itemoption&"||등록하지 않은 상품입니다."
		Else
			strParam = ""
			strParam = fnZilingoQuantitySoldOutJSON(itemid, itemoption, quantity, maylimitEa, strSKUGoodNo)
			Call fnZilingoEditQuantityZero(itemid, itemoption, maylimitEa, strParam, iErrStr)
		End If
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouchEtcOption("zilingo", itemid, itemoption, iErrStr)
		End If 
		'http://wapi.10x10.co.kr/outmall/proc/ZilingoProc.asp?itemid=1867487&itemoption=0000&mallid=zilingo&action=SOLDOUT
	ElseIf action = "PRICE" Then												'가격 수정
		SET oZilingo = new CZilingo
			oZilingo.FRectItemID		= itemid
			oZilingo.FRectitemOption	= itemoption
			oZilingo.getZilingoPriceOneItem
			If (oZilingo.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||"&itemoption&"||가격 수정 가능한 상품이 없습니다."
			Else
				strParam = ""
				strParam = oZilingo.FOneItem.getZilingoPriceJSON()
				Call fnZilingoItemPrice(itemid, itemoption, strParam, oZilingo.FOneItem.FOrgprice, oZilingo.FOneItem.FWonprice, oZilingo.FOneItem.FMultiplerate, oZilingo.FOneItem.FExchangeRate, iErrStr)
			End If
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouchEtcOption("zilingo", itemid, itemoption, iErrStr)
			End If
		SET oZilingo = nothing
		'http://wapi.10x10.co.kr/outmall/proc/ZilingoProc.asp?itemid=1867487&itemoption=0000&mallid=zilingo&action=PRICE
	ElseIf action = "CHKSTAT" Then												'승인 조회
		strTmpGoodNo = getTmpZilingoGoodNo(itemid, itemoption)
		If strTmpGoodNo = "" Then
			iErrStr = "ERR||"&itemid&"||"&itemoption&"||등록하지 않은 상품입니다."
		Else
			Call fnZilingoTmpGoodNo(itemid, itemoption, strTmpGoodNo, iErrStr)
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouchEtcOption("zilingo", itemid, itemoption, iErrStr)
		End If
		response.write iErrStr
		'http://wapi.10x10.co.kr/outmall/proc/ZilingoProc.asp?itemid=1867487&itemoption=0000&mallid=zilingo&action=CHKSTAT
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

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||"&itemoption&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||"&itemoption&"||", "")
				CALL Fn_AcctFailTouchEtcOption("zilingo", itemid, itemoption, SumErrStr)
				response.write "ERR||"&itemid&"||"&itemoption&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||"&itemoption&"||", "")
				response.write "OK||"&itemid&"||"&itemoption&"||"&SumOKStr
			End If
		End If
		'http://wapi.10x10.co.kr/outmall/proc/ZilingoProc.asp?itemid=1867487&itemoption=0000&mallid=zilingo&action=EDITQTY
	End If
End If
'###################################################### ZILINGO API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->