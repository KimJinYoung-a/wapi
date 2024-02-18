<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/marketfor/marketforItemcls.asp"-->
<!-- #include virtual="/outmall/marketfor/incmarketforFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, olotteon, failCnt, chgSellYn, arrRows, getMustprice, addOptErrItem
Dim iErrStr, strParam, strSql, SumErrStr, SumOKStr, isItemIdChk, grpVal, rSkip, rLimit, i, outmallorderserial
Dim requestJson, responseJson
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
grpVal			= request("grpVal")
rSkip			= request("rSkip")
rLimit			= request("rLimit")
requestJson		= request("requestJson")
responseJson	= request("responseJson")
failCnt			= 0
outmallorderserial = request("outmallorderserial")
addOptErrItem	= "N"

Select Case action
	Case "GOSI", "CATE"
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
'######################################################## Marketfor API ########################################################
If action = "GOSI" Then								'상품정보고시 품목 및 항목코드 조회
	Call fnMarketforGetGosiCode(responseJson)
ElseIf action = "CATE" Then							'마켓포상품분류카테고리 조회
	Call fnMarketforGetClsCateCode(responseJson)
ElseIf action = "EDIT" Then
	iErrStr = "AAAA"
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
'###################################################### Marketfor API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
