<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/base64.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/shopee/shopeeItemcls.asp"-->
<!-- #include virtual="/outmall/shopee/incShopeeFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/shopee/FPost.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, oNvstorefarm, failCnt, chgSellYn, arrRows, skipItem, getMustprice, oService, oOperation, chkXML
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, ccd, isItemIdChk, getfarmGoodno
Dim i, chgImageNm, mayOptSoldOut, endItemErrMsgReplace
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
chkXML			= request("chkXML")
ccd				= request("ccd")
failCnt			= 0

Select Case action
	Case "getShopAuth", "getToken", "refreshToken", "uploadImage", "getCategory"
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
		itemid=CLng(getNumeric(itemid))
	End If
End If
'######################################################## Shopee API ########################################################
If action = "getShopAuth" Then					'v2 Shop Authorization
    Call fnShopAuth()
'http://localhost:11117/outmall/shopee/shopeeActProc.asp?act=getShopAuth
ElseIf action = "getToken" Then					'Access Token
     Call fnAccessToken()
'http://localhost:11117/outmall/shopee/shopeeActProc.asp?act=getToken
ElseIf action = "refreshToken" Then				'Refresh Token
    Call fnRefreshToken()
ElseIf action = "uploadImage" Then				'Image Upload
	Call fnShopeeImageUpload()
ElseIf action = "getCategory" Then
	Call fnShopeeCategory()
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
'###################################################### Shopee API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
