<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/ebay/utils.asp"-->
<!-- #include virtual="/outmall/ebay/ebayItemcls.asp"-->
<!-- #include virtual="/outmall/ebay/incEbayFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, oEbay, failCnt, chgSellYn, arrRows, skipItem, t11stGoodno, isAllRegYn, getMustprice
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, depth, isItemIdChk, vOptCnt
Dim isoptionyn, isText, i, vGubun, v, cateCode, sIdx, eIdx
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
depth			= request("depth")
vGubun			= request("vGubun")
sIdx			= request("sIdx")
eIdx			= request("eIdx")
failCnt			= 0

Select Case action
	Case "GETSITECATE", "GETCATE", "GETMATCHCATE"	isItemIdChk = "N"
	Case Else	isItemIdChk = "Y"
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

' rw getToken(vGubun)
 'response.end

'######################################################## ebay API ########################################################
If action = "REG" Then
	SET oEbay = new CEbay
		oEbay.FRectItemID	= itemid
		Select Case vGubun
			Case "A"	oEbay.getAuctionNotRegOneItem
			Case "G"	oEbay.getGmarketNotRegOneItem
		End Select

	    If (oEbay.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		' ElseIf (oEbay.FOneItem.FNotinCate = "Y") Then
		'  	iErrStr = "ERR||"&itemid&"||상품 등록 제외 카테고리입니다."
		Else
			Call dummyDataReg(vGubun, itemid, html2db(oEbay.FOneItem.FItemName))
			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oEbay.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oEbay.FOneItem.getEbayItemRegParameter(vGubun)
'response.write strParam
'response.end
				getMustprice = ""
				getMustprice = oEbay.FOneItem.MustPrice()
				Call fnEbayItemReg(getToken(vGubun), itemid, strParam, iErrStr, getMustprice, oEbay.FOneItem.getEbaySellyn, oEbay.FOneItem.FLimityn, oEbay.FOneItem.FLimitNo, oEbay.FOneItem.FLimitSold, html2db(oEbay.FOneItem.FItemName), oEbay.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[AddItem] 옵션검사 실패"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("auction1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oEbay = nothing
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
'###################################################### LotteiMall API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
