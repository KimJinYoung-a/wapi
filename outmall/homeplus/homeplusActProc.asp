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
Dim itemid, action, oHomeplus, failCnt, chgSellYn, arrRows, skipItem, sellgubun, getMustprice, sellmoney
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, chkparam, optReset, optString
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
failCnt			= 0
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
If action = "REG" Then							'상품등록
	SET oHomeplus = new CHomeplus
		oHomeplus.FRectItemID	= itemid
		oHomeplus.getHomeplusNotRegOneItem
	    If (oHomeplus.FResultCount < 1) Then
			iErrstr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_homeplus_regItem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_homeplus_regItem "
	        strSql = strSql & " (itemid, regdate, reguserid, homeplusstatCD, regitemname)"
	        strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oHomeplus.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql	

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oHomeplus.FOneitem.checkTenItemOptionValid Then
				'//상품등록 파라메터
				strParam = ""
				strParam = oHomeplus.FOneitem.getHomeplusItemRegXML()
				Call fnHomeplusOneItemReg(itemid, strParam, iErrStr, oHomeplus.FOneitem.FSellCash, oHomeplus.FOneitem.getHomeplusSellYn, oHomeplus.FOneitem.FLimityn, oHomeplus.FOneitem.FLimitNo, oHomeplus.FOneitem.FLimitSold, html2db(oHomeplus.FOneitem.FItemName), "createNewProduct")
			Else
				iErrStr = "ERR||"&itemid&"||옵션검사 실패"
			End If

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("homeplus", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("homeplus", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	Set oHomeplus = nothing
ElseIf action = "EditSellYn" Then				'상태변경
	strParam = ""
	strParam = getHomplusSellynParameter(getHomplusGoodNo(itemid), chgSellYn)
	Call fnHomeplusSellyn(itemid, chgSellYn, strParam, iErrStr, "setProductStatus")
'	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("homeplus", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("homeplus", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "ITEMNAME" Then					'정보수정
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

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("homeplus", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("homeplus", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHomeplus = nothing
ElseIf action = "EDIT" Then						'정보외 수정
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
				Call SugiQueLogInsert("homeplus", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
				
				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_homeplus_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				Call SugiQueLogInsert("homeplus", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))

				iErrStr = "OK||"&itemid&"||"&SumOKStr
			End If
		End If
	SET oHomeplus = nothing
ElseIf action = "EditImg" Then					'이미지 수정
	SET oHomeplus = new CHomeplus
		oHomeplus.FRectItemID	= itemid
		oHomeplus.getHomeplusEditOneItem
		If oHomeplus.FResultCount > 0 Then
			strParam = ""
			strParam = oHomeplus.FOneItem.getHomeplusItemEditImgXML()
			Call fnHomeplusOneItemIMGEdit(itemid, oHomeplus.FOneItem.FHomeplusGoodNo, iErrStr, strParam, "updateImage")
		Else
			iErrstr = "ERR||"&itemid&"||이미지 수정 가능한 상품이 아닙니다."
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("homeplus", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("homeplus", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHomeplus = nothing
ElseIf action = "CHKSTAT" Then					'정보조회
	strParam = ""
	strParam = getHomplusStatChkParameter(itemid)
	Call fnHomeplusOneItemStatView(itemid, getHomplusGoodNo(itemid), iErrStr, strParam, "searchProduct")
'	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("homeplus", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("homeplus", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CategoryView" Then				'카테고리조회 및 등록
	Call HomeplusCategoryAPI()
Else
	rw "미지정 ["&action&"]"
End If

If iErrStr <> "" Then
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str += '"&replace(iErrStr, "'", "")&"<br>' " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 2000);" & vbCrLf &_
					"</script>"
End If
'###################################################### Homeplus API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->