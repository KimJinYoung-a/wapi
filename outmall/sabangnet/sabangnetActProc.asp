<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/sabangnet/sabangnetItemcls.asp"-->
<!-- #include virtual="/outmall/sabangnet/incSabangnetFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, oSabangnet, failCnt, chgSellYn, arrRows, skipItem, tOptionCnt, tLimityn, isAllRegYn, getMustprice, tIsChildrenCate
Dim iErrStr, strParam, mustPrice, displayDate, ret1, strSql, SumErrStr, SumOKStr, iitemname, isItemIdChk, isFiftyUpDown, isiframe
Dim isoptionyn, i, chgImageNm, reqDiv
Dim failCnt2
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
reqDiv			= request("reqDiv")
failCnt			= 0
failCnt2		= 0

Select Case action
	Case "Category", "GosiInfo"		isItemIdChk = "N"
	Case Else						isItemIdChk = "Y"
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

'######################################################## sabangnet API ########################################################
If action = "REG" Then									'상품등록
	SET oSabangnet = new CSabangnet
		oSabangnet.FRectItemID	= itemid
		oSabangnet.getSabangnetNotRegOneItem
	    If (oSabangnet.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_sabangnet_regitem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_sabangnet_regitem "
	        strSql = strSql & " 	(itemid, regdate, reguserid, sabangnetstatCD, regitemname, sabangnetSellYn)"
	        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oSabangnet.FOneItem.FItemName)&"', 'N')"
			strSql = strSql & " END "
			dbget.Execute strSql

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oSabangnet.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oSabangnet.FOneItem.getSabangnetItemRegParameter(False, "")
				chgImageNm = oSabangnet.FOneItem.getBasicImage
				Call fnSabangnetItemReg(itemid, strParam, iErrStr, oSabangnet.FOneItem.MustPrice, chgImageNm, oSabangnet.FOneItem.FLimityn, oSabangnet.FOneItem.FLimitno, oSabangnet.FOneItem.FLimitsold)
			Else
				iErrStr = "ERR||"&itemid&"||[REG] 옵션검사 실패"
			End If
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("sabangnet", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("sabangnet", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSabangnet = nothing
ElseIf action = "EditSellYn" Then						'상태 수정
	SET oSabangnet = new CSabangnet
		oSabangnet.FRectItemID	= itemid
		oSabangnet.getSabangnetSimpleEditOneItem
	    If (oSabangnet.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||[요약수정상태] 수정 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = oSabangnet.FOneItem.getSabangnetSimpleEditItemParameter(chgSellYn)
			Call fnSabangnetSimpleEdit(itemid, chgSellYn, oSabangnet.FOneItem.MustPrice, html2db(oSabangnet.FOneItem.FItemName), strParam, iErrStr, "sellyn")
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("sabangnet", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("sabangnet", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSabangnet = nothing
ElseIf action = "PRICE" Then							'가격 수정
	SET oSabangnet = new CSabangnet
		oSabangnet.FRectItemID	= itemid
		oSabangnet.getSabangnetSimpleEditOneItem
	    If (oSabangnet.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||[요약수정가격] 수정 가능한 상품이 아닙니다."
		Else
			strParam = ""
			If (oSabangnet.FOneItem.FmaySoldOut = "Y") OR (oSabangnet.FOneItem.IsMayLimitSoldout = "Y") OR (oSabangnet.FOneItem.IsSoldOut) Then
				chgSellYn = "N"
				strParam = oSabangnet.FOneItem.getSabangnetSimpleEditItemParameter(chgSellYn)
			Else
				chgSellYn = "Y"
				strParam = oSabangnet.FOneItem.getSabangnetSimpleEditItemParameter(chgSellYn)
			End If

			Call fnSabangnetSimpleEdit(itemid, chgSellYn, oSabangnet.FOneItem.MustPrice, html2db(oSabangnet.FOneItem.FItemName), strParam, iErrStr, "price")
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("sabangnet", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("sabangnet", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSabangnet = nothing
ElseIf action = "EDIT" Then								'전체 수정
	SET oSabangnet = new CSabangnet
		oSabangnet.FRectItemID	= itemid
		oSabangnet.getSabangnetEditOneItem

	    If (oSabangnet.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||[전체수정] 수정 가능한 상품이 아닙니다."
		Else
			If (oSabangnet.FOneItem.FmaySoldOut = "Y") OR (oSabangnet.FOneItem.IsMayLimitSoldout = "Y") OR (oSabangnet.FOneItem.IsSoldOut) Then
				chgSellYn = "N"
			Else
				chgSellYn = "Y"
			End If
			strParam = ""
			strParam = oSabangnet.FOneItem.getSabangnetItemRegParameter(True, chgSellYn)
			Call fnSabangnetItemEdit(itemid, strParam, iErrStr, oSabangnet.FOneItem.MustPrice, chgImageNm, oSabangnet.FOneItem.FLimityn, oSabangnet.FOneItem.FLimitno, oSabangnet.FOneItem.FLimitsold, chgSellYn)
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("sabangnet", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("sabangnet", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSabangnet = nothing
ElseIf action = "Category" Then							'관리카테고리를 사방넷에 저장하기
	strParam = ""
	strParam = get10x10CategoryParameter()
	Call fnRegSabangnetCategory(strParam)
ElseIf action = "GosiInfo" Then							'상품정보고시 조회 후 저장
	Call fnGosiInfoSabangnet(reqDiv)
ElseIf action = "SDATA" Then							'상품 쇼핑몰별 DATA 수정
	SET oSabangnet = new CSabangnet
		oSabangnet.FRectItemID	= itemid
		oSabangnet.getSabangnetSimpleEditOneItem
	    If (oSabangnet.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||[DATA수정] 수정 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = oSabangnet.FOneItem.getSabangnetShoppingMallEditParameter()
			Call fnShoppingDataSabangnet(itemid, strParam, iErrStr)
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("sabangnet", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("sabangnet", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSabangnet = nothing
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
'###################################################### sabangnet API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->