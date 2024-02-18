<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/halfclub/halfclubItemcls.asp"-->
<!-- #include virtual="/outmall/halfclub/inchalfclubFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, oHalfclub, failCnt, chgSellYn, arrRows, skipItem, thalfclubGoodno, isAllRegYn, getMustprice
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, ccd, isItemIdChk, vOptCnt, chgImageNm
Dim isoptionyn, isText, i
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
ccd				= request("ccd")
failCnt			= 0

Select Case action
	Case "Delivery"		isItemIdChk = "N"
	Case Else			isItemIdChk = "Y"
End Select

If isItemIdChk = "Y" Then
	If itemid="" or itemid="0" Then
		response.write "<script>alert('��ǰ��ȣ�� �����ϴ�.')</script>"
		response.end
	ElseIf Not(isNumeric(itemid)) Then
		response.write "<script>alert('�߸��� ��ǰ��ȣ�Դϴ�.')</script>"
		response.end
	Else
		'�������·� ��ȯ
		itemid = CLng(getNumeric(itemid))
	End If
End If
'######################################################## halfclub API ########################################################
If action = "REG" Then								'��ǰ���
	SET oHalfclub = new CHalfclub
		oHalfclub.FRectItemID	= itemid
		oHalfclub.gethalfclubNotRegOneItem

	    If (oHalfclub.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_halfclub_regitem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_halfclub_regitem "
	        strSql = strSql & " 	(itemid, regdate, reguserid, HalfClubStatCD, regitemname, HalfClubSellYn)"
	        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oHalfclub.FOneItem.FItemName)&"', 'N')"
			strSql = strSql & " END "
			dbget.Execute strSql
'			If oHalfclub.FOneItem.getMatchingInfoDiv(oHalfclub.FOneItem.FNeedInfoDiv) = "N" Then
'				iErrStr = "ERR||"&itemid&"||ǰ�������� �ٹ����ٰ� ���� �ʽ��ϴ�."
			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oHalfclub.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oHalfclub.FOneItem.getHalfClubItemRegParameter()
				getMustprice = ""
				getMustprice = oHalfclub.FOneItem.MustPrice()
				Call fnhalfclubItemReg(itemid, strParam, iErrStr, getMustprice, oHalfclub.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[REG] �ɼǰ˻� ����"
			End If
		End If
	SET oHalfclub = nothing
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("halfclub", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("halfclub", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDIT" Then							'��ǰ����
	SET oHalfclub = new CHalfclub
		oHalfclub.FRectItemID	= itemid
		oHalfclub.gethalfclubEditOneItem
	    If (oHalfclub.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||[��ǰ����] ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			If (oHalfclub.FOneItem.FmaySoldOut = "Y") OR (oHalfclub.FOneItem.IsSoldOutLimit5Sell) OR (oHalfclub.FOneItem.IsMayLimitSoldout = "Y") Then
				strParam = oHalfclub.FOneItem.getHalfClubItemEditParameter("N")
				chgSellYn = "N"
			Else
				strParam = oHalfclub.FOneItem.getHalfClubItemEditParameter("Y")
				chgSellYn = "Y"
			End If

			getMustprice = ""
			getMustprice = oHalfclub.FOneItem.MustPrice()
			If oHalfclub.FOneItem.isImageChanged Then
				chgImageNm = oHalfclub.FOneItem.getBasicImage
			Else
				chgImageNm = "N"
			End If
			Call fnhalfclubItemEdit(itemid, oHalfclub.FOneItem.FHalfclubGoodNo, iErrStr, strParam, getMustprice, html2db(oHalfclub.FOneItem.FItemName), chgSellYn, chgImageNm)
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("halfclub", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("halfclub", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHalfclub = nothing
ElseIf action = "EditSellYn" Then					'���º���
	SET oHalfclub = new CHalfclub
		oHalfclub.FRectItemID	= itemid
		oHalfclub.gethalfclubEditOneItem

	    If (oHalfclub.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||[���¼���] ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oHalfclub.FOneItem.getHalfClubItemEditParameter(chgSellYn)
			getMustprice = ""
			getMustprice = oHalfclub.FOneItem.MustPrice()
			If oHalfclub.FOneItem.isImageChanged Then
				chgImageNm = oHalfclub.FOneItem.getBasicImage
			Else
				chgImageNm = "N"
			End If
			Call fnhalfclubItemEdit(itemid, oHalfclub.FOneItem.FHalfclubGoodNo, iErrStr, strParam, getMustprice, html2db(oHalfclub.FOneItem.FItemName), chgSellYn, chgImageNm)
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("halfclub", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("halfclub", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHalfclub = nothing
ElseIf action = "PRICE" Then						'��ǰ����
	SET oHalfclub = new CHalfclub
		oHalfclub.FRectItemID	= itemid
		oHalfclub.gethalfclubEditPriceOneItem
	    If (oHalfclub.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||[��ǰ����] ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oHalfclub.FOneItem.getHalfClubPriceParameter()
			getMustprice = ""
			getMustprice = oHalfclub.FOneItem.MustPrice()
			Call fnhalfclubItemEditPrice(itemid, oHalfclub.FOneItem.FHalfclubGoodNo, iErrStr, strParam, getMustprice)
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("halfclub", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("halfclub", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHalfclub = nothing
ElseIf action = "Delivery" Then						'�ù���ڵ� ��ȸ
	Call fnhalfclubDeliveryCode()
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
'###################################################### halfclub API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->