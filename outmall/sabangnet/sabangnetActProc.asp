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
		response.write "<script>alert('��ǰ��ȣ�� �����ϴ�.')</script>"
		response.end
	ElseIf Not(isNumeric(itemid)) Then
		response.write "<script>alert('�߸��� ��ǰ��ȣ�Դϴ�.')</script>"
		response.end
	Else
		'�������·� ��ȯ
		itemid=CLng(getNumeric(itemid))
	End If
End If

'######################################################## sabangnet API ########################################################
If action = "REG" Then									'��ǰ���
	SET oSabangnet = new CSabangnet
		oSabangnet.FRectItemID	= itemid
		oSabangnet.getSabangnetNotRegOneItem
	    If (oSabangnet.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_sabangnet_regitem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_sabangnet_regitem "
	        strSql = strSql & " 	(itemid, regdate, reguserid, sabangnetstatCD, regitemname, sabangnetSellYn)"
	        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oSabangnet.FOneItem.FItemName)&"', 'N')"
			strSql = strSql & " END "
			dbget.Execute strSql

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oSabangnet.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oSabangnet.FOneItem.getSabangnetItemRegParameter(False, "")
				chgImageNm = oSabangnet.FOneItem.getBasicImage
				Call fnSabangnetItemReg(itemid, strParam, iErrStr, oSabangnet.FOneItem.MustPrice, chgImageNm, oSabangnet.FOneItem.FLimityn, oSabangnet.FOneItem.FLimitno, oSabangnet.FOneItem.FLimitsold)
			Else
				iErrStr = "ERR||"&itemid&"||[REG] �ɼǰ˻� ����"
			End If
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("sabangnet", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("sabangnet", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSabangnet = nothing
ElseIf action = "EditSellYn" Then						'���� ����
	SET oSabangnet = new CSabangnet
		oSabangnet.FRectItemID	= itemid
		oSabangnet.getSabangnetSimpleEditOneItem
	    If (oSabangnet.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||[����������] ���� ������ ��ǰ�� �ƴմϴ�."
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
ElseIf action = "PRICE" Then							'���� ����
	SET oSabangnet = new CSabangnet
		oSabangnet.FRectItemID	= itemid
		oSabangnet.getSabangnetSimpleEditOneItem
	    If (oSabangnet.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||[����������] ���� ������ ��ǰ�� �ƴմϴ�."
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
ElseIf action = "EDIT" Then								'��ü ����
	SET oSabangnet = new CSabangnet
		oSabangnet.FRectItemID	= itemid
		oSabangnet.getSabangnetEditOneItem

	    If (oSabangnet.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||[��ü����] ���� ������ ��ǰ�� �ƴմϴ�."
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
ElseIf action = "Category" Then							'����ī�װ��� ���ݿ� �����ϱ�
	strParam = ""
	strParam = get10x10CategoryParameter()
	Call fnRegSabangnetCategory(strParam)
ElseIf action = "GosiInfo" Then							'��ǰ������� ��ȸ �� ����
	Call fnGosiInfoSabangnet(reqDiv)
ElseIf action = "SDATA" Then							'��ǰ ���θ��� DATA ����
	SET oSabangnet = new CSabangnet
		oSabangnet.FRectItemID	= itemid
		oSabangnet.getSabangnetSimpleEditOneItem
	    If (oSabangnet.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||[DATA����] ���� ������ ��ǰ�� �ƴմϴ�."
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