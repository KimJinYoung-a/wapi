<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/gmarket/gmarketItemcls.asp"-->
<!-- #include virtual="/outmall/gmarket/incGmarketFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, oGmarket, oGmarketOpt, failCnt, chgSellYn, arrRows, skipItem, tGmarketGoodno, tOptionCnt, tLimityn, isAllRegYn, getMustprice, tIsChildrenCate
Dim iErrStr, strParam, mustPrice, displayDate, ret1, strSql, SumErrStr, SumOKStr, iitemname, isItemIdChk, isFiftyUpDown, isiframe
Dim gMakername, gBrandname, contentUpdateCnt
Dim isChild, isLife, isElec
Dim isoptionyn, isText, i
Dim failCnt2
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
gMakername		= Trim(request("gMakername"))
gBrandname		= Trim(request("gBrandname"))
failCnt			= 0
failCnt2		= 0

Select Case action
	Case "AddMakerBrand", "AddAddressBook", "RequestAddressBook", "CATE"	isItemIdChk = "N"
	Case Else								isItemIdChk = "Y"
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
'######################################################## Gmarket API ########################################################
If action = "REGAddItem" Then							'��ǰ �⺻ ���� ���
	SET oGmarket = new CGmarket
		oGmarket.FRectItemID	= itemid
		oGmarket.getGmarketNotRegOneItem
	    If (oGmarket.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf (oGmarket.FOneItem.FDepthCode = "0") Then
			iErrStr = "ERR||"&itemid&"||ī�װ� ��Ī�� �ʿ��մϴ�."
		' ElseIf (oGmarket.FOneItem.FBrandCode = "0") Then
		' 	iErrStr = "ERR||"&itemid&"||�귣�� ��Ī�� �ʿ��մϴ�."
		ElseIf oGmarket.FOneItem.checkItemContent = "Y" Then
			iErrStr = "ERR||"&itemid&"||iframe�� ���� ��ǰ�� ��� �� �� �����ϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_gmarket_regitem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_gmarket_regitem "
			strSql = strSql & " 	(itemid, regdate, reguserid, gmarketstatCD, regitemname, gmarketSellYn)"
			strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oGmarket.FOneItem.FItemName)&"', 'N')"
			strSql = strSql & " END "
			dbget.Execute strSql

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oGmarket.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketItemRegParameter(FALSE)
				Call fnGmarketItemReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[AddItem] �ɼǰ˻� ����"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("gmarket1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oGmarket = nothing
ElseIf action = "REGInfoCd" Then						'��ǰ��� + ������� ���
	tGmarketGoodno = getGmarketGoodno(itemid)
	If tGmarketGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
	Else
		strParam = ""
		strParam = getGmarketInfoCdParameter(itemid, tGmarketGoodno)
		Call fnGmarketItemInfoCd(itemid, strParam, iErrStr)
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If

		If failCnt = 0 Then
			Call getGmarketChildrenCate(itemid, isChild, isLife, isElec)
			If isChild = "Y" OR isLife = "Y" OR isElec = "Y" Then
				strParam = ""
				strParam = getGmarketChildrenParameter(itemid, tGmarketGoodno, isChild, isLife, isElec)
				Call fnGmarketItemChildren(itemid, strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		End If
	End If

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
		Call SugiQueLogInsert("gmarket1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("gmarket1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "REGAddOPT" Then
	SET oGmarket = new CGmarket
		oGmarket.FRectItemID	= itemid
		oGmarket.getGmarketNotOptOneItem

	    If (oGmarket.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||�ɼ� ��� ������ ��ǰ�� �ƴմϴ�."
		ElseIf (oGmarket.FOneItem.FGmarketGoodNo = "") Then
			iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
		ElseIf (oGmarket.FOneItem.FAPIadditem = "N") Then
			iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
		ElseIf (oGmarket.FOneItem.getFiftyUpDown = "Y") Then
			iErrStr = "ERR||"&itemid&"||�ɼǰ����� 50%�� �ʰ��մϴ�."
		Else
			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oGmarket.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketItemOptRegParameter()
				Call fnGmarketOPTReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FLimityn, oGmarket.FOneItem.FLimitno, oGmarket.FOneItem.FLimitsold)
			Else
				iErrStr = "ERR||"&itemid&"||[AddOPT] �ɼǰ˻� ����"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("gmarket1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oGmarket = nothing
ElseIf action = "REG" Then								'�⺻���� + ������� + �ɼ����� ���
	'##################################### �⺻ ���� ��� ���� #####################################
	SET oGmarket = new CGmarket
		oGmarket.FRectItemID	= itemid
		oGmarket.getGmarketNotRegOneItem
	    If (oGmarket.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf (oGmarket.FOneItem.FDepthCode = "0") Then
			iErrStr = "ERR||"&itemid&"||ī�װ� ��Ī�� �ʿ��մϴ�."
		ElseIf oGmarket.FOneItem.checkItemContent = "Y" Then
			iErrStr = "ERR||"&itemid&"||iframe�� ���� ��ǰ�� ��� �� �� �����ϴ�."
		' ElseIf (oGmarket.FOneItem.FBrandCode = "0") Then
		' 	iErrStr = "ERR||"&itemid&"||�귣�� ��Ī�� �ʿ��մϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_gmarket_regitem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_gmarket_regitem "
	        strSql = strSql & " 	(itemid, regdate, reguserid, gmarketstatCD, regitemname, gmarketSellYn)"
	        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oGmarket.FOneItem.FItemName)&"', 'N')"
			strSql = strSql & " END "
			dbget.Execute strSql

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oGmarket.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketItemRegParameter(FALSE)
				'getMustprice = ""
				'getMustprice = oGmarket.FOneItem.MustPrice()
				Call fnGmarketItemReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[AddItem] �ɼǰ˻� ����"
			End If
		End If
	SET oGmarket = nothing
	If Left(iErrStr, 2) <> "OK" Then
		failCnt = failCnt + 1
		SumErrStr = SumErrStr & iErrStr
	Else
		SumOKStr = SumOKStr & iErrStr
	End If
	'##################################### �⺻ ���� ��� �� #####################################

	'#################################### ��� ���� ��� ���� ####################################
	If failCnt = 0 Then
		tGmarketGoodno = getGmarketGoodno(itemid)
		If tGmarketGoodno = "" Then
			iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
		Else
			strParam = ""
			strParam = getGmarketInfoCdParameter(itemid, tGmarketGoodno)
			Call fnGmarketItemInfoCd(itemid, strParam, iErrStr)
		End If

		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'#################################### ��� ���� ��� �� ####################################

	'#################################### ��� ���� ��� ���� ####################################
	If failCnt = 0 Then
		Call getGmarketChildrenCate(itemid, isChild, isLife, isElec)
		If isChild = "Y" OR isLife = "Y" OR isElec = "Y" Then
			strParam = ""
			strParam = getGmarketChildrenParameter(itemid, tGmarketGoodno, isChild, isLife, isElec)
			Call fnGmarketItemChildren(itemid, strParam, iErrStr)
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
	End If
	'#################################### ��� ���� ��� �� ####################################

	'#################################### ��ǰ�� ��� ���� ####################################
	If failCnt = 0 Then
		strParam = ""
		strParam = getGmarketReturnFeeParameter(itemid, tGmarketGoodno, CRETURNFEE)
		Call fnGmarketReturnFee(itemid, strParam, CRETURNFEE, iErrStr)

		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'#################################### ��ǰ�� ��� �� ####################################

	'#################################### �ɼ� ���� ��� ���� ####################################
	If failCnt = 0 Then
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketNotOptOneItem
		    If (oGmarket.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||�ɼ� ��� ������ ��ǰ�� �ƴմϴ�."
			ElseIf (oGmarket.FOneItem.FGmarketGoodNo = "") Then
				iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
			ElseIf (oGmarket.FOneItem.FAPIadditem = "N") Then
				iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
			ElseIf (oGmarket.FOneItem.getFiftyUpDown = "Y") Then
				iErrStr = "ERR||"&itemid&"||�ɼǰ����� 50%�� �ʰ��մϴ�."
			Else
				'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
				If oGmarket.FOneItem.checkTenItemOptionValid Then
					strParam = ""
					strParam = oGmarket.FOneItem.getGmarketItemOptRegParameter()
					Call fnGmarketOPTReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FLimityn, oGmarket.FOneItem.FLimitno, oGmarket.FOneItem.FLimitsold)
				Else
					iErrStr = "ERR||"&itemid&"||[AddOPT] �ɼǰ˻� ����"
				End If
			End If
		SET oGmarket = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'#################################### �ɼ� ���� ��� �� ####################################

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
		Call SugiQueLogInsert("gmarket1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("gmarket1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "REGOnSale" Then						'�űԵ�� ��ǰ �Ǹ������� ����
	isAllRegYn = getAllRegChk(itemid, "X")
	If isAllRegYn <> "Y" Then
		iErrStr = "ERR||"&itemid&"||�⺻����, �ɼ�����, ��ǰ��� �Է��� Ȯ���ϼ���"
	Else
		tGmarketGoodno = getGmarketGoodno(itemid)
		strParam = ""
		strParam = getGmarketAddPriceParameter(itemid, tGmarketGoodno, "Y", mustPrice, displayDate)
		Call fnGmarketItemAddPrice(itemid, strParam, mustPrice, displayDate, "Y", iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("gmarket1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "REGOnSale2" Then						'�߰��ݾ� ���� ��ǰ �Ǹ������� ����
	isAllRegYn = getAllRegChk(itemid, "O")
	If isAllRegYn <> "Y" Then
		iErrStr = "ERR||"&itemid&"||�⺻����, ��ǰ��� �Է��� Ȯ���ϼ���"
	Else
		tGmarketGoodno = getGmarketGoodno(itemid)
		strParam = ""
		strParam = getGmarketAddPriceParameter(itemid, tGmarketGoodno, "Y", mustPrice, displayDate)
		Call fnGmarketItemAddPrice(itemid, strParam, mustPrice, displayDate, "Y", iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("gmarket1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "REG2" Then								'OnSale + �ɼ� ���	// �߰��ݾ� ������ ���
	isAllRegYn = getAllRegChk(itemid, "O")
	If isAllRegYn <> "Y" Then
		iErrStr = "ERR||"&itemid&"||�⺻����, ��ǰ��� �Է��� Ȯ���ϼ���"
	Else
		tGmarketGoodno = getGmarketGoodno(itemid)
		strParam = ""
		strParam = getGmarketAddPriceParameter(itemid, tGmarketGoodno, "Y", mustPrice, displayDate)
		Call fnGmarketItemAddPrice(itemid, strParam, mustPrice, displayDate, "Y", iErrStr)
	End If
	If Left(iErrStr, 2) <> "OK" Then
		failCnt = failCnt + 1
		SumErrStr = SumErrStr & iErrStr
	Else
		SumOKStr = SumOKStr & iErrStr
	End If

	'#################################### �ɼ� ���� ��� ���� ####################################
	If failCnt = 0 Then
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketNotOptOneItem
		    If (oGmarket.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||�ɼ� ��� ������ ��ǰ�� �ƴմϴ�."
			ElseIf (oGmarket.FOneItem.FGmarketGoodNo = "") Then
				iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
			ElseIf (oGmarket.FOneItem.FAPIadditem = "N") Then
				iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
			ElseIf (oGmarket.FOneItem.getFiftyUpDown = "Y") Then
				iErrStr = "ERR||"&itemid&"||�ɼǰ����� 50%�� �ʰ��մϴ�."
			Else
				'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
				If oGmarket.FOneItem.checkTenItemOptionValid Then
					strParam = ""
					strParam = oGmarket.FOneItem.getGmarketItemOptRegParameter()
					Call fnGmarketOPTReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FLimityn, oGmarket.FOneItem.FLimitno, oGmarket.FOneItem.FLimitsold)
				Else
					iErrStr = "ERR||"&itemid&"||[AddOPT] �ɼǰ˻� ����"
				End If
			End If
		SET oGmarket = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'#################################### �ɼ� ���� ��� �� ####################################
	If failCnt > 0 Then
		strParam = ""
		strParam = getGmarketAddPriceParameter(itemid, tGmarketGoodno, "N", mustPrice, displayDate)
		Call fnGmarketItemAddPrice(itemid, strParam, mustPrice, displayDate, "N", iErrStr)

		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
			failCnt2 = failCnt2 + 1
		Else
			SumOKStr = SumOKStr & iErrStr
		End If

		If failCnt2 = 0 Then
			SumErrStr = SumErrStr & "[ǰ��ó��]"
		End If
	End If

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
		Call SugiQueLogInsert("gmarket1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("gmarket1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "REGPrice" Then
	tGmarketGoodno = getGmarketGoodno(itemid)
	If tGmarketGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
	Else
		strParam = ""
		strParam = getGmarketAddPriceParameter(itemid, tGmarketGoodno, "N", mustPrice, displayDate)
		Call fnGmarketItemAddPrice(itemid, strParam, mustPrice, displayDate, "N", iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("gmarket1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "EditSellYn" Then
	isAllRegYn = getAllRegChk2(itemid, tGmarketGoodno, tOptionCnt, tLimityn, chgSellYn)
	If tGmarketGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||�⺻����, �ɼ�����, ��ǰ��� �Է��� Ȯ���ϼ���"
	Else
		strParam = ""
		strParam = getGmarketAddPriceParameter(itemid, tGmarketGoodno, chgSellYn, mustPrice, displayDate)
		Call fnGmarketItemAddPrice(itemid, strParam, mustPrice, displayDate, chgSellYn, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("gmarket1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "EditInfo" Then
	SET oGmarket = new CGmarket
		oGmarket.FRectItemID	= itemid
		oGmarket.getGmarketEditOneItem

	    If (oGmarket.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		ElseIf oGmarket.FOneItem.checkItemContent = "Y" Then
			iErrStr = "ERR||"&itemid&"||iframe�� ���� ��ǰ�� ���� �� �� �����ϴ�."
		Else
			strParam = ""
			strParam = oGmarket.FOneItem.getGmarketItemRegParameter(TRUE)
			Call fnGmarketIteminfoEdit(itemid, oGmarket.FOneItem.FGmarketGoodNo, oGmarket.FOneItem.FItemName, iErrStr, strParam)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("gmarket1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oGmarket = nothing
ElseIf action = "EDITRETURNFEE" Then
	tGmarketGoodno = getGmarketGoodno(itemid)
	If tGmarketGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
	Else
		strParam = ""
		strParam = getGmarketReturnFeeParameter(itemid, tGmarketGoodno, CRETURNFEE)
		Call fnGmarketReturnFee(itemid, strParam, CRETURNFEE, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("gmarket1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDITPOLICY" Then
	SET oGmarket = new CGmarket
		oGmarket.FRectItemID	= itemid
		oGmarket.getGmarketEditOneItem

	    If (oGmarket.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		ElseIf oGmarket.FOneItem.checkItemContent = "Y" Then
			iErrStr = "ERR||"&itemid&"||iframe�� ���� ��ǰ�� ���� �� �� �����ϴ�."
		Else
			strParam = ""
			strParam = oGmarket.FOneItem.getGmarketItemRegParameter(TRUE)
			Call fnGmarketIteminfoEdit(itemid, oGmarket.FOneItem.FGmarketGoodNo, oGmarket.FOneItem.FItemName, iErrStr, strParam)
		End If

		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If

		If failCnt = 0 Then
			strParam = ""
			strParam = getGmarketReturnFeeParameter(itemid, oGmarket.FOneItem.FGmarketGoodNo, CRETURNFEE)
			Call fnGmarketReturnFee(itemid, strParam, CRETURNFEE, iErrStr)

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
			CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
			Call SugiQueLogInsert("gmarket1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("gmarket1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oGmarket = nothing
ElseIf action = "PRICE" Then
	SET oGmarket = new CGmarket
		oGmarket.FRectItemID	= itemid
		oGmarket.getGmarketEditPriceOptOneItem
		If oGmarket.FResultCount > 0 Then
			'�ɼ��߰��ݾ��� ��ǰ�ݾ��� 50%�ʰ� �˻�
			isFiftyUpDown = oGmarket.FOneItem.getFiftyUpDown

			getMustprice = ""
			getMustprice = oGmarket.FOneItem.MustPrice()
			'���� ǰ���� �ش��ϰų� 50%�ʰ��ϰų� 0���ɼ��� ��� ǰ���� ��..(������ǰ��� ��� 5�����ϵ� ������)
			If (oGmarket.FOneItem.FmaySoldOut = "Y") OR (isFiftyUpDown = "Y") OR (oGmarket.FOneItem.FLimityn = "Y" AND (oGmarket.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) OR (oGmarket.FOneItem.IsMayLimitSoldout = "Y") Then
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketAddPriceParameter("N", getMustprice, displayDate)
				Call fnGmarketItemAddPrice(itemid, strParam, getMustprice, displayDate, "N", iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
				SET oGmarket = nothing
			Else
			'�� ���ǿ� �ش����� ������ ������ �Ǹ�ó��
				iErrStr = ""
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketAddPriceParameter("Y", getMustprice, displayDate)
				Call fnGmarketItemAddPrice(itemid, strParam, getMustprice, displayDate, "Y", iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
				SET oGmarket = nothing

				SET oGmarket = new CGmarket
					oGmarket.FRectItemID	= itemid
					oGmarket.getGmarketNotOptOneItem
				    If (oGmarket.FResultCount < 1) Then
						iErrStr = "ERR||"&itemid&"||�ɼ� ��� ������ ��ǰ�� �ƴմϴ�."
					ElseIf (oGmarket.FOneItem.FGmarketGoodNo = "") Then
						iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
					ElseIf (oGmarket.FOneItem.FAPIadditem = "N") Then
						iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
					ElseIf (oGmarket.FOneItem.getFiftyUpDown = "Y") Then
						iErrStr = "ERR||"&itemid&"||�ɼǰ����� 50%�� �ʰ��մϴ�."
					Else
						'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
						If oGmarket.FOneItem.checkTenItemOptionValid Then
							strParam = ""
							strParam = oGmarket.FOneItem.getGmarketItemOptRegParameter()
							Call fnGmarketOPTReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FLimityn, oGmarket.FOneItem.FLimitno, oGmarket.FOneItem.FLimitsold)
						Else
							iErrStr = "ERR||"&itemid&"||[AddOPT] �ɼǰ˻� ����"
						End If
					End If
				SET oGmarket = nothing
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		Else
			iErrStr = "ERR||"&itemid&"||������ ������ ����"
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
			CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
			Call SugiQueLogInsert("gmarket1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("gmarket1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
ElseIf action = "EDIT" Then
	SET oGmarket = new CGmarket
		oGmarket.FRectItemID	= itemid
		oGmarket.getGmarketEditOneItem
		'#################################### �⺻ ���� ���� ���� ####################################
	    If (oGmarket.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		ElseIf oGmarket.FOneItem.checkItemContent = "Y" Then
			iErrStr = "ERR||"&itemid&"||iframe�� ���� ��ǰ�� ���� �� �� �����ϴ�."
			isiframe = "Y"
		Else
			strParam = ""
			strParam = oGmarket.FOneItem.getGmarketItemRegParameter(TRUE)
			Call fnGmarketIteminfoEdit(itemid, oGmarket.FOneItem.FGmarketGoodNo, oGmarket.FOneItem.FItemName, iErrStr, strParam)
		End If
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If

		'#################################### ��ǰ�� ���� ���� ���� ####################################
		If (oGmarket.FResultCount > 0) AND (oGmarket.FOneItem.FReturnShippingFee < 100) Then
			strParam = ""
			strParam = getGmarketReturnFeeParameter(itemid, oGmarket.FOneItem.FGmarketGoodNo, CRETURNFEE)
			Call fnGmarketReturnFee(itemid, strParam, CRETURNFEE, iErrStr)

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If

		'#################################### �̹��� ���� ���� ####################################
		If (oGmarket.FResultCount > 0) AND (oGmarket.FOneItem.isImageChanged) Then
			strParam = ""
			strParam = oGmarket.FOneItem.getGmarketItemEditImgParameter()
			Call fnGmarketEditImg(itemid, strParam, iErrStr, oGmarket.FOneItem.FbasicimageNm)
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
	SET oGmarket = nothing

	SET oGmarket = new CGmarket
		oGmarket.FRectItemID	= itemid
		oGmarket.getGmarketEditPriceOptOneItem
		'#################################### ��ǰ ���� ���� ���� ####################################
		If oGmarket.FResultCount > 0 Then
			'�ɼ��߰��ݾ��� ��ǰ�ݾ��� 50%�ʰ� �˻�
			isFiftyUpDown = oGmarket.FOneItem.getFiftyUpDown

			getMustprice = ""
			getMustprice = oGmarket.FOneItem.MustPrice()
			'���� ǰ���� �ش��ϰų� ������������ �ְų� 50%�ʰ��ϰų� 0���ɼ��� ��� ǰ���� ��..(������ǰ��� ��� 5�����ϵ� ������)
			If (oGmarket.FOneItem.FmaySoldOut = "Y") OR (isFiftyUpDown = "Y") OR (isiframe = "Y") OR (oGmarket.FOneItem.FLimityn = "Y" AND (oGmarket.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) OR (oGmarket.FOneItem.IsMayLimitSoldout = "Y") Then
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketAddPriceParameter("N", getMustprice, displayDate)
				Call fnGmarketItemAddPrice(itemid, strParam, getMustprice, displayDate, "N", iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					failCnt = 0
					SumOKStr = SumOKStr & iErrStr
				End If
				SET oGmarket = nothing
			Else
			'�� ���ǿ� �ش����� ������ ������ �Ǹ�ó��
				iErrStr = ""
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketAddPriceParameter("Y", getMustprice, displayDate)
				Call fnGmarketItemAddPrice(itemid, strParam, getMustprice, displayDate, "Y", iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
				SET oGmarket = nothing
		'#################################### ��ǰ �ɼ� ���� ���� ####################################
				SET oGmarket = new CGmarket
					oGmarket.FRectItemID	= itemid
					oGmarket.getGmarketNotOptOneItem
				    If (oGmarket.FResultCount < 1) Then
						iErrStr = "ERR||"&itemid&"||�ɼ� ��� ������ ��ǰ�� �ƴմϴ�."
					ElseIf (oGmarket.FOneItem.FGmarketGoodNo = "") Then
						iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
					ElseIf (oGmarket.FOneItem.FAPIadditem = "N") Then
						iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
					ElseIf (oGmarket.FOneItem.getFiftyUpDown = "Y") Then
						iErrStr = "ERR||"&itemid&"||�ɼǰ����� 50%�� �ʰ��մϴ�."
					Else
						'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
						If oGmarket.FOneItem.checkTenItemOptionValid Then
							strParam = ""
							strParam = oGmarket.FOneItem.getGmarketItemOptRegParameter()
							Call fnGmarketOPTReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FLimityn, oGmarket.FOneItem.FLimitno, oGmarket.FOneItem.FLimitsold)
						Else
							iErrStr = "ERR||"&itemid&"||[AddOPT] �ɼǰ˻� ����"
						End If
					End If
				SET oGmarket = nothing
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		Else
			iErrStr = "ERR||"&itemid&"||������ ������ ����"
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
			CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
			Call SugiQueLogInsert("gmarket1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("gmarket1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
ElseIf action = "EditImg" Then
	SET oGmarket = new CGmarket
		oGmarket.FRectItemID	= itemid
		oGmarket.getGmarketEditImageOneItem
	    If (oGmarket.FOneItem.FGmarketGoodNo = "") Then
			iErrStr = "ERR||"&itemid&"||�̹��� ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oGmarket.FOneItem.getGmarketItemEditImgParameter()
			Call fnGmarketEditImg(itemid, strParam, iErrStr, oGmarket.FOneItem.FbasicimageNm)
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("gmarket1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oGmarket = nothing
ElseIf action = "EditCert" Then
	tGmarketGoodno = getGmarketGoodno(itemid)
	If tGmarketGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
	Else
		Call getGmarketChildrenCate(itemid, isChild, isLife, isElec)
		If isChild = "Y" OR isLife = "Y" OR isElec = "Y" Then
			strParam = ""
			strParam = getGmarketChildrenParameter(itemid, tGmarketGoodno, isChild, isLife, isElec)
			Call fnGmarketItemChildren(itemid, strParam, iErrStr)
		Else
			iErrStr = "ERR||"&itemid&"||������ �ʿ���� ��ǰ�Դϴ�."
		End If
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("gmarket1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "REGG9" Then
	SET oGmarket = new CGmarket
		oGmarket.FRectItemID	= itemid
		oGmarket.getG9NotRegOneItem
		If (oGmarket.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||G9�� ��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oGmarket.FOneItem.getG9ItemRegParameter()
			Call fnG9ItemReg(itemid, strParam, iErrStr)
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("gmarket1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oGmarket = nothing
ElseIf action = "AddMakerBrand" Then
	strParam = ""
	strParam = getGmarketAddMakerBrandParameter(gMakername, gBrandname)
	Call fnGmarketAddMaker(strParam)
ElseIf action = "AddAddressBook" Then
	strParam = ""
	strParam = getGmarketAddAddressBookParameter()
	Call fnGmarketAddAddressBook(strParam)
ElseIf action = "RequestAddressBook" Then
	strParam = ""
	strParam = getGmarketRequestAddressBookParameter()
	Call fnGmarketRequestAddressBook(strParam)
ElseIf action = "CATE" Then
	Call fnGmarketCateGet()
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
'###################################################### Gmarket API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
