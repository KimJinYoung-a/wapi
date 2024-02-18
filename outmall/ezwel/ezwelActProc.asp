<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/aspJSON1.17.asp"-->
<!-- #include virtual="/outmall/ezwel/ezwelItemcls.asp"-->
<!-- #include virtual="/outmall/ezwel/incezwelFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, oEzwel, failCnt, chgSellYn, arrRows, skipItem, sellgubun, getMustprice, chkXML, ezwelGoodno, isItemIdChk
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, chkparam, optReset, optString
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
chkXML			= request("chkXML")
failCnt			= 0

Select Case action
	Case "mafcList", "brandList"	isItemIdChk = "N"
	Case Else				isItemIdChk = "Y"
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
'######################################################## Ezwel API ########################################################
If action = "EditSellYn" Then								'���º���
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem

		If chgSellYn = "N" Then
			sellgubun = "SellN"
		Else
			chgSellYn = "AdminOK"
			sellgubun = "SellY"
		End If

		strParam = ""
		strParam = oEzwel.FOneItem.getEzwelItemRegXML(sellgubun, chkXML)
		getMustprice = ""
		getMustprice = oEzwel.FOneItem.fngetMustPrice()
		Call EzwelOneItemEditSellyn(itemid, oEzwel.FOneItem.FEzwelGoodNo, iErrStr, strParam, getMustprice, chgSellYn, "all", oEzwel.FOneItem.FLimityn, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitsold, chkXML)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oEzwel = nothing
ElseIf action = "EDIT" Then									'��ǰ ����
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���������� ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			chkparam = ""
			iErrStr = ""
			optReset = "N"
			optString = "all"
			'*********************************************************************************************************************************************************
			'2014-11-06 ������ | dev_Comment
			'API�� ���۵Ǵ� ���� ��ǰ�ɼ��� �ν����� ���� | ��ϵ� �ɼ�ī��Ʈ�� ũ�ٸ� 10x10���� �ɼ� ������ ���� �������
			'�ᱹ �������� �ɼǻ��������� ������ �ɼ��� �ʱ�ȭ ���� �߰�
			'�߰� : �ι� API���۽� ���� Ȯ���� ������ �� | �Ƹ� ��������� DB�� ��ǰ���� �����ϴ� �� ���� �ɷ��ִ� �� ��..
			'		���� �켱 �̷� ��ǰ�� ǰ����
			strSql = ""
			strSql = strSql &  "SELECT top 1 r.itemid, i.optioncnt, r.regedoptcnt "
			strSql = strSql & " FROM db_item.dbo.tbl_item as i "
			strSql = strSql & " join db_etcmall.dbo.tbl_ezwel_regitem as r on i.itemid=r.itemid "
			strSql = strSql & " WHERE i.itemid=" & itemid
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If not rsget.EOF Then
				If CInt(rsget("optioncnt")) > 0 Then
					If CInt(rsget("optioncnt")) <> CInt(rsget("regedoptcnt")) Then
						optReset = "Y"
						optString = "optMustN"
					End If
				End If
			End If
			rsget.Close

			If (oEzwel.FOneItem.FmaySoldOut = "Y") OR (oEzwel.FOneItem.IsSoldOutLimit5Sell) OR (optReset = "Y") OR (oEzwel.FOneItem.IsMayLimitSoldout = "Y") Then
				If optReset = "Y" Then
					strParam = oEzwel.FOneItem.getEzwelItemRegXML("MustNotOpt", chkXML)
				Else
					strParam = oEzwel.FOneItem.getEzwelItemRegXML("SellN", chkXML)
				End If
				chgSellYn = "N"
			Else
				strParam = oEzwel.FOneItem.getEzwelItemRegXML("SellY", chkXML)
				chgSellYn = "Y"
			End If

			getMustprice = ""
			getMustprice = oEzwel.FOneItem.fngetMustPrice()
			Call EzwelOneItemEdit(itemid, oEzwel.FOneItem.FEzwelGoodNo, iErrStr, strParam, getMustprice, chgSellYn, optString, oEzwel.FOneItem.FLimityn, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitsold, chkXML, oEzwel.FOneItem.FezwelSellYn)
		End If

		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If

		If InStr(iErrStr, "[���Ǹ�]") = 0 Then
			Call EzwelItemChkstat(itemid, iErrStr, oEzwel.FOneItem.FEzwelGoodNo)
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
			CALL Fn_AcctFailTouch("ezwel", itemid, SumErrStr)
			Call SugiQueLogInsert("ezwel", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("ezwel", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oEzwel = nothing
ElseIf action = "REG" Then									'��ǰ ���
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelNotRegOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oEzwel.FOneItem.FdepthCode = "0" Then
			iErrStr = "ERR||"&itemid&"||ī�װ� ��Ī ���� Ȯ���ϼ���."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_ezwel_regItem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_ezwel_regItem "
	        strSql = strSql & " 	(itemid, regdate, reguserid, ezwelstatCD, regitemname)"
	        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oEzwel.FOneItem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oEzwel.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oEzwel.FOneItem.getEzwelItemRegXML("Reg", chkXML)
				Call EzwelItemReg(itemid, strParam, iErrStr, oEzwel.FOneItem.FSellCash, oEzwel.FOneItem.getEzwelSellYn, oEzwel.FOneItem.FLimityn, oEzwel.FOneItem.FLimitNo, oEzwel.FOneItem.FLimitSold, html2db(oEzwel.FOneItem.FItemName), oEzwel.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||�ɼǰ˻� ����"
			End If
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oEzwel = nothing
ElseIf action = "CHKSTAT" Then									'���� ��ȸ
	ezwelGoodno = getEzwelGoodno(itemid)
	If (ezwelGoodno = "") Then
		iErrStr = "ERR||"&itemid&"||��ȸ ������ ��ǰ�� �ƴմϴ�."
	Else
		Call EzwelItemChkstat(itemid, iErrStr, ezwelGoodno)
	End If

	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ElseIf action = "REG2" Then										'��ǰ���
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelNotRegOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oEzwel.FOneItem.FdepthCode = "0" Then
			iErrStr = "ERR||"&itemid&"||ī�װ� ��Ī ���� Ȯ���ϼ���."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Outmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"', '"&CMALLNAME&"' "
			dbget.execute strSql
			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oEzwel.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oEzwel.FOneItem.getEzwelItemRegJson("N")
				Call EzwelItemNewReg(itemid, strParam, iErrStr, oEzwel.FOneItem.FSellCash, oEzwel.FOneItem.getEzwelSellYn, oEzwel.FOneItem.FLimityn, oEzwel.FOneItem.FLimitNo, oEzwel.FOneItem.FLimitSold, html2db(oEzwel.FOneItem.FItemName), oEzwel.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||�ɼǰ˻� ����"
			End If
		End If

		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	SET oEzwel = nothing

	If failCnt = 0 Then
		SET oEzwel = new cEzwel
			oEzwel.FRectItemID	= itemid
			oEzwel.getEzwelEditOneItem
			If (oEzwel.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ȸ ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				Call EzwelItemNewChkstat(itemid, oEzwel.FOneItem.FEzwelGoodNo, oEzwel.FOneItem.FLimitYN, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitSold, iErrStr)
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		SET oEzwel = nothing
	End If

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("ezwel", itemid, SumErrStr)
		Call SugiQueLogInsert("ezwel", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("ezwel", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=REG2&itemid=2343355
ElseIf action = "CHKSTAT2" Then									'����ȸ
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ȸ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			Call EzwelItemNewChkstat(itemid, oEzwel.FOneItem.FEzwelGoodNo, oEzwel.FOneItem.FLimitYN, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitSold, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oEzwel = nothing
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=CHKSTAT2&itemid=2930817
ElseIf action = "PRICE" Then									'���ݺ���
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = oEzwel.FOneItem.getEzwelItemPriceJson()
			Call EzwelItemPrice(itemid, strParam, oEzwel.FOneItem.MustPrice, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oEzwel = nothing
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=PRICE&itemid=2930817
ElseIf action = "EditSellYn2" Then								'���º���
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		Else
			Call EzwelNewEditSellyn(itemid, chgSellYn, oEzwel.FOneItem.FEzwelGoodNo, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oEzwel = nothing
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=EditSellYn2&itemid=2930817&chgSellYn=Y
ElseIf action = "EDITOPT" Then									'�ɼǺ���
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem
	    If (oEzwel.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = oEzwel.FOneItem.getEzwelItemOptionJson()
			Call EzwelItemOption(itemid, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("ezwel", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oEzwel = nothing
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=EDITOPT&itemid=2930817
ElseIf action = "EDIT2" Then									'��ǰ����
	SET oEzwel = new cEzwel
		oEzwel.FRectItemID	= itemid
		oEzwel.getEzwelEditOneItem

		If oEzwel.FResultCount = 0 Then
	    	failCnt = failCnt + 1
			iErrStr = "ERR||"&itemid&"||���������� ��ǰ�� �ƴմϴ�."
		Else
			If (oEzwel.FOneItem.FmaySoldOut = "Y") OR (oEzwel.FOneItem.IsMayLimitSoldout = "Y") OR (oEzwel.FOneItem.FLimityn = "Y" AND (oEzwel.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) Then
				Call EzwelNewEditSellyn(itemid, "N", oEzwel.FOneItem.FEzwelGoodNo, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
				rw "�ǸŻ��� ����"
				response.flush
				response.clear
			Else
			'##################################### �ǸŻ�ǰ ��ȸ ���� #######################################
				Call EzwelItemNewChkstat(itemid, oEzwel.FOneItem.FEzwelGoodNo, oEzwel.FOneItem.FLimitYN, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitSold, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
				rw "�ǸŻ�ǰ ��ȸ"
				response.flush
				response.clear
			'##################################### �ǸŻ�ǰ ��ȸ �� #########################################

			'##################################### ���� ���� ���� ###########################################
				If failCnt = 0 Then
					strParam = ""
					strParam = oEzwel.FOneItem.getEzwelItemPriceJson()
					Call EzwelItemPrice(itemid, strParam, oEzwel.FOneItem.MustPrice, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "���� ����"
					response.flush
					response.clear
				End If
			'##################################### ���� ���� �� #############################################

			'##################################### ��ǰ ���� ���� ###########################################
				If failCnt = 0 Then
					strParam = ""
					strParam = oEzwel.FOneItem.getEzwelItemRegJson("EDIT")
					Call EzwelItemNewEdit(itemid, strParam, iErrStr, html2db(oEzwel.FOneItem.FItemName), oEzwel.FOneItem.FbasicimageNm)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "��ǰ ����"
					response.flush
					response.clear
				End If
			'####################################### ��ǰ ���� �� ###########################################

			'##################################### �Ǹ� ���� ���� ���� ######################################
				If failCnt = 0 Then
					Call EzwelNewEditSellyn(itemid, "Y", oEzwel.FOneItem.FEzwelGoodNo, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "�ǸŻ��� ����"
					response.flush
					response.clear
				End If
			'##################################### �Ǹ� ���� ���� �� ########################################

			'##################################### �ǸŻ�ǰ ��ȸ ���� #######################################
				If failCnt = 0 Then
					Call EzwelItemNewChkstat(itemid, oEzwel.FOneItem.FEzwelGoodNo, oEzwel.FOneItem.FLimitYN, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitSold, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "�ǸŻ�ǰ ��ȸ"
					response.flush
					response.clear
				End If
			'##################################### �ǸŻ�ǰ ��ȸ �� #########################################
			End If
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("ezwel", itemid, SumErrStr)
			Call SugiQueLogInsert("ezwel", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("ezwel", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oEzwel = nothing
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=EDIT2&itemid=2648990
ElseIf action = "brandList" Then								'�귣����ȸ
	Call fnEzwelBrandList()
	response.end
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=brandList
ElseIf action = "mafcList" Then									'��������ȸ
	Call fnEzwelMafcList()
	response.end
	'http://localhost:11117/outmall/ezwel/ezwelActProc.asp?act=mafcList
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
'###################################################### ezwel API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
