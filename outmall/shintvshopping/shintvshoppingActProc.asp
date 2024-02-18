<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/shintvshopping/inc_authCheck.asp"-->
<!-- #include virtual="/outmall/shintvshopping/shintvshoppingItemcls.asp"-->
<!-- #include virtual="/outmall/shintvshopping/incShintvshoppingFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, oShintvshopping, failCnt, chgSellYn, arrRows, getMustprice, interfaceId, getShipCostCode
Dim iErrStr, strParam, strSql, SumErrStr, SumOKStr, isItemIdChk, grpVal, rSkip, rLimit, i, salegb
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
interfaceId		= request("interfaceId")
failCnt			= 0

Select Case action
	Case "commonCode", "cateList", "certList", "shipCost", "offerList"
		isItemIdChk = "N"
	Case Else
		isItemIdChk = "Y"
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
'ī�װ� ���Ž� Ȯ���� ��
'1. http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=cateList	'ī�װ��� ����� ���� ������
'2. http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=certList	'������ ī�װ��� �������� ������Ʈ
'######################################################## shintvshopping API ########################################################
'��ǰ ��� PROCESS
'IF_API_10_037 >> IF_API_10_001 / IF_API_10_002 / IF_API_10_003 / IF_API_10_006 /  IF_API_10_027 / >> IF_API_10_011
'�����������		��������			��ǰ���		�̹���URL		�������	  ��������(Optional)	���ο�û
' 	reglevel1		reglevel2		reglevel3		reglevel4		reglevel5		reglevel6			reglevel7
If action = "REG" Then
	'##################################### ����������� ���� #####################################
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingNotRegOneItem
	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Outmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"', '"&CMALLNAME&"' "
			dbget.execute strSql

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oShintvshopping.FOneItem.checkTenItemOptionValid Then
				getShipCostCode = ""
				getShipCostCode = oShintvshopping.FOneItem.fnShipCostCode()
				strParam = ""
				strParam = oShintvshopping.FOneItem.getshintvshoppingItemRegParameter(getShipCostCode)

				getMustprice = ""
				getMustprice = oShintvshopping.FOneItem.MustPrice()
				Call fnShintvshoppingItemReg(itemid, strParam, iErrStr, getMustprice, oShintvshopping.FOneItem.getShintvshoppingSellYn, oShintvshopping.FOneItem.FLimityn, oShintvshopping.FOneItem.FLimitNo, oShintvshopping.FOneItem.FLimitSold, html2db(oShintvshopping.FOneItem.FItemName), oShintvshopping.FOneItem.FbasicimageNm)
				rw "�����������"
				response.flush
				response.clear
			Else
				iErrStr = "ERR||"&itemid&"||[�ӽõ��] �ɼǰ˻� ����"
			End If
		End If
	SET oShintvshopping = nothing
	If Left(iErrStr, 2) <> "OK" Then
		failCnt = failCnt + 1
		SumErrStr = SumErrStr & iErrStr
	Else
		SumOKStr = SumOKStr & iErrStr
	End If
	'##################################### ����������� �� #######################################

	'##################################### �������� ���� #######################################
	If failCnt = 0 Then
		iErrStr = ""
		SET oShintvshopping = new CShintvshopping
			oShintvshopping.FRectItemID	= itemid
			oShintvshopping.getShintvshoppingTmpRegedOneItem
			If (oShintvshopping.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			ElseIf oShintvshopping.FOneItem.FReglevel <> 1 Then
				iErrStr = "ERR||"&itemid&"||��ǰ���� ����ϼ���."
			Else
				strParam = ""
				strParam = oShintvshopping.FOneItem.getshintvshoppingContentParameter()
				Call fnShintvshoppingContentReg(itemid, strParam, iErrStr)
				rw "��������"
				response.flush
				response.clear
			End If
		SET oShintvshopping = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### �������� �� #######################################

	'##################################### ��ǰ��� ���� #######################################
	If failCnt = 0 Then
		iErrStr = ""
		SET oShintvshopping = new CShintvshopping
			oShintvshopping.FRectItemID	= itemid
			oShintvshopping.getShintvshoppingTmpRegedOneItem
			If (oShintvshopping.FResultCount < 1) Then
				failCnt = failCnt + 1
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			ElseIf oShintvshopping.FOneItem.FReglevel <> 2 Then
				failCnt = failCnt + 1
				iErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oShintvshopping.FOneItem.FReglevel &") "
			Else
				arrRows = getOptionList(itemid)
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						strParam = ""
						strParam = oShintvshopping.FOneItem.getshintvshoppingOptParameter(arrRows(0, i), arrRows(1, i))
						Call fnShintvshoppingOptReg(itemid, strParam, iErrStr)
						If iErrStr <> "" Then
							SumErrStr = SumErrStr & arrRows(2, i) & ","
						End If
					Next
					iErrStr = ArrErrStrInfo("REGOpt", itemid, SumErrStr)
					rw "��ǰ���"
					response.flush
					response.clear
				End If
			End If
		SET oShintvshopping = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### ��ǰ��� �� #########################################

	'##################################### �̹���URL ���� ######################################
	If failCnt = 0 Then
		iErrStr = ""
		SET oShintvshopping = new CShintvshopping
			oShintvshopping.FRectItemID	= itemid
			oShintvshopping.getShintvshoppingTmpRegedOneItem
			If (oShintvshopping.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			ElseIf oShintvshopping.FOneItem.FReglevel <> 3 Then
				iErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oShintvshopping.FOneItem.FReglevel &") "
			Else
				strParam = ""
				strParam = oShintvshopping.FOneItem.getshintvshoppingImageParameter()
				Call fnShintvshoppingImageReg(itemid, strParam, iErrStr)
				rw "�̹���URL ���"
				response.flush
				response.clear
			End If
		SET oShintvshopping = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### �̹���URL �� ########################################

	'##################################### ������� ���� #######################################
	If failCnt = 0 Then
		iErrStr = ""
		SET oShintvshopping = new CShintvshopping
			oShintvshopping.FRectItemID	= itemid
			oShintvshopping.getShintvshoppingTmpRegedOneItem
			If (oShintvshopping.FResultCount < 1) Then
				failCnt = failCnt + 1
				SumErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			ElseIf oShintvshopping.FOneItem.FReglevel <> 4 Then
				failCnt = failCnt + 1
				SumErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oShintvshopping.FOneItem.FReglevel &") "
			Else
				arrRows = getInfoCodeMapList(itemid)
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						strParam = ""
						strParam = oShintvshopping.FOneItem.getshintvshoppingGosiRegParameter(arrRows(0, i), arrRows(1, i), arrRows(2, i))
						Call fnShintvshoppingGosiReg(itemid, strParam, iErrStr)
						If iErrStr <> "" Then
							SumErrStr = SumErrStr & arrRows(0, i) & ","
						End If
					Next
					iErrStr = ArrErrStrInfo("REGGosi", itemid, SumErrStr)
					rw "������� ���"
					response.flush
					response.clear
				End If
			End If
		SET oShintvshopping = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### ������� �� ##########################################

	'##################################### ����������� ���� ####################################
	If (failCnt = 0) AND (getMayCertYn(itemid)) = "Y" Then
		iErrStr = ""
		SET oShintvshopping = new CShintvshopping
			oShintvshopping.FRectItemID	= itemid
			oShintvshopping.getShintvshoppingTmpRegedOneItem
			If (oShintvshopping.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				strParam = oShintvshopping.FOneItem.getshintvshoppingCertParameter()
				Call fnShintvshoppingCert(itemid, strParam, iErrStr)
				rw "�������� ���"
				response.flush
				response.clear
			End If
		SET oShintvshopping = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### ����������� ���� #####################################

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("shintvshopping", itemid, SumErrStr)
		Call SugiQueLogInsert("shintvshopping", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_shintvshopping_regitem SET " & VBCRLF
		strSql = strSql & " accFailcnt = 0  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("shintvshopping", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=REG&itemid=3322047
ElseIf action = "CONFIRM" Then
	'##################################### ���ο�û ���� #####################################
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingTmpRegedOneItem
	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oShintvshopping.FOneItem.FReglevel <> 5 AND oShintvshopping.FOneItem.FReglevel <> 6 Then
			iErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oShintvshopping.FOneItem.FReglevel &") "
		Else
			strParam = ""
			strParam = oShintvshopping.FOneItem.getshintvshoppingConfirmParameter()
			Call fnShintvshoppingConfirm(itemid, strParam, iErrStr)
			rw "���ο�û"
			response.flush
			response.clear
		End If
	SET oShintvshopping = nothing
	If Left(iErrStr, 2) <> "OK" Then
		failCnt = failCnt + 1
		SumErrStr = SumErrStr & iErrStr
	Else
		SumOKStr = SumOKStr & iErrStr
	End If
	'##################################### ���ο�û �� #######################################

	'##################################### �ǸŻ�ǰ ��ȸ(��)_v2 ���� ########################
	If failCnt = 0 Then
		SET oShintvshopping = new CShintvshopping
			oShintvshopping.FRectItemID	= itemid
			oShintvshopping.getShintvshoppingEditOneItem

			If (oShintvshopping.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ȸ ������ ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				strParam = oShintvshopping.FOneItem.getShintvshoppingItemViewParameter()
				Call fnShintvshoppingItemView(itemid, strParam, iErrStr)
				rw "�ǸŻ�ǰ ��ȸ"
				response.flush
				response.clear
			End If
		SET oShintvshopping = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### �ǸŻ�ǰ ��ȸ(��)_v2 �� ##########################

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("shintvshopping", itemid, SumErrStr)
		Call SugiQueLogInsert("shintvshopping", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_shintvshopping_regitem SET " & VBCRLF
		strSql = strSql & " accFailcnt = 0  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("shintvshopping", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "PRICE" Then							'IF_API_10_029 / ���»� ���ݵ��
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingEditOneItem

	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			getMustprice = ""
			getMustprice = oShintvshopping.FOneItem.MustPrice()
			strParam = ""
			strParam = oShintvshopping.FOneItem.getshintvshoppingEditPriceParameter()
			Call fnShintvshoppingEditPrice(itemid, strParam, getMustprice, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("shintvshopping", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=EDITContent&itemid=3853757
ElseIf action = "EDIT" Then
'��ǰ ���� PROCESS
' ��ȸ > �Ǹſ��� N�̸� ����
' ��ȸ > ���ݼ��� > �⺻�������� > ��������� > �̹������濩�� �˻��� ���� > �ǸŻ��¼��� > ������ > �ɼ��߰�  > ��ȸ
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingEditOneItem
		If oShintvshopping.FResultCount = 0 Then
			iErrStr = "ERR||"&itemid&"||���� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
		Else
			'checkTenItemOptionValid2 => �ɼ� ��뿩�ο� �ɼ��߰��ݾ� üũ
			If (oShintvshopping.FOneItem.FmaySoldOut = "Y") OR (oShintvshopping.FOneItem.IsMayLimitSoldout = "Y") OR (oShintvshopping.FOneItem.checkTenItemOptionValid2 <> "True") Then
				strParam = ""
				strParam = oShintvshopping.FOneItem.getShintvshoppingSellynParameter("N")
				Call fnShintvshoppingSellyn(itemid, "N", strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
	'##################################### �ǸŻ�ǰ ��ȸ ���� #######################################
				strParam = ""
				strParam = oShintvshopping.FOneItem.getShintvshoppingItemViewParameter()
				Call fnShintvshoppingItemView(itemid, strParam, iErrStr)
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

	'##################################### ���»� ���ݵ�� ���� #####################################
				If failCnt = 0 Then
					iErrStr = ""
					getMustprice = ""
					getMustprice = oShintvshopping.FOneItem.MustPrice()
					strParam = ""
					strParam = oShintvshopping.FOneItem.getshintvshoppingEditPriceParameter()
					Call fnShintvshoppingEditPrice(itemid, strParam, getMustprice, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "�ǸŻ�ǰ ���ݼ���"
					response.flush
					response.clear
				End If
	'##################################### ���»� ���ݵ�� �� #######################################

	'##################################### ���»� �������� ����_v2 ���� #############################
				If failCnt = 0 Then
					iErrStr = ""
					strParam = ""
					getShipCostCode = ""
					getShipCostCode = oShintvshopping.FOneItem.fnShipCostCode()
					strParam = oShintvshopping.FOneItem.getshintvshoppingItemEditParameter(getShipCostCode)
					Call fnShintvshoppingItemEdit(itemid, strParam, getShipCostCode, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "�ǸŻ�ǰ �������� ����"
					response.flush
					response.clear
				End If
	'##################################### ���»� �������� ����_v2 �� ###############################

	'##################################### ���»� ����� ���� ���� ##################################
				If failCnt = 0 Then
					iErrStr = ""
					strParam = ""
					strParam = oShintvshopping.FOneItem.getshintvshoppingEditContentParameter()
					Call fnShintvshoppingEditContentReg(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "�ǸŻ�ǰ ����� ����"
					response.flush
					response.clear
				End If
	'##################################### ���»� ����� ���� �� ####################################

	'##################################### �̹��� ���� ���� #########################################
				If failCnt = 0 Then
					If oShintvshopping.FOneItem.isImageChanged = True Then
						iErrStr = ""
						strParam = ""
						strParam = oShintvshopping.FOneItem.getshintvshoppingEditImageParameter()
						Call fnShintvshoppingEditImage(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
						rw "�̹��� ����"
						response.flush
						response.clear
					End If
				End If
	'##################################### �̹��� ���� �� ###########################################

	'##################################### �ɼ� �ǸŻ��� ���� ���� ###################################
				If failCnt = 0 Then
					iErrStr = ""
					arrRows = ""
					arrRows = getOptiopnMapList(itemid)
					If isArray(arrRows) Then
						For i = 0 To UBound(arrRows,2)
							If (arrRows(7, i) < 1) OR (arrRows(11, i)= "1") OR (arrRows(9, i)= "N") OR (arrRows(10, i) = "N") Then	'��� 1�� ���ϰų� �ɼǸ��� �ٸ��ų� �ɼǻ�뿩��N �̰ų� �ɼ��Ǹſ��� N�̰ų�
								salegb = "11"
							Else
								salegb = "00"
							End If

							strParam = ""
							strParam = oShintvshopping.FOneItem.geshintvshoppingOptionStatParam(arrRows(2, i), salegb)
							Call fnShintvshoppingOptSellyn(itemid, strParam, iErrStr)
							If iErrStr <> "" Then
								SumErrStr = SumErrStr & arrRows(2, i) & ","
							End If
						Next
						iErrStr = ArrErrStrInfo("EDITSTAT", itemid, SumErrStr)
						rw "�ɼ� �ǸŻ��� ����"
						response.flush
						response.clear
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				End If
	'##################################### �ɼ� �ǸŻ��� ���� �� #####################################

	'##################################### �ɼ� ��� ���� ���� #######################################
				If failCnt = 0 Then
					iErrStr = ""
					arrRows = ""
					arrRows = getOptiopnMapList(itemid)
					If isArray(arrRows) Then
						For i = 0 To UBound(arrRows,2)
							strParam = ""
							strParam = oShintvshopping.FOneItem.geshintvshoppingOptionQtyParam(arrRows(2, i), arrRows(7, i))
							Call fnShintvshoppingQtyEdit(itemid, strParam, iErrStr)
							If iErrStr <> "" Then
								SumErrStr = SumErrStr & arrRows(2, i) & ","
							End If
						Next
						iErrStr = ArrErrStrInfo("EDITQTY", itemid, SumErrStr)
						rw "�ɼ� ��� ����"
						response.flush
						response.clear
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				End If
	'##################################### �ɼ� ��� ���� �� #########################################

	'##################################### �ǸŻ�ǰ �ɼ��߰� ���� #####################################
				If failCnt = 0 Then
					iErrStr = ""
					arrRows = ""
					arrRows = getOptiopnMayAddList(itemid)
					If isArray(arrRows) Then
						For i = 0 To UBound(arrRows,2)
							strParam = ""
							strParam = oShintvshopping.FOneItem.geshintvshoppingOptionAddParam(arrRows(0, i), arrRows(1, i))
							Call fnShintvshoppingOptAdd(itemid, strParam, iErrStr)
							If iErrStr <> "" Then
								SumErrStr = SumErrStr & arrRows(2, i) & ","
							End If
						Next
						iErrStr = ArrErrStrInfo("EDITADDOPT", itemid, SumErrStr)
					Else
						iErrStr = "OK||"&itemid&"||����[�ɼ��߰�x]]"
					End If
					rw "�ɼ� �߰�"
					response.flush
					response.clear

					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
	'##################################### �ǸŻ�ǰ �ɼ��߰� �� #######################################

	'##################################### �ǸŻ�ǰ ��ȸ ���� #######################################
'				If failCnt = 0 Then
					If InStr(SumErrStr, "������ ����ó��") < 0 Then
						strParam = ""
						strParam = oShintvshopping.FOneItem.getShintvshoppingItemViewParameter()
						Call fnShintvshoppingItemView(itemid, strParam, iErrStr)
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
'				End If
	'##################################### �ǸŻ�ǰ ��ȸ �� #########################################
			End If
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			If InStr(SumErrStr, "������ ����ó��") < 0 Then
				CALL Fn_AcctFailTouch("shintvshopping", itemid, SumErrStr)
			End If
			Call SugiQueLogInsert("shintvshopping", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_shintvshopping_regItem SET " & VBCRLF
			strSql = strSql & " accFailcnt = 0  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("shintvshopping", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oShintvshopping = nothing
ElseIf action = "REGAddItem" Then						'IF_API_10_037 / �ӽû�ǰ �������� ���_v2
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingNotRegOneItem
	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Outmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"', '"&CMALLNAME&"' "
			dbget.execute strSql

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oShintvshopping.FOneItem.checkTenItemOptionValid Then
				getShipCostCode = ""
				getShipCostCode = oShintvshopping.FOneItem.fnShipCostCode()
				strParam = ""
				strParam = oShintvshopping.FOneItem.getshintvshoppingItemRegParameter(getShipCostCode)

				getMustprice = ""
				getMustprice = oShintvshopping.FOneItem.MustPrice()
				Call fnShintvshoppingItemReg(itemid, strParam, iErrStr, getMustprice, oShintvshopping.FOneItem.getShintvshoppingSellYn, oShintvshopping.FOneItem.FLimityn, oShintvshopping.FOneItem.FLimitNo, oShintvshopping.FOneItem.FLimitSold, html2db(oShintvshopping.FOneItem.FItemName), oShintvshopping.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[�ӽõ��] �ɼǰ˻� ����"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("shintvshopping", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=REGAddItem&itemid=3003425		--90001159
ElseIf action = "REGContent" Then					'IF_API_10_001 / �ӽû�ǰ ����� ���
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingTmpRegedOneItem
	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oShintvshopping.FOneItem.FReglevel <> 1 Then
			iErrStr = "ERR||"&itemid&"||��ǰ���� ����ϼ���."
		Else
			strParam = ""
			strParam = oShintvshopping.FOneItem.getshintvshoppingContentParameter()
			Call fnShintvshoppingContentReg(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("shintvshopping", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=REGContent&itemid=3003425		--90001159
ElseIf action = "REGOpt" Then						'IF_API_10_002 / �ӽû�ǰ ��ǰ���� ���
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingTmpRegedOneItem
	    If (oShintvshopping.FResultCount < 1) Then
			failCnt = failCnt + 1
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oShintvshopping.FOneItem.FReglevel <> 2 Then
			failCnt = failCnt + 1
			iErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oShintvshopping.FOneItem.FReglevel &") "
		Else
			arrRows = getOptionList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oShintvshopping.FOneItem.getshintvshoppingOptParameter(arrRows(0, i), arrRows(1, i))
					Call fnShintvshoppingOptReg(itemid, strParam, iErrStr)

					If iErrStr <> "" Then
						SumErrStr = SumErrStr & arrRows(2, i) & ","
					End If
				Next
				iErrStr = ArrErrStrInfo(action, itemid, SumErrStr)

				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("shintvshopping", itemid, SumErrStr)
			Call SugiQueLogInsert("shintvshopping", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("shintvshopping", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=REGOpt&itemid=3003425		--90001159
ElseIf action = "REGImg" Then						'IF_API_10_003 / �ӽû�ǰ �̹��� ���(URL)
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingTmpRegedOneItem
	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oShintvshopping.FOneItem.FReglevel <> 3 Then
			iErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oShintvshopping.FOneItem.FReglevel &") "
		Else
			strParam = ""
			strParam = oShintvshopping.FOneItem.getshintvshoppingImageParameter()
			Call fnShintvshoppingImageReg(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("shintvshopping", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=REGImg&itemid=3003425		--90001159
ElseIf action = "REGGosi" Then						'IF_API_10_006 / �ӽû�ǰ ����������� ���
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingTmpRegedOneItem
	    If (oShintvshopping.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oShintvshopping.FOneItem.FReglevel <> 4 Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oShintvshopping.FOneItem.FReglevel &") "
		Else
			arrRows = getInfoCodeMapList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oShintvshopping.FOneItem.getshintvshoppingGosiRegParameter(arrRows(0, i), arrRows(1, i), arrRows(2, i))
					Call fnShintvshoppingGosiReg(itemid, strParam, iErrStr)
					If iErrStr <> "" Then
						SumErrStr = SumErrStr & arrRows(0, i) & ","
					End If
				Next
				iErrStr = ArrErrStrInfo(action, itemid, SumErrStr)

				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("shintvshopping", itemid, SumErrStr)
			Call SugiQueLogInsert("shintvshopping", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("shintvshopping", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=REGGosi&itemid=3003425		--90001159
ElseIf action = "REGCert" Then						'IF_API_10_027 / �ӽû�ǰ �����������
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingTmpRegedOneItem
	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oShintvshopping.FOneItem.getshintvshoppingCertParameter() 
			Call fnShintvshoppingCert(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("shintvshopping", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=REGCert&itemid=3003425		--90001159
ElseIf action = "REGConfirm" Then					'IF_API_10_011 / �ӽû�ǰ ���ο�û
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingTmpRegedOneItem
	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oShintvshopping.FOneItem.FReglevel <> 5 AND oShintvshopping.FOneItem.FReglevel <> 6 Then
			iErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oShintvshopping.FOneItem.FReglevel &") "
		Else
			strParam = ""
			strParam = oShintvshopping.FOneItem.getshintvshoppingConfirmParameter()
			Call fnShintvshoppingConfirm(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("shintvshopping", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=REGConfirm&itemid=3003425		--90001159
ElseIf action = "REGCHKSTAT" Then					'IF_API_10_039 / �ӽû�ǰ ��ȸ(��)_v2
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingTmpRegedOneItem
	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ȸ ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oShintvshopping.FOneItem.getshintvshoppingConfirmParameter()
			Call fnShintvshoppingRegChkstat(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("shintvshopping", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=REGCHKSTAT&itemid=3003425		--90001159
ElseIf action = "EditSellYn" Then					'IF_API_10_023 / ��ǰ �Ǹ��ߴ� ó��
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingEditOneItem

	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oShintvshopping.FOneItem.getShintvshoppingSellynParameter(chgSellYn)
			Call fnShintvshoppingSellyn(itemid, chgSellYn, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("shintvshopping", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=EditSellYn&itemid=3322050&chgsellyn=N
ElseIf action = "CHKSTAT" Then						'IF_API_10_034 / �ǸŻ�ǰ ��ȸ(��)_v2
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingEditOneItem

	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ȸ ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oShintvshopping.FOneItem.getShintvshoppingItemViewParameter()
			Call fnShintvshoppingItemView(itemid, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("shintvshopping", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=CHKSTAT&itemid=3322050
ElseIf action = "EDITINFO" Then						'�ǸŻ�ǰ �������� ����_v2
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingEditOneItem

	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||�������� ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			getShipCostCode = ""
			getShipCostCode = oShintvshopping.FOneItem.fnShipCostCode()
			strParam = ""
			strParam = oShintvshopping.FOneItem.getshintvshoppingItemEditParameter(getShipCostCode)
			Call fnShintvshoppingItemEdit(itemid, strParam, getShipCostCode, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("shintvshopping", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=CHKSTAT&itemid=3322050
ElseIf action = "EDITContent" Then					'IF_API_10_019 / �ǸŻ�ǰ ����� ���
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingEditOneItem

	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||����� ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oShintvshopping.FOneItem.getshintvshoppingEditContentParameter()
			Call fnShintvshoppingEditContentReg(itemid, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("shintvshopping", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=EDITContent&itemid=3853757
ElseIf action = "EDITImage" Then					'IF_API_10_021 / �ǸŻ�ǰ �̹��� ���(URL)
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingEditOneItem

	    If (oShintvshopping.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||�̹��� ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oShintvshopping.FOneItem.getshintvshoppingEditImageParameter()
			Call fnShintvshoppingEditImage(itemid, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("shintvshopping", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oShintvshopping = nothing
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=EDITImage&itemid=3853757
ElseIf action = "EDITQTY" Then						'IF_API_10_018 / �ǸŻ�ǰ �����
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingEditOneItem

	    If (oShintvshopping.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||��� ���� ������ ��ǰ�� �ƴմϴ�."
		ElseIf getShintvshoppingOptCnt(itemid) = 0 Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||��ȸ���� �����ϼ���."
		Else
			arrRows = getOptiopnMapList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oShintvshopping.FOneItem.geshintvshoppingOptionQtyParam(arrRows(2, i), arrRows(7, i))
					Call fnShintvshoppingQtyEdit(itemid, strParam, iErrStr)
					If iErrStr <> "" Then
						SumErrStr = SumErrStr & arrRows(2, i) & ","
					End If
				Next
				iErrStr = ArrErrStrInfo(action, itemid, SumErrStr)

				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("shintvshopping", itemid, SumErrStr)
			Call SugiQueLogInsert("shintvshopping", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("shintvshopping", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
 	SET oShintvshopping = nothing
 	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=EDITGosi&itemid=3853757
ElseIf action = "EDITSTAT" Then					'IF_API_10_023 / ��ǰ �Ǹ��ߴ� ó��
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingEditOneItem

	    If (oShintvshopping.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		ElseIf getShintvshoppingOptCnt(itemid) = 0 Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||��ȸ���� �����ϼ���."
		Else
			arrRows = getOptiopnMapList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					If (arrRows(7, i) < 1) OR (arrRows(11, i)= "1") OR (arrRows(9, i)= "N") OR (arrRows(10, i) = "N") Then	'��� 1�� ���ϰų� �ɼǸ��� �ٸ��ų� �ɼǻ�뿩��N �̰ų� �ɼ��Ǹſ��� N�̰ų�
						salegb = "11"
					Else
						salegb = "00"
					End If

					strParam = ""
					strParam = oShintvshopping.FOneItem.geshintvshoppingOptionStatParam(arrRows(2, i), salegb)
					Call fnShintvshoppingOptSellyn(itemid, strParam, iErrStr)
					If iErrStr <> "" Then
						SumErrStr = SumErrStr & arrRows(2, i) & ","
					End If
				Next
				iErrStr = ArrErrStrInfo(action, itemid, SumErrStr)

				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("shintvshopping", itemid, SumErrStr)
			Call SugiQueLogInsert("shintvshopping", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("shintvshopping", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
 	SET oShintvshopping = nothing
 	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=EDITGosi&itemid=3853757
ElseIf action = "EDITADDOPT" Then				'IF_API_10_033 / �ǸŻ�ǰ ��ǰ���� ���_v2
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingEditOneItem

	    If (oShintvshopping.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		ElseIf getShintvshoppingOptCnt(itemid) = 0 Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||��ȸ���� �����ϼ���."
		Else
			arrRows = getOptiopnMayAddList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oShintvshopping.FOneItem.geshintvshoppingOptionAddParam(arrRows(0, i), arrRows(1, i))
					Call fnShintvshoppingOptAdd(itemid, strParam, iErrStr)
					If iErrStr <> "" Then
						SumErrStr = SumErrStr & arrRows(2, i) & ","
					End If
				Next
				iErrStr = ArrErrStrInfo(action, itemid, SumErrStr)
			Else
				iErrStr = "OK||"&itemid&"||����[�ɼ��߰�x]]"
			End If

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
			CALL Fn_AcctFailTouch("shintvshopping", itemid, SumErrStr)
			Call SugiQueLogInsert("shintvshopping", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("shintvshopping", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
 	SET oShintvshopping = nothing
 	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=EDITGosi&itemid=3853757
ElseIf action = "EDITGosi" Then					'IF_API_10_016 / �ǸŻ�ǰ ����������� ���
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingEditOneItem

	    If (oShintvshopping.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||����������� ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			arrRows = getInfoCodeMapList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oShintvshopping.FOneItem.getshintvshoppingGosiEditParameter(arrRows(0, i), arrRows(1, i), arrRows(2, i))
					Call fnShintvshoppingGosiEdit(itemid, strParam, iErrStr)
					If iErrStr <> "" Then
						SumErrStr = SumErrStr & arrRows(0, i) & ","
					End If
				Next
				iErrStr = ArrErrStrInfo(action, itemid, SumErrStr)

				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("shintvshopping", itemid, SumErrStr)
			Call SugiQueLogInsert("shintvshopping", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("shintvshopping", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
 	SET oShintvshopping = nothing
 	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=EDITGosi&itemid=3853757
 ElseIf action = "EDITCert" Then					'IF_API_10_028 / �ǸŻ�ǰ �����������
	SET oShintvshopping = new CShintvshopping
		oShintvshopping.FRectItemID	= itemid
		oShintvshopping.getShintvshoppingEditOneItem

 	    If (oShintvshopping.FResultCount < 1) Then
 			iErrStr = "ERR||"&itemid&"||�������� ���� ������ ��ǰ�� �ƴմϴ�."
 		Else
 			strParam = ""
 			strParam = oShintvshopping.FOneItem.getshintvshoppingEditCertParameter()
 			Call fnShintvshoppingEditCert(itemid, strParam, iErrStr)
 		End If

 		If LEFT(iErrStr, 2) <> "OK" Then
 			CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
 		End If
 		Call SugiQueLogInsert("shintvshopping", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
 	SET oShintvshopping = nothing
' 	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=EDITImage&itemid=3853757
ElseIf action = "commonCode" Then					'IF_API_00_001 ~ 
	Call fnGetCommonCodeList(interfaceId)
	response.end
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=commonCode&interfaceId=IF_API_00_001
ElseIf action = "cateList" Then						'IF_API_00_028 / ��ǰ ���з� ��ȸ
	Call fnGetGoodsTgroupList()
	response.end
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=catelist
ElseIf action = "certList" Then						'IF_API_00_027 / �������� �׸���ȸ
	strSql = ""
	strSql = strSql & " SELECT lgroup+mgroup+sgroup+dgroup+tgroup as lmsdCode "
	strSql = strSql & " FROM db_etcmall.dbo.tbl_shintvshopping_category "
	strSql = strSql & " WHERE safetyCertYn IS NULL "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		arrRows = rsget.getRows()
	End If
	rsget.Close

	For i = 0 To ubound(arrRows,2)
		Call fnGetCertList(arrRows(0, i))
		If (i mod 300) = 0 Then
			rw "ȣ���� �Դϴ� : " & i
			response.flush
			response.Clear
		End If
	next
	rw "ȣ�� �Ϸ� : " & i
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=certList
ElseIf action = "offerList" Then					'IF_API_00_023 / ��ǰ����������� �׸� ��ȸ
	Call fnGetOfferList()
	response.end
ElseIf action = "shipCost" Then						'IF_API_00_030 / �� ��ۺ���å ���
	Call fnInputCustShipCost()
	response.end
	'http://localhost:11117/outmall/shintvshopping/shintvshoppingActProc.asp?act=shipCost
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
'###################################################### shintvshopping API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
