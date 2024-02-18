<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/skstoa/skstoaItemcls.asp"-->
<!-- #include virtual="/outmall/skstoa/incskstoaFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, oSkstoa, failCnt, chgSellYn, arrRows, getMustprice, interfaceId
Dim iErrStr, strParam, strSql, SumErrStr, SumOKStr, isItemIdChk, grpVal, rSkip, rLimit, i, salegb, iskGoodNo
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
interfaceId		= request("interfaceId")
failCnt			= 0
iskGoodNo		= ""

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
'######################################################## skstoa API ########################################################
'��ǰ ��� PROCESS
'IF_API_10_005 >> IF_API_10_001 / IF_API_10_002 / IF_API_10_003 / IF_API_10_006 /  IF_API_10_028 / >> IF_API_10_011
'�����������		��������			��ǰ���		�̹���URL		�������	  KC ���� �߰�(Optional)	���ο�û
' 	reglevel1		reglevel2		reglevel3		reglevel4		reglevel5		reglevel6			reglevel7
If action = "REG" Then
	'##################################### ����������� ���� #####################################
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getskstoaNotRegOneItem
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Outmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"', '"&CMALLNAME&"' "
			dbget.execute strSql

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oSkstoa.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oSkstoa.FOneItem.getskstoaItemRegParameter()

				getMustprice = ""
				getMustprice = oSkstoa.FOneItem.MustPrice()
				Call fnskstoaItemReg(itemid, strParam, iErrStr, getMustprice, oSkstoa.FOneItem.getskstoaSellYn, oSkstoa.FOneItem.FLimityn, oSkstoa.FOneItem.FLimitNo, oSkstoa.FOneItem.FLimitSold, html2db(oSkstoa.FOneItem.FItemName), oSkstoa.FOneItem.FbasicimageNm)
				rw "�����������"
				response.flush
				response.clear
			Else
				iErrStr = "ERR||"&itemid&"||[�ӽõ��] �ɼǰ˻� ����"
			End If
		End If
	SET oSkstoa = nothing
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
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaTmpRegedOneItem("")
			If (oSkstoa.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			ElseIf oSkstoa.FOneItem.FReglevel <> 1 Then
				iErrStr = "ERR||"&itemid&"||��ǰ���� ����ϼ���."
			Else
				strParam = ""
				strParam = oSkstoa.FOneItem.getskstoaContentParameter()
				Call fnSkstoaContentReg(itemid, strParam, iErrStr)
				rw "��������"
				response.flush
				response.clear
			End If
		SET oSkstoa = nothing
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
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaTmpRegedOneItem("")
			If (oSkstoa.FResultCount < 1) Then
				failCnt = failCnt + 1
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			ElseIf oSkstoa.FOneItem.FReglevel <> 2 Then
				failCnt = failCnt + 1
				iErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oSkstoa.FOneItem.FReglevel &") "
			Else
				arrRows = getOptionList(itemid)
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						strParam = ""
						strParam = oSkstoa.FOneItem.getskstoaOptParameter(arrRows(0, i), arrRows(1, i))
						Call fnSkstoaOptReg(itemid, strParam, iErrStr)
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
		SET oSkstoa = nothing
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
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaTmpRegedOneItem("")
			If (oSkstoa.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			ElseIf oSkstoa.FOneItem.FReglevel <> 3 Then
				iErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oSkstoa.FOneItem.FReglevel &") "
			Else
				strParam = ""
				strParam = oSkstoa.FOneItem.getskstoaImageParameter()
				Call fnSkstoaImageReg(itemid, strParam, iErrStr)
				rw "�̹���URL ���"
				response.flush
				response.clear
			End If
		SET oSkstoa = nothing
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
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaTmpRegedOneItem("")
			If (oSkstoa.FResultCount < 1) Then
				failCnt = failCnt + 1
				SumErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			ElseIf oSkstoa.FOneItem.FReglevel <> 4 Then
				failCnt = failCnt + 1
				SumErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oSkstoa.FOneItem.FReglevel &") "
			Else
				arrRows = getInfoCodeMapList(itemid)
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						strParam = ""
						strParam = oSkstoa.FOneItem.getskstoaGosiRegParameter(arrRows(0, i), arrRows(1, i), arrRows(2, i))
						Call fnSkstoaGosiReg(itemid, strParam, iErrStr)
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
		SET oSkstoa = nothing
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
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaTmpRegedOneItem("")
			If (oSkstoa.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				strParam = oSkstoa.FOneItem.getskstoaCertParameter() 
				Call fnSkstoaCert(itemid, strParam, iErrStr)
				rw "�������� ���"
				response.flush
				response.clear
			End If
		SET oSkstoa = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### ����������� ���� #####################################

	'##################################### ���ο�û ���� #########################################
	If (failCnt = 0) Then
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaTmpRegedOneItem("")
			If (oSkstoa.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			ElseIf oSkstoa.FOneItem.FReglevel <> 5 AND oSkstoa.FOneItem.FReglevel <> 6 Then
				iErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oSkstoa.FOneItem.FReglevel &") "
			Else
				strParam = ""
				strParam = oSkstoa.FOneItem.getskstoaConfirmParameter()
				Call fnSkstoaConfirm(itemid, strParam, iErrStr)
				rw "���ο�û"
				response.flush
				response.clear
			End If
		SET oSkstoa = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### ���ο�û �� ###########################################
	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
		Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_skstoa_regitem SET " & VBCRLF
		strSql = strSql & " accFailcnt = 0  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REG&itemid=4412416
ElseIf action = "CONFIRM" Then
	'##################################### �ӽû�ǰ ��ȸ(��) ���� ###############################
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("Y")
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ȸ ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaConfirmParameter()
			Call fnskstoaRegChkstat(itemid, strParam, iErrStr, iskGoodNo)
			rw "������ȸ"
			response.flush
			response.clear
		End If
	SET oSkstoa = nothing
	If Left(iErrStr, 2) <> "OK" Then
		failCnt = failCnt + 1
		SumErrStr = SumErrStr & iErrStr
	Else
		SumOKStr = SumOKStr & iErrStr
	End If
	'##################################### �ӽû�ǰ ��ȸ(��) �� #################################

	'##################################### �ǸŻ�ǰ ��ȸ(��) ���� ###############################
	If failCnt = 0 AND iskGoodNo <> "" Then
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaEditOneItem

			If (oSkstoa.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ȸ ������ ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				strParam = oSkstoa.FOneItem.getSkstoaItemViewParameter()
				Call fnSkstoaItemView(itemid, strParam, iErrStr)
				rw "�ǸŻ�ǰ ��ȸ"
				response.flush
				response.clear
			End If
		SET oSkstoa = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### �ǸŻ�ǰ ��ȸ(��) �� #################################

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
		Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_skstoa_regitem SET " & VBCRLF
		strSql = strSql & " accFailcnt = 0  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=CONFIRM&itemid=4412416
ElseIf action = "PRICE" Then							'IF_API_10_031 / �ǸŻ�ǰ ��ǰ���� ����
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			getMustprice = ""
			getMustprice = oSkstoa.FOneItem.MustPrice()
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaEditPriceParameter()
			Call fnSkstoaEditPrice(itemid, strParam, getMustprice, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=PRICE&itemid=3853757
ElseIf action = "EDIT" Then
'��ǰ ���� PROCESS
' ��ȸ > �Ǹſ��� N�̸� ����
' ��ȸ > ���ݼ��� > �⺻�������� > ��������� > �̹������濩�� �˻��� ���� > �ǸŻ��¼��� > ������ > �ɼ��߰�  > ��ȸ
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

		If oSkstoa.FResultCount = 0 Then
			iErrStr = "ERR||"&itemid&"||���� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
		Else
			'checkTenItemOptionValid2 => �ɼ� ��뿩�ο� �ɼ��߰��ݾ� üũ
			If (oSkstoa.FOneItem.FmaySoldOut = "Y") OR (oSkstoa.FOneItem.IsMayLimitSoldout = "Y") OR (oSkstoa.FOneItem.checkTenItemOptionValid2 <> "True") OR (oSkstoa.FOneItem.FLimityn = "Y" AND (oSkstoa.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) Then
				strParam = ""
				strParam = oSkstoa.FOneItem.getSkstoaSellynParameter("N")
				Call fnSkstoaSellyn(itemid, "N", strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
	'##################################### �ǸŻ�ǰ ��ȸ ���� #######################################
				strParam = ""
				strParam = oSkstoa.FOneItem.getSkstoaItemViewParameter()
				Call fnSkstoaItemView(itemid, strParam, iErrStr)
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
					getMustprice = oSkstoa.FOneItem.MustPrice()
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaEditPriceParameter()
					Call fnskstoaEditPrice(itemid, strParam, getMustprice, iErrStr)
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

	'##################################### ���»� �������� ���� ���� ################################
				If failCnt = 0 Then
					iErrStr = ""
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaItemEditParameter()
					Call fnSkstoaItemEdit(itemid, strParam, iErrStr)
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
	'##################################### ���»� �������� ���� �� ##################################

	'##################################### ���»� ����� ���� ���� ##################################
				If failCnt = 0 Then
					iErrStr = ""
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaEditContentParameter()
					Call fnSkstoaEditContentReg(itemid, strParam, iErrStr)
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
					If oSkstoa.FOneItem.isImageChanged = True Then
						iErrStr = ""
						strParam = ""
						strParam = oSkstoa.FOneItem.getskstoaEditImageParameter()
						Call fnSkstoaEditImage(itemid, strParam, iErrStr)

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
					arrRows = getOptiopnMapList(itemid, "S")
					If isArray(arrRows) Then
						For i = 0 To UBound(arrRows,2)
							If (arrRows(7, i) < 1) OR (arrRows(11, i)= "1") OR (arrRows(9, i)= "N") OR (arrRows(10, i) = "N") Then	'��� 1�� ���ϰų� �ɼǸ��� �ٸ��ų� �ɼǻ�뿩��N �̰ų� �ɼ��Ǹſ��� N�̰ų�
								salegb = "11"
							Else
								salegb = "00"
							End If

							strParam = ""
							strParam = oSkstoa.FOneItem.geskstoaOptionStatParam(arrRows(2, i), salegb)
							Call fnSkstoaOptSellyn(itemid, strParam, iErrStr)
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
					arrRows = getOptiopnMapList(itemid, "C")
					If isArray(arrRows) Then
						For i = 0 To UBound(arrRows,2)
							strParam = ""
							strParam = oSkstoa.FOneItem.geskstoaOptionQtyParam(arrRows(2, i), arrRows(7, i))
							Call fnSkstoaQtyEdit(itemid, strParam, iErrStr)
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
							strParam = oSkstoa.FOneItem.getskstoaOptionAddParam(arrRows(0, i), arrRows(1, i))
							Call fnSkstoaOptAdd(itemid, strParam, iErrStr)

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
						strParam = oSkstoa.FOneItem.getSkstoaItemViewParameter()
						Call fnSkstoaItemView(itemid, strParam, iErrStr)
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
				CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			End If
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_skstoa_regItem SET " & VBCRLF
			strSql = strSql & " accFailcnt = 0  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDIT&itemid=4466673
ElseIf action = "REGAddItem" Then					'IF_API_10_005 / �ӽû�ǰ �������� ���
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getskstoaNotRegOneItem
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Outmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"', '"&CMALLNAME&"' "
			dbget.execute strSql

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oSkstoa.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oSkstoa.FOneItem.getskstoaItemRegParameter()

				getMustprice = ""
				getMustprice = oSkstoa.FOneItem.MustPrice()
				Call fnskstoaItemReg(itemid, strParam, iErrStr, getMustprice, oSkstoa.FOneItem.getskstoaSellYn, oSkstoa.FOneItem.FLimityn, oSkstoa.FOneItem.FLimitNo, oSkstoa.FOneItem.FLimitSold, html2db(oSkstoa.FOneItem.FItemName), oSkstoa.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[�ӽõ��] �ɼǰ˻� ����"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGAddItem&itemid=3853757
ElseIf action = "REGContent" Then					'IF_API_10_001 / �ӽû�ǰ ����� ���
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("")
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oSkstoa.FOneItem.FReglevel <> 1 Then
			iErrStr = "ERR||"&itemid&"||��ǰ���� ����ϼ���."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaContentParameter()
			Call fnSkstoaContentReg(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGContent&itemid=3853757
ElseIf action = "REGOpt" Then						'IF_API_10_002 / �ӽû�ǰ ��ǰ���� ���
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("")
	    If (oSkstoa.FResultCount < 1) Then
			failCnt = failCnt + 1
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oSkstoa.FOneItem.FReglevel <> 2 Then
			failCnt = failCnt + 1
			iErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oSkstoa.FOneItem.FReglevel &") "
		Else
			arrRows = getOptionList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaOptParameter(arrRows(0, i), arrRows(1, i))
					Call fnSkstoaOptReg(itemid, strParam, iErrStr)

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
			CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGOpt&itemid=3853757
ElseIf action = "REGImg" Then						'IF_API_10_003 / �ӽû�ǰ �̹��� ���(URL)
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("")
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oSkstoa.FOneItem.FReglevel <> 3 Then
			iErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oSkstoa.FOneItem.FReglevel &") "
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaImageParameter()
			Call fnSkstoaImageReg(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGImg&itemid=3853757
ElseIf action = "REGGosi" Then						'IF_API_10_006 / �ӽû�ǰ ����������� ���
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("")
	    If (oSkstoa.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oSkstoa.FOneItem.FReglevel <> 4 Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oSkstoa.FOneItem.FReglevel &") "
		Else
			arrRows = getInfoCodeMapList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaGosiRegParameter(arrRows(0, i), arrRows(1, i), arrRows(2, i))
					Call fnSkstoaGosiReg(itemid, strParam, iErrStr)
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
			CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGGosi&itemid=3853757
ElseIf action = "REGCert" Then						'IF_API_10_028 / KC ���� �߰�
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("")
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaCertParameter() 
			Call fnSkstoaCert(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGCert&itemid=3853757
ElseIf action = "REGConfirm" Then					'IF_API_10_011 / �ӽû�ǰ ���ο�û
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("")
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oSkstoa.FOneItem.FReglevel <> 5 AND oSkstoa.FOneItem.FReglevel <> 6 Then
			iErrStr = "ERR||"&itemid&"||��� ���� Ȯ�� (���� : "& oSkstoa.FOneItem.FReglevel &") "
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaConfirmParameter()
			Call fnSkstoaConfirm(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGConfirm&itemid=3853757
ElseIf action = "REGCHKSTAT" Then					'IF_API_10_013 / �ӽû�ǰ ��ȸ(��) -> ���⼭ �ǸŻ�ǰ�ڵ带 ��´�.
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("Y")
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ȸ ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaConfirmParameter()
			Call fnskstoaRegChkstat(itemid, strParam, iErrStr, iskGoodNo)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGCHKSTAT&itemid=3853757
ElseIf action = "EditSellYn" Then					'IF_API_10_023 / ��ǰ �Ǹ��ߴ� ó��
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getSkstoaSellynParameter(chgSellYn)
			Call fnSkstoaSellyn(itemid, chgSellYn, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EditSellYn&itemid=3853757&chgsellyn=N
ElseIf action = "CHKSTAT" Then						'IF_API_10_015 / �ǸŻ�ǰ ��ȸ(��)
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ȸ ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getSkstoaItemViewParameter()
			Call fnSkstoaItemView(itemid, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=CHKSTAT&itemid=3853757
ElseIf action = "EDITINFO" Then						'IF_API_10_017 / �ǸŻ�ǰ �������� ����
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||�������� ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaItemEditParameter()
			Call fnSkstoaItemEdit(itemid, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITINFO&itemid=3853757
ElseIf action = "EDITContent" Then					'IF_API_10_019 / �ǸŻ�ǰ ����� ���
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||����� ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaEditContentParameter()
			Call fnSkstoaEditContentReg(itemid, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITContent&itemid=3853757
ElseIf action = "EDITImage" Then					'IF_API_10_021 / �ǸŻ�ǰ �̹��� ���(URL)
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||�̹��� ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaEditImageParameter()
			Call fnSkstoaEditImage(itemid, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITImage&itemid=3853757
ElseIf action = "EDITQTY" Then						'IF_API_10_018 / �ǸŻ�ǰ �����
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||��� ���� ������ ��ǰ�� �ƴմϴ�."
		ElseIf getSkstoaOptCnt(itemid) = 0 Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||��ȸ���� �����ϼ���."
		Else
			arrRows = getOptiopnMapList(itemid, "C")
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oSkstoa.FOneItem.geskstoaOptionQtyParam(arrRows(2, i), arrRows(7, i))
					Call fnSkstoaQtyEdit(itemid, strParam, iErrStr)
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
			CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
 	SET oSkstoa = nothing
 	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITQTY&itemid=3853757
ElseIf action = "EDITSTAT" Then						'IF_API_10_023 / ��ǰ �Ǹ��ߴ� ó��
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		ElseIf getSkstoaOptCnt(itemid) = 0 Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||��ȸ���� �����ϼ���."
		Else
			arrRows = getOptiopnMapList(itemid, "S")
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					If (arrRows(7, i) < 1) OR (arrRows(11, i)= "1") OR (arrRows(9, i)= "N") OR (arrRows(10, i) = "N") Then	'��� 1�� ���ϰų� �ɼǸ��� �ٸ��ų� �ɼǻ�뿩��N �̰ų� �ɼ��Ǹſ��� N�̰ų�
						salegb = "11"
					Else
						salegb = "00"
					End If

					strParam = ""
					strParam = oSkstoa.FOneItem.geskstoaOptionStatParam(arrRows(2, i), salegb)
					Call fnSkstoaOptSellyn(itemid, strParam, iErrStr)
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
			CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
 	SET oSkstoa = nothing
 	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITSTAT&itemid=4466673
ElseIf action = "EDITADDOPT" Then					'IF_API_10_020 / �ǸŻ�ǰ ��ǰ���� ���
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		ElseIf getSkstoaOptCnt(itemid) = 0 Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||��ȸ���� �����ϼ���."
		Else
			arrRows = getOptiopnMayAddList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaOptionAddParam(arrRows(0, i), arrRows(1, i))
					Call fnSkstoaOptAdd(itemid, strParam, iErrStr)
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
			CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
 	SET oSkstoa = nothing
 	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITADDOPT&itemid=4466673
ElseIf action = "EDITGosi" Then					'IF_API_10_016 / �ǸŻ�ǰ ����������� ���
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||����������� ���� ������ ��ǰ�� �ƴմϴ�."
		Else
			arrRows = getInfoCodeMapList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaGosiEditParameter(arrRows(0, i), arrRows(1, i), arrRows(2, i))
					Call fnSkstoaGosiEdit(itemid, strParam, iErrStr)
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
			CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
 	SET oSkstoa = nothing
 	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITGosi&itemid=4466673
 ElseIf action = "EDITCert" Then					'IF_API_10_030 / �ǸŻ�ǰ KC�������� ����
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

 	    If (oSkstoa.FResultCount < 1) Then
 			iErrStr = "ERR||"&itemid&"||�������� ���� ������ ��ǰ�� �ƴմϴ�."
 		Else
 			strParam = ""
 			strParam = oSkstoa.FOneItem.getskstoaEditCertParameter()
 			Call fnSkstoaEditCert(itemid, strParam, iErrStr)
 		End If

 		If LEFT(iErrStr, 2) <> "OK" Then
 			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
 		End If
 		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
 	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITCert&itemid=4466673
ElseIf action = "commonCode" Then					'IF_API_00_001 ~ 
	Call fnGetCommonCodeList(interfaceId)
	response.end
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=commonCode&interfaceId=IF_API_00_001
ElseIf action = "cateList" Then						'IF_API_00_028 / ��ǰ ���з� ��ȸ
	Call fnGetGoodsDgroupList()
	response.end
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=cateList
ElseIf action = "offerList" Then					'IF_API_00_023 / ��ǰ����������� �׸� ��ȸ
	Call fnGetOfferList()
	response.end
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=offerList
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
'###################################################### skstoa API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
