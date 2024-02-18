<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/shintvshopping/inc_authCheck.asp"-->
<!-- #include virtual="/outmall/shintvshopping/shintvshoppingItemcls.asp"-->
<!-- #include virtual="/outmall/shintvshopping/incShintvshoppingFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, mallid, action, oShintvshopping, failCnt, chgSellYn, arrRows, getMustprice, interfaceId, getShipCostCode
Dim iErrStr, strParam, strSql, SumErrStr, SumOKStr, isItemIdChk, grpVal, rSkip, rLimit, i, salegb
Dim jenkinsBatchYn, idx, lastErrStr
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
jenkinsBatchYn	= request("jenkinsBatchYn")
idx				= request("idx")
lastErrStr		= ""
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
'######################################################## shintvshopping API ########################################################
If mallid = "shintvshopping" Then
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
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
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
'					If failCnt = 0 Then
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
							response.flush
							response.clear
						End If
'					End If
		'##################################### �ǸŻ�ǰ ��ȸ �� #########################################
				End If
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				If InStr(SumErrStr, "������ ����ó��") < 0 Then
					CALL Fn_AcctFailTouch("shintvshopping", itemid, SumErrStr)
				End If
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_shintvshopping_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oShintvshopping = nothing
	ElseIf action = "REGL4" Then
		'##################################### ������� ���� #######################################
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
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
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
'					If failCnt = 0 Then
						strParam = ""
						strParam = oShintvshopping.FOneItem.getShintvshoppingItemViewParameter()
						Call fnShintvshoppingItemView(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
						response.flush
						response.clear
'					End If
		'##################################### �ǸŻ�ǰ ��ȸ �� #########################################
				End If
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("shintvshopping", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_shintvshopping_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oShintvshopping = nothing
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
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
	ElseIf action = "SOLDOUT" Then
		SET oShintvshopping = new CShintvshopping
			oShintvshopping.FRectItemID	= itemid
			oShintvshopping.getShintvshoppingEditOneItem

			If (oShintvshopping.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				strParam = oShintvshopping.FOneItem.getShintvshoppingSellynParameter("N")
				Call fnShintvshoppingSellyn(itemid, "N", strParam, iErrStr)
			End If
			lastErrStr = iErrStr
			response.write iErrStr

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
			End If
		SET oShintvshopping = nothing
	ElseIf action = "PRICE" Then
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
			lastErrStr = iErrStr
			response.write iErrStr

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("shintvshopping", itemid, iErrStr)
			End If
		SET oShintvshopping = nothing
	End If
End If
'###################################################### shintvshopping API END ######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
