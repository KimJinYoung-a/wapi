<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/nvstoregift/nvstoregiftItemcls.asp"-->
<!-- #include virtual="/outmall/nvstoregift/incnvstoregiftFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, oNvstoregift, failCnt, chgSellYn, arrRows, skipItem, getMustprice, oService, oOperation, chkXML
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, ccd, isItemIdChk, getfarmGoodno
Dim i, chgImageNm, mayOptSoldOut, endItemErrMsgReplace
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
chkXML			= request("chkXML")
ccd				= request("ccd")
failCnt			= 0
Select Case action
	Case "nvstorefarmCommonCode"	isItemIdChk = "N"
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
'######################################################## ������� API ########################################################
If action = "REG" Then											'��ǰ �⺻ ���� + �ɼ� ���
	SET oNvstoregift = new CNvstoregift
		oNvstoregift.FRectItemID	= itemid
		oNvstoregift.getNvstoregiftNotRegOneItem
		If (oNvstoregift.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oNvstoregift.FOneItem.FAPIaddImg <> "Y" Then
			iErrStr = "ERR||"&itemid&"||�̹��� ���� ���ε� �ϼ���."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_nvstoregift_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstoregift_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, nvstoregiftstatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oNvstoregift.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oNvstoregift.FOneitem.checkTenItemOptionValid Then
				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oNvstoregift.FOneitem.getNvstoregiftItemRegXML(oService, oOperation, "")

				getMustprice = ""
				getMustprice = oNvstoregift.FOneItem.MustPrice()
				Call fnNvstoregiftItemReg(itemid, strParam, iErrStr, getMustprice, oNvstoregift.FOneItem.getNvstoregiftSellYn, oNvstoregift.FOneItem.FLimityn, oNvstoregift.FOneItem.FLimitNo, oNvstoregift.FOneItem.FLimitSold, html2db(oNvstoregift.FOneItem.FItemName), oNvstoregift.FOneItem.FbasicimageNm, oService, oOperation, chkXML)
			Else
				iErrStr = "ERR||"&itemid&"||�ɼǰ˻� ����"
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
'------------------------------------------------------------ ��ǰ �⺻ ���� ��� ------------------------------------------------------------
			If failCnt = 0  Then
				If oNvstoregift.FOneitem.FOptioncnt > 0 Then				'�ɼǼ��� 0�̸� ��ǰ�̹Ƿ� �ɼ� ��� X
					oService		= "ProductService"
					oOperation		= "ManageOption"

					getfarmGoodno = getNvstoregiftGoodNo(itemid)
					If getfarmGoodno <> "" Then
						strParam = ""
						strParam = getNvstoregiftOptionRegXML(itemid, getfarmGoodno, oService, oOperation)
						If strParam <> "X" Then
							Call fnNvstoregiftOptionReg(itemid, strParam, iErrStr, oService, oOperation)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						End If
					Else
						failCnt = failCnt + 1
						iErrStr = "ERR||"&itemid&"||�̵�ϻ�ǰ�Դϴ�."
						SumErrStr = SumErrStr & iErrStr
					End If

					If failCnt > 0 Then
						'�ɼ� ��� �� ������ ���� API �̿�
						oService		= "ProductService"
						oOperation		= "DeleteProduct"

						strParam = ""
						strParam = getNvstoregiftDeleteParameter(getfarmGoodno, oService, oOperation)
						Call fnNvstoregiftDelete(itemid, strParam, iErrStr, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							SumErrStr = SumErrStr & iErrStr
							SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
							SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
							CALL Fn_AcctFailTouch("nvstoregift", itemid, SumErrStr)
							Call SugiQueLogInsert("nvstoregift", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
							iErrStr = "ERR||"&itemid&"||"&SumErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
							SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
							CALL Fn_AcctFailTouch("nvstoregift", itemid, SumErrStr)
							iErrStr = "ERR||"&itemid&"||�ɼ�API ����, ����ó��"
						End If
					Else
						SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
						Call SugiQueLogInsert("nvstoregift", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
						iErrStr = "OK||"&itemid&"||"&SumOKStr
					End If
				Else
					SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
					Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
					iErrStr = "OK||"&itemid&"||"&SumOKStr
				End If
			Else
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("nvstoregift", itemid, SumErrStr)
				Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), SumErrStr, session("ssBctID"))
				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			End If
		End If
	SET oNvstoregift = nothing
ElseIf action = "REGITEM" Then									'��ǰ �⺻ ���� ���
	SET oNvstoregift = new CNvstoregift
		oNvstoregift.FRectItemID	= itemid
		oNvstoregift.getNvstoregiftNotRegOneItem
		If (oNvstoregift.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oNvstoregift.FOneItem.FAPIaddImg <> "Y" Then
			iErrStr = "ERR||"&itemid&"||�̹��� ���� ���ε� �ϼ���."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_nvstoregift_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstoregift_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, nvstoregiftstatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oNvstoregift.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oNvstoregift.FOneitem.checkTenItemOptionValid Then
				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oNvstoregift.FOneitem.getNvstoregiftItemRegXML(oService, oOperation, "")

				getMustprice = ""
				getMustprice = oNvstoregift.FOneItem.MustPrice()
				Call fnNvstoregiftItemReg(itemid, strParam, iErrStr, getMustprice, oNvstoregift.FOneItem.getNvstoregiftSellYn, oNvstoregift.FOneItem.FLimityn, oNvstoregift.FOneItem.FLimitNo, oNvstoregift.FOneItem.FLimitSold, html2db(oNvstoregift.FOneItem.FItemName), oNvstoregift.FOneItem.FbasicimageNm, oService, oOperation, chkXML)
			Else
				iErrStr = "ERR||"&itemid&"||�ɼǰ˻� ����"
			End If
		End If
	SET oNvstoregift = nothing
ElseIf action = "REGOPT" Then									'�ɼ� ���
	oService		= "ProductService"
	oOperation		= "ManageOption"

	getfarmGoodno = getNvstoregiftGoodNo(itemid)
	If getfarmGoodno <> "" Then
		strParam = ""
		strParam = getNvstoregiftOptionRegXML(itemid, getfarmGoodno, oService, oOperation)
		If strParam <> "X" Then
			Call fnNvstoregiftOptionReg(itemid, strParam, iErrStr, oService, oOperation)
		Else
			iErrStr = "OK||"&itemid&"||��ǰ���� API���� ���ʿ�..by������"
		End If
	Else
		iErrStr = "ERR||"&itemid&"||��ǰ�� ��ϵ��� �ʾҽ��ϴ�."
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "Image" Then									'�̹��� ���
	SET oNvstoregift = new CNvstoregift
		oNvstoregift.FRectItemID	= itemid
		oNvstoregift.FRectGubun		= "IMG"
		oNvstoregift.getNvstoregiftNotRegOneItem
		If (oNvstoregift.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_nvstoregift_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstoregift_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, nvstoregiftstatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oNvstoregift.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oNvstoregift.FOneitem.checkTenItemOptionValid Then
				oService		= "ImageService"
				oOperation		= "UploadImage"

				strParam = ""
				strParam = oNvstoregift.FOneitem.getNvstoregiftImageRegXML(oService, oOperation)
				chgImageNm = oNvstoregift.FOneItem.getBasicImage
				Call fnNvstoregiftImageReg(itemid, strParam, iErrStr, chgImageNm, oService, oOperation)
			Else
				iErrStr = "ERR||"&itemid&"||�ɼǰ˻� ����"
			End If
		End If
	SET oNvstoregift = nothing
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKOPT" Then									'�ɼ� ��ȸ
	oService		= "ProductService"
	oOperation		= "GetOption"

	strParam = ""
	strParam = getNvstoregiftOptionSearchParameter(getNvstoregiftGoodNo(itemid), oService, oOperation)
	Call fnNvstoregiftOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKSTAT" Then									'��ǰ ��ȸ
	oService		= "ProductService"
	oOperation		= "GetProduct"

	strParam = ""
	strParam = getNvstoregiftItemSearchParameter(getNvstoregiftGoodNo(itemid), oService, oOperation)
	Call fnNvstoregiftItemSearch(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDIT" Then										'��ǰ��ȸ -> �ɼ���ȸ -> ��ǰ���� -> �ɼǼ��� ��
	SET oNvstoregift = new CNvstoregift
		oNvstoregift.FRectItemID	= itemid
		oNvstoregift.getNvstoregiftEditOneItem

		If (oNvstoregift.FResultCount < 1) OR (oNvstoregift.FOneItem.FNvstoregiftGoodNo = "") Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
			failCnt = failCnt + 1
		Else
			If oNvstoregift.FOneItem.FOptioncnt > 0 Then
				mayOptSoldOut = oNvstoregift.FOneItem.IsMayLimitSoldout
			End If

			If (oNvstoregift.FOneItem.FMaySoldOut = "Y") OR (oNvstoregift.FOneItem.IsSoldOutLimit5Sell) OR (mayOptSoldOut = "Y") OR (oNvstoregift.FOneItem.FLimitYn = "Y" AND oNvstoregift.FOneItem.getiszeroWonSoldOut(itemid) = "Y") Then
				oService		= "ProductService"
				oOperation		= "ChangeProductSaleStatus"

				strParam = ""
				strParam = getNvstoregiftSellynParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, "N", oService, oOperation)
				Call fnNvstoregiftSellyn(itemid, "N", strParam, iErrStr, oService, oOperation)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (oNvstoregift.FOneItem.FNvstoregiftSellYn = "N" AND oNvstoregift.FOneItem.IsSoldOutLimit5Sell = False) Then
					oService		= "ProductService"
					oOperation		= "ChangeProductSaleStatus"

					strParam = ""
					strParam = getNvstoregiftSellynParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, "Y", oService, oOperation)
					Call fnNvstoregiftSellyn(itemid, "Y", strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
	'################################################ 0.��ǰ ��������(ReturnCostReason �ű� �ʵ� ����..) ####################
				If oNvstoregift.FOneItem.isImageChanged = True Then
					chgImageNm = oNvstoregift.FOneItem.getBasicImage
				Else
					chgImageNm = "N"
				End If

				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oNvstoregift.FOneitem.getNvstoregiftItemRegXML(oService, oOperation, "Y")
				getMustprice = ""
				getMustprice = oNvstoregift.FOneItem.MustPrice()
				Call fnNvstoregiftItemEDIT(itemid, strParam, iErrStr, getMustprice, oNvstoregift.FOneItem.getNvstoregiftSellYn, oNvstoregift.FOneItem.FLimityn, oNvstoregift.FOneItem.FLimitNo, oNvstoregift.FOneItem.FLimitSold, oNvstoregift.FOneItem.FItemName, chgImageNm, oService, oOperation)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
	'################################################ 1.�ɼ� ��������(regedoption����) #######################################
				oService		= "ProductService"
				oOperation		= "GetOption"

				strParam = ""
				strParam = getNvstoregiftOptionSearchParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, oService, oOperation)
				Call fnNvstoregiftOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
	'##########################################################################################################################
	'################################################ 2.�̹��� ����� �̹��� ����ε� #########################################
				If chgImageNm <> "N" Then
					oService		= "ImageService"
					oOperation		= "UploadImage"

					strParam = ""
					strParam = oNvstoregift.FOneitem.getNvstoregiftImageRegXML(oService, oOperation)
					chgImageNm = oNvstoregift.FOneItem.getBasicImage
					Call fnNvstoregiftImageReg(itemid, strParam, iErrStr, chgImageNm, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
	'##########################################################################################################################
	'############################################## 3.����Ƚ���� 0�϶� ��ǰ���� ###############################################
				If failCnt = "0" Then
					oService		= "ProductService"
					oOperation		= "ManageProduct"

					strParam = ""
					strParam = oNvstoregift.FOneitem.getNvstoregiftItemRegXML(oService, oOperation, "Y")
					getMustprice = ""
					getMustprice = oNvstoregift.FOneItem.MustPrice()
					Call fnNvstoregiftItemEDIT(itemid, strParam, iErrStr, getMustprice, oNvstoregift.FOneItem.getNvstoregiftSellYn, oNvstoregift.FOneItem.FLimityn, oNvstoregift.FOneItem.FLimitNo, oNvstoregift.FOneItem.FLimitSold, (oNvstoregift.FOneItem.FItemName), chgImageNm, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
	'##########################################################################################################################
	'############################################## 4.�ɼǼ��� ################################################################
					oService		= "ProductService"
					oOperation		= "ManageOption"

					strParam = ""
					strParam = getNvstoregiftOptionRegXML(itemid, oNvstoregift.FOneItem.FNvstoregiftGoodno, oService, oOperation)
					If strParam <> "X" Then
						Call fnNvstoregiftOptionReg(itemid, strParam, iErrStr, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
	'##########################################################################################################################
	'################################################ 5.�ɼ� �������� #######################################
					oService		= "ProductService"
					oOperation		= "GetOption"

					strParam = ""
					strParam = getNvstoregiftOptionSearchParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, oService, oOperation)
					Call fnNvstoregiftOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

	'##########################################################################################################################
				endItemErrMsgReplace = replace(SumErrStr, "OK||"&itemid&"||", "")
				endItemErrMsgReplace = replace(SumErrStr, "ERR||"&itemid&"||", "")

				If (Instr(endItemErrMsgReplace, "��з��� ������ �� �����ϴ�") > 0) OR (Instr(endItemErrMsgReplace, "��з��º����Ҽ������ϴ�") > 0) OR (Instr(endItemErrMsgReplace, "�ɼ��ǿɼǰ�/��뿩���׸���") > 0) OR (Instr(endItemErrMsgReplace, "�ɼ��� �ɼǰ�/��뿩�� �׸���") > 0) OR (Instr(endItemErrMsgReplace, "�ɼǰ��׸��޸�(,)��") > 0) OR (Instr(endItemErrMsgReplace, "�ɼǰ� �׸� �޸�(,)��") > 0) Then
					oService		= "ProductService"
					oOperation		= "DeleteProduct"

					strParam = ""
					strParam = getNvstoregiftDeleteParameter(oNvstoregift.FOneItem.FNvstoregiftGoodNo, oService, oOperation)
					Call fnNvstoregiftDelete(itemid, strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						failCnt = 0
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			End If
		End If
		'OK�� ERR�̴� editQuecnt�� + 1�� ��Ŵ..
		'�����ٸ����� editQuecnt ASC, i.lastupdate DESC�� �ߺ��� ����
		strSql = ""
		strSql = strSql & " UPDATE [db_etcmall].[dbo].tbl_nvstoregift_regItem SET " & VBCRLF
		strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
		strSql = strSql & " ,nvstoregiftlastupdate = getdate()  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("nvstoregift", itemid, SumErrStr)
			Call SugiQueLogInsert("nvstoregift", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_nvstoregift_regItem SET " & VBCRLF
			strSql = strSql & " accFailcnt = 0  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("nvstoregift", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oNvstoregift = nothing
ElseIf action = "EditSellYn" Then								'���� ����
	oService		= "ProductService"
	oOperation		= "ChangeProductSaleStatus"

	strParam = ""
	strParam = getNvstoregiftSellynParameter(getNvstoregiftGoodNo(itemid), chgSellYn, oService, oOperation)
	Call fnNvstoregiftSellyn(itemid, chgSellYn, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "DEL" Then										'��ǰ ����
	oService		= "ProductService"
	oOperation		= "DeleteProduct"

	strParam = ""
	strParam = getNvstoregiftDeleteParameter(getNvstoregiftGoodNo(itemid), oService, oOperation)
	Call fnNvstoregiftDelete(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstoregift", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "nvstorefarmCommonCode" Then					'�����ڵ� �˻�
	If ccd = "GetAddressBookList" Then
		strParam = ""
		strParam = getAddressBookList(ccd)
	End If
'	Call fnAuctionCommonCode(ccd, strParam)
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
'###################################################### ������� API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
