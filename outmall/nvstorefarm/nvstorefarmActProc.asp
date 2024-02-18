<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/nvstorefarm/nvstorefarmItemcls.asp"-->
<!-- #include virtual="/outmall/nvstorefarm/incNvstorefarmFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, oNvstorefarm, failCnt, chgSellYn, arrRows, skipItem, getMustprice, oService, oOperation, chkXML
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, ccd, isItemIdChk, getfarmGoodno
Dim i, chgImageNm, mayOptSoldOut, endItemErrMsgReplace
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
chkXML			= request("chkXML")
ccd				= request("ccd")
failCnt			= 0
Select Case action
	Case "nvstorefarmCommonCode", "CATE", "CATEDETAIL"	isItemIdChk = "N"
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
	SET oNvstorefarm = new CNvstorefarm
		oNvstorefarm.FRectItemID	= itemid
		oNvstorefarm.getNvstorefarmNotRegOneItem
		If (oNvstorefarm.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oNvstorefarm.FOneItem.FAPIaddImg <> "Y" Then
			iErrStr = "ERR||"&itemid&"||�̹��� ���� ���ε� �ϼ���."
		ElseIf oNvstorefarm.FOneItem.FNvstoremoonbanguid <> "0" Then
			iErrStr = "ERR||"&itemid&"||������� ���汸 ��ǰ�� �ߺ��Դϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_nvstorefarm_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstorefarm_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, nvstorefarmstatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oNvstorefarm.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oNvstorefarm.FOneitem.checkTenItemOptionValid Then
				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oNvstorefarm.FOneitem.getNvstorefarmItemRegXML(oService, oOperation, "")

				getMustprice = ""
				getMustprice = oNvstorefarm.FOneItem.getSalePrice()
				Call fnNvstorefarmItemReg(itemid, strParam, iErrStr, getMustprice, oNvstorefarm.FOneItem.getNvstorefarmSellYn, oNvstorefarm.FOneItem.FLimityn, oNvstorefarm.FOneItem.FLimitNo, oNvstorefarm.FOneItem.FLimitSold, html2db(oNvstorefarm.FOneItem.FItemName), oNvstorefarm.FOneItem.FbasicimageNm, oService, oOperation, chkXML)
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
				If oNvstorefarm.FOneitem.FOptioncnt > 0 Then				'�ɼǼ��� 0�̸� ��ǰ�̹Ƿ� �ɼ� ��� X
					oService		= "ProductService"
					oOperation		= "ManageOption"

					getfarmGoodno = getNvstorefarmGoodNo(itemid)
					If getfarmGoodno <> "" Then
						strParam = ""
						strParam = getNvstorefarmOptionRegXML(itemid, getfarmGoodno, oService, oOperation)
						If strParam <> "X" Then
							Call fnNvstorefarmOptionReg(itemid, strParam, iErrStr, oService, oOperation)
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
						strParam = getNvstorefarmDeleteParameter(getfarmGoodno, oService, oOperation)
						Call fnNvstorefarmDelete(itemid, strParam, iErrStr, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							SumErrStr = SumErrStr & iErrStr
							SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
							SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
							CALL Fn_AcctFailTouch("nvstorefarm", itemid, SumErrStr)
							Call SugiQueLogInsert("nvstorefarm", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
							iErrStr = "ERR||"&itemid&"||"&SumErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
							SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
							CALL Fn_AcctFailTouch("nvstorefarm", itemid, SumErrStr)
							iErrStr = "ERR||"&itemid&"||�ɼ�API ����, ����ó��"
						End If
					Else
						SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
						Call SugiQueLogInsert("nvstorefarm", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
						iErrStr = "OK||"&itemid&"||"&SumOKStr
					End If
				Else
					SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
					Call SugiQueLogInsert("nvstorefarm", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
					iErrStr = "OK||"&itemid&"||"&SumOKStr
				End If
			Else
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("nvstorefarm", itemid, SumErrStr)
				Call SugiQueLogInsert("nvstorefarm", action, itemid, Split(iErrStr,"||")(0), SumErrStr, session("ssBctID"))
				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			End If
		End If
	SET oNvstorefarm = nothing
ElseIf action = "REGITEM" Then									'��ǰ �⺻ ���� ���
	SET oNvstorefarm = new CNvstorefarm
		oNvstorefarm.FRectItemID	= itemid
		oNvstorefarm.getNvstorefarmNotRegOneItem
		If (oNvstorefarm.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oNvstorefarm.FOneItem.FAPIaddImg <> "Y" Then
			iErrStr = "ERR||"&itemid&"||�̹��� ���� ���ε� �ϼ���."
		ElseIf oNvstorefarm.FOneItem.FNvstoremoonbanguid <> "0" Then
			iErrStr = "ERR||"&itemid&"||������� ���汸 ��ǰ�� �ߺ��Դϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_nvstorefarm_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstorefarm_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, nvstorefarmstatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oNvstorefarm.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oNvstorefarm.FOneitem.checkTenItemOptionValid Then
				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oNvstorefarm.FOneitem.getNvstorefarmItemRegXML(oService, oOperation, "")

				getMustprice = ""
				getMustprice = oNvstorefarm.FOneItem.getSalePrice()
				Call fnNvstorefarmItemReg(itemid, strParam, iErrStr, getMustprice, oNvstorefarm.FOneItem.getNvstorefarmSellYn, oNvstorefarm.FOneItem.FLimityn, oNvstorefarm.FOneItem.FLimitNo, oNvstorefarm.FOneItem.FLimitSold, html2db(oNvstorefarm.FOneItem.FItemName), oNvstorefarm.FOneItem.FbasicimageNm, oService, oOperation, chkXML)
			Else
				iErrStr = "ERR||"&itemid&"||�ɼǰ˻� ����"
			End If
		End If
	SET oNvstorefarm = nothing
ElseIf action = "REGOPT" Then									'�ɼ� ���
	oService		= "ProductService"
	oOperation		= "ManageOption"

	getfarmGoodno = getNvstorefarmGoodNo(itemid)
	If getfarmGoodno <> "" Then
		strParam = ""
		strParam = getNvstorefarmOptionRegXML(itemid, getfarmGoodno, oService, oOperation)
		If strParam <> "X" Then
			Call fnNvstorefarmOptionReg(itemid, strParam, iErrStr, oService, oOperation)
		End If
	Else
		iErrStr = "ERR||"&itemid&"||��ǰ�� ��ϵ��� �ʾҽ��ϴ�."
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstorefarm", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstorefarm", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "Image" Then									'�̹��� ���
	SET oNvstorefarm = new CNvstorefarm
		oNvstorefarm.FRectItemID	= itemid
		oNvstorefarm.FRectGubun		= "IMG"
		oNvstorefarm.getNvstorefarmNotRegOneItem
		If (oNvstorefarm.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oNvstorefarm.FOneItem.FNvstoremoonbanguid <> "0" Then
			iErrStr = "ERR||"&itemid&"||������� ���汸 ��ǰ�� �ߺ��Դϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_nvstorefarm_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstorefarm_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, nvstorefarmstatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oNvstorefarm.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oNvstorefarm.FOneitem.checkTenItemOptionValid Then
				oService		= "ImageService"
				oOperation		= "UploadImage"

				strParam = ""
				strParam = oNvstorefarm.FOneitem.getNvstorefarmImageRegXML(oService, oOperation)
				chgImageNm = oNvstorefarm.FOneItem.getBasicImage
				Call fnNvstorefarmImageReg(itemid, strParam, iErrStr, chgImageNm, oService, oOperation)
			Else
				iErrStr = "ERR||"&itemid&"||�ɼǰ˻� ����"
			End If
		End If
	SET oNvstorefarm = nothing
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstorefarm", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstorefarm", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKOPT" Then									'�ɼ� ��ȸ
	oService		= "ProductService"
	oOperation		= "GetOption"

	strParam = ""
	strParam = getNvstorefarmOptionSearchParameter(getNvstorefarmGoodNo(itemid), oService, oOperation)
	Call fnNvstorefarmOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstorefarm", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstorefarm", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKSTAT" Then									'��ǰ ��ȸ
	oService		= "ProductService"
	oOperation		= "GetProduct"

	strParam = ""
	strParam = getNvstorefarmItemSearchParameter(getNvstorefarmGoodNo(itemid), oService, oOperation)
	Call fnNvstorefarmItemSearch(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstorefarm", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstorefarm", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDIT" Then										'��ǰ��ȸ -> �ɼ���ȸ -> ��ǰ���� -> �ɼǼ��� ��
	SET oNvstorefarm = new CNvstorefarm
		oNvstorefarm.FRectItemID	= itemid
		oNvstorefarm.getNvstorefarmEditOneItem

		If (oNvstorefarm.FResultCount < 1) OR (oNvstorefarm.FOneItem.FNvstorefarmGoodNo = "") Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
			failCnt = failCnt + 1
		Else
			If oNvstorefarm.FOneItem.FOptioncnt > 0 Then
				mayOptSoldOut = oNvstorefarm.FOneItem.IsMayLimitSoldout
			End If

			If (oNvstorefarm.FOneItem.FMaySoldOut = "Y") OR (oNvstorefarm.FOneItem.IsSoldOutLimit5Sell) OR (mayOptSoldOut = "Y") OR (oNvstorefarm.FOneItem.FLimitYn = "Y" AND oNvstorefarm.FOneItem.getiszeroWonSoldOut(itemid) = "Y") Then
				oService		= "ProductService"
				oOperation		= "ChangeProductSaleStatus"

				strParam = ""
				strParam = getNvstorefarmSellynParameter(oNvstorefarm.FOneItem.FNvstorefarmGoodNo, "N", oService, oOperation)
				Call fnNvstorefarmSellyn(itemid, "N", strParam, iErrStr, oService, oOperation)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (oNvstorefarm.FOneItem.FNvstorefarmSellYn = "N" AND oNvstorefarm.FOneItem.IsSoldOutLimit5Sell = False) Then
					oService		= "ProductService"
					oOperation		= "ChangeProductSaleStatus"

					strParam = ""
					strParam = getNvstorefarmSellynParameter(oNvstorefarm.FOneItem.FNvstorefarmGoodNo, "Y", oService, oOperation)
					Call fnNvstorefarmSellyn(itemid, "Y", strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
	'################################################ 0.��ǰ ��������(ReturnCostReason �ű� �ʵ� ����..) ####################
				If oNvstorefarm.FOneItem.isImageChanged = True Then
					chgImageNm = oNvstorefarm.FOneItem.getBasicImage
				Else
					chgImageNm = "N"
				End If

				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oNvstorefarm.FOneitem.getNvstorefarmItemRegXML(oService, oOperation, "Y")
				getMustprice = ""
				getMustprice = oNvstorefarm.FOneItem.getSalePrice()
				Call fnNvstorefarmItemEDIT(itemid, strParam, iErrStr, getMustprice, oNvstorefarm.FOneItem.getNvstorefarmSellYn, oNvstorefarm.FOneItem.FLimityn, oNvstorefarm.FOneItem.FLimitNo, oNvstorefarm.FOneItem.FLimitSold, oNvstorefarm.FOneItem.FItemName, chgImageNm, oService, oOperation)
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
				strParam = getNvstorefarmOptionSearchParameter(oNvstorefarm.FOneItem.FNvstorefarmGoodNo, oService, oOperation)
				Call fnNvstorefarmOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
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
					strParam = oNvstorefarm.FOneitem.getNvstorefarmImageRegXML(oService, oOperation)
					chgImageNm = oNvstorefarm.FOneItem.getBasicImage
					Call fnNvstorefarmImageReg(itemid, strParam, iErrStr, chgImageNm, oService, oOperation)
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
					strParam = oNvstorefarm.FOneitem.getNvstorefarmItemRegXML(oService, oOperation, "Y")
					getMustprice = ""
					getMustprice = oNvstorefarm.FOneItem.getSalePrice()
					Call fnNvstorefarmItemEDIT(itemid, strParam, iErrStr, getMustprice, oNvstorefarm.FOneItem.getNvstorefarmSellYn, oNvstorefarm.FOneItem.FLimityn, oNvstorefarm.FOneItem.FLimitNo, oNvstorefarm.FOneItem.FLimitSold, (oNvstorefarm.FOneItem.FItemName), chgImageNm, oService, oOperation)
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
					strParam = getNvstorefarmOptionRegXML(itemid, oNvstorefarm.FOneItem.FNvstorefarmGoodno, oService, oOperation)
					If strParam <> "X" Then
						Call fnNvstorefarmOptionReg(itemid, strParam, iErrStr, oService, oOperation)
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
					strParam = getNvstorefarmOptionSearchParameter(oNvstorefarm.FOneItem.FNvstorefarmGoodNo, oService, oOperation)
					Call fnNvstorefarmOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
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
					strParam = getNvstorefarmDeleteParameter(oNvstorefarm.FOneItem.FNvstorefarmGoodNo, oService, oOperation)
					Call fnNvstorefarmDelete(itemid, strParam, iErrStr, oService, oOperation)
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
		strSql = strSql & " UPDATE [db_etcmall].[dbo].tbl_nvstorefarm_regitem SET " & VBCRLF
		strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
		strSql = strSql & " ,nvstorefarmlastupdate = getdate()  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("nvstorefarm", itemid, SumErrStr)
			Call SugiQueLogInsert("nvstorefarm", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_nvstorefarm_regitem SET " & VBCRLF
			strSql = strSql & " accFailcnt = 0  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("nvstorefarm", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oNvstorefarm = nothing
ElseIf action = "EditSellYn" Then								'���� ����
	oService		= "ProductService"
	oOperation		= "ChangeProductSaleStatus"

	strParam = ""
	strParam = getNvstorefarmSellynParameter(getNvstorefarmGoodNo(itemid), chgSellYn, oService, oOperation)
	Call fnNvstorefarmSellyn(itemid, chgSellYn, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstorefarm", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstorefarm", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "DEL" Then										'��ǰ ����
	oService		= "ProductService"
	oOperation		= "DeleteProduct"

	strParam = ""
	strParam = getNvstorefarmDeleteParameter(getNvstorefarmGoodNo(itemid), oService, oOperation)
	Call fnNvstorefarmDelete(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstorefarm", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstorefarm", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "nvstorefarmCommonCode" Then					'�����ڵ� �˻�
	If ccd = "GetAddressBookList" Then
		strParam = ""
		strParam = getAddressBookList(ccd)
	End If
'	Call fnAuctionCommonCode(ccd, strParam)
ElseIf action = "CATE" Then										'ī�װ� �˻�
	ccd = "GetAllCategoryList"
	strParam = getAllCategoryList(ccd)
ElseIf action = "CATEDETAIL" Then								'ī�װ� ����ȸ
	ccd = "GetCategoryInfo"
rw "�Ʒ� catekey�� 1000������ ��� ����..sp���� �ּ������ؾ���..response.end ó��.."
'response.end
'	strSql = ""
'	strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_nvstorefarm_certInfo] "
'	dbget.Execute strSql

'	strSql = ""
'	strSql = strSql & " SELECT CateKey "
'	strSql = strSql & " FROM db_etcmall.dbo.tbl_nvstorefarm_category "
'	'strSql = strSql & " WHERE CateKey <= 50001777 "
'	'strSql = strSql & " WHERE CateKey > 50001777 AND CateKey <= 50002805 "
'	'strSql = strSql & " WHERE CateKey > 50002805 AND CateKey <= 50003911 "
'	'strSql = strSql & " WHERE CateKey > 50003911 AND CateKey <= 50005605 "
'	strSql = strSql & " WHERE CateKey > 50005605 "
'	strSql = strSql & " GROUP BY CateKey "
'	strSql = strSql & " ORDER BY CateKey "
'	rsget.Open strSql,dbget,1
'	If not rsget.Eof Then
'		arrRows = rsget.getRows()
'	End If
'	rsget.Close
'response.end

	strSql = "exec [db_etcmall].[dbo].[usp_API_Nvstorefarm_CatekeyList_Get]"
	rsget.CursorLocation = adUseClient
	rsget.CursorType = adOpenStatic
	rsget.LockType = adLockOptimistic
	rsget.Open strSql, dbget
	If Not(rsget.EOF or rsget.BOF) Then
		arrRows = rsget.getRows()
	End If
	rsget.Close

	For i = 0 To Ubound(arrRows, 2)
		strParam = getCategoryInfo(ccd, arrRows(0, i))
	Next
	rw "OK"
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
