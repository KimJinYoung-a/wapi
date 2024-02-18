<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/nvstorefarmClass/nvClassItemcls.asp"-->
<!-- #include virtual="/outmall/nvstorefarmClass/incNvClassFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, oNvclass, failCnt, chgSellYn, arrRows, skipItem, getMustprice, oService, oOperation
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, ccd, isItemIdChk, getfarmGoodno
Dim i, chgImageNm, mayOptSoldOut
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
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
	SET oNvclass = new CNvClass
		oNvclass.FRectItemID	= itemid
		oNvclass.getNvClassNotRegOneItem
		If (oNvclass.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oNvclass.FOneItem.FAPIaddImg <> "Y" Then
			iErrStr = "ERR||"&itemid&"||�̹��� ���� ���ε� �ϼ���."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, nvClassstatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oNvclass.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oNvclass.FOneitem.checkTenItemOptionValid Then
				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oNvclass.FOneitem.getNvClassItemRegXML(oService, oOperation, "")

				getMustprice = ""
				getMustprice = oNvclass.FOneItem.MustPrice()
				Call fnNvClassItemReg(itemid, strParam, iErrStr, getMustprice, oNvclass.FOneItem.getNvClassSellYn, html2db(oNvclass.FOneItem.FItemName), oNvclass.FOneItem.FbasicimageNm, oService, oOperation)
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
				If oNvclass.FOneitem.FOptioncnt > 0 Then				'�ɼǼ��� 0�̸� ��ǰ�̹Ƿ� �ɼ� ��� X
					oService		= "ProductService"
					oOperation		= "ManageOption"

					getfarmGoodno = getNvClassGoodNo(itemid)
					If getfarmGoodno <> "" Then
						strParam = ""
						strParam = getNvClassOptionRegXML(itemid, getfarmGoodno, oService, oOperation)
						If strParam <> "X" Then
							Call fnNvClassOptionReg(itemid, strParam, iErrStr, oService, oOperation)
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
'------------------------------------------------------------ �ɼ� ���� ��� ------------------------------------------------------------
					oService		= "ProductService"
					oOperation		= "GetOption"

					strParam = ""
					strParam = getNvClassOptionSearchParameter(getfarmGoodno, oService, oOperation)
					Call fnNvClassOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
'------------------------------------------------------------ �ɼ� ���� ��ȸ ------------------------------------------------------------
					If failCnt > 0 Then
						'�ɼ� ��� �� ������ ���� API �̿�
						oService		= "ProductService"
						oOperation		= "DeleteProduct"

						strParam = ""
						strParam = getNvClassDeleteParameter(getfarmGoodno, oService, oOperation)
						Call fnNvClassDelete(itemid, strParam, iErrStr, oService, oOperation)
						If Left(iErrStr, 2) <> "OK" Then
							SumErrStr = SumErrStr & iErrStr
							SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
							SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
							CALL Fn_AcctFailTouch("nvstorefarmclass", itemid, SumErrStr)
							Call SugiQueLogInsert("nvstorefarmclass", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
							iErrStr = "ERR||"&itemid&"||"&SumErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
							SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
							CALL Fn_AcctFailTouch("nvstorefarmclass", itemid, SumErrStr)
							iErrStr = "ERR||"&itemid&"||�ɼ�API ����, ����ó��"
						End If
					Else
						SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
						Call SugiQueLogInsert("nvstorefarmclass", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
						iErrStr = "OK||"&itemid&"||"&SumOKStr
					End If
				Else
					SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
					Call SugiQueLogInsert("nvstorefarmclass", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
					iErrStr = "OK||"&itemid&"||"&SumOKStr
				End If
			Else
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("nvstorefarmclass", itemid, SumErrStr)
				Call SugiQueLogInsert("nvstorefarmclass", action, itemid, Split(iErrStr,"||")(0), SumErrStr, session("ssBctID"))
				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			End If
		End If
	SET oNvclass = nothing
ElseIf action = "REGITEM" Then									'��ǰ �⺻ ���� ���
	SET oNvclass = new CNvClass
		oNvclass.FRectItemID	= itemid
		oNvclass.getNvClassNotRegOneItem
		If (oNvclass.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf oNvclass.FOneItem.FAPIaddImg <> "Y" Then
			iErrStr = "ERR||"&itemid&"||�̹��� ���� ���ε� �ϼ���."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, nvClassstatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oNvclass.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oNvclass.FOneitem.checkTenItemOptionValid Then
				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oNvclass.FOneitem.getNvClassItemRegXML(oService, oOperation, "")

				getMustprice = ""
				getMustprice = oNvclass.FOneItem.MustPrice()
				Call fnNvClassItemReg(itemid, strParam, iErrStr, getMustprice, oNvclass.FOneItem.getNvClassSellYn, html2db(oNvclass.FOneItem.FItemName), oNvclass.FOneItem.FbasicimageNm, oService, oOperation)
			Else
				iErrStr = "ERR||"&itemid&"||�ɼǰ˻� ����"
			End If
		End If
	SET oNvclass = nothing
ElseIf action = "REGOPT" Then									'�ɼ� ���
	oService		= "ProductService"
	oOperation		= "ManageOption"

	getfarmGoodno = getNvClassGoodNo(itemid)
	If getfarmGoodno <> "" Then
		strParam = ""
		strParam = getNvClassOptionRegXML(itemid, getfarmGoodno, oService, oOperation)
		If strParam <> "X" Then
			Call fnNvClassOptionReg(itemid, strParam, iErrStr, oService, oOperation)
		End If
	Else
		iErrStr = "ERR||"&itemid&"||��ǰ�� ��ϵ��� �ʾҽ��ϴ�."
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstorefarmclass", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstorefarmclass", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "Image" Then									'�̹��� ���
	SET oNvclass = new CNvClass
		oNvclass.FRectItemID	= itemid
		oNvclass.FRectGubun		= "IMG"
		oNvclass.getNvClassNotRegOneItem

		If (oNvclass.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, nvClassstatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oNvclass.FOneitem.FItemName)&"')"
			strSql = strSql & " END "
			dbget.Execute strSql

			If oNvclass.FOneitem.checkTenItemOptionValid Then
				oService		= "ImageService"
				oOperation		= "UploadImage"

				strParam = ""
				strParam = oNvclass.FOneitem.getNvClassImageRegXML(oService, oOperation)
				chgImageNm = oNvclass.FOneItem.getBasicImage
				Call fnNvClassImageReg(itemid, strParam, iErrStr, chgImageNm, oService, oOperation)
			Else
				iErrStr = "ERR||"&itemid&"||�ɼǰ˻� ����"
			End If
		End If
	SET oNvclass = nothing
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstorefarmclass", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstorefarmclass", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKOPT" Then									'�ɼ� ��ȸ
	oService		= "ProductService"
	oOperation		= "GetOption"

	strParam = ""
	strParam = getNvClassOptionSearchParameter(getNvClassGoodNo(itemid), oService, oOperation)
	Call fnNvClassOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstorefarmclass", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstorefarmclass", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKSTAT" Then									'��ǰ ��ȸ
	oService		= "ProductService"
	oOperation		= "GetProduct"

	strParam = ""
	strParam = getNvClassItemSearchParameter(getNvClassGoodNo(itemid), oService, oOperation)
	Call fnNvClassItemSearch(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstorefarmclass", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstorefarmclass", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDIT" Then										'��ǰ��ȸ -> �ɼ���ȸ -> ��ǰ���� -> �ɼǼ��� ��
	SET oNvclass = new CNvClass
		oNvclass.FRectItemID	= itemid
		oNvclass.getNvClassEditOneItem

		If (oNvclass.FResultCount < 1) OR (oNvclass.FOneItem.FNvClassGoodNo = "") Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
			failCnt = failCnt + 1
		Else
			If oNvclass.FOneItem.FOptioncnt > 0 Then
				mayOptSoldOut = oNvclass.FOneItem.IsMayLimitSoldout
			End If

			If (oNvclass.FOneItem.FMaySoldOut = "Y") OR (mayOptSoldOut = "Y") OR (oNvclass.FOneItem.FLimitYn = "Y" AND oNvclass.FOneItem.getiszeroWonSoldOut(itemid) = "Y") Then
				oService		= "ProductService"
				oOperation		= "ChangeProductSaleStatus"

				strParam = ""
				strParam = getNvClassSellynParameter(oNvclass.FOneItem.FNvClassGoodNo, "N", oService, oOperation)
				Call fnNvClassSellyn(itemid, "N", strParam, iErrStr, oService, oOperation)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (oNvclass.FOneItem.FNvClassSellyn = "N") Then
					oService		= "ProductService"
					oOperation		= "ChangeProductSaleStatus"

					strParam = ""
					strParam = getNvClassSellynParameter(oNvclass.FOneItem.FNvClassGoodNo, "Y", oService, oOperation)
					Call fnNvClassSellyn(itemid, "Y", strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
	'################################################ 0.��ǰ ��������(ReturnCostReason �ű� �ʵ� ����..) ####################
				oService		= "ProductService"
				oOperation		= "ManageProduct"

				strParam = ""
				strParam = oNvclass.FOneitem.getNvClassItemRegXML(oService, oOperation, "Y")
				getMustprice = ""
				getMustprice = oNvclass.FOneItem.MustPrice()
				Call fnNvClassItemEDIT(itemid, strParam, iErrStr, getMustprice, oNvclass.FOneItem.FItemName, chgImageNm, oService, oOperation)

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
				strParam = getNvClassOptionSearchParameter(oNvclass.FOneItem.FNvClassGoodNo, oService, oOperation)
				Call fnNvClassOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
	'##########################################################################################################################
	'################################################ 2.�̹��� ����� �̹��� ����ε� #########################################
				If oNvclass.FOneItem.isImageChanged = True Then
					chgImageNm = oNvclass.FOneItem.getBasicImage
				Else
					chgImageNm = "N"
				End If

				If chgImageNm <> "N" Then
					oService		= "ImageService"
					oOperation		= "UploadImage"

					strParam = ""
					strParam = oNvclass.FOneitem.getNvClassImageRegXML(oService, oOperation)
					chgImageNm = oNvclass.FOneItem.getBasicImage
					Call fnNvClassImageReg(itemid, strParam, iErrStr, chgImageNm, oService, oOperation)
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
					strParam = oNvclass.FOneitem.getNvClassItemRegXML(oService, oOperation, "Y")
					getMustprice = ""
					getMustprice = oNvclass.FOneItem.MustPrice()
					Call fnNvClassItemEDIT(itemid, strParam, iErrStr, getMustprice, oNvclass.FOneItem.FItemName, chgImageNm, oService, oOperation)
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
					strParam = getNvClassOptionRegXML(itemid, oNvclass.FOneItem.FNvClassGoodNo, oService, oOperation)
					If strParam <> "X" Then
						Call fnNvClassOptionReg(itemid, strParam, iErrStr, oService, oOperation)
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
					strParam = getNvClassOptionSearchParameter(oNvclass.FOneItem.FNvClassGoodNo, oService, oOperation)
					Call fnNvClassOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
	'##########################################################################################################################
			End If
		End If
		'OK�� ERR�̴� editQuecnt�� + 1�� ��Ŵ..
		'�����ٸ����� editQuecnt ASC, i.lastupdate DESC�� �ߺ��� ����
		strSql = ""
		strSql = strSql & " UPDATE [db_etcmall].[dbo].tbl_nvstorefarmclass_regItem SET " & VBCRLF
		strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
		strSql = strSql & " ,nvClasslastupdate = getdate()  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("nvstorefarmclass", itemid, SumErrStr)
			Call SugiQueLogInsert("nvstorefarmclass", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_nvstorefarmclass_regItem SET " & VBCRLF
			strSql = strSql & " accFailcnt = 0  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("nvstorefarmclass", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oNvclass = nothing
ElseIf action = "EditSellYn" Then								'���� ����
	oService		= "ProductService"
	oOperation		= "ChangeProductSaleStatus"

	strParam = ""
	strParam = getNvClassSellynParameter(getNvClassGoodNo(itemid), chgSellYn, oService, oOperation)
	Call fnNvClassSellyn(itemid, chgSellYn, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstorefarmclass", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstorefarmclass", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "DEL" Then										'��ǰ ����
	oService		= "ProductService"
	oOperation		= "DeleteProduct"

	strParam = ""
	strParam = getNvClassDeleteParameter(getNvClassGoodNo(itemid), oService, oOperation)
	Call fnNvClassDelete(itemid, strParam, iErrStr, oService, oOperation)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("nvstorefarmclass", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("nvstorefarmclass", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
End If

If iErrStr <> "" Then
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str += '"&iErrStr&"<br>' " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 500);" & vbCrLf &_
					"</script>"
End If
'###################################################### ������� API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
