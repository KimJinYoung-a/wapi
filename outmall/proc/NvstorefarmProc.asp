<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/nvstorefarm/nvstorefarmItemcls.asp"-->
<!-- #include virtual="/outmall/nvstorefarm/incNvstorefarmFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, oNvstorefarm, failCnt, chgSellYn, arrRows, skipItem, sellgubun, getMustprice, chkXML
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, getfarmGoodno
Dim oService, oOperation, mayOptSoldOut, chgImageNm, endItemErrMsgReplace
Dim jenkinsBatchYn, idx, lastErrStr
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
chkXML			= request("chkXML")
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
'######################################################## ������� API ########################################################
If mallid = "nvstorefarm" Then
	If action = "Image" Then													'�̹��� ���
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
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("nvstorefarm", itemid, iErrStr)
		End If
	ElseIf action = "REG" Then													'��ǰ���
		SET oNvstorefarm = new CNvstorefarm
			oNvstorefarm.FRectItemID	= itemid
			oNvstorefarm.getNvstorefarmNotRegOneItem
			If (oNvstorefarm.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			ElseIf oNvstorefarm.FOneItem.FAPIaddImg <> "Y" Then
				iErrStr = "ERR||"&itemid&"||�̹��� ���� ���ε� �ϼ���."
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			ElseIf oNvstorefarm.FOneItem.FNvstoremoonbanguid <> "0" Then
				iErrStr = "ERR||"&itemid&"||������� ���汸 ��ǰ�� �ߺ��Դϴ�."
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
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
					getMustprice = oNvstorefarm.FOneItem.MustPrice()
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
								iErrStr = "ERR||"&itemid&"||"&SumErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
								SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
								CALL Fn_AcctFailTouch("nvstorefarm", itemid, SumErrStr)
								iErrStr = "ERR||"&itemid&"||�ɼ�API ����, ����ó��"
							End If
						Else
							SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
							iErrStr = "OK||"&itemid&"||"&SumOKStr
						End If
					Else
						SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
						iErrStr = "OK||"&itemid&"||"&SumOKStr
					End If
				Else
					SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
					SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
					CALL Fn_AcctFailTouch("nvstorefarm", itemid, SumErrStr)
					iErrStr = "ERR||"&itemid&"||"&SumErrStr
				End If
			End If
		SET oNvstorefarm = nothing

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("nvstorefarm", itemid, SumErrStr)
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
	ElseIf action = "CHKOPT" Then									'�ɼ� ��ȸ
		oService		= "ProductService"
		oOperation		= "GetOption"

		strParam = ""
		strParam = getNvstorefarmOptionSearchParameter(getNvstorefarmGoodNo(itemid), oService, oOperation)
		Call fnNvstorefarmOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("nvstorefarm", itemid, iErrStr)
		End If
		''Call SugiQueLogInsert("nvstorefarm", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	ElseIf action = "SOLDOUT" Then												'���º���
		oService		= "ProductService"
		oOperation		= "ChangeProductSaleStatus"

		strParam = ""
		strParam = getNvstorefarmSellynParameter(getNvstorefarmGoodNo(itemid), "N", oService, oOperation)
		Call fnNvstorefarmSellyn(itemid, "N", strParam, iErrStr, oService, oOperation)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("nvstorefarm", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/NvstorefarmProc.asp?itemid=699617&mallid=nvstorefarm&action=SOLDOUT
	ElseIf action = "DEL" Then										'��ǰ ����
		oService		= "ProductService"
		oOperation		= "DeleteProduct"

		strParam = ""
		strParam = getNvstorefarmDeleteParameter(getNvstorefarmGoodNo(itemid), oService, oOperation)
		Call fnNvstorefarmDelete(itemid, strParam, iErrStr, oService, oOperation)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("nvstorefarm", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/NvstorefarmProc.asp?itemid=3567941&mallid=nvstorefarm&action=DEL
	ElseIf action = "EDIT" OR action = "ITEMNAME" OR action = "PRICE" Then		'��ǰ����
		SET oNvstorefarm = new CNvstorefarm
			oNvstorefarm.FRectItemID	= itemid
			oNvstorefarm.getNvstorefarmEditOneItem

			If (oNvstorefarm.FResultCount < 1)  Then
				iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
				failCnt = failCnt + 1
			ElseIf (oNvstorefarm.FOneItem.FNvstorefarmGoodNo = "") Then
				iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
				failCnt = failCnt + 1
			Else
				If oNvstorefarm.FOneItem.FOptioncnt > 0 Then
					mayOptSoldOut = oNvstorefarm.FOneItem.IsMayLimitSoldout
				End If

				If (oNvstorefarm.FOneItem.FMaySoldOut = "Y") OR (oNvstorefarm.FOneItem.IsSoldOutLimit5Sell) OR (mayOptSoldOut = "Y") Then
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
					If oNvstorefarm.FOneItem.isImageChanged Then
						chgImageNm = oNvstorefarm.FOneItem.getBasicImage
					Else
						chgImageNm = "N"
					End If

					oService		= "ProductService"
					oOperation		= "ManageProduct"

					strParam = ""
					strParam = oNvstorefarm.FOneitem.getNvstorefarmItemRegXML(oService, oOperation, "Y")
					getMustprice = ""
					getMustprice = oNvstorefarm.FOneItem.MustPrice()
					Call fnNvstorefarmItemEDIT(itemid, strParam, iErrStr, getMustprice, oNvstorefarm.FOneItem.getNvstorefarmSellYn, oNvstorefarm.FOneItem.FLimityn, oNvstorefarm.FOneItem.FLimitNo, oNvstorefarm.FOneItem.FLimitSold, (oNvstorefarm.FOneItem.FItemName), chgImageNm, oService, oOperation)
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
						getMustprice = oNvstorefarm.FOneItem.MustPrice()
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
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_nvstorefarm_regitem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oNvstorefarm = nothing
	End If
End If
'###################################################### ������� API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
