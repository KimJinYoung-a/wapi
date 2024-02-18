<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/nvstoregift/nvstoregiftItemcls.asp"-->
<!-- #include virtual="/outmall/nvstoregift/incNvstoregiftFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, oNvstoregift, failCnt, chgSellYn, arrRows, skipItem, sellgubun, getMustprice, chkXML
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
If mallid = "nvstoregift" Then
	If action = "CHKOPT" Then									'�ɼ� ��ȸ
		oService		= "ProductService"
		oOperation		= "GetOption"

		strParam = ""
		strParam = getNvstoregiftOptionSearchParameter(getNvstoregiftGoodNo(itemid), oService, oOperation)
		Call fnNvstoregiftOptionSearch(itemid, strParam, iErrStr, oService, oOperation)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
		End If
	ElseIf action = "SOLDOUT" Then												'���º���
		oService		= "ProductService"
		oOperation		= "ChangeProductSaleStatus"

		strParam = ""
		strParam = getNvstoregiftSellynParameter(getNvstoregiftGoodNo(itemid), "N", oService, oOperation)
		Call fnNvstoregiftSellyn(itemid, "N", strParam, iErrStr, oService, oOperation)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("nvstoregift", itemid, iErrStr)
		End If
		'http://wapi.10x10.co.kr/outmall/proc/NvstoregiftProc.asp?itemid=3144076&mallid=nvstoregift&action=SOLDOUT
	ElseIf action = "EDIT" OR action = "ITEMNAME" OR action = "PRICE" Then		'��ǰ����
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

				If (oNvstoregift.FOneItem.FMaySoldOut = "Y") OR (oNvstoregift.FOneItem.IsSoldOutLimit5Sell) OR (mayOptSoldOut = "Y") Then
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
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_nvstoregift_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oNvstoregift = nothing
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
