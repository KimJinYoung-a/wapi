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
Dim itemid, mallid, action, oNvclass, failCnt, chgSellYn, arrRows, skipItem, sellgubun, getMustprice
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, getfarmGoodno
Dim oService, oOperation, mayOptSoldOut, chgImageNm
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0

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
If mallid = "nvstorefarmclass" Then
	If action = "SOLDOUT" Then												'���º���
		oService		= "ProductService"
		oOperation		= "ChangeProductSaleStatus"

		strParam = ""
		strParam = getNvClassSellynParameter(getNvClassGoodNo(itemid), "N", oService, oOperation)
		Call fnNvClassSellyn(itemid, "N", strParam, iErrStr, oService, oOperation)
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("nvstorefarmclass", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/NvClassProc.asp?itemid=699617&mallid=nvstorefarmclass&action=SOLDOUT
	ElseIf action = "EDIT" OR action = "ITEMNAME" OR action = "PRICE" Then		'��ǰ����
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
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_nvstorefarmclass_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oNvclass = nothing
	End If
End If
'###################################################### ������� API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
