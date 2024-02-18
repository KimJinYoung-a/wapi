<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/gmarket/gmarketItemcls.asp"-->
<!-- #include virtual="/outmall/gmarket/incGmarketFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, failCnt, arrRows, skipItem, oGmarket, getMustprice
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname
Dim tGmarketGoodno, tOptionCnt, tLimityn, isAllRegYn, displayDate, isFiftyUpDown, isiframe
Dim isChild, isLife, isElec
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
'######################################################## Gmarket API ########################################################
If mallid = "gmarket1010" Then
	If action = "REG" Then					'��ǰ���
	'##################################### �⺻ ���� ��� ���� #####################################
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketNotRegOneItem
		    If (oGmarket.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			ElseIf (oGmarket.FOneItem.FDepthCode = "0") Then
				iErrStr = "ERR||"&itemid&"||ī�װ� ��Ī�� �ʿ��մϴ�."
			' ElseIf (oGmarket.FOneItem.FBrandCode = "0") Then
			' 	iErrStr = "ERR||"&itemid&"||�귣�� ��Ī�� �ʿ��մϴ�."
			Else
				strSql = ""
				strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_gmarket_regitem where itemid="&itemid&")"
				strSql = strSql & " BEGIN"& VbCRLF
				strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_gmarket_regitem "
		        strSql = strSql & " 	(itemid, regdate, reguserid, gmarketstatCD, regitemname, gmarketSellYn)"
		        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oGmarket.FOneItem.FItemName)&"', 'N')"
				strSql = strSql & " END "
				dbget.Execute strSql

				'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
				If oGmarket.FOneItem.checkTenItemOptionValid Then
					strParam = ""
					strParam = oGmarket.FOneItem.getGmarketItemRegParameter(FALSE)
					Call fnGmarketItemReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FbasicimageNm)
				Else
					iErrStr = "ERR||"&itemid&"||[AddItem] �ɼǰ˻� ����"
				End If
			End If
		SET oGmarket = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
		'##################################### �⺻ ���� ��� �� #####################################

		'#################################### ��� ���� ��� ���� ####################################
		If failCnt = 0 Then
			tGmarketGoodno = getGmarketGoodno(itemid)
			If tGmarketGoodno = "" Then
				iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
			Else
				strParam = ""
				strParam = getGmarketInfoCdParameter(itemid, tGmarketGoodno)
				Call fnGmarketItemInfoCd(itemid, strParam, iErrStr)
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
		'#################################### ��� ���� ��� �� ####################################

		'#################################### ��� ���� ��� ���� ####################################
		If failCnt = 0 Then
			Call getGmarketChildrenCate(itemid, isChild, isLife, isElec)
			If isChild = "Y" OR isLife = "Y" OR isElec = "Y" Then
				strParam = ""
				strParam = getGmarketChildrenParameter(itemid, tGmarketGoodno, isChild, isLife, isElec)
				Call fnGmarketItemChildren(itemid, strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		End If
		'#################################### ��� ���� ��� �� ####################################

		'#################################### ��ǰ�� ��� ���� ####################################
		If failCnt = 0 Then
			strParam = ""
			strParam = getGmarketReturnFeeParameter(itemid, tGmarketGoodno, CRETURNFEE)
			Call fnGmarketReturnFee(itemid, strParam, CRETURNFEE, iErrStr)

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
		'#################################### ��ǰ�� ��� �� ####################################

		'#################################### �ɼ� ���� ��� ���� ####################################
		If failCnt = 0 Then
			SET oGmarket = new CGmarket
				oGmarket.FRectItemID	= itemid
				oGmarket.getGmarketNotOptOneItem
			    If (oGmarket.FResultCount < 1) Then
					iErrStr = "ERR||"&itemid&"||�ɼ� ��� ������ ��ǰ�� �ƴմϴ�."
				ElseIf (oGmarket.FOneItem.FGmarketGoodNo = "") Then
					iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
				ElseIf (oGmarket.FOneItem.FAPIadditem = "N") Then
					iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
				ElseIf (oGmarket.FOneItem.getFiftyUpDown = "Y") Then
					iErrStr = "ERR||"&itemid&"||�ɼǰ����� 50%�� �ʰ��մϴ�."
				Else
					'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
					If oGmarket.FOneItem.checkTenItemOptionValid Then
						strParam = ""
						strParam = oGmarket.FOneItem.getGmarketItemOptRegParameter()
						Call fnGmarketOPTReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FLimityn, oGmarket.FOneItem.FLimitno, oGmarket.FOneItem.FLimitsold)
					Else
						iErrStr = "ERR||"&itemid&"||[AddOPT] �ɼǰ˻� ����"
					End If
				End If
			SET oGmarket = nothing
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
		'#################################### �ɼ� ���� ��� �� ####################################
		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/gmarketProc.asp?itemid=699617&mallid=gmarket1010&action=REG
	ElseIf action = "REGOnSale" Then						'�űԵ�� ��ǰ �Ǹ������� ����
		isAllRegYn = getAllRegChk(itemid, "X")
		If isAllRegYn <> "Y" Then
			iErrStr = "ERR||"&itemid&"||�⺻����, �ɼ�����, ��ǰ��� �Է��� Ȯ���ϼ���"
		Else
			tGmarketGoodno = getGmarketGoodno(itemid)
			strParam = ""
			strParam = getGmarketAddPriceParameter(itemid, tGmarketGoodno, "Y", mustPrice, displayDate)
			Call fnGmarketItemAddPrice(itemid, strParam, mustPrice, displayDate, "Y", iErrStr)
		End If
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
	ElseIf action = "SOLDOUT" Then			'���º���
		isAllRegYn = getAllRegChk2(itemid, tGmarketGoodno, tOptionCnt, tLimityn, "Y")
		If tGmarketGoodno = "" Then
			iErrStr = "ERR||"&itemid&"||�⺻����, �ɼ�����, ��ǰ��� �Է��� Ȯ���ϼ���"
		Else
			strParam = ""
			strParam = getGmarketAddPriceParameter(itemid, tGmarketGoodno, "N", mustPrice, displayDate)
			Call fnGmarketItemAddPrice(itemid, strParam, mustPrice, displayDate, "N", iErrStr)
		End If
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=SOLDOUT
	ElseIf action = "EDITIMG" Then		'�̹��� ����
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketEditImageOneItem
		    If (oGmarket.FOneItem.FGmarketGoodNo = "") Then
				iErrStr = "ERR||"&itemid&"||�̹��� ���� ������ ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketItemEditImgParameter()
				Call fnGmarketEditImg(itemid, strParam, iErrStr, oGmarket.FOneItem.FbasicimageNm)
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
			End If
		SET oGmarket = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=EDITIMG
	ElseIf action = "EDITINFO" Then		'�⺻���� ����
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketEditOneItem

			If (oGmarket.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
			ElseIf oGmarket.FOneItem.checkItemContent = "Y" Then
				iErrStr = "ERR||"&itemid&"||iframe�� ���� ��ǰ�� ���� �� �� �����ϴ�."
			Else
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketItemRegParameter(TRUE)
				Call fnGmarketIteminfoEdit(itemid, oGmarket.FOneItem.FGmarketGoodNo, oGmarket.FOneItem.FItemName, iErrStr, strParam)
			End If

			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
			End If
		SET oGmarket = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=EDITINFO
	ElseIf action = "EDITPOLICY" Then	'�⺻���� ���� + ��ǰ�� ����
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketEditOneItem

			If (oGmarket.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
			ElseIf oGmarket.FOneItem.checkItemContent = "Y" Then
				iErrStr = "ERR||"&itemid&"||iframe�� ���� ��ǰ�� ���� �� �� �����ϴ�."
			Else
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketItemRegParameter(TRUE)
				Call fnGmarketIteminfoEdit(itemid, oGmarket.FOneItem.FGmarketGoodNo, oGmarket.FOneItem.FItemName, iErrStr, strParam)
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			If failCnt = 0 Then
				strParam = ""
				strParam = getGmarketReturnFeeParameter(itemid, oGmarket.FOneItem.FGmarketGoodNo, CRETURNFEE)
				Call fnGmarketReturnFee(itemid, strParam, CRETURNFEE, iErrStr)

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
				CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oGmarket = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=EDITPOLICY
	ElseIf action = "KEEPSELL" Then		'��ǰ �Ǹ� ����
		isAllRegYn = getAllRegChk2(itemid, tGmarketGoodno, tOptionCnt, tLimityn, "Y")
		If tGmarketGoodno = "" Then
			iErrStr = "ERR||"&itemid&"||�⺻����, �ɼ�����, ��ǰ��� �Է��� Ȯ���ϼ���"
		Else
			strParam = ""
			strParam = getGmarketAddPriceParameter(itemid, tGmarketGoodno, "Y", mustPrice, displayDate)
			Call fnGmarketItemAddPrice(itemid, strParam, mustPrice, displayDate, "Y", iErrStr)
		End If
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("gmarket1010", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=KEEPSELL
	ElseIf action = "PRICE" Then		'���ݼ���
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketEditPriceOptOneItem
			If oGmarket.FResultCount > 0 Then
				'�ɼ��߰��ݾ��� ��ǰ�ݾ��� 50%�ʰ� �˻�
				isFiftyUpDown = oGmarket.FOneItem.getFiftyUpDown

				getMustprice = ""
				getMustprice = oGmarket.FOneItem.MustPrice()
				'���� ǰ���� �ش��ϰų� 50%�ʰ��ϰų� 0���ɼ��� ��� ǰ���� ��..(������ǰ��� ��� 5�����ϵ� ������)
				If (oGmarket.FOneItem.FmaySoldOut = "Y") OR (isFiftyUpDown = "Y") OR (oGmarket.FOneItem.FLimityn = "Y" AND (oGmarket.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) OR (oGmarket.FOneItem.IsMayLimitSoldout = "Y") Then
					strParam = ""
					strParam = oGmarket.FOneItem.getGmarketAddPriceParameter("N", getMustprice, displayDate)
					Call fnGmarketItemAddPrice(itemid, strParam, getMustprice, displayDate, "N", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					SET oGmarket = nothing
				Else
				'�� ���ǿ� �ش����� ������ ������ �Ǹ�ó��
					iErrStr = ""
					strParam = ""
					strParam = oGmarket.FOneItem.getGmarketAddPriceParameter("Y", getMustprice, displayDate)
					Call fnGmarketItemAddPrice(itemid, strParam, getMustprice, displayDate, "Y", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					SET oGmarket = nothing

					SET oGmarket = new CGmarket
						oGmarket.FRectItemID	= itemid
						oGmarket.getGmarketNotOptOneItem
					    If (oGmarket.FResultCount < 1) Then
							iErrStr = "ERR||"&itemid&"||�ɼ� ��� ������ ��ǰ�� �ƴմϴ�."
						ElseIf (oGmarket.FOneItem.FGmarketGoodNo = "") Then
							iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
						ElseIf (oGmarket.FOneItem.FAPIadditem = "N") Then
							iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
						ElseIf (oGmarket.FOneItem.getFiftyUpDown = "Y") Then
							iErrStr = "ERR||"&itemid&"||�ɼǰ����� 50%�� �ʰ��մϴ�."
						Else
							'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
							If oGmarket.FOneItem.checkTenItemOptionValid Then
								strParam = ""
								strParam = oGmarket.FOneItem.getGmarketItemOptRegParameter()
								Call fnGmarketOPTReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FLimityn, oGmarket.FOneItem.FLimitno, oGmarket.FOneItem.FLimitsold)
							Else
								iErrStr = "ERR||"&itemid&"||[AddOPT] �ɼǰ˻� ����"
							End If
						End If
					SET oGmarket = nothing
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			Else
				iErrStr = "ERR||"&itemid&"||������ ������ ����"
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
				CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_gmarket_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=PRICE
	ElseIf action = "EDIT" Then			'�����ȸ + ��ǰ���� + ���� + �ʿ信 ���� ��ǰ�ǸŻ��¼���
		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketEditOneItem
			'#################################### �⺻ ���� ���� ���� ####################################
		    If (oGmarket.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
			ElseIf oGmarket.FOneItem.checkItemContent = "Y" Then
				iErrStr = "ERR||"&itemid&"||iframe�� ���� ��ǰ�� ���� �� �� �����ϴ�."
				isiframe = "Y"
			Else
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketItemRegParameter(TRUE)
				Call fnGmarketIteminfoEdit(itemid, oGmarket.FOneItem.FGmarketGoodNo, oGmarket.FOneItem.FItemName, iErrStr, strParam)
			End If
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			'#################################### ��ǰ�� ���� ���� ���� ####################################
			If (oGmarket.FResultCount > 0) AND (oGmarket.FOneItem.FReturnShippingFee < 100) Then
				strParam = ""
				strParam = getGmarketReturnFeeParameter(itemid, oGmarket.FOneItem.FGmarketGoodNo, CRETURNFEE)
				Call fnGmarketReturnFee(itemid, strParam, CRETURNFEE, iErrStr)

				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If

			'#################################### �̹��� ���� ���� ####################################
			If (oGmarket.FResultCount > 0) AND (oGmarket.FOneItem.isImageChanged) Then
				strParam = ""
				strParam = oGmarket.FOneItem.getGmarketItemEditImgParameter()
				Call fnGmarketEditImg(itemid, strParam, iErrStr, oGmarket.FOneItem.FbasicimageNm)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		SET oGmarket = nothing

		SET oGmarket = new CGmarket
			oGmarket.FRectItemID	= itemid
			oGmarket.getGmarketEditPriceOptOneItem
			'#################################### ��ǰ ���� ���� ���� ####################################
			If oGmarket.FResultCount > 0 Then
				'�ɼ��߰��ݾ��� ��ǰ�ݾ��� 50%�ʰ� �˻�
				isFiftyUpDown = oGmarket.FOneItem.getFiftyUpDown

				getMustprice = ""
				getMustprice = oGmarket.FOneItem.MustPrice()
				'���� ǰ���� �ش��ϰų� ������������ �ְų� 50%�ʰ��ϰų� 0���ɼ��� ��� ǰ���� ��..(������ǰ��� ��� 5�����ϵ� ������)
				If (oGmarket.FOneItem.FmaySoldOut = "Y") OR (isFiftyUpDown = "Y") OR (isiframe = "Y") OR (oGmarket.FOneItem.FLimityn = "Y" AND (oGmarket.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) OR (oGmarket.FOneItem.IsMayLimitSoldout = "Y") Then
					strParam = ""
					strParam = oGmarket.FOneItem.getGmarketAddPriceParameter("N", getMustprice, displayDate)
					Call fnGmarketItemAddPrice(itemid, strParam, getMustprice, displayDate, "N", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						failCnt = 0
						SumOKStr = SumOKStr & iErrStr
					End If
					SET oGmarket = nothing
				Else
				'�� ���ǿ� �ش����� ������ ������ �Ǹ�ó��
					iErrStr = ""
					strParam = ""
					strParam = oGmarket.FOneItem.getGmarketAddPriceParameter("Y", getMustprice, displayDate)
					Call fnGmarketItemAddPrice(itemid, strParam, getMustprice, displayDate, "Y", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					SET oGmarket = nothing
			'#################################### ��ǰ �ɼ� ���� ���� ####################################
					SET oGmarket = new CGmarket
						oGmarket.FRectItemID	= itemid
						oGmarket.getGmarketNotOptOneItem
					    If (oGmarket.FResultCount < 1) Then
							iErrStr = "ERR||"&itemid&"||�ɼ� ��� ������ ��ǰ�� �ƴմϴ�."
						ElseIf (oGmarket.FOneItem.FGmarketGoodNo = "") Then
							iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
						ElseIf (oGmarket.FOneItem.FAPIadditem = "N") Then
							iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
						ElseIf (oGmarket.FOneItem.getFiftyUpDown = "Y") Then
							iErrStr = "ERR||"&itemid&"||�ɼǰ����� 50%�� �ʰ��մϴ�."
						Else
							'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
							If oGmarket.FOneItem.checkTenItemOptionValid Then
								strParam = ""
								strParam = oGmarket.FOneItem.getGmarketItemOptRegParameter()
								Call fnGmarketOPTReg(itemid, strParam, iErrStr, oGmarket.FOneItem.FLimityn, oGmarket.FOneItem.FLimitno, oGmarket.FOneItem.FLimitsold)
							Else
								iErrStr = "ERR||"&itemid&"||[AddOPT] �ɼǰ˻� ����"
							End If
						End If
					SET oGmarket = nothing
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			Else
				iErrStr = "ERR||"&itemid&"||������ ������ ����"
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If

			'OK�� ERR�̴� editQuecnt�� + 1�� ��Ŵ..
			'�����ٸ����� editQuecnt ASC, i.lastupdate DESC�� �ߺ��� ����
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_gmarket_regItem SET " & VBCRLF
			strSql = strSql & " editQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
			strSql = strSql & " ,gmarketlastupdate = getdate()  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("gmarket1010", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_gmarket_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		'http://testwapi.10x10.co.kr/outmall/proc/GmarketProc.asp?itemid=282197&mallid=gmarket1010&action=EDIT
	End If
End If
'###################################################### Gmarket API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
