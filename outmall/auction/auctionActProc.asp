<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/auction/auctionItemcls.asp"-->
<!-- #include virtual="/outmall/auction/incAuctionFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, oAuction, oAuctionOpt, failCnt, chgSellYn, arrRows, skipItem, tAuctionGoodno, isAllRegYn, getMustprice
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, ccd, isItemIdChk
Dim isoptionyn, isText, i, isiframe
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
ccd				= request("ccd")
failCnt			= 0

Select Case action
	Case "auctionCommonCode"	isItemIdChk = "N"
	Case Else					isItemIdChk = "Y"
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
'######################################################## Auction API ########################################################
If action = "REGAddItem" Then							'��ǰ �⺻ ���� ���
	SET oAuction = new CAuction
		oAuction.FRectItemID	= itemid
		oAuction.getAuctionNotRegOneItem

	    If (oAuction.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf (oAuction.FOneItem.FNotinCate = "Y") Then
			iErrStr = "ERR||"&itemid&"||��ǰ ��� ���� ī�װ��Դϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_auction_regitem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_auction_regitem "
	        strSql = strSql & " 	(itemid, regdate, reguserid, auctionstatCD, regitemname, auctionSellYn)"
	        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oAuction.FOneItem.FItemName)&"', 'N')"
			strSql = strSql & " END "
			dbget.Execute strSql

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oAuction.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oAuction.FOneItem.getAuctionItemRegParameter()
'response.write strParam
'response.end
				getMustprice = ""
				getMustprice = oAuction.FOneItem.MustPrice()
				Call fnAuctionItemReg(itemid, strParam, iErrStr, getMustprice, oAuction.FOneItem.getAuctionSellYn, oAuction.FOneItem.FLimityn, oAuction.FOneItem.FLimitNo, oAuction.FOneItem.FLimitSold, html2db(oAuction.FOneItem.FItemName), oAuction.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[AddItem] �ɼǰ˻� ����"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("auction1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oAuction = nothing
ElseIf action = "REGAddOPT" Then						'��ǰ �ɼ� ���� ���
	SET oAuction = new CAuction
		oAuction.FRectItemID	= itemid
		oAuction.getAuctionNotOptOneItem

	    If (oAuction.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||�ɼ� ��� ������ ��ǰ�� �ƴմϴ�."
		ElseIf (oAuction.FOneItem.FAuctionGoodNo = "") Then
			iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
		ElseIf (oAuction.FOneItem.FAPIadditem = "N") Then
			iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
		ElseIf (oAuction.FOneItem.FAPIaddopt = "Y") Then
			iErrStr = "ERR||"&itemid&"||�̹� �ɼ������� ����ϼ̽��ϴ�."
		Else
			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oAuction.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
				Call fnAuctionOPTReg(itemid, strParam, iErrStr)
			Else
				iErrStr = "ERR||"&itemid&"||[AddOPT] �ɼǰ˻� ����"
			End If

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("auction1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oAuction = nothing
ElseIf action = "REGInfoCd" Then						'��ǰ��� ���
	tAuctionGoodno = getAuctionGoodno(itemid)
	If tAuctionGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
	Else
		strParam = ""
		strParam = getAuctionInfoCdParameter(itemid, tAuctionGoodno)
		Call fnAuctionItemInfoCd(itemid, strParam, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("auction1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "REG" Then								'�⺻���� + �ɼ����� + ������� ���
	'##################################### �⺻ ���� ��� ���� #####################################
	SET oAuction = new CAuction
		oAuction.FRectItemID	= itemid
		oAuction.getAuctionNotRegOneItem
	    If (oAuction.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		ElseIf (oAuction.FOneItem.FNotinCate = "Y") Then
			iErrStr = "ERR||"&itemid&"||��ǰ ��� ���� ī�װ��Դϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_auction_regitem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_auction_regitem "
	        strSql = strSql & " 	(itemid, regdate, reguserid, auctionstatCD, regitemname, auctionSellYn)"
	        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oAuction.FOneItem.FItemName)&"', 'N')"
			strSql = strSql & " END "
			dbget.Execute strSql

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oAuction.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oAuction.FOneItem.getAuctionItemRegParameter()
				getMustprice = ""
				getMustprice = oAuction.FOneItem.MustPrice()
				Call fnAuctionItemReg(itemid, strParam, iErrStr, getMustprice, oAuction.FOneItem.getAuctionSellYn, oAuction.FOneItem.FLimityn, oAuction.FOneItem.FLimitNo, oAuction.FOneItem.FLimitSold, html2db(oAuction.FOneItem.FItemName), oAuction.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[AddItem] �ɼǰ˻� ����"
			End If
		End If
	SET oAuction = nothing
	If Left(iErrStr, 2) <> "OK" Then
		failCnt = failCnt + 1
		SumErrStr = SumErrStr & iErrStr
	Else
		SumOKStr = SumOKStr & iErrStr
	End If
	'##################################### �⺻ ���� ��� �� #####################################

	'#################################### �ɼ� ���� ��� ���� ####################################
	If failCnt = 0 Then
		SET oAuctionOpt = new CAuction
			oAuctionOpt.FRectItemID	= itemid
			oAuctionOpt.getAuctionNotOptOneItem
		    If (oAuctionOpt.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||�ɼ� ��� ������ ��ǰ�� �ƴմϴ�."
			ElseIf (oAuctionOpt.FOneItem.FAuctionGoodNo = "") Then
				iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
			ElseIf (oAuctionOpt.FOneItem.FAPIadditem = "N") Then
				iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
			ElseIf (oAuctionOpt.FOneItem.FAPIaddopt = "Y") Then
				iErrStr = "ERR||"&itemid&"||�̹� �ɼ������� ����ϼ̽��ϴ�."
			Else
				strParam = ""
				strParam = oAuctionOpt.FOneItem.getAuctionOPTRegParameter()
				Call fnAuctionOPTReg(itemid, strParam, iErrStr)
			End If
			tAuctionGoodno = oAuctionOpt.FOneItem.FAuctionGoodNo
		SET oAuctionOpt = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'#################################### �ɼ� ���� ��� �� ####################################

	'################################# ��ǰ��� ���� ��� ���� #################################
	If failCnt = 0 Then
		If tAuctionGoodno = "" Then
			iErrStr = "ERR||"&itemid&"||�⺻�������� �Է��ϼž� �˴ϴ�."
		Else
			strParam = ""
			strParam = getAuctionInfoCdParameter(itemid, tAuctionGoodno)
			Call fnAuctionItemInfoCd(itemid, strParam, iErrStr)
		End If
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'################################## ��ǰ��� ���� ��� �� ##################################
	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("auction1010", itemid, SumErrStr)
		Call SugiQueLogInsert("auction1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("auction1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "REGOnSale" Then						'�ɼ� ��ȸ ��  �űԵ�� ��ǰ �Ǹ������� ����
	isAllRegYn = getAllRegChk(itemid)
	If isAllRegYn <> "Y" Then
		iErrStr = "ERR||"&itemid&"||�⺻����, �ɼ�����, ��ǰ��� �Է��� Ȯ���ϼ���"
	Else
		tAuctionGoodno = getAuctionGoodno(itemid)
		strParam = ""
		strParam = getAuctionOptSellModParameter(tAuctionGoodno)
		Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If

		If failCnt = 0 Then
			strParam = ""
			strParam = getAuctionSellYnParameter("Y", itemid, tAuctionGoodno)
			Call fnAuctionSellyn(itemid, "Y", strParam, iErrStr)
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
			CALL Fn_AcctFailTouch("auction1010", itemid, SumErrStr)
			Call SugiQueLogInsert("auction1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("auction1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	End If
ElseIf action = "EditSellYn" Then						'��ǰ ���� ����
	isAllRegYn = getAllRegChk2(itemid)
	If isAllRegYn <> "Y" Then
		iErrStr = "ERR||"&itemid&"||�⺻����, �ɼ�����, ��ǰ��� �Է��� Ȯ���ϼ���"
	Else
		tAuctionGoodno = getAuctionGoodno(itemid)
		strParam = ""
		strParam = getAuctionSellYnParameter(chgSellYn, itemid, tAuctionGoodno)
		Call fnAuctionSellyn(itemid, chgSellYn, strParam, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("auction1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "EditInfo" Then							'�⺻��������(��ǰ��, ����, �̹���, ��ǰ�󼼵��)
	SET oAuction = new CAuction
		oAuction.FRectItemID	= itemid
		oAuction.getAuctionEditOneItem
	    If (oAuction.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		ElseIf getAllRegChk2(itemid) <> "Y" Then
			iErrStr = "ERR||"&itemid&"||OnSale���� Ȯ���ϼ���"
		Else
			strParam = ""
			strParam = oAuction.FOneItem.getAuctionItemInfoEditParameter()

			getMustprice = ""
			getMustprice = oAuction.FOneItem.MustPrice()
			Call fnAuctionIteminfoEdit(itemid, oAuction.FOneItem.FAuctionGoodNo, iErrStr, strParam, getMustprice)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
			End If

			If (Left(iErrStr,2)) <> "OK" and (Left(iErrStr,2)) <> "ER" Then
				iErrStr = "ERR||"&itemid&"||�߸��� ȣ��"
			End If

			Call SugiQueLogInsert("auction1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oAuction = nothing
ElseIf action = "OPTSTAT" Then							'��ǰ ��ȸ(�ɼ� ���� ��������)
	tAuctionGoodno = getAuctionGoodno(itemid)
	If tAuctionGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||��ϵ� ��ǰ�� �ƴմϴ�."
	Else
		strParam = ""
		strParam = getAuctionOptSellModParameter(tAuctionGoodno)
		Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("auction1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "OPTEDIT" Then							'�ɼ� ���� ����
	SET oAuction = new CAuction
		oAuction.FRectItemID	= itemid
		oAuction.getAuctionEditOneItem
	    If (oAuction.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		ElseIf getAllRegChk2(itemid) <> "Y" Then
			iErrStr = "ERR||"&itemid&"||OnSale���� Ȯ���ϼ���"
		Else
			If (oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt > 0) OR (oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt = 0) Then			'�ٹ����� �ɼ��ְ�, ������ �ɼǵ� ��ϵǾ��ִٸ�..�� �Ѵ� �ɼǻ���

				'## �� 3���� API�� ������ �� ��
				'1.�ɼ��� ���� �ʱ�ȭ
				strParam = ""
				strParam = oAuction.FOneItem.getAuctionOPTDeleteParameter()
				Call fnAuctionOPTDel(itemid, strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'2.�ʱ�ȭ �� �� ����
				strParam = ""
				strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
				Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'3.�ɼ� ��ȸ�� ��������
				strParam = ""
				strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
				Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If oAuction.FOneItem.FOptioncnt = 0 AND oAuction.FOneItem.FRegedoptcnt = 0 Then		'�� �� ��ǰ ����
					If (oAuction.FOneItem.FaccFailCNT > 0 AND InStr(oAuction.FOneItem.FlastErrStr, "�ؽ�Ʈ���� �ּ� 1�� �̻� ����Ǿ�� �մϴ�") > 0) Then
						strParam = ""
						strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
						Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

					strParam = ""
					strParam = oAuction.FOneItem.getAuctionDanPoomModParameter()
					Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					If failCnt = 0 Then
						strSql = ""
						strSql = " DELETE FROM db_item.dbo.tbl_outmall_regedoption WHERE mallid = '"&CMALLNAME&"' and itemid = '"&itemid&"' "
						dbget.Execute(strSql)

						strSql = ""
						strSql = "UPDATE db_etcmall.dbo.tbl_auction_regitem SET regedoptcnt = null WHERE itemid = '"&itemid&"'"
						dbget.Execute(strSql)
					End If

					'2. �ɼ� ��ȸ�� ��������
					strParam = ""
					strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
					Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				ElseIf oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt = 0 Then	'�ٹ����ٻ�ǰ�� �ɼ����� ���� ����ǰ�, ��ϵ� �ɼ��� ���� ����
					'1. �� ����
					strParam = ""
					strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
					Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					If failCnt = 0 Then
						strSql = ""
						strSql = " DELETE FROM db_item.dbo.tbl_outmall_regedoption WHERE mallid = '"&CMALLNAME&"' and itemid = '"&itemid&"' "
						dbget.Execute(strSql)

						strSql = ""
						strSql = "UPDATE db_etcmall.dbo.tbl_auction_regitem SET regedoptcnt = null WHERE itemid = '"&itemid&"'"
						dbget.Execute(strSql)
					End If

					'2. �ɼ� ��ȸ�� ��������
					strParam = ""
					strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
					Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

				ElseIf oAuction.FOneItem.FOptioncnt = 0 AND oAuction.FOneItem.FRegedoptcnt > 0 Then	'�ٹ����ٻ�ǰ�� �ɼ��������� ��ǰ���� ����ǰ�, ��ϵ� �ɼ��� �ִ� ����
					'1.�ɼ��� ���� �ʱ�ȭ
					strParam = ""
					strParam = oAuction.FOneItem.getAuctionOPTDeleteParameter()
					Call fnAuctionOPTDel(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'2.�ɼ� ��ȸ�� ��������
					strParam = ""
					strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
					Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
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
				CALL Fn_AcctFailTouch("auction1010", itemid, SumErrStr)
				Call SugiQueLogInsert("auction1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				Call SugiQueLogInsert("auction1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
				iErrStr = "OK||"&itemid&"||"&SumOKStr
			End If
		End If
	SET oAuction = nothing
ElseIf action = "EDIT" Then
	If getAllRegChk2(itemid) <> "Y" Then
		iErrStr = "ERR||"&itemid&"||OnSale���� Ȯ���ϼ���"
	Else
		SET oAuction = new CAuction
			oAuction.FRectItemID	= itemid
			oAuction.getAuctionEditOneItem
			If oAuction.FResultCount > 0 Then

				If oAuction.FOneItem.checkItemContent = "Y" Then
					isiframe = "Y"
				End If
'rw Instr(oAuction.FOneItem.FItemcontent, "<IFRAME")
'response.end
'response.write isiframe & "<br />"
				If (oAuction.FOneItem.FmaySoldOut = "Y") OR (isiframe = "Y") OR (oAuction.FOneItem.IsMayLimitSoldout = "Y") Then
					strParam = ""
					strParam = getAuctionSellYnParameter("N", itemid, oAuction.FOneItem.FAuctionGoodNo)
					Call fnAuctionSellyn(itemid, "N", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					If (oAuction.FOneItem.FAuctionSellYn = "N" AND oAuction.FOneItem.IsSoldOut = False) Then
						iErrStr = ""
						strParam = ""
						strParam = getAuctionSellYnParameter("Y", itemid, oAuction.FOneItem.FAuctionGoodNo)
						Call fnAuctionSellyn(itemid, "Y", strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

					strParam = ""
					strParam = oAuction.FOneItem.getAuctionItemInfoEditParameter()
					getMustprice = ""
					getMustprice = oAuction.FOneItem.MustPrice()
					Call fnAuctionIteminfoEdit(itemid, oAuction.FOneItem.FAuctionGoodNo, iErrStr, strParam, getMustprice)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					If (oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt > 0) OR (oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt = 0) Then			'�ٹ����� �ɼ��ְ�, ������ �ɼǵ� ��ϵǾ��ִٸ�..�� �Ѵ� �ɼǻ���
						'## �� 3���� API�� ������ �� ��
						'1.�ɼ��� ���� �ʱ�ȭ
						strParam = ""
						strParam = oAuction.FOneItem.getAuctionOPTDeleteParameter()
						Call fnAuctionOPTDel(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If

						'2.�ʱ�ȭ �� �� ����
						strParam = ""
						strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
						Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If

						'3.�ɼ� ��ȸ�� ��������
						strParam = ""
						strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
						Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					Else
						If oAuction.FOneItem.FOptioncnt = 0 AND oAuction.FOneItem.FRegedoptcnt = 0 Then		'�� �� ��ǰ ����
							If (oAuction.FOneItem.FaccFailCNT > 0 AND InStr(oAuction.FOneItem.FlastErrStr, "�ؽ�Ʈ���� �ּ� 1�� �̻� ����Ǿ�� �մϴ�") > 0) Then
								strParam = ""
								strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
								Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
								If Left(iErrStr, 2) <> "OK" Then
									failCnt = failCnt + 1
									SumErrStr = SumErrStr & iErrStr
								Else
									SumOKStr = SumOKStr & iErrStr
								End If
							End If

							strParam = ""
							strParam = oAuction.FOneItem.getAuctionDanPoomModParameter()
							Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If

							If failCnt = 0 Then
								strSql = ""
								strSql = " DELETE FROM db_item.dbo.tbl_outmall_regedoption WHERE mallid = '"&CMALLNAME&"' and itemid = '"&itemid&"' "
								dbget.Execute(strSql)

								strSql = ""
								strSql = "UPDATE db_etcmall.dbo.tbl_auction_regitem SET regedoptcnt = null WHERE itemid = '"&itemid&"'"
								dbget.Execute(strSql)
							End If

							'2. �ɼ� ��ȸ�� ��������
							strParam = ""
							strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
							Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						ElseIf oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt = 0 Then	'�ٹ����ٻ�ǰ�� �ɼ����� ���� ����ǰ�, ��ϵ� �ɼ��� ���� ����
							'1. �� ����
							strParam = ""
							strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
							Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If

							If failCnt = 0 Then
								strSql = ""
								strSql = " DELETE FROM db_item.dbo.tbl_outmall_regedoption WHERE mallid = '"&CMALLNAME&"' and itemid = '"&itemid&"' "
								dbget.Execute(strSql)

								strSql = ""
								strSql = "UPDATE db_etcmall.dbo.tbl_auction_regitem SET regedoptcnt = null WHERE itemid = '"&itemid&"'"
								dbget.Execute(strSql)
							End If

							'2. �ɼ� ��ȸ�� ��������
							strParam = ""
							strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
							Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If

						ElseIf oAuction.FOneItem.FOptioncnt = 0 AND oAuction.FOneItem.FRegedoptcnt > 0 Then	'�ٹ����ٻ�ǰ�� �ɼ��������� ��ǰ���� ����ǰ�, ��ϵ� �ɼ��� �ִ� ����
							'1.�ɼ��� ���� �ʱ�ȭ
							strParam = ""
							strParam = oAuction.FOneItem.getAuctionOPTDeleteParameter()
							Call fnAuctionOPTDel(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If

							'2.�ɼ� ��ȸ�� ��������
							strParam = ""
							strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
							Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						End If
					End If

					strParam = ""
					strParam = getAuctionInfoCdParameter(itemid, oAuction.FOneItem.FAuctionGoodNo)
					Call fnAuctionItemInfoCd(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'OK�� ERR�̴� editQuecnt�� + 1�� ��Ŵ..
					'�����ٸ����� editQuecnt ASC, i.lastupdate DESC�� �ߺ��� ����
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_auction_regItem SET " & VBCRLF
					strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
					strSql = strSql & " ,AuctionLastUpdate = getdate()  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql
				End If

				If failCnt > 0 Then
					SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
					SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
					CALL Fn_AcctFailTouch("auction1010", itemid, SumErrStr)
					Call SugiQueLogInsert("auction1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
					iErrStr = "ERR||"&itemid&"||"&SumErrStr
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_auction_regItem SET " & VBCRLF
					strSql = strSql & " accFailcnt = 0  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql

					SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
					Call SugiQueLogInsert("auction1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
					iErrStr = "OK||"&itemid&"||"&SumOKStr
				End If
			End If
		SET oAuction = nothing
	End If
ElseIf action = "auctionCommonCode" Then
	Dim isday
	If ccd = "GetNationCode" Then
		strParam = ""
		strParam = getAuctionCommonCodeList(ccd)
	ElseIf ccd = "GetShippingPlaceCode" Then
		strParam = ""
		strParam = getAuctionCommonCodeShipPlace(ccd)
	ElseIf ccd = "GetPaidOrderList" Then
		strSql = ""
	    strSql = strSql&"select top 1 convert(varchar(10),selldate,21) as lastOrdInputDt"
	    strSql = strSql&" from db_temp.dbo.tbl_XSite_TMpOrder"
	    strSql = strSql&" where sellsite='auction1010'"
	    strSql = strSql&" order by selldate desc"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		if (Not rsget.Eof) then
			isday = rsget("lastOrdInputDt")
		end if
		rsget.Close
	ElseIf ccd = "GetDeliveryList" Then
		strParam = ""
		strParam = getAuctionCommonCodeGetDeliveryList(ccd)
	ElseIf ccd = "GetDeliveryPrepareList" Then
		strParam = ""
		strParam = getAuctionOrderList2(ccd,isday)
	End If
	Call fnAuctionCommonCode(ccd, strParam)
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
'###################################################### LotteiMall API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->