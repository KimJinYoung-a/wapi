<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/auction/auctionItemcls.asp"-->
<!-- #include virtual="/outmall/auction/incAuctionFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, failCnt, arrRows, skipItem, oAuction, getMustprice, tAuctionGoodno, oAuctionOpt
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, isiframe, isAllRegYn
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
'######################################################## Auction API ########################################################
If mallid = "auction1010" Then
	If action = "REG" Then					'��ǰ���
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
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/AuctionProc.asp?itemid=1159694&mallid=auction1010&action=REG
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
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/AuctionProc.asp?itemid=1159694&mallid=auction1010&action=REGOnSale
	ElseIf action = "SOLDOUT" Then			'���º���
		strParam = ""
		strParam = getAuctionSellYnParameter("N", itemid, getAuctionGoodno(itemid))
		Call fnAuctionSellyn(itemid, "N", strParam, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/AuctionProc.asp?itemid=1159694&mallid=auction1010&action=SOLDOUT
	ElseIf action = "KEEPSELL" Then		'��ǰ �Ǹ� ����
		strParam = ""
		strParam = getAuctionSellYnParameter("Y", itemid, getAuctionGoodno(itemid))
		Call fnAuctionSellyn(itemid, "Y", strParam, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/AuctionProc.asp?itemid=1159694&mallid=auction1010&action=KEEPSELL
	ElseIf action = "PRICE" Then		'���ݼ���
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
			End If

			If (Left(iErrStr,2)) <> "OK" and (Left(iErrStr,2)) <> "ER" Then
				iErrStr = "ERR||"&itemid&"||�߸��� ȣ��"
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
			End If
		SET oAuction = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/AuctionProc.asp?itemid=1159694&mallid=auction1010&action=PRICE
	ElseIf action = "EDIT" Then			'�����ȸ + ��ǰ���� + ���� + �ʿ信 ���� ��ǰ�ǸŻ��¼���
		SET oAuction = new CAuction
			oAuction.FRectItemID	= itemid
			oAuction.getAuctionEditOneItem
			If oAuction.FResultCount > 0 Then
				If oAuction.FOneItem.checkItemContent = "Y" Then
					isiframe = "Y"
				End If

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

					If oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt > 0 Then			'�ٹ����� �ɼ��ְ�, ������ �ɼǵ� ��ϵǾ��ִٸ�..�� �Ѵ� �ɼǻ���
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
				End If

				'OK�� ERR�̴� editQuecnt�� + 1�� ��Ŵ..
				'�����ٸ����� editQuecnt ASC, i.lastupdate DESC�� �ߺ��� ����
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_auction_regItem SET " & VBCRLF
				strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
				strSql = strSql & " ,AuctionLastUpdate = getdate()  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				If failCnt > 0 Then
					SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
					SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
					CALL Fn_AcctFailTouch("auction1010", itemid, SumErrStr)
					lastErrStr = "ERR||"&itemid&"||"&SumErrStr
					response.write "ERR||"&itemid&"||"&SumErrStr
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_auction_regItem SET " & VBCRLF
					strSql = strSql & " accFailcnt = 0  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql

					SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
					lastErrStr = "OK||"&itemid&"||"&SumOKStr
					response.write "OK||"&itemid&"||"&SumOKStr
				End If
			End If
		SET oAuction = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/AuctionProc.asp?itemid=1159694&mallid=auction1010&action=EDIT
	End If
End If
'###################################################### Auction API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->