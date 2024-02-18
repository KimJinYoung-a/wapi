<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/gsshopAddOpt/gsshopItemcls.asp"-->
<!-- #include virtual="/outmall/gsshopAddOpt/incGSShopFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim idx, mallid, action, oGSShop, failCnt
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname
Dim jenkinsBatchYn, qidx, lastErrStr
idx				= requestCheckVar(request("idx"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
jenkinsBatchYn	= request("jenkinsBatchYn")
qidx			= request("qidx")
lastErrStr		= ""
If Not(isNumeric(idx)) Then
	response.write "<script>alert('�߸��� ��ǰ��ȣ�Դϴ�.')</script>"
	response.end
End If
'######################################################## GSShop API ########################################################
If mallid = "gsshop" Then
	If action = "SOLDOUT" Then								'ǰ��ó��
		strParam = ""
		strParam = getGSShopSellynParameter(idx, "N")

		Call fnGSShopNewSellyn(idx, "N", strParam, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouchOption("gsshop", idx, iErrStr)
		End If
		'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopOptProc.asp?idx=1&mallid=gsshop&act=SOLDOUT
	ElseIf action = "PRICE" Then								'���ݼ���
		strParam = ""
		strParam = getGSShopPriceParameter(idx, mustPrice)
		If strParam = "" Then
			lastErrStr = "ERR||"&idx&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
			response.write "ERR||"&idx&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
		Else
			Call fnGSShopNewPrice(idx, strParam, mustPrice, iErrStr)
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouchOption("gsshop", idx, iErrStr)
			End If
		End If
		'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopOptProc.asp?idx=1&mallid=gsshop&act=PRICE
	ElseIf action = "EDIT" Then									'����&���&�ɼ�&���� ���� | ���� -> ���� -> �ɼ� �� ����
		SET oGSShop = new CGSShop
			oGSShop.FRectIdx		= idx
			oGSShop.getGSShopEditOneItem
			If oGSShop.FResultCount > 0 Then
				If (oGSShop.FOneItem.FmaySoldOut = "Y") OR (oGSShop.FOneItem.IsOptionSoldOut) OR (oGSShop.FOneItem.isDiffName) Then
					strParam = ""
					strParam = getGSShopSellynParameter(idx, "N")
					Call fnGSShopNewSellyn(idx, "N", strParam, iErrStr)

					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					If (oGSShop.FOneItem.FGsshopSellYn = "N" AND oGSShop.FOneItem.FmaySoldOut = "N" AND oGSShop.FOneItem.IsOptionSoldOut = False) Then
						iErrStr = ""
						strParam = ""
						strParam = getGSShopSellynParameter(idx, "Y")
						Call fnGSShopNewSellyn(idx, "Y", strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

					If (oGSShop.FOneItem.FRealSellprice <> oGSShop.FOneItem.FGSShopPrice) Then
						strParam = ""
						strParam = getGSShopPriceParameter(idx, mustPrice)
						If strParam = "" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & "ERR||"&idx&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
						Else
							Call fnGSShopNewPrice(idx, strParam, mustPrice, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						End If
					End If

					'Ÿ�Ӿƿ� ������ ��ǰ��ǰ�� regedoption���̺� �Է��� �� �Ǿ��� ���
					If oGSShop.FOneItem.FLimitYn = "Y" Then
						strSql = ""
						strSql = strSql & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_outmall_regedoption where itemid='"&oGSShop.FOneItem.Fitemid&"' and itemoption = '"&oGSShop.FOneItem.FItemoption&"' and mallid = 'gsshop') "
						strSql = strSql & " BEGIN"& VbCRLF
						strSql = strSql & " insert into db_item.dbo.tbl_outmall_regedoption (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate) values " & VBCRLF
						strSql = strSql & " ('"&oGSShop.FOneItem.Fitemid&"', '"&oGSShop.FOneItem.FItemoption&"', 'gsshop', '"&oGSShop.FOneItem.FGsshopGoodNo&"001', '"&oGSShop.FOneItem.FOptionname&"', 'Y', 'Y', '220', '"&oGSShop.FOneItem.FOptaddprice&"', getdate()) " & VBCRLF
						strSql = strSql & " END "
					Else
						strSql = ""
						strSql = strSql & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_outmall_regedoption where itemid='"&oGSShop.FOneItem.Fitemid&"' and itemoption = '"&oGSShop.FOneItem.FItemoption&"' and mallid = 'gsshop') "
						strSql = strSql & " BEGIN"& VbCRLF
						strSql = strSql & " insert into db_item.dbo.tbl_outmall_regedoption (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate) values " & VBCRLF
						strSql = strSql & " ('"&oGSShop.FOneItem.Fitemid&"', '"&oGSShop.FOneItem.FItemoption&"', 'gsshop', '"&oGSShop.FOneItem.FGsshopGoodNo&"001', '"&oGSShop.FOneItem.FOptionname&"', 'Y', 'N', '999', '"&oGSShop.FOneItem.FOptaddprice&"', getdate()) " & VBCRLF
						strSql = strSql & " END "
					End If
					dbget.Execute strSql

					'�⺻ ���� ����
					strParam = ""
					strParam = oGSShop.FOneItem.getGSShopItemEditParameter()
					Call fnGSShopNewItemInfoEdit(idx, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'�ɼ� �߰� �� ��� ����
					strParam = ""
		            strParam = oGSShop.FOneItem.getGSShopOptParameter()
					Call fnGSShopNewOPTSuEdit(oGSShop.FOneItem.Fitemid, strParam, idx, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'�ɼ� �ǸŻ��� ����
					strParam = ""
		            strParam = oGSShop.FOneItem.getGSShopOptSellParameter()
		            Call fnGSShopNewOPTSellEdit(oGSShop.FOneItem.Fitemid, strParam, idx, oGSShop.FOneItem.FItemoption, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'OK�� ERR�̴� editQuecnt�� + 1�� ��Ŵ..
					'�����ٸ����� editQuecnt ASC, i.lastupdate DESC�� �ߺ��� ����
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_gsshopAddoption_regitem SET " & VBCRLF
					strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
					strSql = strSql & " WHERE midx = '"&idx&"' " & VBCRLF
					dbget.Execute strSql
				End If

				If failCnt > 0 Then
					SumErrStr = replace(SumErrStr, "OK||"&idx&"||", "")
					SumErrStr = replace(SumErrStr, "ERR||"&idx&"||", "")
					lastErrStr = "ERR||"&idx&"||"&SumErrStr
					response.write "ERR||"&idx&"||"&SumErrStr
					CALL Fn_AcctFailTouchOption("gsshop", idx, SumErrStr)
				Else
					SumOKStr = replace(SumOKStr, "OK||"&idx&"||", "")
					lastErrStr = "OK||"&idx&"||"&SumOKStr
					response.write "OK||"&idx&"||"&SumOKStr
				End If

			End If
			'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopOptProc.asp?idx=1&mallid=gsshop&act=EDIT
		SET oGSShop = nothing
	End If
End If
'###################################################### GSShop API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_Option_API_Que_ResultWrite] "&qidx&","&idx&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->