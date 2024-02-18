<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/gsshopAddOpt/gsshopItemcls.asp"-->
<!-- #include virtual="/outmall/gsshopAddOpt/incGSShopFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim idx, action, oGSShop, failCnt, chgSellYn
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname
Dim itemoption
idx				= requestCheckVar(request("idx"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
failCnt			= 0

If Not(isNumeric(idx)) Then
	response.write "<script>alert('�߸��� ��ǰ��ȣ�Դϴ�.')</script>"
	response.end
End If
'######################################################## GSShop API ########################################################
If action = "EditSellYn" Then								'���º���
	strParam = ""
	strParam = getGSShopSellynParameter(idx, chgSellYn)
	Call fnGSShopNewSellyn(idx, chgSellYn, strParam, iErrStr)
	'response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchOption("gsshop", idx, iErrStr)
	End If
	Call SugiOptionQueLogInsert("gsshop", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "REG" Then									'��ǰ���
	SET oGSShop = new CGSShop
		oGSShop.FRectIdx		= idx
		oGSShop.getGSShopNotRegOneItem
	    If (oGSShop.FResultCount < 1) Then
			iErrStr = "ERR||"&idx&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			If oGSShop.FOneItem.FDivcode = "" Then		'���� ��ǰ�з� ��Ī�� �� �� ī�װ� ��ǰ�̶��..
				iErrStr = "ERR||"&oGSShop.FOneItem.FItemid&"||��ǰ�з� ��Ī�� ���� ���� ��ǰ��ȣ"
'			ElseIf (oGSShop.FOneItem.FDeliveryType = "9" OR oGSShop.FOneItem.FDeliveryType = "7" OR oGSShop.FOneItem.FDeliveryType = "2") AND (oGSShop.FOneItem.FDeliveryCd = "" OR oGSShop.FOneItem.FDeliveryAddrCd = "") Then
'				iErrStr = "ERR||"&oGSShop.FOneItem.FItemid&"||�ù��/�ּ��� ��Ī�� ���� ���� ��ǰ��ȣ"
'			ElseIf oGSShop.FOneItem.FBrandcd = "" Then
'				iErrStr = "ERR||"&oGSShop.FOneItem.FItemid&"||�귣���ڵ� ��Ī�� ���� ���� ��ǰ��ȣ"
			Else
				strSql = ""
				strSql = strSql & " IF NOT Exists(SELECT TOP 1 * FROM db_etcmall.dbo.tbl_gsshopAddoption_regitem where midx="&idx&")"
				strSql = strSql & " BEGIN"& VbCRLF
				strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_gsshopAddoption_regitem "
		        strSql = strSql & " (midx, regdate, reguserid, gsshopstatCD)"
		        strSql = strSql & " VALUES ("&idx&", getdate(), '"&session("SSBctID")&"', '1')"
				strSql = strSql & " END "
				dbget.Execute strSql

				strParam = ""
				strParam = oGSShop.FOneItem.getGSShopItemNewRegParameter()
				If (session("ssBctID") = "icommang") or (session("ssBctID") = "kjy8517") Then
					'rw gsshopAPIURL &"?"& strParam
					'response.end
				End If
				CALL fnGSShopNewItemReg(oGSShop.FOneItem.FItemid, strParam, iErrStr, oGSShop.FOneItem.FRealSellprice, oGSShop.FOneItem.getGSShopSellYn, oGSShop.FOneItem.FLimityn, oGSShop.FOneItem.FOptlimitno, oGSShop.FOneItem.FOptlimitsold, html2db(oGSShop.FOneItem.getRealItemname), oGSShop.FOneItem.FItemoption, idx, html2db(oGSShop.FOneItem.FOptionname))
'				CALL fnGSShopItemReg(oGSShop.FOneItem.FItemid, strParam, iErrStr, oGSShop.FOneItem.FRealSellprice, oGSShop.FOneItem.getGSShopSellYn, oGSShop.FOneItem.FLimityn, oGSShop.FOneItem.FOptlimitno, oGSShop.FOneItem.FOptlimitsold, html2db(oGSShop.FOneItem.getRealItemname), oGSShop.FOneItem.FItemoption, idx, html2db(oGSShop.FOneItem.FOptionname))

				If LEFT(iErrStr, 2) <> "OK" Then
					CALL Fn_AcctFailTouchOption("gsshop", idx, iErrStr)
				End If
				Call SugiOptionQueLogInsert("gsshop", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
			End If
		End If
	SET oGSShop = nothing
ElseIf action = "PRICE" Then
	strParam = ""
	strParam = getGSShopPriceParameter(idx, mustPrice)
	If strParam = "" Then
		response.write "ERR||"&idx&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
	Else
		Call fnGSShopNewPrice(idx, strParam, mustPrice, iErrStr)
		'response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouchOption("gsshop", idx, iErrStr)
		End If
		Call SugiOptionQueLogInsert("gsshop", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "IMAGE" Then
	SET oGSShop = new CGSShop
		oGSShop.FRectIdx		= idx
		oGSShop.getGSShopEditOneItem
		If oGSShop.FResultCount > 0 Then
			strParam = ""
			strParam = oGSShop.FOneItem.getGSShopImageEditParameter()
			Call fnGSShopNewImageEdit(idx, strParam, iErrStr)
			'response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouchOption("gsshop", idx, iErrStr)
			End If
			Call SugiOptionQueLogInsert("gsshop", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oGSShop = nothing
ElseIf action = "EDITINFO" Then
	SET oGSShop = new CGSShop
		oGSShop.FRectIdx		= idx
		oGSShop.getGSShopEditOneItem
		If oGSShop.FResultCount > 0 Then
			strParam = ""
			strParam = oGSShop.FOneItem.getGSShopItemEditParameter()
			Call fnGSShopNewItemInfoEdit(idx, strParam, iErrStr)
			'response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("gsshop", idx, iErrStr)
			End If
			Call SugiQueLogInsert("gsshop", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oGSShop = nothing
ElseIf action = "CONTENT" Then
	SET oGSShop = new CGSShop
		oGSShop.FRectIdx		= idx
		oGSShop.getGSShopEditOneItem
		If oGSShop.FResultCount > 0 Then
			strParam = ""
			strParam = oGSShop.FOneItem.getGSShopContentsEditParameter()

			Call fnGSShopNewContentsEdit(idx, strParam, iErrStr)
			'response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouchOption("gsshop", idx, iErrStr)
			End If
			Call SugiOptionQueLogInsert("gsshop", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oGSShop = nothing
ElseIf action = "ITEMNAME" Then
	strParam = ""
	strParam = getGSShopItemnameParameter(idx, iitemname)
	Call fnGSShopChgNewItemname(idx, strParam, iitemname, iErrStr)
	'response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchOption("gsshop", idx, iErrStr)
	End If
	Call SugiOptionQueLogInsert("gsshop", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "INFODIV" Then								'���ΰ���׸�
	SET oGSShop = new CGSShop
		oGSShop.FRectIdx		= idx
		oGSShop.getGSShopEditOneItem
		If oGSShop.FResultCount > 0 Then
			strParam = ""
			strParam = oGSShop.FOneItem.getGSShopInfodivEditParameter()
			Call fnGSShopNewInfodivEdit(idx, strParam, iErrStr)
'			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouchOption("gsshop", idx, iErrStr)
			End If
			Call SugiOptionQueLogInsert("gsshop", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oGSShop = nothing
ElseIf action = "EDIT" Then
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
				'response.write "ERR||"&itemid&"||"&SumErrStr
				CALL Fn_AcctFailTouch("gsshop", idx, SumErrStr)
				Call SugiOptionQueLogInsert("gsshop", action, idx, "ERR", "ERR||"&idx&"||"&SumErrStr, session("ssBctID"))

				iErrStr = "ERR||"&idx&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_gsshopAddoption_regitem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " ,GSShopLastUpdate = getdate() " & VBCRLF
				strSql = strSql & " WHERE midx = '"&idx&"' "
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&idx&"||", "")
				'response.write "OK||"&itemid&"||"&SumOKStr
				Call SugiOptionQueLogInsert("gsshop", action, idx, "OK", "OK||"&idx&"||"&SumOKStr, session("ssBctID"))

				iErrStr = "OK||"&idx&"||"&SumOKStr
			End If
		End If
		'testURL : http://wapi.10x10.co.kr/outmall/proc/GSShopProc.asp?itemid=1044802&mallid=gsshop&action=EDIT
	SET oGSShop = nothing
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
'###################################################### GSShop API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->