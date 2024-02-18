<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 60 * 15
%>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/hmall/hmallItemcls.asp"-->
<!-- #include virtual="/outmall/hmall/inchmallFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, oHmall, failCnt, chgSellYn, arrRows, isItemIdChk, maeipdiv, mustPrice, getMustprice
Dim iErrStr, strSql, SumErrStr, SumOKStr, i, tHmallGoodno, isChkStat, strparam, mrgnRate, chgImageNm, endItemErrMsgReplace
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
failCnt			= 0

If action <> "SECTView" and action <> "infoDivView" Then
	If isItemIdChk = "Y" Then
		If itemid="" or itemid="0" Then
			response.write "<script>alert('��ǰ��ȣ�� �����ϴ�.')</script>"
			response.end
		ElseIf Not(isNumeric(itemid)) Then
			response.write "<script>alert('�߸��� ��ǰ��ȣ�Դϴ�.')</script>"
			response.end
		Else
			'�������·� ��ȯ
			itemid = CLng(getNumeric(itemid))
		End If
	End If
End If

'######################################################## HMall API ########################################################
If action = "REG" Then									'��ǰ���
'	SET oHmall = new CHMall
'		oHmall.FRectItemID	= itemid
'		oHmall.getHmallNotRegOneItem
'	    If (oHmall.FResultCount < 1) Then
'			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
'		Else
'			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
'			dbget.execute strSql
'
'			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
'			If oHmall.FOneItem.checkTenItemOptionValid Then
'				Call fnHmallItemReg(itemid, iErrStr)
'			Else
'				iErrStr = "ERR||"&itemid&"||[��ǰ���] �ɼǰ˻� ����"
'			End If
'		End If
'		If LEFT(iErrStr, 2) <> "OK" Then
'			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
'		End If
'		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
'	SET oHmall = nothing
'###################################################	������� �� ���� ##########################################
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		'oHmall.getHmallNotRegOneItem
		oHmall.getHmallNotRegOnlyOneItem
		If oHmall.FResultCount > 0 Then
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
			dbget.execute strSql

			'If oHmall.FOneItem.fnCheckMakerid Then
			'	iErrStr = "ERR||"&itemid&"||[��ǰ���add] ���̰��� ��ϺҰ�"
			If oHmall.FOneItem.checkTenItemOptionValid Then
'				Call fnHmallOnlyItemReg(itemid, iErrStr)
				strParam = ""
				strParam = oHmall.FOneItem.gethmallItemRegParameter()

				getMustprice = ""
				getMustprice = oHmall.FOneItem.MustPrice()
				Call fnHmallItemOnlyReg(itemid, strParam, iErrStr, getMustprice, oHmall.FOneItem.gethmallSellYn, oHmall.FOneItem.FLimityn, oHmall.FOneItem.FLimitNo, oHmall.FOneItem.FLimitSold, html2db(oHmall.FOneItem.FItemName), oHmall.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[��ǰ���add] �ɼǰ˻� ����"
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			If failCnt = 0 Then
				tHmallGoodno = getHmallGoodno(itemid)
				If tHmallGoodno <> "" Then
					chgImageNm = oHmall.FOneItem.getBasicImage
					Call fnHmallImage(itemid, chgImageNm, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			End If

			If failCnt = 0 Then
				tHmallGoodno = getHmallGoodno2(itemid)
				If tHmallGoodno <> "" Then
					Call fnHmallImageConfirm(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			End If
		Else
			failCnt = 1
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
			dbget.execute strSql
			SumErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
			Call SugiQueLogInsert("hmall1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("hmall1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oHmall = nothing
ElseIf action = "REGAddItem" Then						'��ǰ�� ���
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallNotRegOneItem
	    If (oHmall.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
			dbget.execute strSql

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oHmall.FOneItem.checkTenItemOptionValid Then
				Call fnHmallOnlyItemReg(itemid, iErrStr)
			Else
				iErrStr = "ERR||"&itemid&"||[��ǰ���] �ɼǰ˻� ����"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
ElseIf action = "IMAGE" Then							'�̹��� ��� & Ȯ��
	tHmallGoodno = getHmallGoodno(itemid)
	If tHmallGoodno = "" Then
		failCnt = 1
		SumErrStr = "ERR||"&itemid&"||��ǰ���� ��� �ϼž� �˴ϴ�."
	Else
		chgImageNm = getTenBasicImage(itemid)
		Call fnHmallImage(itemid, chgImageNm, iErrStr)
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If

	If failCnt = 0 Then
		tHmallGoodno = getHmallGoodno2(itemid)
		If tHmallGoodno = "" Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||��ǰ �� �̹������� ��� �ϼž� �˴ϴ�."
		Else
			Call fnHmallImageConfirm(itemid, iErrStr)
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
		CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
		Call SugiQueLogInsert("hmall1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("hmall1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "REGImage" Then							'�̹��� ���
	tHmallGoodno = getHmallGoodno(itemid)
	If tHmallGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||��ǰ���� ��� �ϼž� �˴ϴ�."
	Else
		chgImageNm = getTenBasicImage(itemid)
		Call fnHmallImage(itemid, chgImageNm, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "REGImageConfirm" Then					'�̹��� Ȯ��
	tHmallGoodno = getHmallGoodno2(itemid)
	If tHmallGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||��ǰ �� �̹������� ��� �ϼž� �˴ϴ�."
	Else
		Call fnHmallImageConfirm(itemid, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "EditSellYn" Then						'��ǰ ���� ����
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount = 0 Then
			iErrStr = "ERR||"&itemid&"||���¼��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
		Else
			Call fnHmallSellYN(itemid, chgSellYn, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
ElseIf action = "PRICE" Then							'��ǰ ���� ����
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount > 0 Then
			mustPrice = ""
			mustPrice = oHmall.FOneItem.MustPrice()

			mrgnRate = ""
			mrgnRate = oHmall.FOneItem.FMrgnRate
			strParam = oHmall.FOneItem.getHmallPriceParameter()
			Call fnHmallPrice(itemid, mustPrice, mrgnRate, strParam, iErrStr)
		Else
			iErrStr = "ERR||"&itemid&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
' ElseIf action = "CHKSTAT" Then							'��ǰ �� ��ȸ	+ ��� ��ȸ			.NET ���� / �ɼ��� ������ �ʹ�
' 	SET oHmall = new CHMall
' 		oHmall.FRectItemID	= itemid
' 		oHmall.getHmallEditOneItem
' 		If oHmall.FResultCount > 0 Then
' 			Call fnHmallStatChk(itemid, iErrStr)
' 			If Left(iErrStr, 2) <> "OK" Then
' 				failCnt = failCnt + 1
' 				SumErrStr = SumErrStr & iErrStr
' 			Else
' 				SumOKStr = SumOKStr & iErrStr
' 			End If

' 			If INSTR(iErrStr, "���οϷ�") > 0 AND failCnt = 0 Then
' 				Call fnHmallOptionStatCheck(itemid, iErrStr)
' 				If Left(iErrStr, 2) <> "OK" Then
' 					failCnt = failCnt + 1
' 					SumErrStr = SumErrStr & iErrStr
' 				Else
' 					SumOKStr = SumOKStr & iErrStr
' 				End If

' 				strparam = oHmall.FOneItem.fngetOptionEditParam(itemid)
' 				Call fnHmallOptionEdit(itemid, strparam, iErrStr)
' 				If Left(iErrStr, 2) <> "OK" Then
' 					failCnt = failCnt + 1
' 					SumErrStr = SumErrStr & iErrStr
' 				Else
' 					SumOKStr = SumOKStr & iErrStr
' 				End If

' 				Call fnHmallOptionStatCheck(itemid, iErrStr)
' 				If Left(iErrStr, 2) <> "OK" Then
' 					failCnt = failCnt + 1
' 					SumErrStr = SumErrStr & iErrStr
' 				Else
' 					SumOKStr = SumOKStr & iErrStr
' 				End If
' 			End If
' 		Else
' 			failCnt = 1
' 			SumErrStr = "ERR||"&itemid&"||����ȸ �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
' 		End If

' 		If failCnt > 0 Then
' 			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
' 			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
' 			CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
' 			Call SugiQueLogInsert("hmall1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

' 			iErrStr = "ERR||"&itemid&"||"&SumErrStr
' 		Else
' 			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
' 			Call SugiQueLogInsert("hmall1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
' 			iErrStr = "OK||"&itemid&"||"&SumOKStr
' 		End If
' 	SET oHmall = nothing
ElseIf action = "OPTSTAT" Then							'��ǰ ��� ��ȸ
	'Call fnHmallOptionStatChk(itemid, iErrStr)
	Call fnHmallOptionStatCheck(itemid, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "OPTEDIT" Then							'��ǰ �ɼ� ����
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount > 0 Then
			strparam = oHmall.FOneItem.fngetOptionEditParam(itemid)
			Call fnHmallOptionEdit(itemid, strparam, iErrStr)
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			If failCnt = 0 Then
				Call fnHmallOptionStatCheck(itemid, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		Else
			failCnt = 1
			SumErrStr = "ERR||"&itemid&"||�ɼǼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
			Call SugiQueLogInsert("hmall1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("hmall1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
ElseIf action = "EDITItem" Then							'��ǰ ������ ����
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount > 0 Then
			Call fnHmallOnlyItemEdit(itemid, iErrStr)
		Else
			iErrStr = "ERR||"&itemid&"||��ǰ���� ���� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
ElseIf action = "REGOnly" Then						'��ǰ�� ���
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallNotRegOnlyOneItem
	    If (oHmall.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
			dbget.execute strSql

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oHmall.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oHmall.FOneItem.gethmallItemRegParameter()

				getMustprice = ""
				getMustprice = oHmall.FOneItem.MustPrice()
				Call fnHmallItemOnlyReg(itemid, strParam, iErrStr, getMustprice, oHmall.FOneItem.gethmallSellYn, oHmall.FOneItem.FLimityn, oHmall.FOneItem.FLimitNo, oHmall.FOneItem.FLimitSold, html2db(oHmall.FOneItem.FItemName), oHmall.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[��ǰ���] �ɼǰ˻� ����"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
	'http://localhost:11117/outmall/hmall/hmallActProc.asp?act=REGOnly&itemid=3887719
ElseIf action = "EDITonly" Then							'��ǰ ������ ����
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount > 0 Then
			strParam = oHmall.FOneItem.gethmallItemEditParameter()
			Call fnHmallItemOnlyEdit(itemid, strParam, iErrStr)
		Else
			iErrStr = "ERR||"&itemid&"||��ǰ���� ���� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
	'http://localhost:11117/outmall/hmall/hmallActProc.asp?act=EDITonly&itemid=2952364
ElseIf action = "CHKSTAT" Then							'��ǰ �� ��ȸ	+ ��� ��ȸ			.ASP����(����ó����) / 2022-01-05 �߰�
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount > 0 Then
			strParam = ""
			strParam = oHmall.FOneItem.getHmallItemConfirmParameter()
			Call fnHmallStatChk2(itemid, strParam, iErrStr)
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			If INSTR(iErrStr, "���οϷ�") > 0 AND failCnt = 0 Then
				Call fnHmallOptionStatCheck(itemid, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				strparam = oHmall.FOneItem.fngetOptionEditParam(itemid)
				Call fnHmallOptionEdit(itemid, strparam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				Call fnHmallOptionStatCheck(itemid, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		Else
			failCnt = 1
			SumErrStr = "ERR||"&itemid&"||����ȸ �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
			Call SugiQueLogInsert("hmall1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("hmall1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oHmall = nothing
ElseIf action = "EDIT" Then								'��ǰ ����
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount = 0 Then
			iErrStr = "ERR||"&itemid&"||���� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
		Else
			'If (oHmall.FOneItem.FmaySoldOut = "Y") OR (oHmall.FOneItem.IsMayLimitSoldout = "Y") OR (oHmall.FOneItem.IsAllOptionChange = "Y") OR (oHmall.FOneItem.fnCheckMakerid) Then
            If (oHmall.FOneItem.FmaySoldOut = "Y") OR (oHmall.FOneItem.IsMayLimitSoldout = "Y") Then
				Call fnHmallSellYN(itemid, "N", iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
			'############## Hmall ��ǰ ���� #################
'2022-05-09 ������ �ϴ� ����
'				Call fnHmallOnlyItemEdit(itemid, iErrStr)
'				If Left(iErrStr, 2) <> "OK" Then
'					failCnt = failCnt + 1
'					SumErrStr = SumErrStr & iErrStr
'				Else
'					SumOKStr = SumOKStr & iErrStr
'				End If
				strParam = ""
				strParam = oHmall.FOneItem.gethmallItemEditParameter()
				Call fnHmallItemOnlyEdit(itemid, strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			'############## Hmall �̹��� ���� #################
				If oHmall.FOneItem.isImageChanged Then
					chgImageNm = oHmall.FOneItem.getBasicImage
					Call fnHmallImage(itemid, chgImageNm, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					Call fnHmallImageConfirm(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

			'############## Hmall ���� ���� #################
				If failCnt = 0 Then
					mustPrice = ""
					mustPrice = oHmall.FOneItem.MustPrice()

					mrgnRate = ""
					mrgnRate = oHmall.FOneItem.FMrgnRate
					strParam = ""
					strParam = oHmall.FOneItem.getHmallPriceParameter()
					Call fnHmallPrice(itemid, mustPrice, mrgnRate, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

			'############## Hmall �ɼ� ���� #################
				If failCnt = 0 Then
					strparam = oHmall.FOneItem.fngetOptionEditParam(itemid)
					Call fnHmallOptionEdit(itemid, strparam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

			'############## Hmall ��� ��ȸ #################
				If failCnt = 0 Then
					Call fnHmallOptionStatCheck(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

			'############## Hmall �Ǹ� ���� ���� #################
				If failCnt = 0 Then
					Call fnHmallSellYN(itemid, "Y", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				endItemErrMsgReplace = replace(SumErrStr, "OK||"&itemid&"||", "")
				endItemErrMsgReplace = replace(SumErrStr, "ERR||"&itemid&"||", "")

				If (oHmall.FOneItem.IsAllOptionChange = "Y") OR (Instr(endItemErrMsgReplace, "�ǸŰ����� �Ӽ� ������ �����ϴ�") > 0) OR (Instr(endItemErrMsgReplace, "�ǸŰ����ѼӼ������������ϴ�") > 0) Then
					Call fnHmallSellYN(itemid, "N", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					strSql = "	DECLARE @Temp CHAR(1) " & _
								"	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = 'hmall1010' AND itemid = '"& itemid &"') " & _
								"		BEGIN " & _
								"			INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid(itemid,mallgubun,bigo) VALUES('"& itemid &"','hmall1010', '�ɼ���ü����ǰ��(system)') " & _
								"		END	"
					dbget.execute strSql
				End If
			End If
		End If
	SET oHmall = nothing

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
		Call SugiQueLogInsert("hmall1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regItem SET " & VBCRLF
		strSql = strSql & " accFailcnt = 0  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("hmall1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "SECTView" Then
	Call fnHmallSectView()
	'http://localhost:11117/outmall/hmall/hmallActProc.asp?act=SECTView
ElseIf action = "infoDivView" Then
	Call fnHmallInfoDivView()
	'http://localhost:11117/outmall/hmall/hmallActProc.asp?act=infoDivView
End If

response.write  "<script>" & vbCrLf &_
				"	var str, t; " & vbCrLf &_
				"	t = parent.document.getElementById('actStr') " & vbCrLf &_
				"	str = t.innerHTML; " & vbCrLf &_
				"	str += '"&iErrStr&"<br>' " & vbCrLf &_
				"	t.innerHTML = str; " & vbCrLf &_
				"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
				"</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->