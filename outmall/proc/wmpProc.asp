<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/wmp/wmpItemcls.asp"-->
<!-- #include virtual="/outmall/wmp/incWmpFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, mallid, action, failCnt, oWmp, getMustprice, chgImageNm, chgSellYn
Dim iErrStr, strParam, mustPrice, strSql, SumErrStr, SumOKStr, getOptSellValid, isOK
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
'######################################################## ������ API ########################################################
If mallid = "WMP" Then
	If action = "REG" Then					'��ǰ���
		SET oWmp = new CWmp
			oWmp.FRectItemID	= itemid
			oWmp.getWmpNotRegOneItem
			If (oWmp.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			Else
				strSql = "EXEC [db_etcmall].[dbo].[usp_API_Wemake_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
				dbget.execute strSql

				strSql = "SELECT db_etcmall.[dbo].[getWemakeAvailableString] ('"&itemid&"') as isOK"
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If not rsget.EOF Then
					isOK = rsget("isOK")
				End If
				rsget.Close

				'##�ɼ� �߰��ݾ��� �ְų� 3�ܿɼ� �̻��� ��� False
				If oWmp.FOneItem.checkTenItemOptionValid2 <> "True" Then
					iErrStr = "ERR||"&itemid&"||[��ǰ���] 3�ܿɼ� or �ɼ��߰��ݾ� �Ұ� or �ɼǰ���200�ʰ�"
				'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
				ElseIf oWmp.FOneItem.FoptionCnt > 0 AND oWmp.FOneItem.FLimitYN = "Y" AND oWmp.FOneItem.checkTenItemOptionUseValid <> "True" Then
					iErrStr = "ERR||"&itemid&"||[��ǰ���] �ɼ� �������� 5�� �̸�"
				ElseIf isOK = "N" Then
					iErrStr = "ERR||"&itemid&"||[��ǰ���] ��Ģ�� or �ɼ�Ÿ�Ա������� ��ϺҰ�"
				ElseIf oWmp.FOneItem.FinfoDiv = "" OR oWmp.FOneItem.FinfoDiv = "38" Then
					iErrStr = "ERR||"&itemid&"||[��ǰ���] ��ǰ ������� �׸� ����"
				ElseIf oWmp.FOneItem.checkTenItemOptionValid Then
					Call fnWmpItemReg(itemid, iErrStr)
				Else
					iErrStr = "ERR||"&itemid&"||[��ǰ���] �ɼǰ˻� ����"
				End If
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			If failCnt = 0 Then
				Call fnWmpStatCheck(itemid, iErrStr, oWmp.FOneItem.FLimitYN)
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
				CALL Fn_AcctFailTouch("WMP", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_wemake_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oWmp = nothing
		'http://wapi.10x10.co.kr/outmall/proc/wmpProc.asp?itemid=325046&mallid=WMP&action=REG
	ElseIf action = "SOLDOUT" Then			'���º���
		SET oWmp = new CWmp
			oWmp.FRectItemID	= itemid
			oWmp.getWmpEditSaleOneItem

			If (oWmp.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||[���¼���] ���� ������ ��ǰ�� �ƴմϴ�."
			Else
				getMustprice = ""
				getMustprice = oWmp.FOneItem.FMustPrice
				Call fnWmpSellyn(itemid, iErrStr, getMustprice, oWmp.FOneItem.FStockCount, "N")
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("WMP", itemid, iErrStr)
			End If
		SET oWmp = nothing
		'http://wapi.10x10.co.kr/outmall/proc/wmpProc.asp?itemid=325046&mallid=WMP&action=SOLDOUT
	ElseIf action = "CHKSTAT" Then			'��ȸ
		SET oWmp = new CWmp
			oWmp.FRectItemID	= itemid
			oWmp.getWmpEditSaleOneItem

			If (oWmp.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||[��ȸ] ��ȸ ������ ��ǰ�� �ƴմϴ�."
			Else
				Call fnWmpStatCheck(itemid, iErrStr, oWmp.FOneItem.FLimitYN)
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("WMP", itemid, iErrStr)
			End If
		SET oWmp = nothing
		'http://wapi.10x10.co.kr/outmall/proc/wmpProc.asp?itemid=325046&mallid=WMP&action=CHKSTAT
	ElseIf action = "PRICE" Then		'���� ����
		SET oWmp = new CWmp
			oWmp.FRectItemID	= itemid
			oWmp.getWmpEditSaleOneItem

			If (oWmp.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||[���ݼ���] ���� ������ ��ǰ�� �ƴմϴ�."
			Else
				getMustprice = ""
				getMustprice = oWmp.FOneItem.FMustPrice

				getOptSellValid = true
				If oWmp.FOneItem.FLimitYN = "Y" Then
					getOptSellValid = oWmp.FOneItem.checkTenItemOptionUseValid
				End If
				Call fnWmpPrice(itemid, iErrStr, getMustprice, oWmp.FOneItem.FStockCount, getOptSellValid)
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("WMP", itemid, iErrStr)
			End If
		SET oWmp = nothing
		'http://wapi.10x10.co.kr/outmall/proc/wmpProc.asp?itemid=325046&mallid=WMP&action=PRICE
	ElseIf action = "EDIT" Then		'��ǰ ����
		SET oWmp = new CWmp
			oWmp.FRectItemID	= itemid
			oWmp.getWmpEditOneItem
			If oWmp.FResultCount = 0 Then
				iErrStr = "ERR||"&itemid&"||���� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
			Else
				getMustprice = ""
				getMustprice = oWmp.FOneItem.FMustPrice
				If (oWmp.FOneItem.FmaySoldOut = "Y") OR (oWmp.FOneItem.IsMayLimitSoldout = "Y") OR (oWmp.FOneItem.checkTenItemOptionValid2 <> "True") Then
					Call fnWmpSellyn(itemid, iErrStr, getMustprice, oWmp.FOneItem.FStockCount, "N")
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
				'############## ������ ��ǰ ���� #################
					Call fnWmpItemEdit(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

				'############## ������ ��ǰ ��ȸ #################
					If failCnt = 0 Then
						Call fnWmpStatCheck(itemid, iErrStr, oWmp.FOneItem.FLimitYN)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

				'############## ������ ��ǰ �Ǹ� #################
					If failCnt = 0 Then
						Call fnWmpSellyn(itemid, iErrStr, getMustprice, oWmp.FOneItem.FStockCount, "Y")
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				End If
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("WMP", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_wemake_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oWmp = nothing
		'http://wapi.10x10.co.kr/outmall/proc/wmpProc.asp?itemid=325046&mallid=WMP&action=EDIT
	End If
End If
'###################################################### ������ API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->