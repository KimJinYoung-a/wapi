<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/sabangnet/sabangnetItemcls.asp"-->
<!-- #include virtual="/outmall/sabangnet/incSabangnetFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, failCnt, oSabangnet, getMustprice, chgSellYn, vOptCnt
Dim iErrStr, strParam, mustPrice, strSql, SumErrStr, SumOKStr, chgImageNm
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
'######################################################## Sabangnet API ########################################################
If mallid = "sabangnet" Then
	If action = "SOLDOUT" Then				'���º���
		SET oSabangnet = new CSabangnet
			oSabangnet.FRectItemID	= itemid
			oSabangnet.getSabangnetSimpleEditOneItem
		    If (oSabangnet.FResultCount < 1) Then
				lastErrStr = "ERR||"&itemid&"||[����������] ���� ������ ��ǰ�� �ƴմϴ�."
				response.write "ERR||"&itemid&"||[����������] ���� ������ ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				strParam = oSabangnet.FOneItem.getSabangnetSimpleEditItemParameter("N")
				Call fnSabangnetSimpleEdit(itemid, "N", oSabangnet.FOneItem.MustPrice, html2db(oSabangnet.FOneItem.FItemName), strParam, iErrStr, "sellyn")
				lastErrStr = iErrStr
				response.write iErrStr
			End If
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("sabangnet", itemid, iErrStr)
			End If
		SET oSabangnet = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/sabangnetProc.asp?itemid=1649348&mallid=sabangnet&action=SOLDOUT
	ElseIf action = "PRICE" Then			'���ݼ���
		SET oSabangnet = new CSabangnet
			oSabangnet.FRectItemID	= itemid
			oSabangnet.getSabangnetSimpleEditOneItem
		    If (oSabangnet.FResultCount < 1) Then
				lastErrStr = "ERR||"&itemid&"||[����������] ���� ������ ��ǰ�� �ƴմϴ�."
				response.write "ERR||"&itemid&"||[����������] ���� ������ ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				If (oSabangnet.FOneItem.FmaySoldOut = "Y") OR (oSabangnet.FOneItem.IsMayLimitSoldout = "Y") OR (oSabangnet.FOneItem.IsSoldOut) Then
					chgSellYn = "N"
					strParam = oSabangnet.FOneItem.getSabangnetSimpleEditItemParameter(chgSellYn)
				Else
					chgSellYn = "Y"
					strParam = oSabangnet.FOneItem.getSabangnetSimpleEditItemParameter(chgSellYn)
				End If

				Call fnSabangnetSimpleEdit(itemid, chgSellYn, oSabangnet.FOneItem.MustPrice, html2db(oSabangnet.FOneItem.FItemName), strParam, iErrStr, "price")
				lastErrStr = iErrStr
				response.write iErrStr
			End If
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("sabangnet", itemid, iErrStr)
			End If
		SET oSabangnet = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/sabangnetProc.asp?itemid=1649348&mallid=sabangnet&action=PRICE
	ElseIf action = "EDIT" Then				'��ǰ����
		SET oSabangnet = new CSabangnet
			oSabangnet.FRectItemID	= itemid
			oSabangnet.getSabangnetEditOneItem

		    If (oSabangnet.FResultCount < 1) Then
				lastErrStr = "ERR||"&itemid&"||[��ü����] ���� ������ ��ǰ�� �ƴմϴ�."
				response.write "ERR||"&itemid&"||[��ü����] ���� ������ ��ǰ�� �ƴմϴ�."
			Else
				If (oSabangnet.FOneItem.FmaySoldOut = "Y") OR (oSabangnet.FOneItem.IsMayLimitSoldout = "Y") OR (oSabangnet.FOneItem.IsSoldOut) Then
					chgSellYn = "N"
				Else
					chgSellYn = "Y"
				End If
				strParam = ""
				strParam = oSabangnet.FOneItem.getSabangnetItemRegParameter(True, chgSellYn)
				Call fnSabangnetItemEdit(itemid, strParam, iErrStr, oSabangnet.FOneItem.MustPrice, chgImageNm, oSabangnet.FOneItem.FLimityn, oSabangnet.FOneItem.FLimitno, oSabangnet.FOneItem.FLimitsold, chgSellYn)
				lastErrStr = iErrStr
				response.write iErrStr
			End If
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("sabangnet", itemid, iErrStr)
			End If
		SET oSabangnet = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/sabangnetProc.asp?itemid=1649348&mallid=sabangnet&action=PRICE
	End If
End If
'###################################################### Sabangnet API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->