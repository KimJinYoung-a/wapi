<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/ezwel/ezwelItemcls.asp"-->
<!-- #include virtual="/outmall/ezwel/incezwelFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, oEzwel, failCnt, chgSellYn, arrRows, skipItem, sellgubun, getMustprice, chkXML
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, optReset, optString, mode, ezwelGoodno
Dim jenkinsBatchYn, idx
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
chkXML			= request("chkXML")
failCnt			= 0

mode			= request("mode")
jenkinsBatchYn	= request("jenkinsBatchYn")
idx				= request("idx")

If mode = "updateSendState" Then
	Dim sqlStr, AssignedRow
	sqlStr = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr & "	Set sendState='"&request("updateSendState")&"'"
	sqlStr = sqlStr & "	,sendReqCnt=sendReqCnt+1"

	if (request("updateSendState") = "952") then
		'// ����ֹ��� �μ����۵� skip
		sqlStr = sqlStr & " , recvSendState = 100 "
		sqlStr = sqlStr & " , recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1 "
	end if

	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ord_no")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	sqlStr = sqlStr & "	and sellsite='ezwel'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
	response.end
Else
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
'######################################################## Ezwel API ########################################################
If mallid = "ezwel" Then
	If action = "SOLDOUT" Then													'���º���
		SET oEzwel = new cEzwel
			oEzwel.FRectItemID	= itemid
			oEzwel.getEzwelEditOneItem
			If (oEzwel.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				strParam = oEzwel.FOneItem.getEzwelItemRegXML("SellN", chkXML)
				getMustprice = ""
				getMustprice = oEzwel.FOneItem.fngetMustPrice()
				Call EzwelOneItemEditSellyn(itemid, oEzwel.FOneItem.FEzwelGoodNo, iErrStr, strParam, getMustprice, "N", "all", oEzwel.FOneItem.FLimityn, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitsold, chkXML)
				response.write iErrStr
				If LEFT(iErrStr, 2) <> "OK" Then
					CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
				End If
			End If
			'http://testwapi.10x10.co.kr/outmall/proc/EzwelProc.asp?itemid=699617&mallid=ezwel&action=SOLDOUT
		SET oEzwel = nothing
	ElseIf action = "REG" Then													'��ǰ���
		SET oEzwel = new cEzwel
			oEzwel.FRectItemID	= itemid
			oEzwel.getEzwelNotRegOneItem
		    If (oEzwel.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			ElseIf oEzwel.FOneItem.FdepthCode = "0" Then
				iErrStr = "ERR||"&itemid&"||ī�װ� ��Ī ���� Ȯ���ϼ���."
			Else
				strSql = ""
				strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_ezwel_regItem where itemid="&itemid&")"
				strSql = strSql & " BEGIN"& VbCRLF
				strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_ezwel_regItem "
		        strSql = strSql & " 	(itemid, regdate, reguserid, ezwelstatCD, regitemname)"
		        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oEzwel.FOneItem.FItemName)&"')"
				strSql = strSql & " END "
				dbget.Execute strSql

				'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
				If oEzwel.FOneItem.checkTenItemOptionValid Then
					strParam = ""
					strParam = oEzwel.FOneItem.getEzwelItemRegXML("Reg", chkXML)
					Call EzwelItemReg(itemid, strParam, iErrStr, oEzwel.FOneItem.FSellCash, oEzwel.FOneItem.getEzwelSellYn, oEzwel.FOneItem.FLimityn, oEzwel.FOneItem.FLimitNo, oEzwel.FOneItem.FLimitSold, html2db(oEzwel.FOneItem.FItemName), oEzwel.FOneItem.FbasicimageNm)
				Else
					iErrStr = "ERR||"&itemid&"||�ɼǰ˻� ����"
				End If
			End If
		SET oEzwel = nothing
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
	ElseIf action = "EDIT" OR action = "ITEMNAME" OR action = "PRICE" Then		'��ǰ����
		SET oEzwel = new cEzwel
			oEzwel.FRectItemID	= itemid
			oEzwel.getEzwelEditOneItem
			If (oEzwel.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||���������� ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				iErrStr = ""
				optReset = "N"
				optString = "all"
				'*********************************************************************************************************************************************************
				'2014-11-06 ������ | dev_Comment
				'API�� ���۵Ǵ� ���� ��ǰ�ɼ��� �ν����� ���� | ��ϵ� �ɼ�ī��Ʈ�� ũ�ٸ� 10x10���� �ɼ� ������ ���� �������
				'�ᱹ �������� �ɼǻ��������� ������ �ɼ��� �ʱ�ȭ ���� �߰�
				'�߰� : �ι� API���۽� ���� Ȯ���� ������ �� | �Ƹ� ��������� DB�� ��ǰ���� �����ϴ� �� ���� �ɷ��ִ� �� ��..
				'		���� �켱 �̷� ��ǰ�� ǰ����
				strSql = ""
				strSql = strSql &  "SELECT top 1 r.itemid, i.optioncnt, r.regedoptcnt "
				strSql = strSql & " FROM db_item.dbo.tbl_item as i "
				strSql = strSql & " join db_etcmall.dbo.tbl_ezwel_regitem as r on i.itemid=r.itemid "
				strSql = strSql & " WHERE i.itemid=" & itemid
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If not rsget.EOF Then
					If CInt(rsget("optioncnt")) > 0 Then
						If CInt(rsget("optioncnt")) <> CInt(rsget("regedoptcnt")) Then
							optReset = "Y"
							optString = "optMustN"
						End If
					End If
				End If
				rsget.Close

				If (oEzwel.FOneItem.FmaySoldOut = "Y") OR (oEzwel.FOneItem.IsSoldOutLimit5Sell) OR (optReset = "Y") OR (oEzwel.FOneItem.IsMayLimitSoldout = "Y") Then
					If optReset = "Y" Then
						strParam = oEzwel.FOneItem.getEzwelItemRegXML("MustNotOpt", chkXML)
					Else
						strParam = oEzwel.FOneItem.getEzwelItemRegXML("SellN", chkXML)
					End If
					chgSellYn = "N"
				Else
					strParam = oEzwel.FOneItem.getEzwelItemRegXML("SellY", chkXML)
					chgSellYn = "Y"
				End If

				getMustprice = ""
				getMustprice = oEzwel.FOneItem.fngetMustPrice()
				Call EzwelOneItemEdit(itemid, oEzwel.FOneItem.FEzwelGoodNo, iErrStr, strParam, getMustprice, chgSellYn, optString, oEzwel.FOneItem.FLimityn, oEzwel.FOneItem.FLimitno, oEzwel.FOneItem.FLimitsold, chkXML, oEzwel.FOneItem.FezwelSellYn)
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			If InStr(iErrStr, "[���Ǹ�]") = 0 Then
				Call EzwelItemChkstat(itemid, iErrStr, oEzwel.FOneItem.FEzwelGoodNo)
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
				CALL Fn_AcctFailTouch("ezwel", itemid, SumErrStr)
				iErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				iErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
			'http://testwapi.10x10.co.kr/outmall/proc/EzwelProc.asp?itemid=699617&mallid=ezwel&action=EDIT
			'http://testwapi.10x10.co.kr/outmall/proc/EzwelProc.asp?itemid=699617&mallid=ezwel&action=ITEMNAME
			'http://testwapi.10x10.co.kr/outmall/proc/EzwelProc.asp?itemid=699617&mallid=ezwel&action=PRICE
		SET oEzwel = nothing
	ElseIf action = "CHKSTAT" Then												'���� ��ȸ
		ezwelGoodno = getEzwelGoodno(itemid)
		If (ezwelGoodno = "") Then
			iErrStr = "ERR||"&itemid&"||��ȸ ������ ��ǰ�� �ƴմϴ�."
		Else
			Call EzwelItemChkstat(itemid, iErrStr, ezwelGoodno)
		End If
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ezwel", itemid, iErrStr)
		End If
	End If
End If
'###################################################### ezwel API END #######################################################
If jenkinsBatchYn = "Y" and iErrStr <> "" Then
	sqlStr = ""
	sqlStr = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(iErrStr, "||")(0)&"','"&html2DB(Split(iErrStr, "||")(2))&"'"
	dbget.Execute sqlStr
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->