<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/lotteCom/lotteItemcls.asp"-->
<!-- #include virtual="/outmall/lotteCom/incLotteFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/lotteCom/inc_dailyAuthCheck.asp"-->
<%
Dim itemid, mallid, action, oLotteitem, failCnt, chgSellYn, arrRows, skipItem, assin, isMayEndItem, isMayEndItem2
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, mode, tLotteGoodno
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
mode			= request("mode")

If mode = "updateSendState" Then
	strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendState='"&request("updateSendState")&"'"
	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
	strSql = strSql & "	where OutMallOrderSerial='"&request("ORG_ord_no")&"'"
	strSql = strSql & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	dbget.Execute strSql,assin
	response.write "<script>alert('"&assin&"�� �Ϸ� ó��.');opener.close();window.close()</script>"
	response.end
ElseIf mode = "etcSongjangFin" Then
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendState=7"
	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
    strSql = strSql & "	where OutMallOrderSerial='"&request("ORG_ord_no")&"'"
    strSql = strSql & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
    dbget.Execute strSql,assin
    response.write "<script>alert('"&assin&"�� �Ϸ� ó��.');opener.close();window.close()</script>"
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
'######################################################## LotteCom API ########################################################

If action = "SOLDOUT" Then								'���º���
	strParam = ""
	strParam = getLotteComSellynParameter("N", getLotteGoodno(itemid))
	Call fnLotteComSellyn(itemid, "N", strParam, iErrStr)
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/LotteComProc.asp?itemid=279397&mallid=lotteCom&action=SOLDOUT
ElseIf action = "PRICE" Then								'���ݼ���
	SET oLotteitem = new CLotte
		oLotteitem.FRectItemID	= itemid
		oLotteitem.getLotteEditOneItem
		If oLotteitem.FResultCount > 0 Then
			strParam = ""
			mustPrice = ""
			mustPrice = oLotteitem.FOneItem.MustPrice()
			strParam = getLotteComPriceParameter(itemid, getLotteGoodno(itemid), mustPrice)
			If strParam = "" Then
				response.write "ERR||"&itemid&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
			Else
				Call fnLotteComPrice(itemid, strParam, mustPrice, iErrStr)
				response.write iErrStr
			End If
		else
			response.write "ERR||"&itemid&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�.[1]"
		end if
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
		End If
	SET oLotteitem = nothing
	'http://wapi.10x10.co.kr/outmall/proc/LotteComProc.asp?itemid=279397&mallid=lotteCom&action=PRICE
ElseIf action = "ITEMNAME" Then								'��ǰ�����
	strParam = ""
	strParam = getLotteItemnameParameter(itemid, iitemname, getLotteGoodno(itemid))
	Call fnLotteComChgItemname(itemid, strParam, iErrStr)
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/LotteComProc.asp?itemid=279397&mallid=lotteCom&action=ITEMNAME
ElseIf action = "CHKSTAT" Then								'�űԻ�ǰ��ȸ
	Call fnLotteComStatChk(itemid, iErrStr)
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/LotteComProc.asp?itemid=279397&mallid=lotteCom&action=CHKSTAT
ElseIf action = "CHKSTOCK" Then								'�����ȸ
	Call fnLotteComStockChk(itemid, iErrStr)
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/LotteComProc.asp?itemid=279397&mallid=lotteCom&action=CHKSTOCK
ElseIf action = "EDIT" Then									'�����ȸ + ��ǰ���� + ���� + �ʿ信 ���� ��ǰ�ǸŻ��¼���
	SET oLotteitem = new CLotte
		oLotteitem.FRectItemID	= itemid
		oLotteitem.getLotteEditOneItem
		If oLotteitem.FResultCount > 0 Then
			'1. ǰ���� �ش��ϸ� ǰ��ó��
			If (oLotteitem.FOneItem.FmaySoldOut = "Y") OR (oLotteitem.FOneItem.IsSoldOutLimit5Sell) OR (oLotteitem.FOneItem.IsMayLimitSoldout = "Y") Then
				strParam = ""
				strParam = getLotteComSellynParameter("N", getLotteGoodno(itemid))
				Call fnLotteComSellyn(itemid, "N", strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				'2. ��� ��ȸ
				Call fnLotteComStockChk(itemid, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'3. ��� ��ȸ�� �ǸŰ� �Ұ����ϸ� �Ǹ�����ó��
				isMayEndItem = getOptCntCompare(itemid)
				If isMayEndItem = "Y" Then
					strParam = ""
					strParam = getLotteComSellynParameter("X", getLotteGoodno(itemid))

					Call fnLotteComSellyn(itemid, "X", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					'3-1 �ǸŰ� ������ ��ǰ�̸� ��ǰ ����
					strParam = ""
					strParam = oLotteitem.FOneItem.getLotteComItemEditParameter()
					Call fnLotteComInfoEdit(itemid, strParam, iErrStr, FALSE)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'3-2 �Ǹŷ� ����
					strParam = ""
					strParam = getLotteComSellynParameter("Y", getLotteGoodno(itemid))
					Call fnLotteComSellyn(itemid, "Y", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'4. �ǸŰ� ����
					strParam = ""
					mustPrice = ""
					mustPrice = oLotteitem.FOneItem.MustPrice()
					strParam = getLotteComPriceParameter(itemid, getLotteGoodno(itemid), mustPrice)
					If strParam = "" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & "ERR||"&itemid&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
					Else
						Call fnLotteComPrice(itemid, strParam, mustPrice, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

					'5. �̹��� ����
					If oLotteitem.FOneItem.isImageChanged Then
						strParam = ""
						strParam = oLotteitem.FOneItem.getLotteItemImageEdit()
						Call fnLotteComImageEdit(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

					'6. ��� ��ȸ
					Call fnLotteComStockChk(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					isMayEndItem2 = getUseOption(itemid)
					If isMayEndItem2 = "N" Then
						strParam = ""
						strParam = getLotteComSellynParameter("X", getLotteGoodno(itemid))

						Call fnLotteComSellyn(itemid, "X", strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					Else
						'���û�ǰ ��ȸ�ؼ� ���� ��ǰ���¸� ��������
						Call fnCheckLotteComItemStat(itemid, iErrStr, getLotteGoodno(itemid))
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
					strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem SET " & VBCRLF
					strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
					strSql = strSql & " ,LotteLastUpdate = getdate()  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql
				End If
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("lotteCom", itemid, SumErrStr)
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				response.write "OK||"&itemid&"||"&SumOKStr
			End If

		End If
	SET oLotteitem = nothing
ElseIf action = "REG" Then									'��ǰ ���
	SET oLotteitem = new CLotte
		oLotteitem.FRectItemID	= itemid
		oLotteitem.getLotteNotRegOneItem

		tLotteGoodno = getLotteGoodno(itemid)
		If tLotteGoodno <> "" Then
			iErrStr = "ERR||"&itemid&"||�̹� ��ϵ� ��ǰ �Դϴ�."
	    ElseIf (oLotteitem.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_lotte_regItem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_item.dbo.tbl_lotte_regItem "
	        strSql = strSql & " 	(itemid, regdate, reguserid, LotteStatCd)"
	        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '10')"
			strSql = strSql & " END "
		    dbget.Execute strSql
			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oLotteitem.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oLotteitem.FOneItem.getLotteComItemRegParameter(FALSE)
				Call fnLotteComItemReg(itemid, strParam, iErrStr, oLotteitem.FOneItem.FSellCash, oLotteitem.FOneItem.getLotteSellYn, oLotteitem.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||�ɼǰ˻� ����"
			End If
		End If
	SET oLotteitem = nothing
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("lotteCom", itemid, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/LotteComProc.asp?itemid=1860480&mallid=lotteCom&action=REG
End If
'###################################################### LotteCom API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
