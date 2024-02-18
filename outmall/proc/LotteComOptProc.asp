<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/lotteComAddOpt/lotteItemcls.asp"-->
<!-- #include virtual="/outmall/lotteComAddOpt/incLotteFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/lotteComAddOpt/inc_dailyAuthCheck.asp"-->
<%
Dim idx, mallid, action, oLotteitem, failCnt, chgSellYn, arrRows
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, mode
idx			= requestCheckVar(request("idx"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
mode			= request("mode")

If Not(isNumeric(idx)) Then
	response.write "<script>alert('�߸��� ��ǰ��ȣ�Դϴ�.')</script>"
	response.end
End If
'######################################################## LotteCom API ########################################################

If action = "SOLDOUT" Then								'���º���
	strParam = ""
	strParam = getLotteComSellynParameter("N", getLotteGoodno(idx))
	Call fnLotteComSellyn(idx, "N", strParam, iErrStr)
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchOption("lotteCom", idx, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/LotteComOptProc.asp?idx=1003&mallid=lotteCom&action=SOLDOUT
ElseIf action = "PRICE" Then								'���ݼ���
	strParam = ""
	strParam = getLotteComPriceParameter(idx, getLotteGoodno(idx), mustPrice)
	If strParam = "" Then
		response.write "ERR||"&idx&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
	Else
		Call fnLotteComPrice(idx, strParam, mustPrice, iErrStr)
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouchOption("lotteCom", idx, iErrStr)
		End If
	End If
	'http://wapi.10x10.co.kr/outmall/proc/LotteComOptProc.asp?idx=1003&mallid=lotteCom&action=PRICE
ElseIf action = "CHKSTAT" Then								'�űԻ�ǰ��ȸ
	Call fnLotteComStatChk(idx, iErrStr)
	response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchOption("lotteCom", idx, iErrStr)
	End If
	'http://wapi.10x10.co.kr/outmall/proc/LotteComOptProc.asp?idx=1003&mallid=lotteCom&action=CHKSTAT
ElseIf action = "EDIT" Then									'�����ȸ + ��ǰ���� + ���� + �ʿ信 ���� ��ǰ�ǸŻ��¼���
	SET oLotteitem = new CLotte
		oLotteitem.FRectIdx = idx
		oLotteitem.getLotteEditOneItem
		If oLotteitem.FResultCount > 0 Then
			If (oLotteitem.FOneItem.FmaySoldOut = "Y") OR (oLotteitem.FOneItem.IsOptionSoldOut) OR (oLotteitem.FOneItem.isDiffName) Then
				strParam = ""
				strParam = getLotteComSellynParameter("N", getLotteGoodno(idx))
				Call fnLotteComSellyn(idx, "N", strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (oLotteitem.FOneItem.FLotteSellYn = "N" AND oLotteitem.FOneItem.FmaySoldOut = "N" AND oLotteitem.FOneItem.IsOptionSoldOut = False) Then
					iErrStr = ""
					strParam = ""
					strParam = getLotteComSellynParameter("Y", getLotteGoodno(idx))
					Call fnLotteComSellyn(idx, "Y", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strParam = ""
				strParam = getLotteComPriceParameter(idx, getLotteGoodno(idx), mustPrice)
				If strParam = "" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & "ERR||"&itemid&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
				Else
					Call fnLotteComPrice(idx, strParam, mustPrice, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strParam = ""
				strParam = oLotteitem.FOneItem.getLotteComItemEditParameter()
				Call fnLotteComInfoEdit(idx, strParam, iErrStr, FALSE)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				If oLotteitem.FOneItem.isImageChanged Then
					strParam = ""
					strParam = oLotteitem.FOneItem.getLotteItemImageEdit()
					Call fnLotteComImageEdit(idx, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				'���û�ǰ ��ȸ�ؼ� ���� ��ǰ���¸� ��������
				Call fnCheckLotteComItemStat(idx, iErrStr, getLotteGoodno(idx))
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If


				'OK�� ERR�̴� editQuecnt�� + 1�� ��Ŵ..
				'�����ٸ����� editQuecnt ASC, i.lastupdate DESC�� �ߺ��� ����
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_lotteAddOption_regItem] SET " & VBCRLF
				strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
				strSql = strSql & " ,LotteLastUpdate = getdate()  " & VBCRLF
				strSql = strSql & " WHERE midx = '"&idx&"' " & VBCRLF
				dbget.Execute strSql
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&idx&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&idx&"||", "")
				CALL Fn_AcctFailTouchOption("lotteCom", idx, SumErrStr)
				response.write "ERR||"&idx&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_lotteAddOption_regItem] SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE midx = '"&idx&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&idx&"||", "")
				response.write "OK||"&idx&"||"&SumOKStr
			End If
		End If
	SET oLotteitem = nothing
	'http://wapi.10x10.co.kr/outmall/proc/LotteComOptProc.asp?idx=1003&mallid=lotteCom&action=EDIT
End If
'###################################################### LotteiMall API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->