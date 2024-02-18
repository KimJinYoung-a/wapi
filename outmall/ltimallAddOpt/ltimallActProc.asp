<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/ltimallAddOpt/LtimallItemcls.asp"-->
<!-- #include virtual="/outmall/ltimallAddOpt/incLtimallFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/ltimallAddOpt/inc_dailyAuthCheck.asp"-->
<%
Dim idx, action, oiMall, failCnt, chgSellYn, arrRows, skipItem
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, iAddOptCnt
idx				= requestCheckVar(request("idx"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
failCnt			= 0
If Not(isNumeric(idx)) Then
	response.write "<script>alert('�߸��� ��ǰ��ȣ�Դϴ�.')</script>"
	response.end
End If

'######################################################## LotteiMall API ########################################################
If action = "EditSellYn" Then								'���º���
	strParam = ""
	strParam = getLtiMallSellynParameter(chgSellYn, getLtimallGoodno(idx))
	Call fnLtiMallSellyn(idx, chgSellYn, strParam, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchOption("lotteimall", idx, iErrStr)
	End If
	Call SugiOptionQueLogInsert("lotteimall", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "PRICE" Then								'���ݼ���
	strParam = ""
	strParam = getLtiMallPriceParameter(idx, getLtimallGoodno(idx), mustPrice)
	If strParam = "" Then
		iErrStr = "ERR||"&idx&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
	Else
		Call fnLtimallPrice(idx, strParam, mustPrice, iErrStr)
		'response.write iErrStr
	End If

	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchOption("lotteimall", idx, iErrStr)
	End If
	Call SugiOptionQueLogInsert("lotteimall", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "ITEMNAME" Then								'��ǰ�����
	strParam = ""
	strParam = getLtiMallItemnameParameter(idx, iitemname, getLtimallGoodno(idx))
	Call fnLtiMallChgItemname(idx, strParam, iErrStr)
	'response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchOption("lotteimall", idx, iErrStr)
	End If
	Call SugiOptionQueLogInsert("lotteimall", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKSTAT" Then								'�űԻ�ǰ��ȸ
	Call fnLtiMallstatChk(idx, iErrStr)
	'response.write iErrStr
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchOption("lotteimall", idx, iErrStr)
	End If
	Call SugiOptionQueLogInsert("lotteimall", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDIT" Then									'�����ȸ + ��ǰ���� + ���� + �ʿ信 ���� ��ǰ�ǸŻ��¼���(2015-10-06 ������ ���û�����ȸ �ּ�ó��)
	SET oiMall = new CLotteiMall
		oiMall.FRectIdx = idx
		oiMall.getLtimallEditOneItem
		If oiMall.FResultCount > 0 Then
			If (oiMall.FOneItem.FmaySoldOut = "Y") OR (oiMall.FOneItem.IsOptionSoldOut) OR (oiMall.FOneItem.isDiffName) Then
				strParam = ""
				strParam = getLtiMallSellynParameter("N", getLtimallGoodno(idx))
				Call fnLtiMallSellyn(idx, "N", strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (oiMall.FOneItem.FLtimallSellYn = "N" AND oiMall.FOneItem.FmaySoldOut = "N" AND oiMall.FOneItem.IsOptionSoldOut = False) Then
					iErrStr = ""
					strParam = ""
					strParam = getLtiMallSellynParameter("Y", getLtimallGoodno(idx))
					Call fnLtiMallSellyn(idx, "Y", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strParam = ""
				strParam = getLtiMallPriceParameter(idx, getLtimallGoodno(idx), mustPrice)
				If strParam = "" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & "ERR||"&idx&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
				Else
					Call fnLtimallPrice(idx, strParam, mustPrice, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strParam = ""
				strParam = oiMall.FOneItem.getLotteiMallItemEditParameter()
				Call fnLtiMallInfoEdit(idx, strParam, iErrStr, FALSE)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'���û�ǰ ��ȸ�ؼ� ���� ��ǰ���¸� ��������
				Call fnLtiMallDisView(idx, iErrStr, getLtimallGoodno(idx))
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'OK�� ERR�̴� editQuecnt�� + 1�� ��Ŵ..
				'�����ٸ����� editQuecnt ASC, i.lastupdate DESC�� �ߺ��� ����
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] SET " & VBCRLF
				strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
				strSql = strSql & " ,LTiMallLastupdate = getdate()  " & VBCRLF
				strSql = strSql & " WHERE midx = '"&idx&"' " & VBCRLF
				dbget.Execute strSql
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&idx&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&idx&"||", "")
				CALL Fn_AcctFailTouch("lotteimall", idx, SumErrStr)
				Call SugiOptionQueLogInsert("lotteimall", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))

				iErrStr = "ERR||"&idx&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE midx = '"&idx&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&idx&"||", "")
				Call SugiOptionQueLogInsert("lotteimall", action, idx, "OK", "OK||"&idx&"||"&SumOKStr, session("ssBctID"))
				iErrStr = "OK||"&idx&"||"&SumOKStr
			End If
		End If
	SET oiMall = nothing
ElseIf action = "EDIT2" Then								'���ο�����ǰ ����
	SET oiMall = new CLotteiMall
		oiMall.FRectItemID	= itemid
		oiMall.getLtimallEditOneItem

		strParam = ""
		strParam = oiMall.FOneItem.getLotteiMallItemEditParameter2()
		Call fnLtiMallInfoEdit(itemid, strParam, iErrStr, TRUE)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
		End If
		Call SugiOptionQueLogInsert("lotteimall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oiMall = nothing
ElseIf action = "REG" Then									'��ǰ ���
	SET oiMall = new CLotteiMall
		oiMall.FRectIdx	= idx
		oiMall.getLtimallNotRegOneItem

	    If (oiMall.FResultCount < 1) Then
			iErrStr = "ERR||"&idx&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT TOP 1 * FROM db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] where midx="&idx&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] "
	        strSql = strSql & " 	(midx, regdate, reguserid, LtiMallStatCD)"
	        strSql = strSql & " 	VALUES ("&idx&", getdate(), '"&session("SSBctID")&"', '0')"
			strSql = strSql & " END "
		    dbget.Execute strSql

			strParam = ""
			strParam = oiMall.FOneItem.getLotteiMallItemRegParameter(FALSE)
			Call LotteiMallItemReg(oiMall.FOneItem.Fitemid, strParam, iErrStr, oiMall.FOneItem.FRealSellprice, oiMall.FOneItem.getLotteiMallSellYn, idx)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouchOption("lotteimall", idx, iErrStr)
			End If
			Call SugiOptionQueLogInsert("lotteimall", action, idx, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oiMall = nothing
ElseIf action = "DISPVIEW" Then
	Call fnLtiMallDisView(idx, iErrStr, getLtimallGoodno(idx))
	'response.write iErrStr
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