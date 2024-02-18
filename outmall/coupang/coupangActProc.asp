<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/coupang/coupangItemcls.asp"-->
<!-- #include virtual="/outmall/coupang/incCoupangFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, oCoupang, failCnt, chgSellYn, arrRows, isItemIdChk, maeipdiv, mustPrice
Dim iErrStr, strParam, strSql, SumErrStr, SumOKStr, i, tCoupangGoodno, errVendorItemId, isChkStat, couponRegCnt, getRequestedId, couponId, couponItemRegCnt
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
couponId		= request("couponId")
failCnt			= 0
Select Case action
	Case "REGDELIVERY", "CATEMETA"
		isItemIdChk = "N"
		itemid			= requestCheckVar(request("itemid"),32)
	Case Else			isItemIdChk = "Y"
End Select

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

'######################################################## Coupang API ########################################################
If action = "REGDELIVERY" Then								'��������
	maeipdiv = fnBrandmaeipdiv(itemid)
	Call fnCoupangDeliveryReg(itemid, maeipdiv, iErrStr)
ElseIf action = "CATEMETA" Then								'ī�װ� ��Ÿ����
	Call fnCoupangCateMeta(itemid)
ElseIf action = "REG" Then									'��ǰ���
	SET oCoupang = new CCoupang
		oCoupang.FRectItemID	= itemid
		oCoupang.getCoupangNotRegOneItem
	    If (oCoupang.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Coupang_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
			dbget.execute strSql

			if (oCoupang.FOneItem.FoptionCnt>=120) then
				iErrStr = "ERR||"&itemid&"||[��ǰ���] �ɼǼ��� 120�� �̻� ��ϺҰ�"
			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			elseIf oCoupang.FOneItem.checkTenItemOptionValid Then
				Call fnCoupangItemReg(itemid, iErrStr)
			Else
				iErrStr = "ERR||"&itemid&"||[��ǰ���] �ɼǰ˻� ����"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("coupang", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oCoupang = nothing
ElseIf action = "CHKSTAT" Then								'���� ��ȸ + ��� ����
	tCoupangGoodno = getCoupangGoodno(itemid)
	Call fnCoupangStatChk(itemid, tCoupangGoodno, iErrStr)
	If Left(iErrStr, 2) <> "OK" Then
		failCnt = failCnt + 1
		SumErrStr = SumErrStr & iErrStr
	Else
		SumOKStr = SumOKStr & iErrStr
	End If

	If (failCnt = 0) AND (Instr(SumOKStr, "���οϷ�") > 0) Then
		arrRows = getCoupangVendorItemidChkStatList(itemid)
		If isArray(arrRows) Then
			For i = 0 To UBound(arrRows,2)
				Call fnCoupangQuantity(itemid, arrRows(0, i), arrRows(3, i), arrRows(4, i), errVendorItemId)
				If errVendorItemId <> "" Then
					SumErrStr = SumErrStr & errVendorItemId & ","
				End If
			Next
			iErrStr = ArrErrStrInfo("QTY", "", itemid, SumErrStr)
			isChkStat = "Y"
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			If isChkStat = "Y" Then
				strSql = ""
				strSql = strSql & " UPDATE R "
				strSql = strSql & " SET regedOptCnt = isNull(T.regedOptCnt, 0) "
				strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_regItem R "
				strSql = strSql & " JOIN ( "
				strSql = strSql & " 	SELECT R.itemid, count(*) as CNT "
				strSql = strSql & " 	,sum(CASE WHEN itemoption <> '0000' THEN 1 ELSE 0 END) as regedOptCnt "
				strSql = strSql & " 	FROM db_etcmall.dbo.tbl_coupang_regItem R "
				strSql = strSql & " 	JOIN db_etcmall.dbo.tbl_coupang_regedoption as Ro on R.itemid = Ro.itemid and Ro.itemid = '"&itemid&"' "
				strSql = strSql & " 	WHERE Ro.outmallsellyn = 'Y' "
				strSql = strSql & " 	and R.itemid = '"&itemid&"' "
				strSql = strSql & " 	GROUP BY R.itemid "
				strSql = strSql & " ) as T on R.itemid = T.itemid and R.itemid = '"&itemid&"' "
				dbget.Execute(strSql)
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
	End If

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("coupang", itemid, SumErrStr)
		Call SugiQueLogInsert("coupang", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("coupang", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "EditSellYn" Then							'��ǰ ���� ����
	If chgSellYn = "N" Then
		arrRows = getCoupangVendorItemidSellNList(itemid)
	Else
		arrRows = getCoupangVendorItemidList(itemid)
	End If

	If isArray(arrRows) Then
		For i = 0 To UBound(arrRows,2)
			Call fnCoupangSellyn(itemid, chgSellYn, arrRows(0, i), errVendorItemId)
			If errVendorItemId <> "" Then
				SumErrStr = SumErrStr & errVendorItemId & ","
			End If
		Next
		iErrStr = ArrErrStrInfo(action, chgSellYn, itemid, SumErrStr)
	Else
		iErrStr = "ERR||"&itemid&"||[���º���] ��ȸ���� �����ϼ���. by kjy"
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("coupang", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "DELETE" Then								'��ǰ ����
	tCoupangGoodno = getCoupangGoodno(itemid)
	Call fnCoupangDelete(itemid, tCoupangGoodno, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("coupang", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "PRICE" Then								'���� ����
	arrRows = getCoupangVendorItemidList(itemid)
	If isArray(arrRows) Then
		For i = 0 To UBound(arrRows,2)
			Call fnCoupangPrice(itemid, arrRows(0, i), arrRows(1, i), arrRows(2, i), errVendorItemId)
			If errVendorItemId <> "" Then
				SumErrStr = SumErrStr & errVendorItemId & ","
			End If
			mustPrice = arrRows(1, i)
		Next
		iErrStr = ArrErrStrInfo(action, mustPrice, itemid, SumErrStr)
	Else
		iErrStr = "ERR||"&itemid&"||[���ݼ���] ��ȸ���� �����ϼ���. by kjy"
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("coupang", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "QTY" Then									'��� ����
	arrRows = getCoupangVendorItemidList(itemid)
	If isArray(arrRows) Then
		For i = 0 To UBound(arrRows,2)
			Call fnCoupangQuantity(itemid, arrRows(0, i), arrRows(3, i), arrRows(4, i), errVendorItemId)
			If errVendorItemId <> "" Then
				SumErrStr = SumErrStr & errVendorItemId & ","
			End If
		Next
		iErrStr = ArrErrStrInfo(action, "", itemid, SumErrStr)
	Else
		iErrStr = "ERR||"&itemid&"||[������] ��ȸ���� �����ϼ���. by kjy"
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
	Else
		strSql = ""
		strSql = strSql & " UPDATE R "
		strSql = strSql & " SET regedOptCnt = isNull(T.regedOptCnt, 0) "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_regItem R "
		strSql = strSql & " JOIN ( "
		strSql = strSql & " 	SELECT R.itemid, count(*) as CNT "
		strSql = strSql & " 	,sum(CASE WHEN itemoption <> '0000' THEN 1 ELSE 0 END) as regedOptCnt "
		strSql = strSql & " 	FROM db_etcmall.dbo.tbl_coupang_regItem R "
		strSql = strSql & " 	JOIN db_etcmall.dbo.tbl_coupang_regedoption as Ro on R.itemid = Ro.itemid and Ro.itemid = '"&itemid&"' "
		strSql = strSql & " 	WHERE Ro.outmallsellyn = 'Y' "
		strSql = strSql & " 	and R.itemid = '"&itemid&"' "
		strSql = strSql & " 	GROUP BY R.itemid "
		strSql = strSql & " ) as T on R.itemid = T.itemid and R.itemid = '"&itemid&"' "
		dbget.Execute(strSql)
	End If
	Call SugiQueLogInsert("coupang", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDIT" Then									'��ǰ ����
	SET oCoupang = new CCoupang
		oCoupang.FRectItemID	= itemid
		oCoupang.getCoupangEditOneItem
		If oCoupang.FResultCount = 0 Then
	    	failCnt = failCnt + 1
			iErrStr = "ERR||"&itemid&"||���������� ��ǰ�� �ƴմϴ�."
		Else
            '######################################## 1. ��ǰ��ȸ ########################################
            Call fnCoupangStatChk(itemid, oCoupang.FOneItem.FcoupangGoodNo, iErrStr)
            If Left(iErrStr, 2) <> "OK" Then
                failCnt = failCnt + 1
                SumErrStr = SumErrStr & iErrStr
            Else
                SumOKStr = SumOKStr & iErrStr
            End If

			arrRows = getCoupangVendorItemidList(itemid)
			'######################################## 2-1. ǰ�� ó�� ########################################
			If (oCoupang.FOneItem.FmaySoldOut = "Y") OR (oCoupang.FOneItem.IsMayLimitSoldout = "Y") OR (oCoupang.FOneItem.IsAllOptionChange = "Y") Then
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						Call fnCoupangSellyn(itemid, "N", arrRows(0, i), errVendorItemId)
						If errVendorItemId <> "" Then
							SumErrStr = SumErrStr & errVendorItemId & ","
						End If
					Next
					iErrStr = ArrErrStrInfo("EditSellYn", "N", itemid, SumErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			Else
			'######################################## 2-2. �Ǹ� ó�� ########################################
				If (failCnt = 0) AND (Instr(SumOKStr, "���οϷ�") > 0) Then
					If (oCoupang.FOneItem.FCoupangSellYn = "N" AND oCoupang.FOneItem.IsSoldOut = False) Then
						If isArray(arrRows) Then
							For i = 0 To UBound(arrRows,2)
								Call fnCoupangSellyn(itemid, "Y", arrRows(0, i), errVendorItemId)
								If errVendorItemId <> "" Then
									SumErrStr = SumErrStr & errVendorItemId & ","
								End If
							Next
							iErrStr = ArrErrStrInfo("EditSellYn", "Y", itemid, SumErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						End If
					End If
			'######################################## 3. ���� ó�� ########################################
					Call fnCoupangItemEdit(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

			'######################################## 4. ���� ���� ########################################

					If isArray(arrRows) Then
						For i = 0 To UBound(arrRows,2)
							Call fnCoupangPrice(itemid, arrRows(0, i), arrRows(1, i), arrRows(2, i), errVendorItemId)
							If errVendorItemId <> "" Then
								SumErrStr = SumErrStr & errVendorItemId & ","
							End If
							mustPrice = arrRows(1, i)
						Next
						iErrStr = ArrErrStrInfo("PRICE", mustPrice, itemid, SumErrStr)
					End If

					If LEFT(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

			'######################################## 5. ������ ########################################
					If isArray(arrRows) Then
						For i = 0 To UBound(arrRows,2)
							Call fnCoupangQuantity(itemid, arrRows(0, i), arrRows(3, i), arrRows(4, i), errVendorItemId)
							If errVendorItemId <> "" Then
								SumErrStr = SumErrStr & errVendorItemId & ","
							End If
						Next
						iErrStr = ArrErrStrInfo("QTY", "", itemid, SumErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							strSql = ""
							strSql = strSql & " UPDATE R "
							strSql = strSql & " SET regedOptCnt = isNull(T.regedOptCnt, 0) "
							strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_regItem R "
							strSql = strSql & " JOIN ( "
							strSql = strSql & " 	SELECT R.itemid, count(*) as CNT "
							strSql = strSql & " 	,sum(CASE WHEN itemoption <> '0000' THEN 1 ELSE 0 END) as regedOptCnt "
							strSql = strSql & " 	FROM db_etcmall.dbo.tbl_coupang_regItem R "
							strSql = strSql & " 	JOIN db_etcmall.dbo.tbl_coupang_regedoption as Ro on R.itemid = Ro.itemid and Ro.itemid = '"&itemid&"' "
							strSql = strSql & " 	WHERE Ro.outmallsellyn = 'Y' "
							strSql = strSql & " 	and R.itemid = '"&itemid&"' "
							strSql = strSql & " 	GROUP BY R.itemid "
							strSql = strSql & " ) as T on R.itemid = T.itemid and R.itemid = '"&itemid&"' "
							dbget.Execute(strSql)

							strSql = ""
							strSql = strSql & " DECLARE @A int, @B int"&VbCRLF
							strSql = strSql & " SELECT @A = COUNT(CASE WHEN (outmallsellyn <> 'Y' OR outmalllimitno < 1) THEN 1 END)"&VbCRLF
							strSql = strSql & " , @B = COUNT(*)"&VbCRLF
							strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_regedoption"&VbCRLF
							strSql = strSql & " WHERE itemid = '"&itemid&"'"&VbCRLF
							strSql = strSql & " IF @A = @B"&VbCRLF
							strSql = strSql & " 	BEGIN"&VbCRLF
							strSql = strSql & " 		UPDATE db_etcmall.dbo.tbl_coupang_regItem"&VbCRLF
							strSql = strSql & " 		SET CoupangSellYn = 'N'"&VbCRLF
							strSql = strSql & " 		WHERE itemid = '"&itemid&"'"&VbCRLF
							strSql = strSql & " 	END"
							dbget.Execute(strSql)
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
					If failCnt = 0 Then
						strSql = ""
						strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_regItem SET " & VBCRLF
						strSql = strSql & " accFailcnt = 0  " & VBCRLF
						strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
						dbget.Execute strSql
					End If
				End If
			End If
		End If
	SET oCoupang = nothing

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		Call SugiQueLogInsert("coupang", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("coupang", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "COUPONREG" Then								'����������� ��� / ��ȸ
	couponRegCnt = getCoupangCouponNotRegCount(itemid)
	If couponRegCnt = 0 Then
		failCnt = failCnt + 1
		iErrStr = "ERR||"&itemid&"||��ϰ����� ������ �ƴմϴ�."
	Else
		Call fnCoupangCouponReg(itemid, iErrStr)
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If

		If failCnt = 0 Then
			getRequestedId = getCoupangCouponRequestedId(itemid)
			Call fnCoupangCouponStat(itemid, getRequestedId, iErrStr)
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
		Call SugiQueLogInsert("coupang", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("coupang", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "COUPONDETAILREG" Then
	Call fnCoupangCouponItemReg(itemid, couponId, iErrStr)
	Call SugiQueLogInsert("coupang", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "COUPONDETAILDEL" Then
	Call fnCoupangCouponItemDel(itemid, couponId, iErrStr)
	Call SugiQueLogInsert("coupang", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
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