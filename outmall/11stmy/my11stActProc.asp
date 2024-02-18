<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/11stmy/my11stItemcls.asp"-->
<!-- #include virtual="/outmall/11stmy/incmy11stFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, omy11st, failCnt, chgSellYn, arrRows, isItemIdChk
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, ccd, vMy11stGoodno
Dim isoptionyn, isText, i, chgOptCnt, vOrgprice, mayOptSoldOut, vExchangeRate, vMultiplerate, vMaySellPrice
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
ccd				= request("ccd")
failCnt			= 0
Select Case action
	Case "my11stCommonCode"			isItemIdChk = "N"
	Case Else						isItemIdChk = "Y"
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
		itemid=CLng(getNumeric(itemid))
	End If
End If
'######################################################## 11���� API ########################################################
If action = "REG" Then									'��ǰ ���
	SET omy11st = new CMy11st
		omy11st.FRectItemID	= itemid
		omy11st.getmy11stNotRegOneItem
		If (omy11st.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			chgOptCnt = omy11st.getchangeOptionNameCnt(itemid)
			If (omy11st.FOneItem.FOptioncnt > 0) AND (chgOptCnt = 0) Then
				iErrStr = "ERR||"&itemid&"||�ɼ� ���� �� �ɼ� ��뿩�� Ȯ���ϼ���."
			Else
				If (omy11st.FOneItem.FOptioncnt > 0) AND omy11st.FOneItem.IsMayLimitSoldout = "Y" Then
					iErrStr = "ERR||"&itemid&"||�ɼ� ���� ����"
				Else
					'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
					If omy11st.FOneItem.checkTenItemOptionValid Then
						strSql = ""
						strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_my11st_regItem] where itemid="&itemid&")"
						strSql = strSql & " BEGIN"& VbCRLF
						strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_my11st_regItem] "
						strSql = strSql & " (itemid, regdate, reguserid, my11ststatCD, transItemname, regitemname)"
						strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(omy11st.FOneitem.FTransItemName)&"',  '"&html2db(omy11st.FOneitem.FItemName)&"')"
						strSql = strSql & " END "
						dbget.Execute strSql

						strParam = ""
						strParam = omy11st.FOneItem.getMy11stItemRegXML("")

						Call fnMy11stItemReg(itemid, strParam, omy11st.FOneItem.FOrgprice, omy11st.FOneItem.FMaySellPrice, omy11st.FOneItem.FOptRecordCnt, omy11st.FOneItem.FMultiplerate, omy11st.FOneItem.FExchangeRate, iErrStr)
					Else
						iErrStr = "ERR||"&itemid&"||�ɼǰ˻� ����"
					End If
				End If
			End If
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET omy11st = nothing
ElseIf action = "EDIT" Then								'��ǰ ����
	SET omy11st = new CMy11st
		omy11st.FRectItemID	= itemid
		omy11st.getmy11stlEditOneItem
		If omy11st.FResultCount > 0 Then
			If omy11st.FOneItem.FOptioncnt > 0 Then
				mayOptSoldOut = omy11st.FOneItem.IsMayLimitSoldout
			End If

			If (omy11st.FOneItem.FMaySoldOut = "Y") OR (omy11st.FOneItem.IsSoldOutLimit5Sell) OR (mayOptSoldOut = "Y") Then
				Call fnMy11stSoldOut(itemid, omy11st.FOneItem.FMy11stGoodNo, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (omy11st.FOneItem.FMy11stSellYn = "N" AND omy11st.FOneItem.IsSoldOut = False) Then
					iErrStr = ""
					Call fnMy11stOnSale(itemid, omy11st.FOneItem.FMy11stGoodNo, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				'��ǰ ����
				strParam = ""
				strParam = omy11st.FOneItem.getMy11stItemRegXML(omy11st.FOneItem.FMy11stGoodNo)
				Call fnMy11stItemEdit(itemid, omy11st.FOneItem.FMy11stGoodNo, strParam, omy11st.FOneItem.FOrgprice, omy11st.FOneItem.FExchangeRate, omy11st.FOneItem.FMultiplerate, omy11st.FOneItem.FMaySellPrice, omy11st.FOneItem.FOptRecordCnt, omy11st.FOneItem.FNotdb2HTMLitemname, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'�ɼ� ��ȸ
				Call fnMy11stOptView(itemid, omy11st.FOneItem.FMy11stGoodNo, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "'", "")
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("11stmy", itemid, SumErrStr)
				Call SugiQueLogInsert("11stmy", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				Call SugiQueLogInsert("11stmy", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
				iErrStr = "OK||"&itemid&"||"&SumOKStr
			End If
		End If
	SET omy11st = nothing
ElseIf action = "PRICE" Then							'�Ǹ� ���� ����
	vMy11stGoodno = getMy11stGoodNo(itemid)
	If vMy11stGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||11���� ��ǰ�ڵ� ����"
	Else
		Call getMy11stRatePrice(itemid, vOrgprice, vExchangeRate, vMultiplerate, vMaySellPrice)
		Call fnMy11stPrice(itemid, vMy11stGoodno, vOrgprice, vExchangeRate, vMultiplerate, vMaySellPrice, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "SOLDOUT" Then							'�Ǹ� ���� ���� N
	vMy11stGoodno = getMy11stGoodNo(itemid)
	If vMy11stGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||11���� ��ǰ�ڵ� ����"
	Else
		Call fnMy11stSoldOut(itemid, vMy11stGoodno, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "ONSALE" Then							'�Ǹ� ���� ���� Y
	vMy11stGoodno = getMy11stGoodNo(itemid)
	If vMy11stGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||11���� ��ǰ�ڵ� ����"
	Else
		Call fnMy11stOnSale(itemid, vMy11stGoodno, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDITOPT" Then							'�ɼ� ����
	SET omy11st = new CMy11st
		omy11st.FRectItemID	= itemid
		omy11st.getmy11stlEditOneItem
		If omy11st.FResultCount > 0 Then
			strParam = ""
			strParam = omy11st.FOneItem.getMy11stOptEditXML()
			Call fnMy11stOptEdit(itemid, omy11st.FOneItem.FMy11stGoodNo, strParam, omy11st.FOneItem.FOptRecordCnt, iErrStr)
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET omy11st = nothing
ElseIf action = "VIEW" Then								'��ǰ ��ȸ
	vMy11stGoodno = getMy11stGoodNo(itemid)
	If vMy11stGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||11���� ��ǰ�ڵ� ����"
	Else
		Call fnMy11stView(itemid, vMy11stGoodno, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "VIEWOPT" Then							'�ɼ� ��ȸ
	vMy11stGoodno = getMy11stGoodNo(itemid)
	If vMy11stGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||11���� ��ǰ�ڵ� ����"
	Else
		Call fnMy11stOptView(itemid, vMy11stGoodno, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "my11stCommonCode" Then					'�����ڵ� �˻�
	If ccd = "CATEGORYLIST" Then
		strParam = ""
		strParam = getCommCode(ccd)
	End If
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
'###################################################### 11���� API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->