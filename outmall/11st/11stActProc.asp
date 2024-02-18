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
<!-- #include virtual="/outmall/11st/11stItemcls.asp"-->
<!-- #include virtual="/outmall/11st/inc11stFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, o11st, oAuctionOpt, failCnt, chgSellYn, arrRows, skipItem, t11stGoodno, isAllRegYn, getMustprice
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, ccd, isItemIdChk, vOptCnt
Dim isoptionyn, isText, i
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
ccd				= request("ccd")
failCnt			= 0

' If ((session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang")) Then
' Else
' 	response.write "<script>alert('�г�Ƽ�� API���� ���ҽ��ϴ�.by ������')</script>"
' 	response.end
' End If

Select Case action
	Case "11stCommonCode", "GETCATE"	isItemIdChk = "N"
	Case Else				isItemIdChk = "Y"
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
'######################################################## 11st API ########################################################
If action = "REG" Then								'��ǰ���
	SET o11st = new C11st
		o11st.FRectItemID	= itemid
		o11st.get11stNotRegOneItem
	    If (o11st.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = ""
			strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_11st_regitem where itemid="&itemid&")"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_11st_regitem "
	        strSql = strSql & " 	(itemid, regdate, reguserid, st11statCD, regitemname, st11SellYn)"
	        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(o11st.FOneItem.FItemName)&"', 'N')"
			strSql = strSql & " END "
			dbget.Execute strSql
			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If o11st.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = o11st.FOneItem.get11stItemRegParameter()
				getMustprice = ""
				getMustprice = o11st.FOneItem.MustPrice()
				Call fn11stItemReg(itemid, strParam, iErrStr, getMustprice, o11st.FOneItem.get11stSellYn, o11st.FOneItem.FLimityn, o11st.FOneItem.FLimitNo, o11st.FOneItem.FLimitSold, html2db(o11st.FOneItem.FItemName), o11st.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[AddItem] �ɼǰ˻� ����"
			End If
		End If

		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	SET o11st = nothing

	If failCnt = 0 Then
		Call get11stGoodno3(itemid, t11stGoodno, vOptCnt)
		Call fn11stStockChk(itemid, t11stGoodno, vOptCnt, iErrStr)
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
		CALL Fn_AcctFailTouch("11st1010", itemid, SumErrStr)
		Call SugiQueLogInsert("11st1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("11st1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "EditSellYn" Then							'��ǰ ���� ����
	t11stGoodno = get11stGoodno(itemid)
	Call fn11stSellyn(itemid, chgSellYn, t11stGoodno, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11st1010", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11st1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "PRICE" Then								'���ݼ���
	Call get11stGoodno2(itemid, t11stGoodno, mustPrice)
	Call fn11stPrice(itemid, t11stGoodno, mustPrice, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11st1010", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11st1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKSTOCK" Then								'�����ȸ
	Call get11stGoodno3(itemid, t11stGoodno, vOptCnt)
	Call fn11stStockChk(itemid, t11stGoodno, vOptCnt, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11st1010", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11st1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "CHKSTAT" Then								'10x10��ǰ�ڵ�� 11���� ��ǰ ��ȸ
	Call fn11stStatChk(itemid, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11st1010", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11st1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDIT" Then									'����
	SET o11st = new C11st
		o11st.FRectItemID	= itemid
		o11st.get11stEditOneItem
		If o11st.FResultCount = 0 Then
	    	failCnt = failCnt + 1
			iErrStr = "ERR||"&itemid&"||���������� ��ǰ�� �ƴմϴ�."
		Else
			If (o11st.FOneItem.FmaySoldOut = "Y") OR (o11st.FOneItem.IsMayLimitSoldout = "Y") OR (o11st.FOneItem.FLimityn = "Y" AND (o11st.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) Then
				Call fn11stSellyn(itemid, "N", o11st.FOneItem.FSt11GoodNo, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				strParam = ""
				strParam = o11st.FOneItem.get11stItemRegParameter()
				getMustprice = ""
				getMustprice = o11st.FOneItem.MustPrice()
				Call fn11stItemEdit(itemid, strParam, iErrStr, o11st.FOneItem.FbasicimageNm, getMustprice, o11st.FOneItem.Fst11GoodNo)
				SumErrStr = replace(SumErrStr, "'", "��")

				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				If failCnt = 0 Then
					Call fn11stPrice(itemid, o11st.FOneItem.Fst11GoodNo, o11st.FOneItem.MustPrice(), iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

'				If (o11st.FOneItem.FSt11SellYn = "N" AND o11st.FOneItem.IsSoldOut = False) Then
				If failCnt = 0 Then
					Call fn11stSellyn(itemid, "Y", o11st.FOneItem.Fst11GoodNo, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
'				End If

			End If

			'OK�� ERR�̴� editQuecnt�� + 1�� ��Ŵ..
			'�����ٸ����� editQuecnt ASC, i.lastupdate DESC�� �ߺ��� ����
			'st11LastUpdate �� �����ÿ���
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_11st_regItem SET " & VBCRLF
			strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
			If failCnt = 0 Then
				strSql = strSql & " ,st11LastUpdate = getdate()  " & VBCRLF
			End If
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			If failCnt = 0 Then
				Call get11stGoodno3(itemid, t11stGoodno, vOptCnt)
				Call fn11stStockChk(itemid, t11stGoodno, vOptCnt, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		End If
	SET o11st = nothing

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("11st1010", itemid, SumErrStr)
		Call SugiQueLogInsert("11st1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_11st_regItem SET " & VBCRLF
		strSql = strSql & " accFailcnt = 0  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("11st1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "11stCommonCode" Then
	If ccd = "category" Then
		Call fn11stCommonCode(ccd, strParam)
	ElseIf ccd = "outboundarea" OR ccd = "inboundarea" Then
		Call fn11stoutinboundCode(ccd, strParam)
	End If
ElseIf action = "GETCATE" Then
	Dim tmpSafeDiv
	strSql = ""
	strSql = "EXEC [db_etcmall].[dbo].[usp_Ten_OutMall_11st_setSafeGosi] "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) then
		arrRows = rsget.getRows()
	End If
	rsget.Close
'rw UBound(arrRows,2) / 100

	For i = 0 To UBound(arrRows,2)
		Call fn11stGetCate(arrRows(0,i), iErrStr)
		If (i mod 100 = 0) Then
			rw "------------ API ȣ�����Դϴ� ------------"
			response.flush
			response.clear
		End If
	Next
	rw "------------ API �� ------------"
End If

If iErrStr <> "" Then
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str = '"&iErrStr&"<br>' + str " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
					"</script>"
End If
'###################################################### LotteiMall API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
