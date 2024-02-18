<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/lotteon/lotteonItemcls.asp"-->
<!-- #include virtual="/outmall/lotteon/inclotteonFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, olotteon, failCnt, chgSellYn, arrRows, getMustprice, addOptErrItem
Dim iErrStr, strParam, strSql, SumErrStr, SumOKStr, isItemIdChk, grpVal, rSkip, rLimit, i, outmallorderserial
Dim requestJson, responseJson, callComplete, hasnext
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
grpVal			= request("grpVal")
rSkip			= request("rSkip")
rLimit			= request("rLimit")
requestJson		= request("requestJson")
responseJson	= request("responseJson")
failCnt			= 0
outmallorderserial = request("outmallorderserial")
addOptErrItem	= "N"
callComplete = "N"

''ī�װ� ����ý� �ϴ� �����ؾ� ��..
' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_StdCategory]
' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_StdCategory_Disp]
' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_StdCategory_Attr] 
' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_tmpStdCategory] 

' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_DispCategory]
' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_tmpDispCategory]

' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_Attribute] 
' --TRUNCATE TABLE db_etcmall.[dbo].[tbl_lotteon_Attribute_Values]
Select Case action
	Case "DVPVIEW", "GRPCD", "GRPDTLCD", "ATTRVIEW", "DISPCATE", "STDCATE", "BRANDVIEW", "ORDVIEW"
		isItemIdChk = "N"
	Case Else
		isItemIdChk = "Y"
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
'######################################################## LotteOn API ########################################################
'http://localhost:11117/outmall/lotteon/lotteonActProc.asp?itemid=214560&act=PRICE&requestJson=Y&responseJson=Y
If action = "REG" Then									'��ǰ ���
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotRegOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Outmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"', '"&CMALLNAME&"' "
			dbget.execute strSql
			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If olotteon.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = olotteon.FOneItem.getLotteonItemRegParameter()
				If requestJson = "Y" Then
					response.write strParam
				End If
				getMustprice = ""
				getMustprice = olotteon.FOneItem.MustPrice()
				CALL fnLotteonItemReg(itemid, strParam, iErrStr, getMustprice, olotteon.FOneItem.getLotteonSellYn, olotteon.FOneItem.FLimityn, olotteon.FOneItem.FLimitNo, olotteon.FOneItem.FLimitSold, html2db(olotteon.FOneItem.FItemName), olotteon.FOneItem.FbasicimageNm, responseJson)
			Else
				iErrStr = "ERR||"&itemid&"||[��ǰ���] �ɼǰ˻� ����"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteon", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET olotteon = nothing
ElseIf action = "CHKSTAT" Then							'��ǰ �� ��ȸ
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ȸ ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = olotteon.FOneItem.getLotteonItemViewParameter()
			If requestJson = "Y" Then
				response.write strParam
			End If
			CALL fnLotteonItemView(itemid, strParam, iErrStr, responseJson)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteon", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET olotteon = nothing
ElseIf action = "EDITINFO" Then							'��ǰ�� ����
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = olotteon.FOneItem.getLotteonItemEditParameter()			'���� ��ǰ ����
			If requestJson = "Y" Then
				response.write strParam
			End If
			CALL fnLotteonItemEdit(itemid, olotteon.FOneItem.FItemName, strParam, iErrStr, responseJson)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteon", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET olotteon = nothing
ElseIf action = "EDIT" Then								'��ǰ ����
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strSql = "exec db_etcmall.dbo.usp_Ten_OutMall_optEditParamList_lotteon '"&CMallName&"'," & itemid
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows
			End If
			rsget.close

			If UBound(arrRows,2) = 0 AND arrRows(0,0) = "Z1" Then
				addOptErrItem = "Y"
			End If

			If (oLotteon.FOneItem.FmaySoldOut = "Y") OR (oLotteon.FOneItem.IsMayLimitSoldout = "Y") OR (oLotteon.FOneItem.IsSoldOut) OR (oLotteon.FOneItem.FOptionCnt = 0 AND oLotteon.FOneItem.getRegedOptionCnt > 0)  OR (oLotteon.FOneItem.FLimityn = "Y" AND (oLotteon.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) OR (addOptErrItem = "Y") Then
				chgSellYn = "N"
			Else
				chgSellYn = "Y"
			End If

            If chgSellYn = "N" Then
				strParam = ""
				strParam = olotteon.FOneItem.getLotteonSellynParameter(chgSellYn)
				Call fnLotteOnSellyn(itemid, chgSellYn, strParam, iErrStr, responseJson)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
            Else
				If (olotteon.FOneItem.FLotteonSellYn <> "Y" AND chgSellYn = "Y") Then
					strParam = ""
					strParam = olotteon.FOneItem.getLotteonSellynParameter(chgSellYn)
					Call fnLotteOnSellyn(itemid, chgSellYn, strParam, iErrStr, responseJson)
					response.flush
					response.clear
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				strParam = ""
				strParam = olotteon.FOneItem.getLotteonItemViewParameter()				'��ǰ �� ��ȸ
				CALL fnLotteonItemView(itemid, strParam, iErrStr, responseJson)
			    rw "��ǰ �� ��ȸ"
				response.flush
				response.clear
                If Left(iErrStr, 2) <> "OK" Then
                    failCnt = failCnt + 1
                    SumErrStr = SumErrStr & iErrStr
                Else
                    SumOKStr = SumOKStr & iErrStr
                End If

				'���� ��ǰ ����, ��ǰ ���� ����, ��ǰ ��� ������� �ּ�ó��
				' If failCnt = 0 Then
				' 	strParam = ""
				' 	strParam = olotteon.FOneItem.getLotteonItemEditParameter()			'���� ��ǰ ����
				' 	CALL fnLotteonItemEdit(itemid, olotteon.FOneItem.FItemName, strParam, iErrStr, responseJson)
				' 	rw "���� ��ǰ ����"
				' 	response.flush
				' 	response.clear
				' 	If Left(iErrStr, 2) <> "OK" Then
				' 		failCnt = failCnt + 1
				' 		SumErrStr = SumErrStr & iErrStr
				' 	Else
				' 		SumOKStr = SumOKStr & iErrStr
				' 	End If
				' End If

				' If failCnt = 0 Then
				' 	getMustprice = ""
				' 	getMustprice = olotteon.FOneItem.MustPrice()
				' 	strParam = ""
				' 	strParam = olotteon.FOneItem.getLotteonPriceParameter()				'��ǰ ���� ����
				' 	Call fnLotteOnPrice(itemid, strParam, getMustprice, iErrStr, responseJson)
				' 	rw "��ǰ ���� ����"
				' 	response.flush
				' 	response.clear
				' 	If Left(iErrStr, 2) <> "OK" Then
				' 		failCnt = failCnt + 1
				' 		SumErrStr = SumErrStr & iErrStr
				' 	Else
				' 		SumOKStr = SumOKStr & iErrStr
				' 	End If
				' End If

				' If failCnt = 0 Then
				' 	strParam = ""
				' 	strParam = olotteon.FOneItem.getLotteonQuantityParameter()			'��ǰ ��� ����
				' 	Call fnLotteOnQuantity(itemid, strParam, iErrStr, responseJson)
				' 	rw "��ǰ ��� ����"
				' 	response.flush
				' 	response.clear
				' 	If Left(iErrStr, 2) <> "OK" Then
				' 		failCnt = failCnt + 1
				' 		SumErrStr = SumErrStr & iErrStr
				' 	Else
				' 		SumOKStr = SumOKStr & iErrStr
				' 	End If
				' End If

				'################## �� �ּ��� �Ʒ��� ���� ����..2020-05-07 ���� ######################
				If failCnt = 0 Then
				 	getMustprice = ""
				 	getMustprice = olotteon.FOneItem.MustPrice()

					strParam = ""
					strParam = olotteon.FOneItem.getLotteonItemEditParameter()			'���� ��ǰ ����
					CALL fnLotteonItemEdit2(itemid, olotteon.FOneItem.FItemName, getMustprice, strParam, iErrStr, responseJson)
					rw "���� ��ǰ ����"
					response.flush
					response.clear
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
				'################## ���� ����..2020-05-07 ���� ######################

				If failCnt = 0 Then
					strParam = ""
					strParam = olotteon.FOneItem.getLotteonOptStatusParameter()			'��ǰ �ǸŻ��� ����
					Call fnLotteOnOptStat(itemid, strParam, iErrStr, responseJson)
					rw "��ǰ �ǸŻ��� ����"
					response.flush
					response.clear
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				If failCnt = 0 Then
					strParam = ""
					strParam = olotteon.FOneItem.getLotteonItemViewParameter()			'��ǰ �� ��ȸ
					CALL fnLotteonItemView(itemid, strParam, iErrStr, responseJson)
					rw "��ǰ �� ��ȸ"
					response.flush
					response.clear
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
                strSql = strSql & " UPDATE db_etcmall.dbo.tbl_lotteon_regitem SET " & VBCRLF
                strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
                strSql = strSql & " ,lotteonlastupdate = getdate()  " & VBCRLF
                strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
                dbget.Execute strSql
			End If
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("lotteon", itemid, SumErrStr)
			Call SugiQueLogInsert("lotteon", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_lotteon_regitem SET " & VBCRLF
			strSql = strSql & " accFailcnt = 0  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("lotteon", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET olotteon = nothing
ElseIf action = "QTY" Then								'��ǰ ��� ����
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = olotteon.FOneItem.getLotteonQuantityParameter()
			If requestJson = "Y" Then
				response.write strParam
			End If
			Call fnLotteOnQuantity(itemid, strParam, iErrStr, responseJson)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteon", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET olotteon = nothing
ElseIf action = "PRICE" Then							'��ǰ ���� ����
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		Else
			If LEFT(olotteon.FOneItem.FLastStatCheckDate, 10) = "1900-01-01" Then
				strParam = ""
				strParam = olotteon.FOneItem.getLotteonItemViewParameter()			'��ǰ �� ��ȸ
				CALL fnLotteonItemView(itemid, strParam, iErrStr, responseJson)
				rw "��ǰ �� ��ȸ"
				response.flush
				response.clear
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If

			If failCnt = 0 Then
				getMustprice = ""
				getMustprice = olotteon.FOneItem.MustPrice()
				strParam = ""
				strParam = olotteon.FOneItem.getLotteonPriceParameter()				'��ǰ ���� ����
				Call fnLotteOnPrice(itemid, strParam, getMustprice, iErrStr, responseJson)
				rw "��ǰ ���� ����"
				response.flush
				response.clear
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
			CALL Fn_AcctFailTouch("lotteon", itemid, SumErrStr)
			Call SugiQueLogInsert("lotteon", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("lotteon", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET olotteon = nothing
ElseIf action = "EditSellYn" Then						'��ǰ ���� ����
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = olotteon.FOneItem.getLotteonSellynParameter(chgSellYn)
			If requestJson = "Y" Then
				response.write strParam
			End If
			Call fnLotteOnSellyn(itemid, chgSellYn, strParam, iErrStr, responseJson)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteon", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET olotteon = nothing
ElseIf action = "OPTSTAT" Then							'��ǰ �ǸŻ��� ����
	SET olotteon = new CLotteon
		olotteon.FRectItemID	= itemid
		olotteon.getLotteonNotEditOneItem
	    If (olotteon.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
		Else
			strParam = ""
			strParam = olotteon.FOneItem.getLotteonOptStatusParameter()
			If requestJson = "Y" Then
				response.write strParam
			End If
			Call fnLotteOnOptStat(itemid, strParam, iErrStr, responseJson)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("lotteon", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET olotteon = nothing
ElseIf action = "DVPVIEW" Then							'�Ǹ��� �����/��ǰ�� ����Ʈ ��ȸ
	Call fnlotteonDVPView()
ElseIf action = "ATTRVIEW" Then							'�Ӽ� �⺻ ��ȸ
	Do Until callComplete = "Y"
		Call fnlotteonAttrView(rSkip, hasnext)
		If hasnext = "N" Then
			callComplete = "Y"
			rw "�Ϸ�"
		Else
			rw "API ȣ�� �� �Դϴ�. "
			rw "-------------------------"
		End If
		response.flush
	Loop
ElseIf action = "DISPCATE" Then							'����ī�װ� ��ȸ
	Do Until callComplete = "Y"
		Call fnlotteonDispCateView(rSkip, hasnext)
		If hasnext = "N" Then
			callComplete = "Y"
			rw "�Ϸ�"
		Else
			rw "API ȣ�� �� �Դϴ�. "
			rw "-------------------------"
		End If
		response.flush
	Loop
ElseIf action = "STDCATE" Then							'ǥ��ī�װ� ��ȸ
	Do Until callComplete = "Y"
		Call fnlotteonStdCateView(rSkip, hasnext)
		If hasnext = "N" Then
			callComplete = "Y"
			rw "�Ϸ�"
		Else
			rw "API ȣ�� �� �Դϴ�. "
			rw "-------------------------"
		End If
		response.flush
	Loop
ElseIf action = "BRANDVIEW" Then						'�귣�� ��ȸ
	Call fnlotteonBrandView(rSkip, rLimit)
ElseIf action = "GRPCD" Then							'�����ڵ� ��ȸ
	Call fnlotteonGetGroupCode()
ElseIf action = "GRPDTLCD" Then							'�����ڵ� �� ��ȸ
	Call fnlotteonGetGroupCodeDetail(grpVal)
ElseIf action = "ORDVIEW" Then							'��ۻ��� ��ȸ / JSON ��¸� ����
	Call fnlotteonViewOrder(outmallorderserial)
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
'###################################################### LotteOn API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
