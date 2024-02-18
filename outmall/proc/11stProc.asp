<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/11st/11stItemcls.asp"-->
<!-- #include virtual="/outmall/11st/inc11stFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, failCnt, o11st, getMustprice, t11stGoodno, vOptCnt
Dim iErrStr, strParam, mustPrice, strSql, SumErrStr, SumOKStr
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
'######################################################## 11st API ########################################################
If mallid = "11st1010" Then
	If action = "REG" Then					'��ǰ���
		SET o11st = new C11st
			o11st.FRectItemID	= itemid
			o11st.get11stNotRegOneItem

			t11stGoodno = get11stGoodno(itemid)
			If t11stGoodno <> "" Then
				iErrStr = "ERR||"&itemid&"||�̹� ��ϵ� ��ǰ �Դϴ�."
		    ElseIf (o11st.FResultCount < 1) Then
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
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/11stProc.asp?itemid=1706248&mallid=11st1010&action=REG
	ElseIf action = "SOLDOUT" Then			'���º���
		t11stGoodno = get11stGoodno(itemid)
		Call fn11stSellyn(itemid, "N", t11stGoodno, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("11st1010", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/11stProc.asp?itemid=325046&mallid=11st1010&action=SOLDOUT
	ElseIf action = "PRICE" Then			'���ݼ���
		Call get11stGoodno2(itemid, t11stGoodno, mustPrice)
		Call fn11stPrice(itemid, t11stGoodno, mustPrice, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("11st1010", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/11stProc.asp?itemid=325046&mallid=11st1010&action=PRICE
	ElseIf action = "CHKSTAT" Then			'10x10��ǰ�ڵ�� 11���� ��ǰ ��ȸ
		Call fn11stStatChk(itemid, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("11st1010", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/11stProc.asp?itemid=325046&mallid=11st1010&action=CHKSTAT
	ElseIf action = "EDIT" Then				'��ǰ����
		SET o11st = new C11st
			o11st.FRectItemID	= itemid
			o11st.get11stEditOneItem
			If o11st.FResultCount > 0 Then
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

'					If (o11st.FOneItem.FSt11SellYn = "N" AND o11st.FOneItem.IsSoldOut = False) Then
					If failCnt = 0 Then
						Call fn11stSellyn(itemid, "Y", o11st.FOneItem.Fst11GoodNo, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
'					End If
				End If
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


				If failCnt > 0 Then
					SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
					SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
					CALL Fn_AcctFailTouch("11st1010", itemid, SumErrStr)
					lastErrStr = "ERR||"&itemid&"||"&SumErrStr
					response.write "ERR||"&itemid&"||"&SumErrStr
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_11st_regitem SET " & VBCRLF
					strSql = strSql & " accFailcnt = 0  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql

					SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
					lastErrStr = "OK||"&itemid&"||"&SumOKStr
					response.write "OK||"&itemid&"||"&SumOKStr
				End If
			End If
		SET o11st = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/11stProc.asp?itemid=325046&mallid=11st1010&action=EDIT
	End If
End If
'###################################################### 11st API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
