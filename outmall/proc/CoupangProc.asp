<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/coupang/coupangItemcls.asp"-->
<!-- #include virtual="/outmall/coupang/incCoupangFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, mallid, action, failCnt, oCoupang, getMustprice, chgSellYn, vOptCnt, i, isChkStat
Dim iErrStr, strParam, mustPrice, strSql, SumErrStr, SumOKStr, chgImageNm, arrRows, errVendorItemId, tCoupangGoodno
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
'######################################################## Coupang API ########################################################
If mallid = "coupang" Then
	If action = "SOLDOUT" Then				'���º���
		arrRows = getCoupangVendorItemidList(itemid)
		If isArray(arrRows) Then
			For i = 0 To UBound(arrRows,2)
				Call fnCoupangSellyn(itemid, "N", arrRows(0, i), errVendorItemId)
				If errVendorItemId <> "" Then
					SumErrStr = SumErrStr & errVendorItemId & ","
				End If
			Next
			iErrStr = ArrErrStrInfo("EditSellYn", "N", itemid, SumErrStr)
		Else
			iErrStr = "ERR||"&itemid&"||[���º���] ��ȸ���� �����ϼ���. by kjy"
		End If
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
		End If
		'http://wapi.10x10.co.kr/outmall/proc/CoupangProc.asp?itemid=1891798&mallid=coupang&action=SOLDOUT
	ElseIf action = "REG" Then				'��ǰ���
		SET oCoupang = new CCoupang
			oCoupang.FRectItemID	= itemid
			oCoupang.getCoupangNotRegOneItem
		    If (oCoupang.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			Else
				strSql = "EXEC [db_etcmall].[dbo].[usp_API_Coupang_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
				dbget.execute strSql

				'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
				If oCoupang.FOneItem.checkTenItemOptionValid Then
					Call fnCoupangItemReg(itemid, iErrStr)
				Else
					iErrStr = "ERR||"&itemid&"||[��ǰ���] �ɼǰ˻� ����"
				End If
			End If
		SET oCoupang = nothing
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
		End If
		'http://wapi.10x10.co.kr/outmall/proc/CoupangProc.asp?itemid=1489931&mallid=coupang&action=REG
	ElseIf action = "PRICE" Then			'���ݼ���
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
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
		End If
		'http://wapi.10x10.co.kr/outmall/proc/CoupangProc.asp?itemid=1891798&mallid=coupang&action=PRICE
	ElseIf action = "CHKSTAT" Then			'��ǰ��ȸ
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
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
		'http://wapi.10x10.co.kr/outmall/proc/CoupangProc.asp?itemid=1891798&mallid=coupang&action=CHKSTAT
	ElseIf action = "EDIT" Then				'��ǰ����
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

						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
				'######################################## 5. ��� ���� ########################################
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

		'OK�� ERR�̴� editQuecnt�� + 1�� ��Ŵ..
		'�����ٸ����� editQuecnt ASC, i.lastupdate DESC�� �ߺ��� ����
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_regItem SET " & VBCRLF
		strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
		strSql = strSql & " ,coupangLastUpdate = getdate()  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("coupang", itemid, SumErrStr)
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
		'http://wapi.10x10.co.kr/outmall/proc/CoupangProc.asp?itemid=1891798&mallid=coupang&action=EDIT
	End If
End If
'###################################################### Coupang API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->