<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/ltimall/LtimallItemcls.asp"-->
<!-- #include virtual="/outmall/ltimall/incLtimallFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<!-- #include virtual="/outmall/ltimall/inc_dailyAuthCheck.asp"-->
<%
Dim itemid, mallid, action, oiMall, failCnt, arrRows, skipItem
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, iAddOptCnt, mode, assin
Dim jenkinsBatchYn, idx, lastErrStr
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
mode			= request("mode")
jenkinsBatchYn	= request("jenkinsBatchYn")
idx				= request("idx")
lastErrStr		= ""
If mode = "updateSendState" Then
	strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendState='"&request("updateSendState")&"'"
	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
	strSql = strSql & "	where OutMallOrderSerial='"&request("ord_no")&"'"
	strSql = strSql & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	dbget.Execute strSql,assin
	response.write "<script>alert('"&assin&"�� �Ϸ� ó��.');opener.close();window.close()</script>"
	response.end
ElseIf mode = "etcSongjangFin" Then
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendState=7"
	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
    strSql = strSql & "	where OutMallOrderSerial='"&request("ord_no")&"'"
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
'######################################################## LotteiMall API ########################################################
If mallid = "lotteimall" Then
	If action = "SOLDOUT" Then									'���º���
		strParam = ""
		strParam = getLtiMallSellynParameter("N", getLtimallGoodno(itemid))
		Call fnLtiMallSellyn(itemid, "N", strParam, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=SOLDOUT
	ElseIf action = "PRICE" Then								'���ݼ���
		SET oiMall = new CLotteiMall
			oiMall.FRectItemID	= itemid
			oiMall.getLtimallEditOneItem
			If oiMall.FResultCount > 0 Then
				mustPrice = ""
				mustPrice = oiMall.FOneItem.MustPrice()
				strParam = ""
				strParam = getLtiMallPriceParameter(itemid, getLtimallGoodno(itemid), mustPrice)
				If strParam = "" Then
					lastErrStr = "ERR||"&itemid&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
					response.write "ERR||"&itemid&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
				Else
					Call fnLtimallPrice(itemid, strParam, mustPrice, iErrStr)
					lastErrStr = iErrStr
					response.write iErrStr
				End If
			else
				lastErrStr = "ERR||"&itemid&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�.[1]"
				response.write "ERR||"&itemid&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�.[1]"
			end if

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
			End If
		SET oiMall = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=PRICE
	ElseIf action = "ITEMNAME" Then								'��ǰ�����
		strParam = ""
		strParam = getLtiMallItemnameParameter(itemid, iitemname, getLtimallGoodno(itemid))
		Call fnLtiMallChgItemname(itemid, strParam, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=ITEMNAME
	ElseIf action = "CHKSTAT" Then								'�űԻ�ǰ��ȸ
		Call fnLtiMallstatChk(itemid, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=CHKSTAT
	ElseIf action = "CHKSTOCK" Then								'�����ȸ
		Call fnLtiMallStockChk(itemid, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=CHKSTOCK
	ElseIf action = "DISPVIEW" Then
		Call fnLtiMallDisView(itemid, iErrStr, getLtimallGoodno(itemid))
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=DISPVIEW
	ElseIf action = "EDIT" Then									'�����ȸ + ��ǰ���� + ���� + �ʿ信 ���� ��ǰ�ǸŻ��¼���(2015-10-06 ������ ���û�����ȸ �ּ�ó��)
		SET oiMall = new CLotteiMall
			oiMall.FRectItemID	= itemid
			oiMall.getLtimallEditOneItem
			If oiMall.FResultCount > 0 Then
				If (oiMall.FOneItem.FmaySoldOut = "Y") OR (oiMall.FOneItem.IsSoldOutLimit5Sell) OR (oiMall.FOneItem.IsMayLimitSoldout = "Y") Then
					strParam = ""
					strParam = getLtiMallSellynParameter("N", getLtimallGoodno(itemid))
					Call fnLtiMallSellyn(itemid, "N", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				ElseIf oiMall.FOneItem.isDuppOptionItemYn = "Y" Then
					strParam = ""
					strParam = getLtiMallSellynParameter("X", getLtimallGoodno(itemid))

					Call fnLtiMallSellyn(itemid, "X", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					If (oiMall.FOneItem.FLtimallSellYn = "N" AND oiMall.FOneItem.IsSoldOut = False) Then
						iErrStr = ""
						strParam = ""
						strParam = getLtiMallSellynParameter("Y", getLtimallGoodno(itemid))
						Call fnLtiMallSellyn(itemid, "Y", strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

			        If (oiMall.FOneItem.FSellcash <> oiMall.FOneItem.FLtiMallPrice) Then
						mustPrice = ""
						mustPrice = oiMall.FOneItem.MustPrice()
						strParam = ""
						strParam = getLtiMallPriceParameter(itemid, getLtimallGoodno(itemid), mustPrice)
						If strParam = "" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & "ERR||"&itemid&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
						Else
							Call fnLtimallPrice(itemid, strParam, mustPrice, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						End If
					End If

					'2013-07-01 ���û�ǰ ��ǰ�߰� �� ��� �߰�
					Dim dp, aoptNm, aoptDc
					strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_ltimall '"&CMALLNAME&"'," & itemid
					rsget.CursorLocation = adUseClient
					rsget.CursorType = adOpenStatic
					rsget.LockType = adLockOptimistic
					rsget.Open strSql, dbget
					If Not(rsget.EOF or rsget.BOF) Then
					    arrRows = rsget.getRows
					End If
					rsget.close

					''�߰��� �ɼ� ���
					If isArray(arrRows) Then
						For dp = 0 To UBound(ArrRows, 2)
							If (ArrRows(11,dp)=0) and ArrRows(12,dp) = "1" AND ArrRows(15,dp) = "" Then		'�ɼǸ��� �ٸ��� �ɼ��ڵ尪�� ���� �� ==> ��ǰ�߰� �ǹ�// preged 0
								aoptNm = Replace(db2Html(ArrRows(2,dp)),":","")
								If aoptNm = "" Then
									aoptNm = "�ɼ�"
								End If
								aoptDc = aoptDc & Replace(Replace(db2Html(ArrRows(3,dp)),":",""),"'","")&","
							End If
						Next

						If aoptDc <> "" Then
	'						rw "��ǰ�߰�:"&aoptDc
							aoptDc = Left(aoptDc, Len(aoptDc) - 1)
							strParam = ""
							strParam = getLtiMallAddOptParameter(aoptNm, aoptDc, getLtimallGoodno(itemid))
							CALL fnLtiMallAddOpt(itemid, strParam, iErrStr, iAddOptCnt)
							If iAddOptCnt > 0 Then
								If Left(iErrStr, 2) <> "OK" Then
									failCnt = failCnt + 1
									SumErrStr = SumErrStr & iErrStr
								Else
									SumOKStr = SumOKStr & iErrStr
								End If
							End If
						End If
					End If

					'������ ��ǰ�߰� ��찡 �ֱ� ������ ���Ȯ��
					Call fnLtiMallStockChk(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					strParam = ""
					strParam = oiMall.FOneItem.getLotteiMallItemEditParameter()
					Call fnLtiMallInfoEdit(itemid, strParam, iErrStr, FALSE)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'���û�ǰ ��ȸ�ؼ� ���� ��ǰ���¸� ��������
					Call fnCheckLtiMallItemStat(itemid, iErrStr, getLtimallGoodno(itemid))
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					''���� �� ��ǰ �����ȸ�� �Դ����� 2018/12/17 �߰�
					Call fnLtiMallStockChk(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					'OK�� ERR�̴� editQuecnt�� + 1�� ��Ŵ..
					'�����ٸ����� editQuecnt ASC, i.lastupdate DESC�� �ߺ��� ����
					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_ltimall_regitem SET " & VBCRLF
					strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
					strSql = strSql & " ,LTiMallLastupdate = getdate()  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql
				End If
		    Else
				iErrstr = "ERR||"&itemid&"||���������� ��ǰ�� �ƴմϴ�."
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			End If
			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("lotteimall", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oiMall = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=EDIT
	ElseIf action = "EDIT2" Then								'���ο�����ǰ ����
		SET oiMall = new CLotteiMall
			oiMall.FRectItemID	= itemid
			oiMall.getLtimallEditOneItem

			strParam = ""
			strParam = oiMall.FOneItem.getLotteiMallItemEditParameter2()
			Call fnLtiMallInfoEdit(itemid, strParam, iErrStr, TRUE)
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
			End If
		SET oiMall = nothing
	ElseIf action = "REG" Then									'��ǰ ���
		SET oiMall = new CLotteiMall
			oiMall.FRectItemID	= itemid
			oiMall.getLtimallNotRegOneItem
		    If (oiMall.FResultCount < 1) Then
				lastErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
				response.write "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			Else
				strSql = ""
				strSql = strSql & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_LTiMall_regitem where itemid="&itemid&")"
				strSql = strSql & " BEGIN"& VbCRLF
				strSql = strSql & " 	INSERT INTO db_item.dbo.tbl_LTiMall_regitem "
		        strSql = strSql & " 	(itemid, regdate, reguserid, LTiMallstatCD)"
		        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1')"
				strSql = strSql & " END "
			    dbget.Execute strSql

				If oiMall.FOneItem.checkNotRegWords = "N" Then
					iErrStr = "ERR||"&itemid&"||��ϺҰ� �ܾ� ����(����, 1+1, ����, ����)"
				Else
					'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
					If oiMall.FOneItem.checkTenItemOptionValid Then
						strParam = ""
						strParam = oiMall.FOneItem.getLotteiMallItemRegParameter(FALSE)
						Call LotteiMallItemReg(itemid, strParam, iErrStr, oiMall.FOneItem.FSellCash, oiMall.FOneItem.getLotteiMallSellYn)
					Else
						iErrStr = "ERR||"&itemid&"||�ɼǰ˻� ����"
					End If
				End If
				lastErrStr = iErrStr
				response.write iErrStr
				If LEFT(iErrStr, 2) <> "OK" Then
					CALL Fn_AcctFailTouch("lotteimall", itemid, iErrStr)
				End If
			End If
		SET oiMall = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/LtimallProc.asp?itemid=279397&mallid=lotteimall&action=REG
	End If
End If
'###################################################### LotteiMall API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
