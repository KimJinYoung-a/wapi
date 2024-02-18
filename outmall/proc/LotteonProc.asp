<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/lotteon/lotteonItemcls.asp"-->
<!-- #include virtual="/outmall/lotteon/inclotteonFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, mallid, action, failCnt, oLotteon, getMustprice, chgSellYn, vOptCnt, i, isChkStat, addOptErrItem
Dim iErrStr, strParam, mustPrice, strSql, SumErrStr, SumOKStr, chgImageNm, arrRows, errVendorItemId
Dim jenkinsBatchYn, idx, lastErrStr
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
jenkinsBatchYn	= request("jenkinsBatchYn")
idx				= request("idx")
lastErrStr		= ""
addOptErrItem	= "N"

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
'######################################################## Lotteon API ########################################################
If mallid = "lotteon" Then
	If action = "REG" Then						'��ǰ ���
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
					getMustprice = ""
					getMustprice = olotteon.FOneItem.MustPrice()
					CALL fnLotteonItemReg(itemid, strParam, iErrStr, getMustprice, olotteon.FOneItem.getLotteonSellYn, olotteon.FOneItem.FLimityn, olotteon.FOneItem.FLimitNo, olotteon.FOneItem.FLimitSold, html2db(olotteon.FOneItem.FItemName), olotteon.FOneItem.FbasicimageNm, "")
				Else
					iErrStr = "ERR||"&itemid&"||[��ǰ���] �ɼǰ˻� ����"
				End If
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
			End If
		'http://wapi.10x10.co.kr/outmall/proc/LotteonProc.asp?itemid=795138&mallid=lotteon&action=REG
		SET olotteon = nothing
	ElseIf action = "CHKSTAT" Then				'��ǰ �� ��ȸ
		SET olotteon = new CLotteon
			olotteon.FRectItemID	= itemid
			olotteon.getLotteonNotEditOneItem

			If (olotteon.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ȸ ������ ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				strParam = olotteon.FOneItem.getLotteonItemViewParameter()
				CALL fnLotteonItemView(itemid, strParam, iErrStr, "")
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
			End If
			'http://wapi.10x10.co.kr/outmall/proc/LotteonProc.asp?itemid=396356&mallid=lotteon&action=CHKSTAT
		SET olotteon = nothing
	ElseIf action = "EDIT" Then								'��ǰ ����
		SET olotteon = new CLotteon
			olotteon.FRectItemID	= itemid
			olotteon.getLotteonNotEditOneItem

			If (olotteon.FResultCount < 1) Then
				failCnt = failCnt + 1
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

				If Not IsArray(arrRows) then 
					strParam = ""
					strParam = olotteon.FOneItem.getLotteonItemViewParameter()				'��ǰ �� ��ȸ
					CALL fnLotteonItemView(itemid, strParam, iErrStr, "")
					response.flush
					response.clear
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					If UBound(arrRows,2) = 0 AND arrRows(0,0) = "Z1" Then
						addOptErrItem = "Y"
					End If
				End If

				If (oLotteon.FOneItem.FmaySoldOut = "Y") OR (oLotteon.FOneItem.IsMayLimitSoldout = "Y") OR (oLotteon.FOneItem.IsSoldOut) OR (oLotteon.FOneItem.FOptionCnt = 0 AND oLotteon.FOneItem.getRegedOptionCnt > 0)  OR (oLotteon.FOneItem.FLimityn = "Y" AND (oLotteon.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) OR (addOptErrItem = "Y") Then
					chgSellYn = "N"
				Else
					chgSellYn = "Y"
				End If

				If chgSellYn = "N" Then
					strParam = ""
					strParam = olotteon.FOneItem.getLotteonSellynParameter(chgSellYn)
					Call fnLotteOnSellyn(itemid, chgSellYn, strParam, iErrStr, "")
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
						Call fnLotteOnSellyn(itemid, chgSellYn, strParam, iErrStr, "")
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
					CALL fnLotteonItemView(itemid, strParam, iErrStr, "")
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
					' 	CALL fnLotteonItemEdit(itemid, olotteon.FOneItem.FItemName, strParam, iErrStr, "")
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
					' 	Call fnLotteOnPrice(itemid, strParam, getMustprice, iErrStr, "")
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
					' 	Call fnLotteOnQuantity(itemid, strParam, iErrStr, "")
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
						CALL fnLotteonItemEdit2(itemid, olotteon.FOneItem.FItemName, getMustprice, strParam, iErrStr, "")
						response.flush
						response.clear
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
					'################## ���� ����..2020-05-07 ���� #######################################

					If failCnt = 0 Then
						strParam = ""
						strParam = olotteon.FOneItem.getLotteonOptStatusParameter()			'��ǰ �ǸŻ��� ����
						Call fnLotteOnOptStat(itemid, strParam, iErrStr, "")
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
						CALL fnLotteonItemView(itemid, strParam, iErrStr, "")
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
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_lotteon_regitem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
			'http://wapi.10x10.co.kr/outmall/proc/LotteonProc.asp?itemid=1933230&mallid=lotteon&action=EDIT
		SET olotteon = nothing
	ElseIf action = "PRICE" Then				'��ǰ ���� ����
		SET olotteon = new CLotteon
			olotteon.FRectItemID	= itemid
			olotteon.getLotteonNotEditOneItem
			If (olotteon.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
			Else
				If LEFT(olotteon.FOneItem.FLastStatCheckDate, 10) = "1900-01-01" Then
					strParam = ""
					strParam = olotteon.FOneItem.getLotteonItemViewParameter()			'��ǰ �� ��ȸ
					CALL fnLotteonItemView(itemid, strParam, iErrStr, "")
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
					Call fnLotteOnPrice(itemid, strParam, getMustprice, iErrStr, "")
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
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
			'http://wapi.10x10.co.kr/outmall/proc/LotteonProc.asp?itemid=795138&mallid=lotteon&action=PRICE
		SET olotteon = nothing
	ElseIf action = "SOLDOUT" Then				'��ǰ ���� ����
		SET olotteon = new CLotteon
			olotteon.FRectItemID	= itemid
			olotteon.getLotteonNotEditOneItem
			If (olotteon.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				strParam = olotteon.FOneItem.getLotteonSellynParameter("N")
				Call fnLotteOnSellyn(itemid, "N", strParam, iErrStr, "")
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
			End If
			'http://wapi.10x10.co.kr/outmall/proc/LotteonProc.asp?itemid=795138&mallid=lotteon&action=SOLDOUT
		SET olotteon = nothing
	ElseIf action = "DELETE" Then				'�Ǹ�����(����)
		SET olotteon = new CLotteon
			olotteon.FRectItemID	= itemid
			olotteon.getLotteonNotEditOneItem
			If (olotteon.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				strParam = olotteon.FOneItem.getLotteonSellynParameter("X")
				Call fnLotteOnSellyn(itemid, "X", strParam, iErrStr, "")
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("lotteon", itemid, iErrStr)
			End If
			'http://wapi.10x10.co.kr/outmall/proc/LotteonProc.asp?itemid=123404&mallid=lotteon&action=DELETE
		SET olotteon = nothing
	End If
End If
'###################################################### Lotteon API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->