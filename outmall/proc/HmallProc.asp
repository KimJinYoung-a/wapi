<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/hmall/hmallItemcls.asp"-->
<!-- #include virtual="/outmall/hmall/inchmallFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, mallid, action, failCnt, oHmall, getMustprice, chgSellYn, i
Dim iErrStr, strParam, mustPrice, strSql, SumErrStr, SumOKStr, arrRows, mrgnRate, chgImageNm, tHmallGoodno
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
'######################################################## Hmall API ########################################################
If mallid = "hmall1010" Then
	If action = "SOLDOUT" Then				'���º���
		SET oHmall = new CHMall
			oHmall.FRectItemID	= itemid
			oHmall.getHmallEditOneItem
			If oHmall.FResultCount = 0 Then
				iErrStr = "ERR||"&itemid&"||���¼��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
			Else
				Call fnHmallSellYN(itemid, "N", iErrStr)
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
			End If
		SET oHmall = nothing
		'http://wapi.10x10.co.kr/outmall/proc/HmallProc.asp?itemid=1079251&mallid=hmall&action=SOLDOUT
	ElseIf action = "REG" Then				'��ǰ���
		SET oHmall = new CHMall
			oHmall.FRectItemID	= itemid
			oHmall.getHmallNotRegOnlyOneItem

			If oHmall.FResultCount > 0 Then
				strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
				dbget.execute strSql

				'If oHmall.FOneItem.fnCheckMakerid Then
				'	iErrStr = "ERR||"&itemid&"||[��ǰ���add] ���̰��� ��ϺҰ�"
				If oHmall.FOneItem.checkTenItemOptionValid Then
					strParam = ""
					strParam = oHmall.FOneItem.gethmallItemRegParameter()
					getMustprice = ""
					getMustprice = oHmall.FOneItem.MustPrice()
					Call fnHmallItemOnlyReg(itemid, strParam, iErrStr, getMustprice, oHmall.FOneItem.gethmallSellYn, oHmall.FOneItem.FLimityn, oHmall.FOneItem.FLimitNo, oHmall.FOneItem.FLimitSold, html2db(oHmall.FOneItem.FItemName), oHmall.FOneItem.FbasicimageNm)
					'Call fnHmallOnlyItemReg(itemid, iErrStr)
				Else
					iErrStr = "ERR||"&itemid&"||[��ǰ���add] �ɼǰ˻� ����"
				End If

				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				If failCnt = 0 Then
					tHmallGoodno = getHmallGoodno(itemid)
					If tHmallGoodno <> "" Then
						chgImageNm = oHmall.FOneItem.getBasicImage
						Call fnHmallImage(itemid, chgImageNm, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				End If

				If failCnt = 0 Then
					tHmallGoodno = getHmallGoodno2(itemid)
					If tHmallGoodno <> "" Then
						Call fnHmallImageConfirm(itemid, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				End If
			Else
				failCnt = 1
				strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
				dbget.execute strSql
				SumErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oHmall = nothing
		'http://wapi.10x10.co.kr/outmall/proc/HmallProc.asp?itemid=1079251&mallid=hmall&action=REG
	ElseIf action = "PRICE" Then			'���ݼ���
		SET oHmall = new CHMall
			oHmall.FRectItemID	= itemid
			oHmall.getHmallEditOneItem
			If oHmall.FResultCount > 0 Then
				mustPrice = ""
				mustPrice = oHmall.FOneItem.MustPrice()

				mrgnRate = ""
				mrgnRate = oHmall.FOneItem.FMrgnRate
				strParam = oHmall.FOneItem.getHmallPriceParameter()
				Call fnHmallPrice(itemid, mustPrice, mrgnRate, iErrStr)
			Else
				iErrStr = "ERR||"&itemid&"||���ݼ��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
			End If
		SET oHmall = nothing
		'http://wapi.10x10.co.kr/outmall/proc/HmallProc.asp?itemid=1079251&mallid=hmall&action=PRICE
	ElseIf action = "IMAGE" Then							'�̹��� ��� & Ȯ��
		tHmallGoodno = getHmallGoodno(itemid)
		If tHmallGoodno = "" Then
			failCnt = 1
			SumErrStr = "ERR||"&itemid&"||��ǰ���� ��� �ϼž� �˴ϴ�."
		Else
			chgImageNm = getTenBasicImage(itemid)
			Call fnHmallImage(itemid, chgImageNm, iErrStr)
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If

		If failCnt = 0 Then
			tHmallGoodno = getHmallGoodno2(itemid)
			If tHmallGoodno = "" Then
				failCnt = failCnt + 1
				SumErrStr = "ERR||"&itemid&"||��ǰ �� �̹������� ��� �ϼž� �˴ϴ�."
			Else
				Call fnHmallImageConfirm(itemid, iErrStr)
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
			CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
		'http://wapi.10x10.co.kr/outmall/proc/HmallProc.asp?itemid=1079251&mallid=hmall&action=IMAGE
	' ElseIf action = "CHKSTAT" Then			'��ǰ��ȸ
	' 	SET oHmall = new CHMall
	' 		oHmall.FRectItemID	= itemid
	' 		oHmall.getHmallEditOneItem
	' 		If oHmall.FResultCount > 0 Then
	' 			Call fnHmallStatChk(itemid, iErrStr)
	' 			If Left(iErrStr, 2) <> "OK" Then
	' 				failCnt = failCnt + 1
	' 				SumErrStr = SumErrStr & iErrStr
	' 			Else
	' 				SumOKStr = SumOKStr & iErrStr
	' 			End If

	' 			If INSTR(iErrStr, "���οϷ�") > 0 AND failCnt = 0 Then
	' 				Call fnHmallOptionStatCheck(itemid, iErrStr)
	' 				If Left(iErrStr, 2) <> "OK" Then
	' 					failCnt = failCnt + 1
	' 					SumErrStr = SumErrStr & iErrStr
	' 				Else
	' 					SumOKStr = SumOKStr & iErrStr
	' 				End If

	' 				strparam = oHmall.FOneItem.fngetOptionEditParam(itemid)
	' 				Call fnHmallOptionEdit(itemid, strparam, iErrStr)
	' 				If Left(iErrStr, 2) <> "OK" Then
	' 					failCnt = failCnt + 1
	' 					SumErrStr = SumErrStr & iErrStr
	' 				Else
	' 					SumOKStr = SumOKStr & iErrStr
	' 				End If

	' 				Call fnHmallOptionStatCheck(itemid, iErrStr)
	' 				If Left(iErrStr, 2) <> "OK" Then
	' 					failCnt = failCnt + 1
	' 					SumErrStr = SumErrStr & iErrStr
	' 				Else
	' 					SumOKStr = SumOKStr & iErrStr
	' 				End If
	' 			End If
	' 		Else
	' 			failCnt = 1
	' 			SumErrStr = "ERR||"&itemid&"||����ȸ �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
	' 		End If

	' 		If failCnt > 0 Then
	' 			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
	' 			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
	' 			CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
	' 			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
	' 			response.write "ERR||"&itemid&"||"&SumErrStr
	' 		Else
	' 			strSql = ""
	' 			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regItem SET " & VBCRLF
	' 			strSql = strSql & " accFailcnt = 0  " & VBCRLF
	' 			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
	' 			dbget.Execute strSql

	' 			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
	' 			lastErrStr = "OK||"&itemid&"||"&SumOKStr
	' 			response.write "OK||"&itemid&"||"&SumOKStr
	' 		End If
	' 	SET oHmall = nothing
	' 	'http://wapi.10x10.co.kr/outmall/proc/HmallProc.asp?itemid=1079251&mallid=hmall&action=CHKSTAT
	ElseIf action = "CHKSTAT" Then			'��ǰ��ȸ
		SET oHmall = new CHMall
			oHmall.FRectItemID	= itemid
			oHmall.getHmallEditOneItem
			If oHmall.FResultCount > 0 Then
				strParam = ""
				strParam = oHmall.FOneItem.getHmallItemConfirmParameter()
				Call fnHmallStatChk2(itemid, strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				If INSTR(iErrStr, "���οϷ�") > 0 AND failCnt = 0 Then
					Call fnHmallOptionStatCheck(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					strparam = oHmall.FOneItem.fngetOptionEditParam(itemid)
					Call fnHmallOptionEdit(itemid, strparam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					Call fnHmallOptionStatCheck(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			Else
				failCnt = 1
				SumErrStr = "ERR||"&itemid&"||����ȸ �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oHmall = nothing
		'http://wapi.10x10.co.kr/outmall/proc/HmallProc.asp?itemid=1079251&mallid=hmall&action=CHKSTAT
	ElseIf action = "OPTSTAT" Then			'��ǰ ��� ��ȸ
		Call fnHmallOptionStatCheck(itemid, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		'http://wapi.10x10.co.kr/outmall/proc/HmallProc.asp?itemid=1079251&mallid=hmall&action=OPTSTAT
	ElseIf action = "EDIT" Then				'��ǰ����
		SET oHmall = new CHMall
			oHmall.FRectItemID	= itemid
			oHmall.getHmallEditOneItem
			If oHmall.FResultCount = 0 Then
				iErrStr = "ERR||"&itemid&"||���� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
			Else
				'If (oHmall.FOneItem.FmaySoldOut = "Y") OR (oHmall.FOneItem.IsMayLimitSoldout = "Y") OR (oHmall.FOneItem.IsAllOptionChange = "Y") OR (oHmall.FOneItem.fnCheckMakerid) Then
				If (oHmall.FOneItem.FmaySoldOut = "Y") OR (oHmall.FOneItem.IsMayLimitSoldout = "Y") OR (oHmall.FOneItem.IsAllOptionChange = "Y") Then
					Call fnHmallSellYN(itemid, "N", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
				'############## Hmall ��ǰ ���� #################
					'2022-05-09 ������ �ϴ� ����
					' Call fnHmallOnlyItemEdit(itemid, iErrStr)
					' If Left(iErrStr, 2) <> "OK" Then
					' 	failCnt = failCnt + 1
					' 	SumErrStr = SumErrStr & iErrStr
					' Else
					' 	SumOKStr = SumOKStr & iErrStr
					' End If
					strParam = ""
					strParam = oHmall.FOneItem.gethmallItemEditParameter()
					Call fnHmallItemOnlyEdit(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				'############## Hmall �̹��� ���� #################
					If oHmall.FOneItem.isImageChanged Then
						chgImageNm = oHmall.FOneItem.getBasicImage
						Call fnHmallImage(itemid, chgImageNm, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If

						Call fnHmallImageConfirm(itemid, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

				'############## Hmall ���� ���� #################
					If failCnt = 0 Then
						mustPrice = ""
						mustPrice = oHmall.FOneItem.MustPrice()

						mrgnRate = ""
						mrgnRate = oHmall.FOneItem.FMrgnRate
						strParam = oHmall.FOneItem.getHmallPriceParameter()
						Call fnHmallPrice(itemid, mustPrice, mrgnRate, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

				'############## Hmall �ɼ� ���� #################
					If failCnt = 0 Then
						strparam = oHmall.FOneItem.fngetOptionEditParam(itemid)
						Call fnHmallOptionEdit(itemid, strparam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

				'############## Hmall ��� ��ȸ #################
					If failCnt = 0 Then
						Call fnHmallOptionStatCheck(itemid, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

				'############## Hmall �Ǹ� ���� ���� #################
					If failCnt = 0 Then
						Call fnHmallSellYN(itemid, "Y", iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				End If
			End If
		SET oHmall = nothing

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regItem SET " & VBCRLF
			strSql = strSql & " accFailcnt = 0  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
		'http://wapi.10x10.co.kr/outmall/proc/HmallProc.asp?itemid=1079251&mallid=hmall&action=EDIT
	End If
End If
'###################################################### Hmall API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->