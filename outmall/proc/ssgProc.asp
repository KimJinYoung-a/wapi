<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/ssg/ssgItemcls.asp"-->
<!-- #include virtual="/outmall/ssg/incssgFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, failCnt, oSsg, getMustprice, tSsgGoodNo, vOptCnt, chgImageNm, chgSellYn
Dim iErrStr, strParam, mustPrice, strSql, SumErrStr, SumOKStr, endItemErrMsgReplace
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
'######################################################## ssg API ########################################################
If mallid = "ssg" Then
	If action = "REG" Then					'��ǰ���
		SET oSsg = new CSsg
			oSsg.FRectItemID	= itemid
			oSsg.getSsgNotRegOneItem
			If (oSsg.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			ElseIf (oSsg.FOneItem.FMapcnt = 0) Then
				iErrStr = "ERR||"&itemid&"||ī�װ� ��Ī�� �ʿ��մϴ�."
			Else
				strSql = ""
				strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_ssg_regitem where itemid="&itemid&")"
				strSql = strSql & " BEGIN"& VbCRLF
				strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_ssg_regitem "
				strSql = strSql & " 	(itemid, regdate, reguserid, ssgstatCD, regitemname, ssgSellYn)"
				strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oSsg.FOneItem.FItemName)&"', 'N')"
				strSql = strSql & " END "
				dbget.Execute strSql
				'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
				If oSsg.FOneItem.checkTenItemOptionValid Then
					getMustprice = ""
					getMustprice = oSsg.FOneItem.MustPrice()

					strParam = ""
					strParam = oSsg.FOneItem.getssgItemRegParameter(getMustprice)
					Call fnSsgItemReg(itemid, strParam, iErrStr, getMustprice, oSsg.FOneItem.FbasicimageNm, oSsg.FOneItem.getSSGMargin)
				Else
					iErrStr = "ERR||"&itemid&"||[��ǰ���] �ɼǰ˻� ����"
				End If
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("ssg", itemid, iErrStr)
			End If
		SET oSsg = nothing
	ElseIf action = "SOLDOUT" Then			'���º���
		SET oSsg = new Cssg
			oSsg.FRectItemID	= itemid
			oSsg.FRectMustSellyn= "Y"
			oSsg.getSsgEditOneItem

		    If (oSsg.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||[��ǰ����] ���� ������ ��ǰ�� �ƴմϴ�."
			Else
				strParam = ""
				strParam = oSsg.FOneItem.getssgItemEditSellynParameter("N")
				getMustprice = ""
				getMustprice = oSsg.FOneItem.MustPrice()
				If oSsg.FOneItem.isImageChanged Then
					chgImageNm = oSsg.FOneItem.getBasicImage
				Else
					chgImageNm = "N"
				End If
				Call fnSsgItemEditSellyn(itemid, oSsg.FOneItem.FSsgGoodNo, iErrStr, strParam, getMustprice, html2db(oSsg.FOneItem.FItemName), "N", chgImageNm)
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("ssg", itemid, iErrStr)
			End If
		SET oSsg = nothing
		'http://wapi.10x10.co.kr/outmall/proc/ssgProc.asp?itemid=325046&mallid=ssg&action=SOLDOUT
	ElseIf action = "CHKSTAT" Then			'����Ȯ��
		tSsgGoodNo = getSsgGoodNo(itemid)
		Call fnSsgStatChk(itemid, tSsgGoodNo, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("ssg", itemid, iErrStr)
		End If
		'http://wapi.10x10.co.kr/outmall/proc/ssgProc.asp?itemid=325046&mallid=ssg&action=CHKSTAT
	ElseIf (action = "EDIT") OR (action = "PRICE") Then		'���� �� ��ǰ����
		SET oSsg = new Cssg
			oSsg.FRectItemID	= itemid
			oSsg.getSsgEditOneItem
		    If (oSsg.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||[��ǰ����] ���� ������ ��ǰ�� �ƴմϴ�."
			Else
				If (oSsg.FOneItem.getiszeroWonSoldOut(itemid) = "Y") OR (oSsg.FOneItem.FmaySoldOut = "Y") OR (oSsg.FOneItem.IsMayLimitSoldout = "Y") OR (oSsg.FOneItem.IsSoldOut) OR (oSsg.FOneItem.FOptionCnt = 0 AND oSsg.FOneItem.getRegedOptionCnt > 0) Then
					chgSellYn = "N"
				Else
					chgSellYn = "Y"
				End If

				If chgSellYn = "N" Then
					strParam = ""
					strParam = oSsg.FOneItem.getssgItemEditSellynParameter(chgSellYn)
					getMustprice = ""
					getMustprice = oSsg.FOneItem.MustPrice()
					If oSsg.FOneItem.isImageChanged Then
						chgImageNm = oSsg.FOneItem.getBasicImage
					Else
						chgImageNm = "N"
					End If
					Call fnSsgItemEditSellyn(itemid, oSsg.FOneItem.FSsgGoodNo, iErrStr, strParam, getMustprice, html2db(oSsg.FOneItem.FItemName), chgSellYn, chgImageNm)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					strParam = ""
					strParam = oSsg.FOneItem.getssgItemEditParameter(chgSellYn)
					getMustprice = ""
					getMustprice = oSsg.FOneItem.MustPrice()
					If oSsg.FOneItem.isImageChanged Then
						chgImageNm = oSsg.FOneItem.getBasicImage
					Else
						chgImageNm = "N"
					End If
					Call fnSsgItemEdit(itemid, oSsg.FOneItem.FSsgGoodNo, iErrStr, strParam, getMustprice, html2db(oSsg.FOneItem.FItemName), chgSellYn, chgImageNm)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					If failCnt > 0 Then
						endItemErrMsgReplace = replace(SumErrStr, "OK||"&itemid&"||", "")
						endItemErrMsgReplace = replace(SumErrStr, "ERR||"&itemid&"||", "")

						If (Instr(endItemErrMsgReplace, "�ߺ��ȿɼ��������մϴ�") > 0) OR (Instr(endItemErrMsgReplace, "�ߺ� �� �ɼ���") > 0) Then
							strParam = ""
							strParam = oSsg.FOneItem.getssgItemEditSellynParameter("X")
							Call fnSsgItemEditSellyn(itemid, oSsg.FOneItem.FSsgGoodNo, iErrStr, strParam, getMustprice, html2db(oSsg.FOneItem.FItemName), "X", chgImageNm)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						End If
					Else
						Call fnViewItemInfo(itemid, oSsg.FOneItem.FSsgGoodNo, iErrStr)
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
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ssg_regitem SET " & VBCRLF
					strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
					strSql = strSql & " ,ssglastupdate = getdate()  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql
				End If
			End If
			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("ssg", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oSsg = nothing
		'http://wapi.10x10.co.kr/outmall/proc/ssgProc.asp?itemid=325046&mallid=ssg&action=EDIT
	End If
End If
'###################################################### ssg API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->