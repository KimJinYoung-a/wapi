<%
'############################################## ���� �����ϴ� API �Լ� ���� ##############################################
''�Ե����̸� ��ǰ ���
Function LotteiMallItemReg(iitemid, strParam, byRef iErrStr, iSellCash, iLotteSellYn)
	Dim objXML,xmlDOM,strRst, lp, resultcode, resultmsg
	Dim buf, LotteGoodNo, strSql, buf_item_list, pp, OptDesc, StockQty, AssignedRow
	Dim ArgLength, NameValueArr(), j
	On Error Resume Next
	LotteiMallItemReg = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/registApiGoodsInfo.lotte", false
'rw ltiMallAPIURL & "/openapi/registApiGoodsInfo.lotte?"&strparam
'response.end
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)

		If session("ssBctID")="kjy8517" Then
			response.write "<textarea cols=100 rows=30>"&strParam&"</textarea>"
		End If

		If objXML.Status = "200" Then
		    buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			'response.write buf

			If session("ssBctID")="kjy8517" Then
				response.write "<textarea cols=100 rows=30>"&buf&"</textarea>"
			End If

			LotteGoodNo = ""
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf ''BinaryToText(objXML.ResponseBody, "euc-kr")

				resultcode	= xmlDOM.getElementsByTagName("Result").item(0).text
				resultmsg	= xmlDOM.getElementsByTagName("Message").item(0).text
				LotteGoodNo = xmlDOM.getElementsByTagName("goods_no").item(0).text

				If resultcode <> 1 Then
		            iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(��ǰ���)"
				Else
					'��ǰ���翩�� Ȯ��
					strSql = "Select count(itemid) From db_item.dbo.tbl_LTiMall_regItem Where itemid='" & iitemid & "'"
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If rsget(0) > 0 Then
						'// ���� -> ����
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCRLF
						strSql = strSql & "	Set LTiMallLastUpdate = getdate() "  & VbCRLF
						strSql = strSql & "	, LTiMallTmpGoodNo = '" & LotteGoodNo & "'"  & VbCRLF
						strSql = strSql & "	, LTiMallPrice = " &iSellCash& VbCRLF
						strSql = strSql & "	, accFailCnt = 0"& VbCRLF
						strSql = strSql & "	, LTiMallRegdate = isNULL(LTiMallRegdate, getdate())" ''�߰� 2013/02/26
						If (LotteGoodNo <> "") Then
						    strSql = strSql & "	, LTiMallstatCD = '20'"& VbCRLF			'���ο�û
						Else
							strSql = strSql & "	, LTiMallstatCD = '1'"& VbCRLF			'���۽õ�
						End If
						strSql = strSql & "	From db_item.dbo.tbl_LTiMall_regItem R"& VbCRLF
						strSql = strSql & " Where R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
					Else
						'// ���� -> �űԵ��
						strSql = ""
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_LTiMall_regItem "
						strSql = strSql & " (itemid, reguserid, LTiMallRegdate, LTiMallLastUpdate, LTiMallTmpGoodNo, LTiMallPrice, LTiMallSellYn, LTiMallStatCd) VALUES " & VbCRLF
						strSql = strSql & " ('" & iitemid & "'" & VBCRLF
						strSql = strSql & " , '" & session("ssBctId") & "'" &_
						strSql = strSql & " , getdate(), getdate()" & VBCRLF
						strSql = strSql & " , '" & LotteGoodNo & "'" & VBCRLF
						strSql = strSql & " , '" & iSellCash & "'" & VBCRLF
						strSql = strSql & " , '" & iLotteSellYn & "'" & VBCRLF
						If (LotteGoodNo <> "") Then
						    strSql = strSql & ",'20'"
						Else
						    strSql = strSql & ",'10'"
						End If
						strSql = strSql & ")"
						dbget.Execute(strSql)
					End If
					rsget.Close

'					############### ���� �ҽ� �ּ� �� API���μ��� ��� �����ϸ鼭 �Ʒ��� ����..2015-04-17 ���� #############
'					If xmlDOM.getElementsByTagName("Argument").item(18).getAttribute("name")="item_list" Then
'					    buf_item_list = xmlDOM.getElementsByTagName("Argument").item(18).getAttribute("value")
'					Else
'					    buf_item_list = xmlDOM.getElementsByTagName("Argument").item(29).getAttribute("value")
'					End If

					ArgLength = xmlDOM.getElementsByTagName("Argument").length
					Redim NameValueArr(ArgLength, 1)			'Name�� Value ���� �ϳ��� name(1) + Value(1) = 2 => 0���͹Ƿ� 1
					For j=0 to ArgLength
						NameValueArr(j,0) = xmlDOM.getElementsByTagName("Argument")(j).getAttribute("name")
						NameValueArr(j,1) = xmlDOM.getElementsByTagName("Argument")(j).getAttribute("value")
						If NameValueArr(j,0) = "item_list" Then
							buf_item_list = NameValueArr(j,1)
						End If
					Next
'					##########################################################################################################

					pp = 1
					If (buf_item_list <> "") Then
						'rw "["&iitemid&"]=="&LotteGoodNo&"=="&buf_item_list
			            buf_item_list = split(buf_item_list,":")
			            For lp = Lbound(buf_item_list) to Ubound(buf_item_list)
			                OptDesc = split(buf_item_list(lp),",")(0)
			                StockQty = split(buf_item_list(lp),",")(1)
							strSql = ""
							strSql = strSql & " Insert into db_item.dbo.tbl_OutMall_regedoption"
							strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outMallSellyn, outmalllimityn, outMallLimitNo)"
							strSql = strSql & " values("&iitemid
							strSql = strSql & " ,''"
							strSql = strSql & " ,'lotteimall'"
							strSql = strSql & " ,'"&pp&"'"
							strSql = strSql & " ,'"&html2DB(OptDesc)&"'"
							strSql = strSql & " ,'Y'"
							strSql = strSql & " ,'Y'"
							strSql = strSql & " ,"&StockQty
							strSql = strSql & ")"
							dbget.Execute strSql, AssignedRow

							''�ɼ� �ڵ� ��Ī.
							If (AssignedRow>0) Then
								strSql = ""
								strSql = strSql & " update oP"   & VBCRLF
								strSql = strSql & " SET itemoption = O.itemoption " & VBCRLF
								strSql = strSql & " FROM db_item.dbo.tbl_OutMall_regedoption oP " & VBCRLF
								strSql = strSql & " JOIN db_item.dbo.tbl_item_option o on oP.itemid=o.itemid " & VBCRLF
								strSql = strSql & " WHERE oP.mallid = 'lotteimall' " & VBCRLF
								strSql = strSql & " and o.itemid = "&iitemid & VBCRLF
								strSql = strSql & " and oP.itemid = "&iitemid & VBCRLF
								strSql = strSql & " and op.outmallOptCode = '"&pp&"'" & VBCRLF
								strSql = strSql & " and op.outmallOptName = o.optionname" & VBCRLF
								dbget.Execute strSql, AssignedRow
							End If
							pp = pp + 1
						Next
						strSql = ""
						strSql = strSql & " UPDATE R " & VBCRLF
						strSql = strSql & " SET regedOptCnt = isNULL(T.CNT,0) " & VBCRLF
						strSql = strSql & " FROM db_item.dbo.tbl_LTiMall_regItem R " & VBCRLF
						strSql = strSql & " Join ( " & VBCRLF
						strSql = strSql & " 	SELECT R.itemid, count(*) as CNT FROM db_item.dbo.tbl_LTiMall_regItem R " & VBCRLF
						strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = 'lotteimall' and Ro.itemid = " &iitemid & VBCRLF
						strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
						strSql = strSql & " ) T on R.itemid = T.itemid " & VBCRLF
						dbget.Execute strSql
					End If
					iErrStr =  "OK||"&iitemid&"||��ϼ���(��ǰ���)"
				End If
			Set xmlDOM = Nothing
			LotteiMallItemReg= true
		Else
			iErrStr = "ERR||"&iitemid&"||LotteiMall ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'ǰ�� ���� �Լ�
Public Function fnLtiMallSellyn(iitemid, ichgSellYn, istrParam, byRef iErrStr)
    Dim strParam, resultcode, resultmsg
    Dim objXML, xmlDOM
    Dim strRst, strSql, buf
    fnLtiMallSellyn = False
	on Error resume next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", ltiMallAPIURL & "/openapi/updateGoodsSaleStat.lotte" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
'rw istrParam
'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") or (session("ssBctID")="skyer9") Then
					''response.write "�׽�Ʈ �α� ��� ����-�Ǹ����� ����<br />"
					''response.write buf & "<br />"
					''response.write "�׽�Ʈ �α� ��� ����-�Ǹ����� ����<br />"
				End If

				resultcode = xmlDOM.getElementsByTagName("Result").item(0).text
				resultmsg = xmlDOM.getElementsByTagName("Message").item(0).text

				'// ���� ����
				If resultcode <> 1 Then
		            iErrStr = "ERR||"&iitemid&"||"&resultmsg
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_ltiMall_regItem " & VbCRLF
					strSql = strSql & " SET LtiMallLastUpdate = getdate() " & VbCRLF
					strSql = strSql & " ,LtiMallSellYn = '" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " ,accFailCnt = 0 " & VbCRLF
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute(strSql)
					If ichgSellYn = "N" Then
						iErrStr = "OK||"&iitemid&"||ǰ��ó��"
					ElseIf ichgSellYn = "Y" Then
						iErrStr = "OK||"&iitemid&"||�Ǹ������� ����"
					Else
						strSql = ""
						strSql = strSql &" INSERT INTO [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] " & VBCRLF
						strSql = strSql &" SELECT TOP 1 'lotteimall', i.itemid, r.ltimallGoodNo, r.ltimallRegdate, getdate(), r.lastErrStr" & VBCRLF
						strSql = strSql &" FROM db_item.dbo.tbl_item as i " & VBCRLF
						strSql = strSql &" JOIN db_item.dbo.tbl_ltimall_regitem as r on i.itemid = r.itemid " & VBCRLF
						strSql = strSql &" WHERE i.itemid = "&iitemid & VBCRLF
						dbget.Execute(strSql)

						strSql = ""
						strSql = strSql & " DELETE FROM db_item.dbo.tbl_ltimall_regitem " & vbcrlf
						strSql = strSql & " WHERE itemid = '"&iitemid&"' "
						dbget.Execute(strSql)

						strSql = ""
						strSql = strSql & " DELETE FROM db_item.dbo.tbl_outmall_regedoption " & vbcrlf
						strSql = strSql & " WHERE itemid = '"&iitemid&"' " & vbcrlf
						strSql = strSql & " and mallid = '"&CMALLNAME&"' " & vbcrlf
						dbget.Execute(strSql)

						strSql = ""
						strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_outmall_API_Que " & vbcrlf
						strSql = strSql & " WHERE itemid = '"&iitemid&"' " & vbcrlf
						strSql = strSql & " and mallid = '"&CMALLNAME&"' " & vbcrlf
						dbget.Execute(strSql)
						iErrStr = "OK||"&iitemid&"||�Ǹ�����"
					End If
		        End If
			Set xmlDOM = Nothing
			fnLtiMallSellyn = True
		Else
			iErrStr = "ERR||"&iitemid&"||LotteiMall ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-SELLEDIT-001]"
		End If
	Set objXML = Nothing
	on Error Goto 0
End Function

'���û�ǰ �ǸŰ� ����
Public Function fnLtimallPrice(iitemid, istrParam, imustprice, byRef iErrStr)
    Dim objXML,xmlDOM,strRst
    Dim buf, strSql
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnLtimallPrice = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/updateGoodsSalePrcOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)

		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				resultcode = xmlDOM.getElementsByTagName("Result").item(0).text
				resultmsg = xmlDOM.getElementsByTagName("Message").item(0).text

				If resultcode <> 1 Then
		            iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(��ǰ����)"
		            fnLtimallPrice = False
				Else
				    '// ��ǰ�������� ����
				    strSql = ""
	    			strSql = strSql & " UPDATE db_item.dbo.tbl_ltiMall_regItem  " & VbCRLF
	    			strSql = strSql & "	SET ltimallLastUpdate=getdate() " & VbCRLF
	    			strSql = strSql & "	, ltimallPrice = " & imustprice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " Where itemid='" & iitemid & "'"& VbCRLF
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||��������(��ǰ����)"
					fnLtimallPrice = True
				End If
			Set xmlDOM = Nothing
		Else
			fnLtimallPrice = False
			iErrStr = "ERR||"&iitemid&"||LotteiMall ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRICE-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Public Function fnLtiMallChgItemname(iitemid, strParam, iErrStr)
	Dim objXML, xmlDOM, strRst, resultmsg, resultcode, strSql
	On Error Resume Next
	fnLtiMallChgItemname = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/updateGoodsNmOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)

		If (session("ssBctID")="kjy8517") Then
			'rw ltiMallAPIURL & "/openapi/updateGoodsNmOpenApi.lotte"
			'rw strParam
		End If

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				resultcode = xmlDOM.getElementsByTagName("Result").item(0).text
			    resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text

				If resultcode <> 1 Then
		            iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(��ǰ��)"
		            fnLtiMallChgItemname = False
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_ltiMall_regItem " & VbCRLF
					strSql = strSql & " SET regitemname = B.itemname "& VbCRLF
					strSql = strSql & " FROM db_item.dbo.tbl_ltiMall_regItem A "& VbCRLF
					strSql = strSql & " JOIN db_item.dbo.tbl_item B on A.itemid = B.itemid "& VbCRLF
					strSql = strSql & " WHERE A.itemid='" & iitemid & "'"& VbCRLF
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||��������(��ǰ��)"
					fnLtiMallChgItemname = True
			    End If
			Set xmlDOM = Nothing
		Else
			fnLtiMallChgItemname = False
			iErrStr = "ERR||"&iitemid&"||LotteiMall ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-NMEDIT-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Public Function fnLtiMallstatChk(iitemid, iErrStr)
	Dim objXML,xmlDOM,strRst,resultmsg, iLotteGoodNo, strSql
	Dim strParam, iLotteTmpID, SaleStatCd, GoodsViewCount
	Dim iRbody, ltimallStatName
	On Error Resume Next
	fnLtiMallstatChk = False
	iLotteTmpID = getLtiMallTmpItemIdByTenItemID(iitemid)

	If (iLotteTmpID = "") OR (iLotteTmpID = "���û�ǰ") then
		iErrStr =  "ERR||"&iitemid&"||�̹� ���û�ǰ �Դϴ�.(�űԻ�ǰ��ȸ)"
		Exit function
	End If

	strParam = "subscriptionId=" & ltiMallAuthNo & "&search_type=3&search_value="&iLotteTmpID
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/searchNewGoodsInfoOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				GoodsViewCount 	= Trim(xmlDOM.getElementsByTagName("GoodsCount").item(0).text)		'�˻���
				resultmsg 		= xmlDOM.getElementsByTagName("Message").Item(0).Text
				iLotteGoodNo	= Trim(xmlDOM.getElementsByTagName("GoodsNo").item(0).text)			'���û�ǰ��ȣ
				SaleStatCd		= Trim(xmlDOM.getElementsByTagName("ConfStatCd").item(0).text)		'���������ڵ�

				If GoodsViewCount <> 1 Then
					If resultmsg = "" Then
						resultmsg = "��ȸ��� ����"
					End If
		            iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(�űԻ�ǰ��ȸ)"
		            fnLtiMallstatChk = False
				Else

					Select Case SaleStatCd
						Case "10"	ltimallStatName = "�ӽõ��"
						Case "20"	ltimallStatName = "���ο�û"
						Case "30"	ltimallStatName = "���οϷ�"
						Case "40"	ltimallStatName = "�ݷ�"
						Case "50"	ltimallStatName = "���κҰ�"
						Case "51"	ltimallStatName = "����ο�û"
						Case "52"	ltimallStatName = "������û"
						Case "60"	ltimallStatName = "���"
					End Select

					If SaleStatCd = "30" Then				'���οϷ�(LtiMallStatCd, LtiMallGoodNo, lastConfirmdate ����)
						strSql = ""
						strSql = strSql & " UPDATE db_item.dbo.tbl_ltiMall_regItem " & VbCRLF
						strSql = strSql & " SET lastConfirmdate = getdate() "& VbCRLF
						strSql = strSql & "	,LtiMallStatCd='7' "
						strSql = strSql & " ,LtiMallGoodNo='" & iLotteGoodNo & "' "
						strSql = strSql & " WHERE itemid='" & iitemid & "'"& VbCRLF
						dbget.Execute(strSql)
					Else
						strSql = ""
						strSql = strSql & " UPDATE db_item.dbo.tbl_ltiMall_regItem " & VbCRLF
						strSql = strSql & " SET lastConfirmdate = getdate() "& VbCRLF
						strSql = strSql & "	,LtiMallStatCd='"&SaleStatCd&"' "& VbCRLF
						strSql = strSql & " WHERE itemid='" & iitemid & "'"& VbCRLF
						dbget.Execute(strSql)
					End If
					iErrStr =  "OK||"&iitemid&"||����(�űԻ�ǰ��ȸ) : "&ltimallStatName
					fnLtiMallstatChk = True
			    End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-STATCHK-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Public Function fnLtiMallStockChk(iitemid, iErrStr)
	Dim ilottegoods_no
	Dim objXML,xmlDOM,strRst, iMessage
	Dim ProdCount, buf, AssignedRow, oneProdInfo, strParam
	Dim GoodNo,ItemNo,OptDesc,DispYn,SaleStatCd,StockQty, bufopt
	Dim strSql, actCnt, CorpItemNo, getRegOptCD, SubNodes

	On Error Resume Next
	fnLtiMallStockChk = False
	ilottegoods_no = getLtimallGoodno(itemid)

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		strParam = "?subscriptionId=" & ltiMallAuthNo				'�Ե����̸� ������ȣ	(*)
		strParam = strParam & "&search_type=goods_no"
		strParam = strParam & "&search_value=" & ilottegoods_no		'�Ե����̸� ��ǰ��ȣ	(*)

		objXML.Open "GET", ltiMallAPIURL & "/openapi/searchStockList.lotte"&strParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML replace(buf,"&","��")
				ProdCount   = Trim(xmlDOM.getElementsByTagName("GoodsCount").item(0).text)   '' ��ǰ ����

				If (ProdCount <> "") Then
			        Set oneProdInfo = xmlDOM.getElementsByTagName("GoodsInfoList")
			        ' strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and itemoption='')"
			        ' strSql = strSql & " BEGIN"
			        ' strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and itemoption=''"
			        ' strSql = strSql & " END"
			        ' dbget.Execute strSql

			        ' ''2013/05/30 �߰�
			        ' strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and Len(outmalloptCode)>6)"
			        ' strSql = strSql & " BEGIN"
			        ' strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and Len(outmalloptCode)>6"
			        ' strSql = strSql & " END"
			        ' dbget.Execute strSql

					''���� ����ȭ 2018/12/17
					strSql = "DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and (itemoption='' or Len(outmalloptCode)>6)"
					dbget.Execute strSql

					For each SubNodes in oneProdInfo
						GoodNo	    = Trim(SubNodes.getElementsByTagName("GoodNo").item(0).text)
					    ItemNo	    = Trim(SubNodes.getElementsByTagName("ItemNo").item(0).text)        '' ��ǰ�ڵ� (���� 0,1,2,)
					    OptDesc	    = Trim(SubNodes.getElementsByTagName("OptDesc").item(0).text)
					    SaleStatCd	= Trim(SubNodes.getElementsByTagName("SaleStatCd").item(0).text) ''�Ǹ�����, �Ǹ�����, ǰ��'
					    StockQty	= Trim(SubNodes.getElementsByTagName("StockQty").item(0).text)
                        CorpItemNo  = Trim(SubNodes.getElementsByTagName("CorpItemNo").item(0).text)  '' ��ǰ�ڵ�_�ɼ��ڵ�

						getRegOptCD = Split(CorpItemNo,"_")(1)
					    OptDesc = replace(OptDesc, "��", "&")
					    If (SaleStatCd <> "10") Then
					        DispYn = "N"
					    else
					        DispYn = "Y"
					    End If

					    If (StockQty = "null") Then
					        StockQty = "0"
					    End If

					    bufopt = OptDesc
						If InStr(bufopt,",") > 0 then
						    If (splitValue(bufopt,",",0) <> "") Then
						        OptDesc = splitValue(splitValue(bufopt,",",0),":",1)
						    End If

						    If (splitValue(bufopt,",",1) <> "") Then
						        OptDesc = OptDesc+","+splitValue(splitValue(bufopt,",",1),":",1)
						    End If

						    If (splitValue(bufopt,",",2)<>"") Then
						        OptDesc = OptDesc+","+splitValue(splitValue(bufopt,",",2),":",1)
						    End If
						Else
							OptDesc = splitValue(OptDesc,":",1)
						End If

						'################  2014-02-21 14:34 ������ ######################
						'OptDesc = replace(OptDesc, ",,", ",")���� �߰�..
						'���� : [db_item].[dbo].tbl_item_option_Multiple �� optionTypeName�� ,�� �� �ִ� ��찡 ����..
 						'ex)�̴ϼȰ���(5,000��)..�̷��� �Ǿ����� �� ,�� split��Ŵ�� ���� ,,�� ,�� ġȯ��
 						OptDesc = replace(OptDesc, ",,", ",")
 						'#################################################################

						''response.write "�׽�Ʈ �α� ��� ����-4-1<br />"
						''rw GoodNo&"|"&ItemNo&"|"&CorpItemNo&"|"&OptDesc&"|"&DispYn&"|"&SaleStatCd&"|"&StockQty & "<br />"
						''response.write "�׽�Ʈ �α� ��� ����-4-1<br />"
						strSql = ""
						strSql = strSql & " UPDATE oP "
					    strSql = strSql & " SET outmallOptName='"&html2DB(OptDesc)&"'"&VbCRLF
						strSql = strSql & " ,outmallOptCode='"&ItemNo&"'"&VbCRLF
					    strSql = strSql & " ,lastupdate=getdate()"&VbCRLF
					    strSql = strSql & " ,outMallSellyn='"&DispYn&"'"&VbCRLF
					    strSql = strSql & " ,outmalllimityn='Y'"&VbCRLF
					    strSql = strSql & " ,outMallLimitNo="&StockQty&VbCRLF
					    strSql = strSql & " ,checkdate=getdate()"&VbCRLF
					    strSql = strSql & " FROM db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
					    strSql = strSql & " WHERE itemid="&iitemid&VbCRLF
					    strSql = strSql & " and convert(int, outmallOptCode)='"&ItemNo&"'"&VbCRLF				'������ outmallOptCode�� 001,002,003 �̷��� �������� ���� �Ŀ� 1,2,3�̷��� ����
					    strSql = strSql & " and mallid='lotteimall'"&VbCRLF
					    dbget.Execute strSql, AssignedRow
					    If (AssignedRow < 1) Then
							''������ �̹� ��������? �ּ�ó�� 2018/12/17
					        ' strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and itemoption='')"
					        ' strSql = strSql & " BEGIN"
					        ' strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteimall' and itemid="&iitemid&" and itemoption=''"
					        ' strSql = strSql & " END"
					        ' dbget.Execute strSql

					        strSql = " Insert into db_item.dbo.tbl_OutMall_regedoption"
					        strSql = strSql & " (itemid,itemoption,mallid,outmallOptCode,outmallOptName,outMallSellyn,outmalllimityn,outMallLimitNo,checkdate)"
					        strSql = strSql & " values("&iitemid
					        strSql = strSql & " ,'"&ItemNo&"'" ''�ӽ÷� �Ե� �ڵ� ���� //2013/04/01
					        strSql = strSql & " ,'lotteimall'"
					        strSql = strSql & " ,'"&ItemNo&"'"
					        strSql = strSql & " ,'"&html2DB(OptDesc)&"'"
					        strSql = strSql & " ,'"&DispYn&"'"
					        strSql = strSql & " ,'Y'"
					        strSql = strSql & " ,"&StockQty
					        strSql = strSql & " ,getdate()"
					        strSql = strSql & ")"
					        dbget.Execute strSql, AssignedRow

							If getRegOptCD = "" Then
								Dim newOptSQL
								newOptSQL = ""
								newOptSQL = newOptSQL & " SELECT TOP 1 itemoption FROM [db_item].[dbo].tbl_item_option WHERE itemid = '"&iitemid&"' and optionname = '"&html2DB(OptDesc)&"' "

								rsget.CursorLocation = adUseClient
        						rsget.Open newOptSQL, dbget, adOpenForwardOnly, adLockReadOnly
								If Not(rsget.EOF or rsget.BOF) Then
									getRegOptCD = rsget("itemoption")
								Else
									getRegOptCD = "0000"
								End If
								rsget.Close
							End If

					        ''�ɼ� �ڵ� ��Ī.
					        If (AssignedRow > 0) Then
					            strSql = " update oP"   &VbCRLF
					            strSql = strSql & " set itemoption='"&getRegOptCD&"'"&VbCRLF
					            strSql = strSql & " From db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
					            strSql = strSql & " where oP.mallid='lotteimall'"&VbCRLF
					            strSql = strSql & " and oP.itemid="&iitemid&VbCRLF
					            strSql = strSql & " and op.outmallOptCode='"&ItemNo&"'"&VbCRLF
					            dbget.Execute strSql, AssignedRow
					        End If
					        getRegOptCD = ""
					    Else
					    	''���� ��ǰ�� �� tbl_OutMall_regedoption�� �����Ͱ� ������ tbl_item_option�� �����Ͱ� ���⿡ �ϴ� ���ν��� ȣ��
							Dim DanChkArr
							strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_ltimall '"&CMallName&"'," & iitemid
							rsget.CursorLocation = adUseClient
							rsget.CursorType = adOpenStatic
							rsget.LockType = adLockOptimistic
							rsget.Open strSql, dbget
							If Not(rsget.EOF or rsget.BOF) Then
							    DanChkArr = rsget.getRows
							End If
							rsget.close
							If UBound(DanChkArr,2) = 0 AND DanChkArr(0,1) = "0000"  Then

							Else
						        strSql = " update oP"   &VbCRLF
						        strSql = strSql & " set itemoption=o.itemoption"&VbCRLF
						        strSql = strSql & " From db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
						        strSql = strSql & "     Join db_item.dbo.tbl_item_option o"&VbCRLF
						        strSql = strSql & "     on oP.itemid=o.itemid"&VbCRLF
						        strSql = strSql & " where oP.mallid='lotteimall'"&VbCRLF
						        strSql = strSql & " and o.itemid="&iitemid&VbCRLF
						        strSql = strSql & " and oP.itemid="&iitemid&VbCRLF
						        strSql = strSql & " and op.outmallOptCode='"&ItemNo&"'"&VbCRLF
						        strSql = strSql & " and Replace(Replace(op.outmallOptName,' ',''),':','')=Replace(Replace(o.optionname,' ',''),':','')"&VbCRLF
						        dbget.Execute strSql, AssignedRow
							End If
					    End If
					    actCnt = actCnt+AssignedRow
					Next

					''If (actCnt > 0) Then
					    strSql = " update R"   &VbCRLF
			            strSql = strSql & " set regedOptCnt=isNULL(T.optSellYCNT,0)"   &VbCRLF  ''regedOptCnt => optSellYCNT
						strSql = strSql & " ,lastStatcheckdate=getdate()"&VbCRLF 				''�߰�
			            strSql = strSql & " from db_item.dbo.tbl_LTiMall_regItem R"   &VbCRLF
			            strSql = strSql & " 	Join ("   &VbCRLF
			            strSql = strSql & " 		select R.itemid,count(*) as CNT "
			            strSql = strSql & " 		, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
						strSql = strSql & " 		, sum(CASE WHEN itemoption<>'0000' and [outmallSellyn]='Y' and (outmalllimityn='N' or (outmalllimityn='Y' and outmalllimitno>0)) THEN 1 ELSE 0 END) as optSellYCNT"
			            strSql = strSql & "        from db_item.dbo.tbl_LTiMall_regItem R"   &VbCRLF
			            strSql = strSql & " 			Join db_item.dbo.tbl_OutMall_regedoption Ro"   &VbCRLF
			            strSql = strSql & " 			on R.itemid=Ro.itemid"   &VbCRLF
			            strSql = strSql & " 			and Ro.mallid='lotteimall'"   &VbCRLF
			            strSql = strSql & "             and Ro.itemid="&iitemid&VbCRLF
			            strSql = strSql & " 		group by R.itemid"   &VbCRLF
			            strSql = strSql & " 	) T on R.itemid=T.itemid"   &VbCRLF
			            dbget.Execute strSql
					''End If
				End if
				iErrStr =  "OK||"&iitemid&"||����(�����ȸ)"
				fnLtiMallStockChk =true
			Set xmlDOM = Nothing
		Else
		    iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-STOCKCHK-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function fnLtiMallDisView(iitemid,byRef iErrStr, iLottegoodNo)
	Dim objXML, xmlDOM, strRst, resultmsg, assignedRow
	Dim strParam, iLotteItemID , SaleStatCd, GoodsViewCount
	Dim iRbody, LotteSellyn, iSalePrc, iGoodsNm, sqlstr

	fnLtiMallDisView = False
	strParam = "subscriptionId=" & ltiMallAuthNo & "&goods_no="&iLottegoodNo
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/searchGoodsListOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				GoodsViewCount = xmlDOM.getElementsByTagName("GoodsCount").item(0).text
				If (GoodsViewCount = "1") Then
			    	SaleStatCd = xmlDOM.getElementsByTagName("SaleStatCd").item(0).text
			    	iSalePrc	= xmlDOM.getElementsByTagName("SalePrc").item(0).text
			    	iGoodsNm	= xmlDOM.getElementsByTagName("GoodsNm").item(0).text
			    	iGoodsNm	= replace(iGoodsNm,"@@amp@@","&")
					iGoodsNm	= Replace(iGoodsNm,"&gt;",">")
					iGoodsNm	= Replace(iGoodsNm,"&lt;","<")
					iGoodsNm	= Replace(iGoodsNm,"&nbsp;"," ")
					iGoodsNm	= Replace(iGoodsNm,"&amp;","&")

					' response.write "�׽�Ʈ �α� ��� ����<br />"
					' response.write "�Ե����̸� ���û��� : " & SaleStatCd & "<br />"
					' response.write "�׽�Ʈ �α� ��� ����<br />"

					If (SaleStatCd="10") Then
					    LotteSellyn = "Y"
					ElseIf (SaleStatCd="20") Then
					    LotteSellyn = "N"
					ElseIf (SaleStatCd="30") Then
					    LotteSellyn = "X"
					End If

					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_ltimall_regitem SET " & VBCRLF
					strSql = strSql & " DispViewCnt = isnull(DispViewCnt, 0) + 1 " & VBCRLF
					strSql = strSql & " ,ltiMallPrice="&iSalePrc & VbCRLF
					If (LotteSellyn <> "") Then
					    strSql = strSql & " ,ltiMallSellyn='"&LotteSellyn&"'"
					End If
					strSql = strSql & " WHERE itemid = '"&iitemid&"' " & VBCRLF
					dbget.Execute strSql, assignedRow
			    	iErrStr =  "OK||"&iitemid&"||����(���û�ǰ��ȸ)"
					fnLtiMallDisView = True
			    Else
			    	resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text
			    	iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(���û�ǰ��ȸ)"
		            fnLtiMallDisView = False
			    End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-ItemChk-001]"
			fnLtiMallDisView = False
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Public Function fnLtiMallAddOpt(iitemid, istrParam, byRef iErrStr, iAddOptCnt)
	Dim objXML, xmlDOM, strRst, resultmsg, strSql
	Dim addOptCount, opt1Nm, opt2Nm, opt3Nm, opt4Nm, opt5Nm, opt1Tval, opt2Tval, opt3Tval, opt4Tval, opt5Tval
	On Error Resume Next
	fnLtiMallAddOpt = False
	addOptCount = 0
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/addGoodsItemInfo.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				If xmlDOM.getElementsByTagName("itemCount").item(0).text = "" Then
					addOptCount = 0
					iAddOptCnt = 0
				Else
    				addOptCount = xmlDOM.getElementsByTagName("itemCount").item(0).text
    				iAddOptCnt = addOptCount
    			End If

				If addOptCount > 0 Then
					opt1Nm 		= xmlDOM.getElementsByTagName("opt1Nm").item(0).text
					opt1Tval	= xmlDOM.getElementsByTagName("opt1Tval").item(0).text
					opt2Nm 		= xmlDOM.getElementsByTagName("opt2Nm").item(0).text
					opt2Tval	= xmlDOM.getElementsByTagName("opt2Tval").item(0).text
					opt3Nm 		= xmlDOM.getElementsByTagName("opt3Nm").item(0).text
					opt3Tval	= xmlDOM.getElementsByTagName("opt3Tval").item(0).text
					opt4Nm 		= xmlDOM.getElementsByTagName("opt4Nm").item(0).text
					opt4Tval	= xmlDOM.getElementsByTagName("opt4Tval").item(0).text
					opt5Nm 		= xmlDOM.getElementsByTagName("opt5Nm").item(0).text
					opt5Tval	= xmlDOM.getElementsByTagName("opt5Tval").item(0).text
				    resultmsg 	= xmlDOM.getElementsByTagName("Message").Item(0).Text
					If resultmsg = "" Then
						iErrStr =  "OK||"&iitemid&"||����(�ɼ��߰�)"
						fnLtiMallAddOpt = True
				    Else
			            iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(�ɼ��߰�)"
			            fnLtiMallAddOpt = False
				    End If
			    End If
			Set xmlDOM = Nothing
		Else
			fnLtiMallAddOpt = False
			iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-ADDOPT-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

''�Ե����̸� ��ǰ���� ����
Function fnLtiMallInfoEdit(iitemid, strParam, byRef iErrStr, isVer2)
	Dim objXML, xmlDOM, strRst, resultmsg, resultcode
	On Error Resume Next
	fnLtiMallInfoEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	If (isVer2) Then
	    objXML.Open "POST", ltiMallAPIURL & "/openapi/upateApiNewGoodsInfo.lotte", false          ''��ǰ����
	Else
	    objXML.Open "POST", ltiMallAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte", false      ''���û�ǰ����
	End If
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				If session("ssBctID")="kjy8517" Then
					response.write "<textarea cols=100 rows=30>"&BinaryToText(objXML.ResponseBody, "euc-kr")&"</textarea>"
				End If

				' response.write "�׽�Ʈ �α� ��� ����-��ǰ����<br />"
				' response.write BinaryToText(objXML.ResponseBody, "euc-kr")
				' response.write "�׽�Ʈ �α� ��� ����-��ǰ����<br />"

				resultcode = xmlDOM.getElementsByTagName("Result").item(0).text
			    resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text

			    If (resultcode = "1") Then
					iErrStr =  "OK||"&iitemid&"||����(��ǰ����)"
					fnLtiMallInfoEdit = True
				Else
		            iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(��ǰ����)"
		            fnLtiMallInfoEdit = False
			    End If
			Set xmlDOM = Nothing
		Else
			fnLtiMallInfoEdit = False
			iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-EDIT-001]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

''���û�ǰ ��ȸ
Function fnCheckLtiMallItemStat(iitemid,byRef iErrStr, iLottegoodNo)
	Dim objXML, xmlDOM, strRst, resultmsg, assignedRow
	Dim strParam, iLotteItemID , SaleStatCd, GoodsViewCount
	Dim iRbody, LotteSellyn, iSalePrc, iGoodsNm, sqlstr

	fnCheckLtiMallItemStat = False
	strParam = "subscriptionId=" & ltiMallAuthNo & "&goods_no="&iLottegoodNo
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/searchGoodsListOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody

				GoodsViewCount = xmlDOM.getElementsByTagName("GoodsCount").item(0).text
				If (GoodsViewCount = "1") Then
			    	SaleStatCd = xmlDOM.getElementsByTagName("SaleStatCd").item(0).text
			    	iSalePrc	= xmlDOM.getElementsByTagName("SalePrc").item(0).text
			    	iGoodsNm	= xmlDOM.getElementsByTagName("GoodsNm").item(0).text
			    	iGoodsNm	= replace(iGoodsNm,"@@amp@@","&")
					iGoodsNm	= Replace(iGoodsNm,"&gt;",">")
					iGoodsNm	= Replace(iGoodsNm,"&lt;","<")
					iGoodsNm	= Replace(iGoodsNm,"&nbsp;"," ")
					iGoodsNm	= Replace(iGoodsNm,"&amp;","&")

					If (SaleStatCd="10") Then
					    LotteSellyn = "Y"
					ElseIf (SaleStatCd="20") Then
					    LotteSellyn = "N"
					ElseIf (SaleStatCd="30") Then
					    LotteSellyn = "X"
					End If

					sqlstr = "Update R" & VbCRLF
					sqlstr = sqlstr & " SET regitemname='"&html2db(iGoodsNm)&"'"
'					If (LotteSellyn <> "") Then
'					    sqlstr = sqlstr & " ,ltiMallSellyn='"&LotteSellyn&"'"
'					End If
'					sqlstr = sqlstr & " ,LtiMallStatCd=(CASE WHEN isNULL(LtiMallStatCd,-9)<7 THEN 7 ELSE LtiMallStatCd END)" ''2013/09/03 �߰�
'					sqlstr = sqlstr & " ,lastStatCheckDate=getdate()"
					sqlstr = sqlstr & " FROM db_item.dbo.tbl_LTiMall_regItem R" & VbCRLF
					sqlstr = sqlstr & " WHERE R.itemid="&iitemid & VbCRLF
					dbget.Execute sqlstr, assignedRow
			    	iErrStr =  "OK||"&iitemid&"||����(���û�ǰ��ȸ)"
					fnCheckLtiMallItemStat = True
			    Else
			    	resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text
			    	iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(���û�ǰ��ȸ)"
		            fnCheckLtiMallItemStat = False
			    End If
			Set xmlDOM = Nothing
		Else
			iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-ItemChk-001]"
			fnCheckLtiMallItemStat = False
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function fnLtiMallGoodsList(styyyymmdd,edyyyymmdd, byRef iErrStr)
	Dim objXML, xmlDOM, strRst, resultmsg
	Dim strParam, SaleStatCd, GoodsCount, iSalePrc, iGoodsNm, GoodsRegDtime, DispYn, iGoodsNo
	Dim iRbody, LotteSellyn, sqlStr, assignedRow
	Dim oneGoodsInfo, SubNodes
	Dim regDtKey : regDtKey=LEFT(NOW(),10) & " " &FormatDateTime(NOW(),4)&":"&RIGHT("0"&second(time),2)
	dim ii : ii=0
    fnLtiMallGoodsList = False

	strParam = "subscriptionId=" & ltiMallAuthNo
	strParam = strParam & "&req_start_dtime="&styyyymmdd
	strParam = strParam & "&req_end_dtime="&edyyyymmdd

	if styyyymmdd="20180905" then
		 strParam = strParam & "&sale_stat_cd=10"
	end if

	'On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", ltiMallAPIURL & "/openapi/searchGoodsListOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		response.write "."
		response.flush
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				response.write "."
				response.flush
				GoodsCount = xmlDOM.getElementsByTagName("GoodsCount").item(0).text
				If (GoodsCount >= "0") Then
					Set oneGoodsInfo = xmlDOM.getElementsByTagName("GoodsInfoList")
					For each SubNodes in oneGoodsInfo
						iGoodsNo	= SubNodes.getElementsByTagName("GoodsNo").item(0).text
						SaleStatCd  = SubNodes.getElementsByTagName("SaleStatCd").item(0).text
						iSalePrc	= SubNodes.getElementsByTagName("SalePrc").item(0).text
						iGoodsNm	= SubNodes.getElementsByTagName("GoodsNm").item(0).text
						iGoodsNm	= replace(iGoodsNm,"@@amp@@","&")
						iGoodsNm	= Replace(iGoodsNm,"&gt;",">")
						iGoodsNm	= Replace(iGoodsNm,"&lt;","<")
						iGoodsNm	= Replace(iGoodsNm,"&nbsp;"," ")
						iGoodsNm	= Replace(iGoodsNm,"&amp;","&")

						GoodsRegDtime = SubNodes.getElementsByTagName("GoodsRegDtime").item(0).text
						DispYn 		 = SubNodes.getElementsByTagName("DispYn").item(0).text

						If (SaleStatCd="10") Then
							LotteSellyn = "Y"
						ElseIf (SaleStatCd="20") Then
							LotteSellyn = "N"
						ElseIf (SaleStatCd="30") Then
							LotteSellyn = "X"
						End If

						''rw  iGoodsNo&"|"&SaleStatCd&"|"&iSalePrc&"|"&iGoodsNm&"|"&GoodsRegDtime&"|"&DispYn
						sqlstr = "exec [db_temp].[dbo].[usp_TEN_OutMall_CheckRegItemLIST] 'lotteimall','"&iGoodsNo&"','"&LotteSellyn&"',"&iSalePrc&",'"&replace(iGoodsNm,"'","''")&"','"&GoodsRegDtime&"','"&DispYn&"','"&regDtKey&"'"
						dbget.Execute sqlstr

						ii = ii+1
						if ii mod 2000=0 then response.flush
					Next
					Set oneGoodsInfo = Nothing

					sqlstr = "exec [db_temp].[dbo].[usp_TEN_OutMall_CheckRegItemLIST_MAP] 'lotteimall','"&regDtKey&"'"
					dbget.Execute sqlstr
''response.write sqlstr
					iErrStr =  "OK||"&styyyymmdd&"-"&edyyyymmdd&"||����(���û�ǰ��ȸ)-"&GoodsCount&"��"
					fnLtiMallGoodsList = True
			    Else
			    	resultmsg = xmlDOM.getElementsByTagName("Message").Item(0).Text
			    	iErrStr =  "ERR||"&iitemid&"||"&resultmsg&"(���û�ǰ��ȸ)"
		            fnLtiMallGoodsList = False
			    End If

			Set xmlDOM = Nothing
		Else
			iErrStr = "�Ե����̸��� ����߿� ������ �߻��߽��ϴ�..[ERR-ItemChk-001]"
			fnLtiMallGoodsList = False
		End If
	Set objXML = Nothing
	'On Error Goto 0
End Function
'############################################## ���� �����ϴ� API �Լ� ���� �� ############################################

'################################################# �� ��� �� �Ķ���� ���� ###############################################
'ǰ�� �Ķ��Ÿ
Function getLtiMallSellynParameter(ichgSellYn, iLotteGoodNo)
'iLotteGoodNo = 1035664985
    Dim strRst
	strRst = "?subscriptionId=" & ltiMallAuthNo										'�Ե����̸� ������ȣ	(*)
	strRst = strRst & "&goods_no=" & iLotteGoodNo                       			'�Ե����̸� ��ǰ��ȣ	(*)
	If ichgSellYn = "Y" Then														'�Ǹſ���(10:�Ǹ�, 20:ǰ��, 30:�Ǹ�����)
		strRst = strRst & "&sale_stat_cd=10"
	ElseIf ichgSellYn = "N" Then
		strRst = strRst & "&sale_stat_cd=20"
	ElseIf ichgSellYn = "X" Then													'''X ��� ������
		strRst = strRst & "&sale_stat_cd=30"
	End If
	getLtiMallSellynParameter = strRst
End Function

'���� ���� �Ķ���� ����
Function getLtiMallPriceParameter(iitemid, iLotteGoodNo, MustPrice)
	Dim strRst, strSql
	Dim sellcash, orgprice, buycash, specialPrice
	Dim GetTenTenMargin

	strRst = "subscriptionId=" & ltiMallAuthNo
	strRst = strRst & "&strGoodsNo=" & iLotteGoodNo
	strRst = strRst & "&strReqSalePrc=" & MustPrice
	getLtiMallPriceParameter = strRst
End Function

''//��ǰ�� ���� �Ķ���� ����(�Ե����İ� �Ķ��Ÿ���� �ٸ�)
Function getLtiMallItemnameParameter(iitemid, byref iitemname, iLotteGoodNo)
	Dim strSql, chgname, strRst
	strSql = ""
	strSql = strSql & " SELECT TOP 1 r.itemid, i.ItemName "
	strSql = strSql & "	FROM db_item.dbo.tbl_ltiMall_regItem r "
	strSql = strSql & "	JOIN db_item.dbo.tbl_item i on r.itemid = i.itemid "
	strSql = strSql & "	WHERE i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.Eof Then
		iitemname = rsget("ItemName")
	End If
	rsget.close

	chgname = ""
	chgname = replace(db2html(iitemname),"'","")
	chgname = replace(chgname,"<B>","")
	chgname = replace(chgname,"</B>","")
	chgname = replace(chgname,"~","-")
	chgname = replace(chgname,"?","-")
	chgname = replace(chgname,"<","[")
	chgname = replace(chgname,">","]")
	chgname = replace(chgname,"%","����")
	chgname = replace(chgname,"+","%2B")
	chgname = replace(chgname,"&","%26")
	chgname = replace(chgname,"[������]","")
	chgname = replace(chgname,"[���� ���]","")
	strRst = "subscriptionId=" & ltiMallAuthNo
	strRst = strRst & "&goods_no=" & iLotteGoodNo
	strRst = strRst & "&goods_nm=" & Trim(chgname)
	strRst = strRst & "&chg_caus_cont=api ��ǰ�� ����"
	getLtiMallItemnameParameter = strRst
End Function

'�ɼ� �߰� �Ķ��Ÿ
Function getLtiMallAddOptParameter(nm, dc, iLotteGoodNo)
	Dim strRst
	strRst = "subscriptionId=" & ltiMallAuthNo											'(*)���������Ű
	strRst = strRst & "&goods_no=" & iLotteGoodNo										'(*)�Ե����̸� ��ǰ��ȣ
	strRst = strRst & "&opt_nm=" & nm													'(*)�Ե����̸� �߰��� �ɼǸ�
	strRst = strRst & "&item_nm=" & dc													'(*)�Ե����̸� �߰��� �ɼ�������
	getLtiMallAddOptParameter = strRst
End Function
'################################################ �� ��� �� �Ķ���� ���� �� #############################################

'################################################ ���ϴ� �� ����ϱ� ���� �Լ� ############################################
Public Function GetRaiseValue(value)
    If Fix(value) < value Then
    	GetRaiseValue = Fix(value) + 1
    Else
    	GetRaiseValue = Fix(value)
    End If
End Function

Function getLtiMallTmpItemIdByTenItemID(iitemid)
	Dim sqlStr, retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT ltiMallTmpGoodNo, isnull(ltiMallGoodNo,'') as ltiMallGoodNo " & VBCRLF
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_LTiMall_regItem" & VBCRLF
	sqlStr = sqlStr & " WHERE itemid = "&iitemid & VBCRLF
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		If rsget("ltiMallGoodNo") <> "" Then
			retVal = "���û�ǰ"
		Else
			retVal = rsget("ltiMallTmpGoodNo")
		End If
	End If
	rsget.Close

	If IsNULL(retVal) Then retVal = ""
	getLtiMallTmpItemIdByTenItemID = retVal
End Function
%>
