<%
Public gsshopAPIURL
Public gsshopNewAPIURL
IF application("Svr_Info") = "Dev" THEN
	'gsshopAPIURL = "http://test1.gsshop.com/alia/aliaCommonPrd.gs"	'�׽�Ʈ����
	'gsshopNewAPIURL = "http://testapi.gsshop.com/alia/aliaCommonPrd.gs"
	gsshopAPIURL = "http://ecb2b.gsshop.com/alia/aliaCommonPrd.gs"	'�Ǽ���
	gsshopNewAPIURL = "http://realapi.gsshop.com/alia/aliaCommonPrd.gs"
Else
	gsshopAPIURL = "http://ecb2b.gsshop.com/alia/aliaCommonPrd.gs"	'�Ǽ���
	gsshopNewAPIURL = "http://realapi.gsshop.com/alia/aliaCommonPrd.gs"
End If
'############################################## ���� �����ϴ� API �Լ� ���� ##############################################
'New ��ǰ ��� �Լ�
Function fnGSShopNewItemReg(iitemid, strParam, byRef iErrStr, iSellCash, iGSShopSellYn, ilimityn, ilimitno, ilimiysold, iitemname, iimagename)
	Dim objXML, xmlDOM, strRst
	Dim buf, strSql, AssignedRow
	Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	Dim attrPrdlist, lp, tenOptcd, gsOptcd, strObj, iRbody
	Dim Tlimitno, Tlimitsold, Titemoption, Toptionname, Toptlimitno, Toptlimitsold, Toptsellyn, Toptlimityn, Toptaddprice, Tlimityn, Tsellyn, Titemsu, Tsellcash
	Dim isAttrYn

	On Error Resume Next
	fnGSShopNewItemReg = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			If (session("ssBctID")="kjy8517") Then
				rw "REQ : <textarea cols=40 rows=10>"&strParam&"</textarea>"
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If

			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					If Err.number <> 0 Then
						iErrStr = "ERR||"&iitemid&"||"&Err.Description&"(ERR.��ǰ���)"
					Else
						iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(��ǰ���)"
					End If
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					'��ǰ���翩�� Ȯ��
					strSql = "Select count(itemid) From db_item.dbo.tbl_gsshop_regitem Where itemid='" & iitemid & "'"
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If rsget(0) > 0 Then
						'// ���� -> ����
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCRLF
						strSql = strSql & "	Set GSShopLastUpdate = getdate() "  & VbCRLF
						strSql = strSql & "	, GSShopGoodNo = '" & prdCd & "'"  & VbCRLF
						strSql = strSql & "	, GSShopPrice = " &iSellCash& VbCRLF
						strSql = strSql & "	, accFailCnt = 0"& VbCRLF
						strSql = strSql & "	, GSShopRegdate = isNULL(GSShopRegdate, getdate())"& VbCRLF
						strSql = strSql & "	, GSShopSellYn = '" & iGSShopSellYn & "'"& VbCRLF
						If (prdCd <> "") Then
						    strSql = strSql & "	, GSShopstatCD = '3'"& VbCRLF					'��ϿϷ�(�ӽ�)
						Else
							strSql = strSql & "	, GSShopstatCD = '1'"& VbCRLF					'���۽õ�
						End If
						strSql = strSql & "	, regImageName = '" & iimagename & "'"& VbCRLF
						strSql = strSql & "	From db_item.dbo.tbl_gsshop_regItem R"& VbCRLF
						strSql = strSql & " Where R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
					Else
						'// ���� -> �űԵ��
						strSql = ""
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_gsshop_regItem "
						strSql = strSql & " (itemid, regitemname, reguserid, GSShopRegdate, GSShopLastUpdate, GSShopGoodNo, GSShopPrice, GSShopSellYn, GSShopStatCd, regImageName) VALUES " & VbCRLF
						strSql = strSql & " ('" & iitemid & "'" & VBCRLF
						strSql = strSql & " , '" & iitemname & "'" &_
						strSql = strSql & " , '" & session("ssBctId") & "'" &_
						strSql = strSql & " , getdate(), getdate()" & VBCRLF
						strSql = strSql & " , '" & prdCd & "'" & VBCRLF
						strSql = strSql & " , '" & iSellCash & "'" & VBCRLF
						strSql = strSql & " , '" & iGSShopSellYn & "'" & VBCRLF
						If (prdCd <> "") Then
						    strSql = strSql & ",'3'"											'��ϿϷ�(�ӽ�)
						Else
						    strSql = strSql & ",'1'"											'���۽õ�
						End If
						strSql = strSql & " , '" & iimagename & "'" & VBCRLF
						strSql = strSql & ")"
						dbget.Execute(strSql)
					End If
					rsget.Close

					' On Error Resume Next
					' If isobject(strObj.attr) Then
					' 	If Err = 0 then
					' 		isAttrYn = "Y"
					' 	Else
					' 		isAttrYn = "N"
					' 	End If
					' End If
					' On Error Goto 0

					' If isAttrYn = "Y" Then				'�ɼ��̶��
					' 	Set attrPrdlist = strObj.attr
					' 		For lp=0 to attrPrdlist.length-1
					' 			tenOptcd = attrPrdlist.get(lp).supAttrPrdCd
					' 			gsOptcd = attrPrdlist.get(lp).attrPrdCd
					' 			strSql = ""
					' 			strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
					' 			strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
					' 			strSql = strSql & " SELECT itemid, itemoption, 'gsshop', '"&gsOptcd&"', optionname, optsellyn, optlimityn, " & VBCRLF
					' 			strSql = strSql & " Case WHEN optlimityn = 'Y' AND optlimitno - optlimitsold <= 5 THEN '0' " & VBCRLF
					' 			strSql = strSql & " 	 WHEN optlimityn = 'Y' AND optlimitno - optlimitsold > 5 THEN optlimitno - optlimitsold - 5 " & VBCRLF
					' 			strSql = strSql & " 	 WHEN optlimityn = 'N' THEN '999' End " & VBCRLF
					' 			strSql = strSql & " , '0', getdate() " & VBCRLF
					' 			strSql = strSql & " FROM db_item.dbo.tbl_item_option " & VBCRLF
					' 			strSql = strSql & " WHERE itemid= '"&iitemid&"' " & VBCRLF
					' 			strSql = strSql & " and itemoption = '"& tenOptcd &"' "
					' 			dbget.Execute strSql
					' 		Next

					' 	Set attrPrdlist = nothing
					' Else								'��ǰ�̶��
					' 	strSql = ""
					' 	strSql = strSql & " SELECT COUNT(*) FROM db_item.dbo.tbl_item_option WHERE itemid = '"&iitemid&"' "
					' 	rsget.CursorLocation = adUseClient
					' 	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					' 	If rsget(0) = 0 Then
					' 		tenOptcd	= "0000"
					' 	End If
					' 	rsget.Close

					' 	If (tenOptcd = "0000")  Then	'���� ��ǰ�̶��
					' 		'gsOptcd			= split(attrPrdCd,"^")(0)
					' 		gsOptcd			= ""
					' 		Toptionname		= "����"
					' 		Tlimitno		= ilimitno
					' 		Tlimitsold		= ilimiysold
					' 		Tlimityn		= ilimityn
					' 		If (Tlimityn="Y") then
					' 			If (Tlimitno - Tlimitsold) < 5 Then
					' 				Titemsu = 0
					' 			Else
					' 				Titemsu = Tlimitno - Tlimitsold - 5
					' 			End If
					' 		Else
					' 			Titemsu = 999
					' 		End If
					' 		strSql = ""
					' 		strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
					' 		strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
					' 		strSql = strSql & " VALUES " & VBCRLF
					' 		strSql = strSql & " ('"&iitemid&"',  '"&tenOptcd&"', 'gsshop', '"&gsOptcd&"', '"&html2db(Toptionname)&"', 'Y', '"&Tlimityn&"', '"&Titemsu&"', '0', getdate()) "
					' 		dbget.Execute strSql
					' 	End If
					' End If
					' strSql = ""
					' strSql = strSql & " UPDATE R " & VBCRLF
					' strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VBCRLF
					' strSql = strSql & " FROM db_item.dbo.tbl_gsshop_regItem R " & VBCRLF
					' strSql = strSql & " Join ( " & VBCRLF
					' strSql = strSql & " 	SELECT R.itemid, count(*) as CNT, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt "
					' strSql = strSql & " 	FROM db_item.dbo.tbl_gsshop_regItem R " & VBCRLF
					' strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = 'gsshop' and Ro.itemid = " &iitemid & VBCRLF
					' strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
					' strSql = strSql & " ) T on R.itemid = T.itemid " & VBCRLF
					' dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(��ǰ���)"
					fnGSShopNewItemReg = True
		        End If
			Set strObj = nothing
		Else
			fnGSShopNewItemReg = False
			iErrStr = "ERR||"&iitemid&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG-002]"
		End If
	Set objXML= nothing
End Function

'New ǰ�� ���� �Լ�
Public Function fnGSShopNewSellyn(iitemid, ichgSellYn, istrParam, byRef iErrStr)
    Dim strParam, resultcode, resultmsg, supPrdCd, supCd, prdCd
    Dim objXML, xmlDOM, strObj
    Dim strRst, strSql, iRbody
    On Error Resume Next
    fnGSShopNewSellyn = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(istrParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(���º���)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem " & VbCRLF
					strSql = strSql & " SET GSShopLastUpdate = getdate() " & VbCRLF
					strSql = strSql & " ,lastStatCheckDate = getdate() " & VbCRLF
					strSql = strSql & " ,GSShopSellYn = '" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " ,accFailCnt = 0 " & VbCRLF
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute(strSql)
					If ichgSellYn = "N" Then
						iErrStr = "OK||"&iitemid&"||ǰ��ó��"
					Else
						iErrStr = "OK||"&iitemid&"||�Ǹ������� ����"
					End If
		        End If
			Set strObj = nothing
			fnGSShopNewSellyn = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-SELLEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'��ȸ �Լ�
Public Function fnGSShopItemView(iitemid, istrParam, byRef iErrStr, iVal)
    Dim strParam, resultcode, resultmsg, supPrdCd, supCd, prdCd
    Dim objXML, xmlDOM, strObj, i, AssignedRow
    Dim strRst, strSql, iRbody

	Dim gsshopsellYn, GSShopGoodNo, gsshopPrice, prdPrcList, prdAttrInfoList, outmallSellyn, outmallOptCode, outmallOptName, tenOptcd, outmalllimitno
    On Error Resume Next
    fnGSShopItemView = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", "http://realapi.gsshop.com/api/v3/getPrdInfo.gs", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(istrParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.result
				resultmsg	= replaceMsg(strObj.message)
				If resultcode <> "success" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(��ȸ)"
				Else
					strSql = ""
					strSql =  strSql & " DELETE FROM db_item.dbo.tbl_OutMall_regedoption WHERE mallid='"&CMALLNAME&"' and itemid="&iitemid&" "
					dbget.Execute strSql

					GSShopGoodNo		= strObj.data.prdBaseInfo.prdCd
					gsshopsellYn = ""
					If strObj.data.prdBaseInfo.prdStCd = "Y" Then	'�ǸŻ��� | "�ǸŴ�� : N (MD���� �� ���� ���γ��� ��), �Ǹ��� : Y, �Ǹ����� : E, �Ͻ�ǰ�� : T, �����Ǹ����� : D (���������ڵ尡 31)"
						gsshopsellYn = "Y"
					ElseIf strObj.data.prdBaseInfo.prdStCd = "N" Then
						gsshopsellYn = "E"
					Else
						gsshopsellYn = "N"
					End If
					Set prdPrcList = strObj.data.prdPrcList		'��ǰ���
						For i=0 to prdPrcList.length-1
							gsshopPrice = prdPrcList.get(i).prdPrcSalePrc	'�ǸŰ���
						Next
					Set prdPrcList = nothing

					Set prdAttrInfoList = strObj.data.prdAttrInfoList		'��ǰ �Ӽ� ���� ����Ʈ
						For i=0 to prdAttrInfoList.length-1
							outmallSellyn = ""
							outmallOptCode	= prdAttrInfoList.get(i).attrPrdCd			'GS�Ӽ���ǰ�ڵ�
							outmallOptName	= prdAttrInfoList.get(i).prdAttrVal1		'�Ӽ���1
							tenOptcd		= prdAttrInfoList.get(i).supAttrPrdCd						'���»�Ӽ���ǰ�ڵ�
							If prdAttrInfoList.get(i).attrStCd = "Y" Then				'�Ӽ��ǸŻ��� | �ǸŴ�� : N (MD���� �� ���� ���γ��� ��), �Ǹ��� : Y, �Ǹ����� : E, �Ͻ�ǰ�� : T, �����Ǹ����� : D (���������ڵ尡 31)
								outmallSellyn = "Y"
							Else
								outmallSellyn = "N"
							End If
							outmalllimitno	= prdAttrInfoList.get(i).attrOrdPsblQty		'�ֹ����ɼ���

							strSql = " INSERT INTO db_item.dbo.tbl_OutMall_regedoption"
							strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outMallSellyn, outmalllimityn, outMallLimitNo)"
							strSql = strSql & " VALUES ("&iitemid
							If i = 0 AND outmallOptName = "����" Then
								strSql = strSql & " ,'0000'"
							Else
								strSql = strSql & " ,'"& tenOptcd &"'"
							End If
							strSql = strSql & " ,'"&CMALLNAME&"'"
							strSql = strSql & " ,'"&outmallOptCode&"'"
							strSql = strSql & " ,'"&html2DB(outmallOptName)&"'"
							strSql = strSql & " ,'"&outmallSellyn&"'"
							strSql = strSql & " ,'Y'"
							strSql = strSql & " ,"&outmalllimitno
							strSql = strSql & ")"
							dbget.Execute strSql, AssignedRow
						Next
					Set prdAttrInfoList = nothing
					strSql = ""
					strSql = strSql & " UPDATE R " & VbCRLF
					strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VbCRLF
					strSql = strSql & " ,lastconfirmdate = getdate()"& VbCRLF
					strSql = strSql & " ,GSShopSellYn = '"& gsshopsellYn &"' "& VbCRLF
					strSql = strSql & " ,GSShopPrice = '"& gsshopPrice &"' "& VbCRLF
					strSql = strSql & " ,GSShopGoodNo = CASE WHEN isNull(GSShopGoodNo, '') = '' THEN '"& GSShopGoodNo &"' ELSE GSShopGoodNo END "& VbCRLF
					If GSShopGoodNo <> "" Then
						strSql = strSql & " ,GSShopstatCD = 3 "& VbCRLF
					End If

					If iVal = "REG" Then
						strSql = strSql & " ,GSShopLastUpdate = GETDATE() "& VbCRLF
						strSql = strSql & " ,accFailCnt = 0 "& VbCRLF
						strSql = strSql & " ,GSShopRegdate = GETDATE()  "& VbCRLF
					End If
					strSql = strSql & " FROM db_item.dbo.tbl_gsshop_regItem R " & VbCRLF
					strSql = strSql & " JOIN ( " & VbCRLF
					strSql = strSql & " 	SELECT R.itemid,count(*) as CNT "
					strSql = strSql & " 	, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
					strSql = strSql & "		FROM db_item.dbo.tbl_gsshop_regItem R " & VbCRLF
					strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro " & VbCRLF
					strSql = strSql & " 		on R.itemid = Ro.itemid"   & VbCRLF
					strSql = strSql & " 		and Ro.mallid = '"&CMALLNAME&"'"   & VbCRLF
					strSql = strSql & "         and Ro.itemid = "&iitemid & VbCRLF
					strSql = strSql & " 	GROUP BY R.itemid "   & VbCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid " & VbCRLF
					dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||����(��ȸ)"
		        End If
			Set strObj = nothing
			fnGSShopItemView = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-CHKSTAT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ���û�ǰ �ǸŰ� ����
Public Function fnGSShopNewPrice(iitemid, istrParam, imustprice, byRef iErrStr)
    Dim objXML,xmlDOM,strRst
    Dim buf, strSql, strObj, iRbody
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewPrice = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(istrParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)

				If Instr(resultmsg, "���»����޾� �ݾװ� �����մϴ�") > 0 Then
					resultcode = "True"
				End If

				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(��ǰ����)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

				    strSql = ""
	    			strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem  " & VbCRLF
	    			strSql = strSql & "	SET GSShopLastUpdate=getdate() " & VbCRLF
	    			strSql = strSql & "	, GSShopPrice = " & imustprice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " Where itemid='" & iitemid & "'"& VbCRLF
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(��ǰ����)"
					fnGSShopNewPrice = True
		        End If
			Set strObj = nothing
		Else
			fnGSShopNewPrice = False
			iErrStr = "ERR||"&iitemid&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRICE-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New��ǰ �ɼ� �߰� �� ���� ����
Function fnGSShopNewOPTSuEdit(iitemid, strParam, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage, i
    Dim buf, tenOptcd, lp, gsOptcd, sqlStr, strObj, iRbody
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd, attrPrdlist, Assignedrow
	On Error Resume Next
	fnGSShopNewOPTSuEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)

				If Instr(resultmsg, "�ֹ����ɼ����� �����մϴ�") > 0 Then
					resultcode = "True"
				End If

				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(�ɼ� �߰� �� ���� ����)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					Set attrPrdlist = strObj.attr
						For i=0 to attrPrdlist.length-1
							tenOptcd = attrPrdlist.get(i).supAttrPrdCd
							gsOptcd = attrPrdlist.get(i).attrPrdCd
							If attrPrdlist.length-1 = 0 AND tenOptcd = "0000" Then	'��ǰ�̶��
								sqlStr = ""
								sqlStr = sqlStr & "UPDATE db_item.dbo.tbl_OutMall_regedoption SET "
								sqlStr = sqlStr & "outmalllimitno =  "
								sqlStr = sqlStr & "Case WHEN B.limityn = 'Y' and B.limitno - B.limitsold <= 5 THEN '0'  "
								sqlStr = sqlStr & "	 WHEN B.limityn = 'Y' and B.limitno - B.limitsold > 5 THEN B.limitno - B.limitsold - 5 "
								sqlStr = sqlStr & "	 WHEN B.limityn = 'N' THEN '999' END "
								sqlStr = sqlStr & "FROM db_item.dbo.tbl_OutMall_regedoption A  "
								sqlStr = sqlStr & "JOIN db_item.dbo.tbl_item B on A.itemid = B.itemid "
								sqlStr = sqlStr & "WHERE A.itemid = '"&iitemid&"' and A.itemoption = '"&tenOptcd&"' and A.mallid = 'gsshop' "
								dbget.Execute sqlStr
							Else
								sqlStr = ""
								sqlStr = sqlStr & " IF Exists(SELECT * FROM db_item.dbo.tbl_OutMall_regedoption where itemid='"&iitemid&"' and itemoption = '"&tenOptcd&"' and mallid = 'gsshop') "
								sqlStr = sqlStr & " BEGIN"& VbCRLF
								sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_OutMall_regedoption " & VbCRLF
								sqlStr = sqlStr & " SET outmalllimitno = " & VbCRLF
								sqlStr = sqlStr & " Case WHEN optlimityn = 'Y' AND optlimitno - optlimitsold <= 5 THEN '0' " & VbCRLF
								sqlStr = sqlStr & " 	 WHEN optlimityn = 'Y' AND optlimitno - optlimitsold > 5 THEN optlimitno - optlimitsold - 5" & VbCRLF
								sqlStr = sqlStr & " 	 WHEN optlimityn = 'N' THEN '999' End" & VbCRLF
								sqlStr = sqlStr & " ,outmalllimityn = B.optlimityn " & VbCRLF
								sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption A  " & VbCRLF
								sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_option B on A.itemid = B.itemid and A.itemoption = B.itemoption " & VbCRLF
								sqlStr = sqlStr & " WHERE B.itemid = '"&iitemid&"' and B.itemoption = '"&tenOptcd&"' and A.mallid = 'gsshop' "
								sqlStr = sqlStr & " END ELSE "
								sqlStr = sqlStr & " BEGIN"& VbCRLF
								sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
								sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
								sqlStr = sqlStr & " SELECT itemid, itemoption, 'gsshop', '"&gsOptcd&"', optionname, optsellyn, optlimityn, " & VBCRLF
								sqlStr = sqlStr & " Case WHEN optlimityn = 'Y' AND optlimitno - optlimitsold <= 5 THEN '0' " & VBCRLF
								sqlStr = sqlStr & " 	 WHEN optlimityn = 'Y' AND optlimitno - optlimitsold > 5 THEN optlimitno - optlimitsold - 5 " & VBCRLF
								sqlStr = sqlStr & " 	 WHEN optlimityn = 'N' THEN '999' End " & VBCRLF
								sqlStr = sqlStr & " , '0', getdate() " & VBCRLF
								sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option " & VBCRLF
								sqlStr = sqlStr & " WHERE itemid= '"&iitemid&"' " & VBCRLF
								sqlStr = sqlStr & " and itemoption = '"& tenOptcd &"' "
								sqlStr = sqlStr & " END "
								dbget.Execute sqlStr
							End If
						Next
					Set attrPrdlist = nothing
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(�ɼ� �߰� �� ���� ����)"
					fnGSShopNewOPTSuEdit = True
		        End If
			Set strObj = nothing
		Else
			fnGSShopNewOPTSuEdit = False
			iErrStr = "ERR||"&iitemid&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-OPTSuEdit-002]"
		End If
	Set objXML= nothing
End Function

'New��ǰ �ɼ� ���º���
Function fnGSShopNewOPTSellEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, tenOptcd, lp, gsOptcd, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd, attrPrdlist, sqlStr
    On Error Resume Next
    fnGSShopNewOPTSellEdit = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(�ɼǻ��º���)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

	                ' sqlStr = ""
					' sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_OutMall_regedoption " & VbCRLF
					' sqlStr = sqlStr & " SET outmallsellyn = " & VbCRLF
					' sqlStr = sqlStr & " Case WHEN (B.isusing = 'N' OR B.optsellyn = 'N') THEN 'N' " & VbCRLF
					' sqlStr = sqlStr & " 	 WHEN (B.optlimityn = 'Y' AND B.optlimitno - B.optlimitsold <= 5) THEN 'N'  " & VbCRLF
					' sqlStr = sqlStr & " 	 WHEN (A.outmallOptName <> B.optionname) THEN 'N'  " & VbCRLF
					' sqlStr = sqlStr & " 	 WHEN (isNull(B.itemoption, '') = '') THEN 'N'  " & VbCRLF
					' sqlStr = sqlStr & " ELSE 'Y' END " & VbCRLF
					' sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption A  " & VbCRLF
					' sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item_option B on A.itemid = B.itemid and A.itemoption = B.itemoption " & VbCRLF
					' sqlStr = sqlStr & " WHERE A.itemid = '"&iitemid&"' and A.mallid = 'gsshop' "
				    ' dbget.Execute sqlStr
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(�ɼǻ��º���)"
		        End If
			Set strObj = nothing
			fnGSShopNewOPTSellEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-OPTSellEdit-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ��ǰ�� ���� ���� �Լ�
Public Function fnGSShopChgNewItemname(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, strSql, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
    On Error Resume Next
    fnGSShopChgNewItemname = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)

				If Instr(resultmsg, "��ϵ� ������ ������ ��û�Դϴ�") > 0 Then
					resultcode = "True"
				End If

				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(��ǰ�� ����)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem " & VbCRLF
					strSql = strSql & " SET regitemname = B.itemname "& VbCRLF
					strSql = strSql & " FROM db_item.dbo.tbl_gsshop_regItem A "& VbCRLF
					strSql = strSql & " JOIN db_item.dbo.tbl_item B on A.itemid = B.itemid "& VbCRLF
					strSql = strSql & " WHERE A.itemid='" & iitemid & "'"& VbCRLF
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(��ǰ�� ����)"
		        End If
			Set strObj = nothing
			fnGSShopChgNewItemname = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-NMEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New �̹��� ���� ���� �Լ�
Function fnGSShopNewImageEdit(iitemid, strParam, iErrStr, ichgImageNm)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj, strSql
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewImageEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(�̹�������)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regitem "
					strSql = strSql & " SET regimageName='"&ichgImageNm&"'"
					strSql = strSql & " WHERE itemid = '"& iitemid &"' "
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(�̹�������)"
		        End If
			Set strObj = nothing
			fnGSShopNewImageEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-IMGEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ������������ ���� �Լ�
Function fnGSShopNewSafeCertEdit(iitemid, strParam, iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewSafeCertEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(���ȹ�����)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(���ȹ�����)"
		        End If
			Set strObj = nothing
			fnGSShopNewSafeCertEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-SAFEEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ��ǰ���� ���� ���� �Լ�
Function fnGSShopNewItemInfoEdit(iitemid, strParam, iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewItemInfoEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				prdCd		= strObj.prdCd
				supPrdCd	= strObj.supPrdCd
				resultmsg	= replaceMsg(strObj.msg)
				resultcode	= strObj.success
				supCd 		= strObj.supCd

				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(��ǰ����)"
				Else
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET GSShopLastUpdate = getdate()"
					strSql = strSql & " FROM db_item.dbo.tbl_gsshop_regitem R" & VbCrlf
					strSql = strSql & " where R.itemid = " & iitemid
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(��ǰ����)"
		        End If
			Set strObj = nothing
			fnGSShopNewItemInfoEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-INFOEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ���û�ǰ ���� ����
Function fnGSShopNewContentsEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewContentsEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=euc-kr"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(��ǰ�������)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(��ǰ�������)"

					strSql = ""
					strSql = strSql & " UPDATE db_item.[dbo].[tbl_gsshop_regitem] "
					strSql = strSql & " SET isRegHtmlErr = NULL "
					strSql = strSql & " WHERE itemid = '"& itemid &"' "
					dbget.Execute strSql
		        End If
			Set strObj = nothing
			fnGSShopNewContentsEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-CONTEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ���ΰ���׸� ����
Function fnGSShopNewInfodivEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewInfodivEdit = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(���ΰ���׸����)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(���ΰ���׸� ����)"
		        End If
			Set strObj = nothing
			fnGSShopNewInfodivEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-DIVEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'�������� ����
Function fnGSShopCateEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopCateEdit = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(������������)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iitemid&"||"&resultmsg&"(������������)"
		        End If
			Set strObj = nothing
			fnGSShopCateEdit = true
		Else
			iErrStr = "ERR||"&iitemid&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-CATEEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function getGSShopDivCodeView()
    Dim objXML, xmlDOM, strRst, resultList, strObj
    Dim strSql, iRbody, i, j
	Dim lrgCd, lrgNm, midCd, midNm, smCd, smNm, dtlCd, dtlNm, isusing, infoDivArr, infoDivNameArr, infoDiv, infoDivName, safeGbnCd
'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", "http://realapi.gsshop.com/SupSendPrdClsInfo.gs", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			Set strObj = JSON.parse(iRbody)
				Set resultList = strObj.resultList
					strSql = ""
					strSql = " DELETE FROM db_temp.[dbo].[tbl_gsshopMng_metaInfo] "
					dbget.Execute(strSql)
					For i=0 to resultList.length-1
						lrgCd = resultList.get(i).lrgClsCd
						lrgNm = resultList.get(i).lrgClsNm
						midCd = resultList.get(i).midClsCd
						midNm = resultList.get(i).midClsNm
						smCd = resultList.get(i).smlClsCd
						smNm = resultList.get(i).smlClsNm
						dtlCd = resultList.get(i).dtlClsCd
						dtlNm = resultList.get(i).dtlClsNm
						isusing = resultList.get(i).useYn
						safeGbnCd = resultList.get(i).safeCertTgtGbnCd
						infoDivArr = Split(resultList.get(i).govPublsPrdGrpCd, "$")
						infoDivNameArr = Split(resultList.get(i).govPublsPrdGrpNm, "$")

						strSql = ""
						strSql = strSql & " IF EXISTS (SELECT TOP 1 dtlCd FROM db_temp.[dbo].[tbl_gsshopMng_category] WHERE lrgCd = '"& lrgCd &"' and midCd = '"& midCd &"' and smCd = '"& smCd &"' and dtlCd = '"& dtlCd &"' ) "
						strSql = strSql & " 	BEGIN "
						strSql = strSql & " 		UPDATE db_temp.[dbo].[tbl_gsshopMng_category] "
						strSql = strSql & " 		SET isusing = '"& isusing &"'"
						strSql = strSql & " 		,safeGbnCd = '"& safeGbnCd &"'"
						strSql = strSql & " 		WHERE lrgCd = '"& lrgCd &"' and midCd = '"& midCd &"' and smCd = '"& smCd &"' and dtlCd = '"& dtlCd &"' "
						strSql = strSql & " 	END "
						strSql = strSql & " ELSE "
						strSql = strSql & " 	BEGIN "
						strSql = strSql & " 		INSERT INTO db_temp.[dbo].[tbl_gsshopMng_category] "
						strSql = strSql & " 		(lrgCd, lrgNm, midCd, midNm, smCd, smNm, dtlCd, dtlNm, isusing) VALUES "
						strSql = strSql & " 		('"&lrgCd&"', '"&lrgNm&"', '"&midCd&"', '"&midNm&"', '"&smCd&"', '"&smNm&"', '"&dtlCd&"', '"&dtlNm&"', '"&isusing&"') "
						strSql = strSql & " 	END "
						dbget.Execute(strSql)
						 For j = 0 to Ubound(infoDivArr)
						 	strSql = ""
						 	strSql = strSql & " INSERT INTO db_temp.[dbo].[tbl_gsshopMng_metaInfo]  "
						 	strSql = strSql & " (dtlCd, infoDiv, infoDivName) VALUES "
						 	strSql = strSql & " ('"&dtlCd&"', '"&infoDivArr(j)&"', '"&infoDivNameArr(j)&"') "
						 	dbget.Execute(strSql)
						 Next

						' rw "��з��ڵ� : " & resultList.get(i).lrgClsCd
						' rw "��з��� : " & resultList.get(i).lrgClsNm
						' rw "�ߺз��ڵ� : " & resultList.get(i).midClsCd
						' rw "�ߺз��� : " & resultList.get(i).midClsNm
						' rw "�Һз��ڵ� : " & resultList.get(i).smlClsCd
						' rw "�Һз��� : " & resultList.get(i).smlClsNm
						' rw "���з��ڵ� : " & resultList.get(i).dtlClsCd
						' rw "���з��� : " & resultList.get(i).dtlClsNm
						' rw "��뿩�� : " & resultList.get(i).useYn
						' rw "����������󱸺��ڵ� : " & resultList.get(i).safeCertTgtGbnCd
						' rw "��ȿ���������󿩺� : " & resultList.get(i).validTermMngYn
						' rw "��������ǥ�ô�󿩺� : " & resultList.get(i).unitPrcYn
						' rw "���������ڵ� : " & resultList.get(i).taxTypSelCd
						' rw "������ñ׷��ڵ� : " & resultList.get(i).govPublsPrdGrpCd
						' rw "������ñ׷�� : " & resultList.get(i).govPublsPrdGrpNm
						' rw "--------------"
					Next
					rw "�Ϸ�"
				Set resultList = nothing
			Set strObj = nothing
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function getGSShopCateCodeView(iDate)
    Dim objXML, xmlDOM, strRst, resultList, strObj
    Dim strSql, iRbody, i, j, strParam
	Dim fromDtm, toDtm, sectId, sectNm, sectLrgId, sectLrgNm, sectMidId, sectMidNm, sectDtlId, sectDtlNm, prdDispYn, shopAttrCd, shopAttrNm
	fromDtm = replace(iDate, "-", "") & "000000"
	toDtm	= replace(dateadd("d", 6, iDate), "-", "") & "235959"
	strParam = "?fromDtm=" & fromDtm & "&toDtm=" & toDtm
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", "http://realapi.gsshop.com/DispSectInfo.gs" & strParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		Dim v : v= 0
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
rw iRbody
			Set strObj = JSON.parse(iRbody)
				Set resultList = strObj.resultList
					rw iDate & " ~ " & dateadd("d", 6, iDate)
					For i=0 to resultList.length-1
						sectId		= resultList.get(i).sectId				'����������̵�
						sectNm		= resultList.get(i).sectNm				'���������
						sectLrgId	= resultList.get(i).sectLrgId			'��з�������̵�
						sectLrgNm	= resultList.get(i).sectLrgNm			'��з������
						sectMidId	= resultList.get(i).sectMidId			'�ߺз�������̵�
						sectMidNm	= resultList.get(i).sectMidNm			'�ߺз������
						sectDtlId	= resultList.get(i).sectDtlId			'�Һз�������̵�
						sectDtlNm	= resultList.get(i).sectDtlNm			'�Һз������
						prdDispYn	= resultList.get(i).prdDispYn			'��ǰ�������ɿ���
						shopAttrCd	= resultList.get(i).shopAttrCd			'����Ӽ��ڵ�
						shopAttrNm	= resultList.get(i).shopAttrNm			'����Ӽ���

						If prdDispYn = "Y" Then
							' If shopAttrNm = "�Ϲݸ���" and (Trim(sectId) <> Trim(sectDtlId)) Then
							' 	rw "����������̵� : " & sectId
							' 	rw "��������� : " & sectNm
							' 	rw "��з�������̵� : " & sectLrgId
							' 	rw "��з������ : " & sectLrgNm
							' 	rw "�ߺз�������̵� : " & sectMidId
							' 	rw "�ߺз������ : " & sectMidNm
							' 	rw "�Һз�������̵� : " & sectDtlId
							' 	rw "�Һз������ : " & sectDtlNm
							' 	rw "��ǰ�������ɿ��� : " & prdDispYn
							' 	rw "����Ӽ��ڵ� : " & shopAttrCd
							' 	rw "����Ӽ��� : " & shopAttrNm
							' 	rw "-------------"
							' End If

							If sectLrgNm = "10x10" Then
								v = v + 1
								strSql = ""
								strSql = strSql & " IF NOT EXISTS (SELECT * FROM db_temp.dbo.tbl_gsshop_category WHERE catekey = '"& sectId &"'  ) "
								strSql = strSql & " 	BEGIN "
								strSql = strSql & " 		INSERT INTO db_temp.dbo.tbl_gsshop_category (CateKey, categbn, L_NAME, L_CODE, M_NAME, M_CODE, S_NAME, S_CODE, D_NAME, D_CODE, lastupdate, isusing) VALUES "
								strSql = strSql & " 		('"& sectId &"', 'B', '"&sectLrgNm&"', '"&sectLrgId&"', '"&sectMidNm&"', '"&sectMidId&"', '"&sectDtlNm&"', '"&sectDtlId&"', NULL, NULL, GETDATE(), 'Y') "
								strSql = strSql & " 	END "
								dbget.Execute strSql

								rw "����������̵� : " & sectId
								rw "��������� : " & sectNm
								rw "��з�������̵� : " & sectLrgId
								rw "��з������ : " & sectLrgNm
								rw "�ߺз�������̵� : " & sectMidId
								rw "�ߺз������ : " & sectMidNm
								rw "�Һз�������̵� : " & sectDtlId
								rw "�Һз������ : " & sectDtlNm
								rw "��ǰ�������ɿ��� : " & prdDispYn
								rw "����Ӽ��ڵ� : " & shopAttrCd
								rw "����Ӽ��� : " & shopAttrNm
								rw "-------------" & v
							End IF
						End IF
					Next
					response.write "<input type='button' value='Go' onclick=location.replace('/outmall/gsshop/gsshopActproc.asp?act=CateCodeView&sDate="&dateadd("d", 7, iDate)&"');>"
				Set resultList = nothing
			Set strObj = nothing
		End If
	Set objXML = Nothing
End Function
'############################################## ���� �����ϴ� API �Լ� ���� �� ############################################

'################################################# �� ��� �� �Ķ���� ���� ###############################################
'ǰ�� �Ķ��Ÿ
Function getGSShopSellynParameter(iitemid, ichgSellYn)
	Dim strRst, strSql
	strRst = ""
	strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
	strRst = strRst & "&modGbn=S"														'(*)�������� S : �ǸŻ��� ����
	strRst = strRst & "&regId="&COurRedId												'(*)�����
	'��ǰ�⺻(prdBaseInfo)
	strRst = strRst & "&supPrdCd="&iitemid												'(*)���»��ǰ�ڵ�
	strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
	'��ǰ����(prdPrc)

	If ichgSellYn = "Y" Then
		strRst = strRst & "&saleEndDtm=29991231235959"									'(*)�Ǹ������Ͻ� | ��ǰ�� �ߴ�(�Ǹ�����)�Ϸ��� �ߴܽ����� �Ǹ������Ͻø� �Է��մϴ�.
	ElseIf (ichgSellYn = "N") Then
		strRst = strRst & "&saleEndDtm="&FormatDate(now(), "00000000000000")			'(*)�Ǹ������Ͻ� | ��ǰ�� �ߴ�(�Ǹ�����)�Ϸ��� �ߴܽ����� �Ǹ������Ͻø� �Է��մϴ�.
	End If
	'strRst = strRst & "&attrSaleEndStModYn=N"											'(*)�Ӽ��Ǹ�������¼������� | �Ӽ�����(S) ��ǰ�ǸŻ��¸� ������ �� ����ϴ� �׸�����, ��ǰ������ ���� �� ���� �� �Ӽ���ǰ�� ���µ� �Բ� ���� �� �����Ϸ��� Y, ��ǰ�����Ϳ� �Ӽ� ������ ���º��� ���� �ÿ� N
	strRst = strRst & "&attrSaleEndStModYn=Y"											'(*)�Ӽ��Ǹ�������¼������� | �Ӽ�����(S) ��ǰ�ǸŻ��¸� ������ �� ����ϴ� �׸�����, ��ǰ������ ���� �� ���� �� �Ӽ���ǰ�� ���µ� �Բ� ���� �� �����Ϸ��� Y, ��ǰ�����Ϳ� �Ӽ� ������ ���º��� ���� �ÿ� N

	getGSShopSellynParameter = strRst
End Function

'��ȸ �Ķ��Ÿ
Function getGSShopItemViewParameter(iitemid)
	Dim strRst, strSql
	strRst = ""
	strRst = strRst & "supPrdCd="&iitemid												'(*)���»��ǰ(�Ӽ�)�ڵ�
	strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
	strRst = strRst & "&searchItmCd=PRC,ATTR"											'��ȸ�׸��ڵ� ���� | ���� �׸����� �Է°� ���� ���� ��ǰ�⺻������ ��ȸ '���ϴ� �׸��� �߰������� ��ȸ�� , (�޸�)�� ����Ͽ� ���� (��ü�׸���ȸ:ALL, ��ǰ��:NM, ����:PRC, �Ӽ�:ATTR, ���:DLV, Ȯ������:ADD, ��������:CMP, ����:SECT, ��������:SAFE, �������:GOV, �������:SPEC, �����:HTML) ex) ��ü�׸���ȸ ALL �⺻����, ����, �Ӽ�, ���� ��ȸ�� : PRC,ATTR,SECT �Ӽ�, ��ǰ��, �������, ����� : ATTR,NM,GOV,HTML
	getGSShopItemViewParameter = strRst
End Function

Public Function getGSShopPriceParameter(iitemid, mustprice)
	Dim strRst, strSql
	Dim sellcash, orgprice, buycash
	Dim GetTenTenMargin

	'���� ���� �� �ݺ�����Ʈ �Ǽ�
	strRst = ""
	strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
	strRst = strRst & "&modGbn=P"														'(*)�������� P : ���� ����
	strRst = strRst & "&regId="&COurRedId												'(*)�����
	strRst = strRst & "&regSubjCd=SUP"													'(*)�����ü�ڵ� | ���� ������ ��� : MD, ���»簡 ������ ��� : SUP
	'��ǰ�⺻(prdBaseInfo)
	strRst = strRst & "&supPrdCd="&iitemid												'(*)���»��ǰ�ڵ�
	strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
	strRst = strRst & "&subSupCd="&COurCompanyCode										'(*)�������»��ڵ� | �������»簡 ������ supCd�� ������ �Է�
	'��ǰ����(prdPrc)
	strRst = strRst & "&prdPrcValidStrDtm="&FormatDate(now(), "00000000000000")			'(*)��ȿ�����Ͻ�
	strRst = strRst & "&prdPrcValidEndDtm=29991231235959"								'(*)��ȿ�����Ͻ�
	strRst = strRst & "&prdPrcSalePrc="&mustprice										'(*)�ǸŰ���
	'strRst = strRst & "&prdPrcPrchPrc="												'(SYS)���԰��� | (SYS�� �����ʿ��� �ڵ����� �������ִ� �ڵ� �� ���� ���մϴ�. Null�� �����ֽø� �˴ϴ�.)
	strRst = strRst & "&prdPrcSupGivRtamtCd=01"											'(*)���»�������/���ڵ� | 01 : ��
	strRst = strRst & "&prdPrcSupGivRtamt="&getGSShopSuplyPrice_update(MustPrice)		'(*)���»�������/�� | �⺻�� : �ǸŰ�*(1-0.12)
	getGSShopPriceParameter = strRst
End Function

Public Function getGSShopItemnameParameter(iitemid, byref iitemname)
	Dim strRst, chgname, strSql, brandName
	strSql = ""
	strSql = strSql & " SELECT TOP 1 r.itemid, r.GSShopGoodNo, i.ItemName, c.socname_kor "
	strSql = strSql & "	FROM db_item.dbo.tbl_gsshop_regItem r "
	strSql = strSql & "	JOIN db_item.dbo.tbl_item i on r.itemid = i.itemid "
	strSql = strSql & "	JOIN db_user.dbo.tbl_user_c as c on i.makerid = c.userid "
	strSql = strSql & "	WHERE r.regitemname is Not NULL "
	strSql = strSql & "	and (r.GSShopStatCd=3 OR r.GSShopStatCd=7) "
	strSql = strSql & "	and r.GSShopGoodNo is Not Null "
	strSql = strSql & " and	i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.Eof Then
		iitemname = rsget("ItemName")
		brandName = Trim(rsget("socname_kor"))
	End If
	rsget.close

	chgname = ""
'	chgname = "[�ٹ�����]"&replace(iitemname,"'","")			'���� ��ǰ�� �տ� [�ٹ�����] �̶�� ����
	chgname = replace(iitemname,"'","")							'���� ��ǰ�� �տ� [�ٹ�����] ����

	If Left(iitemname, Len(Trim(brandName)) + 2) = "[" & brandName & "]" Then
	ElseIf (Left(iitemname, len(brandName)) <> brandName) Then
		chgname = brandName & " " & Replace(iitemname,"'","")		'[�ٹ�����] ���� ���� / �귣���ѱ۸� ���� / 2020-07-30 ���� ����
	End If

	chgname = replace(chgname,"&#8211;","-")
	chgname = replace(chgname,"~","-")
	chgname = replace(chgname,"&","��")
	chgname = replace(chgname,"<","[")
	chgname = replace(chgname,">","]")
	chgname = replace(chgname,"%","����")
	chgname = replace(chgname,"+","%2B")
	chgname = replace(chgname,":","%3A")
	chgname = replace(chgname,"[������]","")
	chgname = replace(chgname,"[���� ���]","")

	strRst = ""
	strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
	strRst = strRst & "&modGbn=N"														'(*)�������� N : �����ǰ�� ����
	strRst = strRst & "&regId="&COurRedId												'(*)�����
	'��ǰ�⺻(prdBaseInfo)
	strRst = strRst & "&supPrdCd="&iitemid												'(*)���»��ǰ�ڵ�
	strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
	'�����ǰ��(prdNmChg)
	strRst = strRst & "&prdNmChgValidStrDtm="&FormatDate(now(), "00000000000000")		'(*)��ȿ�����Ͻ�
	strRst = strRst & "&prdNmChgValidEndDtm=29991231235959"								'(*)��ȿ�����Ͻ�
	strRst = strRst & "&prdNmChgExposPrdNm=" & Trim(chrbyte(chgname,56,"Y"))							'(*)�����ǰ�� | GSShop�����ǰ��
	getGSShopItemnameParameter = strRst
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

Public Function getGSShopSuplyPrice_update(iMustPrice)
	getGSShopSuplyPrice_update = CLNG(iMustPrice * (100-CGSSHOPMARGIN) / 100)
End Function
%>
