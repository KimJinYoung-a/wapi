<%
Public gsshopAPIURL
Public gsshopNewAPIURL
IF application("Svr_Info") = "Dev" THEN
	gsshopAPIURL = "http://test1.gsshop.com/alia/aliaCommonPrd.gs"	'�׽�Ʈ����
	gsshopNewAPIURL = "http://realapi.gsshop.com/alia/aliaCommonPrd.gs"
Else
	gsshopAPIURL = "http://ecb2b.gsshop.com/alia/aliaCommonPrd.gs"	'�Ǽ���
	gsshopNewAPIURL = "http://realapi.gsshop.com/alia/aliaCommonPrd.gs"
End If
'############################################## ���� �����ϴ� API �Լ� ���� ##############################################
'New ��ǰ ��� �Լ�
Function fnGSShopNewItemReg(iitemid, strParam, byRef iErrStr, iRealSellprice, iGSShopSellYn, ilimityn, ilimitno, ilimiysold, iitemname, iitemoption, imidx, ioptionname)
	Dim objXML, xmlDOM, strRst
	Dim buf, strSql, AssignedRow
	Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	Dim attrPrdlist, lp, tenOptcd, gsOptcd, strObj, iRbody
	Dim Tlimitno, Tlimitsold, Titemoption, Toptionname, Toptlimitno, Toptlimitsold, Toptsellyn, Toptlimityn, Toptaddprice, Tlimityn, Tsellyn, Titemsu, Tsellcash

'	On Error Resume Next
	fnGSShopNewItemReg = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopNewAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
			If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
				response.write iRbody
			End If
			Set strObj = JSON.parse(iRbody)
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					If Err.number <> 0 Then
						iErrStr = "ERR||"&imidx&"||"&Err.Description&"(ERR.��ǰ���)"
					Else
						iErrStr = "ERR||"&imidx&"||"&resultmsg&"(��ǰ���)"
					End If
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					'��ǰ���翩�� Ȯ��
					strSql = "Select count(*) From db_etcmall.dbo.tbl_gsshopAddoption_regitem Where midx='" & imidx & "'"
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If rsget(0) > 0 Then
						'// ���� -> ����
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCRLF
						strSql = strSql & "	Set GSShopLastUpdate = getdate() "  & VbCRLF
						strSql = strSql & "	, GSShopGoodNo = '" & prdCd & "'"  & VbCRLF
						strSql = strSql & "	, GSShopPrice = " &iRealSellprice& VbCRLF
						strSql = strSql & "	, accFailCnt = 0"& VbCRLF
						strSql = strSql & "	, GSShopRegdate = isNULL(GSShopRegdate, getdate())"& VbCRLF
						strSql = strSql & "	, GSShopSellYn = '" & iGSShopSellYn & "'"& VbCRLF
						If (prdCd <> "") Then
						    strSql = strSql & "	, GSShopstatCD = '3'"& VbCRLF					'��ϿϷ�(�ӽ�)
						Else
							strSql = strSql & "	, GSShopstatCD = '1'"& VbCRLF					'���۽õ�
						End If
						strSql = strSql & "	From db_etcmall.dbo.tbl_gsshopAddoption_regitem R"& VbCRLF
						strSql = strSql & " Where R.midx = '" & imidx & "'"
						dbget.Execute(strSql)
					Else
						'// ���� -> �űԵ��
						strSql = ""
						strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_gsshopAddoption_regitem "
						strSql = strSql & " (regedOptCnt, reguserid, GSShopRegdate, GSShopLastUpdate, GSShopGoodNo, GSShopPrice, GSShopSellYn, GSShopStatCd, accFailCnt) VALUES " & VbCRLF
						strSql = strSql & " (0" &_
						strSql = strSql & " , '" & session("ssBctId") & "'" &_
						strSql = strSql & " , getdate(), getdate()" & VBCRLF
						strSql = strSql & " , '" & prdCd & "'" & VBCRLF
						strSql = strSql & " , '" & iRealSellprice & "'" & VBCRLF
						strSql = strSql & " , '" & iGSShopSellYn & "'" & VBCRLF
						If (prdCd <> "") Then
						    strSql = strSql & ",'3'"											'��ϿϷ�(�ӽ�)
						Else
						    strSql = strSql & ",'1'"											'���۽õ�
						End If
						strSql = strSql & " , 0" & VBCRLF
						strSql = strSql & ")"
						dbget.Execute(strSql)
					End If
					rsget.Close

					strSql = ""
					strSql = strSql & " UPDATE R "
					strSql = strSql & " SET itemname = i.itemname "
					strSql = strSql & " ,optionname = o.optionname "
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] R "
					strSql = strSql & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid "
					strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on R.itemid = o.itemid and R.itemoption = o.itemoption "
					strSql = strSql & " WHERE R.idx = '"&imidx&"' "
					strSql = strSql & " and R.mallid= 'gsshop' "
					iErrStr =  "OK||"&imidx&"||"&resultmsg&"(��ǰ���)"
					fnGSShopNewItemReg = True
		        End If
			Set strObj = nothing
		Else
			fnGSShopNewItemReg = False
			iErrStr = "ERR||"&imidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG-002]"
		End If
	Set objXML= nothing
End Function

'��ǰ ��� �Լ�
Function fnGSShopItemReg(iitemid, strParam, byRef iErrStr, iRealSellprice, iGSShopSellYn, ilimityn, ilimitno, ilimiysold, iitemname, iitemoption, imidx, ioptionname)
	Dim objXML, xmlDOM, strRst
	Dim buf, strSql, AssignedRow
	Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	Dim attrPrdlist, lp, tenOptcd, gsOptcd
	Dim Tlimitno, Tlimitsold, Titemoption, Toptionname, Toptlimitno, Toptlimitsold, Toptsellyn, Toptlimityn, Toptaddprice, Tlimityn, Tsellyn, Titemsu, Tsellcash

	On Error Resume Next
	fnGSShopItemReg = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
'rw gsshopAPIURL&"?"&strparam
'response.end
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
		    buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf
				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�
				attrPrdCd	= Split(buf, "|")(5)	'�̼��Ӽ���ǰ�ڵ�^���»�Ӽ���ǰ�ڵ�,�̼��Ӽ���ǰ�ڵ�^���»�Ӽ���ǰ�ڵ�	'�Ӽ��Ķ��Ÿ �������� ������ ���� �� ����

				If resultcode = "S" Then	'����(S)
					'��ǰ���翩�� Ȯ��
					strSql = "Select count(itemid) From db_etcmall.dbo.tbl_gsshopAddoption_regitem Where midx='" & imidx & "'"
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If rsget(0) > 0 Then
						'// ���� -> ����
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCRLF
						strSql = strSql & "	Set GSShopLastUpdate = getdate() "  & VbCRLF
						strSql = strSql & "	, GSShopGoodNo = '" & prdCd & "'"  & VbCRLF
						strSql = strSql & "	, GSShopPrice = " &iRealSellprice& VbCRLF
						strSql = strSql & "	, accFailCnt = 0"& VbCRLF
						strSql = strSql & "	, GSShopRegdate = isNULL(GSShopRegdate, getdate())"& VbCRLF
						strSql = strSql & "	, GSShopSellYn = '" & iGSShopSellYn & "'"& VbCRLF
						If (prdCd <> "") Then
						    strSql = strSql & "	, GSShopstatCD = '3'"& VbCRLF					'��ϿϷ�(�ӽ�)
						Else
							strSql = strSql & "	, GSShopstatCD = '1'"& VbCRLF					'���۽õ�
						End If
						strSql = strSql & "	From db_etcmall.dbo.tbl_gsshopAddoption_regitem R"& VbCRLF
						strSql = strSql & " Where R.midx = '" & imidx & "'"
						dbget.Execute(strSql)
					Else
						'// ���� -> �űԵ��
						strSql = ""
						strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_gsshopAddoption_regitem "
						strSql = strSql & " (regedOptCnt, reguserid, GSShopRegdate, GSShopLastUpdate, GSShopGoodNo, GSShopPrice, GSShopSellYn, GSShopStatCd, accFailCnt) VALUES " & VbCRLF
						strSql = strSql & " (0" &_
						strSql = strSql & " , '" & session("ssBctId") & "'" &_
						strSql = strSql & " , getdate(), getdate()" & VBCRLF
						strSql = strSql & " , '" & prdCd & "'" & VBCRLF
						strSql = strSql & " , '" & iRealSellprice & "'" & VBCRLF
						strSql = strSql & " , '" & iGSShopSellYn & "'" & VBCRLF
						If (prdCd <> "") Then
						    strSql = strSql & ",'3'"											'��ϿϷ�(�ӽ�)
						Else
						    strSql = strSql & ",'1'"											'���۽õ�
						End If
						strSql = strSql & " , 0" & VBCRLF
						strSql = strSql & ")"
						dbget.Execute(strSql)
					End If
					rsget.Close

					attrPrdlist = split(attrPrdCd,",")
					gsOptcd			= split(attrPrdCd,"^")(0)
					Toptionname		= ioptionname
					Tlimitno		= ilimitno
					Tlimitsold		= ilimiysold
					Tlimityn		= ilimityn
					If (Tlimityn="Y") then
						If (Tlimitno - Tlimitsold) < 5 Then
							Titemsu = 0
						Else
							Titemsu = Tlimitno - Tlimitsold - 5
						End If
					Else
						Titemsu = 999
					End If
					strSql = ""
					strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
					strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
					strSql = strSql & " SELECT TOP 1 itemid, itemoption, 'gsshop', '"&gsOptcd&"', '"&Toptionname&"', 'Y', '"&Tlimityn&"', '"&Titemsu&"', optaddprice, getdate() " & VBCRLF
					strSql = strSql & " FROM db_item.dbo.tbl_item_option " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&iitemid&"' " & VBCRLF
					strSql = strSql & " and itemoption = '"&iitemoption&"' " & VBCRLF
					dbget.Execute strSql

					strSql = ""
					strSql = strSql & " UPDATE R "
					strSql = strSql & " SET itemname = i.itemname "
					strSql = strSql & " ,optionname = o.optionname "
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] R "
					strSql = strSql & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid "
					strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on R.itemid = o.itemid and R.itemoption = o.itemoption "
					strSql = strSql & " WHERE R.idx = '"&midx&"' "
					strSql = strSql & " and R.mallid= 'gsshop' "
					iErrStr =  "OK||"&imidx&"||"&resultmsg&"(��ǰ���)"
				Else						'����(E)
	                iErrStr =  "ERR||"&imidx&"||"&resultmsg&"(��ǰ���)"
				End If
			Set xmlDOM = Nothing
			fnGSShopItemReg= true
		Else
			iErrStr = "ERR||"&imidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ǰ�� ���� �Լ�
Public Function fnGSShopNewSellyn(iidx, ichgSellYn, istrParam, byRef iErrStr)
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
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(���º���)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_gsshopAddoption_regitem] " & VbCRLF
					strSql = strSql & " SET GSShopLastUpdate = getdate() " & VbCRLF
					strSql = strSql & " ,lastStatCheckDate = getdate() " & VbCRLF
					strSql = strSql & " ,GSShopSellYn = '" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " ,accFailCnt = 0 " & VbCRLF
					strSql = strSql & " WHERE midx = '" & iidx & "'"
					dbget.Execute(strSql)
					If ichgSellYn = "N" Then
						iErrStr = "OK||"&iidx&"||ǰ��ó��"
					Else
						iErrStr = "OK||"&iidx&"||�Ǹ������� ����"
					End If
		        End If
			Set strObj = nothing
			fnGSShopNewSellyn = true
		Else
			iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-SELLEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'ǰ�� ���� �Լ�
Public Function fnGSShopSellyn(iidx, ichgSellYn, istrParam, byRef iErrStr)
    Dim strParam, resultcode, resultmsg, supPrdCd, supCd, prdCd
    Dim objXML, xmlDOM
    Dim strRst, strSql, buf
    On Error Resume Next
    fnGSShopSellyn = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(istrParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
'					rw buf
				End If

				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�

				If Err <> 0 Then
					If (IsAutoScript) Then
						iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-SELLEDIT-001]"
					Else
						iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-SELLEDIT-001]"
					End If
					Set objXML = Nothing
				    Set xmlDOM = Nothing
				    On Error Goto 0
				    Exit Function
			    End If

				If resultcode <> "S" Then
					iErrStr = "ERR||"&iidx&"||"&resultmsg
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_gsshopAddoption_regitem] " & VbCRLF
					strSql = strSql & " SET GSShopLastUpdate = getdate() " & VbCRLF
					strSql = strSql & " ,lastStatCheckDate = getdate() " & VbCRLF
					strSql = strSql & " ,GSShopSellYn = '" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " ,accFailCnt = 0 " & VbCRLF
					strSql = strSql & " WHERE midx = '" & iidx & "'"
					dbget.Execute(strSql)
					If ichgSellYn = "N" Then
						iErrStr = "OK||"&iidx&"||ǰ��ó��"
					Else
						iErrStr = "OK||"&iidx&"||�Ǹ������� ����"
					End If
		        End If
			Set xmlDOM = Nothing
			fnGSShopSellyn = True
		Else
			iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-SELLEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ���û�ǰ �ǸŰ� ����
Public Function fnGSShopNewPrice(iidx, istrParam, imustprice, byRef iErrStr)
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
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(��ǰ����)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
				    '// ��ǰ�������� ����
				    strSql = ""
	    			strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_gsshopAddoption_regitem] " & VbCRLF
	    			strSql = strSql & "	SET GSShopLastUpdate=getdate() " & VbCRLF
	    			strSql = strSql & "	, GSShopPrice = " & imustprice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " Where midx='" & iidx & "'"& VbCRLF
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(��ǰ����)"
					fnGSShopNewPrice = True
		        End If
			Set strObj = nothing
		Else
			fnGSShopNewPrice = False
			iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRICE-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'���û�ǰ �ǸŰ� ����
Public Function fnGSShopPrice(iidx, istrParam, imustprice, byRef iErrStr)
    Dim objXML,xmlDOM,strRst
    Dim buf, strSql
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopPrice = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(istrParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�

				If Err <> 0 Then
					If (IsAutoScript) Then
						iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRICE-001]"
					Else
						iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRICE-001]"
					End If
					Set objXML = Nothing
				    Set xmlDOM = Nothing
				    On Error Goto 0
				    Exit Function
			    End If

				If resultcode = "S" Then
				    '// ��ǰ�������� ����
				    strSql = ""
	    			strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_gsshopAddoption_regitem] " & VbCRLF
	    			strSql = strSql & "	SET GSShopLastUpdate=getdate() " & VbCRLF
	    			strSql = strSql & "	, GSShopPrice = " & imustprice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " Where midx='" & iidx & "'"& VbCRLF
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&idx&"||"&resultmsg&"(��ǰ����)"
					fnGSShopPrice = True
				Else
	                iErrStr =  "ERR||"&idx&"||"&resultmsg&"(��ǰ����)"
					fnGSShopPrice = False
				End If
			Set xmlDOM = Nothing
		Else
			fnGSShopPrice = False
			iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRICE-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ��ǰ �ɼ� ���� ����
Function fnGSShopNewOPTSuEdit(iitemid, strParam, iidx, byRef iErrStr)
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
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(�ɼ� ���� ����)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					Set attrPrdlist = strObj.attr
						tenOptcd = attrPrdlist.get(0).supAttrPrdCd
						gsOptcd = attrPrdlist.get(0).attrPrdCd
						sqlStr = ""
						sqlStr = sqlStr & "UPDATE db_item.dbo.tbl_OutMall_regedoption SET "
						sqlStr = sqlStr & "outmalllimitno =  "
						sqlStr = sqlStr & "Case WHEN i.limityn = 'Y' and o.optlimitno - o.optlimitsold <= 5 THEN '0' "
						sqlStr = sqlStr & "	 WHEN i.limityn = 'Y' and o.optlimitno - o.optlimitsold > 5 THEN o.optlimitno - o.optlimitsold - 5 "
						sqlStr = sqlStr & "	 WHEN i.limityn = 'N' THEN '999' END "
						sqlStr = sqlStr & "FROM db_item.dbo.tbl_OutMall_regedoption R  "
						sqlStr = sqlStr & "JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid "
						sqlStr = sqlStr & "JOIN db_item.dbo.tbl_item_option o on i.itemid = o.itemid and o.itemoption = '"&tenOptcd&"' "
						sqlStr = sqlStr & "WHERE R.itemid = '"&iitemid&"' and R.itemoption = '"&tenOptcd&"' and R.mallid = 'gsshop' "
						dbget.Execute sqlStr
					Set attrPrdlist = nothing
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(�ɼ� ���� ����)"
					fnGSShopNewOPTSuEdit = True
		        End If
			Set strObj = nothing
		Else
			fnGSShopNewOPTSuEdit = False
			iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-OPTSuEdit-002]"
		End If
	Set objXML= nothing
	On Error Goto 0
End Function


'��ǰ �ɼ� ���� ����
Function fnGSShopOPTSuEdit(iitemid, strParam, iidx, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, tenOptcd, lp, gsOptcd, sqlStr
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd, attrPrdlist, Assignedrow
	On Error Resume Next
	fnGSShopOPTSuEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				If Err <> 0 Then
					iErrStr =  "ERR||"&iidx&"||GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-OPTSuEdit-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If

				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�
				attrPrdCd	= Split(buf, "|")(5)	'�̼��Ӽ���ǰ�ڵ�^���»�Ӽ���ǰ�ڵ�,�̼��Ӽ���ǰ�ڵ�^���»�Ӽ���ǰ�ڵ�	'�Ӽ��Ķ��Ÿ �������� ������ ���� �� ����

				If resultcode = "S" Then	'����(S)
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(�ɼ� ���� ����)"
					attrPrdlist = split(attrPrdCd,",")
					gsOptcd		= split(attrPrdlist(0),"^")(0)
	                tenOptcd	= split(attrPrdlist(0),"^")(1)
					sqlStr = ""
					sqlStr = sqlStr & "UPDATE db_item.dbo.tbl_OutMall_regedoption SET "
					sqlStr = sqlStr & "outmalllimitno =  "
					sqlStr = sqlStr & "Case WHEN i.limityn = 'Y' and o.optlimitno - o.optlimitsold <= 5 THEN '0' "
					sqlStr = sqlStr & "	 WHEN i.limityn = 'Y' and o.optlimitno - o.optlimitsold > 5 THEN o.optlimitno - o.optlimitsold - 5 "
					sqlStr = sqlStr & "	 WHEN i.limityn = 'N' THEN '999' END "
					sqlStr = sqlStr & "FROM db_item.dbo.tbl_OutMall_regedoption R  "
					sqlStr = sqlStr & "JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid "
					sqlStr = sqlStr & "JOIN db_item.dbo.tbl_item_option o on i.itemid = o.itemid and o.itemoption = '"&tenOptcd&"' "
					sqlStr = sqlStr & "WHERE R.itemid = '"&iitemid&"' and R.itemoption = '"&tenOptcd&"' and R.mallid = 'gsshop' "
					dbget.Execute sqlStr
				Else						'����(E)
				    iErrStr =  "ERR||"&iidx&"||"&resultmsg&"(�ɼ� ���� ����)"
			        Set objXML = Nothing
			        Set xmlDOM = Nothing
			        On Error Goto 0
				    Exit Function
				End If
			Set xmlDOM = Nothing
			fnGSShopOPTSuEdit = True
		Else
			iErrStr =  "ERR||"&iidx&"||GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-OPTSuEdit-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ��ǰ �ɼ� ���º���
Function fnGSShopNewOPTSellEdit(iitemid, strParam, iidx, iitemoption, byRef iErrStr)
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
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(�ɼǻ��º���)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

	                sqlStr = ""
					sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_OutMall_regedoption SET " & VbCRLF
					sqlStr = sqlStr & " outmallsellyn = " & VbCRLF
					sqlStr = sqlStr & " Case WHEN (o.isusing <> 'Y' OR o.optsellyn <> 'Y') THEN 'N'  " & VbCRLF
					sqlStr = sqlStr & " 	 WHEN (i.limityn = 'Y' AND o.optlimitno - o.optlimitsold <= 5) THEN 'N' " & VbCRLF
					sqlStr = sqlStr & " 	 WHEN (R.outmallOptName <> o.optionname) THEN 'N' " & VbCRLF
					sqlStr = sqlStr & " ELSE 'Y' END " & VbCRLF
					sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption R  " & VbCRLF
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid  " & VbCRLF
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_option o on i.itemid = o.itemid and o.itemoption = '"&iitemoption&"'  " & VbCRLF
					sqlStr = sqlStr & " WHERE R.itemid = '"&iitemid&"' and R.itemoption = '"&iitemoption&"' and R.mallid = 'gsshop' "
				    dbget.Execute sqlStr
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(�ɼǻ��º���)"
				End If
			Set strObj = nothing
			fnGSShopNewOPTSellEdit = True
		Else
			iErrStr = "ERR||"&iidx&"||GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-OPTSellEdit-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'��ǰ �ɼ� ���º���
Function fnGSShopOPTSellEdit(iitemid, strParam, iidx, iitemoption, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, tenOptcd, lp, gsOptcd
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd, attrPrdlist, sqlStr
	On Error Resume Next
	fnGSShopOPTSellEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				If Err <> 0 Then
					iErrStr = "GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-OPTSellEdit-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If

				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�
				''���ߵ� �� S->P�� ���� ���� �ִ�.(�ɼǻ�ǰ���� ��ǰ��ǰ���� ���ϴ� ��쿡�� ���� ������ ���� �����ؾߵǴ� �� �ݵ�� ó���ؾ���..

				If resultcode = "S" Then	'����(S)
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(�ɼǻ��º���)"
	                sqlStr = ""
					sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_OutMall_regedoption SET " & VbCRLF
					sqlStr = sqlStr & " outmallsellyn = " & VbCRLF
					sqlStr = sqlStr & " Case WHEN (o.isusing <> 'Y' OR o.optsellyn <> 'Y') THEN 'N'  " & VbCRLF
					sqlStr = sqlStr & " 	 WHEN (i.limityn = 'Y' AND o.optlimitno - o.optlimitsold <= 5) THEN 'N' " & VbCRLF
					sqlStr = sqlStr & " 	 WHEN (R.outmallOptName <> o.optionname) THEN 'N' " & VbCRLF
					sqlStr = sqlStr & " ELSE 'Y' END " & VbCRLF
					sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption R  " & VbCRLF
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid  " & VbCRLF
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_option o on i.itemid = o.itemid and o.itemoption = '"&iitemoption&"'  " & VbCRLF
					sqlStr = sqlStr & " WHERE R.itemid = '"&iitemid&"' and R.itemoption = '"&iitemoption&"' and R.mallid = 'gsshop' "
				    dbget.Execute sqlStr
				Else						'����(E)
				    iErrStr =  "ERR||"&iidx&"||"&resultmsg&"(�ɼǻ��º���)"
			        Set objXML = Nothing
			        Set xmlDOM = Nothing
				    Exit Function
				End If
			Set xmlDOM = Nothing
			fnGSShopOPTSellEdit = True
		Else
			iErrStr =  "ERR||"&iidx&"||GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-OPTSellEdit-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ��ǰ�� ���� ���� �Լ�
Public Function fnGSShopChgNewItemname(iidx, strParam, iitemname, byRef iErrStr)
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
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(��ǰ�� ����)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd

					strSql = ""
					strSql = strSql & " UPDATE R "
					strSql = strSql & " SET itemname = i.itemname "
					strSql = strSql & " ,optionname = o.optionname "
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] R "
					strSql = strSql & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid "
					strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on R.itemid = o.itemid and R.itemoption = o.itemoption "
					strSql = strSql & " WHERE R.idx = '"&iidx&"' "
					strSql = strSql & " and R.mallid= 'gsshop' "
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(��ǰ�� ����)"
		        End If
			Set strObj = nothing
			fnGSShopChgNewItemname = true
		Else
			iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-NMEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function fnGSShopChgItemname(iidx, strParam, iitemname, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, strSql
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopChgItemname = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				If Err <> 0 Then
					iErrStr =  "ERR||"&iidx&"||GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-NMEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�

				If resultcode = "S" Then
					strSql = ""
					strSql = strSql & " UPDATE R "
					strSql = strSql & " SET itemname = i.itemname "
					strSql = strSql & " ,optionname = o.optionname "
					strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] R "
					strSql = strSql & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid "
					strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on R.itemid = o.itemid and R.itemoption = o.itemoption "
					strSql = strSql & " WHERE R.idx = '"&iidx&"' "
					strSql = strSql & " and R.mallid= 'gsshop' "
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(��ǰ�� ����)"
				Else
					iErrStr =  "ERR||"&iidx&"||"&resultmsg
				End If

			Set xmlDOM = Nothing
			fnGSShopChgItemname = True
		Else
			iErrStr =  "ERR||"&iidx&"||GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-NMEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New �̹��� ���� ���� �Լ�
Function fnGSShopNewImageEdit(iidx, strParam, iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
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
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(�̹�������)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(�̹�������)"
		        End If
			Set strObj = nothing
			fnGSShopNewImageEdit = true
		Else
			iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-IMGEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ��ǰ���� ���� ���� �Լ�
Function fnGSShopNewItemInfoEdit(iidx, strParam, iErrStr)
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
				resultcode	= strObj.success
				resultmsg	= replaceMsg(strObj.msg)
				If resultcode <> "True" Then
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(��ǰ����)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(��ǰ����)"
		        End If
			Set strObj = nothing
			fnGSShopNewItemInfoEdit = true
		Else
			iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-INFOEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

Function fnGSShopImageEdit(iidx, strParam, iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopImageEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				If Err <> 0 Then
					iErrStr =  "ERR||"&iidx&"||GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-IMGEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�

				If resultcode = "S" Then
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(�̹��� ����)"
				Else
					iErrStr =  "ERR||"&iidx&"||"&resultmsg
				End If

			Set xmlDOM = Nothing
			fnGSShopImageEdit = True
		Else
			iErrStr =  "ERR||"&iidx&"||GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-IMGEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ���û�ǰ ���� ����
Function fnGSShopNewContentsEdit(iidx, strParam, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, iRbody, strObj
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopNewContentsEdit = False

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
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(��ǰ�������)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(��ǰ�������)"
		        End If
			Set strObj = nothing
			fnGSShopNewContentsEdit = true
		Else
			iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-CONTEDIT-002]"
		End If
End Function

'���û�ǰ ���� ����
Function fnGSShopContentsEdit(iidx, strParam, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopContentsEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				If Err <> 0 Then
					iErrStr =  "ERR||"&iidx&"||GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-CONTEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�

				If resultcode = "S" Then
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(��ǰ���� ����)"
				Else
					iErrStr =  "ERR||"&iidx&"||"&resultmsg
				End If

			Set xmlDOM = Nothing
			fnGSShopContentsEdit = True
		Else
			iErrStr =  "ERR||"&iidx&"||GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-CONTEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'New ���ΰ���׸� ����
Function fnGSShopNewInfodivEdit(iidx, strParam, byRef iErrStr)
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
					iErrStr = "ERR||"&iidx&"||"&resultmsg&"(���ΰ���׸� ����)"
				Else
					prdCd		= strObj.prdCd
					supPrdCd	= strObj.supPrdCd
					supCd 		= strObj.supCd
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(���ΰ���׸� ����)"
		        End If
			Set strObj = nothing
			fnGSShopNewInfodivEdit = true
		Else
			iErrStr = "ERR||"&iidx&"||GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-DIVEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'���ΰ���׸� ����
Function fnGSShopInfodivEdit(iidx, strParam, byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	fnGSShopInfodivEdit = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					'rw buf
				End If

				If Err <> 0 Then
					iErrStr =  "ERR||"&iidx&"||GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-DIVEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�

				If resultcode = "S" Then
					iErrStr =  "OK||"&iidx&"||"&resultmsg&"(���ΰ���׸� ����)"
				Else
					iErrStr =  "ERR||"&iidx&"||"&resultmsg
				End If

			Set xmlDOM = Nothing
			fnGSShopInfodivEdit = True
		Else
			iErrStr =  "ERR||"&iidx&"||GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-DIVEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function
'############################################## ���� �����ϴ� API �Լ� ���� �� ############################################

'################################################# �� ��� �� �Ķ���� ���� ###############################################
'ǰ�� �Ķ��Ÿ
Function getGSShopSellynParameter(iidx, ichgSellYn)
	Dim strRst, strSql, newCode

	strSql = ""
	strSql = strSql & " SELECT TOP 1 convert(varchar(30),itemid)+convert(varchar(30),itemoption) as newCode "
	strSql = strSql & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] "
	strSql = strSql & " where idx = '"&iidx&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		newCode	= rsget("newCode")
	End If
	rsget.close

	strRst = ""
	strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
	strRst = strRst & "&modGbn=S"														'(*)�������� S : �ǸŻ��� ����
	strRst = strRst & "&regId="&COurRedId												'(*)�����
	'��ǰ�⺻(prdBaseInfo)
	strRst = strRst & "&supPrdCd="&newCode												'(*)���»��ǰ�ڵ�
	strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
	'��ǰ����(prdPrc)

	If ichgSellYn = "Y" Then
		strRst = strRst & "&saleEndDtm=29991231235959"									'(*)�Ǹ������Ͻ� | ��ǰ�� �ߴ�(�Ǹ�����)�Ϸ��� �ߴܽ����� �Ǹ������Ͻø� �Է��մϴ�.
	ElseIf (ichgSellYn = "N") Then
		strRst = strRst & "&saleEndDtm="&FormatDate(now(), "00000000000000")			'(*)�Ǹ������Ͻ� | ��ǰ�� �ߴ�(�Ǹ�����)�Ϸ��� �ߴܽ����� �Ǹ������Ͻø� �Է��մϴ�.
	End If
	strRst = strRst & "&attrSaleEndStModYn=N"											'(*)�Ӽ��Ǹ�������¼������� | �Ӽ�����(S) ��ǰ�ǸŻ��¸� ������ �� ����ϴ� �׸�����, ��ǰ������ ���� �� ���� �� �Ӽ���ǰ�� ���µ� �Բ� ���� �� �����Ϸ��� Y, ��ǰ�����Ϳ� �Ӽ� ������ ���º��� ���� �ÿ� N

	getGSShopSellynParameter = strRst
End Function

Public Function getGSShopPriceParameter(iidx, byref mustprice)
	Dim strRst, strSql
	Dim sellcash, orgprice, buycash, optaddprice, newCode
	Dim GetTenTenMargin

	strSql = ""
	strSql = strSql & " SELECT TOP 1 sellcash, buycash, orgprice, o.optaddprice, convert(varchar(30),m.itemid) + convert(varchar(30),m.itemoption) as newCode "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
	strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as m on i.itemid = m.itemid and o.itemoption = m.itemoption "
	strSql = strSql & " WHERE m.idx = '"&iidx&"' "
	strSql = strSql & " and m.mallid = 'gsshop' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		sellcash	= rsget("sellcash")
		orgprice	= rsget("orgprice")
		buycash		= rsget("buycash")
		optaddprice	= rsget("optaddprice")
		newCode		= rsget("newCode")
	Else
		getGSShopPriceParameter = ""
		Exit Function
		response.end
	End If
	rsget.close

	GetTenTenMargin = CLng(10000 - buycash / sellcash * 100 * 100) / 100
	If GetTenTenMargin < CMAXMARGIN Then
		MustPrice = orgprice + optaddprice
	Else
		MustPrice = sellcash + optaddprice
	End If

	'���� ���� �� �ݺ�����Ʈ �Ǽ�
	strRst = ""
	strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
	strRst = strRst & "&modGbn=P"														'(*)�������� P : ���� ����
	strRst = strRst & "&regId="&COurRedId												'(*)�����
	strRst = strRst & "&regSubjCd=SUP"													'(*)�����ü�ڵ� | ���� ������ ��� : MD, ���»簡 ������ ��� : SUP
	'��ǰ�⺻(prdBaseInfo)
	strRst = strRst & "&supPrdCd="&newCode												'(*)���»��ǰ�ڵ�
	strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
	strRst = strRst & "&subSupCd="&COurCompanyCode										'(*)�������»��ڵ� | �������»簡 ������ supCd�� ������ �Է�
	'��ǰ����(prdPrc)
	strRst = strRst & "&prdPrcValidStrDtm="&FormatDate(now(), "00000000000000")			'(*)��ȿ�����Ͻ�
	strRst = strRst & "&prdPrcValidEndDtm=29991231235959"								'(*)��ȿ�����Ͻ�
	strRst = strRst & "&prdPrcSalePrc="&Clng(GetRaiseValue(MustPrice/10)*10)			'(*)�ǸŰ���
	'strRst = strRst & "&prdPrcPrchPrc="													'(SYS)���԰��� | (SYS�� �����ʿ��� �ڵ����� �������ִ� �ڵ� �� ���� ���մϴ�. Null�� �����ֽø� �˴ϴ�.)
	strRst = strRst & "&prdPrcSupGivRtamtCd=01"											'(*)���»�������/���ڵ� | 01 : ��
	strRst = strRst & "&prdPrcSupGivRtamt="&getGSShopSuplyPrice_update(MustPrice)		'(*)���»�������/�� | �⺻�� : �ǸŰ�*(1-0.12)
	getGSShopPriceParameter = strRst
End Function

Public Function getGSShopItemnameParameter(iidx, byref iitemname)
	Dim strRst, chgname, strSql, newitemname, itemnameChange, newCode
	strSql = ""
	strSql = strSql & " SELECT TOP 1 M.itemid, convert(varchar(30),m.itemid) + convert(varchar(30),m.itemoption) as newCode, isnull(M.newitemname, '') as newitemname, isnull(M.itemnameChange, '') as itemnameChange "
	strSql = strSql & "	FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M "
	strSql = strSql & "	JOIN db_etcmall.[dbo].[tbl_gsshopAddoption_regitem] as R on M.idx = R.midx "
	strSql = strSql & "	WHERE M.idx = '"&iidx&"' "
	strSql = strSql & "	and M.mallid = 'gsshop' "
	strSql = strSql & "	and (R.GSShopStatCd=3 OR R.GSShopStatCd=7) "
	strSql = strSql & " and R.GSShopGoodNo is Not Null "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.Eof Then
		newCode			= rsget("newCode")
		newitemname		= rsget("newitemname")
		itemnameChange	= rsget("itemnameChange")
	End If
	rsget.close

	If itemnameChange = "" Then
		iitemname = newitemname
	Else
		iitemname = itemnameChange
	End If

	chgname = ""
	chgname = "[�ٹ�����]"&replace(iitemname,"'","")		'���� ��ǰ�� �տ� [�ٹ�����] �̶�� ����
	chgname = replace(chgname,"&#8211;","-")
	chgname = replace(chgname,"~","-")
	chgname = replace(chgname,"&","��")
	chgname = replace(chgname,"<","[")
	chgname = replace(chgname,">","]")
	chgname = replace(chgname,"%","����")
	chgname = replace(chgname,"+","%2B")
	chgname = replace(chgname,"[������]","")
	chgname = replace(chgname,"[���� ���]","")

	strRst = ""
	strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
	strRst = strRst & "&modGbn=N"														'(*)�������� N : �����ǰ�� ����
	strRst = strRst & "&regId="&COurRedId												'(*)�����
	'��ǰ�⺻(prdBaseInfo)
	strRst = strRst & "&supPrdCd="&newCode												'(*)���»��ǰ�ڵ�
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