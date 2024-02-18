<%
Public homeplusAPIURL
Public strInterface
Public homeplusVenderID
Public homepluspasswd

IF application("Svr_Info") = "Dev" THEN
	homeplusAPIURL = "http://112.108.7.201:7006/services/API2?wsdl"
	strInterface = "http://112.108.7.201:7006/api/services/API2"
	homeplusVenderID = "292811"
	homepluspasswd = "qwer1234"
Else
	homeplusAPIURL = "http://api.direct.homeplus.co.kr:17004/services/API2?wsdl"
	strInterface = "http://api.direct.homeplus.co.kr:17004/api/services/API2"
	homeplusVenderID = "292811"
	homepluspasswd = "cube1010!!"
End if
'############################################## ���� �����ϴ� API �Լ� ���� ##############################################
Function getXMLString(mode)
	Dim strRst
	If mode = "login" Then
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/""  xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:ns1=""http://xml.apache.org/axis/"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:"&mode&" xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<venderId>"&homeplusVenderID&"</venderId>"
		strRst = strRst & "			<passwd>"&homepluspasswd&"</passwd>"
		strRst = strRst & "		</m:"&mode&">"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
	ElseIf mode = "getCategories" Then
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/""  xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:"&mode&" xmlns:m=""" & strInterface & """></m:"&mode&">"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
	End If
	getXMLString = strRst
End Function

'ī�װ� API
Function HomeplusCategoryAPI()
    Dim mode : mode = "getCategories"
	Dim xmlStr : xmlStr = getXMLString(mode)
	Dim objXML, xmlDOM, xmlDOM2, retCode, resultmsg, hplist, SubNodes, strSql
	Dim hDIVISION, hGROUP, hDEPT, hCLASS, hSUBCLASS, hDIV_NAME, hGROUP_NAME, hDEPT_NAME, hCLASS_NAME, hSUB_NAME, hCATEGORY_ID, hCATEGORY_NAME
	Dim AssignedRow

    On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#login"
		objXML.setTimeouts 5000,90000,90000,90000
		objXML.send(getXMLString("login"))
	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.ValidateOnParse= True
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

		If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then	'�α��� �����̶��
			objXML.open "post", "" & homeplusAPIURL & "", False
			objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
			objXML.setRequestHeader "Content-Length", Len(xmlStr)
			objXML.setRequestHeader "SOAPAction", strInterface & "#"&mode
			objXML.send(xmlStr)
			If objXML.Status = "200" Then
				Set xmlDOM2 = Server.CreateObject("MSXML.DOMDocument")
					xmlDOM2.async = False
					xmlDOM2.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
					retCode = xmlDOM2.getElementsByTagName("ns1:code").item(0).text
					resultmsg = xmlDOM2.getElementsByTagName("ns1:message").item(0).text
					rw retCode
					rw resultmsg
					response.end
'				If retCode = "E0000" Then
'					strSql = ""
'					strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_homeplus_dftcategory "
'					dbget.Execute(strSql)
'					Set hplist = xmlDOM2.getElementsByTagName("ns1:list")
'						For each SubNodes in hplist
'							hDIVISION		= Trim(SubNodes.getElementsByTagName("ns1:DIVISION").item(0).text)		'�ֻ��� �з��ڵ�
'							hGROUP			= Trim(SubNodes.getElementsByTagName("ns1:GROUP").item(0).text)			'DIVISION ���� �з� �ڵ�
'							hDEPT			= Trim(SubNodes.getElementsByTagName("ns1:DEPT").item(0).text)			'GROUP ���� �з� �ڵ�
'							hCLASS			= Trim(SubNodes.getElementsByTagName("ns1:CLASS").item(0).text)			'DEPT ���� �з� �ڵ�
'							hSUBCLASS		= Trim(SubNodes.getElementsByTagName("ns1:SUBCLASS").item(0).text)		'CLASS ���� �з� �ڵ�
'							hDIV_NAME		= Trim(SubNodes.getElementsByTagName("ns1:DIV_NAME").item(0).text)		'DIVISION �з���
'							hGROUP_NAME		= Trim(SubNodes.getElementsByTagName("ns1:GROUP_NAME").item(0).text)	'GROUP �з���
'							hDEPT_NAME		= Trim(SubNodes.getElementsByTagName("ns1:DEPT_NAME").item(0).text)		'DEPT �з���
'							hCLASS_NAME		= Trim(SubNodes.getElementsByTagName("ns1:CLASS_NAME").item(0).text)	'CLASS �з���
'							hSUB_NAME		= Trim(SubNodes.getElementsByTagName("ns1:SUB_NAME").item(0).text)		'SUBCLASS �з���
'							hCATEGORY_ID	= Trim(SubNodes.getElementsByTagName("ns1:CATEGORY_ID").item(0).text)	'��ǰ����������ø� ���� ī�װ� ���̵�
'							hCATEGORY_NAME	= Trim(SubNodes.getElementsByTagName("ns1:CATEGORY_NAME").item(0).text)	'��ǰ����������ø� ���� ī�װ� ��
'
'							strSql = ""
'							strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_homeplus_dftcategory (hDIVISION, hGROUP, hDEPT, hCLASS, hSUBCLASS, hDIV_NAME, hGROUP_NAME, hDEPT_NAME, hCLASS_NAME, hSUB_NAME, hCATEGORY_ID, hCATEGORY_NAME) VALUES " & VBCRLF
'							strSql = strSql & " ('"&db2html(hDIVISION)&"', '"&db2html(hGROUP)&"', '"&db2html(hDEPT)&"', '"&db2html(hCLASS)&"', '"&db2html(hSUBCLASS)&"', '"&db2html(hDIV_NAME)&"', '"&hGROUP_NAME&"', '"&db2html(hDEPT_NAME)&"', '"&db2html(hCLASS_NAME)&"', '"&db2html(hSUB_NAME)&"', '"&db2html(hCATEGORY_ID)&"', '"&db2html(hCATEGORY_NAME)&"')" & VBCRLF
'							dbget.Execute strSql, AssignedRow
'						Next
'					Set hplist = nothing
'				End If
				Set xmlDOM2 = nothing
			End If
		End If
		Set xmlDOM = nothing
	End If
	Set objXML = nothing
End Function

'��ǰ���
Function fnHomeplusOneItemReg(iitemid, istrParam, byRef iErrStr, iSellCash, ihomeplusSellYn, ilimityn, ilimitno, ilimitsold, iitemname, mode)
	Dim objXML, xmlDOM, xmlDOM2, SubNodes, strSql
	Dim xmlStr : xmlStr = getXMLString("login")
	Dim retCode, homegoodNo, resultmsg, optlist, s_ITEMNO, i_ITEMNO, s_OPTION_NAME
	Dim AssignedRow
	Dim Tlimitno, Tlimitsold, Tlimityn, Titemsu
    On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#login"
		objXML.setTimeouts 5000,90000,90000,90000
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.ValidateOnParse= True
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then	'�α��� �����̶��
			objXML.open "post", "" & homeplusAPIURL & "", False
			objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
			objXML.setRequestHeader "Content-Length", Len(strParam)
			objXML.setRequestHeader "SOAPAction", strInterface & "#" &mode
			objXML.send(istrParam)
			If objXML.Status = "200" Then
				Set xmlDOM2 = Server.CreateObject("MSXML.DOMDocument")
					xmlDOM2.async = False
					xmlDOM2.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
					retCode		= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/createNewProductResponse/ns1:createNewProductReturn/ns1:code").text
					homegoodNo	= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/createNewProductResponse/ns1:createNewProductReturn/ns1:i_STYLENO").text
					resultmsg	= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/createNewProductResponse/ns1:createNewProductReturn/ns1:message").text

				If retCode = "E0000" Then	'����(E0000)
					'��ǰ���翩�� Ȯ��
					strSql = "SELECT COUNT(itemid) FROM db_etcmall.dbo.tbl_homeplus_regItem WHERE itemid='" & iitemid & "'"
					rsget.Open strSql, dbget, 1
					If rsget(0) > 0 Then
						'// ���� -> ����
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCRLF
						strSql = strSql & "	Set homeplusLastUpdate = getdate() "  & VbCRLF
						strSql = strSql & "	, homeplusGoodNo = '" & homegoodNo & "'"  & VbCRLF
						strSql = strSql & "	, homeplusPrice = " &iSellCash& VbCRLF
						strSql = strSql & "	, accFailCnt = 0"& VbCRLF
						strSql = strSql & "	, homeplusRegdate = isNULL(homeplusRegdate, getdate())"
						If (homegoodNo <> "") Then
						    strSql = strSql & "	, homeplusstatCD = '7'"& VbCRLF					'��ϿϷ�(�ӽ�)
						Else
							strSql = strSql & "	, homeplusstatCD = '1'"& VbCRLF					'���۽õ�
						End If
						strSql = strSql & "	From db_etcmall.dbo.tbl_homeplus_regItem R"& VbCRLF
						strSql = strSql & " Where R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
					Else
						'// ���� -> �űԵ��
						strSql = ""
						strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_homeplus_regItem "
						strSql = strSql & " (itemid, regitemname, reguserid, homeplusRegdate, homeplusLastUpdate, homeplusGoodNo, homeplusPrice, homeplusSellYn, homeplusStatCd) VALUES " & VbCRLF
						strSql = strSql & " ('" & iitemid & "'" & VBCRLF
						strSql = strSql & " , '" & iitemname & "'" &_
						strSql = strSql & " , '" & session("ssBctId") & "'" &_
						strSql = strSql & " , getdate(), getdate()" & VBCRLF
						strSql = strSql & " , '" & homegoodNo & "'" & VBCRLF
						strSql = strSql & " , '" & iSellCash & "'" & VBCRLF
						strSql = strSql & " , '" & ihomeplusSellYn & "'" & VBCRLF
						If (homegoodNo <> "") Then
						    strSql = strSql & ",'7'"											'��ϿϷ�(�ӽ�)
						Else
						    strSql = strSql & ",'1'"											'���۽õ�
						End If
						strSql = strSql & ")"
						dbget.Execute(strSql)
					End If
					rsget.Close
		
					Set optlist = xmlDOM2.SelectNodes("soapenv:Envelope/soapenv:Body/createNewProductResponse/ns1:createNewProductReturn/ns1:ITEMRESULT/ns1:ITEMRESULT")
						For each SubNodes in optlist
							s_ITEMNO		= Trim(SubNodes.SelectSingleNode("ns1:s_ITEMNO").text)		'�ٹ����� �ɼ��ڵ�
							i_ITEMNO		= Trim(SubNodes.SelectSingleNode("ns1:i_ITEMNO").text)		'Ȩ�÷��� �ɼ��ڵ�
							s_OPTION_NAME	= Trim(SubNodes.SelectSingleNode("ns1:s_OPTION_NAME").text)	'�ɼǸ�
							If s_ITEMNO = "0000" Then
								Tlimitno		= ilimitno
								Tlimitsold		= ilimitsold
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
								sqlStr = ""
								sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_OutMall_regedoption " & VBCRLF
								sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
								sqlStr = sqlStr & " VALUES " & VBCRLF
								sqlStr = sqlStr & " ('"&iitemid&"',  '"&s_ITEMNO&"', 'homeplus', '"&i_ITEMNO&"', '"&html2db(s_OPTION_NAME)&"', 'Y', '"&ilimityn&"', '"&Titemsu&"', '0', getdate()) "
								dbget.Execute sqlStr
							Else
								sqlStr = ""
								sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_OutMall_regedoption " & VBCRLF
								sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
								sqlStr = sqlStr & " SELECT itemid, itemoption, 'homeplus', '"&i_ITEMNO&"', optionname, optsellyn, 'Y', " & VBCRLF
								sqlStr = sqlStr & " Case WHEN optlimityn = 'Y' AND optlimitno - optlimitsold <= 5 THEN '0' " & VBCRLF
								sqlStr = sqlStr & " 	 WHEN optlimityn = 'Y' AND optlimitno - optlimitsold > 5 THEN optlimitno - optlimitsold - 5 " & VBCRLF
								sqlStr = sqlStr & " 	 WHEN optlimityn = 'N' THEN '999' End " & VBCRLF
								sqlStr = sqlStr & " , '0', getdate() " & VBCRLF
								sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option " & VBCRLF
								sqlStr = sqlStr & " WHERE itemid= '"&iitemid&"' " & VBCRLF
								sqlStr = sqlStr & " and itemoption = '"& s_ITEMNO &"' "
								dbget.Execute sqlStr
							End If
						Next
					Set optlist = nothing
					strSql = ""
					strSql = strSql & " UPDATE R " & VBCRLF
					strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VBCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_homeplus_regItem R " & VBCRLF
					strSql = strSql & " Join ( " & VBCRLF
					strSql = strSql & " 	SELECT R.itemid, count(*) as CNT, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt "
					strSql = strSql & " 	FROM db_etcmall.dbo.tbl_homeplus_regItem R " & VBCRLF
					strSql = strSql & " 	JOIN db_etcmall.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = 'homeplus' and Ro.itemid = " &iitemid & VBCRLF
					strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid " & VBCRLF
					dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||��ϼ���(��ǰ���)"
				Else						'����(E)
					iErrStr = "ERR||"&iitemid&"||"&resultmsg&"(��ǰ���)"
				End If
				Set xmlDOM2 = nothing
			Else
				iErrStr = "ERR||"&iitemid&"||Homeplus ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG-001]"
			End If
		Else
			iErrStr = "ERR||"&iitemid&"||Homeplus �α��� ����[ERR-REG-001]"
		End If
		Set xmlDOM = nothing
	End If
	Set objXML = nothing
	On Error Goto 0
End Function

'Ȩ�÷��� ���� ����
Function fnHomeplusSellyn(iitemid, ichgSellYn, istrParam, byRef iErrStr, mode)
	Dim objXML, xmlDOM, xmlDOM2, strSql
	Dim retCode, resultmsg
	Dim xmlStr : xmlStr = getXMLString("login")
    On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#login"
		objXML.setTimeouts 5000,90000,90000,90000
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.ValidateOnParse= True
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then	'�α��� �����̶��
			objXML.open "post", "" & homeplusAPIURL & "", False
			objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
			objXML.setRequestHeader "Content-Length", Len(xmlStr)
			objXML.setRequestHeader "SOAPAction", strInterface & "#"&mode
			objXML.send(istrParam)
			If objXML.Status = "200" Then
				Set xmlDOM2 = Server.CreateObject("MSXML.DOMDocument")
					xmlDOM2.async = False
					xmlDOM2.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
					retCode		= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/setProductStatusResponse/ns1:setProductStatusReturn/ns1:code").text
					resultmsg	= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/setProductStatusResponse/ns1:setProductStatusReturn/ns1:message").text
				If retCode = "E0000" Then	'����(E0000)
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_homeplus_regItem " & VbCRLF
					strSql = strSql & " SET homeplusLastUpdate = getdate() " & VbCRLF
					strSql = strSql & " ,homeplusSellYn = '" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " ,accFailCNT=0" & VbCRLF
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute(strSql)
					If ichgSellYn = "N" Then
						iErrStr = "OK||"&iitemid&"||ǰ��ó��"
					Else
						iErrStr = "OK||"&iitemid&"||�Ǹ������� ����"
					End If
				Else						'����(E)
				    iErrStr = "ERR||"&iitemid&"||"&resultmsg
				End If
				Set xmlDOM2 = nothing
			Else
				iErrStr = "ERR||"&iitemid&"||Homeplus ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EditSellYn-001]"
			End If			
		Else
			iErrStr = "ERR||"&iitemid&"||Homeplus �α��� ����[ERR-EditSellYn-001]"
		End If
		Set xmlDOM = nothing
	End If
	Set objXML = nothing
	On Error Goto 0
End Function

'���� ����
Function fnHomeplusOneItemEdit(iitemid, iHomeplusGoodNo, byRef iErrStr, istrParam, mode)
	Dim objXML, xmlDOM, xmlDOM2, strSql
	Dim retCode, resultmsg
	Dim xmlStr : xmlStr = getXMLString("login")
    On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#login"
		objXML.setTimeouts 5000,90000,90000,90000
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.ValidateOnParse= True
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then	'�α��� �����̶��
			objXML.open "post", "" & homeplusAPIURL & "", False
			objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
			objXML.setRequestHeader "Content-Length", Len(xmlStr)
			objXML.setRequestHeader "SOAPAction", strInterface & "#" &mode
			objXML.send(istrParam)
		'			response.write objXML.ResponseText
		'			response.end	
			If objXML.Status = "200" Then
				Set xmlDOM2 = Server.CreateObject("MSXML.DOMDocument")
					xmlDOM2.async = False
					xmlDOM2.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		'			response.write objXML.ResponseText
		'			response.end
					retCode		= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/updateProductResponse/ns1:updateProductReturn/ns1:code").text
					resultmsg	= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/updateProductResponse/ns1:updateProductReturn/ns1:message").text
		
				If retCode = "E0000" Then	'����(E0000)�̸� ����Ƚ�� �ʱ�ȭ
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_homeplus_regItem " & VbCRLF
					strSql = strSql & " SET accFailCnt = 0, regitemname = B.itemname "& VbCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_homeplus_regItem A "& VbCRLF
					strSql = strSql & " JOIN db_item.dbo.tbl_item B on A.itemid = B.itemid "& VbCRLF
					strSql = strSql & " WHERE A.itemid='" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||��������(��ǰ����)"
				Else						'����(E)
				    iErrStr = "ERR||"&iitemid&"||"&resultmsg
				End If
				Set xmlDOM2 = nothing
			Else
				iErrStr = "ERR||"&iitemid&"||Homeplus ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-ITEMNAME-001]"
			End If			
		Else
			iErrStr = "ERR||"&iitemid&"||Homeplus �α��� ����[ERR-ITEMNAME-002]"
		End If
		Set xmlDOM = nothing
	End If
	Set objXML = nothing
	On Error Goto 0
End Function

'������ ����
Function fnHomeplusOneItemOPTEdit(iitemid, iHomeplusGoodNo, byRef iErrStr, istrParam, imustprice, mode)
	Dim objXML, xmlDOM, xmlDOM2, strSql
	Dim retCode, resultmsg
	Dim xmlStr : xmlStr = getXMLString("login")
    On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#login"
		objXML.setTimeouts 5000,90000,90000,90000
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.ValidateOnParse= True
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then	'�α��� �����̶��
			objXML.open "post", "" & homeplusAPIURL & "", False
			objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
			objXML.setRequestHeader "Content-Length", Len(xmlStr)
			objXML.setRequestHeader "SOAPAction", strInterface & "#" &mode
			objXML.send(istrParam)
	
			If objXML.Status = "200" Then
				Set xmlDOM2 = Server.CreateObject("MSXML.DOMDocument")
					xmlDOM2.async = False
					xmlDOM2.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		'			response.write objXML.ResponseText
		'			response.end
					retCode		= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/updateProductItemResponse/ns1:updateProductItemReturn/ns1:code").text
					resultmsg	= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/updateProductItemResponse/ns1:updateProductItemReturn/ns1:message").text
		
				If retCode = "E0000" Then	'����(E0000)�̸� ����Ƚ�� �ʱ�ȭ
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_homeplus_regItem " & VbCRLF
					strSql = strSql & " SET accFailCnt = 0 " & VbCRLF
					strSql = strSql & " , homeplusLastUpdate = getdate() " & VbCRLF
					strSql = strSql & " , homeplusprice = '"&imustprice&"' " & VbCRLF
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||��������(�����ܼ���)"
				Else						'����(E)
				    iErrStr = "ERR||"&iitemid&"||"&resultmsg
				End If
				Set xmlDOM2 = nothing
			Else
				iErrStr = "ERR||"&iitemid&"||Homeplus ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDIT-001]"
			End If			
		Else
			iErrStr = "ERR||"&iitemid&"||Homeplus �α��� ����[ERR-EDIT-001]"
		End If
		Set xmlDOM = nothing
	End If
	Set objXML = nothing
	On Error Goto 0
End Function
'�̹��� ����
Function fnHomeplusOneItemIMGEdit(iitemid, iHomeplusGoodNo, byRef iErrStr, istrParam, mode)
	Dim objXML, xmlDOM, xmlDOM2, strSql
	Dim retCode, resultmsg
	Dim xmlStr : xmlStr = getXMLString("login")
    On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#login"
		objXML.setTimeouts 5000,90000,90000,90000
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.ValidateOnParse= True
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then	'�α��� �����̶��
			objXML.open "post", "" & homeplusAPIURL & "", False
			objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
			objXML.setRequestHeader "Content-Length", Len(xmlStr)
			objXML.setRequestHeader "SOAPAction", strInterface & "#" &mode
			objXML.send(istrParam)
			If objXML.Status = "200" Then
				Set xmlDOM2 = Server.CreateObject("MSXML.DOMDocument")
					xmlDOM2.async = False
					xmlDOM2.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
'response.end
					retCode		= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/updateImageResponse/ns1:updateImageReturn/ns1:code").text
					resultmsg	= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/updateImageResponse/ns1:updateImageReturn/ns1:message").text
				If retCode = "E0000" Then	'����(E0000)
					iErrStr =  "OK||"&iitemid&"||��������(�̹���)"
				Else						'����(E)
				    iErrStr = "ERR||"&iitemid&"||"&resultmsg
				End If
				Set xmlDOM2 = nothing
			Else
				iErrStr = "ERR||"&iitemid&"||Homeplus ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EditImg-001]"
			End If			
		Else
			iErrStr = "ERR||"&iitemid&"||Homeplus �α��� ����[ERR-EditImg-001]"
		End If
		Set xmlDOM = nothing
	End If
	Set objXML = nothing
	On Error Goto 0
End Function
'��ǰ ��ȸ
Function fnHomeplusOneItemView(iitemid, iHomeplusGoodNo, byRef iErrStr, istrParam, mode)
	Dim objXML, xmlDOM, xmlDOM2, strSql
	Dim retCode, resultmsg, regedItemStatus, actCnt
	Dim oneProdInfo, SubNodes, AssignedRow, StockQty
	Dim hplOptStatus, hplOptno, regedOpt10x10OptNo, regedOpt10x10OptNm
	Dim Soptioncnt, Slimityn, Slimitno, Slimitsold 
	Dim xmlStr : xmlStr = getXMLString("login")
    On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#login"
		objXML.setTimeouts 5000,90000,90000,90000
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.ValidateOnParse= True
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then	'�α��� �����̶��
			objXML.open "post", "" & homeplusAPIURL & "", False
			objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
			objXML.setRequestHeader "Content-Length", Len(xmlStr)
			objXML.setRequestHeader "SOAPAction", strInterface & "#" &mode
			objXML.send(istrParam)
			If objXML.Status = "200" Then
				Set xmlDOM2 = Server.CreateObject("MSXML.DOMDocument")
					xmlDOM2.async = False
					xmlDOM2.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		'			response.write objXML.ResponseText
		'			response.end
					retCode			= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/searchProductResponse/ns1:searchProductReturn/ns1:code").text
					resultmsg		= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/searchProductResponse/ns1:searchProductReturn/ns1:message").text
					regedItemStatus	= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/searchProductResponse/ns1:searchProductReturn/ns1:SALE").text
				If retCode = "E0000" Then	'����(E0000)
					strSql = ""
					strSql = strSql & " SELECT TOP 1 i.optioncnt, i.limityn, i.limitno, i.limitsold "
					strSql = strSql & " FROM db_item.dbo.tbl_item as i "
					strSql = strSql & " JOIN db_etcmall.dbo.tbl_homeplus_regitem as r on i.itemid = r.itemid "
					strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
			        rsget.Open strSql, dbget
					If Not(rsget.EOF or rsget.BOF) Then
						Soptioncnt	= rsget("optioncnt")
						Slimityn	= rsget("limityn")
						Slimitno	= rsget("limitno")
						Slimitsold	= rsget("limitsold")
					End If
					rsget.Close

					Set oneProdInfo = xmlDOM2.SelectNodes("soapenv:Envelope/soapenv:Body/searchProductResponse/ns1:searchProductReturn/ns1:ITEMRESULT/ns1:ITEMRESULT")
						For Each SubNodes In oneProdInfo
							hplOptStatus			= SubNodes.SelectSingleNode("ns1:SALE").text
							hplOptno				= SubNodes.SelectSingleNode("ns1:i_ITEMNO").text
							regedOpt10x10OptNo		= SubNodes.SelectSingleNode("ns1:s_ITEMNO").text
							regedOpt10x10OptNm		= SubNodes.SelectSingleNode("ns1:s_OPTION_NAME").text
		
							If Soptioncnt = 0 Then		'��ǰ�̶��
								If Slimityn = "Y" Then
									StockQty = Slimitno - Slimitsold - 5
								Else
									StockQty = 999
								End If
							Else												'�ɼ��̶��
								If Slimityn = "Y" Then
									strSql = ""
									strSql = strSql & " SELECT CASE WHEN (optlimitno - optlimitsold) <= 5 Then '0' Else (optlimitno - optlimitsold - 5) End as StockQty "
									strSql = strSql & " FROM db_item.dbo.tbl_item_option  "
									strSql = strSql & " WHERE itemid='"&iitemid&"' and itemoption = '"&regedOpt10x10OptNo&"' "
							        rsget.Open strSql, dbget
									If Not(rsget.EOF or rsget.BOF) Then
										StockQty = rsget("StockQty")
									Else
										StockQty = 0
									End If
									rsget.Close
								Else
									StockQty = 999
								End If
							End If
							'1.������ �μ�Ʈ ������ ������Ʈ
							strSql = ""
							strSql = strSql & " IF Exists(SELECT * FROM db_item.dbo.tbl_OutMall_regedoption where itemid='"&iitemid&"' and itemoption = '"&regedOpt10x10OptNo&"' and outmallOptCode = '"&hplOptno&"' and mallid = 'homeplus') "
							strSql = strSql & " BEGIN"& VbCRLF
							strSql = strSql & " UPDATE oP "
						    strSql = strSql & " SET outmallOptName='"&html2DB(regedOpt10x10OptNm)&"'"&VbCRLF
							strSql = strSql & " ,outmallOptCode='"&hplOptno&"'"&VbCRLF
						    strSql = strSql & " ,lastupdate=getdate()"&VbCRLF
						    strSql = strSql & " ,outMallSellyn='"&Chkiif(hplOptStatus="true", "Y", "N")&"'"&VbCRLF
						    strSql = strSql & " ,outmalllimityn='Y'"&VbCRLF
						    strSql = strSql & " ,outMallLimitNo="&StockQty&VbCRLF
						    strSql = strSql & " ,checkdate=getdate()"&VbCRLF
						    strSql = strSql & " FROM db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
						    strSql = strSql & " WHERE itemid="&iitemid&VbCRLF
						    strSql = strSql & " and convert(int, outmallOptCode)='"&hplOptno&"'"&VbCRLF				'������ outmallOptCode�� 001,002,003 �̷��� �������� ���� �Ŀ� 1,2,3�̷��� ����
						    strSql = strSql & " and mallid='homeplus'"&VbCRLF
							strSql = strSql & " END ELSE "
							strSql = strSql & " BEGIN"& VbCRLF
							strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
					        strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate)"
					        strSql = strSql & " VALUES ('"&iitemid&"', '"&regedOpt10x10OptNo&"', 'homeplus', '"&hplOptno&"', '"&html2DB(regedOpt10x10OptNm)&"', '"&Chkiif(hplOptStatus="true", "Y", "N")&"', 'Y', '"&StockQty&"', '', getdate())"
							strSql = strSql & " END "
						    dbget.Execute strSql, AssignedRow
							actCnt = actCnt+AssignedRow
'							rw "retCode : " &retCode
'							rw "resultmsg : "&resultmsg
'							rw "regedItemStatus : "&regedItemStatus
'							rw "hplOptStatus : "&hplOptStatus
'							rw "hplOptno : "&hplOptno
'							rw "regedOpt10x10OptNo : "&regedOpt10x10OptNo
'							rw "regedOpt10x10OptNm : "&regedOpt10x10OptNm
'							rw "----------------------------"
						Next
					Set oneProdInfo = nothing
'					2.regedItemStatus�� ���� ��ǰ�Ǹ� ���� ���� / regrdOptcnt�� ����
					If (actCnt > 0) Then
						strSql = " update R"   &VbCRLF
						strSql = strSql & " set regedOptCnt=isNULL(T.regedOptCnt,0)"   &VbCRLF
'						strSql = strSql & " ,homeplusSellYn = '"&Chkiif(regedItemStatus="true", "Y", "N")&"'"   &VbCRLF
						strSql = strSql & " from db_etcmall.dbo.tbl_homeplus_regItem R"   &VbCRLF
						strSql = strSql & " 	Join ("   &VbCRLF
						strSql = strSql & " 		select R.itemid,count(*) as CNT "
						strSql = strSql & " 		, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
						strSql = strSql & "        from db_etcmall.dbo.tbl_homeplus_regItem R"   &VbCRLF
						strSql = strSql & " 			Join db_item.dbo.tbl_OutMall_regedoption Ro"   &VbCRLF
						strSql = strSql & " 			on R.itemid=Ro.itemid"   &VbCRLF
						strSql = strSql & " 			and Ro.mallid='homeplus'"   &VbCRLF
						strSql = strSql & "             and Ro.itemid="&iitemid&VbCRLF
						strSql = strSql & " 		group by R.itemid"   &VbCRLF
						strSql = strSql & " 	) T on R.itemid=T.itemid"   &VbCRLF
						dbget.Execute strSql
					End If
		
					strSql = ""
					strSql = strSql & " SELECT count(*) as cnt FROM db_item.dbo.tbl_OutMall_regedoption where itemid='"&iitemid&"' and outmallSellyn = 'Y' and mallid = 'homeplus' "
					rsget.Open strSql, dbget
					If rsget("cnt") = 0 Then
						strSql = ""
						strSql = strSql & " UPDATE oP "
					    strSql = strSql & " SET homeplusSellYn ='N'"&VbCRLF
					    strSql = strSql & " FROM db_etcmall.dbo.tbl_homeplus_regitem oP"&VbCRLF
					    strSql = strSql & " WHERE itemid="&iitemid&VbCRLF
						dbget.Execute strSql
					End If
					rsget.Close
					iErrStr =  "OK||"&iitemid&"||����(��ǰ��ȸ)"
				Else						'����(E)
				    iErrStr = "ERR||"&iitemid&"||"&resultmsg
				End If
				Set xmlDOM2 = nothing
			Else
				iErrStr = "ERR||"&iitemid&"||Homeplus ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-CHKSTAT-001]"
			End If			
		Else
			iErrStr = "ERR||"&iitemid&"||Homeplus �α��� ����[ERR-CHKSTAT-001]"
		End If
		Set xmlDOM = nothing
	End If
	Set objXML = nothing
	On Error Goto 0
End Function

'��ǰ ��ȸ
Function fnHomeplusOneItemStatView(iitemid, iHomeplusGoodNo, byRef iErrStr, istrParam, mode)
	Dim objXML, xmlDOM, xmlDOM2, strSql
	Dim retCode, resultmsg, regedItemStatus, actCnt
	Dim oneProdInfo, SubNodes, AssignedRow, StockQty
	Dim hplOptStatus, hplOptno, regedOpt10x10OptNo, regedOpt10x10OptNm
	Dim Soptioncnt, Slimityn, Slimitno, Slimitsold 
	Dim xmlStr : xmlStr = getXMLString("login")
    On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#login"
		objXML.setTimeouts 5000,90000,90000,90000
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.ValidateOnParse= True
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then	'�α��� �����̶��
			objXML.open "post", "" & homeplusAPIURL & "", False
			objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
			objXML.setRequestHeader "Content-Length", Len(xmlStr)
			objXML.setRequestHeader "SOAPAction", strInterface & "#" &mode
			objXML.send(istrParam)
			If objXML.Status = "200" Then
				Set xmlDOM2 = Server.CreateObject("MSXML.DOMDocument")
					xmlDOM2.async = False
					xmlDOM2.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		'			response.write objXML.ResponseText
		'			response.end
					retCode			= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/searchProductResponse/ns1:searchProductReturn/ns1:code").text
					resultmsg		= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/searchProductResponse/ns1:searchProductReturn/ns1:message").text
					regedItemStatus	= xmlDOM2.selectSingleNode("soapenv:Envelope/soapenv:Body/searchProductResponse/ns1:searchProductReturn/ns1:SALE").text
				If retCode = "E0000" Then	'����(E0000)
					strSql = ""
					strSql = strSql & " SELECT TOP 1 i.optioncnt, i.limityn, i.limitno, i.limitsold "
					strSql = strSql & " FROM db_item.dbo.tbl_item as i "
					strSql = strSql & " JOIN db_etcmall.dbo.tbl_homeplus_regitem as r on i.itemid = r.itemid "
					strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
			        rsget.Open strSql, dbget
					If Not(rsget.EOF or rsget.BOF) Then
						Soptioncnt	= rsget("optioncnt")
						Slimityn	= rsget("limityn")
						Slimitno	= rsget("limitno")
						Slimitsold	= rsget("limitsold")
					End If
					rsget.Close

					Set oneProdInfo = xmlDOM2.SelectNodes("soapenv:Envelope/soapenv:Body/searchProductResponse/ns1:searchProductReturn/ns1:ITEMRESULT/ns1:ITEMRESULT")
						For Each SubNodes In oneProdInfo
							hplOptStatus			= SubNodes.SelectSingleNode("ns1:SALE").text
							hplOptno				= SubNodes.SelectSingleNode("ns1:i_ITEMNO").text
							regedOpt10x10OptNo		= SubNodes.SelectSingleNode("ns1:s_ITEMNO").text
							regedOpt10x10OptNm		= SubNodes.SelectSingleNode("ns1:s_OPTION_NAME").text
		
							If Soptioncnt = 0 Then		'��ǰ�̶��
								If Slimityn = "Y" Then
									StockQty = Slimitno - Slimitsold - 5
								Else
									StockQty = 999
								End If
							Else												'�ɼ��̶��
								If Slimityn = "Y" Then
									strSql = ""
									strSql = strSql & " SELECT CASE WHEN (optlimitno - optlimitsold) <= 5 Then '0' Else (optlimitno - optlimitsold - 5) End as StockQty "
									strSql = strSql & " FROM db_item.dbo.tbl_item_option  "
									strSql = strSql & " WHERE itemid='"&iitemid&"' and itemoption = '"&regedOpt10x10OptNo&"' "
							        rsget.Open strSql, dbget
									If Not(rsget.EOF or rsget.BOF) Then
										StockQty = rsget("StockQty")
									Else
										StockQty = 0
									End If
									rsget.Close
								Else
									StockQty = 999
								End If
							End If
							'1.������ �μ�Ʈ ������ ������Ʈ
							strSql = ""
							strSql = strSql & " IF Exists(SELECT * FROM db_item.dbo.tbl_OutMall_regedoption where itemid='"&iitemid&"' and itemoption = '"&regedOpt10x10OptNo&"' and outmallOptCode = '"&hplOptno&"' and mallid = 'homeplus') "
							strSql = strSql & " BEGIN"& VbCRLF
							strSql = strSql & " UPDATE oP "
						    strSql = strSql & " SET outmallOptName='"&html2DB(regedOpt10x10OptNm)&"'"&VbCRLF
							strSql = strSql & " ,outmallOptCode='"&hplOptno&"'"&VbCRLF
						    strSql = strSql & " ,lastupdate=getdate()"&VbCRLF
						    strSql = strSql & " ,outMallSellyn='"&Chkiif(hplOptStatus="true", "Y", "N")&"'"&VbCRLF
						    strSql = strSql & " ,outmalllimityn='Y'"&VbCRLF
						    strSql = strSql & " ,outMallLimitNo="&StockQty&VbCRLF
						    strSql = strSql & " ,checkdate=getdate()"&VbCRLF
						    strSql = strSql & " FROM db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
						    strSql = strSql & " WHERE itemid="&iitemid&VbCRLF
						    strSql = strSql & " and convert(int, outmallOptCode)='"&hplOptno&"'"&VbCRLF				'������ outmallOptCode�� 001,002,003 �̷��� �������� ���� �Ŀ� 1,2,3�̷��� ����
						    strSql = strSql & " and mallid='homeplus'"&VbCRLF
							strSql = strSql & " END ELSE "
							strSql = strSql & " BEGIN"& VbCRLF
							strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
					        strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate)"
					        strSql = strSql & " VALUES ('"&iitemid&"', '"&regedOpt10x10OptNo&"', 'homeplus', '"&hplOptno&"', '"&html2DB(regedOpt10x10OptNm)&"', '"&Chkiif(hplOptStatus="true", "Y", "N")&"', 'Y', '"&StockQty&"', '', getdate())"
							strSql = strSql & " END "
						    dbget.Execute strSql, AssignedRow
							actCnt = actCnt+AssignedRow
'							rw "retCode : " &retCode
'							rw "resultmsg : "&resultmsg
'							rw "regedItemStatus : "&regedItemStatus
'							rw "hplOptStatus : "&hplOptStatus
'							rw "hplOptno : "&hplOptno
'							rw "regedOpt10x10OptNo : "&regedOpt10x10OptNo
'							rw "regedOpt10x10OptNm : "&regedOpt10x10OptNm
'							rw "----------------------------"
						Next
					Set oneProdInfo = nothing
'					2.regedItemStatus�� ���� ��ǰ�Ǹ� ���� ���� / regrdOptcnt�� ����
					If (actCnt > 0) Then
						strSql = " update R"   &VbCRLF
						strSql = strSql & " set regedOptCnt=isNULL(T.regedOptCnt,0)"   &VbCRLF
						strSql = strSql & " ,homeplusSellYn = '"&Chkiif(regedItemStatus="true", "Y", "N")&"'"   &VbCRLF
						strSql = strSql & " from db_etcmall.dbo.tbl_homeplus_regItem R"   &VbCRLF
						strSql = strSql & " 	Join ("   &VbCRLF
						strSql = strSql & " 		select R.itemid,count(*) as CNT "
						strSql = strSql & " 		, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
						strSql = strSql & "        from db_etcmall.dbo.tbl_homeplus_regItem R"   &VbCRLF
						strSql = strSql & " 			Join db_item.dbo.tbl_OutMall_regedoption Ro"   &VbCRLF
						strSql = strSql & " 			on R.itemid=Ro.itemid"   &VbCRLF
						strSql = strSql & " 			and Ro.mallid='homeplus'"   &VbCRLF
						strSql = strSql & "             and Ro.itemid="&iitemid&VbCRLF
						strSql = strSql & " 		group by R.itemid"   &VbCRLF
						strSql = strSql & " 	) T on R.itemid=T.itemid"   &VbCRLF
						dbget.Execute strSql
					End If
		
					strSql = ""
					strSql = strSql & " SELECT count(*) as cnt FROM db_item.dbo.tbl_OutMall_regedoption where itemid='"&iitemid&"' and outmallSellyn = 'Y' and mallid = 'homeplus' "
					rsget.Open strSql, dbget
					If rsget("cnt") = 0 Then
						strSql = ""
						strSql = strSql & " UPDATE oP "
					    strSql = strSql & " SET homeplusSellYn ='N'"&VbCRLF
					    strSql = strSql & " FROM db_etcmall.dbo.tbl_homeplus_regitem oP"&VbCRLF
					    strSql = strSql & " WHERE itemid="&iitemid&VbCRLF
						dbget.Execute strSql
					End If
					rsget.Close
					iErrStr =  "OK||"&iitemid&"||����(��ǰ��ȸ)"
				Else						'����(E)
				    iErrStr = "ERR||"&iitemid&"||"&resultmsg
				End If
				Set xmlDOM2 = nothing
			Else
				iErrStr = "ERR||"&iitemid&"||Homeplus ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-CHKSTAT-001]"
			End If			
		Else
			iErrStr = "ERR||"&iitemid&"||Homeplus �α��� ����[ERR-CHKSTAT-001]"
		End If
		Set xmlDOM = nothing
	End If
	Set objXML = nothing
	On Error Goto 0
End Function
'############################################## ���� �����ϴ� API �Լ� ���� �� ############################################

'################################################# �� ��� �� �Ķ���� ���� ###############################################
Function getHomplusSellynParameter(iHomplusGoodno, ichgSellYn)
	Dim strRst, ckSellyn
	If (ichgSellYn = "N") Then
		ckSellyn = False
	Else
		ckSellyn = True
	End If

	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
	strRst = strRst & "	<SOAP-ENV:Body>"
	strRst = strRst & "		<m:setProductStatus xmlns:m=""" & strInterface & """>"
	strRst = strRst & "			<I_STYLENO>"&iHomplusGoodno&"</I_STYLENO>"
	strRst = strRst & "			<B_Status>"&ckSellyn&"</B_Status>"
	strRst = strRst & "		</m:setProductStatus>"
	strRst = strRst & "	</SOAP-ENV:Body>"
	strRst = strRst & "</SOAP-ENV:Envelope>"
	getHomplusSellynParameter = strRst
End Function

Function getHomplusStatChkParameter(iitemid)
	Dim strRst
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
	strRst = strRst & "	<SOAP-ENV:Body>"
	strRst = strRst & "		<m:searchProduct xmlns:m=""" & strInterface & """>"
	strRst = strRst & "			<PRODUCT_CODE>"&iitemid&"</PRODUCT_CODE>"
	strRst = strRst & "		</m:searchProduct>"
	strRst = strRst & "	</SOAP-ENV:Body>"
	strRst = strRst & "</SOAP-ENV:Envelope>"
	getHomplusStatChkParameter = strRst
End Function
%>