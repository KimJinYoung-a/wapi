<%
'��ǰ �շ�
Public Function fnMy11stItemReg(iitemid, istrParam, iOrgprice, irateprice, iregedOptcnt, iMultiplerate, iExchangeRate, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, resultCode, productNo
	Dim iMessage, AssignedRow
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", my11stAPIURL&"/prodservices/product", false
		objXML.setRequestHeader "Content-Type", "application/xml; charset=utf8"
		objXML.setRequestHeader "openapikey", apiKEY
		objXML.send(istrParam)
		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
				resultCode = xmlDOM.getElementsByTagName("resultCode").item(0).text
				If resultCode = "200" Then
					productNo = xmlDOM.getElementsByTagName("productNo").item(0).text
					strSql = ""
					strSql = strSql & " UPDATE R " & vbcrlf
					strSql = strSql & " SET my11stGoodNo = '"&productNo&"' " & vbcrlf
					strSql = strSql & " , my11stLastUpdate = getdate() " & vbcrlf
					strSql = strSql & " , regOrgprice = "&iOrgprice&" " & vbcrlf
					strSql = strSql & " , my11stPrice = '" & irateprice & "'" & vbcrlf
					strSql = strSql & " , accFailCnt = 0 " & vbcrlf
					strSql = strSql & " , regedOptCnt = '"&iregedOptcnt&"' " & vbcrlf
					strSql = strSql & " , my11stRegdate = getdate() " & vbcrlf
					strSql = strSql & " , my11stStatCd = 7 " & vbcrlf
					strSql = strSql & " , multiplerate = '"&iMultiplerate&"' " & vbcrlf
					strSql = strSql & " , exchangeRate = '"&iExchangeRate&"' " & vbcrlf
					strSql = strSql & "	FROM db_etcmall.[dbo].[tbl_my11st_regItem] R " & vbcrlf
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute strSql
					iErrStr = "OK||"&iitemid&"||����(��ǰ���)"
				Else
					iMessage = xmlDOM.getElementsByTagName("message")(0).Text
					iErrStr = "ERR||"&iitemid&"||"&iMessage
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||11���� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'��ǰ ����
Public Function fnMy11stItemEdit(iitemid, imy11stGoodno, istrParam, iOrgprice, iExchangeRate, iMultiplerate, irateprice, iregedOptcnt, iItemname, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, resultCode, productNo
	Dim iMessage, AssignedRow
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", my11stAPIURL&"/prodservices/product/"&imy11stGoodno, false
		objXML.setRequestHeader "Content-Type", "application/xml; charset=utf8"
		objXML.setRequestHeader "openapikey", apiKEY
		objXML.send(istrParam)
		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				resultCode = xmlDOM.getElementsByTagName("resultCode").item(0).text
				If resultCode = "200" Then
					productNo = xmlDOM.getElementsByTagName("productNo").item(0).text
					strSql = ""
					strSql = strSql & " UPDATE R " & vbcrlf
					strSql = strSql & " SET my11stLastUpdate = getdate() " & vbcrlf
					strSql = strSql & " , regOrgprice = "&iOrgprice&" " & vbcrlf
					strSql = strSql & " , multiplerate = '"&iMultiplerate&"' " & vbcrlf
					strSql = strSql & " , exchangeRate = '"&iExchangeRate&"' " & vbcrlf
					strSql = strSql & " , my11stPrice = '" & irateprice & "'" & vbcrlf
					strSql = strSql & " , accFailCnt = 0 " & vbcrlf
					strSql = strSql & " , regedOptCnt = '"&iregedOptcnt&"' " & vbcrlf
					strSql = strSql & " , regitemname = '"&html2db(iItemname)&"' " & vbcrlf
					strSql = strSql & "	FROM db_etcmall.[dbo].[tbl_my11st_regItem] R " & vbcrlf
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute strSql
					iErrStr = "OK||"&iitemid&"||����(��ǰ����)"
				Else
					iMessage = xmlDOM.getElementsByTagName("message")(0).Text
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(��ǰ����)"
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||11���� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDIT]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'�ǸŻ��� ���� N
Public Function fnMy11stSoldOut(iitemid, i11stGoodno, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, resultCode, productNo
	Dim iMessage, AssignedRow
	On Error Resume Next
	strSql = ""
	strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_my11st_regItem] " & VbCRLF
	strSql = strSql & " SET my11stSellYn = 'N'" & VbCRLF
	strSql = strSql & " ,my11stLastUpdate = getdate()" & VbCRLF
	strSql = strSql & " ,accFailCNT=0" & VbCRLF
	strSql = strSql & " WHERE itemid = "&iitemid
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "PUT", my11stAPIURL&"/prodstatservice/stat/stopdisplay/"&i11stGoodno, false
		objXML.setRequestHeader "Content-Type", "application/xml; charset=utf8"
		objXML.setRequestHeader "openapikey", apiKEY
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				resultCode = xmlDOM.getElementsByTagName("resultCode").item(0).text
				If resultCode = "200" Then
					dbget.Execute(strSql)
					iErrStr = "OK||"&iitemid&"||�Ǹ�����(Hidden)"
				Else
					iMessage = xmlDOM.getElementsByTagName("message")(0).Text
					If InStr(iMessage, "Only items under 'On Sale' or 'Before Listing' status can be put under 'Listing on Hold.'") Then
						dbget.Execute(strSql)
						iErrStr = "OK||"&iitemid&"||�ǸŻ��� ����(SKIP_Hidden)"
					Else
						iErrStr = "ERR||"&iitemid&"||"&iMessage&" �Ǹ�����(Hidden)"
					End If
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||11���� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-SOLDOUT]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'�ǸŻ��� ���� Y
Public Function fnMy11stOnSale(iitemid, i11stGoodno, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, resultCode, productNo
	Dim iMessage, AssignedRow
	On Error Resume Next
	strSql = ""
	strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_my11st_regItem] " & VbCRLF
	strSql = strSql & " SET my11stSellYn = 'Y'" & VbCRLF
	strSql = strSql & " ,my11stLastUpdate = getdate()" & VbCRLF
	strSql = strSql & " ,accFailCNT=0" & VbCRLF
	strSql = strSql & " WHERE itemid = "&iitemid

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "PUT", my11stAPIURL&"/prodstatservice/stat/restartdisplay/"&i11stGoodno, false
		objXML.setRequestHeader "Content-Type", "application/xml; charset=utf8"
		objXML.setRequestHeader "openapikey", apiKEY
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				resultCode = xmlDOM.getElementsByTagName("resultCode").item(0).text
				If resultCode = "200" Then
					dbget.Execute(strSql)
					iErrStr = "OK||"&iitemid&"||�Ǹ�(Active)"
				Else
					iMessage = xmlDOM.getElementsByTagName("message")(0).Text
					If InStr(iMessage, "Only products under 'Listing on Hold' can be released from the state") Then
						dbget.Execute(strSql)
						iErrStr = "OK||"&iitemid&"||�ǸŻ��� ����(SKIP_Active)"
					Else
						iErrStr = "ERR||"&iitemid&"||"&iMessage&" �Ǹ�(Active)"
					End If
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||11���� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-ONSALE]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'�Ǹ� ���� ����
Public Function fnMy11stPrice(iitemid, i11stGoodno, iOrgprice, iExchangeRate, iMultiplerate, irateprice, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, resultCode, productNo
	Dim iMessage, AssignedRow
	On Error Resume Next
    strSql = ""
	strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_my11st_regItem]  " & VbCRLF
	strSql = strSql & "	SET my11stLastUpdate = getdate() " & VbCRLF
	strSql = strSql & "	, my11stPrice = " & irateprice & VbCRLF
	strSql = strSql & "	, regOrgprice = " & iOrgprice & VbCRLF
	strSql = strSql & " , multiplerate = '"&iMultiplerate&"' " & vbcrlf
	strSql = strSql & " , exchangeRate = '"&iExchangeRate&"' " & vbcrlf
	strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	strSql = strSql & " WHERE itemid='" & iitemid & "'"
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", my11stAPIURL&"/prodservices/product/price/"&i11stGoodno&"/"&irateprice, false
		objXML.setRequestHeader "Content-Type", "application/xml; charset=utf8"
		objXML.setRequestHeader "openapikey", apiKEY
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				resultCode = xmlDOM.getElementsByTagName("resultCode").item(0).text
				If resultCode = "200" Then
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||��������(��ǰ����)"
				Else
					iMessage = xmlDOM.getElementsByTagName("message")(0).Text
					If InStr(iMessage, "The price is not different from the original price") Then
						dbget.Execute(strSql)
						iErrStr = "OK||"&iitemid&"||���� ����(SKIP_PRICE)"
					Else
						iErrStr = "ERR||"&iitemid&"||"&iMessage&" (PRICE)"	
					End If
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||11���� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRICE]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'�ɼ� ����
Public Function fnMy11stOptEdit(iitemid, iMy11stGoodNo, istrParam, iregedOptcnt, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, resultCode, productNo
	Dim iMessage, AssignedRow
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", my11stAPIURL&"/prodservices/updateProductOption/"&iMy11stGoodNo, false
		objXML.setRequestHeader "Content-Type", "application/xml; charset=utf8"
		objXML.setRequestHeader "openapikey", apiKEY
		objXML.send(istrParam)
		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				resultCode = xmlDOM.getElementsByTagName("resultCode").item(0).text
				If resultCode = "200" Then
				    strSql = ""
	    			strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_my11st_regItem]  " & VbCRLF
	    			strSql = strSql & "	SET my11stLastUpdate = getdate() " & VbCRLF
	    			strSql = strSql & "	, regedOptcnt = '" & iregedOptcnt &"'" & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " WHERE itemid='" & iitemid & "'"
	    			dbget.Execute(strSql)
					iErrStr =  "OK||"&iitemid&"||��������(EditOPT)"
				Else
					iMessage = xmlDOM.getElementsByTagName("message")(0).Text
					iErrStr = "ERR||"&iitemid&"||"&iMessage&" (EditOPT)"
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||11���� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EditOPT]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'�ɼ� ��ȸ
Public Function fnMy11stOptView(iitemid, i11stGoodno, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, resultCode, productNo, prdStckNo, prdStckStatCd, stckQty, mixDtlOptNm, addPrc
	Dim iMessage, AssignedRow, Nodes, SubNodes, i
	i = 0
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", my11stAPIURL&"/prodmarketservice/prodmarket/stck/"&i11stGoodno, false
		objXML.setRequestHeader "Content-Type", "application/xml; charset=utf8"
		objXML.setRequestHeader "openapikey", apiKEY
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
'response.end
				productNo = xmlDOM.getElementsByTagName("prdNo").item(0).text			'���� �������ΰ� ����
				If productNo <> "" Then
					strSQL = ""
					strSQL = strSQL & " DELETE FROM db_item.dbo.tbl_Outmall_regedoption WHERE itemid = '"&iitemid&"' and mallid = '"&CMALLNAME&"' "
					dbget.Execute strSQL
					Set Nodes = xmlDOM.getElementsByTagName("ProductStock")
						For each SubNodes in Nodes
							prdStckNo		= SubNodes.getElementsByTagName("prdStckNo")(0).Text 		'11���� �ɼ��ڵ�
							prdStckStatCd	= SubNodes.getElementsByTagName("prdStckStatCd")(0).Text	'01 : Use, 02 : SoldOut
							stckQty			= SubNodes.getElementsByTagName("stckQty")(0).Text			'���� ����
							mixDtlOptNm		= SubNodes.getElementsByTagName("mixDtlOptNm")(0).Text		'��� �ɼǸ�
							addPrc			= SubNodes.getElementsByTagName("addPrc")(0).Text			'�ɼ� �ݾ�

							Select Case SubNodes.getElementsByTagName("prdStckStatCd")(0).Text
								Case "01"			prdStckStatCd = "Y"
								Case "02"			prdStckStatCd = "N"
							End Select
							
							If mixDtlOptNm <> "" Then
								strSQL = ""
								strSQL = strSQL & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastUpdate, checkdate) "
								strSQL = strSQL & " SELECT TOP 1 '"&iitemid&"', itemoption, '11stmy', '"&prdStckNo&"', optionname, '"&prdStckStatCd&"', 'Y', '"&stckQty&"', getdate(), getdate() "
								strSQL = strSQL & " FROM [db_item].[dbo].[tbl_item_multiLang_option] "
								strSQL = strSQL & " WHERE itemid = '"&iitemid&"' "
								strSQL = strSQL & " and countryCd ='EN' "
								strSQL = strSQL & " and optionname ='"&mixDtlOptNm&"' "
								dbget.Execute strSQL
								i = i + 1
							End If
						Next
					Set Nodes = nothing

					If i > 0 Then
						strSQL = ""
						strSQL = strSQL & " UPDATE R"   &VbCRLF
						strSQL = strSQL & " SET regedOptCnt= "&i&""   &VbCRLF
						strSQL = strSQL & " ,lastOptConfirmdate = getdate() "   &VbCRLF
						strSQL = strSQL & " FROM db_etcmall.[dbo].[tbl_my11st_regItem] R"   &VbCRLF
						strSQL = strSQL & " WHERE itemid = '"&iitemid&"' "
						dbget.Execute strSQL
					End If
					iErrStr =  "OK||"&iitemid&"||����(VIEWOPT)"
				Else
					iMessage = xmlDOM.getElementsByTagName("message")(0).Text
					iErrStr = "ERR||"&iitemid&"||"&iMessage&" (VIEWOPT)"
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||11���� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-VIEWOPT]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'�Ǹ� ��ȸ
Public Function fnMy11stView(iitemid, i11stGoodno, byRef iErrStr)
	Dim objXML, xmlDOM, strSql, resultCode, productNo
	Dim iMessage, AssignedRow
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", my11stAPIURL&"/prodservices/product/details/"&i11stGoodno, false
		objXML.setRequestHeader "Content-Type", "application/xml; charset=utf8"
		objXML.setRequestHeader "openapikey", apiKEY
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
response.write BinaryToText(objXML.ResponseBody, "euc-kr")
response.end

				resultCode = xmlDOM.getElementsByTagName("resultCode").item(0).text
				If resultCode = "200" Then
'				    '// ��ǰ�������� ����
'				    strSql = ""
'	    			strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_my11st_regItem]  " & VbCRLF
'	    			strSql = strSql & "	SET my11stLastUpdate = getdate() " & VbCRLF
'	    			strSql = strSql & "	, my11stPrice = " & imustprice & VbCRLF
'	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
'	    			strSql = strSql & " Where itemid='" & iitemid & "'"
'	    			dbget.Execute(strSql)
'					iErrStr =  "OK||"&iitemid&"||��������(��ǰ����)"
				Else
					iMessage = xmlDOM.getElementsByTagName("message")(0).Text
					iErrStr = "ERR||"&iitemid&"||"&iMessage&" (VIEW)"
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||11���� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-VIEW]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'�����ڵ�
Public Function getCommCode(iccd)
	Dim strRst
	Dim objXML, xmlDOM, strSql, Nodes, SubNodes
	Dim retCode, goodsCd, iMessage, AssignedRow
	Dim depth, dispNm, CateKey, parentCateKey
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", my11stAPIURL&"/cateservice/category", false
		objXML.setRequestHeader "Content-Type", "application/xml; charset=utf8"
		objXML.setRequestHeader "openapikey", apiKEY
		objXML.send()
	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
'			response.write BinaryToText(objXML.ResponseBody, "euc-kr")

			Set Nodes = xmlDOM.getElementsByTagName("ns2:category")
				For each SubNodes in Nodes
					depth			= SubNodes.getElementsByTagName("depth")(0).Text 
					dispNm			= SubNodes.getElementsByTagName("dispNm")(0).Text
					CateKey			= SubNodes.getElementsByTagName("dispNo")(0).Text
					parentCateKey	= SubNodes.getElementsByTagName("parentDispNo")(0).Text

					strSql = ""
					strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_my11st_TmpCategory (depth, dispNm, CateKey, parentCateKey) VALUES "
					strSql = strSql & " ('"&depth&"', '"&dispNm&"', '"&CateKey&"', '"&parentCateKey&"') "
					dbget.Execute(strSql)
				Next
			Set Nodes = nothing
		Set xmlDOM = nothing
	End If
End Function
%>