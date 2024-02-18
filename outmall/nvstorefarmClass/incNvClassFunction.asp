<%
Public Function getsecretKey(iaccessLicense, iTimestamp, isignature, iserv, ioper)
	Dim cryptoLib, oLicense, osecretKey, otimeStamp, osignature
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
		If (application("Svr_Info") = "Dev") Then
'			iaccessLicense = "01000100004b035a25d67f991849cad1c7042b8da528d13e9ddce6878f2e43ac88080e0a5e" 'AccessLicense Key �Է�, PDF��������
'			osecretKey = "AQABAAAWPWagCrPjFQnFEtxs5j+oyZFwuzomdNq0XZSricPuMw=="  'SecreKey �Է�, PDF��������
			iaccessLicense = "010001000019133c715650b9c85b820961612f2b90b431ddd8654b42c097c4df1a43d0be09" 'AccessLicense Key �Է�, PDF��������
			osecretKey = "AQABAADX6Hz/wORFJS5pSIy4KQXkH83gC9G1aXChxBjcnUMqWw=="  'SecreKey �Է�, PDF��������
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		Else
			iaccessLicense = "010001000019133c715650b9c85b820961612f2b90b431ddd8654b42c097c4df1a43d0be09" 'AccessLicense Key �Է�, PDF��������
			osecretKey = "AQABAADX6Hz/wORFJS5pSIy4KQXkH83gC9G1aXChxBjcnUMqWw=="  'SecreKey �Է�, PDF��������
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		End If
	Set cryptoLib = nothing
End Function

'�̹��� ���ε�
Public Function fnNvClassImageReg(iitemid, strParam, byRef iErrStr, ichgImageNm, iservice, ioperation)
	Dim objXML, xmlDOM, strSql, iMessage, nvstorefarmURL, SubNodes, ResponseType, imglist
	Dim myURL, yourURL
	If (application("Svr_Info") = "Dev") Then
		nvstorefarmURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstorefarmURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstorefarmURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iservice & "#" & ioperation
		objXML.send(strParam)

 		If objXML.Status = 200 Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(objXML.responseText)
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
				If ResponseType = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " DELETE FROM " & vbcrlf
					strSql = strSql & " db_etcmall.[dbo].[tbl_nvstorefarmclass_Image] " & vbcrlf
					strSql = strSql & " WHERE itemid = '"&iitemid&"' "
					dbget.Execute strSql

					Set imglist = xmlDOM.getElementsByTagName("n:Image")
					For Each SubNodes in imglist
						myURL	= SubNodes.SelectSingleNode("n:Source").text
						yourURL	= SubNodes.SelectSingleNode("n:URL").text

						If InStr(myURL, "/basic/") Then
							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstorefarmclass_Image] (itemid, imgType, tenURL, storefarmURL) VALUES " & vbcrlf
							strSql = strSql & " ('"&iitemid&"', '1', '"&myURL&"', '"&yourURL&"') " & vbcrlf
							dbget.Execute strSql
						Else
							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstorefarmclass_Image] (itemid, imgType, tenURL, storefarmURL) VALUES " & vbcrlf
							strSql = strSql & " ('"&iitemid&"', '2', '"&myURL&"', '"&yourURL&"') " & vbcrlf
							dbget.Execute strSql
						End If
					Next
					Set imglist = nothing
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] SET "
					strSql = strSql & " APIaddImg = 'Y' "
					strSql = strSql & " ,regimageName = '"&ichgImageNm&"'"& VbCrlf
					strSql = strSql & " WHERE itemid = '"&iitemid&"' "
					dbget.Execute strSql
					iErrStr = "OK||"&iitemid&"||�̹��� ���ε� ����"
				Else
					iMessage = xmlDOM.getElementsByTagName("n:Message")(0).Text
					If InStr(iMessage, "���ռ� ����") Then
						iMessage = xmlDOM.getElementsByTagName("n:Detail")(0).Text
					End If
					iErrStr = "ERR||"&iitemid&"||"&iMessage
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||������� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-IMAGE]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'��ǰ ���
Public Function fnNvClassItemReg(iitemid, strParam, byRef iErrStr, iSellCash, invClassSellyn, iitemname, iimageNm, iservice, ioperation)
	Dim objXML, xmlDOM, strSql, iMessage, nvstorefarmURL, SubNodes, ResponseType, imglist
	Dim ProductId
	If (application("Svr_Info") = "Dev") Then
		nvstorefarmURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstorefarmURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstorefarmURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iservice & "#" & ioperation
		objXML.send(strParam)
 		If objXML.Status = 200 Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(objXML.responseText)
If iitemid = "2525634" Then
	response.write objXML.responseText
End If
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
'				If ResponseType = "SUCCESS" Then
					ProductId = xmlDOM.getElementsByTagName("ProductId").item(0).text
				If ResponseType = "SUCCESS" AND ProductId <> "" Then
					strSql = strSql & " UPDATE R " & vbcrlf
					strSql = strSql & " SET nvClassGoodNo = '"&ProductId&"' " & vbcrlf
					strSql = strSql & " , nvClassLastUpdate = getdate() " & vbcrlf
					strSql = strSql & " , nvClassPrice = " & iSellCash & vbcrlf
					strSql = strSql & " , accFailCnt = 0 " & vbcrlf
					strSql = strSql & " , nvClassRegdate = getdate() " & vbcrlf
					strSql = strSql & " , nvClassStatCd = 7 " & vbcrlf
					strSql = strSql & "	FROM db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] R " & vbcrlf
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute strSql
					iErrStr = "OK||"&iitemid&"||����(��ǰ���)"
				Else
					iMessage = xmlDOM.getElementsByTagName("n:Message")(0).Text
					If InStr(iMessage, "���ռ� ����") Then
						iMessage = xmlDOM.getElementsByTagName("n:Detail")(0).Text
					End If
					iErrStr = "ERR||"&iitemid&"||"&iMessage
'response.write xmlDOM.getElementsByTagName("n:Detail")(0).Text
'response.end
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||������� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'��ǰ ����
Public Function fnNvClassItemEDIT(iitemid, strParam, byRef iErrStr, iSellCash, iitemname, ichgImageNm, iservice, ioperation)
	Dim objXML, xmlDOM, strSql, iMessage, nvstorefarmURL, SubNodes, ResponseType, imglist
	Dim ProductId
	If (application("Svr_Info") = "Dev") Then
		nvstorefarmURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstorefarmURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstorefarmURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iservice & "#" & ioperation
		objXML.send(strParam)
 		If objXML.Status = 200 Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(objXML.responseText)
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
				If ResponseType = "SUCCESS" Then
					ProductId = xmlDOM.getElementsByTagName("ProductId").item(0).text
					strSql = ""
					strSql = strSql & " UPDATE R " & vbcrlf
					strSql = strSql & " SET nvClassLastUpdate = getdate() " & vbcrlf
					strSql = strSql & " , nvClassPrice = " & iSellCash & vbcrlf
					strSql = strSql & " , accFailCnt = 0 " & vbcrlf
					If (ichgImageNm <> "N") Then
						strSql = strSql & " ,regimageName='"&ichgImageNm&"'"& VbCrlf
					End If
					strSql = strSql & " , regitemname = '"&html2db(iitemname)&"'" & vbcrlf
					strSql = strSql & "	FROM db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] R " & vbcrlf
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute strSql
					iErrStr = "OK||"&iitemid&"||����(��ǰ����)"
				Else
					iMessage = xmlDOM.getElementsByTagName("n:Message")(0).Text
					If InStr(iMessage, "���ռ� ����") Then
						iMessage = xmlDOM.getElementsByTagName("n:Detail")(0).Text
					End If
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(��ǰ����)"
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||������� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-ITEMEDIT]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'�ɼ� ���
Public Function fnNvClassOptionReg(iitemid, strParam, byRef iErrStr, iservice, ioperation)
	Dim objXML, xmlDOM, strSql, iMessage, nvstorefarmURL, SubNodes, ResponseType, imglist
	Dim myURL, yourURL, statusType, nvRegitemname, MasterPrice, ProductId
	If (application("Svr_Info") = "Dev") Then
		nvstorefarmURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstorefarmURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstorefarmURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iservice & "#" & ioperation
		objXML.send(strParam)
 		If objXML.Status = 200 Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(objXML.responseText)
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text

				If ResponseType = "SUCCESS" Then
					iErrStr = "OK||"&iitemid&"||����(�ɼǼ���)"
				Else
					iMessage = xmlDOM.getElementsByTagName("n:Message")(0).Text
					If InStr(iMessage, "���ռ� ����") OR InStr(iMessage, "id �׸���") Then
						iMessage = xmlDOM.getElementsByTagName("n:Detail")(0).Text
					End If
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(�ɼǼ���)"
'response.write objXML.responseText
'response.end
'response.write xmlDOM.getElementsByTagName("n:Detail")(0).Text
'response.end
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||������� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-OPTION]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'��ǰ ��ȸ
Public Function fnNvClassItemSearch(iitemid, strParam, byRef iErrStr, iservice, ioperation)
	Dim objXML, xmlDOM, strSql, iMessage, nvstorefarmURL, SubNodes, ResponseType, imglist
	Dim myURL, yourURL, statusType, nvRegitemname, MasterPrice, ProductId
	If (application("Svr_Info") = "Dev") Then
		nvstorefarmURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstorefarmURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstorefarmURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iservice & "#" & ioperation
		objXML.send(strParam)
 		If objXML.Status = 200 Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(objXML.responseText)
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
				If ResponseType = "SUCCESS" Then
					Select Case xmlDOM.getElementsByTagName("n:StatusType").item(0).text
						Case "SALE"		statusType		= "Y" 		'�Ǹ�
						Case "SUSP"		statusType		= "N" 		'�Ͻ� ����
						Case "OSTK"		statusType		= "N" 		'ǰ��
					End Select
					nvRegitemname	= xmlDOM.getElementsByTagName("n:Name").item(0).text
					MasterPrice		= xmlDOM.getElementsByTagName("n:SalePrice").item(0).text
					ProductId		= xmlDOM.getElementsByTagName("n:ProductId").item(0).text

					strSQL = ""
					strSQL = strSQL & " UPDATE R" & VbCRLF
					strSQL = strSQL & " SET nvClassPrice = " & MasterPrice & VbCRLF
					strSQL = strSQL & " ,nvClassSellyn='"&statusType&"'" & VbCRLF
					strSQL = strSQL & " ,regitemname='"&html2db(nvRegitemname)&"'" & VbCRLF
					strSQL = strSQL & " ,lastStatCheckDate = getdate()" & VbCRLF
					strSQL = strSQL & " ,nvClassGoodNo = isNULL(R.nvClassGoodNo,'"&ProductId&"')"&VbCRLF
					strSQL = strSQL & " FROM db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] R" & VbCRLF
					strSQL = strSQL & " where R.itemid="&iitemid & VbCRLF
					strSQL = strSQL & " and isNULL(nvClassGoodNo,'') in ('','"&ProductId&"')"&VbCRLF    ''�ߺ���ϵ�CaSE ���
					strSQL = strSQL & " and (isNULL(nvClassPrice,0)<>"&MasterPrice&"" & VbCRLF
					strSQL = strSQL & "     or isNULL(nvClassSellyn,'')<>'"&statusType&"'"& VbCRLF
					strSQL = strSQL & "     or isNULL(regitemname,'')<>'"&html2db(nvRegitemname)&"'"& VbCRLF
					strSQL = strSQL & "     or isNULL(nvClassGoodNo,'')<>'"&ProductId&"'"& VbCRLF
					strSQL = strSQL & " )"
				    dbget.Execute strSQL
					iErrStr =  "OK||"&iitemid&"||����(�ǸŻ�����ȸ)"
				Else
					iMessage = xmlDOM.getElementsByTagName("n:Message")(0).Text
					If InStr(iMessage, "���ռ� ����") Then
						iMessage = xmlDOM.getElementsByTagName("n:Detail")(0).Text
					End If
					iErrStr = "ERR||"&iitemid&"||"&iMessage
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||������� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-ITEMSEARCH]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'�ɼ� ��ȸ
Public Function fnNvClassOptionSearch(iitemid, strParam, byRef iErrStr, iservice, ioperation)
	Dim objXML, xmlDOM, strSQL, iMessage, nvstorefarmURL, SubNodes, ResponseType, imglist
	Dim myURL, yourURL, statusType, nvRegitemname, MasterPrice, ProductId
	Dim Nodes, onvOptId, myOptNo, addprice, saleLmtQty, nvOptval1, nvOptval2, nvOptval3, nvOptval4, nvOptval5, AlloptNm
	If (application("Svr_Info") = "Dev") Then
		nvstorefarmURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstorefarmURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstorefarmURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iservice & "#" & ioperation
		objXML.send(strParam)
 		If objXML.Status = 200 Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(objXML.responseText)
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
'rw replace(objXML.responseText, "?xml", "?AAAAAl")
				If ResponseType = "SUCCESS" Then
					strSQL = ""
					strSQL = strSQL & " DELETE FROM db_item.dbo.tbl_Outmall_regedoption WHERE itemid = '"&iitemid&"' and mallid = '"&CMALLNAME&"' "
					dbget.Execute strSQL
					If xmlDOM.getElementsByTagName("n:Item").length > 0 Then
						Set Nodes = xmlDOM.getElementsByTagName("n:Item")
							For each SubNodes in Nodes
								onvOptId		= SubNodes.getElementsByTagName("n:Id")(0).Text					'���̹� �ɼ��ڵ�
								saleLmtQty		= SubNodes.getElementsByTagName("n:Quantity")(0).Text			'�ɼ� ����
								AlloptNm		= SubNodes.getElementsByTagName("n:Value1")(0).Text				'�ɼǸ�
								nvOptval2		= SubNodes.getElementsByTagName("n:Value2")(0).Text
								nvOptval3		= SubNodes.getElementsByTagName("n:Value3")(0).Text
								nvOptval4		= SubNodes.getElementsByTagName("n:Value4")(0).Text
								nvOptval5		= SubNodes.getElementsByTagName("n:Value5")(0).Text
								If nvOptval2 <> "" Then
									AlloptNm = AlloptNm & ","&nvOptval2
								ElseIf nvOptval3 <> "" Then
									AlloptNm = AlloptNm & ","&nvOptval3
								ElseIf nvOptval4 <> "" Then
									AlloptNm = AlloptNm & ","&nvOptval4
								End If
								addprice		= SubNodes.getElementsByTagName("n:Price")(0).Text				'�߰��ݾ�
								myOptNo			= SubNodes.getElementsByTagName("n:SellerManagerCode")(0).Text	'10x10 �ɼ��ڵ�

								strSQL = ""
								strSQL = strSQL & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
								strSQL = strSQL & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastUpdate, checkdate) "
								strSQL = strSQL & " VALUES "
								strSQL = strSQL & " ('"&iitemid&"'"
								strSQL = strSQL & ",  '"&myOptNo&"'"
								strSQL = strSQL & ", '"&CMALLNAME&"'"
								strSQL = strSQL & ", '"&onvOptId&"'"
								strSQL = strSQL & ", '"&html2db(AlloptNm)&"'"
								strSQL = strSQL & ", 'Y'"
								strSQL = strSQL & ", '"&"Y"&"'"
								strSQL = strSQL & ", '"&saleLmtQty&"'"
								strSQL = strSQL & ", getdate() "
								strSQL = strSQL & ", getdate()) "
								dbget.Execute strSQL
							Next
						Set Nodes = nothing
					End If
					strSQL = ""
					strSQL = strSQL & " UPDATE R"   &VbCRLF
					strSQL = strSQL & " SET regedOptCnt=isNULL(T.regedOptCnt,0)"   &VbCRLF
					strSQL = strSQL & " FROM db_etcmall.dbo.tbl_nvstorefarmclass_regItem R"   &VbCRLF
					strSQL = strSQL & " JOIN ("   &VbCRLF
					strSQL = strSQL & " 	SELECT R.itemid,count(*) as CNT , sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt "   &VbCRLF
					strSQL = strSQL & " 	FROM db_etcmall.dbo.tbl_nvstorefarmclass_regItem R "   &VbCRLF
					strSQL = strSQL & " 	LEFT JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = '"&CMALLNAME&"' "   &VbCRLF
					strSQL = strSQL & " 	WHERE R.itemid ="&itemid &VbCRLF
					strSQL = strSQL & " 	GROUP BY R.itemid "   &VbCRLF
					strSQL = strSQL & " ) T on R.itemid=T.itemid"   &VbCRLF
					dbget.Execute strSQL
					iErrStr =  "OK||"&iitemid&"||����(�ɼ���ȸ)"
				Else
					iMessage = xmlDOM.getElementsByTagName("n:Message")(0).Text
					If InStr(iMessage, "���ռ� ����") Then
						iMessage = xmlDOM.getElementsByTagName("n:Detail")(0).Text
					End If
					iErrStr = "ERR||"&iitemid&"||"&iMessage
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||������� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-OPTIONSEARCH]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'�Ǹ� ���� ����
Public Function fnNvClassSellyn(iitemid, ichgSellYn, strParam, byRef iErrStr, iservice, ioperation)
	Dim objXML, xmlDOM, strSql, iMessage, nvstorefarmURL, SubNodes, ResponseType, imglist
	Dim myURL, yourURL
	If (application("Svr_Info") = "Dev") Then
		nvstorefarmURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstorefarmURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstorefarmURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iservice & "#" & ioperation
		objXML.send(strParam)
 		If objXML.Status = 200 Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(objXML.responseText)
'response.write objXML.responseText
'response.end
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
				If ResponseType = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] " & VbCRLF
					strSql = strSql & " SET nvClassSellyn = '"&ichgSellYn&"'" & VbCRLF
					strSql = strSql & " ,nvClassLastUpdate = getdate()" & VbCRLF
					strSql = strSql & " ,accFailCNT=0" & VbCRLF
					strSql = strSql & " WHERE itemid = "&iitemid
					dbget.Execute(strSql)

					If ichgSellYn = "N" Then
						iErrStr = "OK||"&iitemid&"||�Ǹ�����(����)"
					Else
						iErrStr = "OK||"&iitemid&"||�Ǹ���(����)"
					End If
				Else
					iMessage = xmlDOM.getElementsByTagName("n:Message")(0).Text
					If InStr(iMessage, "���ռ� ����") OR InStr(iMessage, "�ڼ��� ������ Detail ������Ʈ") Then
						iMessage = xmlDOM.getElementsByTagName("n:Detail")(0).Text
					End If
					If InStr(iMessage, "ǰ�� ��ǰ�� �Ǹ��� ���·θ� ������ �� �ֽ��ϴ�") OR InStr(iMessage, "�Ǹű��� ������") Then
						strSql = ""
						strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] " & VbCRLF
						strSql = strSql & " SET nvClassSellyn = 'N'" & VbCRLF
						strSql = strSql & " ,nvClassLastUpdate = getdate()" & VbCRLF
						strSql = strSql & " ,accFailCNT=0" & VbCRLF
						strSql = strSql & " WHERE itemid = "&iitemid
						dbget.Execute(strSql)
						iErrStr = "OK||"&iitemid&"||�Ǹ�����(����)/������ ����ó��"
					Else
						iErrStr = "ERR||"&iitemid&"||"&iMessage
					End If
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||������� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-SELLEDIT]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'��ǰ ����
Public Function fnNvClassDelete(iitemid, strParam, byRef iErrStr, iservice, ioperation)
	Dim objXML, xmlDOM, strSql, iMessage, nvstorefarmURL, SubNodes, ResponseType, imglist
	Dim myURL, yourURL
	If (application("Svr_Info") = "Dev") Then
		nvstorefarmURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstorefarmURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstorefarmURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iservice & "#" & ioperation
		objXML.send(strParam)
 		If objXML.Status = 200 Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(objXML.responseText)
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
				If ResponseType = "SUCCESS" Then
					strSql = ""
					strSql = strSql &" INSERT INTO [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] " & VBCRLF
					strSql = strSql &" SELECT TOP 1 'nvstorefarmclass', i.itemid, r.nvClassGoodNo, r.nvClassRegdate, getdate(), r.lastErrStr" & VBCRLF
					strSql = strSql &" FROM db_item.dbo.tbl_item as i " & VBCRLF
					strSql = strSql &" JOIN db_etcmall.dbo.tbl_nvstorefarmclass_regItem as r on i.itemid = r.itemid " & VBCRLF
					strSql = strSql &" WHERE i.itemid = "&iitemid & VBCRLF
					dbget.Execute(strSql)

					strSql = ""
					strSql = strSql & " DELETE FROM [db_etcmall].[dbo].[tbl_nvstorefarmclass_regItem] " & vbcrlf
					strSql = strSql & " WHERE itemid = '"&iitemid&"' "
					dbget.Execute(strSql)

					strSql = ""
					strSql = strSql & " DELETE FROM [db_etcmall].[dbo].[tbl_nvstorefarmclass_Image] " & vbcrlf
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

					iErrStr = "OK||"&iitemid&"||����(��ǰ)"
				Else
					iMessage = xmlDOM.getElementsByTagName("n:Message")(0).Text
					If InStr(iMessage, "���ռ� ����") Then
						iMessage = xmlDOM.getElementsByTagName("n:Detail")(0).Text
					End If
					iErrStr = "ERR||"&iitemid&"||"&iMessage
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||������� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-SELLEDIT]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'��ǰ ���� ���� XML
Public Function getNvClassSellynParameter(invClassGoodNo, ichgSellYn, iservice, ioperation)
	Dim stopYN, strRst, oaccessLicense, oTimestamp, osignature, reqID
	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		reqID = "ncp_1np6kl_01"
	End If

	If ichgSellYn = "N" Then
		stopYN = "SUSP"		'�Ǹ�����
	ElseIf ichgSellYn = "Y" Then
		stopYN = "SALE"		'�Ǹ�
	End If
	Call getsecretKey(oaccessLicense, oTimestamp, osignature, iservice, ioperation)
	strRst = ""
	strRst = strRst &"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
	strRst = strRst &"	<soapenv:Header/>"
	strRst = strRst &"	<soapenv:Body>"
	strRst = strRst &"		<shop:ChangeProductSaleStatusRequest>"
	strRst = strRst &"			<shop:RequestID>"&reqID&"</shop:RequestID>"
	strRst = strRst &"			<shop:AccessCredentials>"
	strRst = strRst &"				<shop:AccessLicense>"&oaccessLicense&"</shop:AccessLicense>"
	strRst = strRst &"				<shop:Timestamp>"&oTimestamp&"</shop:Timestamp>"
	strRst = strRst &"				<shop:Signature>"&osignature&"</shop:Signature>"
	strRst = strRst &"			</shop:AccessCredentials>"
	strRst = strRst &"			<shop:Version>2.0</shop:Version>"
	strRst = strRst &"			<SellerId>"&reqID&"</SellerId>"
	strRst = strRst &"			<SaleStatus>"
	strRst = strRst &"				<shop:ProductId>"&invClassGoodNo&"</shop:ProductId>"
	strRst = strRst &"				<shop:StatusType>"&stopYN&"</shop:StatusType>"
	strRst = strRst &"			</SaleStatus>"
	strRst = strRst &"		</shop:ChangeProductSaleStatusRequest>"
	strRst = strRst &"	</soapenv:Body>"
	strRst = strRst &"</soapenv:Envelope>"
	getNvClassSellynParameter = strRst
End Function

'��ǰ �ɼ� ��� XML
Public Function getNvClassOptionRegXML(iitemid, invClassGoodNo, iservice, ioperation)
	Dim strRst, oaccessLicense, oTimestamp, osignature, limitYCnt
	Dim strSql, iitemdiv, ioptioncnt, chkMultiOpt, MultiTypeCnt, arrMultiTypeNm, i, j, k
	Dim optNm, optLimit, ilimityn, itemoption, optDc, optIsusing, optSellYn, optaddprice, optionTypeName, reqID
	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		reqID = "ncp_1np6kl_01"
	End If

	Call getsecretKey(oaccessLicense, oTimestamp, osignature, iservice, ioperation)
	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.limityn, i.itemdiv, i.optioncnt, isnull(o.optionTypeName, '') as optionTypeName "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
	strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		ilimityn		= rsget("limityn")
		iitemdiv		= rsget("itemdiv")
		ioptioncnt		= rsget("optioncnt")
		If Trim(rsget("optionTypeName")) = "" Then
			optionTypeName	= "�ɼ�"
		Else
			optionTypeName	= rsget("optionTypeName")
		End If
	rsget.Close

	strSql = ""
	strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & iitemid
    rsget.CursorLocation = adUseClient
	rsget.CursorType = adOpenStatic
	rsget.LockType = adLockOptimistic
    rsget.Open strSql, dbget
	If Not(rsget.EOF or rsget.BOF) Then
		chkMultiOpt = true
		MultiTypeCnt = rsget.recordcount
		For i = 1 to rsget.recordcount
			arrMultiTypeNm = arrMultiTypeNm &"						<shop:Name"&i&">"&db2Html(rsget("optionTypeName"))&"</shop:Name"&i&">"			'#�ɼǸ�1~5
			rsget.MoveNext
			If i > 4 Then Exit For
		Next
	End If
	rsget.Close

	If (ioptioncnt > 0) Then
		strRst = ""
		strRst = strRst &"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
		strRst = strRst &"	<soapenv:Header/>"
		strRst = strRst &"	<soapenv:Body>"
		strRst = strRst &"		<shop:ManageOptionRequest>"
		strRst = strRst &"			<shop:RequestID>"&reqID&"</shop:RequestID>"
		strRst = strRst &"			<shop:AccessCredentials>"
		strRst = strRst &"				<shop:AccessLicense>"&oaccessLicense&"</shop:AccessLicense>"
		strRst = strRst &"				<shop:Timestamp>"&oTimestamp&"</shop:Timestamp>"
		strRst = strRst &"				<shop:Signature>"&osignature&"</shop:Signature>"
		strRst = strRst &"			</shop:AccessCredentials>"
		strRst = strRst &"			<shop:Version>2.0</shop:Version>"
		strRst = strRst &"			<SellerId>"&reqID&"</SellerId>"
		strRst = strRst &"			<Option>"
		strRst = strRst &"				<shop:ProductId>"&invClassGoodNo&"</shop:ProductId>"

		If ioptioncnt > 0 Then
			strRst = strRst &"				<shop:Combination>"
			strRst = strRst &"					<shop:Names>"
			If chkMultiOpt = true Then
				strRst = strRst & arrMultiTypeNm
			Else
				strRst = strRst &"						<shop:Name1><![CDATA["&optionTypeName&"]]></shop:Name1>"	'#�ɼǸ�1
			End If
			strRst = strRst &"					</shop:Names>"
			strRst = strRst &"					<shop:ItemList>"
			If chkMultiOpt = true Then																'���տɼ� �̶��
				strSql = ""
				strSql = strSql &"  SELECT itemoption, isusing, optsellyn, optaddprice, optionname, (optlimitno-optlimitsold) as optLimit "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE isUsing='Y' and optsellyn='Y' and itemid=" & iitemid
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) then
					For j = 1 to rsget.recordcount
						optLimit = rsget("optLimit")
						If (optLimit < 1) Then optLimit = 0
						If (ilimityn <> "Y") Then optLimit = 9999
						itemoption	= rsget("itemoption")
						optDc		= db2Html(rsget("optionname"))
						optaddprice	= rsget("optaddprice")

						strRst = strRst &"						<shop:Item>"
			'			strRst = strRst &"							<shop:Id></shop:Id>"															'�ɼ�ID | �ɼ� ID �Է½� ���� �ɼ� ����
						For k = 1 to MultiTypeCnt
							If InStr(optDc, ",") Then
								strRst = strRst &"							<shop:Value"&k&">"&Split(optDc,",")(k-1)&"</shop:Value"&k&">"				'#�ɼǸ�1�� �ش��ϴ� �ɼǰ�
							Else
								strRst = strRst &"							<shop:Value"&k&">"&optDc&"</shop:Value"&k&">"								'#�ɼǸ�1�� �ش��ϴ� �ɼǰ�
							End If
						Next
						strRst = strRst &"							<shop:Price>"&optaddprice&"</shop:Price>"										'�ɼǰ� | ���Է½� 0��
						strRst = strRst &"							<shop:Quantity>"&optLimit&"</shop:Quantity>"									'��� ���� | ���Է½� 0��
						strRst = strRst &"							<shop:SellerManagerCode><![CDATA["&itemoption&"]]></shop:SellerManagerCode>"	'�ǸŰ� ���� �ڵ�
						strRst = strRst &"							<shop:Usable>Y</shop:Usable>"													'#��� ���� | Y or N
						strRst = strRst &"						</shop:Item>"
						rsget.MoveNext
					Next
				end if
				rsget.Close
			Else																						'���� �ɼ� �̶��
				strSql = ""
				strSql = strSql &"  SELECT itemoption, isusing, optsellyn, optaddprice, optionname, (optlimitno-optlimitsold) as optLimit "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE isUsing='Y' and optsellyn='Y' and itemid=" & iitemid
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
						optLimit = rsget("optLimit")
						If (optLimit < 1) Then optLimit = 0
						If (ilimityn <> "Y") Then optLimit = 9999
						itemoption	= rsget("itemoption")
						optDc		= db2Html(rsget("optionname"))
						optaddprice	= rsget("optaddprice")

						If (optLimit > 0) Then
							 limitYCnt =  limitYCnt + 1
						End If

						strRst = strRst &"						<shop:Item>"
			'			strRst = strRst &"							<shop:Id></shop:Id>"															'�ɼ�ID | �ɼ� ID �Է½� ���� �ɼ� ����
						strRst = strRst &"							<shop:Value1><![CDATA["&optDc&"]]></shop:Value1>"								'#�ɼǸ�1�� �ش��ϴ� �ɼǰ�
						strRst = strRst &"							<shop:Price>"&optaddprice&"</shop:Price>"										'�ɼǰ� | ���Է½� 0��
						strRst = strRst &"							<shop:Quantity>"&optLimit&"</shop:Quantity>"									'��� ���� | ���Է½� 0��
						strRst = strRst &"							<shop:SellerManagerCode><![CDATA["&itemoption&"]]></shop:SellerManagerCode>"	'�ǸŰ� ���� �ڵ�
						strRst = strRst &"							<shop:Usable>Y</shop:Usable>"													'#��� ���� | Y or N
						strRst = strRst &"						</shop:Item>"
						rsget.MoveNext
					Loop

'					If limitYCnt = 0 Then
'						FMayLimitSoldout = "Y"
'					Else
'						FMayLimitSoldout = "N"
'					End If
				end if
				rsget.Close
			End If
			strRst = strRst &"					</shop:ItemList>"
			strRst = strRst &"				</shop:Combination>"
		End If
		strRst = strRst &"			</Option>"
		strRst = strRst &"		</shop:ManageOptionRequest>"
		strRst = strRst &"	</soapenv:Body>"
		strRst = strRst &"</soapenv:Envelope>"
		getNvClassOptionRegXML = strRst
	Else
		Dim isRegedOptCnt
		strSql = ""
		strSql = strSql &"  SELECT TOP 1 isnull(regedOptcnt, 0) as regedOptcnt "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem]"
		strSql = strSql & " WHERE nvClassStatcd = 7 and itemid=" & iitemid
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			isRegedOptCnt = rsget("regedOptcnt")
		End If
		rsget.Close

		If (ioptioncnt = 0) and (isRegedOptCnt > 0) Then
			strRst = ""
			strRst = strRst &"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
			strRst = strRst &"	<soapenv:Header/>"
			strRst = strRst &"	<soapenv:Body>"
			strRst = strRst &"		<shop:ManageOptionRequest>"
			strRst = strRst &"			<shop:RequestID>"&reqID&"</shop:RequestID>"
			strRst = strRst &"			<shop:AccessCredentials>"
			strRst = strRst &"				<shop:AccessLicense>"&oaccessLicense&"</shop:AccessLicense>"
			strRst = strRst &"				<shop:Timestamp>"&oTimestamp&"</shop:Timestamp>"
			strRst = strRst &"				<shop:Signature>"&osignature&"</shop:Signature>"
			strRst = strRst &"			</shop:AccessCredentials>"
			strRst = strRst &"			<shop:Version>2.0</shop:Version>"
			strRst = strRst &"			<SellerId>"&reqID&"</SellerId>"
			strRst = strRst &"			<Option>"
			strRst = strRst &"				<shop:ProductId>"&invClassGoodNo&"</shop:ProductId>"
			strRst = strRst &"			</Option>"
			strRst = strRst &"		</shop:ManageOptionRequest>"
			strRst = strRst &"	</soapenv:Body>"
			strRst = strRst &"</soapenv:Envelope>"
			getNvClassOptionRegXML = strRst
		Else
			getNvClassOptionRegXML = "X"
		End If
	End If
End Function

'��ǰ ��ȸ XML
Public Function getNvClassItemSearchParameter(invClassGoodNo, iservice, ioperation)
	Dim strRst, oaccessLicense, oTimestamp, osignature, reqID
	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		reqID = "ncp_1np6kl_01"
	End If

	Call getsecretKey(oaccessLicense, oTimestamp, osignature, iservice, ioperation)
	strRst = ""
	strRst = strRst &"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
	strRst = strRst &"	<soapenv:Header/>"
	strRst = strRst &"	<soapenv:Body>"
	strRst = strRst &"		<shop:GetProductRequest>"
	strRst = strRst &"			<shop:RequestID>"&reqID&"</shop:RequestID>"
	strRst = strRst &"			<shop:AccessCredentials>"
	strRst = strRst &"				<shop:AccessLicense>"&oaccessLicense&"</shop:AccessLicense>"
	strRst = strRst &"				<shop:Timestamp>"&oTimestamp&"</shop:Timestamp>"
	strRst = strRst &"				<shop:Signature>"&osignature&"</shop:Signature>"
	strRst = strRst &"			</shop:AccessCredentials>"
	strRst = strRst &"			<shop:Version>2.0</shop:Version>"
	strRst = strRst &"			<SellerId>"&reqID&"</SellerId>"
	strRst = strRst &"			<ProductId>"&invClassGoodNo&"</ProductId>"
	strRst = strRst &"		</shop:GetProductRequest>"
	strRst = strRst &"	</soapenv:Body>"
	strRst = strRst &"</soapenv:Envelope>"
	getNvClassItemSearchParameter = strRst
End Function

'�ɼ� ��ȸ XML
Public Function getNvClassOptionSearchParameter(invClassGoodNo, iservice, ioperation)
	Dim strRst, oaccessLicense, oTimestamp, osignature, reqID
	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		reqID = "ncp_1np6kl_01"
	End If

	Call getsecretKey(oaccessLicense, oTimestamp, osignature, iservice, ioperation)
	strRst = ""
	strRst = strRst &"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
	strRst = strRst &"	<soapenv:Header/>"
	strRst = strRst &"	<soapenv:Body>"
	strRst = strRst &"		<shop:GetOptionRequest>"
	strRst = strRst &"			<shop:RequestID>"&reqID&"</shop:RequestID>"
	strRst = strRst &"			<shop:AccessCredentials>"
	strRst = strRst &"				<shop:AccessLicense>"&oaccessLicense&"</shop:AccessLicense>"
	strRst = strRst &"				<shop:Timestamp>"&oTimestamp&"</shop:Timestamp>"
	strRst = strRst &"				<shop:Signature>"&osignature&"</shop:Signature>"
	strRst = strRst &"			</shop:AccessCredentials>"
	strRst = strRst &"			<shop:Version>2.0</shop:Version>"
	strRst = strRst &"			<SellerId>"&reqID&"</SellerId>"
	strRst = strRst &"			<ProductId>"&invClassGoodNo&"</ProductId>"
	strRst = strRst &"		</shop:GetOptionRequest>"
	strRst = strRst &"	</soapenv:Body>"
	strRst = strRst &"</soapenv:Envelope>"
	getNvClassOptionSearchParameter = strRst
End Function

'��ǰ ���� XML
Public Function getNvClassDeleteParameter(invClassGoodNo, iservice, ioperation)
	Dim strRst, oaccessLicense, oTimestamp, osignature, reqID
	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		reqID = "ncp_1np6kl_01"
	End If

	Call getsecretKey(oaccessLicense, oTimestamp, osignature, iservice, ioperation)
	strRst = ""
	strRst = strRst &"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
	strRst = strRst &"	<soapenv:Header/>"
	strRst = strRst &"	<soapenv:Body>"
	strRst = strRst &"		<shop:DeleteProductRequest>"
	strRst = strRst &"			<shop:RequestID>"&reqID&"</shop:RequestID>"
	strRst = strRst &"			<shop:AccessCredentials>"
	strRst = strRst &"				<shop:AccessLicense>"&oaccessLicense&"</shop:AccessLicense>"
	strRst = strRst &"				<shop:Timestamp>"&oTimestamp&"</shop:Timestamp>"
	strRst = strRst &"				<shop:Signature>"&osignature&"</shop:Signature>"
	strRst = strRst &"			</shop:AccessCredentials>"
	strRst = strRst &"			<shop:Version>2.0</shop:Version>"
	strRst = strRst &"			<SellerId>"&reqID&"</SellerId>"
	strRst = strRst &"			<ProductId>"&invClassGoodNo&"</ProductId>"
	strRst = strRst &"		</shop:DeleteProductRequest>"
	strRst = strRst &"	</soapenv:Body>"
	strRst = strRst &"</soapenv:Envelope>"
	getNvClassDeleteParameter = strRst
End Function
%>