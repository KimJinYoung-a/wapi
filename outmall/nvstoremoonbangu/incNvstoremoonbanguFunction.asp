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
Public Function fnNvstoremoonbanguImageReg(iitemid, strParam, byRef iErrStr, ichgImageNm, iservice, ioperation)
	Dim objXML, xmlDOM, strSql, iMessage, nvstoremoonbanguURL, SubNodes, ResponseType, imglist
	Dim myURL, yourURL
	If (application("Svr_Info") = "Dev") Then
		nvstoremoonbanguURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstoremoonbanguURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstoremoonbanguURL, False
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
					strSql = strSql & " db_etcmall.[dbo].[tbl_nvstoremoonbangu_Image] " & vbcrlf
					strSql = strSql & " WHERE itemid = '"&iitemid&"' "
					dbget.Execute strSql

					Set imglist = xmlDOM.getElementsByTagName("n:Image")
					For Each SubNodes in imglist
						myURL	= SubNodes.SelectSingleNode("n:Source").text
						yourURL	= SubNodes.SelectSingleNode("n:URL").text

						If InStr(myURL, "/basic/") OR InStr(myURL, "/nvadd1/") Then
							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstoremoonbangu_Image] (itemid, imgType, tenURL, storefarmURL) VALUES " & vbcrlf
							strSql = strSql & " ('"&iitemid&"', '1', '"&myURL&"', '"&yourURL&"') " & vbcrlf
							dbget.Execute strSql
						Else
							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstoremoonbangu_Image] (itemid, imgType, tenURL, storefarmURL) VALUES " & vbcrlf
							strSql = strSql & " ('"&iitemid&"', '2', '"&myURL&"', '"&yourURL&"') " & vbcrlf
							dbget.Execute strSql
						End If
					Next
					Set imglist = nothing
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] SET "
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

'��ǰ �շ�
Public Function fnNvstoremoonbanguItemReg(iitemid, strParam, byRef iErrStr, iSellCash, iNvstoremoonbanguSellYn, ilimityn, ilimitno, ilimiysold, iitemname, iimageNm, iservice, ioperation, ichkXML)
	Dim objXML, xmlDOM, strSql, iMessage, nvstoremoonbanguURL, SubNodes, ResponseType, imglist
	Dim ProductId
	If (application("Svr_Info") = "Dev") Then
		nvstoremoonbanguURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstoremoonbanguURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If
	On Error Resume Next
	If ichkXML = "Y" Then
		response.write strParam
'		response.end
	End If
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstoremoonbanguURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iservice & "#" & ioperation
		objXML.send(strParam)
 		If objXML.Status = 200 Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(objXML.responseText)
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
'				If ResponseType = "SUCCESS" Then
					ProductId = xmlDOM.getElementsByTagName("ProductId").item(0).text
				If ResponseType = "SUCCESS" AND ProductId <> "" Then
					strSql = strSql & " UPDATE R " & vbcrlf
					strSql = strSql & " SET nvstoremoonbanguGoodNo = '"&ProductId&"' " & vbcrlf
					strSql = strSql & " , nvstoremoonbanguLastUpdate = getdate() " & vbcrlf
					strSql = strSql & " , nvstoremoonbanguPrice = " & iSellCash & vbcrlf
					strSql = strSql & " , accFailCnt = 0 " & vbcrlf
					strSql = strSql & " , nvstoremoonbanguRegdate = getdate() " & vbcrlf
					strSql = strSql & " , nvstoremoonbanguStatCd = 7 " & vbcrlf
					strSql = strSql & "	FROM db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] R " & vbcrlf
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute strSql
					iErrStr = "OK||"&iitemid&"||����(��ǰ���)"
				Else
					If ichkXML = "Y" Then
						response.write objXML.responseText
					End If

					iMessage = xmlDOM.getElementsByTagName("n:Message")(0).Text
					If InStr(iMessage, "���ռ� ����") Then
						iMessage = xmlDOM.getElementsByTagName("n:Detail")(0).Text
					End If
					iErrStr = "ERR||"&iitemid&"||"&iMessage
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||������� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'��ǰ ����
Public Function fnNvstoremoonbanguItemEDIT(iitemid, strParam, byRef iErrStr, iSellCash, iNvstoremoonbanguSellYn, ilimityn, ilimitno, ilimiysold, iitemname, ichgImageNm, iservice, ioperation)
	Dim objXML, xmlDOM, strSql, iMessage, nvstoremoonbanguURL, SubNodes, ResponseType, imglist
	Dim ProductId
	If (application("Svr_Info") = "Dev") Then
		nvstoremoonbanguURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstoremoonbanguURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstoremoonbanguURL, False
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
					strSql = strSql & " SET nvstoremoonbanguLastUpdate = getdate() " & vbcrlf
					strSql = strSql & " , nvstoremoonbanguPrice = " & iSellCash & vbcrlf
					strSql = strSql & " , accFailCnt = 0 " & vbcrlf
					If (ichgImageNm <> "N") Then
						strSql = strSql & " ,regimageName='"&ichgImageNm&"'"& VbCrlf
					End If
					strSql = strSql & " , regitemname = '"&html2db(iitemname)&"'" & vbcrlf
					strSql = strSql & "	FROM db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] R " & vbcrlf
					strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
					dbget.Execute strSql
					iErrStr = "OK||"&iitemid&"||����(��ǰ����)"
				Else
					iMessage = xmlDOM.getElementsByTagName("n:Message")(0).Text
					If InStr(iMessage, "���ռ� ����") Then
						iMessage = xmlDOM.getElementsByTagName("n:Detail")(0).Text
						If InStr(iMessage, "�ɼ��� �ɼǰ�/��뿩��") Then
							iMessage = "����� �ɼ��� �ɼǰ�/��뿩�� �׸��� Ȯ�� �� ������ �ּ���."
						End If
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
Public Function fnNvstoremoonbanguOptionReg(iitemid, strParam, byRef iErrStr, iservice, ioperation)
	Dim objXML, xmlDOM, strSql, iMessage, nvstoremoonbanguURL, SubNodes, ResponseType, imglist
	Dim myURL, yourURL, statusType, nvRegitemname, MasterPrice, ProductId
	If (application("Svr_Info") = "Dev") Then
		nvstoremoonbanguURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstoremoonbanguURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstoremoonbanguURL, False
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
					If InStr(iMessage, "���ռ� ����") Then
						iMessage = xmlDOM.getElementsByTagName("n:Detail")(0).Text
						If InStr(iMessage, "�ɼ��� �ɼǰ�/��뿩��") Then
							iMessage = "����� �ɼ��� �ɼǰ�/��뿩�� �׸��� Ȯ�� �� ������ �ּ���."
						End If
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
Public Function fnNvstoremoonbanguItemSearch(iitemid, strParam, byRef iErrStr, iservice, ioperation)
	Dim objXML, xmlDOM, strSql, iMessage, nvstoremoonbanguURL, SubNodes, ResponseType, imglist
	Dim myURL, yourURL, statusType, nvRegitemname, MasterPrice, ProductId
	If (application("Svr_Info") = "Dev") Then
		nvstoremoonbanguURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstoremoonbanguURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstoremoonbanguURL, False
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
					strSQL = strSQL & " SET nvstoremoonbanguPrice = " & MasterPrice & VbCRLF
					strSQL = strSQL & " ,nvstoremoonbanguSellyn='"&statusType&"'" & VbCRLF
					strSQL = strSQL & " ,regitemname='"&html2db(nvRegitemname)&"'" & VbCRLF
					strSQL = strSQL & " ,lastStatCheckDate = getdate()" & VbCRLF
					strSQL = strSQL & " ,nvstoremoonbanguGoodNo = isNULL(R.nvstoremoonbanguGoodNo,'"&ProductId&"')"&VbCRLF
					strSQL = strSQL & " FROM db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] R" & VbCRLF
					strSQL = strSQL & " where R.itemid="&iitemid & VbCRLF
					strSQL = strSQL & " and isNULL(nvstoremoonbanguGoodNo,'') in ('','"&ProductId&"')"&VbCRLF    ''�ߺ���ϵ�CaSE ���
					strSQL = strSQL & " and (isNULL(nvstoremoonbanguPrice,0)<>"&MasterPrice&"" & VbCRLF
					strSQL = strSQL & "     or isNULL(nvstoremoonbanguSellyn,'')<>'"&statusType&"'"& VbCRLF
					strSQL = strSQL & "     or isNULL(regitemname,'')<>'"&html2db(nvRegitemname)&"'"& VbCRLF
					strSQL = strSQL & "     or isNULL(nvstoremoonbanguGoodNo,'')<>'"&ProductId&"'"& VbCRLF
					strSQL = strSQL & " )"
				    dbget.Execute strSQL
					iErrStr =  "OK||"&iitemid&"||����(�ǸŻ�����ȸ)"
				Else
					iMessage = xmlDOM.getElementsByTagName("n:Message")(0).Text
					If InStr(iMessage, "���ռ� ����") Then
						iMessage = xmlDOM.getElementsByTagName("n:Detail")(0).Text
						If InStr(iMessage, "�ɼ��� �ɼǰ�/��뿩��") Then
							iMessage = "����� �ɼ��� �ɼǰ�/��뿩�� �׸��� Ȯ�� �� ������ �ּ���."
						End If
					End If
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(�ǸŻ�����ȸ)"
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||������� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-ITEMSEARCH]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'�ɼ� ��ȸ
Public Function fnNvstoremoonbanguOptionSearch(iitemid, strParam, byRef iErrStr, iservice, ioperation)
	Dim objXML, xmlDOM, strSQL, iMessage, nvstoremoonbanguURL, SubNodes, ResponseType, imglist
	Dim myURL, yourURL, statusType, nvRegitemname, MasterPrice, ProductId
	Dim Nodes, onvOptId, myOptNo, addprice, saleLmtQty, nvOptval1, nvOptval2, nvOptval3, nvOptval4, nvOptval5, AlloptNm
	If (application("Svr_Info") = "Dev") Then
		nvstoremoonbanguURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstoremoonbanguURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstoremoonbanguURL, False
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

								' strSQL = ""
								' strSQL = strSQL & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
								' strSQL = strSQL & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastUpdate, checkdate) "
								' strSQL = strSQL & " VALUES "
								' strSQL = strSQL & " ('"&iitemid&"'"
								' strSQL = strSQL & ",  '"&myOptNo&"'"
								' strSQL = strSQL & ", '"&CMALLNAME&"'"
								' strSQL = strSQL & ", '"&onvOptId&"'"
								' strSQL = strSQL & ", '"&html2db(AlloptNm)&"'"
								' strSQL = strSQL & ", 'Y'"
								' strSQL = strSQL & ", '"&"Y"&"'"
								' strSQL = strSQL & ", '"&saleLmtQty&"'"
								' strSQL = strSQL & ", getdate() "
								' strSQL = strSQL & ", getdate()) "
								' dbget.Execute strSQL

								strSQL = ""
								strSQL = strSQL & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption "
								strSQL = strSQL & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, lastUpdate, checkdate) "
								strSQL = strSQL & " SELECT itemid, itemoption, '"&CMALLNAME&"', '"&onvOptId&"', optionname, '"&Chkiif(saleLmtQty > 0,"Y","N")&"', 'Y', '"&saleLmtQty&"', getdate(), getdate() "
								strSQL = strSQL & " FROM db_item.dbo.tbl_item_option "
								strSQL = strSQL & " WHERE itemid = '"& iitemid &"' "
								strSQL = strSQL & " and itemoption = '"& myOptNo &"' "
								dbget.Execute strSQL
							Next
						Set Nodes = nothing
					End If
					strSQL = ""
					strSQL = strSQL & " UPDATE R"   &VbCRLF
					strSQL = strSQL & " SET regedOptCnt=isNULL(T.regedOptCnt,0)"   &VbCRLF
					strSQL = strSQL & " FROM db_etcmall.dbo.tbl_nvstoremoonbangu_regItem R"   &VbCRLF
					strSQL = strSQL & " JOIN ("   &VbCRLF
					strSQL = strSQL & " 	SELECT R.itemid,count(*) as CNT , sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt "   &VbCRLF
					strSQL = strSQL & " 	FROM db_etcmall.dbo.tbl_nvstoremoonbangu_regItem R "   &VbCRLF
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
					iErrStr = "ERR||"&iitemid&"||"&iMessage&"(�ɼ���ȸ)"
				End If
			Set xmlDOM = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||������� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-OPTIONSEARCH]"
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'�Ǹ� ���� ����
Public Function fnNvstoremoonbanguSellyn(iitemid, ichgSellYn, strParam, byRef iErrStr, iservice, ioperation)
	Dim objXML, xmlDOM, strSql, iMessage, nvstoremoonbanguURL, SubNodes, ResponseType, imglist
	Dim myURL, yourURL
	If (application("Svr_Info") = "Dev") Then
		nvstoremoonbanguURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstoremoonbanguURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstoremoonbanguURL, False
		objXML.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
		objXML.setRequestHeader "SOAPAction", iservice & "#" & ioperation
		objXML.send(strParam)
 		If objXML.Status = 200 Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.loadXML(objXML.responseText)
If (session("ssBctID") = "kjy8517") Then
	'response.write objXML.responseText
end if
'response.write objXML.responseText
'response.end
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
				If ResponseType = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] " & VbCRLF
					strSql = strSql & " SET nvstoremoonbanguSellYn = '"&ichgSellYn&"'" & VbCRLF
					strSql = strSql & " ,nvstoremoonbanguLastUpdate = getdate()" & VbCRLF
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
						strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] " & VbCRLF
						strSql = strSql & " SET nvstoremoonbanguSellYn = 'N'" & VbCRLF
						strSql = strSql & " ,nvstoremoonbanguLastUpdate = getdate()" & VbCRLF
						strSql = strSql & " ,accFailCNT=0" & VbCRLF
						strSql = strSql & " WHERE itemid = "&iitemid
						dbget.Execute(strSql)
						iErrStr = "OK||"&iitemid&"||�Ǹ�����(����)/������ ����ó��"
					Else
						iErrStr = "ERR||"&iitemid&"||"&iMessage&"(����)"
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
Public Function fnNvstoremoonbanguDelete(iitemid, strParam, byRef iErrStr, iservice, ioperation)
	Dim objXML, xmlDOM, strSql, iMessage, nvstoremoonbanguURL, SubNodes, ResponseType, imglist
	Dim myURL, yourURL
	If (application("Svr_Info") = "Dev") Then
		nvstoremoonbanguURL = "http://sandbox.api.naver.com/ShopN/"&iservice
	Else
		nvstoremoonbanguURL = "http://ec.api.naver.com/ShopN/"&iservice
	End If

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", nvstoremoonbanguURL, False
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
					strSql = strSql &" SELECT TOP 1 'nvstoremoonbangu', i.itemid, r.nvstoremoonbanguGoodNo, r.nvstoremoonbanguregdate, getdate(), r.lastErrStr" & VBCRLF
					strSql = strSql &" FROM db_item.dbo.tbl_item as i " & VBCRLF
					strSql = strSql &" JOIN db_etcmall.dbo.tbl_nvstoremoonbangu_regItem as r on i.itemid = r.itemid " & VBCRLF
					strSql = strSql &" WHERE i.itemid = "&iitemid & VBCRLF
					dbget.Execute(strSql)

					strSql = ""
					strSql = strSql & " DELETE FROM [db_etcmall].[dbo].[tbl_nvstoremoonbangu_regItem] " & vbcrlf
					strSql = strSql & " WHERE itemid = '"&iitemid&"' "
					dbget.Execute(strSql)

					strSql = ""
					strSql = strSql & " DELETE FROM [db_etcmall].[dbo].[tbl_nvstoremoonbangu_Image] " & vbcrlf
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
Public Function getNvstoremoonbanguSellynParameter(iNvstoremoonbangugoodno, ichgSellYn, iservice, ioperation)
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
	strRst = strRst &"				<shop:ProductId>"&iNvstoremoonbangugoodno&"</shop:ProductId>"
	strRst = strRst &"				<shop:StatusType>"&stopYN&"</shop:StatusType>"
	strRst = strRst &"			</SaleStatus>"
	strRst = strRst &"		</shop:ChangeProductSaleStatusRequest>"
	strRst = strRst &"	</soapenv:Body>"
	strRst = strRst &"</soapenv:Envelope>"
	getNvstoremoonbanguSellynParameter = strRst
End Function

'��ǰ �ɼ� ��� XML
Public Function getNvstoremoonbanguOptionRegXML(iitemid, invstoremoonbangugoodno, iservice, ioperation)
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

	If (ioptioncnt > 0) OR (iitemdiv = "06") Then
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
		strRst = strRst &"				<shop:ProductId>"&invstoremoonbangugoodno&"</shop:ProductId>"

		If iitemdiv = "06" Then
			strRst = strRst &"				<shop:CustomList>"										'���� �Է��� �ɼ� | �ܵ��� �ɼ�, ������ �ɼ�, ���� �Է��� �ɼ� �� �ּ� �Ѱ��� �Է�
			strRst = strRst &"					<shop:Custom>"
'			strRst = strRst &"						<shop:Id></shop:Id>"							'�ɼ�ID | �ɼ� ID �Է� �� ���� �ɼ� ����
			strRst = strRst &"						<shop:Name><![CDATA[�����Է�]]></shop:Name>"	'#�ɼǸ�
			strRst = strRst &"						<shop:Usable>Y</shop:Usable>"					'#��� ���� | Y or N
			strRst = strRst &"					</shop:Custom>"
			strRst = strRst &"				</shop:CustomList>"
		End If

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
						optLimit = optLimit-5
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
						optLimit = optLimit-5
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
		getNvstoremoonbanguOptionRegXML = strRst
	Else
		Dim isRegedOptCnt
		strSql = ""
		strSql = strSql &"  SELECT TOP 1 isnull(regedOptcnt, 0) as regedOptcnt "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem]"
		strSql = strSql & " WHERE nvstoremoonbanguStatcd = 7 and itemid=" & iitemid
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
			strRst = strRst &"				<shop:ProductId>"&invstoremoonbangugoodno&"</shop:ProductId>"
			strRst = strRst &"			</Option>"
			strRst = strRst &"		</shop:ManageOptionRequest>"
			strRst = strRst &"	</soapenv:Body>"
			strRst = strRst &"</soapenv:Envelope>"
			getNvstoremoonbanguOptionRegXML = strRst
		Else
			getNvstoremoonbanguOptionRegXML = "X"
		End If
	End If
End Function

'��ǰ ��ȸ XML
Public Function getNvstoremoonbanguItemSearchParameter(invstoremoonbangugoodno, iservice, ioperation)
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
	strRst = strRst &"			<ProductId>"&invstoremoonbangugoodno&"</ProductId>"
	strRst = strRst &"		</shop:GetProductRequest>"
	strRst = strRst &"	</soapenv:Body>"
	strRst = strRst &"</soapenv:Envelope>"
	getNvstoremoonbanguItemSearchParameter = strRst
End Function

'�ɼ� ��ȸ XML
Public Function getNvstoremoonbanguOptionSearchParameter(invstoremoonbangugoodno, iservice, ioperation)
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
	strRst = strRst &"			<ProductId>"&invstoremoonbangugoodno&"</ProductId>"
	strRst = strRst &"		</shop:GetOptionRequest>"
	strRst = strRst &"	</soapenv:Body>"
	strRst = strRst &"</soapenv:Envelope>"
	getNvstoremoonbanguOptionSearchParameter = strRst
End Function

'�ɼ� ��ȸ XML
Public Function fnAuctionCommonCode(iservice, ioperation)
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
	strRst = strRst &"			<ProductId>"&invstoremoonbangugoodno&"</ProductId>"
	strRst = strRst &"		</shop:GetOptionRequest>"
	strRst = strRst &"	</soapenv:Body>"
	strRst = strRst &"</soapenv:Envelope>"
	getNvstoremoonbanguOptionSearchParameter = strRst
End Function

'��ǰ ���� XML
Public Function getNvstoremoonbanguDeleteParameter(invstoremoonbangugoodno, iservice, ioperation)
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
	strRst = strRst &"			<ProductId>"&invstoremoonbangugoodno&"</ProductId>"
	strRst = strRst &"		</shop:DeleteProductRequest>"
	strRst = strRst &"	</soapenv:Body>"
	strRst = strRst &"</soapenv:Envelope>"
	getNvstoremoonbanguDeleteParameter = strRst
End Function


'�����ڵ� �� �Ǹ��� �ּҷ�
Public Function getAddressBookList(iccd)
	Dim strRst, oaccessLicense, oTimestamp, osignature, oServ, oOper, reqID
	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		reqID = "ncp_1np6kl_01"
	End If

	oServ		= "AddressBookService"
	Call getsecretKey(oaccessLicense, oTimestamp, osignature, oServ, iccd)
	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
	strRst = strRst & "   <soapenv:Header/>"
	strRst = strRst & "   <soapenv:Body>"
	strRst = strRst & "      <shop:GetAddressBookListRequest>"
	strRst = strRst & "         <shop:RequestID>"&reqID&"</shop:RequestID>"
	strRst = strRst & "         <shop:AccessCredentials>"
	strRst = strRst & "            <shop:AccessLicense>"&oaccessLicense&"</shop:AccessLicense>"
	strRst = strRst & "            <shop:Timestamp>"&oTimestamp&"</shop:Timestamp>"
	strRst = strRst & "            <shop:Signature>"&osignature&"</shop:Signature>"
	strRst = strRst & "         </shop:AccessCredentials>"
	strRst = strRst & "         <shop:Version>2.0</shop:Version>"
	strRst = strRst & "         <SellerId>"&reqID&"</SellerId>"
	strRst = strRst & "      </shop:GetAddressBookListRequest>"
	strRst = strRst & "   </soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	Dim nvstoremoonbanguURL
	If (application("Svr_Info") = "Dev") Then
		nvstoremoonbanguURL = "http://sandbox.api.naver.com/ShopN/"&oServ
	Else
		nvstoremoonbanguURL = "http://ec.api.naver.com/ShopN/"&oServ
	End If

	Dim httpRequest, xmlDOM
	Set httpRequest = CreateObject("MSXML2.XMLHTTP")

	httpRequest.Open "POST", nvstoremoonbanguURL, False
	httpRequest.SetRequestHeader "Content-Type", "text/xml;charset=UTF-8"
	httpRequest.SetRequestHeader "SOAPAction", oServ & "#" & iccd
	httpRequest.send strRst
	If httpRequest.Status = 200 Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(Replace(httpRequest.responseText,"soap:",""))

			response.write (Replace(httpRequest.responseText,"soap:",""))
			response.end
	End If
End Function

'ī�װ� ��ü
Public Function getAllCategoryList(iccd)
	Dim strRst, oaccessLicense, oTimestamp, osignature, oServ, oOper, reqID, ResponseType
	Dim SubNodes, Nodes, strSql
	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		reqID = "ncp_1np6kl_01"
	End If

	oServ		= "ProductService"
	Call getsecretKey(oaccessLicense, oTimestamp, osignature, oServ, iccd)
	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
	strRst = strRst & "   <soapenv:Header/>"
	strRst = strRst & "   <soapenv:Body>"
	strRst = strRst & "      <shop:GetAllCategoryListRequest>"
	strRst = strRst & "         <shop:RequestID>"&reqID&"</shop:RequestID>"
	strRst = strRst & "         <shop:AccessCredentials>"
	strRst = strRst & "            <shop:AccessLicense>"&oaccessLicense&"</shop:AccessLicense>"
	strRst = strRst & "            <shop:Timestamp>"&oTimestamp&"</shop:Timestamp>"
	strRst = strRst & "            <shop:Signature>"&osignature&"</shop:Signature>"
	strRst = strRst & "         </shop:AccessCredentials>"
	strRst = strRst & "         <shop:Version>2.0</shop:Version>"
	strRst = strRst & "         <Last>Y</Last>"
	strRst = strRst & "      </shop:GetAllCategoryListRequest>"
	strRst = strRst & "   </soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	Dim nvstoremoonbanguURL
	If (application("Svr_Info") = "Dev") Then
		nvstoremoonbanguURL = "http://sandbox.api.naver.com/ShopN/"&oServ
	Else
		nvstoremoonbanguURL = "http://ec.api.naver.com/ShopN/"&oServ
	End If

	Dim httpRequest, xmlDOM, objXML
	Dim categoryName, categoryId, cateSplit, i, cateDeplth1Name, cateDeplth2Name, cateDeplth3Name, cateDeplth4Name
	Set objXML = CreateObject("MSXML2.XMLHTTP")
	objXML.Open "POST", nvstoremoonbanguURL, False
	objXML.SetRequestHeader "Content-Type", "text/xml;charset=UTF-8"
	objXML.SetRequestHeader "SOAPAction", oServ & "#" & iccd
	objXML.send strRst
	If objXML.Status = 200 Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
'	rw replace(objXML.responseText, "?xml", "?AAAAAl")
			If ResponseType = "SUCCESS" Then
				If xmlDOM.getElementsByTagName("n:Category").length > 0 Then
					Set Nodes = xmlDOM.getElementsByTagName("n:Category")
						For each SubNodes in Nodes
							categoryName =""
							categoryName	= SubNodes.getElementsByTagName("n:CategoryName")(0).Text

							cateSplit = ""
							cateSplit = Split(categoryName, ">")

							cateDeplth1Name = ""
							cateDeplth2Name = ""
							cateDeplth3Name = ""
							cateDeplth4Name = ""

							cateDeplth1Name = cateSplit(0)
							cateDeplth2Name = cateSplit(1)
							If Ubound(cateSplit) >= 2 Then
								cateDeplth3Name = cateSplit(2)
							End If
							If Ubound(cateSplit) = 3 Then
								cateDeplth4Name = cateSplit(3)
							End If

							categoryId = ""
							categoryId		= SubNodes.getElementsByTagName("n:Id")(0).Text
							strSql = ""
							strSql =  strSql & " IF Exists(SELECT * FROM db_etcmall.dbo.tbl_nvstoremoonbangu_category WHERE catekey = '"&categoryId&"') "
							strSql =  strSql & " 	BEGIN "
							strSql =  strSql & " 		UPDATE db_etcmall.dbo.tbl_nvstoremoonbangu_category SET "
							strSql =  strSql & " 		Depth1Nm = '"&cateDeplth1Name&"' "
							strSql =  strSql & " 		,Depth2Nm = '"&cateDeplth2Name&"' "
							strSql =  strSql & " 		,Depth3Nm = '"&cateDeplth3Name&"' "
							strSql =  strSql & " 		,Depth4Nm = '"&cateDeplth4Name&"' "
							strSql =  strSql & " 		WHERE catekey = '"&categoryId&"' "
							strSql =  strSql & " 	END "
							strSql =  strSql & " ELSE "
							strSql =  strSql & " 	BEGIN "
							strSql =  strSql & " 		INSERT INTO db_etcmall.dbo.tbl_nvstoremoonbangu_category "
							strSql =  strSql & " 		(CateKey, Depth1Nm, Depth2Nm, Depth3Nm, Depth4Nm) "
							strSql =  strSql & " 		VALUES ('"&categoryId&"', '"&cateDeplth1Name&"', '"&cateDeplth2Name&"', '"&cateDeplth3Name&"', '"&cateDeplth4Name&"') "
							strSql =  strSql & " 	END "
							dbget.Execute strSql
						Next
					Set Nodes = nothing
				End If
			End If
			rw "OK"
		Set xmlDOM = nothing
	End If
End Function

'ī�װ� ��
Public Function getCategoryInfo(iccd, iCateId)
	Dim strRst, oaccessLicense, oTimestamp, osignature, oServ, oOper, reqID, ResponseType
	Dim SubNodes, Nodes, strSql
	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		reqID = "ncp_1np6kl_01"
	End If

	oServ		= "ProductService"
	Call getsecretKey(oaccessLicense, oTimestamp, osignature, oServ, iccd)
	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
	strRst = strRst & "   <soapenv:Header/>"
	strRst = strRst & "   <soapenv:Body>"
	strRst = strRst & "      <shop:GetCategoryInfoRequest>"
	strRst = strRst & "         <shop:RequestID>"&reqID&"</shop:RequestID>"
	strRst = strRst & "         <shop:AccessCredentials>"
	strRst = strRst & "            <shop:AccessLicense>"&oaccessLicense&"</shop:AccessLicense>"
	strRst = strRst & "            <shop:Timestamp>"&oTimestamp&"</shop:Timestamp>"
	strRst = strRst & "            <shop:Signature>"&osignature&"</shop:Signature>"
	strRst = strRst & "         </shop:AccessCredentials>"
	strRst = strRst & "         <shop:Version>2.0</shop:Version>"
	strRst = strRst & "         <CategoryId>"&iCateId&"</CategoryId>"
	strRst = strRst & "      </shop:GetCategoryInfoRequest>"
	strRst = strRst & "   </soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	Dim nvstoremoonbanguURL
	If (application("Svr_Info") = "Dev") Then
		nvstoremoonbanguURL = "http://sandbox.api.naver.com/ShopN/"&oServ
	Else
		nvstoremoonbanguURL = "http://ec.api.naver.com/ShopN/"&oServ
	End If

	Dim httpRequest, xmlDOM, objXML
	Dim expCode, expName
	Set objXML = CreateObject("MSXML2.XMLHTTP")
	objXML.Open "POST", nvstoremoonbanguURL, False
	objXML.SetRequestHeader "Content-Type", "text/xml;charset=UTF-8"
	objXML.SetRequestHeader "SOAPAction", oServ & "#" & iccd
	objXML.send strRst
	If objXML.Status = 200 Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
'	rw replace(objXML.responseText, "?xml", "?AAAAAl")
			If ResponseType = "SUCCESS" Then
				If xmlDOM.getElementsByTagName("n:ExceptionalCategoryList").length > 0 Then
					Set Nodes = xmlDOM.getElementsByTagName("n:ExceptionalCategory")
						For each SubNodes in Nodes
							expCode	= SubNodes.getElementsByTagName("n:Code")(0).Text
							expName	= SubNodes.getElementsByTagName("n:Name")(0).Text

							strSql = ""
							strSql =  strSql & " INSERT INTO db_etcmall.[dbo].[tbl_nvstoremoonbangu_certInfo] (CateKey, certCode, certName, lastUpdate) "
							strSql =  strSql & " VALUES ('"&iCateId&"', '"&expCode&"', '"&expName&"', getdate()) "
							dbget.Execute strSql
						Next
						'rw expCode & " | " & expName
					Set Nodes = nothing
				End If
			End If
		Set xmlDOM = nothing
	End If
End Function
%>