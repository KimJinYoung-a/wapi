<%
'//���̳ʸ� ������ TEXT���·� ��ȯ
	Function  BinaryToText(BinaryData, CharSet)
		 Const adTypeText = 2
		 Const adTypeBinary = 1

		 Dim BinaryStream
		 Set BinaryStream = CreateObject("ADODB.Stream")

		'���� ������ Ÿ��
		 BinaryStream.Type = adTypeBinary

		 BinaryStream.Open
		 BinaryStream.Write BinaryData
		 ' binary -> text
		 BinaryStream.Position = 0
		 BinaryStream.Type = adTypeText

		' ��ȯ�� ������ ĳ���ͼ�
		 BinaryStream.CharSet = CharSet

		'��ȯ�� ������ ��ȯ
		 BinaryToText = BinaryStream.ReadText

		 Set BinaryStream = Nothing
	End Function
	
Dim xmlHttp, vntPostedData, url, postdata
url = "https://fcmobile.inicis.com/smart/pay_req_url.php"
''url = "http://fcmobile.inicis.com/smart/pay_req_url.php"

'' Set xmlHttp = server.CreateObject("Microsoft.XMLHTTP") ''Msxml2
Set xmlHttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0") ''.3.0

    xmlHttp.Open "POST", url, False
	xmlHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xmlHttp.Send postdata
	
	vntPostedData = BinaryToText(xmlHttp.responseBody, "euc-kr")
Set xmlHttp = Nothing

response.write vntPostedData
%>