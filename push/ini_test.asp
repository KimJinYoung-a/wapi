<%
'//바이너리 데이터 TEXT형태로 변환
	Function  BinaryToText(BinaryData, CharSet)
		 Const adTypeText = 2
		 Const adTypeBinary = 1

		 Dim BinaryStream
		 Set BinaryStream = CreateObject("ADODB.Stream")

		'원본 데이터 타입
		 BinaryStream.Type = adTypeBinary

		 BinaryStream.Open
		 BinaryStream.Write BinaryData
		 ' binary -> text
		 BinaryStream.Position = 0
		 BinaryStream.Type = adTypeText

		' 변환할 데이터 캐릭터셋
		 BinaryStream.CharSet = CharSet

		'변환한 데이터 반환
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