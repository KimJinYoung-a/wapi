<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
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

Function sendJandiMgs(istrParam)
	Dim call_url : call_url = "https://wh.jandi.com/connect-api/webhook/15400820/649002d5a799377e4e0ae56d087851d9"
	Dim xmlHttp, SendReqPost
	Set xmlHttp = CreateObject("Msxml2.ServerXMLHTTP.3.0")
		xmlHttp.open "POST", call_url, False
		xmlHttp.setRequestHeader "Accept", "application/vnd.tosslab.jandi-v2+json"
		xmlHttp.setRequestHeader "Content-Type", "application/json"
		xmlHttp.setTimeouts 5000, 60000, 60000, 60000 ''2013/03/14 �߰�
		xmlHttp.Send(istrParam)

		SendReqPost = BinaryToText(xmlHttp.responseBody, "UTF-8")
		rw SendReqPost
	Set xmlHttp = Nothing
	sendJandiMgs = SendReqPost
End Function

Dim sqlStr
Dim vMallid, vMadedt, vEpCnt, vCompletedt
sqlStr = sqlStr & ""
sqlStr = sqlStr & " SELECT TOP 1 mallid, madedt, isNull(epCnt, 0) as epCnt, completedt " & vbcrlf
sqlStr = sqlStr & " FROM db_outmall.[dbo].[tbl_EpShop_Report] " & vbcrlf
sqlStr = sqlStr & " WHERE madedt = CONVERT(Date, GETDATE())  " & vbcrlf
rsCTget.CursorLocation = adUseClient
rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly
IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
	vMallid		= rsCTget("mallid")
	vMadedt		= rsCTget("madedt")
	vEpCnt		= rsCTget("epCnt")
	vCompletedt	= rsCTget("completedt")
END IF
rsCTget.close

Dim obj, strParam
Dim titleStr, descriptionStr
Dim isSucessStr

If vEpCnt > 0 Then
	isSucessStr = "����"
Else
	isSucessStr = "����"
End If

Set obj = jsObject()
	obj("body") = "���̹� ��üEP ���� (" & vMadedt & ")"
	obj("connectColor") = "#FAC11B"
	Set obj("connectInfo")= jsArray()
		titleStr		= "EP���� Report"
		descriptionStr	= "�������� : " & isSucessStr & VBCRLF & "EP �����Ǽ� : " & vEpCnt & VBCRLF & "�Ϸ�ð� : " & vCompletedt
		Set obj("connectInfo")(0) = jsObject()
			obj("connectInfo")(0)("title") = titleStr
			obj("connectInfo")(0)("description") = descriptionStr
			strParam = obj.jsString
Set obj = nothing
Call sendJandiMgs(strParam)
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->