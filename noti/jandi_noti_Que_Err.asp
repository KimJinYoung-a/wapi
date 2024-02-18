<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbDatamartOpen.asp" -->
<!-- #include virtual="/lib/db/dbOpen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%
Dim ref : ref = Request.ServerVariables("REMOTE_ADDR")
If (Not CheckVaildIP(ref)) Then
	'dbDatamart_dbget.Close()
	'response.write "접속이 불가능한 IP 입니다." : response.ed
End If

Dim sedata
Dim sqlStr, ArrRows
' sqlStr = sqlStr & ""
' sqlStr = sqlStr & " SELECT TOP 100 itemid, itemname " & vbcrlf
' sqlStr = sqlStr & " FROM db_item.dbo.tbl_item " & vbcrlf
' dbDatamart_rsget.Open sqlStr,dbDatamart_dbget,1
' IF not dbDatamart_rsget.EOF THEN
' 	ArrRows = dbDatamart_rsget.getRows()
' END IF
' dbDatamart_rsget.close
If (application("Svr_Info") = "Dev") Then
	sqlStr = sqlStr & ""
	sqlStr = sqlStr & " SELECT TOP 10 mallId, apiAction, lastErrMsg, 37 as cnt, itemid " & vbcrlf
	sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_outmall_API_Que " & vbcrlf
	rsget.Open sqlStr,dbget,1
	IF not rsget.EOF THEN
		ArrRows = rsget.getRows()
	END IF
	rsget.close
Else
	sqlStr = "EXEC [db_datamart].[dbo].[usp_Ten_Que_Err_Alarm]"
	dbDatamart_rsget.CursorLocation = adUseClient
	dbDatamart_rsget.CursorType = adOpenStatic
	dbDatamart_rsget.LockType = adLockOptimistic
	dbDatamart_rsget.Open sqlStr, dbDatamart_dbget
	If Not(dbDatamart_rsget.EOF or dbDatamart_rsget.BOF) Then
		ArrRows = dbDatamart_rsget.getRows()
	End If
	dbDatamart_rsget.close
End If

Dim i
Dim obj, strParam
Dim titleStr
Dim descriptionStr

If IsArray(ArrRows) then
	Set obj = jsObject()
		obj("body") = "10회이상 반복 오류 건 (3시간 전)"
		obj("connectColor") = "#FAC11B"
		Set obj("connectInfo")= jsArray()

		For i = 0 To UBound(ArrRows, 2)
			titleStr = ""
			descriptionStr = ""
			titleStr		= ArrRows(2, i) & " ("& ArrRows(3, i) &"건)"
			descriptionStr	= "몰ID : " & ArrRows(0, i) & VBCRLF & "Action : " & ArrRows(1, i) & VBCRLF & "상품코드 : " &  ArrRows(4, i)

			Set obj("connectInfo")(i) = jsObject()
				obj("connectInfo")(i)("title") = titleStr
				obj("connectInfo")(i)("description") = descriptionStr
		Next
		strParam = obj.jsString
	Set obj = nothing
	Call sendJandiMgs(strParam)
End If

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.73","61.252.133.70","61.252.133.10","61.252.133.80","110.93.128.114","110.93.128.113","192.168.1.73","192.168.1.70")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

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

Function sendJandiMgs(sedata)
	Dim call_url : call_url = "https://wh.jandi.com/connect-api/webhook/15400820/80337c466f6ae711d4d074aa4a783eab"
	Dim xmlHttp, SendReqPost
	Set xmlHttp = CreateObject("Msxml2.ServerXMLHTTP.3.0")
		xmlHttp.open "POST", call_url, False
		xmlHttp.setRequestHeader "Accept", "application/vnd.tosslab.jandi-v2+json"
		xmlHttp.setRequestHeader "Content-Type", "application/json"
		xmlHttp.setTimeouts 5000,60000,60000,60000 ''2013/03/14 추가
		xmlHttp.Send(sedata)
		SendReqPost = BinaryToText(xmlHttp.responseBody, "UTF-8")
	Set xmlHttp=Nothing
	sendJandiMgs = SendReqPost
End Function
%>
<!-- #include virtual="/lib/db/dbDatamartClose.asp" -->
<!-- #include virtual="/lib/db/dbClose.asp" -->