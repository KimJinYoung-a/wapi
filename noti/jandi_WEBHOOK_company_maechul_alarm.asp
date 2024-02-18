<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbDatamartOpen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<%
dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    dbDatamart_dbget.Close()
    response.write "접속이 불가능한 IP 입니다." : response.end
end if

dim sedata
dim sqlStr, ArrRows
sqlStr = "select top 1 yyyymmdd_alarm, yyyymmdd_realmaechul, messageStr" & vbcrlf
sqlStr = sqlStr & " from db_datamart.dbo.tbl_company_maechul_alarm_log with (nolock)" & vbcrlf
sqlStr = sqlStr & " where yyyymmdd_alarm='"& date() &"'" & vbcrlf

'response.write sqlStr & "<Br>"
dbDatamart_rsget.Open sqlStr,dbDatamart_dbget,1
IF not dbDatamart_rsget.EOF THEN
	ArrRows = dbDatamart_rsget.getRows()
END IF
dbDatamart_rsget.close

dim i
dim yyyymmdd_alarm, messageStr
dim descMsg, titleMsg
if IsArray(ArrRows) then
    yyyymmdd_alarm = ArrRows(0,i)
    messageStr = ArrRows(2,i)

    titleMsg = "[텐바이텐] 매출보고"
    descMsg = messageStr

    sedata = "{"
    sedata = sedata & "'body': '[텐바이텐] 매출보고',"
    sedata = sedata & "'connectColor': '#FAC11B',"
    sedata = sedata & "'connectInfo': ["
    sedata = sedata & "{"
    if (titleMsg<>"") then
        sedata = sedata & "'title': '"&titleMsg&"',"
    end if
    if (descMsg<>"") then
        sedata = sedata & "'description': '"& replace(descMsg,vbcrlf,"\n") &"'"
    end if
    sedata = sedata & "}"
    sedata = sedata & "]"
    sedata = sedata & "}"
    
    sedata = replace(sedata,"'","""")
'	response.write sedata
'	response.end

    ' [텐바이텐] 매출보고
    response.write sendJandiMgs1(sedata)

    ' [GSR]실적공유
    response.write sendJandiMgs2(sedata)
end if

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

function sendJandiMgs1(sedata)
    dim call_url : call_url = "https://wh.jandi.com/connect-api/webhook/15400820/9e467d41afa8c054e3f7af31fe4309ed"
    dim xmlHttp, SendReqPost
    Set xmlHttp = CreateObject("Msxml2.ServerXMLHTTP.3.0")
    
    xmlHttp.open "POST",call_url, False
    xmlHttp.setRequestHeader "Accept", "application/vnd.tosslab.jandi-v2+json"  
    xmlHttp.setRequestHeader "Content-Type", "application/json"  
    
    xmlHttp.setTimeouts 5000,60000,60000,60000 ''2013/03/14 추가
    xmlHttp.Send(sedata)
    
    SendReqPost = BinaryToText(xmlHttp.responseBody, "UTF-8")
    set xmlHttp=Nothing
    
    sendJandiMgs1 = SendReqPost
end function

function sendJandiMgs2(sedata)
    dim call_url : call_url = "https://wh.jandi.com/connect-api/webhook/15400820/29e01b82a5071990567db098fccb3e35"
    dim xmlHttp, SendReqPost
    Set xmlHttp = CreateObject("Msxml2.ServerXMLHTTP.3.0")
    
    xmlHttp.open "POST",call_url, False
    xmlHttp.setRequestHeader "Accept", "application/vnd.tosslab.jandi-v2+json"  
    xmlHttp.setRequestHeader "Content-Type", "application/json"  
    
    xmlHttp.setTimeouts 5000,60000,60000,60000 ''2013/03/14 추가
    xmlHttp.Send(sedata)
    
    SendReqPost = BinaryToText(xmlHttp.responseBody, "UTF-8")
    set xmlHttp=Nothing
    
    sendJandiMgs2 = SendReqPost
end function
%>
<!-- #include virtual="/lib/db/dbDatamartClose.asp" -->