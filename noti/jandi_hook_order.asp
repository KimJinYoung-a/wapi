<%@  codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Response.CharSet="UTF-8"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/util/aspJSON1.17.asp" -->
<%
''response
'{
'    "body" : "[[PizzaHouse]](http://url_to_text) You have a new Pizza order.",
'    "connectColor" : "#FAC11B",
'    "connectInfo" : [
'    {
'        "title" : "Topping",
'        "description" : "Pepperoni"
'    },
'    {
'        "title": "Location",
'        "description": "Empire State Building, 5th Ave, New York",
'    }
'    ]
'}


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

function sendJandiMgs(sedata)
    dim call_url : call_url = "https://wh.jandi.com/connect-api/webhook/15400820/72565878422057bd00faabc3c0e85454"
    dim xmlHttp, SendReqPost
    Set xmlHttp = CreateObject("Msxml2.ServerXMLHTTP.3.0")
    
    xmlHttp.open "POST",call_url, False
    xmlHttp.setRequestHeader "Accept", "application/vnd.tosslab.jandi-v2+json"  
    xmlHttp.setRequestHeader "Content-Type", "application/json"  
    
    xmlHttp.setTimeouts 5000,60000,60000,60000 ''2013/03/14 추가
    xmlHttp.Send(sedata)
    
    SendReqPost = BinaryToText(xmlHttp.responseBody, "UTF-8")
    set xmlHttp=Nothing
    
    sendJandiMgs = SendReqPost
end function


function IsValidToken(itoken, byref itokenIdx, byref iaddparam, keyword, text)
    dim i
    IsValidToken = False
    itokenIdx = -1
    for i=LBound(g_TokenArray) to UBound(g_TokenArray)
        if (LCASE(g_TokenArray(i))=LCASE(itoken)) then
            IsValidToken = True
            itokenIdx = i
            iaddparam = Trim(replace(text,"/"&keyword,""))
            Exit function
        end if
    next
end function


Dim g_TokenArray : g_TokenArray = Array("6fe2cb9ed68937bf194b194ae50092b8","TTTTTT")

dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

'dim token : token = request("token")
'dim teamName : teamName = request("teamName")
'dim roomName : roomName = request("roomName")
'dim writerName : writerName = request("writerName")
'dim text : text = request("text")
'dim keyword : keyword = request("keyword")
'dim createdAt : createdAt = request("createdAt")

dim token, teamName, roomName, writerName, text, keyword
dim rcvData , lngBytesCount
dim oJSON
If (Request.TotalBytes > 0) Then
    lngBytesCount = Request.TotalBytes
    rcvData = BinaryToText(Request.BinaryRead(lngBytesCount),"utf-8")
    
    Set oJSON = New aspJSON
    oJSON.loadJSON(rcvData)
    token = oJSON.data("token")
    teamName = oJSON.data("teamName")
    roomName = oJSON.data("roomName")
    writerName = oJSON.data("writerName")
    text = oJSON.data("text")
    keyword = oJSON.data("keyword")
    Set oJSON = Nothing
    
else
    rcvData = "TTT"
End If

''token = "6fe2cb9ed68937bf194b194ae50092b8"

Dim iTokenIdx, iaddparam
if Not IsValidToken(token,iTokenIdx,iaddparam,keyword,text) then
    dbget.close()
    response.end
end if

dim addparam1
if (iaddparam<>"") then
    if IsNumeric(iaddparam) then
        addparam1 = iaddparam
    end if
end if

dim sqlStr, ArrRows
dim retData, i
dim itemid,itemname,ttlordcnt,ttlcnt,itemcostsum, avgprice
if (iTokenIdx=0) then  ''6fe2cb9ed68937bf194b194ae50092b8 : 베스트셀러
    if (addparam1="") then
        addparam1 = "3"
    end if
    sqlStr = "db_order.[dbo].[usp_Ten_RecentBestSellItem]("&addparam1&")"
    rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    IF not rsget.EOF THEN
    	ArrRows = rsget.getRows()
    END IF
    rsget.close
    
    retData = "{"
    retData = retData & """body"" : ""[최근 "&addparam1&"시간 자사몰 베스트 셀러]"","
    retData = retData & """connectColor"" : ""#FAC11B"","
    retData = retData & """connectInfo"" : ["
    if IsArray(ArrRows) then
        For i=0 To UBound(ArrRows,2)
            itemid = ArrRows(0,i)
            itemname = ArrRows(1,i)
            ttlordcnt = ArrRows(2,i)
            ttlcnt = ArrRows(3,i)
            itemcostsum = ArrRows(4,i)
            
            avgprice = 0 
            if ttlcnt<>0 then avgprice = itemcostsum/ttlcnt
            
            itemname = replace(itemname,"""","")
            retData = retData & "	{"
            retData = retData & "	""title"" : ""["&itemid&"] "&itemname&""","
            retData = retData & "	""description"" : ""주문수 "&ttlordcnt&" : 판매수 "&ttlcnt&" : 평단가 : "&formatNumber(avgprice,0)&""","
            retData = retData & "	""imageUrl"": ""http://www.10x10.co.kr/"&itemid&""""
            retData = retData & "	}"
            
            if i<>UBound(ArrRows,2) then
                retData = retData & ","
            end if
                
        Next
    end if
    retData = retData & "]"
    retData = retData & "}"
    
    
end if

response.write retData


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->