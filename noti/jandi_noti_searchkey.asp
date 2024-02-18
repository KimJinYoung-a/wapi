<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<%
''https://wh.jandi.com/connect-api/webhook/15400820/72565878422057bd00faabc3c0e85454
''Accept : application/vnd.tosslab.jandi-v2+json
''Content-Type : application/json

'{
'  "body": "[[PizzaHouse]](http://url_to_text) You have a new Pizza order.",
'  "connectColor": "#FAC11B",
'  "connectInfo": [
'    {
'      "title": "Topping",
'      "description": "Pepperoni"
'    },
'    {
'      "title": "Location",
'      "description": "Empire State Building, 5th Ave, New York",
'      "imageUrl": "http://url_to_text"
'    }
'  ]
'}

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.70","61.252.133.10","61.252.133.80","110.93.128.114","110.93.128.113","192.168.1.70","121.78.103.60")
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

function sendJandiMgs(sedata)
    dim call_url : call_url = "https://wh.jandi.com/connect-api/webhook/15400820/f65d84e45181a85f26a67fdd6492c1d2"
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

dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    dbAnalget.Close()
    response.write "nonono"
    response.end
end if

dim sedata
dim sqlStr, ArrRows
sqlStr = "db_Evt.[dbo].[usp_Ten_Keyword_ZoomUpList]"
rsAnalget.Open sqlStr,dbAnalget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
IF not rsAnalget.EOF THEN
	ArrRows = rsAnalget.getRows()
END IF
rsAnalget.close

dim i, sn
dim rect, searchCnt, Rnk, pRnk, RnkUp, mxrectCnt
dim descMsg, titleMsg

if IsArray(ArrRows) then
    titleMsg = "전일["&LEFT(CStr(dateadd("d",-1,now())),10)&"] 검색 상승 키워드 (내부검색 기준, -14일 비교)"
    descMsg = ""
    For i=0 To UBound(ArrRows,2)
        rect = ArrRows(0,i)
        searchCnt = ArrRows(1,i)
        Rnk = ArrRows(2,i)
        pRnk = ArrRows(3,i)
        RnkUp = ArrRows(4,i)
        mxrectCnt = ArrRows(5,i)
        
        descMsg = descMsg & "{"
        descMsg = descMsg & "'title': '"&CStr(i+1)&". "& rect& " (+"&RnkUp&")',"
        descMsg = descMsg & "'description': '검색결과수 :"&mxrectCnt&"   |   랭킹 :"&Rnk& "',"
        descMsg = descMsg & "'imageUrl': 'http://www.10x10.co.kr/search/search_result.asp?rect="&rect&"'"
        
        descMsg = descMsg & "}"
        if (i<>UBound(ArrRows,2)) then
            descMsg = descMsg & ","
        end if
        
    next
        
        sedata = "{"
        sedata = sedata & "'body': '["&titleMsg&"]',"
        sedata = sedata & "'connectColor': '#FAC11B',"
        sedata = sedata & "'connectInfo': ["
        sedata = sedata & descMsg
'        sedata = sedata & "{"
'        if (titleMsg<>"") then
'            sedata = sedata & "'title': '"&titleMsg&"',"
'        end if
'        if (descMsg<>"") then
'            sedata = sedata & "'description': '"&descMsg&"'"
'        end if
'        sedata = sedata & "}"
        'sedata = sedata & ",{"
        'sedata = sedata & "'title': '타이틀2',"
        'sedata = sedata & "'description': '디스크립션2',"
        ''sedata = sedata & "'imageUrl': 'http://webimage.10x10.co.kr/image/list/195/L001959793.jpg'"
        'sedata = sedata & "}"
        sedata = sedata & "]"
        sedata = sedata & "}"
        
        sedata = replace(sedata,"'","""")
        
        if (descMsg<>"") then
            response.write sendJandiMgs(sedata)
        end if
end if


%>
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->