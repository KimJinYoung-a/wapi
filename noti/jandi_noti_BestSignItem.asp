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

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.70","61.252.133.10","61.252.133.80","110.93.128.114","110.93.128.113","192.168.1.70")
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
    dim call_url : call_url = "https://wh.jandi.com/connect-api/webhook/15400820/1fdbb7c3b49277b0e1ea871ee41928a1"
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
    dbget.Close()
    response.write "nonono"
    response.end
end if

dim sedata
dim sqlStr, ArrRows
sqlStr = "[db_analyze_data_raw].[dbo].[usp_Ten_Sign_best_item_get]"
rsAnalget.Open sqlStr,dbAnalget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
IF not rsAnalget.EOF THEN
	ArrRows = rsAnalget.getRows()
END IF
rsAnalget.close

dim i, sn
dim itemid, itemname, makerid, sellSTDate, yyyymmdd, sellOrdCNT, sellRnk, favCnt, favRnk, baguniSum, baguniRnk
dim descMsg, titleMsg

''itemid	itemname	makerid	sellSTDate	yyyymmdd	sellOrdCNT	sellRnk	favCnt	favRnk	baguniSum	baguniRnk	regdate	zozimscore
''1999929	오난코리아 루메나 N9-BK 모기퇴치기	n9	2018-06-08 17:11:48.940	2018-06-14	1999929	38	6	170	1	332	1	2018-06-15 20:15:09.700	1062

if IsArray(ArrRows) then
    
    For i=0 To UBound(ArrRows,2)
        itemid      = ArrRows(0,i)
        itemname    = ArrRows(1,i)
        makerid     = ArrRows(2,i)
        sellSTDate  = ArrRows(3,i)
        yyyymmdd    = ArrRows(4,i)
        sellOrdCNT  = ArrRows(5,i)
        sellRnk     = ArrRows(6,i)
        favCnt      = ArrRows(7,i)
        favRnk      = ArrRows(8,i)
        baguniSum   = ArrRows(9,i)
        baguniRnk   = ArrRows(10,i)
        
        titleMsg = "전일 ("&LEFT(CStr(dateadd("d",-1,now())),10)&") 조짐(Sign) 상품 (최근7일등록상품) (http://www.10x10.co.kr/"&itemid&")"
        descMsg = ""
        descMsg = descMsg & "{"
        descMsg = descMsg & "'title': '"&CStr(i+1)&". 브랜드ID :"&makerid&" | 판매시작일 :"&sellSTDate& "',"
        descMsg = descMsg & "'description': '주문수 :"&sellOrdCNT&"("&sellRnk&" 위) | 위시수 :"&favCnt&"("&favRnk&" 위) | 장바구니 :"&baguniSum&"("&baguniRnk&" 위) '"
        'descMsg = descMsg & "'imageUrl': 'http://www.10x10.co.kr/search/search_result.asp?rect="&rect&"'"
        
        descMsg = descMsg & "}"
        
        
    
        
        sedata = "{"
        sedata = sedata & "'body': '["&titleMsg&"]',"
        sedata = sedata & "'connectColor': '#FAC11B',"
        sedata = sedata & "'connectInfo': ["
        sedata = sedata & descMsg

        sedata = sedata & "]"
        sedata = sedata & "}"
        
        sedata = replace(sedata,"'","""")
        
        if (sedata<>"") then
            'response.write sedata&"<br>"
            response.write sendJandiMgs(sedata)
        end if
   next
end if


%>
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->