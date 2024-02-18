<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<%
''https://wh.jandi.com/connect-api/webhook/15400820/56bd0ac3e6db69b6b4aea837894d9469
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
    if (fnIsLocalDev) then
        CheckVaildIP = true
        Exit function
    end if

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
    dim call_url : call_url = "https://wh.jandi.com/connect-api/webhook/15400820/56bd0ac3e6db69b6b4aea837894d9469"
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
sqlStr = "db_order.[dbo].[sp_Ten_CHK_order_SellCost_ByHook]"
rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
IF not rsget.EOF THEN
	ArrRows = rsget.getRows()
END IF
rsget.close

dim i, sn
dim sitename,itemid,sellitemcost, preSellcash,ttlcnt,makerid,ttlordcnt,mwdiv,dismargin, sttime
''itemid	sellitemcost	preSellcash	ttlcnt	makerid	mwdiv	dismargin
''371635	5300.00	198000.00	2	poog	U	97.33

dim descMsg, titleMsg, is3plMall
dim MaxLoopN : MaxLoopN=5
if IsArray(ArrRows) then
    For i=0 To UBound(ArrRows,2)
        ''sitename = ArrRows(0,i)
        itemid = ArrRows(0,i)
        sellitemcost = ArrRows(1,i)
        preSellcash = ArrRows(2,i)
        ttlcnt = ArrRows(3,i)
        ttlordcnt = ArrRows(4,i)
        makerid = ArrRows(5,i)
        mwdiv = ArrRows(6,i)
        dismargin = ArrRows(7,i)
        
        ''is3plMall = ((LEFT(sitename,3)="its") or (sitename="ithinksoshop"))
        
        ''titleMsg = "site:"&sitename&" ("&sttime&"~)"
        descMsg = "주문건수:"&ttlordcnt& " 주문수량:"&formatNumber(ttlcnt,0)
        if ttlcnt<>0 then 
            descMsg = descMsg & " 판매가:"&formatNumber(sellitemcost,0)&" / 이전판매가:"&formatNumber(preSellcash,0)

            if (preSellcash<>0) then
                descMsg = descMsg & " ("&CLNG((preSellcash-sellitemcost)/preSellcash*100)&"%)"
            end if

            descMsg = descMsg&"\r\n"
            descMsg = descMsg&"브랜드:"&makerid&" 매입구분:"&mwdiv
        end if
        
        sedata = "{"
        sedata = sedata & "'body': '[판매가체크] "&itemid&"( http://www.10x10.co.kr/"&itemid&" )',"
        sedata = sedata & "'connectColor': '#FAC11B',"
        sedata = sedata & "'connectInfo': ["
        sedata = sedata & "{"
        if (titleMsg<>"") then
            sedata = sedata & "'title': '"&titleMsg&"',"
        end if
        if (descMsg<>"") then
            sedata = sedata & "'description': '"&descMsg&"'"
        end if
        sedata = sedata & "}"
        'sedata = sedata & ",{"
        'sedata = sedata & "'title': '타이틀2',"
        'sedata = sedata & "'description': '디스크립션2',"
        ''sedata = sedata & "'imageUrl': 'http://webimage.10x10.co.kr/image/list/195/L001959793.jpg'"
        'sedata = sedata & "}"
        sedata = sedata & "]"
        sedata = sedata & "}"
        
        sedata = replace(sedata,"'","""")
        
        if (sn>=MaxLoopN) then Exit For
        
       
        ''response.write sedata
        response.write sendJandiMgs(sedata)
        sn = sn+1
    
    
    next
end if


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->