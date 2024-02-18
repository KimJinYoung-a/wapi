<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.charset = "utf-8" %>
<!-- #include virtual="/lib/util/aspJSON1.17.asp" -->
<%
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("192.168.1.70","192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

function getGaAPI(ids,stDt,edDt,metrics,dimensions,iGaToken)
    ''https://www.googleapis.com/analytics/v3/data/ga?ids=ga%3A85556678&start-date=2014-01-01&end-date=2016-01-10&metrics=ga%3Asessions&dimensions=ga%3Adate&access_token=ya29.aQKhg7UQcqeEP8QqulQvzGgA5b6XjQg6a2G8o5v5dIXeWlglpZk65TYsJH7t6hzdCSoGIA
    Dim QurURL : QurURL = "https://www.googleapis.com/analytics/v3/data/ga"
    Dim OueParam : OueParam = "ids="&ids&"&start-date="&stDt&"&end-date="&edDt&"&metrics="&metrics&"&dimensions="&dimensions&"&access_token="&iGaToken
    Dim retTxt
    
    'response.write QurURL&"?"
    'response.write OueParam&"<br>"
    'response.write "getGaToken:"&getGaToken&"<br>"
    
    if (iGaToken<>"") then
        retTxt = SendReqGet(QurURL,OueParam)
        getGaAPI =retTxt
    end if
end function

function getGaToken()
    Dim retTxt, oJSON, itoken
    retTxt = SendReqGet("http://ec2-52-79-73-177.ap-northeast-2.compute.amazonaws.com:8081/token","")
    
    Set oJSON = New aspJSON

    ''response.write retTxt
    'Load JSON string
    oJSON.loadJSON(retTxt)
    itoken = oJSON.data("token")
    Set oJSON = Nothing
    
    getGaToken = itoken
end function

function AddGaData(idt,gaid,metrics,metricSub,ival)
    
    dim iAnalCon : set iAnalCon = CreateObject("ADODB.Connection")
    Dim strSql
    
    strSql = "[db_analyze_data_raw].[dbo].[sp_TEN_ga_data_ADD] '"&idt&"','"&gaid&"','"&metrics&"','"&metricSub&"',"&ival
    iAnalCon.Open Application("db_analyze")
    iAnalCon.Execute strSql
    iAnalCon.Close
    SET iAnalCon = Nothing
    
    
end function

Function SendReqGet(call_url, sedata)
    dim igetURL
    igetURL = call_url
    if sedata<>"" then igetURL = igetURL&"?"&sedata
    
    dim xmlHttp
    Set xmlHttp = CreateObject("Msxml2.ServerXMLHTTP.3.0")
    
    xmlHttp.open "GET",igetURL, False
    xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"  ''UTF-8 charset 필요.
    xmlHttp.setTimeouts 5000,30000,30000,30000 ''2013/03/14 추가  30000 으로 변경
    xmlHttp.Send
    
    SendReqGet = BinaryToText(xmlHttp.responseBody, "UTF-8")
    set xmlHttp=Nothing
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



dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    response.write ref
    response.end
end if


Dim iGaToken
iGaToken = getGaToken
''response.end

'' 사용자/신규사용자의 경우 하루씩 조회 하는것과 기간별 조회한것의 값이 다름
Dim stDt : stDt=request("stDt") '"2014-06-01"
Dim edDt : edDt=request("edDt") '"2014-06-30"
Dim ids : ids=request("ids")  '' 33744458-www,  48997153 - mob,  85556678-app , 82072655-app(mod) 48996365-fingers ,,85559666(X) 비트윈.
Dim idsArray
Dim metrics : metrics = request("metrics") ''ga:sessions, pageviews, users, newUsers //, percentNewSessions
Dim dimensions : dimensions = "ga:date" ''dateHour, date, week, month,  year, weekday
Dim metricSub : metricSub =""
Dim retTxt

if (metrics="") then metrics = "ga:pageviews"  
    
if (stDt="") then
    stDt = LEFT(dateadd("d",-10,now()),10)  ''최대 이전 10일
end if

if (edDt="") then
    edDt = LEFT(now(),10)
end if

if (ids="") then 
    idsArray = Array("ga:33744458","ga:48997153","ga:85556678","ga:82072655")
else
    idsArray = Array(ids)
end if    

''idsArray = Array("ga:33744458")

if (metrics="ga:sessions") or (metrics="ga:pageviews") then
    dimensions = "ga:dateHour"
else
    dimensions = "ga:date"
end if

Dim i,j
Dim sqlStr
Dim oJSON
Dim totalResults
Dim idatetime, iVal
Dim resText
for j=LBound(idsArray) to UBound(idsArray)
    ids = idsArray(j)
    retTxt = getGaAPI(ids,stDt,edDt,metrics,dimensions,iGaToken)
    
    
    Set oJSON = New aspJSON
    
    
    oJSON.loadJSON(retTxt)
    totalResults = 0
    totalResults = oJSON.data("totalResults")
    
    for i=0 to totalResults-1
        idatetime = oJSON.data("rows").item(i).item(0)
        iVal      = oJSON.data("rows").item(i).item(1)
        if (dimensions = "ga:dateHour") then
            idatetime = LEFT(idatetime,4)&"-"&Mid(idatetime,5,2)&"-"&Mid(idatetime,7,2)&" "&Mid(idatetime,9,2)&":00:00"
        else
            idatetime = LEFT(idatetime,4)&"-"&Mid(idatetime,5,2)&"-"&Mid(idatetime,7,2)&" 00:00:00"
        end if
        'response.write idatetime
        
        Call AddGaData(idatetime,replace(ids,"ga:",""),replace(metrics,"ga:",""),metricSub,ival)
    Next
    
    Set oJSON = Nothing


    resText = resText & "ids:"&ids&"_metrics:"&metrics&"_dimensions:"&dimensions&"_TTL:"&totalResults&"("&stDt&"~"&edDt&")"&"<br>"
Next

response.write resText
%>