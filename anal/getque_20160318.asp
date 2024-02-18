<%@ language=vbscript %>
<% option explicit %>
<%
'// Closs Domain Setting 안됨. => jsonp 로 변경 : 응답 callback() 로감쌈.
'Response.AddHeader "Access-Control-Allow-Origin","*"
'Response.AddHeader "Access-Control-Allow-Headers","X-Requested-With"
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "Expires","0"

%>

<%
''<!-- #include virtual="/lib/util/aspJSON1.17.asp" -->

Response.ContentType = "text/html"

dim TMP_check_UserIP : TMP_check_UserIP= request.ServerVariables("REMOTE_ADDR")
dim TMP_check_UserRef : TMP_check_UserRef= LCASE(request.ServerVariables("HTTP_REFERER"))

Dim C_isGroupHeader : C_isGroupHeader= Array("userjoin","meachul","meachulbylevel","mktall","meachulall") ''

dim C_ALLOWIPLIST, C_ALLOWREFERERLIST
C_ALLOWIPLIST = Array(  "115.94.163.42","115.94.163.43","115.94.163.44","115.94.163.45" _
                        ,"61.252.133.2","61.252.133.3","61.252.133.4","61.252.133.5","61.252.133.6" _
                        ,"61.252.133.7","61.252.133.8","61.252.133.9","61.252.133.10","61.252.133.11" _
                        ,"61.252.133.12","61.252.133.13","61.252.133.14","61.252.133.15","61.252.133.16" _
                        ,"61.252.133.17","61.252.133.18","61.252.133.19","61.252.133.20","61.252.133.21" _
                        ,"61.252.133.22","61.252.133.23","61.252.133.24","61.252.133.25","61.252.133.26" _
                        ,"61.252.133.27","61.252.133.28","61.252.133.29","61.252.133.30","61.252.133.31" _
                        ,"61.252.133.32","61.252.133.33","61.252.133.34","61.252.133.35","61.252.133.36" _
                        ,"61.252.133.37","61.252.133.38","61.252.133.39","61.252.133.40","61.252.133.41" _
                        ,"61.252.133.67","61.252.133.68","61.252.133.69","61.252.133.70" _
                        ,"61.252.133.71","61.252.133.72","61.252.133.73","61.252.133.74","61.252.133.75" _
                        ,"61.252.133.76","61.252.133.77","61.252.133.78","61.252.133.79","61.252.133.80" _
                        ,"61.252.133.81","61.252.133.82","61.252.133.83","61.252.133.84","61.252.133.85","61.252.133.86","61.252.133.91" _
                        ,"61.252.133.100","61.252.133.103","61.252.133.104","61.252.133.105","61.252.133.106","61.252.133.107" _
                        ,"61.252.133.113","61.252.133.114","61.252.133.115","61.252.133.116","61.252.133.117","61.252.133.118" _
                        ,"61.252.133.121","61.252.133.122","61.252.133.123","61.252.133.124","61.252.133.125", "61.252.133.92" _
                        ,"52.79.73.177" _
                )

C_ALLOWREFERERLIST = Array ("http://testscm.10x10.co.kr","https://testscm.10x10.co.kr","http://scm.10x10.co.kr","https://scm.10x10.co.kr" )

dim IPCheckOK,REfererCheckOK
dim tmp_ip_i, tmp_ip_buf1, tmp_ref_buf1
IPCheckOK = false
REfererCheckOK = false
for tmp_ip_i=0 to UBound(C_ALLOWIPLIST)
    tmp_ip_buf1 = C_ALLOWIPLIST(tmp_ip_i)
    if (TMP_check_UserIP=tmp_ip_buf1) then
        IPCheckOK = true
        Exit For
    end if
next

if (NOT IPCheckOK) then
    for tmp_ip_i=0 to UBound(C_ALLOWREFERERLIST)
        tmp_ref_buf1 = C_ALLOWREFERERLIST(tmp_ip_i)
        ''if (InStr(TMP_check_UserRef,tmp_ref_buf1)>0) then
        if (LEFT(TMP_check_UserRef,Len(tmp_ref_buf1))=tmp_ref_buf1) then
            REfererCheckOK = true
            Exit For
        end if
    next
    
    if (Not REfererCheckOK) then
        call retErrorJoson("not authentication")
    end if
end if

function getGroupHeaderResult(ikind,isGroupHeader)
    dim i, tmp_buf1
    isGroupHeader=false
    for i=0 to UBound(C_isGroupHeader)
        tmp_buf1 = C_isGroupHeader(i)
        if (ikind=tmp_buf1) then
            isGroupHeader = true
            Exit For
        end if
    next
    
    if (isGroupHeader) then
        getGroupHeaderResult=",""isGroupHeader"":true"
    else
        getGroupHeaderResult=",""isGroupHeader"":false"
    end if

end function

function retErrorJoson(ierrStr)
    dim iretJson
    iretJson = "{""response"":""error"""
    iretJson = iretJson&",""errmsg"":"""&ierrStr&""""
    iretJson = iretJson &"}"
    
    response.write "callback("&iretJson&")"
    'response.write iretJson
    
    response.end
end function

function fnConvFiledTypeTodataTypeStr(iFiledType,byRef iFieldName, byRef bufdataAttr, byRef bufFieldGrpName, byVal isGroupHeader)
    '' adVarChar:200, adInteger:3, adCurrency:6  ''http://www.w3schools.com/asp/prop_field_type.asp
    '' STRING, INTEGER, CURRENCY, PERCENT, FLOATE
    '' IMAGEURL, LINKURL(N), HIDDEN
    dim nPos
    dim iGrpName
    
    bufdataAttr = ""
    select case iFiledType
        case "3","20"
            fnConvFiledTypeTodataTypeStr = "INTEGER"
        case "6"
            fnConvFiledTypeTodataTypeStr = "CURRENCY"   
        case "131"
            fnConvFiledTypeTodataTypeStr = "FLOATE" ''""
        case ELSE
            fnConvFiledTypeTodataTypeStr = "STRING"
    end select
    if LEFT(iFieldName,6)="hidden" then bufdataAttr="HIDDEN"
    if LEFT(iFieldName,8)="IMAGEURL" then bufdataAttr="IMAGEURL"
    if LEFT(iFieldName,5)="LINK(" and InStr(iFieldName,"_")>0 then 
        bufdataAttr=LEFT(iFieldName,InStr(iFieldName,"_")-1)
    end if
    
    iGrpName=""
    nPos = 0
    if (isGroupHeader) and (LEFT(iFieldName,4)="GRP_") then
        nPos = InStr(5,iFieldName,"_")
        if (nPos>0) then
            iGrpName = MID(iFieldName,5,nPos-5)
            iFieldName = MID(iFieldName,5+nPos-5+1,255)
        end if
    else
        iGrpName = iFieldName
    end if
    
    if (isGroupHeader) then  ''그룹헤더인 경우
        bufFieldGrpName = iGrpName    
    else
        bufFieldGrpName = iFieldName
    end if
end function

function getDbDataJson(kind,startdate,enddate,dimensions,channel,param2,param3,pretype,ordtype,shdate,byref iretJson)
    dim iAnalCon : set iAnalCon = CreateObject("ADODB.Connection")
    dim iAnaleRs : set iAnaleRs = CreateObject("ADODB.Recordset")
    dim iField, iFieldCount : iFieldCount=0
    dim i, iVal, totalResults : totalResults=0
    dim isGroupHeader
    dim igroupHeaderResult : igroupHeaderResult = getGroupHeaderResult(kind, isGroupHeader)
    
    dim otime : otime=Timer()   
    dim MAXResult : MAXResult = 1000
    getDbDataJson = false
    
    dim istrSql  
    select CASE kind
        CASE "aaaaa"
            istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_getOrderMData_regdate_day]('"&startdate&"','"&enddate&"')"
        CASE "gadata"
            ''[db_analyze_data_raw].[dbo].[sp_TEN_getGaData] '2015-12-01','2016-01-11','dimensions(dateHour, date, week, month,  year, weekday)','33744458','metrics(sessions,pageviews,users,newUsers)','metricSub','desc'
            if (pretype<>"") then
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_getGaData_withPre]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','"&param2&"','"&ordtype&"',"&MAXResult&",'"&pretype&"',2)"
            else
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_getGaData]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','"&param2&"','"&ordtype&"',"&MAXResult&")"
            end if
            
            ''2016/02/17 전년비교로 강제
            if (param2="sessions") or (param2="pageviews") or (param2="users") or (param2="newusers")   then
                pretype="pyearwd"
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_getGaData_withPre]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','"&param2&"','"&ordtype&"',"&MAXResult&",'"&pretype&"',1)"
            end if
            
        CASE "gadatachannel"
            istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_getGaData_Channel]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','"&param2&"','"&ordtype&"',"&MAXResult&")"
        'CASE "userjoin" ''회원가입
        '    istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_getUserJoinData]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','','"&ordtype&"',"&MAXResult&")"
        CASE "meachul"  ''매출 (TEST)
            istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_getOrderMData]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','"&param2&"','"&ordtype&"','',"&MAXResult&")"
        CASE "meachulall"  ''매출
            if (param2="master") then
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_getOrderMData]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','regdate','"&ordtype&"','',"&MAXResult&")"
            elseif (param2="dispcate") then
                igroupHeaderResult = getGroupHeaderResult("XXX", isGroupHeader) '' 그룹이 아님..
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_get_MD_DispCateSell]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','"&ordtype&"',"&MAXResult&")"
            elseif (param2="brand") then
                igroupHeaderResult = getGroupHeaderResult("XXX", isGroupHeader) '' 그룹이 아님..
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_get_MD_BrandSell]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','"&ordtype&"',"&MAXResult&")"
            elseif (param2="rdsite") then
                igroupHeaderResult = getGroupHeaderResult("XXX", isGroupHeader) '' 그룹이 아님..
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_get_MD_RdsiteSell]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','"&ordtype&"',"&MAXResult&")"
              
                
            else
                CALL retErrorJoson("'"&param2&"' param2 not defined")
            end if
        CASE "meachulbylevel"  ''등급별매출
            istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_getOrderByUserLevel]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','"&param2&"','"&ordtype&"','',"&MAXResult&")"
        CASE "mktall" '' 마케팅관련
            if (param2="gavisit") then
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_get_MKT_GA_metrics]('"&startdate&"','"&enddate&"','users','"&dimensions&"','"&channel&"','"&ordtype&"',"&MAXResult&")"
            elseif (param2="sessions") then
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_get_MKT_GA_metrics]('"&startdate&"','"&enddate&"','sessions','"&dimensions&"','"&channel&"','"&ordtype&"',"&MAXResult&")"
            
            elseif (param2="userjoin") then
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_get_MKT_UserJoin]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','"&ordtype&"',"&MAXResult&")"
            elseif (param2="meachul") then
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_get_MKT_Meachul]('"&startdate&"','"&enddate&"','"&shdate&"','"&dimensions&"','"&channel&"','"&ordtype&"',"&MAXResult&")"
            elseif (param2="meachulbylevel") then
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_get_MKT_UserLevelMeachul]('"&startdate&"','"&enddate&"','','"&dimensions&"','"&channel&"','','"&ordtype&"',"&MAXResult&")"
            elseif (param2="danga") then
                if (pretype<>"") then
                    istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_get_MKT_Danga_WithPre]('"&startdate&"','"&enddate&"','','"&dimensions&"','"&channel&"','"&pretype&"','"&ordtype&"',"&MAXResult&")"
                else
                    istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_get_MKT_Danga]('"&startdate&"','"&enddate&"','','"&dimensions&"','"&channel&"','','"&ordtype&"',"&MAXResult&")"
                end if
            elseif (param2="cr") then
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_get_MKT_CR]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','"&ordtype&"',"&MAXResult&")"
            elseif (param2="dangabylevel") then
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_get_MKT_DangaByLevel]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','"&ordtype&"',"&MAXResult&")"
            elseif (param2="uhold") then
                igroupHeaderResult = getGroupHeaderResult("XXX", isGroupHeader) '' 그룹이 아님..
                istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_get_MKT_UHold]('"&startdate&"','"&enddate&"','"&dimensions&"','"&channel&"','"&ordtype&"',"&MAXResult&")"
           
            else
                CALL retErrorJoson("'"&param2&"' param2 not defined")
            end if
        CASE "bestseller"  ''이걸로 조회시 막힘..? "=>\"로 변환시 json..pjson?
            istrSql = "[db_analyze_data_raw].[dbo].[sp_TEN_getOrder_BestSellItem]('"&startdate&"','"&enddate&"')"
        CASE ELSE
            CALL retErrorJoson("'"&kind&"' kind not defined")
    End SElect
    
    
    ''response.write istrSql
    
    Dim buf, ifieldtype, bufFieldName, bufdataAttr, bufFieldGrpName
    iAnalCon.Open Application("db_analyze") 
    on Error resume Next
    iAnaleRs.Open istrSql, iAnalCon, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    If Err THEN
        call retErrorJoson("ERR:"&Err.description)
        on Error Goto 0
            
    end if
    
    if (iAnaleRs.Fields.count>0) then
        iVal = ",""columnHeaders"":["
        for each iField in iAnaleRs.Fields
           '''"columnHeaders":[{"name":"ga:date","columnType":"DIMENSION","dataType":"STRING"},{"name":"ga:sessions","columnType":"METRIC","dataType":"INTEGER"}]
           buf = iField.Type                '' adVarChar:200, adInteger:3, adCurrency:6  ''http://www.w3schools.com/asp/prop_field_type.asp
           bufFieldName = iField.Name
           
           buf = fnConvFiledTypeTodataTypeStr(buf, bufFieldName, bufdataAttr, bufFieldGrpName, isGroupHeader)
           ''bufFieldName = server.UrlEncode(bufFieldName)  ''??
           
           if (NOT (bufdataAttr="HIDDEN")) then
               iVal = iVal&"{""name"":"""&bufFieldName&""""
               iVal = iVal&",""datatype"":"""&buf&""""
               iVal = iVal&",""dataattr"":"""&bufdataAttr&""""
               if (isGroupHeader) then
                  iVal = iVal&",""groupName"":"""&bufFieldGrpName&"""" 
               end if
               iVal = iVal&"},"
          end if  
          
          iFieldCount = iFieldCount+1
          
          'response.write("Attr:" & prop.Attributes & "<br>")
          'response.write("Name:" & prop.Name & "<br>")
          'response.write("Value:" & prop.Value & "<br>")
        next
        if Right(iVal,1)="," then iVal=Left(iVal,Len(iVal)-1)
        iVal = iVal&"]"
    end if
    if (not iAnaleRs.EOF) then
        iVal = iVal&",""rows"":["
        Do Until iAnaleRs.Eof
            iVal = iVal&"["
            for i=0 to iFieldCount-1
                if NOT (LEFT(iAnaleRs.Fields(i).name,6)="hidden") then
                    ifieldtype=iAnaleRs.Fields(i).Type
                    if (ifieldtype=3 or ifieldtype=6 or ifieldtype=131 or ifieldtype=20) then
                        iVal = iVal&""&iAnaleRs(i)&","
                    else
                        iVal = iVal&""""&replace(iAnaleRs(i),chr(34),"\"&chr(34))&""","   ''doubleQouta
                    end if
                end if
            next
            iAnaleRs.moveNext
            if Right(iVal,1)="," then iVal=Left(iVal,Len(iVal)-1)
            iVal = iVal&"],"
            
            totalResults = totalResults+1
        Loop
        if Right(iVal,1)="," then iVal=Left(iVal,Len(iVal)-1)
        iVal = iVal&"]"
    end if
    iAnaleRs.Close
    iAnalCon.Close
    
    SET iAnaleRs = Nothing
    SET iAnalCon = Nothing
    
    dim iquerys
    iquerys = "{"
    iquerys = iquerys&"""kind"":"""&kind&""","
    iquerys = iquerys&"""startdate"":"""&startdate&""","
    iquerys = iquerys&"""enddate"":"""&enddate&""","
    iquerys = iquerys&"""dimensions"":"""&dimensions&""","
    iquerys = iquerys&"""channel"":"""&channel&""","
    iquerys = iquerys&"""param2"":"""&param2&""","
    iquerys = iquerys&"""param3"":"""&param3&""","
    
    if (right(iquerys,1)=",") then iquerys=Left(iquerys,Len(iquerys)-1)
    iquerys = iquerys&"}"
    
    iretJson = "{""response"":""success"""
    iretJson = iretJson&",""errmsg"":"""&""&""""
    iretJson = iretJson &",""query"":"&iquerys
    iretJson = iretJson &",""maxresults"":"&MAXResult
    iretJson = iretJson &",""totalresults"":"&totalResults
    iretJson = iretJson &igroupHeaderResult
    iretJson = iretJson &iVal
    iretJson = iretJson &"}"
    
    getDbDataJson = true
end function



dim retVal
dim kind      : kind          =   request("kind")
dim startdate  : startdate    =   request("startdate")
dim enddate    : enddate      =   request("enddate")
dim dimensions : dimensions   =   request("dimensions")  ''(dateHour, date, month, year, week, weekday ...)
dim pretype   : pretype     =   request("pretype")  ''(pday, pweek, pmonth, pyear, pyearwd)
dim param1    : param1      =   request("param1") '' gaids, == channel
dim channel   : channel     =   request("channel")
dim param2    : param2      =   request("param2") 
dim param3    : param3      =   request("param3") 
dim ordtype   : ordtype     =   request("ordtype") ''정렬

dim shdate    : shdate = request("shdate")
dim shchannel : shchannel = request("shchannel")
'dim stnum   : stnum         =   request("stnum")
'dim pagesize   : pagesize   =   request("pagesize")

'if (kind="") then retErrorJoson("required kind param")
'if (startdate="") then retErrorJoson("required startdate param")
'if (enddate="") then retErrorJoson("required enddate param")
'if (dimensions="") then retErrorJoson("required dimensions param")


if (channel="") and (param1<>"") then channel=param1 ''구버전
if (shchannel<>"") then channel=shchannel
        
Dim iretJson

if (getDbDataJson(kind,startdate,enddate,dimensions,channel,param2,param3,pretype,ordtype,shdate,iretJson)) then
    '' response.write iretJson
    response.write "callback("&iretJson&")"
end if


%>
