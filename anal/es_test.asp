<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.charset = "utf-8" %>
<!-- #include virtual="/lib/util/aspJSON1.17.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
response.write "TTT"  ''2016/09/19 ELK 오류 // 2016/10/10 //2016/12/12 //2017/01/09 // 2020-07-11(진영)
response.end


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

function debugWrite(iTxt)
    if (application("Svr_Info")	= "Dev") then
        if (request("isautoscript")<>"on") then
            response.write iTxt&"<br>"
        end if
    end if
end function

function getESAPI(stDt,edDt,channel,reqdata,vCode)
    Dim retTxt
    Dim oquery : oquery = getESQuery(stDt,edDt,channel,reqdata,vCode)
 '' response.write "<textarea cols=110 rows=30>"&oquery&"</textarea>"

	retTxt = SendReqPost(getESURI(stDt,edDt,channel), oquery)

	getESAPI = retTxt
end function

Function SendReqPost(call_url, sedata)
    dim xmlHttp
    Set xmlHttp = CreateObject("Msxml2.ServerXMLHTTP.3.0")

    xmlHttp.open "POST",call_url, False
    xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"  ''UTF-8 charset 필요.
    xmlHttp.setTimeouts 5000,60000,60000,60000 ''2013/03/14 추가
    xmlHttp.Send(sedata)

    SendReqPost = BinaryToText(xmlHttp.responseBody, "UTF-8")
    set xmlHttp=Nothing
end function

'function getESURI(stDt,edDt,channel)
'	dim resultSTR
'	getESURI = ""
'
'	Select Case channel
'		Case "Web"
'			getESURI = "http://52.79.58.27:9200/weblog-one-*/_search?pretty"
'		Case "Mob"
'			getESURI = "http://52.79.58.27:9200/moblog-one-*/_search"
'		Case "App"
'			getESURI = "http://52.79.58.27:9200/moblog-one-*/_search"
'		Case Else
'			''
'	End Select
'end function

'' new Kibana v5
function getESURI(stDt,edDt,channel)
	dim resultSTR
	getESURI = ""

	Select Case channel
		Case "Web"
			getESURI = "http://k.tenbyten.kr:9200/weblog-one-*/_search?pretty"
		Case "Mob"
			getESURI = "http://k.tenbyten.kr:9200/moblog-one-*/_search"
		Case "App"
			getESURI = "http://k.tenbyten.kr:9200/moblog-one-*/_search"
		Case Else
			''
	End Select
end function

function getESQuery(stDt,edDt,channel,reqdata,vCode)
	dim resultSTR
	dim start_ts, end_ts
    dim oquery
    dim retMaxRows

	getESQuery = ""

	if (channel = "Web") or (channel = "Mob") or (channel = "App") then

        SELECT CASE reqdata
            CASE "uqipCntPerDay", "uqipCntOrderPerDay", "pageViewPerHour"
                '// KST 시간대는 UTC+9 이다.
        		''start_ts = DateDiff("s", "1970-01-01 00:00:00", Left(DateAdd("d", -1, stDt),10) & " 15:00:00") & "000"
        		''end_ts = DateDiff("s", "1970-01-01 00:00:00", edDt & " 14:59:59") & "999"

        		''CALL debugWrite(stDt)
        		''CALL debugWrite(edDt)

        		start_ts = DateDiff("s", "1970-01-01 00:00:00", DateAdd("h", -9, stDt)) & "000"
        		end_ts = DateDiff("s", "1970-01-01 00:00:00", DateAdd("h", -9, edDt)) & "999"

        		oquery="host : \'$HOST$\' AND (NOT status : 302)"
        		if (reqdata="uqipCntOrderPerDay") then
        		    oquery="host : \'$HOST$\' AND (NOT status : 302) AND 페이지명 : \'주문완료\'"
        		end if

        		if (channel="Mob") then oquery=oquery& " AND (NOT user-agent : (*tenapp*))" ''2016/05/17 추가

        		resultSTR = "{"
                resultSTR = resultSTR & "  'size': 0,"
'                resultSTR = resultSTR & "  'query': {"
'                resultSTR = resultSTR & "    'filtered': {"
'                resultSTR = resultSTR & "      'query': {"
'                resultSTR = resultSTR & "        'query_string': {"
'                resultSTR = resultSTR & "          'query': '"&oquery&"', "
'                resultSTR = resultSTR & "          'analyze_wildcard': true"
'                resultSTR = resultSTR & "        }"
'                resultSTR = resultSTR & "      },"
'                resultSTR = resultSTR & "      'filter': {"
'                resultSTR = resultSTR & "        'bool': {"
'                resultSTR = resultSTR & "          'must': ["
'                resultSTR = resultSTR & "            {"
'                resultSTR = resultSTR & "              'range': {"
'                resultSTR = resultSTR & "                '@timestamp': {"
'                resultSTR = resultSTR & "                  'gte': $START_TS$,"
'                resultSTR = resultSTR & "                  'lte': $END_TS$"
'                resultSTR = resultSTR & "                }"
'                resultSTR = resultSTR & "              }"
'                resultSTR = resultSTR & "            }"
'                resultSTR = resultSTR & "          ],"
'                resultSTR = resultSTR & "          'must_not': []"
'                resultSTR = resultSTR & "        }"
'                resultSTR = resultSTR & "      }"
'                resultSTR = resultSTR & "    }"
'                resultSTR = resultSTR & "  },"

                '' new kibana ------------------------
                resultSTR = resultSTR & "  'query': { "
				resultSTR = resultSTR & "    'bool': { "
				resultSTR = resultSTR & "      'must': [ "
				resultSTR = resultSTR & "        { "
				resultSTR = resultSTR & "          'query_string': { "
				resultSTR = resultSTR & "            'query': '"&oquery&"', "
				resultSTR = resultSTR & "            'analyze_wildcard': true "
				resultSTR = resultSTR & "          } "
				resultSTR = resultSTR & "        }, "
				resultSTR = resultSTR & "        { "
				resultSTR = resultSTR & "          'range': { "
				resultSTR = resultSTR & "            '@timestamp': { "
				resultSTR = resultSTR & "              'gte': $START_TS$, "
				resultSTR = resultSTR & "              'lte': $END_TS$, "
				resultSTR = resultSTR & "              'format': 'epoch_millis' "
				resultSTR = resultSTR & "            } "
				resultSTR = resultSTR & "          } "
				resultSTR = resultSTR & "        } "
				resultSTR = resultSTR & "      ], "
				resultSTR = resultSTR & "      'must_not': [] "
				resultSTR = resultSTR & "    } "
				resultSTR = resultSTR & "  }, "
                '' new kibana ------------------------


                '' new kibana ------------------------
                resultSTR = resultSTR & "  '_source': { "
				resultSTR = resultSTR & "    'excludes': [] "
				resultSTR = resultSTR & "  }, "
				'' new kibana ------------------------

                resultSTR = resultSTR & "  'aggs': {"
                resultSTR = resultSTR & "    'retVal1': {"
                resultSTR = resultSTR & "      'date_histogram': {"
                resultSTR = resultSTR & "        'field': '@timestamp',"
                if (reqdata="pageViewPerHour") then
                    resultSTR = resultSTR & "        'interval': '1h',"
                else
                    resultSTR = resultSTR & "        'interval': '1d',"
                end if

                '' new kibana ------------------------
                resultSTR = resultSTR & "          'time_zone': 'Asia/Tokyo', "
                '' new kibana ------------------------

                ''resultSTR = resultSTR & "        'pre_zone': '+09:00',"
                ''resultSTR = resultSTR & "        'pre_zone_adjust_large_interval': true,"
                resultSTR = resultSTR & "        'min_doc_count': 1,"
                resultSTR = resultSTR & "        'extended_bounds': {"
                resultSTR = resultSTR & "          'min': $START_TS$,"
                resultSTR = resultSTR & "          'max': $END_TS$"
                resultSTR = resultSTR & "        }"
                resultSTR = resultSTR & "      }"



                if (reqdata="pageViewPerHour") then

                else
                    resultSTR = resultSTR & "      ,'aggs': {"
                    resultSTR = resultSTR & "        'retVal2': {"
                    resultSTR = resultSTR & "          'cardinality': {"
                    resultSTR = resultSTR & "            'field': 'clientip.raw'"
                    resultSTR = resultSTR & "          }"
                    resultSTR = resultSTR & "        }"
                    resultSTR = resultSTR & "      }"
                end if
                resultSTR = resultSTR & "    }"
                resultSTR = resultSTR & "  }"
                resultSTR = resultSTR & "}"

        	CASE "page_orderfinish", "page_ref_tmailer","page_item"
        	    '// KST 시간대는 UTC+9 이다.
        	    ''CALL debugWrite(stDt)
        		''CALL debugWrite(edDt)

        		''start_ts = DateDiff("s", "1970-01-01 00:00:00", Left(DateAdd("d", -1, stDt),10) & " 15:00:00") & "000"
        		''end_ts = DateDiff("s", "1970-01-01 00:00:00", LEFT(edDt,10) & " 14:59:59") & "999"

        		''CALL debugWrite(start_ts)
        		''CALL debugWrite(end_ts)

        		start_ts = DateDiff("s", "1970-01-01 00:00:00", DateAdd("h", -9, stDt)) & "000"
        		end_ts = DateDiff("s", "1970-01-01 00:00:00", DateAdd("h", -9, edDt)) & "999"

        		''CALL debugWrite(start_ts)
        		''CALL debugWrite(end_ts)

        		retMaxRows = 500

        		oquery="host : \'$HOST$\' AND (NOT status : 302)"
        		if (reqdata="page_orderfinish") then
        		    oquery="host : \'$HOST$\' AND (NOT status : 302) AND 페이지명 : \'주문상품목록\'"  ''주문완료=>주문상품목록
        		    retMaxRows = 3000
        		elseif (reqdata="page_ref_tmailer") then
        		    oquery="host : \'$HOST$\' AND (NOT status : 302) AND referer_domain:tmailer.10x10.co.kr"
        		    retMaxRows = 10000 '' 10000
        		elseif (reqdata="page_item") then
        		    oquery="host : \'$HOST$\' AND (NOT status : 302) AND 페이지명 : \'상품\' AND 상품코드:"&vCode
        		    retMaxRows = 10000 '' 10000
        		end if

        		if (channel="Mob") then oquery=oquery& " AND (NOT user-agent : (*tenapp*))" ''2016/05/17 추가

        	    resultSTR = "{"
                resultSTR = resultSTR & "  'size': "&retMaxRows&","            '' 0 or N
                resultSTR = resultSTR & "  'sort': ["
                resultSTR = resultSTR & "    {"
                resultSTR = resultSTR & "      '@timestamp': {"
                resultSTR = resultSTR & "        'order': 'desc',"
                resultSTR = resultSTR & "        'unmapped_type': 'boolean'"
                resultSTR = resultSTR & "      }"
                resultSTR = resultSTR & "    }"
                resultSTR = resultSTR & "  ],"

                resultSTR = resultSTR & "  'query': {"
                resultSTR = resultSTR & "    'filtered': {"
                resultSTR = resultSTR & "      'query': {"
                resultSTR = resultSTR & "        'query_string': {"
                resultSTR = resultSTR & "          'query': '"&oquery&"',"
                resultSTR = resultSTR & "          'analyze_wildcard': true"
                resultSTR = resultSTR & "        }"
                resultSTR = resultSTR & "      },"
                resultSTR = resultSTR & "      'filter': {"
                resultSTR = resultSTR & "        'bool': {"
                resultSTR = resultSTR & "          'must': ["
                resultSTR = resultSTR & "            {"
                resultSTR = resultSTR & "              'range': {"
                resultSTR = resultSTR & "                '@timestamp': {"
                resultSTR = resultSTR & "                  'gte': $START_TS$,"  ''1463339855390
                resultSTR = resultSTR & "                  'lte': $END_TS$" ''1463383055390
                resultSTR = resultSTR & "                }"
                resultSTR = resultSTR & "              }"
                resultSTR = resultSTR & "            }"
                resultSTR = resultSTR & "          ],"
                resultSTR = resultSTR & "          'must_not': []"
                resultSTR = resultSTR & "        }"
                resultSTR = resultSTR & "      }"
                resultSTR = resultSTR & "    }"
                resultSTR = resultSTR & "  },"

                resultSTR = resultSTR & "  'fields': ["
                resultSTR = resultSTR & "    '*'"
                resultSTR = resultSTR & "  ],"
                resultSTR = resultSTR & "  'fielddata_fields': ["
                resultSTR = resultSTR & "    '@timestamp'"
                resultSTR = resultSTR & "    ,'clientip'"
                resultSTR = resultSTR & "    ,'request'"

                resultSTR = resultSTR & "    ,'queryparam'"
                resultSTR = resultSTR & "    ,'referer_url'"
                resultSTR = resultSTR & "    ,'referer_queryparam'"

                resultSTR = resultSTR & "    ,'페이지명'"
                resultSTR = resultSTR & "  ]"
                resultSTR = resultSTR & "}"
            CASE ELSE
        		'// KST 시간대는 UTC+9 이다.
        		start_ts = DateDiff("s", "1970-01-01 00:00:00", Left(DateAdd("d", -1, stDt),10) & " 15:00:00") & "000"
        		end_ts = DateDiff("s", "1970-01-01 00:00:00", edDt & " 14:59:59") & "999"

        		resultSTR = "{ "
        		resultSTR = resultSTR & "  'query': { "
        		resultSTR = resultSTR & "    'filtered': { "
        		resultSTR = resultSTR & "      'query': { "
        		resultSTR = resultSTR & "        'query_string': { "
        		resultSTR = resultSTR & "          'query': 'host : \'$HOST$\' AND (NOT status : 302)', "
        		resultSTR = resultSTR & "          'analyze_wildcard': true "
        		resultSTR = resultSTR & "        } "
        		resultSTR = resultSTR & "      }, "
        		resultSTR = resultSTR & "      'filter': { "
        		resultSTR = resultSTR & "        'bool': { "
        		resultSTR = resultSTR & "          'must': [ "
        		resultSTR = resultSTR & "            { "
        		resultSTR = resultSTR & "              'query': { "
        		resultSTR = resultSTR & "                'query_string': { "
        		resultSTR = resultSTR & "                  'query': '*', "
        		resultSTR = resultSTR & "                  'analyze_wildcard': true "
        		resultSTR = resultSTR & "                } "
        		resultSTR = resultSTR & "              } "
        		resultSTR = resultSTR & "            }, "
        		resultSTR = resultSTR & "            { "
        		resultSTR = resultSTR & "              'range': { "
        		resultSTR = resultSTR & "                '@timestamp': { "
        		resultSTR = resultSTR & "                  'gte': $START_TS$, "
        		resultSTR = resultSTR & "                  'lte': $END_TS$ "
        		resultSTR = resultSTR & "                } "
        		resultSTR = resultSTR & "              } "
        		resultSTR = resultSTR & "            } "
        		resultSTR = resultSTR & "          ], "
        		resultSTR = resultSTR & "          'must_not': [] "
        		resultSTR = resultSTR & "        } "
        		resultSTR = resultSTR & "      } "
        		resultSTR = resultSTR & "    } "
        		resultSTR = resultSTR & "  }, "
        		resultSTR = resultSTR & "  'size': 0, "
        		resultSTR = resultSTR & "  'aggs': { "
        		resultSTR = resultSTR & "    'retVal1': { "
        		resultSTR = resultSTR & "      'date_histogram': { "
        		resultSTR = resultSTR & "        'field': '@timestamp', "
        		resultSTR = resultSTR & "        'interval': '1d', "
        		resultSTR = resultSTR & "        'pre_zone': '+09:00', "
        		resultSTR = resultSTR & "        'pre_zone_adjust_large_interval': true, "
        		resultSTR = resultSTR & "        'min_doc_count': 1, "
        		resultSTR = resultSTR & "        'extended_bounds': { "
        		resultSTR = resultSTR & "          'min': $START_TS$, "
        		resultSTR = resultSTR & "          'max': $END_TS$ "
        		resultSTR = resultSTR & "        } "
        		resultSTR = resultSTR & "      }, "
        		resultSTR = resultSTR & "      'aggs': { "
        		resultSTR = resultSTR & "        'retVal2': { "
        		resultSTR = resultSTR & "          'filters': { "
        		resultSTR = resultSTR & "            'filters': { "
        		resultSTR = resultSTR & "              '접속자(IP)': { "
        		resultSTR = resultSTR & "                'query': { "
        		resultSTR = resultSTR & "                  'query_string': { "
        		resultSTR = resultSTR & "                    'query': '*', "
        		resultSTR = resultSTR & "                    'analyze_wildcard': true "
        		resultSTR = resultSTR & "                  } "
        		resultSTR = resultSTR & "                } "
        		resultSTR = resultSTR & "              }, "
        		resultSTR = resultSTR & "              '주문완료': { "
        		resultSTR = resultSTR & "                'query': { "
        		resultSTR = resultSTR & "                  'query_string': { "
        		resultSTR = resultSTR & "                    'query': '페이지명 : \'주문완료\'', "
        		resultSTR = resultSTR & "                    'analyze_wildcard': true "
        		resultSTR = resultSTR & "                  } "
        		resultSTR = resultSTR & "                } "
        		resultSTR = resultSTR & "              } "
        		resultSTR = resultSTR & "            } "
        		resultSTR = resultSTR & "          }, "
        		resultSTR = resultSTR & "          'aggs': { "
        		resultSTR = resultSTR & "            'retVal3': { "
        		resultSTR = resultSTR & "              'cardinality': { "
        		resultSTR = resultSTR & "                'field': 'clientip' "
        		resultSTR = resultSTR & "              } "
        		resultSTR = resultSTR & "            } "
        		resultSTR = resultSTR & "          } "
        		resultSTR = resultSTR & "        } "
        		resultSTR = resultSTR & "      } "
        		resultSTR = resultSTR & "    } "
        		resultSTR = resultSTR & "  } "
        		resultSTR = resultSTR & "}"
        END SELECT

		resultSTR = Replace(resultSTR, "'", """")
		resultSTR = Replace(resultSTR, "$HOST$", channel)
		resultSTR = Replace(resultSTR, "$START_TS$", start_ts)
		resultSTR = Replace(resultSTR, "$END_TS$", end_ts)

		''response.write resultSTR

		getESQuery = resultSTR
	end if
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

function getParseType(reqdata)
    dim i
    for i=LBound(C_ArrOfReqdata) to UBound(C_ArrOfReqdata)
        if (C_ArrOfReqdata(i)=reqdata) then
            getParseType = C_ArrOfParseType(i)
            Exit function
        end if
    next
    getParseType = -1
end function

function UnixTimeToKST(ounixtime,leftN)
    dim v : v = DateAdd("h",9,DateAdd("s",Left(ounixtime,10),#1970/1/1#))
    ''UnixTimeToKST = v
    dim isec
    if (Len(formatDatetime(v,0))=10) then
        isec = ":00"
    else
        isec = Right(formatDatetime(v,0),3)
    end if

    UnixTimeToKST = Left(formatDatetime(v,2)&" "&formatDatetime(v,4)&isec,leftN)
end function

function fnSaveResultSimple(iparseType,channel,reqdata,vTime,vNoVal)
    dim iAnalCon : set iAnalCon = CreateObject("ADODB.Connection")
    Dim strSql

    if (iparseType=2) then
        strSql = "[db_analyze_data_raw].[dbo].[sp_TEN_ELK_data_ADD_simple_byHour] '"&vTime&"','"&reqdata&"','"&channel&"',"&vNoVal
    else
        strSql = "[db_analyze_data_raw].[dbo].[sp_TEN_ELK_data_ADD_simple] '"&vTime&"','"&reqdata&"','"&channel&"',"&vNoVal
    end if
    iAnalCon.Open Application("db_analyze")
    iAnalCon.Execute strSql
    iAnalCon.Close
    SET iAnalCon = Nothing
''  response.write strSql
end function

function fnSaveResultComplex(iparseType,channel,reqdata,vTime,vClientIp,vPageName,vRequestPage,vQueryparam,vReferer_url,vReferer_queryparam,vCode)
    dim iAnalCon : set iAnalCon = CreateObject("ADODB.Connection")
    Dim strSql

    if (iparseType=3) then
        strSql = "[db_analyze_data_raw].[dbo].[sp_TEN_ELK_data_ADD_complex] '"&vTime&"','"&reqdata&"','"&channel&"','"&vClientIp&"','"&vPageName&"','"&vRequestPage&"','"&vQueryparam&"','"&vReferer_url&"','"&vReferer_queryparam&"',"&vCode
    end if
    iAnalCon.Open Application("db_analyze")
    iAnalCon.Execute strSql
    iAnalCon.Close
    SET iAnalCon = Nothing
''  response.write strSql
end function



'''------------------------------------------------------------------------------
dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    response.write ref
    response.end
end if



Dim stDt : stDt=request("stDt") '"2016-05-09"
Dim edDt : edDt=request("edDt") '"2016-05-11"
Dim channel : channel = request("channel")			'// channel = Web or Mob or App
Dim reqdata : reqdata = request("reqdata")          '// test, uqipCntPerDay, uqipCntOrderPerDay, pageViewPerHour
Dim vCode : vCode = request("vCode")
Dim isautoscript : isautoscript = request("isautoscript") ''배치실행.

Dim retTxt

Dim C_ArrOfReqdata : C_ArrOfReqdata = array("test","uqipCntPerDay","uqipCntOrderPerDay","pageViewPerHour","page_orderfinish","page_ref_tmailer","page_item")
Dim C_ArrOfParseType : C_ArrOfParseType = array(9,1,1,2,3,3,3)

if (stDt="") then
    stDt = LEFT(dateadd("d",-1,now()),10)  ''최대 이전 10일 =>3
end if

if (edDt="") then
    edDt = LEFT(now(),10) & " 23:59:59"
end if

if (channel="") then channel = "Web"
if (reqdata="") then reqdata="uqipCntPerDay"


'// ============================================================================

dim iparseType : iparseType = getParseType(reqdata)

if (iparseType=3) then  '' rowData 가져 오는것은 하루치만.
    if (request("research")="") then
        stDt = LEFT(now(),10)
        edDt = LEFT(now(),10)&" 23:59:59"
    end if
end if

'if (isautoscript="on") then
'    if (iparseType="1") or (iparseType="2") then
'        if (stDt="") or (edDt="") then
'            stDt = LEFT(dateadd("d",-1,now()),10)
'            edDt = LEFT(dateadd("d",+1,now()),10)
'        end if
'    end if
'end if

dim i,j,vvv
dim ibucket, bucketItem
dim vUnixTime, vUniqIpNo, vUniqIpOrderNo, vLocaleTime, vQueryparam, vReferer_url, vReferer_queryparam
dim vClientIp, vPageName, vRequestPage, vNoVal
dim otime : otime=Timer()

CALL debugWrite("iparseType:"&iparseType)

retTxt = getESAPI(stDt,edDt,channel,reqdata,vCode)

dim qTime : qTime = FormatNumber(Timer()-otime,4)
otime=Timer()
CALL debugWrite("qTime:"&qTime)

''response.write "<textarea cols=110 rows=30>"&retTxt&"</textarea>"
''response.end

dim oJSON
Dim assignedRow : assignedRow=0
''if (LEFT(reqdata,5)="page_") then  ''page_orderfinish
if (iparseType=3) then '' 데이터가 많음..
    ''아래 파서가 많이 빠름..

    ''retTxt = LEFT(retTxt,10000)
    Set oJSON = JSON.parse(retTxt)

    CALL debugWrite("step2:"&FormatNumber(Timer()-otime,4))
    CALL debugWrite("hits.total:"&oJson.hits.total)
    CALL debugWrite("hits.hits.length:"&oJson.hits.hits.length)
    ''CALL debugWrite("hits-count:"&oJson.hits.hits.get(0).request)

    for i=0 to oJson.hits.hits.length-1
        'CALL debugWrite("oJson.hits.hits.get("&i&").fields.@timestamp:"&oJson.hits.hits.get(i).fields.[@timestamp])
        'CALL debugWrite("oJson.hits.hits.get("&i&").fields.[페이지명]:"&oJson.hits.hits.get(i).fields.[페이지명])
        'CALL debugWrite("oJson.hits.hits.get("&i&").fields.clientip:"&oJson.hits.hits.get(i).fields.[clientip])

        vUnixTime = oJson.hits.hits.get(i).fields.[@timestamp]
        vClientIp = oJson.hits.hits.get(i).fields.[clientip]
        vRequestPage = oJson.hits.hits.get(i).fields.[request]
     on Error resume Next
        vQueryparam = oJson.hits.hits.get(i).fields.[queryparam]
        if Err then vQueryparam=""
     on Error Goto 0

     on Error resume Next
        vReferer_url = oJson.hits.hits.get(i).fields.[referer_url]
        if Err then vReferer_url=""
     on Error Goto 0

     on Error resume Next
        vReferer_queryparam = oJson.hits.hits.get(i).fields.[referer_queryparam]
        if Err then vReferer_queryparam=""
     on Error Goto 0

     on Error resume Next
       vPageName = oJson.hits.hits.get(i).fields.[페이지명]  ''없는경우가 있음..
       if Err then vPageName="UNKNOWN"
     on Error Goto 0

        CALL debugWrite(UnixTimeToKST(vUnixTime,19)&"||"&vClientIp&"||"&vPageName&"||"&vRequestPage&"||"&vQueryparam&"||"&vReferer_url&"||"&vReferer_queryparam)

        if (application("Svr_Info")	= "Dev") then
            if (request("svresult")<>"") then
                Call fnSaveResultComplex(iparseType,channel,reqdata,UnixTimeToKST(vUnixTime,19),vClientIp,vPageName,vRequestPage,vQueryparam,vReferer_url,vReferer_queryparam,vCode)
                assignedRow = assignedRow+1
            end if
        end if
    next


    Set oJSON = Nothing
elseif (iparseType=1) then

    Set oJSON = JSON.parse(retTxt)
    CALL debugWrite("aggregations.retVal1.buckets.length:"&oJson.aggregations.retVal1.buckets.length)

    for i=0 to oJson.aggregations.retVal1.buckets.length-1
        vUnixTime = oJson.aggregations.retVal1.buckets.get(i).[key]
        vNoVal = oJson.aggregations.retVal1.buckets.get(i).[retVal2].value

        CALL debugWrite(UnixTimeToKST(vUnixTime,10)&"||"&vNoVal)

        if (request("svresult")<>"") then
            Call fnSaveResultSimple(iparseType,channel,reqdata,UnixTimeToKST(vUnixTime,10),vNoVal)
            assignedRow = assignedRow+1
        end if
    next

    Set oJSON = Nothing
elseif (iparseType=2) then  ''pageViewPerHour

    Set oJSON = JSON.parse(retTxt)
    CALL debugWrite("aggregations.retVal1.buckets.length:"&oJson.aggregations.retVal1.buckets.length)

    for i=0 to oJson.aggregations.retVal1.buckets.length-1
        vUnixTime = oJson.aggregations.retVal1.buckets.get(i).[key]
        vNoVal = oJson.aggregations.retVal1.buckets.get(i).[doc_count]

        CALL debugWrite(UnixTimeToKST(vUnixTime,19)&"||"&vNoVal)

        if (request("svresult")<>"") then
            Call fnSaveResultSimple(iparseType,channel,reqdata,UnixTimeToKST(vUnixTime,19),vNoVal)
            assignedRow = assignedRow+1
        end if
    next

    Set oJSON = Nothing
else
'// ============================================================================

    Set oJSON = New aspJSON
    oJSON.loadJSON(retTxt)

    CALL debugWrite("step2:"&FormatNumber(Timer()-otime,4))

    ''response.write "bb" & oJSON.data("aggregations").item(0).item(0).count
    ''response.write "bb" & oJSON.data("aggregations/2/buckets").count
    ''response.write "bb" & oJSON.data("aggregations").item("2").item("buckets").count
    '' 'aggs': {'2' 값을 변경하면 그대로 응답함..


    SET vvv =  oJSON.data("aggregations").item("retVal1")
    CALL debugWrite(vvv.item("buckets").count)

    For Each bucketItem In vvv.item("buckets")
        SET ibucket = vvv.item("buckets").item(bucketItem)
        vUnixTime =   ibucket.item("key") ''ibucket.item("key_as_string")

        if (iparseType=9) then
            ''2depth CASE
            ''vUniqIpNo =   ibucket.item("retVal2").item("buckets").item("접속자(IP)").item("retVal3").item("value")
            ''vUniqIpOrderNo =   ibucket.item("retVal2").item("buckets").item("주문완료").item("retVal3").item("value")
        elseif (iparseType=1) then
            ''1depth CASE
            vUniqIpNo =   ibucket.item("retVal2").item("value")

        else

        end if

        CALL debugWrite(UnixTimeToKST(vUnixTime,10)&"||"&vUniqIpNo&"||"&vUniqIpOrderNo)

        SET ibucket = Nothing
    Next

    Set oJSON = Nothing

end if

dim parseTime : parseTime = FormatNumber(Timer()-otime,4)
CALL debugWrite("parseTime:"&parseTime)
%>
<% if (isautoscript="on") then %>
<%
response.write "S_OK|"&assignedRow&"|iparseType="&iparseType&"|stDt="&stDt&"|edDt="&edDt&"|channel="&channel&"|reqdata="&reqdata&"|vCode="&vCode
response.write "<br>"
response.write "qTime="&qTime&"|parseTime="&parseTime&"|pagename=/anal/es_test.asp"
%>
<% else %>
    <script language='javascript'>
        function fnSearch(){
            var frm=document.frmSb;
            frm.submit();
        }

        function reselectThis(comp){
            if (comp.value!=""){
                fnSearch();
                //location.href='?reqdata='+comp.value;
            }
        }
    </script>
    <form name="frmSb" method="get" action="">
    <input type="hidden" name="research" value="on">
    <select name="channel" >
        <option value='Web' <% if channel="Web" then response.write "selected" %> >Web</option>
        <option value='Mob' <% if channel="Mob" then response.write "selected" %> >Mob</option>
        <option value='App' <% if channel="App" then response.write "selected" %> >App</option>
    </select>

    기간:
    <input type="text" name="stDt" value="<%=stDt%>" size="19">~
    <input type="text" name="edDt" value="<%=edDt%>" size="19">

    코드
    <input type="text" name="vCode" value="<%=vCode%>" size="10">

    <select name="reqdata" onChange="reselectThis(this)">
        <option value='uqipCntPerDay' <% if reqdata="uqipCntPerDay" then response.write "selected" %> >접속자수(IP) 일별</option>
        <option value='uqipCntOrderPerDay' <% if reqdata="uqipCntOrderPerDay" then response.write "selected" %> >주문자수(IP) displayorder</option>
        <option value='pageViewPerHour' <% if reqdata="pageViewPerHour" then response.write "selected" %> >pageViewPerHour</option>


        <option value='page_item' <% if reqdata="page_item" then response.write "selected" %> >상품페이지</option>
        <option value='page_orderfinish' <% if reqdata="page_orderfinish" then response.write "selected" %> >주문완료페이지</option>
        <option value='page_ref_tmailer' <% if reqdata="page_ref_tmailer" then response.write "selected" %> >레퍼러:tmailer</option>

    </select>
    <input type="checkbox" name="svresult">결과업데이트
    <input type="button" value="검색" onClick="fnSearch()">
    </form>

    <br>
    <% if (iparseType=3) then %>
    <textarea cols="100" rows="20"><%=LEFT(retTxt,10000)%></textarea>
    <% else %>
    <textarea cols="100" rows="20"><%=retTxt%></textarea>
    <% end if %>
<% end if %>