<%
Dim C_ArrOfReqdata   : C_ArrOfReqdata = array("test" _
                                              ,"best_keywords" _
                                              ,"uqipCntPerDay","uqipCntOrderPerDay" _
                                              ,"pageViewPerHour" _
                                              ,"page_exists_uk","page_orderfinish","page_orderbaguni","page_ref_tmailer","page_item" _
                                              ,"page_orderbaguni_app" _
                                              )
Dim C_ArrOfParseType : C_ArrOfParseType = array(9 _
                                              ,0 _
                                              ,1,1 _
                                              ,2 _
                                              ,3,3,3,3,3 _
                                              ,4 _
                                              )


function debugWrite(iTxt)
    if (application("Svr_Info")	= "Dev") then
        if (request("isautoscript")<>"on") then
            response.write iTxt&"<br>"
        end if
    end if
end function

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

function GtmToKST(ogtmtime,leftN)
    dim ret : ret=ogtmtime
    
    GtmToKST = ret
    
    if Instr(ogtmtime,"T")<1 then exit function
    if Instr(ogtmtime,"Z")<1 then exit function
    
    dim v : v = DateAdd("h",9,replace(LEFT(ret,leftN),"T"," "))
    
    dim isec
    
    if (Len(formatDatetime(v,0))=10) then
        isec = ":00"
    else
        isec = Right(formatDatetime(v,0),3)
    end if

    GtmToKST = Left(formatDatetime(v,2)&" "&formatDatetime(v,4)&isec,leftN)
end function


function getRequestJoson(ichannel, istart_ts, iend_ts, iquery, isize, iinterval, isearchField, isUinqClientIP)
    Dim resultSTR
    resultSTR = "{"&vbCRLF
    resultSTR = resultSTR & "  'query': {"&vbCRLF
    resultSTR = resultSTR & "    'filtered': {"&vbCRLF
    resultSTR = resultSTR & "      'query': {"&vbCRLF
    resultSTR = resultSTR & "        'query_string': {"&vbCRLF
    resultSTR = resultSTR & "          'query': '"&iquery&"',"&vbCRLF
    resultSTR = resultSTR & "          'analyze_wildcard': true"&vbCRLF
    resultSTR = resultSTR & "        }"&vbCRLF
    resultSTR = resultSTR & "      },"&vbCRLF
    resultSTR = resultSTR & "      'filter': {"&vbCRLF
    resultSTR = resultSTR & "        'bool': {"&vbCRLF
    resultSTR = resultSTR & "          'must': ["&vbCRLF
    resultSTR = resultSTR & "            {"&vbCRLF
    resultSTR = resultSTR & "              'range': {"&vbCRLF
    resultSTR = resultSTR & "                '@timestamp': {"&vbCRLF
    resultSTR = resultSTR & "                  'gte': $START_TS$,"&vbCRLF
    resultSTR = resultSTR & "                  'lte': $END_TS$"&vbCRLF
    resultSTR = resultSTR & "                }"&vbCRLF
    resultSTR = resultSTR & "              }"&vbCRLF
    resultSTR = resultSTR & "            }"&vbCRLF
    resultSTR = resultSTR & "          ],"&vbCRLF
    resultSTR = resultSTR & "          'must_not': []"&vbCRLF
    resultSTR = resultSTR & "        }"&vbCRLF
    resultSTR = resultSTR & "      }"&vbCRLF
    resultSTR = resultSTR & "    }"&vbCRLF
    resultSTR = resultSTR & "  }"&vbCRLF
    resultSTR = resultSTR & "  ,'size': 0"&vbCRLF          '' 0 or N
    if (iinterval<>"") then
      resultSTR = resultSTR & ",'aggs': {"&vbCRLF
      resultSTR = resultSTR & "  'retValInterval': {"&vbCRLF
      resultSTR = resultSTR & "    'date_histogram': {"&vbCRLF
      resultSTR = resultSTR & "      'field': '@timestamp',"&vbCRLF
      resultSTR = resultSTR & "      'interval': '1d'," &vbCRLF                        ''iinterval 1d, 1h
      resultSTR = resultSTR & "      'pre_zone': '+09:00',"&vbCRLF
      resultSTR = resultSTR & "      'pre_zone_adjust_large_interval': true,"&vbCRLF
      resultSTR = resultSTR & "      'min_doc_count': 1,"&vbCRLF
      resultSTR = resultSTR & "      'extended_bounds': {"&vbCRLF
      resultSTR = resultSTR & "        'min': $START_TS$,"&vbCRLF
      resultSTR = resultSTR & "        'max': $END_TS$"&vbCRLF
      resultSTR = resultSTR & "      }"&vbCRLF
      resultSTR = resultSTR & "    }"&vbCRLF
      resultSTR = resultSTR & " }"&vbCRLF
    end if
    resultSTR = resultSTR & "  ,'aggs': {"&vbCRLF
    resultSTR = resultSTR & "    'retVal2': {"&vbCRLF
    resultSTR = resultSTR & "      'terms': {"&vbCRLF
    resultSTR = resultSTR & "        'field': '"&isearchField&"'," &vbCRLF
    resultSTR = resultSTR & "        'size': "&isize&","&vbCRLF
    resultSTR = resultSTR & "        'order': {"&vbCRLF
    resultSTR = resultSTR & "          '1': 'desc'"&vbCRLF
    resultSTR = resultSTR & "        }"&vbCRLF
    resultSTR = resultSTR & "      }"&vbCRLF
    if (isUinqClientIP) then
        resultSTR = resultSTR & "      ,'aggs': {"&vbCRLF
        resultSTR = resultSTR & "        '1': {"&vbCRLF
        resultSTR = resultSTR & "          'cardinality': {"&vbCRLF
        resultSTR = resultSTR & "            'field': 'clientip'"&vbCRLF
        resultSTR = resultSTR & "          }"&vbCRLF
        resultSTR = resultSTR & "        }"&vbCRLF
        resultSTR = resultSTR & "      }"&vbCRLF
    end if
    resultSTR = resultSTR & "    }"&vbCRLF
    resultSTR = resultSTR & "  }"&vbCRLF
    if (iinterval<>"") then
        resultSTR = resultSTR & "  }"&vbCRLF
    end if
    resultSTR = resultSTR & "}"

    dim bufChannel
	if (ichannel="FWb") then
	    bufChannel="Web"
    elseif (ichannel="FMb") then
		    bufChannel="Mob"
	else
	    bufChannel = ichannel
	end if

    resultSTR = Replace(resultSTR, "'", """")
	resultSTR = Replace(resultSTR, "$HOST$", bufChannel)
	resultSTR = Replace(resultSTR, "$START_TS$", istart_ts)
	resultSTR = Replace(resultSTR, "$END_TS$", iend_ts)

    getRequestJoson = resultSTR
end function

'' OLD version
function getESQuery(stDt,edDt,channel,reqdata,vCode)
	dim resultSTR
	dim start_ts, end_ts
    dim oquery
    dim retMaxRows
    dim isize, iinterval, isearchField, isUinqClientIP
    dim bufChannel

	getESQuery = ""

	if (channel = "Web") or (channel = "Mob") or (channel = "App") or (channel = "FWb") or (channel = "FMb") then

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
        		if (reqdata="uqipCntOrderPerDay") then                  ''주문완료
        		    if (channel = "FWb") then       ''핑거스웹
        		        oquery="host : \'$HOST$\' AND (NOT status : 302) AND (request:\'/lecpay/applyresult.asp\' OR request:\'/lecpay/DIYresultOrder.asp\')"
        		    elseif (channel = "FMb") then   ''핑거스 모바일
        		        oquery="host : \'$HOST$\' AND (NOT status : 302) AND (request:\'/lecpay/applyresult.asp\' OR request:\'/lecpay/DIYapplyresult.asp\')"
        		    else
        		        oquery="host : \'$HOST$\' AND (NOT status : 302) AND 페이지명 : \'주문완료\'"
        		    end if
        		end if

        		if (channel="Mob") then oquery=oquery& " AND (NOT user-agent : (*tenapp*))" ''2016/05/17 추가

        		resultSTR = "{"
                resultSTR = resultSTR & "  'size': 0,"
                resultSTR = resultSTR & "  'query': {"
                resultSTR = resultSTR & "    'filtered': {"
                resultSTR = resultSTR & "      'query': {"
                resultSTR = resultSTR & "        'query_string': {"
                resultSTR = resultSTR & "          'query': '"&oquery&"', "
                resultSTR = resultSTR & "          'analyze_wildcard': true"
                resultSTR = resultSTR & "        }"
                resultSTR = resultSTR & "      },"
                resultSTR = resultSTR & "      'filter': {"
                resultSTR = resultSTR & "        'bool': {"
                resultSTR = resultSTR & "          'must': ["
                resultSTR = resultSTR & "            {"
                resultSTR = resultSTR & "              'range': {"
                resultSTR = resultSTR & "                '@timestamp': {"
                resultSTR = resultSTR & "                  'gte': $START_TS$,"
                resultSTR = resultSTR & "                  'lte': $END_TS$"
                resultSTR = resultSTR & "                }"
                resultSTR = resultSTR & "              }"
                resultSTR = resultSTR & "            }"
                resultSTR = resultSTR & "          ],"
                resultSTR = resultSTR & "          'must_not': []"
                resultSTR = resultSTR & "        }"
                resultSTR = resultSTR & "      }"
                resultSTR = resultSTR & "    }"
                resultSTR = resultSTR & "  },"
                resultSTR = resultSTR & "  'aggs': {"
                resultSTR = resultSTR & "    'retVal1': {"
                resultSTR = resultSTR & "      'date_histogram': {"
                resultSTR = resultSTR & "        'field': '@timestamp',"
                if (reqdata="pageViewPerHour") then
                    resultSTR = resultSTR & "        'interval': '1h',"
                else
                    resultSTR = resultSTR & "        'interval': '1d',"
                end if
                resultSTR = resultSTR & "        'pre_zone': '+09:00',"
                resultSTR = resultSTR & "        'pre_zone_adjust_large_interval': true,"
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
                    resultSTR = resultSTR & "            'field': 'clientip'"
                    resultSTR = resultSTR & "          }"
                    resultSTR = resultSTR & "        }"
                    resultSTR = resultSTR & "      }"
                end if
                resultSTR = resultSTR & "    }"
                resultSTR = resultSTR & "  }"
                resultSTR = resultSTR & "}"
            
            CASE "page_exists_uk", "page_orderfinish", "page_orderbaguni", "page_ref_tmailer","page_item"
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

        		'' 302를 무조건 빼면 안됨..
        		oquery="host : \'$HOST$\' AND (NOT status : 302)"
        		if (reqdata="page_orderfinish") then
        		    oquery="host : \'$HOST$\' AND (NOT status : 302) AND 페이지명 : \'주문완료\'"  ''주문완료=>주문상품목록=>주문완료
        		    retMaxRows = 5000
        		elseif (reqdata="page_orderbaguni") then
        		    oquery="host : \'$HOST$\' AND (NOT status : 302) AND 페이지명 : \'장바구니담기\'"  ''주문완료=>주문상품목록=>주문완료
        		    retMaxRows = 5000
        		elseif (reqdata="page_ref_tmailer") then
        		    oquery="host : \'$HOST$\' AND (NOT status : 302) AND referer_domain:tmailer.10x10.co.kr"
        		    retMaxRows = 10000 '' 10000
        		elseif (reqdata="page_item") then
        		    oquery="host : \'$HOST$\' AND (NOT status : 302) AND 페이지명 : \'상품\' AND 상품코드:"&vCode
        		    retMaxRows = 10000 '' 10000
        		elseif (reqdata="page_exists_uk") then  ''약 2시간 단위 적당.
        		    ''oquery="host : \'$HOST$\' AND (_exists_:queryparam_uk OR _exists_:queryparam_dumi)"  ''uk 파람 존재, dumi 추가(2016/10/05)
        		    if (channel="Web") then  '' 2016/10/07 clientip:211.204.81.236  - bot 으로 되어 있음?
        		        oquery="(_exists_:queryparam_uk OR _exists_:queryparam_dumi)"  ''uk 파람 존재, dumi 추가(2016/10/05)
        		    else
        		        oquery="host : \'$HOST$\' AND (_exists_:queryparam_uk OR _exists_:queryparam_dumi)"  ''uk 파람 존재, dumi 추가(2016/10/05)
        		    end if
        		    retMaxRows = 10000
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

				resultSTR = resultSTR & "  '_source': ["
                resultSTR = resultSTR & "      '@timestamp'"
				resultSTR = resultSTR & "       ,'clientip'"
				resultSTR = resultSTR & "       ,'request'"
				resultSTR = resultSTR & "       ,'queryparam'"
				resultSTR = resultSTR & "       ,'referer_url'"
				resultSTR = resultSTR & "       ,'referer_queryparam'"
				resultSTR = resultSTR & "       ,'페이지명'"
                resultSTR = resultSTR & "    ]"
                
                				
'				resultSTR = resultSTR & "  'docvalue_fields' : "
'				resultSTR = resultSTR & "   'clientip.raw'"
'				resultSTR = resultSTR & "   ,'request.raw'"
'				resultSTR = resultSTR & " ]"
				

                  
'				resultSTR = resultSTR & "  'stored_fields': ["
'                resultSTR = resultSTR & "      '*',"
'                resultSTR = resultSTR & "      '_source'"
'                resultSTR = resultSTR & "    ]"
'                
'				resultSTR = resultSTR & "  '_source':  "
'				''resultSTR = resultSTR & "    'excludes': [] "
'				resultSTR = resultSTR & "   ["
'				resultSTR = resultSTR & "       '@timestamp':{'format': 'epoch_millis'}"
'				resultSTR = resultSTR & "       ,'clientip'"
'				resultSTR = resultSTR & "       ,'request'"
'				resultSTR = resultSTR & "       ,'queryparam'"
'				resultSTR = resultSTR & "       ,'referer_url'"
'				resultSTR = resultSTR & "       ,'referer_queryparam'"
'				resultSTR = resultSTR & "       ,'페이지명'"
'				resultSTR = resultSTR & "   ]"
'				resultSTR = resultSTR & "  "
				
               


'                resultSTR = resultSTR & "  'stored_fields': ["
'                resultSTR = resultSTR & "    '*'"
'                resultSTR = resultSTR & "   ,'_source'"
'                resultSTR = resultSTR & "  ],"
''                resultSTR = resultSTR & " 'script_fields': {},"
''                resultSTR = resultSTR & "  'fielddata': true, "
'                resultSTR = resultSTR & "  'fielddata_fields': ["
'                resultSTR = resultSTR & "    '@timestamp'"
'                resultSTR = resultSTR & "    ,'clientip'"
'                resultSTR = resultSTR & "    ,'request'"
'
'                resultSTR = resultSTR & "    ,'queryparam'"
'                resultSTR = resultSTR & "    ,'referer_url'"
'                resultSTR = resultSTR & "    ,'referer_queryparam'"
'
'                resultSTR = resultSTR & "    ,'페이지명'"
'                resultSTR = resultSTR & "  ]"
                resultSTR = resultSTR & "}"
                
            '' 이전 키바나 (AWS)
        	CASE "OLD_page_exists_uk", "OLD_page_orderfinish", "OLD_page_orderbaguni", "OLD_page_ref_tmailer","OLD_page_item"
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

        		'' 302를 무조건 빼면 안됨..
        		oquery="host : \'$HOST$\' AND (NOT status : 302)"
        		if (reqdata="page_orderfinish") then
        		    oquery="host : \'$HOST$\' AND (NOT status : 302) AND 페이지명 : \'주문완료\'"  ''주문완료=>주문상품목록=>주문완료
        		    retMaxRows = 5000
        		elseif (reqdata="page_orderbaguni") then
        		    oquery="host : \'$HOST$\' AND (NOT status : 302) AND 페이지명 : \'장바구니담기\'"  ''주문완료=>주문상품목록=>주문완료
        		    retMaxRows = 5000
        		elseif (reqdata="page_ref_tmailer") then
        		    oquery="host : \'$HOST$\' AND (NOT status : 302) AND referer_domain:tmailer.10x10.co.kr"
        		    retMaxRows = 10000 '' 10000
        		elseif (reqdata="page_item") then
        		    oquery="host : \'$HOST$\' AND (NOT status : 302) AND 페이지명 : \'상품\' AND 상품코드:"&vCode
        		    retMaxRows = 10000 '' 10000
        		elseif (reqdata="page_exists_uk") then  ''약 2시간 단위 적당.
        		    ''oquery="host : \'$HOST$\' AND (_exists_:queryparam_uk OR _exists_:queryparam_dumi)"  ''uk 파람 존재, dumi 추가(2016/10/05)
        		    if (channel="Web") then  '' 2016/10/07 clientip:211.204.81.236  - bot 으로 되어 있음?
        		        oquery="(_exists_:queryparam_uk OR _exists_:queryparam_dumi)"  ''uk 파람 존재, dumi 추가(2016/10/05)
        		    else
        		        oquery="host : \'$HOST$\' AND (_exists_:queryparam_uk OR _exists_:queryparam_dumi)"  ''uk 파람 존재, dumi 추가(2016/10/05)
        		    end if
        		    retMaxRows = 10000
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
       '         resultSTR = resultSTR & "  'highlight': {"
       '         resultSTR = resultSTR & "    'pre_tags': ["
       '         resultSTR = resultSTR & "      '@kibana-highlighted-field@'"
       '         resultSTR = resultSTR & "    ],"
       '         resultSTR = resultSTR & "    'post_tags': ["
       '         resultSTR = resultSTR & "      '@/kibana-highlighted-field@'"
       '         resultSTR = resultSTR & "    ],"
       '         resultSTR = resultSTR & "    'fields': {"
       '         resultSTR = resultSTR & "      '*': {}"
       '         resultSTR = resultSTR & "    },"
       '         resultSTR = resultSTR & "    'fragment_size': 2147483647"
       '         resultSTR = resultSTR & "  },"
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
'                resultSTR = resultSTR & "  'aggs': {"
'                resultSTR = resultSTR & "    '2': {"
'                resultSTR = resultSTR & "      'date_histogram': {"
'                resultSTR = resultSTR & "        'field': '@timestamp',"
'                resultSTR = resultSTR & "        'interval': '10m',"
'                resultSTR = resultSTR & "        'pre_zone': '+09:00',"
'                resultSTR = resultSTR & "        'pre_zone_adjust_large_interval': true,"
'                resultSTR = resultSTR & "        'min_doc_count': 0,"
'                resultSTR = resultSTR & "        'extended_bounds': {"
'                resultSTR = resultSTR & "          'min': 1463339855390,"
'                resultSTR = resultSTR & "          'max': 1463383055390"
'                resultSTR = resultSTR & "        }"
'                resultSTR = resultSTR & "      }"
'                resultSTR = resultSTR & "    }"
'                resultSTR = resultSTR & "  },"
                resultSTR = resultSTR & "  'fields': ["
                resultSTR = resultSTR & "    '*'"
           '     resultSTR = resultSTR & "    ,'_source'"
                resultSTR = resultSTR & "  ],"
'                resultSTR = resultSTR & "  'script_fields': {},"
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

            CASE "page_orderbaguni_app"
            	start_ts = DateDiff("s", "1970-01-01 00:00:00", DateAdd("h", -9, stDt)) & "000"
        		end_ts = DateDiff("s", "1970-01-01 00:00:00", DateAdd("h", -9, edDt)) & "999"

				''response.write start_ts
				''response.write end_ts
        		retMaxRows = 500

        		oquery="host : \'$HOST$\' AND !status : 302 AND 페이지명 : \'장바구니담기\' AND os: (\'iOS\' OR \'Android\') "

				resultSTR = " { "
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
				''resultSTR = resultSTR & "              'gte': 1487635200000, "
				''resultSTR = resultSTR & "              'lte': 1487652161319, "
				resultSTR = resultSTR & "              'format': 'epoch_millis' "
				resultSTR = resultSTR & "            } "
				resultSTR = resultSTR & "          } "
				resultSTR = resultSTR & "        } "
				resultSTR = resultSTR & "      ], "
				resultSTR = resultSTR & "      'must_not': [] "
				resultSTR = resultSTR & "    } "
				resultSTR = resultSTR & "  }, "
				resultSTR = resultSTR & "  'size': 0, "
				resultSTR = resultSTR & "  '_source': { "
				resultSTR = resultSTR & "    'excludes': [] "
				resultSTR = resultSTR & "  }, "
				resultSTR = resultSTR & "    'aggs': { "
				resultSTR = resultSTR & "      '2': { "
				resultSTR = resultSTR & "        'date_histogram': { "
				resultSTR = resultSTR & "          'field': '@timestamp', "
				resultSTR = resultSTR & "          'interval': '1h', "
				resultSTR = resultSTR & "          'time_zone': 'Asia/Tokyo', "
				resultSTR = resultSTR & "          'min_doc_count': 1 "
				resultSTR = resultSTR & "        }, "
				resultSTR = resultSTR & "        'aggs': { "
				resultSTR = resultSTR & "          '3': { "
				resultSTR = resultSTR & "            'terms': { "
				resultSTR = resultSTR & "              'field': 'os.raw', "
				resultSTR = resultSTR & "              'size': 5, "
				resultSTR = resultSTR & "              'order': { "
				resultSTR = resultSTR & "                '_count': 'desc' "
				resultSTR = resultSTR & "              } "
				resultSTR = resultSTR & "            }, "
				resultSTR = resultSTR & "            'aggs': { "
				resultSTR = resultSTR & "              '5': { "
				resultSTR = resultSTR & "                'terms': { "
				resultSTR = resultSTR & "                  'field': '상품코드.raw', "
				resultSTR = resultSTR & "                  'size': 100, "
				resultSTR = resultSTR & "                  'order': { "
				resultSTR = resultSTR & "                    '_count': 'desc' "
				resultSTR = resultSTR & "                  } "
				resultSTR = resultSTR & "                }, "
				resultSTR = resultSTR & "                'aggs': { "
				resultSTR = resultSTR & "                  '6': { "
				resultSTR = resultSTR & "                    'terms': { "
				resultSTR = resultSTR & "                      'field': '즉시구매.raw', "
				resultSTR = resultSTR & "                      'size': 5, "
				resultSTR = resultSTR & "                      'order': { "
				resultSTR = resultSTR & "                        '_count': 'desc' "
				resultSTR = resultSTR & "                      } "
				resultSTR = resultSTR & "                    } "
				resultSTR = resultSTR & "                  } "
				resultSTR = resultSTR & "                } "
				resultSTR = resultSTR & "              } "
				resultSTR = resultSTR & "            } "
				resultSTR = resultSTR & "          } "
				resultSTR = resultSTR & "        } "
				resultSTR = resultSTR & "      } "
				resultSTR = resultSTR & "    } "
				resultSTR = resultSTR & "  } "
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

        if (reqdata<>"best_keywords") then
    		if (channel="FWb") then
    		    bufChannel="Web"
    		elseif (channel="FMb") then
    		    bufChannel="Mob"
    		else
    		    bufChannel = channel
    		end if

    	    resultSTR = Replace(resultSTR, "'", """")
    	    resultSTR = Replace(resultSTR, "$HOST$", bufChannel)
    	    resultSTR = Replace(resultSTR, "$START_TS$", start_ts)
    	    resultSTR = Replace(resultSTR, "$END_TS$", end_ts)
	    end if

		getESQuery = resultSTR
	end if
end function


function getESAPI(stDt,edDt,channel,reqdata,vCode)
    Dim retTxt
    dim reqQry

    Dim istart_ts, iend_ts, iquery, isize, iinterval, isearchField, isUinqClientIP

    SELECT CASE reqdata
        CASE "best_keywords"
            istart_ts   = DateDiff("s", "1970-01-01 00:00:00", DateAdd("h", -9, stDt)) & "000"
    		iend_ts     = DateDiff("s", "1970-01-01 00:00:00", DateAdd("h", -9, edDt)) & "999"

    		iquery="host : \'$HOST$\' AND (NOT status : 302)"
    		iquery=iquery&" AND request:\'/search/search_result.asp\' AND queryparam_cpg : \'1\'"

    		isize = 100
    		if (channel = "FWb") or (channel = "FMb") then isize = 20
    		iinterval = "" '"1d"
    		isearchField = "검색키워드"
    		isUinqClientIP = true

    		reqQry = getRequestJoson(channel, istart_ts, iend_ts, iquery, isize, iinterval, isearchField, isUinqClientIP)
        CASE ELSE ''구코드
            reqQry = getESQuery(stDt,edDt,channel,reqdata,vCode)
    END SELECT

    ''response.write "<textarea cols=""100"" rows=""20"">"&LEFT(reqQry,10000)&"</textarea><br>"

	retTxt = SendReqPost(getESURI(stDt,edDt,channel), reqQry)

	getESAPI = retTxt
end function

function getESAPI_NEW(stDt,edDt,channel,reqdata,vCode)
    Dim retTxt
    dim reqQry

    Dim istart_ts, iend_ts, iquery, isize, iinterval, isearchField, isUinqClientIP

    SELECT CASE reqdata
        CASE "best_keywords"
            istart_ts   = DateDiff("s", "1970-01-01 00:00:00", DateAdd("h", -9, stDt)) & "000"
    		iend_ts     = DateDiff("s", "1970-01-01 00:00:00", DateAdd("h", -9, edDt)) & "999"

    		iquery="host : \'$HOST$\' AND (NOT status : 302)"
    		iquery=iquery&" AND request:\'/search/search_result.asp\' AND queryparam_cpg : \'1\'"

    		isize = 100
    		if (channel = "FWb") or (channel = "FMb") then isize = 20
    		iinterval = "" '"1d"
    		isearchField = "검색키워드"
    		isUinqClientIP = true

    		reqQry = getRequestJoson(channel, istart_ts, iend_ts, iquery, isize, iinterval, isearchField, isUinqClientIP)
        CASE ELSE ''구코드
            reqQry = getESQuery(stDt,edDt,channel,reqdata,vCode)
    END SELECT

 '' response.write "<textarea cols=""100"" rows=""20"">"&LEFT(reqQry,10000)&"</textarea><br>"

	retTxt = SendReqPost(getESURI_NEW(stDt,edDt,channel), reqQry)

	getESAPI_NEW = retTxt
end function

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("192.168.1.70","192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72","61.252.133.67")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

function getESURI(stDt,edDt,channel)
	dim resultSTR
	getESURI = ""

	Select Case channel
		Case "Web"
			getESURI = "http://52.79.58.27:9200/weblog-one-*/_search?pretty"
		Case "Mob"
			getESURI = "http://52.79.58.27:9200/moblog-one-*/_search"
		Case "App"
			getESURI = "http://52.79.58.27:9200/moblog-one-*/_search"
		Case "FWb"
			getESURI = "http://52.79.58.27:9200/fin-weblog-one-*/_search"  'fin-weblog-one-
		Case "FMb"
			getESURI = "http://52.79.58.27:9200/fin-moblog-one-*/_search"

		Case Else
			''
	End Select
end function

function getESURI_NEW(stDt,edDt,channel)
	dim resultSTR
	getESURI_NEW = ""

	Select Case channel
		Case "Web"
			getESURI_NEW = "http://k.tenbyten.kr:9200/weblog-one-*/_search?pretty"
		Case "Mob"
			getESURI_NEW = "http://k.tenbyten.kr:9200/moblog-one-*/_search"
		Case "App"
			getESURI_NEW = "http://k.tenbyten.kr:9200/moblog-one-*/_search"
		Case "FWb"
			getESURI_NEW = "http://k.tenbyten.kr:9200/fin-weblog-one-*/_search"  'fin-weblog-one-
		Case "FMb"
			getESURI_NEW = "http://k.tenbyten.kr:9200/fin-moblog-one-*/_search"

		Case Else
			''
	End Select
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

function fnSaveResultSimple_NEW(iparseType, channel, reqdata, dataArr)
    dim iAnalCon : set iAnalCon = CreateObject("ADODB.Connection")
    Dim strSql

    dim yyyymmdd,hh,os,itemid,addgubun,cnt,tmpData

    tmpData = Split(dataArr(0), " ")
    yyyymmdd = tmpData(0)
    hh = tmpData(1)
    os = dataArr(1)
    itemid = dataArr(2)
    addgubun = "CRT"
    if (dataArr(3) = "Y") then
    	addgubun = "ORD"
    end if
    cnt = dataArr(4)

    if (reqdata = "page_orderbaguni_app") and (itemid <> "") then
        strSql = "[db_analyze_data_raw].[dbo].[sp_TEN_ELK_data_add_to_cart] '"&yyyymmdd & " " & hh & ":00:00" & "','"&channel&"','"&os&"',"&itemid&",'"&addgubun&"',"&cnt
        ''response.write strSql & "<br />"
	    iAnalCon.Open Application("db_analyze")
	    iAnalCon.Execute strSql
	    iAnalCon.Close
    end if

    SET iAnalCon = Nothing
end function

function fnSaveResultSimpleGbn(iparseType,channel,reqdata,vTime,vGbnVal,vNoVal)
    dim iAnalCon : set iAnalCon = CreateObject("ADODB.Connection")
    Dim strSql

    strSql = "[db_analyze_data_raw].[dbo].[sp_TEN_ELK_data_ADD_simple_Gbn] '"&vTime&"','"&reqdata&"','"&channel&"','"&vGbnVal&"',"&vNoVal

    iAnalCon.Open Application("db_analyze")
    iAnalCon.Execute strSql
    iAnalCon.Close
    SET iAnalCon = Nothing
''  response.write strSql
end function

function fnSaveResultComplex(iparseType,channel,reqdata,vTime,vClientIp,vPageName,vRequestPage,vQueryparam,vReferer_url,vReferer_queryparam,vCode)
    dim iAnalCon : set iAnalCon = CreateObject("ADODB.Connection")
    Dim strSql

    'vPageName = replace(vPageName,vbCr,"")
    'vPageName = replace(vPageName,vbLf,"")
    'vPageName = replace(vPageName,vbCrLf,"")
    if (vCode="") then vCode=0

    ''SplitValue

    if (iparseType=3) then
        vQueryparam = replace(vQueryparam,"'","")
        vReferer_queryparam = replace(vReferer_queryparam,"'","")
        strSql = "[db_analyze_data_raw].[dbo].[sp_TEN_ELK_data_ADD_complex] '"&vTime&"','"&reqdata&"','"&channel&"','"&vClientIp&"','"&vPageName&"','"&vRequestPage&"','"&vQueryparam&"','"&vReferer_url&"','"&vReferer_queryparam&"',"&vCode

'    response.write strSql&VBCRLF
'if InStr(strSql,"only")>0 then
'    response.write strSql&VBCRLF
'else
        iAnalCon.Open Application("db_analyze")
        iAnalCon.Execute strSql
        iAnalCon.Close
'end if
    end if

    SET iAnalCon = Nothing
''  response.write strSql
end function

function fnElkParseUK()
exit function  ''2018/01/04 스케줄 방식으로 변경([ELK 데이터 취합 주문전환 매시 36분])
''    dim iAnalCon : set iAnalCon = CreateObject("ADODB.Connection")
''    Dim strSql
''    strSql = "[db_analyze_data_raw].[dbo].[sp_TEN_ELK_Parse_complex] "
''    iAnalCon.Open Application("db_analyze")
''    iAnalCon.Execute strSql
''    iAnalCon.Close
''
''    SET iAnalCon = Nothing
end function

%>