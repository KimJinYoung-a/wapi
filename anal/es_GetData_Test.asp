<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.charset = "utf-8" %>
<!-- #include virtual="/anal/es_queLib.asp" -->
<!-- #include virtual="/lib/util/aspJSON1.17.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<script language="jscript" runat="server">
    function jsURLDecode(v){ return decodeURI(v); }
    function jsURLEncode(v){ return encodeURI(v); }
</script>

<%
'response.write "TTT"  '' 2016/10/10 ''2016/12/12
'response.end

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
dim iparseType : iparseType = getParseType(reqdata)  ''파싱타입.

if (stDt="") then
    if (iparseType=3) then  ''complexType
        if (hour(now())="7") then
            stDt = dateadd("h",-7,now())
        else
            stDt = dateadd("h",-2,now())
        end if
        stDt = FormatDateTime(stDt,2)&" "&LEFT(FormatDateTime(stDt,4),2)&":00:00"
    else
        stDt = LEFT(dateadd("d",-1,now()),10)  ''전일
    end if
end if

if (edDt="") then
    if (iparseType=3) then
        if (hour(now())="7") then
            edDt = dateadd("h",+6,stDt)
        else
            edDt = dateadd("h",+1,stDt)
        end if
        edDt = FormatDateTime(edDt,2)&" "&LEFT(FormatDateTime(edDt,4),2)&":59:59"
    else
        edDt = LEFT(now(),10) & " 23:59:59" ''금일 23:59:59
    end if
end if

if (channel="") then channel = "Web"
if (reqdata="") then reqdata = "uqipCntPerDay"


''response.write stDt &"|"
''response.write edDt
'response.end

'// ============================================================================
'if (isautoscript="on") then
'    if (iparseType="1") or (iparseType="2") then
'        if (stDt="") or (edDt="") then
'            stDt = LEFT(dateadd("d",-1,now()),10)
'            edDt = LEFT(dateadd("d",+1,now()),10)
'        end if
'    end if
'end if

dim i, j, k, l, vvv
dim ibucket, bucketItem, bucketItem2, bucketItem3
dim vUnixTime, vUniqIpNo, vUniqIpOrderNo, vLocaleTime, vQueryparam, vReferer_url, vReferer_queryparam
dim vClientIp, vPageName, vRequestPage, vNoVal
dim otime : otime=Timer()

CALL debugWrite("iparseType:"&iparseType)

if (iparseType=4) or (iparseType=3) then
	retTxt = getESAPI_NEW(stDt,edDt,channel,reqdata,vCode)
else
	retTxt = getESAPI(stDt,edDt,channel,reqdata,vCode)
end if

dim qTime : qTime = FormatNumber(Timer()-otime,4)
otime=Timer()
CALL debugWrite("qTime:"&qTime)

response.write "<textarea cols=110 rows=30>"&retTxt&"</textarea>"
'response.end

dim oJSON
Dim assignedRow : assignedRow=0
''if (LEFT(reqdata,5)="page_") then  ''page_orderfinish

if (iparseType=3) then '' 데이터가 많음..  page_exists_uk
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

        vUnixTime = oJson.hits.hits.get(i).[_source].[@timestamp]
        vClientIp = oJson.hits.hits.get(i).[_source].[clientip]
        vRequestPage = oJson.hits.hits.get(i).[_source].[request]
     on Error resume Next
        vQueryparam = oJson.hits.hits.get(i).[_source].[queryparam]
        if Err then vQueryparam=""
     on Error Goto 0

     on Error resume Next
        vReferer_url = oJson.hits.hits.get(i).[_source].[referer_url]
        if Err then vReferer_url=""
     on Error Goto 0

     on Error resume Next
        vReferer_queryparam = oJson.hits.hits.get(i).[_source].[referer_queryparam]
        vReferer_queryparam = jsURLDecode(vReferer_queryparam)
        if Err then vReferer_queryparam=""
     on Error Goto 0

     on Error resume Next
       vPageName = oJson.hits.hits.get(i).[_source].[페이지명]  ''없는경우가 있음..
       if Err then vPageName="UNKNOWN"
     on Error Goto 0
       ''CALL debugWrite(GtmToKST(vUnixTime,19)&"||"&vClientIp&"||"&vPageName&"||"&vRequestPage&"||"&vQueryparam&"||"&vReferer_url&"||"&vReferer_queryparam)
       '' CALL debugWrite(UnixTimeToKST(vUnixTime,19)&"||"&vClientIp&"||"&vPageName&"||"&vRequestPage&"||"&vQueryparam&"||"&vReferer_url&"||"&vReferer_queryparam)

        ''if (application("Svr_Info")	= "Dev") then
            if (request("svresult")<>"") then
               '' Call fnSaveResultComplex(iparseType,channel,reqdata,UnixTimeToKST(vUnixTime,19),vClientIp,vPageName,vRequestPage,vQueryparam,vReferer_url,vReferer_queryparam,vCode)
                Call fnSaveResultComplex(iparseType,channel,reqdata,GtmToKST(vUnixTime,19),vClientIp,vPageName,vRequestPage,vQueryparam,vReferer_url,vReferer_queryparam,vCode)
                assignedRow = assignedRow+1
            end if
        ''end if
    next


    Set oJSON = Nothing

    if (reqdata="page_exists_uk") then
        call fnElkParseUK
    end if

elseif (iparseType=0) then
    dim keyword

    ''response.write retTxt
    Set oJSON = JSON.parse(retTxt)
    CALL debugWrite("aggregations.[retVal2].buckets.length:"&oJson.aggregations.[retVal2].buckets.length)

    for i=0 to oJson.aggregations.[retVal2].buckets.length-1
        keyword = oJson.aggregations.[retVal2].buckets.get(i).[key]
        vNoVal = oJson.aggregations.[retVal2].buckets.get(i).[doc_count]

        keyword = Trim(keyword)
        ''CALL debugWrite(keyword&"||"&keywordCnt)

        if (request("svresult")<>"") then
            Call fnSaveResultSimpleGbn(iparseType,channel,reqdata,stDt,keyword,vNoVal)
            assignedRow = assignedRow+1
        end if
    next

    Set oJSON = Nothing

elseif (iparseType=1) then  ''uqipCntPerDay

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
elseif (iparseType=4) then
	Set oJSON = JSON.parse(retTxt)

	for i = 0 to oJson.aggregations.[2].buckets.length - 1
		set ibucket = oJson.aggregations.[2].buckets.get(i)

		for j = 0 to ibucket.[3].buckets.length - 1
			set bucketItem = ibucket.[3].buckets.get(j)

			for k = 0 to bucketItem.[5].buckets.length - 1
				set bucketItem2 = bucketItem.[5].buckets.get(k)

				if (bucketItem2.[6].buckets.length = 0) then
					Call fnSaveResultSimple_NEW(iparseType, channel, reqdata, Array(UnixTimeToKST(ibucket.key, 13), bucketItem.key, bucketItem2.key, "N", bucketItem2.doc_count))
	            	assignedRow = assignedRow+1
				else
					for l = 0 to bucketItem2.[6].buckets.length - 1
						set bucketItem3 = bucketItem2.[6].buckets.get(l)

						Call fnSaveResultSimple_NEW(iparseType, channel, reqdata, Array(UnixTimeToKST(ibucket.key, 13), bucketItem.key, bucketItem2.key, bucketItem3.key, bucketItem3.doc_count))
		            	assignedRow = assignedRow+1

						set bucketItem3 = Nothing
					next
				end if

				set bucketItem2 = Nothing
			next

			set bucketItem = Nothing
		next

		set ibucket = Nothing
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
response.write "qTime="&qTime&"|parseTime="&parseTime
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

        <option value='FWb' <% if channel="FWb" then response.write "selected" %> >FWb</option>
        <option value='FMb' <% if channel="FMb" then response.write "selected" %> >FMb</option>
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


        <option value='page_exists_uk' <% if reqdata="page_exists_uk" then response.write "selected" %> >uk존재</option>

        <option value='page_orderfinish' <% if reqdata="page_orderfinish" then response.write "selected" %> >주문완료페이지</option>
        <option value='page_orderbaguni' <% if reqdata="page_orderbaguni" then response.write "selected" %> >장바구니담기</option>

        <option value='page_item' <% if reqdata="page_item" then response.write "selected" %> >상품페이지</option>
        <option value='page_ref_tmailer' <% if reqdata="page_ref_tmailer" then response.write "selected" %> >레퍼러:tmailer</option>

        <option value='best_keywords' <% if reqdata="best_keywords" then response.write "selected" %> >베스트 키워드</option>
        <option value='page_orderbaguni_app' <% if reqdata="page_orderbaguni_app" then response.write "selected" %> >장바구니담기(앱)</option>
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