<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% Server.ScriptTimeOut = 240 %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Response.CharSet="UTF-8"
%>
<!-- #include virtual="/lib/util/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib_UTF8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/util/aspJSON1.17.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<%

''select case mode
''    case "agvSendErrClear":
''        '// AGV 오차 초기화
''    case "agvipgo":
''        '// AGV 진열등록
''    case "agvipgodel":
''        '// AGV 진열등록삭제
''    case "agvSendOrder":
''        '// AGV 주문전송
''    case "agvSendBalju":
''        '// AGV 발주전송
''    case "agvSendBaljuCancel":
''        '// AGV 발주전송취소
''    case "agvSendBaljuMulti":
''        '// AGV 발주전송[멀티스테이션]
''    case "agvSendBaljuCancelMulti":
''        '// AGV 발주전송취소[멀티스테이션]
''    case "currstock":
''        '// 특정 상품(한개) 현재고 조회
''    case "currstockListView":
''        '// 상품(브랜드 리스트 또는 상품리스트) 현재고 조회, 저장안함
''    case "currstockList":
''        '// 상품(브랜드 리스트 또는 상품리스트) 현재고 조회, 재고정보 저장
''    case "chgwarehouse2bulk":
''        '// 상품 재고구분 벌크전환
''    case "currstockall":
''        '// 전체 상품 현재고 조회
''    case "agvpickCS":
''        '// CS 피킹 지시등록
''    case "agvpickCSdel":
''        '// CS 피킹 지시삭제
''    case "agvpickup":
''        '// 피킹 지시등록
''    case "agvpickupdel":
''        '// 피킹 지시삭제
''    case "agvstockinvest":
''        '// 재고조사 지시등록
''    case "agvstockinvestdel":
''        '// 재고조사 지시삭제
''    case "senditeminfo":
''        '// 상품정보전송
''    case "recvagvipgo":
''        '// 진열상태변경 수신
''    case "recvagvpick":
''        '// 피킹상태변경 수신
''    case "recvagvpickfinish":
''        '// 피킹완료시 수신(피킹완료수량 포함)
''    case "recvagvstockchange":
''        '// 재고증감정보 수신
''    case "recvagvsurveychange":
''        '// 재고조사 상태변경 보고 수신
''    case "recvagvdistribfinish":
''        '// 분배완료 상태보고(결품수량 전송)
''    case "recvagvWareHouseChg":
''        '// 재고위치 변경정보 수신
''    case else
''        response.write "1111"
''end select

''포천물류 공인IP주소
'' - 175.194.235.185 : 1F
'' - 175.194.235.186 : 2F
'' - 175.194.235.182    ' 175.194.235.187 : 3F

''개발서버 IP
'' - 121.78.103.2

''webadmin 서버 IP
'' - 110.93.128.93

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("192.168.1.67","192.168.1.70","192.168.1.71","192.168.1.72","110.93.128.107","121.78.103.60","110.93.128.114","110.93.128.94","110.93.128.113","175.194.235.185","175.194.235.186","175.194.235.187","175.194.235.182", "121.78.103.2", "110.93.128.93", "::1")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

'// ============================================================================
dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) and (Not CheckJenkinsServerIP(ref)) then
    response.write ref
    response.end
end if


''dim AGVSERVERURL : AGVSERVERURL = "http://218.155.198.200/tenbyten"
dim AGVSERVERURL : AGVSERVERURL = "http://175.194.235.182:58080/tenbyten"
''dim AGVSERVERURL : AGVSERVERURL = "http://175.194.235.187:58080/tenbyten"

If (application("Svr_Info")	= "Dev") Then
    AGVSERVERURL = "http://218.155.198.200/tenbyten"
end if


dim mode, requestNo, requestMaster, skuCd, warehouseCd, jsonResult, skuCdArr, totalQty, totalQtyArr, requestNoArr
dim pickingOrderTypeCd, stationCd, baljuKey
dim resultCode, resultMessage, failCode, pickingOrderNo, displayOrderNo, displayOrderNoArr, ordertype
dim i, j, k, responseJson, callback, affectedRows, jsonString
dim sqlStr
dim orderserial, orderserialArr, orderJson, siteBaljuid, valuesArr, baljuJson, dasindex, lngBytesCount
dim progressStatusCd, resultJson, dataCount
dim brandArray, skuCdArray, dataCountForSend, baljucode, baljuStr, masteridx, status
dim minidx, maxidx, finishedSkuQty, resultData, item, orderData, caseNo, subItem, orderQty, distributionQty
dim itemgubun, itemid, itemoption, itemgubunArr, itemidArr, itemoptionArr
dim locationCd, itemno, yyyymmdd
dim inventorySurveyOrderId
dim newBaljuKey, newBaljuKeyArr
dim successCnt, failCnt, workTypeCd

mode  			= requestCheckVar(request("mode"), 32)
requestNo		= requestCheckVar(request("requestNo"), 32)
requestMaster	= requestCheckVar(request("requestMaster"), 32)
skuCd			= requestCheckVar(request("skuCd"), 32)
stationCd		= requestCheckVar(request("stationCd"), 32)
baljuKey		= requestCheckVar(request("baljuKey"), 32)
callback		= requestCheckVar(request("callback"), 32)
ordertype		= requestCheckVar(request("ordertype"), 32)
baljucode		= requestCheckVar(request("baljucode"), 32)
baljuStr		= requestCheckVar(request("baljuStr"), 64)
masteridx		= requestCheckVar(request("masteridx"), 64)

brandArray		= requestCheckVar(request("brandArray"), 4000)
skuCdArray		= requestCheckVar(request("skuCdArray"), 4000)

itemgubun		= requestCheckVar(request("itemgubun"), 4000)
itemid			= requestCheckVar(request("itemid"), 4000)
itemoption		= requestCheckVar(request("itemoption"), 4000)

Function BytesToStr(bytes, charset)
    Dim Stream
    Set Stream = Server.CreateObject("Adodb.Stream")
        Stream.Type = 1 'adTypeBinary
        Stream.Open
        Stream.Write bytes
        Stream.Position = 0
        Stream.Type = 2 'adTypeText
        Stream.Charset = charset
        BytesToStr = Stream.ReadText
        Stream.Close
    Set Stream = Nothing
End Function


'// ============================================================================
'// ILC 호출 메인함수
'// ============================================================================
function fnSendReceiveJson(url, method, jsonString, jsonResult)
    Dim HTTP_Object
    dim SiteURL :SiteURL = AGVSERVERURL
    dim result

    SiteURL = SiteURL & "/" & url

    Set HTTP_Object = Server.CreateObject("MSXML2.ServerXMLHTTP")

    With HTTP_Object
        .SetTimeouts 30000, 30000, 30000, 30000
        .Open method, SiteURL, False
        .SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
        .Send jsonString
        .WaitForResponse 60
    End With

    If HTTP_Object.Status = "200" or HTTP_Object.Status = "201" Then
        result = HTTP_Object.ResponseText
    else
    	select case HTTP_Object.Status
            case "201":
                result = fnCreateCustomJsonResult(HTTP_Object.Status, "Created", HTTP_Object.Status)
            case "401":
                result = fnCreateCustomJsonResult(HTTP_Object.Status, "Unauthorized", HTTP_Object.Status)
            case "403":
                result = fnCreateCustomJsonResult(HTTP_Object.Status, "Forbidden", HTTP_Object.Status)
            case "404":
                result = fnCreateCustomJsonResult(HTTP_Object.Status, "Not Found", HTTP_Object.Status)
            case "500":
                result = fnCreateCustomJsonResult(HTTP_Object.Status, "Server Error", HTTP_Object.Status)
            case else:
                result = fnCreateCustomJsonResult(HTTP_Object.Status, "Unknown", HTTP_Object.Status)
        end select
    end if

    Set HTTP_Object = Nothing

    jsonResult = result
end function

function fnWriteLog(mode, logtext)
    dim sqlStr

    sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_api_log](mode, logtext)"
    sqlStr = sqlStr + " values('" & mode & "', '" & Replace(logtext, "'", "^") & "')"
    dbget_Logistics.Execute sqlStr
end function

'// ============================================================================
'// JSON 리턴값 생성함수
'// ============================================================================
function fnCreateCustomJsonResult(resultCode, resultMessage, failCode)
    dim jsonString

    jsonString = ""
    jsonString = jsonString & "{" & vbCrLf
    jsonString = jsonString & "  ""resultCode"": """ & resultCode & """," & vbCrLf
    jsonString = jsonString & "  ""resultMessage"": """ & resultMessage & """," & vbCrLf
    jsonString = jsonString & "  ""failCode"": """ & failCode & """" & vbCrLf
    jsonString = jsonString & "}"

    fnCreateCustomJsonResult = jsonString
end function

'// ============================================================================
'// ILC JSON 리턴값 파싱함수
'// ============================================================================
function fnParseResultJson(jsonResult, resultCode, resultMessage, failCode)
    dim resultJson

    Set resultJson = New aspJson
    resultJson.loadJSON(jsonResult)

    resultCode = resultJson.data("resultCode")
    resultMessage = resultJson.data("resultMessage")
    failCode = resultJson.data("failCode")
end function

'// ============================================================================
'// ILC 피킹 JSON 리턴값 파싱함수
'// ============================================================================
function fnParseResultPickingJson(jsonResult, ByRef resultCode, ByRef resultMessage, ByRef failCode, ByRef pickingOrderNo)
    dim resultJson

    Set resultJson = New aspJson
    resultJson.loadJSON(jsonResult)

    resultCode = resultJson.data("resultCode")
    resultMessage = resultJson.data("resultMessage")
    failCode = resultJson.data("failCode")
    pickingOrderNo = resultJson.data("pickingOrderNo")
end function

'// ============================================================================
'// ILC 재고조사 JSON 리턴값 파싱함수
'// ============================================================================
function fnParseResultStockInvestJson(jsonResult, ByRef resultCode, ByRef resultMessage, ByRef failCode, ByRef inventorySurveyOrderId)
    dim resultJson

    Set resultJson = New aspJson
    resultJson.loadJSON(jsonResult)

    resultCode = resultJson.data("resultCode")
    resultMessage = resultJson.data("resultMessage")
    failCode = resultJson.data("failCode")
    inventorySurveyOrderId = resultJson.data("inventorySurveyOrderId")
end function

'// ============================================================================
'// ILC 발주 JSON 리턴값 파싱함수
'// ============================================================================
function fnParseResultSendBaljuJson(jsonResult, ByRef resultCode, ByRef resultMessage, ByRef failCode)
    dim resultJson

    Set resultJson = New aspJson
    resultJson.loadJSON(jsonResult)

    resultCode = resultJson.data("resultCode")
    resultMessage = resultJson.data("resultMessage")
    failCode = resultJson.data("failCode")
end function

'// ============================================================================
'// ILC 재고조회 JSON 리턴값 파싱함수
'// ============================================================================
function fnParseResultProductStockInfoByBrandBySkuCd(jsonResult, ByRef resultCode, ByRef resultMessage, ByRef failCode)
    dim resultJson

    Set resultJson = New aspJson
    resultJson.loadJSON(jsonResult)

    resultCode = resultJson.data("resultCode")
    resultMessage = resultJson.data("resultMessage")
    failCode = resultJson.data("failCode")
end function

'// ============================================================================
'// ILC 진열 JSON 리턴값 파싱함수
'// ============================================================================
function fnParseResultDisplayJson(jsonResult, ByRef resultCode, ByRef resultMessage, ByRef failCode, ByRef displayOrderNo)
    dim resultJson

    Set resultJson = New aspJson
    resultJson.loadJSON(jsonResult)

    resultCode = resultJson.data("resultCode")
    resultMessage = resultJson.data("resultMessage")
    failCode = resultJson.data("failCode")
    displayOrderNo = resultJson.data("displayOrderNo")
end function

'// ============================================================================
'// ILC 브랜드정보(여러 건 한번에) 전송 API
'// ============================================================================
function fnSendBrandInfo(method, jsonString, ByRef jsonResult)
    dim url : url = "api/site/brand"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// ILC 상품정보 전송 API
'// ============================================================================
function fnSendProductInfo(method, jsonString, ByRef jsonResult)
    dim url : url = "api/site/sku"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// ILC 진열요청 전송 API
'// ============================================================================
function fnSendDisplayRequest(method, jsonString, ByRef jsonResult)
    dim url : url = "api/site/savedisplayorder"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// ILC 진열요청취소 전송 API
'// ============================================================================
function fnSendCancelDisplayRequest(method, jsonString, ByRef jsonResult)
    dim url : url = "api/site/canceldisplayorder"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// ILC 피킹요청 전송 API
'// ============================================================================
function fnSendPickRequest(method, jsonString, ByRef jsonResult)
    dim url : url = "api/site/savepickingorder"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// ILC 피킹취소요청 전송 API
'// ============================================================================
function fnSendPickCancelRequest(method, jsonString, ByRef jsonResult)
    dim url : url = "api/site/cancelpickingorder"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// ILC 재고조사 전송 API
'// ============================================================================
function fnSendStockInvestRequest(method, jsonString, ByRef jsonResult)
    dim url : url = "api/site/addInventorySurvey"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// ILC 재고조사 취소 전송 API
'// ============================================================================
function fnSendStockInvestCancelRequest(method, jsonString, ByRef jsonResult)
    dim url : url = "api/site/cancelInventorySurvey"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// ILC 재고조회(건별) 전송 API
'// ============================================================================
function fnGetProductStockInfo(method, jsonString, ByRef jsonResult)
    dim url : url = "api/standard/inventory/item"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// ILC 재고조회(리스트) 전송 API
'// ============================================================================
function fnGetProductStockInfoList(method, jsonString, ByRef jsonResult)
    dim url : url = "api/standard/inventory/item/brand"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// ILC AGV 오차 초기화 전송 API
'// ============================================================================
function fnSendErrClearInfo(method, jsonString, ByRef jsonResult)
    dim url : url = "api/standard/inventory/item/ajustclear"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// 날짜 yyyy-mm-dd hh:mm:ss 생성 함수
'// ============================================================================
function getDateFormatedWithDash(DateVal)
    dim rtnDateStr, m, d, h, min, sec
    rtnDateStr = year(DateVal)
    m = month(DateVal)
    d = day(DateVal)
    h = Hour(DateVal)
    Min = Minute(DateVal)
    sec = second(DateVal)

    if month(DateVal) < 10 then
        m = "0" & month(DateVal)
    end if

    if day(DateVal) < 10 then
        d = "0" & day(DateVal)
    end if

    if Hour(DateVal) < 10 then
        h = "0" & Hour(DateVal)
    end if

    if Minute(DateVal) < 10 then
        Min = "0" & Minute(DateVal)
    end if

    if second(DateVal) < 10 then
        sec = "0" & second(DateVal)
    end if

    rtnDateStr = rtnDateStr & "-" & m & "-" & d & " " & h & ":" & Min & ":" & sec
    getDateFormatedWithDash = rtnDateStr
end function

'// ============================================================================
'// uuid 로 구분된 브랜드정보 일괄전송 함수
'// ============================================================================
function fnSendBrandInfoByUUID(uuid)
    dim sqlStr, i
    dim method, jsonString, jsonResult
    dim resultCode, resultMessage, failCode

    method = "POST"

    sqlStr = " select brandCd, brandName from [db_aLogistics].[dbo].[tbl_agv_sendState_BrandInfo] "
    sqlStr = sqlStr + " where uuid = '" & uuid & "' "
    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    jsonString = ""
    jsonString = jsonString & "{" & vbCrLf
    jsonString = jsonString & "  ""brandList"": [" & vbCrLf

	i=0
	if  not rsget_Logistics.EOF  then
		do until rsget_Logistics.eof
            if (i = 0) then
                jsonString = jsonString & "    {" & vbCrLf
                jsonString = jsonString & "      ""brandCd"": """ & rsget_Logistics("brandCd") & """," & vbCrLf
                jsonString = jsonString & "      ""brandName"": """ & rsget_Logistics("brandName") & """" & vbCrLf
                jsonString = jsonString & "    }" & vbCrLf
            else
                jsonString = jsonString & "    ,{" & vbCrLf
                jsonString = jsonString & "      ""brandId"": """ & rsget_Logistics("brandCd") & """," & vbCrLf
                jsonString = jsonString & "      ""brandName"": """ & rsget_Logistics("brandName") & """" & vbCrLf
                jsonString = jsonString & "    }" & vbCrLf
            end if

			i=i+1
			rsget_Logistics.moveNext
		loop
	end if
	rsget_Logistics.close

    jsonString = jsonString & "  ]" & vbCrLf
    jsonString = jsonString & "}"

    Call fnSendBrandInfo(method, jsonString, jsonResult)

    Call fnParseResultJson(jsonResult, resultCode, resultMessage, failCode)
end function

'// JSON Escape
Function json_escape(str)
	If isEmpty(str) Or isNull(str) Then Exit Function

	str = Trim(str)
	str = Replace(str, """", "")		'// 따옴표 제거
	str = Replace(str, "'", "'")
	str = Replace(str, Chr(92), "\\")
	str = Replace(str, Chr(13) & Chr(10), "\n")
	str = Replace(str, Chr(13), "\n")
	str = Replace(str, Chr(10), "\n")
	str = Replace(str, Chr(9), "\t")
	json_escape = str
End Function

'// ============================================================================
'// uuid 로 구분된 상품정보 일괄전송 함수
'// ============================================================================
function fnSendProductInfoByUUID(uuid, pickingWarehouseCd)
    dim sqlStr, i
    dim method, jsonString, jsonResult
    dim resultCode, resultMessage, failCode, optionname, productName
    dim dataCountForSend

    dim testSkuCd : testSkuCd = "10016595790000"
    dim testSkuCdExists : testSkuCdExists = False

    method = "POST"

    jsonString = ""
    jsonString = jsonString & "{" & vbCrLf
    jsonString = jsonString & "  ""skuList"": [" & vbCrLf

    sqlStr = " select *, IsNull(pickingWarehouseCd, 'BULK') as pickingWarehouseCd from [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] "
    sqlStr = sqlStr + " where uuid = '" & uuid & "' "
    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

	i=0
    dataCountForSend = 0
	if  not rsget_Logistics.EOF  then
		do until rsget_Logistics.eof
            productName = db2html(rsget_Logistics("productName"))
            optionname = db2html(rsget_Logistics("optionName"))

            if IsNull(optionname) or optionname = "" then
                optionname = "-"
            end if

            productName = json_escape(productName)
            optionname = json_escape(optionname)

            if (rsget_Logistics("skuCd") = testSkuCd) then
                testSkuCdExists = True
            end if

            if (i = 0) then
                jsonString = jsonString & "    {" & vbCrLf
                jsonString = jsonString & "      ""barcode"": """ & rsget_Logistics("barcode") & """," & vbCrLf
                jsonString = jsonString & "      ""brandCd"": """ & rsget_Logistics("brandCd") & """," & vbCrLf
                jsonString = jsonString & "      ""optionName"": """ & optionname & """," & vbCrLf
                jsonString = jsonString & "      ""photoUrl"": """ & rsget_Logistics("imgUrl") & """," & vbCrLf
                if (pickingWarehouseCd <> "") then
                    if (pickingWarehouseCd = "AGV") then
                        jsonString = jsonString & "      ""pickingWarehouseCd"": ""AGV_WAREHOUSE""," & vbCrLf
                    elseif (pickingWarehouseCd = "BLK") then
                        jsonString = jsonString & "      ""pickingWarehouseCd"": ""BULK_WAREHOUSE""," & vbCrLf
                    end if
                end if
                jsonString = jsonString & "      ""productName"": """ & productName & """," & vbCrLf
                jsonString = jsonString & "      ""sizeX"": ""0""," & vbCrLf
                jsonString = jsonString & "      ""sizeY"": ""0""," & vbCrLf
                jsonString = jsonString & "      ""sizeZ"": ""0""," & vbCrLf
                jsonString = jsonString & "      ""skuCd"": """ & rsget_Logistics("skuCd") & """," & vbCrLf
                jsonString = jsonString & "      ""skuTypeCd"": """ & rsget_Logistics("skuTypeCd") & """," & vbCrLf
                jsonString = jsonString & "      ""weight"": ""0""" & vbCrLf
                jsonString = jsonString & "    }" & vbCrLf
            else
                jsonString = jsonString & "    , {" & vbCrLf
                jsonString = jsonString & "      ""barcode"": """ & rsget_Logistics("barcode") & """," & vbCrLf
                jsonString = jsonString & "      ""brandCd"": """ & rsget_Logistics("brandCd") & """," & vbCrLf
                jsonString = jsonString & "      ""optionName"": """ & optionname & """," & vbCrLf
                jsonString = jsonString & "      ""photoUrl"": """ & rsget_Logistics("imgUrl") & """," & vbCrLf
                if (pickingWarehouseCd <> "") then
                    if (pickingWarehouseCd = "AGV") then
                        jsonString = jsonString & "      ""pickingWarehouseCd"": ""AGV_WAREHOUSE""," & vbCrLf
                    elseif (pickingWarehouseCd = "BLK") then
                        jsonString = jsonString & "      ""pickingWarehouseCd"": ""BULK_WAREHOUSE""," & vbCrLf
                    end if
                end if
                jsonString = jsonString & "      ""productName"": """ & productName & """," & vbCrLf
                jsonString = jsonString & "      ""sizeX"": ""0""," & vbCrLf
                jsonString = jsonString & "      ""sizeY"": ""0""," & vbCrLf
                jsonString = jsonString & "      ""sizeZ"": ""0""," & vbCrLf
                jsonString = jsonString & "      ""skuCd"": """ & rsget_Logistics("skuCd") & """," & vbCrLf
                jsonString = jsonString & "      ""skuTypeCd"": """ & rsget_Logistics("skuTypeCd") & """," & vbCrLf
                jsonString = jsonString & "      ""weight"": ""0""" & vbCrLf
                jsonString = jsonString & "    }" & vbCrLf
            end if

			i=i+1
            dataCountForSend = dataCountForSend + 1
			rsget_Logistics.moveNext
		loop
	end if
	rsget_Logistics.close

    jsonString = jsonString & "  ]" & vbCrLf
    jsonString = jsonString & "}" & vbCrLf

    if (dataCountForSend = 0) then
        '// 전송할 데이타 없음
    else
        Call fnSendProductInfo(method, jsonString, jsonResult)

        Call fnParseResultJson(jsonResult, resultCode, resultMessage, failCode)

        if (resultCode <> "00") then
            Call fnWriteLog(mode, jsonString)
            Call fnWriteLog(mode, jsonResult)
        else
            if (testSkuCdExists = True) then
                Call fnWriteLog(mode, jsonString)
                Call fnWriteLog(mode, jsonResult)
            end if
        end if
    end if
end function

'// ============================================================================
'// 피킹요청 일괄전송
'// ============================================================================
function fnSendPickingByRequestNo(requestNo, ByRef jsonResult)
    dim sqlStr, i
    dim method, jsonString
    dim resultCode, resultMessage, failCode

    method = "POST"

    sqlStr = " select p.*, i.skuCd "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_pickupItems] p "
    sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] i "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and p.itemgubun = i.itemGubun "
    sqlStr = sqlStr + " 		and p.itemid = i.itemid "
    sqlStr = sqlStr + " 		and p.itemoption = i.itemoption "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	requestNo = '" & requestNo & "' "

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    if rsget_Logistics.EOF  then
        jsonResult = fnCreateCustomJsonResult("500", "No Data", "500")
        rsget_Logistics.close
        Exit Function
    end if

    if  not rsget_Logistics.EOF  then
        jsonString = ""
        jsonString = jsonString & "{" & vbCrLf
        jsonString = jsonString & "  ""pickingOrderTypeCd"": """ & rsget_Logistics("pickingOrderTypeCd") & """," & vbCrLf
        jsonString = jsonString & "  ""requestNo"": """ & requestNo & """," & vbCrLf
        jsonString = jsonString & "  ""stationCd"": """ & rsget_Logistics("stationCd") & """," & vbCrLf
        jsonString = jsonString & "  ""skuList"": [" & vbCrLf

        i = 0
        do until rsget_Logistics.eof
            if (i = 0) then
                jsonString = jsonString & "    {" & vbCrLf
                jsonString = jsonString & "      ""qty"": " & rsget_Logistics("itemno") & "," & vbCrLf
                jsonString = jsonString & "      ""skuCd"": """ & rsget_Logistics("skuCd") & """" & vbCrLf
                jsonString = jsonString & "    }" & vbCrLf
            else
                jsonString = jsonString & "    ,{" & vbCrLf
                jsonString = jsonString & "      ""qty"": " & rsget_Logistics("itemno") & "," & vbCrLf
                jsonString = jsonString & "      ""skuCd"": """ & rsget_Logistics("skuCd") & """" & vbCrLf
                jsonString = jsonString & "    }" & vbCrLf
            end if

			i=i+1
			rsget_Logistics.moveNext
		loop

        jsonString = jsonString & "  ]" & vbCrLf
        jsonString = jsonString & "}" & vbCrLf
    end if
    rsget_Logistics.close

    Call fnSendPickRequest(method, jsonString, jsonResult)
end function

'// ============================================================================
'// 피킹요청 일괄전송
'// ============================================================================
function fnSendPickingByMasterIDX(masteridx, ByRef jsonResult)
    dim sqlStr, i
    dim method, jsonString
    dim resultCode, resultMessage, failCode

    method = "POST"

    sqlStr = " select m.pickingOrderTypeCd, m.requestNo, m.stationCd, i.skuCd, d.itemno "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_pickup_master] m "
    sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_agv_pickup_detail] d on m.idx = d.masteridx "
    sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] i "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and d.itemgubun = i.itemGubun "
    sqlStr = sqlStr + " 		and d.itemid = i.itemid "
    sqlStr = sqlStr + " 		and d.itemoption = i.itemoption "
    sqlStr = sqlStr + " 		and d.deldt is NULL "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.idx = '" & masteridx & "' "

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    if rsget_Logistics.EOF  then
        jsonResult = fnCreateCustomJsonResult("500", "No Data", "500")
        rsget_Logistics.close
        Exit Function
    end if

    if  not rsget_Logistics.EOF  then
        jsonString = ""
        jsonString = jsonString & "{" & vbCrLf
        jsonString = jsonString & "  ""pickingOrderTypeCd"": """ & rsget_Logistics("pickingOrderTypeCd") & """," & vbCrLf
        jsonString = jsonString & "  ""requestNo"": """ & rsget_Logistics("requestNo") & """," & vbCrLf
        jsonString = jsonString & "  ""stationCd"": """ & rsget_Logistics("stationCd") & """," & vbCrLf
        jsonString = jsonString & "  ""skuList"": [" & vbCrLf

        i = 0
        do until rsget_Logistics.eof
            if (i = 0) then
                jsonString = jsonString & "    {" & vbCrLf
                jsonString = jsonString & "      ""qty"": " & rsget_Logistics("itemno") & "," & vbCrLf
                jsonString = jsonString & "      ""skuCd"": """ & rsget_Logistics("skuCd") & """" & vbCrLf
                jsonString = jsonString & "    }" & vbCrLf
            else
                jsonString = jsonString & "    ,{" & vbCrLf
                jsonString = jsonString & "      ""qty"": " & rsget_Logistics("itemno") & "," & vbCrLf
                jsonString = jsonString & "      ""skuCd"": """ & rsget_Logistics("skuCd") & """" & vbCrLf
                jsonString = jsonString & "    }" & vbCrLf
            end if

			i=i+1
			rsget_Logistics.moveNext
		loop

        jsonString = jsonString & "  ]" & vbCrLf
        jsonString = jsonString & "}" & vbCrLf
    end if
    rsget_Logistics.close

    Call fnWriteLog(mode, jsonString)

    Call fnSendPickRequest(method, jsonString, jsonResult)

    Call fnWriteLog(mode, jsonResult)
end function

'// ============================================================================
'// 피킹요청 일괄전송
'// ============================================================================
function fnSendPickingCancelByMasterIDX(masteridx, ByRef jsonResult)
    dim sqlStr, i
    dim method, jsonString
    dim resultCode, resultMessage, failCode

    method = "PUT"

    sqlStr = " select m.pickingOrderNo "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_pickup_master] m "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.idx = '" & masteridx & "' "

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    if rsget_Logistics.EOF  then
        jsonResult = fnCreateCustomJsonResult("500", "No Data", "500")
        rsget_Logistics.close
        Exit Function
    end if

    if  not rsget_Logistics.EOF  then
        jsonString = ""
        jsonString = jsonString & "{" & vbCrLf
        jsonString = jsonString & "  ""pickingOrderNo"": """ & rsget_Logistics("pickingOrderNo") & """" & vbCrLf
        jsonString = jsonString & "}" & vbCrLf
    end if
    rsget_Logistics.close

    Call fnWriteLog(mode, jsonString)

    Call fnSendPickCancelRequest(method, jsonString, jsonResult)
end function

'// ============================================================================
'// 재고조사 일괄전송
'// ============================================================================
function fnSendStockInvestByMasterIDX(masteridx, ByRef jsonResult)
    dim sqlStr, i
    dim method, jsonString
    dim resultCode, resultMessage, failCode

    method = "POST"

    sqlStr = " select m.requestNo, m.stationCd, i.skuCd "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_stock_invest_master] m "
    sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_agv_stock_invest_detail] d on m.idx = d.masteridx "
    sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] i "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and d.itemgubun = i.itemGubun "
    sqlStr = sqlStr + " 		and d.itemid = i.itemid "
    sqlStr = sqlStr + " 		and d.itemoption = i.itemoption "
    sqlStr = sqlStr + " 		and d.deldt is NULL "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.idx = '" & masteridx & "' "

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    if rsget_Logistics.EOF  then
        jsonResult = fnCreateCustomJsonResult("500", "No Data", "500")
        rsget_Logistics.close
        Exit Function
    end if

    if  not rsget_Logistics.EOF  then
        jsonString = ""
        jsonString = jsonString & "{" & vbCrLf
        jsonString = jsonString & "  ""stocktakingOrderNote"": ""재고조사지시 노트""," & vbCrLf
        jsonString = jsonString & "  ""requestNo"": """ & rsget_Logistics("requestNo") & """," & vbCrLf
        jsonString = jsonString & "  ""stationCd"": """ & rsget_Logistics("stationCd") & """," & vbCrLf
        jsonString = jsonString & "  ""skuList"": [" & vbCrLf

        i = 0
        do until rsget_Logistics.eof
            if (i = 0) then
                jsonString = jsonString & "    {" & vbCrLf
                jsonString = jsonString & "      ""skuCd"": """ & rsget_Logistics("skuCd") & """" & vbCrLf
                jsonString = jsonString & "    }" & vbCrLf
            else
                jsonString = jsonString & "    ,{" & vbCrLf
                jsonString = jsonString & "      ""skuCd"": """ & rsget_Logistics("skuCd") & """" & vbCrLf
                jsonString = jsonString & "    }" & vbCrLf
            end if

			i=i+1
			rsget_Logistics.moveNext
		loop

        jsonString = jsonString & "  ]" & vbCrLf
        jsonString = jsonString & "}" & vbCrLf
    end if
    rsget_Logistics.close

    Call fnWriteLog(mode, jsonString)

    Call fnSendStockInvestRequest(method, jsonString, jsonResult)
end function

'// ============================================================================
'// 재고조사 취소 전송
'// ============================================================================
function fnSendStockInvestCancelByMasterIDX(masteridx, ByRef jsonResult)
    dim sqlStr, i
    dim method, jsonString
    dim resultCode, resultMessage, failCode

    method = "PUT"

    sqlStr = " select m.inventorySurveyOrderId "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_stock_invest_master] m "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.idx = '" & masteridx & "' "

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    if rsget_Logistics.EOF  then
        jsonResult = fnCreateCustomJsonResult("500", "No Data", "500")
        rsget_Logistics.close
        Exit Function
    end if

    if  not rsget_Logistics.EOF  then
        jsonString = ""
        jsonString = jsonString & "{" & vbCrLf
        jsonString = jsonString & "  ""inventorySurveyOrderId"": """ & rsget_Logistics("inventorySurveyOrderId") & """" & vbCrLf
        jsonString = jsonString & "}" & vbCrLf
    end if
    rsget_Logistics.close

    Call fnWriteLog(mode, jsonString)

    Call fnSendStockInvestCancelRequest(method, jsonString, jsonResult)

    Call fnWriteLog(mode, jsonResult)
end function

'// ============================================================================
'// 전시요청 건별전송
'// ============================================================================
function fnSendDisplayByRequestNo(requestNo, ByRef jsonResult)
    dim sqlStr, i
    dim method, jsonString
    dim resultCode, resultMessage, failCode

    method = "POST"

    sqlStr = " select top 1 s.idx as requestNo, s.displayOrderTypeCd, i.skuCd, s.realstock, s.requestMaster "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_scheduledItems] s "
    sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] i "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and s.itemGubun = i.itemGubun "
    sqlStr = sqlStr + " 		and s.itemid = i.itemid "
    sqlStr = sqlStr + " 		and s.itemoption = i.itemoption "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and s.idx = " & requestNo

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    if  not rsget_Logistics.EOF  then
        jsonString = ""
        jsonString = jsonString & "{"
        jsonString = jsonString & "  ""displayOrderTypeCd"": """ & rsget_Logistics("displayOrderTypeCd") & ""","
        jsonString = jsonString & "  ""memo"": """ & rsget_Logistics("requestMaster") & ""","
        jsonString = jsonString & "  ""qty"": " & rsget_Logistics("realstock") & ","
        jsonString = jsonString & "  ""requestNo"": """ & rsget_Logistics("requestNo") & ""","
        jsonString = jsonString & "  ""skuCd"": """ & rsget_Logistics("skuCd") & """"
        jsonString = jsonString & "}"
    end if
    rsget_Logistics.close

    Call fnSendDisplayRequest(method, jsonString, jsonResult)
end function

'// ============================================================================
'// 전시요청 취소전송
'// ============================================================================
function fnSendCancelDisplayByDisplayOrderNo(displayOrderNo, ByRef jsonResult)
    dim sqlStr, i
    dim method, jsonString
    dim resultCode, resultMessage, failCode

    method = "PUT"

    jsonString = ""
    jsonString = jsonString & "{"
    jsonString = jsonString & "  ""displayOrderNo"": """ & displayOrderNo & """"
    jsonString = jsonString & "}"

    Call fnSendCancelDisplayRequest(method, jsonString, jsonResult)
end function

'// ============================================================================
'// 재고정보 요청(건별)
'// ============================================================================
function fnGetProductStockInfoBySkuCd(skuCd, ByRef jsonResult)
    dim method, jsonString
    dim resultCode, resultMessage, failCode
    dim resultData, skuList, totalQty, item

    method = "POST"

    jsonString = ""
    jsonString = "["
    jsonString = jsonString + "  {"
    jsonString = jsonString + "    ""skuCd"": """ & skuCd & """"
    jsonString = jsonString + "  }"
    jsonString = jsonString + "]"

    Call fnGetProductStockInfo(method, jsonString, jsonResult)
end function

'// ============================================================================
'// 재고정보 요청(전체, 또는 일부)
'// ============================================================================
function fnGetProductStockInfoByBrandBySkuCd(brandArray, skuCdArray, ByRef jsonResult)
    dim method, jsonString
    dim resultCode, resultMessage, failCode
    dim resultData, skuList, totalQty, item
    dim brand, skuCd, i

    method = "POST"

    jsonString = ""
    jsonString = jsonString + "{"
    jsonString = jsonString + "  ""brandList"": ["

    i = 0
    For Each brand In brandArray
        if Trim(brand) <> "" then
            if (i = 0) then
                jsonString = jsonString + "    {"
                jsonString = jsonString + "      ""brandCd"": """ & Trim(brand) & """"
                jsonString = jsonString + "    }"
            else
                jsonString = jsonString + "    ,{"
                jsonString = jsonString + "      ""brandCd"": """ & Trim(brand) & """"
                jsonString = jsonString + "    }"
            end if
            i = i + 1
        end if
    Next

    jsonString = jsonString + "  ],"
    jsonString = jsonString + "  ""skuList"": ["

    i = 0
    For Each skuCd In skuCdArray
        if Trim(skuCd) <> "" then
            if (i = 0) then
                jsonString = jsonString + "    {"
                jsonString = jsonString + "      ""skuCd"": """ & Trim(skuCd) & """"
                jsonString = jsonString + "    }"
            else
                jsonString = jsonString + "    ,{"
                jsonString = jsonString + "      ""skuCd"": """ & Trim(skuCd) & """"
                jsonString = jsonString + "    }"
            end if
            i = i + 1
        end if
    Next

    jsonString = jsonString + "  ]"
    jsonString = jsonString + "}"

    Call fnWriteLog(mode, jsonString)

    Call fnGetProductStockInfoList(method, jsonString, jsonResult)
end function

'// ============================================================================
'// AGV 오차 초기화
'// ============================================================================
function fnSendErrClear(skuCdArray, ByRef jsonResult)
    dim method, jsonString
    dim resultCode, resultMessage, failCode
    dim resultData, skuList, totalQty, item
    dim brand, skuCd, i

    method = "POST"

    jsonString = ""
    jsonString = jsonString + "{"
    jsonString = jsonString + "  ""skuList"": ["

    i = 0
    For Each skuCd In skuCdArray
        if Trim(skuCd) <> "" then
            if (i = 0) then
                jsonString = jsonString + "    {"
                jsonString = jsonString + "      ""skuCd"": """ & Trim(skuCd) & """"
                jsonString = jsonString + "    }"
            else
                jsonString = jsonString + "    ,{"
                jsonString = jsonString + "      ""skuCd"": """ & Trim(skuCd) & """"
                jsonString = jsonString + "    }"
            end if
            i = i + 1
        end if
    Next

    jsonString = jsonString + "  ]"
    jsonString = jsonString + "}"

    Call fnWriteLog(mode, jsonString)

    Call fnSendErrClearInfo(method, jsonString, jsonResult)
end function

'// ============================================================================
'// 재고정보 리턴 JSON 파싱함수
'// ============================================================================
function fnGetStockByResult(jsonResult)
    dim resultJson, resultData, totalQty, item

    totalQty = 0

    Set resultJson = New aspJson
    resultJson.loadJSON(jsonResult)

    Set resultData = resultJson.data("resultData")
    For Each item In resultData.item("skuList")
        totalQty = resultData.item("skuList").item(item).item("totalQty")
        Exit For
    Next

    fnGetStockByResult = totalQty
end function

'// ============================================================================
'// 브랜드정보/상품정보 AGV 전송함수
'// ============================================================================
function fnSendBrandItemInfo2AGV(requestMaster)
    dim uuid_obj, uuid, responseJson

	'// UUID 생성
	set uuid_obj = Server.CreateObject("Scriptlet.Typelib")
	uuid = uuid_obj.guid
	set uuid_obj = Nothing

	uuid = Mid(uuid, 2, 36)

	'// 상품정보/브랜드정보 전송 필요목록 생성
	sqlStr = " exec db_aLogistics.dbo.usp_AGV_GetBrandItemInfoNeedSend '" & requestMaster & "','" & uuid & "' "
	''response.write sqlStr
	dbget_Logistics.Execute sqlStr

	'// 브랜드정보 전송
	Call fnSendBrandInfoByUUID(uuid)

	sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_sendState_BrandInfo] "
	sqlStr = sqlStr + " set lastSendDate = getdate() "
	sqlStr = sqlStr + " where uuid = '" & uuid & "' "
	dbget_Logistics.Execute sqlStr

	'// 상품정보 전송
	Call fnSendProductInfoByUUID(uuid, "")

	sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] "
	sqlStr = sqlStr + " set lastSendDate = getdate() "
	sqlStr = sqlStr + " where uuid = '" & uuid & "' "
	dbget_Logistics.Execute sqlStr
end function

function fnResetLastSendDate(requestMaster)
    dim sqlStr

    sqlStr = " update i "
    sqlStr = sqlStr + " set i.lastSendDate = DateAdd(day, -1, getdate()) "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_scheduledItems] s "
    sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] i "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and s.requestMaster = '" & requestMaster & "' "
    sqlStr = sqlStr + " 		and s.itemGubun = i.itemGubun "
    sqlStr = sqlStr + " 		and s.itemid = i.itemid "
    sqlStr = sqlStr + " 		and s.itemoption = i.itemoption "
    sqlStr = sqlStr + " 		and DateDiff(day, i.lastSendDate, getdate()) <= 0 "
    dbget_Logistics.Execute sqlStr
end function

function fnChangeWarehouseToBulk(skuCdArr)
    dim uuid_obj, uuid
    dim skuCd, sqlStr, i

	'// UUID 생성
	set uuid_obj = Server.CreateObject("Scriptlet.Typelib")
	uuid = uuid_obj.guid
	set uuid_obj = Nothing

	uuid = Mid(uuid, 2, 36)

    skuCdArr = Split(skuCdArr, ",")
    for i = 0 to UBound(skuCdArr)
        skuCd = Trim(skuCdArr(i))

        if (skuCd <> "") then
            sqlStr = " update "
            sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] "
            sqlStr = sqlStr + "         set uuid = '" & uuid & "' "
            sqlStr = sqlStr + "         where skuCd = '" & skuCd & "' "
            dbget_Logistics.Execute sqlStr
        end if
    next

	'// 상품정보 전송
	Call fnSendProductInfoByUUID(uuid, "BLK")

	sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] "
	sqlStr = sqlStr + " set lastSendDate = getdate() "
	sqlStr = sqlStr + " where uuid = '" & uuid & "' "
	dbget_Logistics.Execute sqlStr
end function

'// ============================================================================
'// 재고정보 어드민 저장
'// ============================================================================
function fnSaveStockInfo(jsonData, updateRackde)
    dim uuid_obj, uuid
    dim sqlStr, resultJson
    dim resultData, item, totalQty, skuCd, adjustQty, warehouseCd
    dim valuesStr, i, tempSql

	'// UUID 생성
	set uuid_obj = Server.CreateObject("Scriptlet.Typelib")
	uuid = uuid_obj.guid
	set uuid_obj = Nothing

	uuid = Mid(uuid, 2, 36)

    sqlStr = " delete from [db_aLogistics].[dbo].[tbl_agv_stock_Info] where DateDiff(day, regdate, getdate()) > 0 "
    dbget_Logistics.Execute sqlStr

    sqlStr = " delete from [db_aLogistics].[dbo].[tbl_agv_stock_Info] where uuid = '" & uuid & "' "
    dbget_Logistics.Execute sqlStr

    Set resultJson = New aspJson
    resultJson.loadJSON(jsonData)

    i = 0
    sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_stock_Info](skuCd, uuid, totsysstock, errrealcheckno, currstockno, warehouseCd)"
    valuesStr = ""
    Set resultData = resultJson.data("resultData")
    For Each item In resultData.item("skuList")
        if (i >= 400) then
            '// values 가 1000개가 넘어가면 오류가 생긴다.
            tempSql = sqlStr + valuesStr
            dbget_Logistics.Execute tempSql

            valuesStr = ""
            i = 0
        end if

        skuCd = resultData.item("skuList").item(item).item("skuCd")
        totalQty = resultData.item("skuList").item(item).item("totalQty")			'// 시스템재고
        adjustQty = resultData.item("skuList").item(item).item("adjustQty")			'// 오차

        '// AGV_WAREHOUSE / BULK_WAREHOUSE
        warehouseCd = resultData.item("skuList").item(item).item("pickingWarehouseCd")
        if (warehouseCd = "BULK_WAREHOUSE") then
            warehouseCd = "BLK"
        else
            warehouseCd = "AGV"
        end if

        if (valuesStr = "") then
            valuesStr = "values('" & skuCd & "', '" & uuid & "', " & totalQty & ", " & adjustQty & ", " & totalQty + adjustQty & ", '" & warehouseCd & "')"
        else
            valuesStr = valuesStr + ",('" & skuCd & "', '" & uuid & "', " & totalQty & ", " & adjustQty & ", " & totalQty + adjustQty & ", '" & warehouseCd & "')"
        end if

        i = i + 1
    Next

    Set resultJson = Nothing

    if (valuesStr <> "") then
        sqlStr = sqlStr + valuesStr
        dbget_Logistics.Execute sqlStr
    end if

    sqlStr = " update a "
    sqlStr = sqlStr +  " set a.itemgubun = Left(a.skuCd, 2), a.itemid = substring(a.skuCd, 3, Len(a.skuCd) - 6), a.itemoption = right(a.skuCd,4) "
    sqlStr = sqlStr +  " from "
    sqlStr = sqlStr +  " [db_aLogistics].[dbo].[tbl_agv_stock_Info] a "
    sqlStr = sqlStr +  " where uuid = '" & uuid & "' "
    dbget_Logistics.Execute sqlStr

    sqlStr = " EXEC [db_summary].[dbo].[usp_AGV_GetStockData] '" & uuid & "' "
    dbget.Execute sqlStr

    if (updateRackde = True) then
        sqlStr = " exec db_aLogistics.dbo.usp_AGV_RecvItemRackcodeInfo_10x10 '" & uuid & "' "
        dbget_Logistics.Execute sqlStr
    end if
end function

'// ============================================================================
'// 주문정보 JSON 생성함수
'// ============================================================================
function fnGetOerInfoOnlineJSONBatch(baljuKey, siteBaljuid, ByRef jsonResult)
    dim sqlStr, skuCnt
    dim jsonString, i, skuListJson, orderListJson, orderserial, currorderserial

    '// 주문정보
	sqlStr = " select "
	sqlStr = sqlStr + " 	m.orderserial, format(m.regdate,'yyyy-MM-dd HH:mm:ss') as regdate, format(m.ipkumdate,'yyyy-MM-dd HH:mm:ss') as ipkumdate "
	sqlStr = sqlStr + " 	, (case when (m.sitename like 'its%') or (m.sitename like 'ithinksoshop%') then 'ithinkso' else '10x10' end) shopGroupCd "
	sqlStr = sqlStr + " 	, (case when m.DlvcountryCode = 'KR' then 'CJ' else 'EMS' end) shippingGroupCd "
	sqlStr = sqlStr + " 	, '<<' + m.orderserial + '_skuGroupCd>>' skuGroupCd "
	sqlStr = sqlStr + " 	, (case when IsNull(a.boxType, '') in ('ABC', 'ABCD', 'A1', 'B1', 'C1', 'D1', 'AB') then 'SMALL' else 'ETC' end) boxGroupCd "
	sqlStr = sqlStr + " 	, IsNull(a.boxType, 'NULL') as boxTypeCd "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_baljumaster] bm "
	sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_baljudetail] bd on bm.id = bd.baljuid "
	sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] m on bd.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_order_logics_add_info] a on a.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_order_detail] d on m.orderserial = d.orderserial "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and bm.id = " & siteBaljuid
	sqlStr = sqlStr + " 	and d.itemid <> 0 "
	sqlStr = sqlStr + " 	and d.isupchebeasong = 'N' "
	sqlStr = sqlStr + " group by "
	sqlStr = sqlStr + " 	m.orderserial, format(m.regdate,'yyyy-MM-dd HH:mm:ss'), format(m.ipkumdate,'yyyy-MM-dd HH:mm:ss') "
	sqlStr = sqlStr + " 	, (case when (m.sitename like 'its%') or (m.sitename like 'ithinksoshop%') then 'ithinkso' else '10x10' end) "
	sqlStr = sqlStr + " 	, (case when m.DlvcountryCode = 'KR' then 'CJ' else 'EMS' end) "
	sqlStr = sqlStr + " 	, m.orderserial + '_skuGroupCd' "
	sqlStr = sqlStr + " 	, (case when IsNull(a.boxType, '') in ('ABC', 'ABCD', 'A1', 'B1', 'C1', 'D1', 'AB') then 'SMALL' else 'ETC' end) "
	sqlStr = sqlStr + " 	, IsNull(a.boxType, 'NULL') "

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    i = 0
    if  not rsget.EOF  then
        do until rsget.eof
            jsonString = ""

            if (i = 0) then
                jsonString = jsonString + "    {"
                jsonString = jsonString + "      ""boxGroupCd"": """ & rsget("boxGroupCd") & ""","
                jsonString = jsonString + "      ""boxTypeCd"": """ & rsget("boxTypeCd") & ""","
                jsonString = jsonString + "      ""paymentDt"": """ & rsget("ipkumdate") & ""","
                jsonString = jsonString + "      ""salesDt"": """ & rsget("regdate") & ""","
                jsonString = jsonString + "      ""salesNo"": """ & rsget("orderserial") & ""","
                jsonString = jsonString + "      ""shippingGroupCd"": """ & rsget("shippingGroupCd") & ""","
                jsonString = jsonString + "      ""shopGroupCd"": """ & rsget("shopGroupCd") & ""","
                jsonString = jsonString + "      ""skuGroupCd"": """ & rsget("skuGroupCd") & ""","
                jsonString = jsonString + "      ""skuList"": ["
                jsonString = jsonString + "<<" + rsget("orderserial") + "_skuListJson>>"
                jsonString = jsonString + "<<" + rsget("orderserial") + "_skuGiftListJson>>"
                jsonString = jsonString + "      ]"
                jsonString = jsonString + "    }"
            else
                jsonString = jsonString + "    ,{"
                jsonString = jsonString + "      ""boxGroupCd"": """ & rsget("boxGroupCd") & ""","
                jsonString = jsonString + "      ""boxTypeCd"": """ & rsget("boxTypeCd") & ""","
                jsonString = jsonString + "      ""paymentDt"": """ & rsget("ipkumdate") & ""","
                jsonString = jsonString + "      ""salesDt"": """ & rsget("regdate") & ""","
                jsonString = jsonString + "      ""salesNo"": """ & rsget("orderserial") & ""","
                jsonString = jsonString + "      ""shippingGroupCd"": """ & rsget("shippingGroupCd") & ""","
                jsonString = jsonString + "      ""shopGroupCd"": """ & rsget("shopGroupCd") & ""","
                jsonString = jsonString + "      ""skuGroupCd"": """ & rsget("skuGroupCd") & ""","
                jsonString = jsonString + "      ""skuList"": ["
                jsonString = jsonString + "<<" + rsget("orderserial") + "_skuListJson>>"
                jsonString = jsonString + "<<" + rsget("orderserial") + "_skuGiftListJson>>"
                jsonString = jsonString + "      ]"
                jsonString = jsonString + "    }"
            end if

            orderListJson = orderListJson + jsonString

			i=i+1
			rsget.moveNext
		loop
    end if
    rsget.close

    '// 상품정보
	sqlStr = " select "
	sqlStr = sqlStr + " 	m.orderserial, d.itemgubun, d.itemid, d.itemoption, d.itemno, i.skuCd "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_Logistics_baljumaster] bm "
	sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_baljudetail] bd on bm.baljuKey = bd.baljuKey "
	sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_order_master] m on bd.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_order_detail] d on m.orderserial = d.orderserial "
	sqlStr = sqlStr + " 	left join [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] i "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and i.itemGubun = d.itemgubun "
	sqlStr = sqlStr + " 		and i.itemid = d.itemid "
	sqlStr = sqlStr + " 		and i.itemoption = d.itemoption "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and bm.baljuKey = " & baljuKey
	sqlStr = sqlStr + " 	and d.isupchebeasong = 'N' "
    sqlStr = sqlStr + " order by m.orderserial, d.itemgubun, d.itemid, d.itemoption "

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    if  not rsget_Logistics.EOF  then
        skuCnt = 0
        orderserial = ""
        currorderserial = ""
        skuListJson = ""
        jsonString = ""
        do until rsget_Logistics.eof
            currorderserial = rsget_Logistics("orderserial")

            if (currorderserial <> orderserial) then
                if (jsonString <> "") then
                    '// 생성된 데이타 있으면
                    if (skuCnt > 1) then
                        orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGroupCd>>", "COMPOUND")
                    else
                        orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGroupCd>>", "SINGLE")
                    end if
                    orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuListJson>>", jsonString)
                end if

                skuCnt = 0
                skuListJson = ""
                jsonString = ""
                orderserial = currorderserial
            end if

            if (skuCnt = 0) then
                jsonString = jsonString + "{"
                jsonString = jsonString + "  ""qty"": " & rsget_Logistics("itemno") & ","
                jsonString = jsonString + "  ""skuCd"": """ & rsget_Logistics("skuCd") & """"
                jsonString = jsonString + "}"
            else
                jsonString = jsonString + ",{"
                jsonString = jsonString + "  ""qty"": " & rsget_Logistics("itemno") & ","
                jsonString = jsonString + "  ""skuCd"": """ & rsget_Logistics("skuCd") & """"
                jsonString = jsonString + "}"
            end if

            skuCnt = skuCnt + 1

			i=i+1
			rsget_Logistics.moveNext
		loop
    end if
    rsget_Logistics.close

    if (jsonString <> "") then
        '// 생성된 데이타 있으면
        if (skuCnt > 1) then
            orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGroupCd>>", "COMPOUND")
        else
            orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGroupCd>>", "SINGLE")
        end if
        orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuListJson>>", jsonString)
    end if

    jsonString = ""

    '// 사은품정보
	sqlStr = " select o.orderserial, (k.prd_itemgubun + (case when k.prd_itemid < 1000000 then right(convert(varchar, (1000000 + k.prd_itemid)), 6) else right(convert(varchar, (100000000 + k.prd_itemid)), 8) end) +  k.prd_itemoption) as skucd, o.giftkind_cnt as itemno "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_baljumaster] bm "
    sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_baljudetail] bd on bm.id = bd.baljuid "
    sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_gift o on bd.orderserial = o.orderserial "
    sqlStr = sqlStr + " 	Join [db_event].[dbo].tbl_gift g on o.gift_code=g.gift_code "
    sqlStr = sqlStr + " 	left Join db_event.dbo.tbl_giftkind k on o.giftkind_code=k.giftkind_code "
    sqlStr = sqlStr + " 	left Join db_event.dbo.tbl_event e on g.evt_code=e.evt_code "
    sqlStr = sqlStr + " 	join [db_shop].[dbo].[tbl_shop_item] i "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and k.prd_itemgubun = i.itemgubun "
    sqlStr = sqlStr + " 		and k.prd_itemid = i.shopitemid "
    sqlStr = sqlStr + " 		and k.prd_itemoption = i.itemoption "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and bm.id = " & siteBaljuid
    sqlStr = sqlStr + " 	and o.gift_delivery = 'N' "
    sqlStr = sqlStr + " order by "
    sqlStr = sqlStr + " 	o.orderserial "

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    if  not rsget.EOF  then
        skuCnt = 0
        orderserial = ""
        currorderserial = ""
        skuListJson = ""
        jsonString = ""
        do until rsget.eof
            currorderserial = rsget("orderserial")

            if (currorderserial <> orderserial) then
                if (jsonString <> "") then
                    '// 생성된 데이타 있으면
                    orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGiftListJson>>", jsonString)
                end if

                skuCnt = 0
                skuListJson = ""
                jsonString = ""
                orderserial = currorderserial
            end if

            jsonString = jsonString + ",{"
            jsonString = jsonString + "  ""qty"": " & rsget("itemno") & ","
            jsonString = jsonString + "  ""skuCd"": """ & rsget("skuCd") & """"
            jsonString = jsonString + "}"

            skuCnt = skuCnt + 1

			i=i+1
			rsget.moveNext
		loop
    end if
    rsget.close

    if (jsonString <> "") then
        '// 생성된 데이타 있으면
        orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGiftListJson>>", jsonString)
    end if

    Dim regEx
    Set regEx = New RegExp

    With regEx
        .Pattern = "<<[0-9a-zA-z\s\-]+_skuGiftListJson>>"
        .IgnoreCase = True
        .Global = True
    End With
    orderListJson = regEx.Replace(orderListJson, "")

    With regEx
        .Pattern = "<<[0-9a-zA-z]+_skuGroupCd>>"
        .IgnoreCase = True
        .Global = True
    End With
    orderListJson = regEx.Replace(orderListJson, "복합")

    With regEx
        .Pattern = "<<[0-9a-zA-z]+_skuListJson>>"
        .IgnoreCase = True
        .Global = True
    End With
    orderListJson = regEx.Replace(orderListJson, "")

    set regEx = nothing

    jsonResult = orderListJson

end function

'// ============================================================================
'// 발주별 주문정보 JSON 생성함수
'// ============================================================================
function fnGetOerInfoOnlineJSONBatchBalju(baljuKey, siteBaljuid, ByRef jsonResult)
    dim sqlStr, skuCnt
    dim jsonString, i, skuListJson, orderListJson, orderserial, currorderserial

    '// 주문정보
	sqlStr = " select "
	sqlStr = sqlStr + " 	m.orderserial, format(m.regdate,'yyyy-MM-dd HH:mm:ss') as regdate, format(m.ipkumdate,'yyyy-MM-dd HH:mm:ss') as ipkumdate "
	sqlStr = sqlStr + " 	, (case when (m.sitename like 'its%') or (m.sitename like 'ithinksoshop%') then 'ithinkso' else '10x10' end) shopGroupCd "
	sqlStr = sqlStr + " 	, (case when m.DlvcountryCode = 'KR' then 'CJ' else 'EMS' end) shippingGroupCd "
	sqlStr = sqlStr + " 	, '<<' + m.orderserial + '_skuGroupCd>>' skuGroupCd "
	sqlStr = sqlStr + " 	, (case when IsNull(a.boxType, '') in ('ABC', 'ABCD', 'A1', 'B1', 'C1', 'D1', 'AB') then 'SMALL' else 'ETC' end) boxGroupCd "
	sqlStr = sqlStr + " 	, IsNull(a.boxType, 'NULL') as boxTypeCd "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_baljumaster] bm "
	sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_baljudetail] bd on bm.id = bd.baljuid "
	sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] m on bd.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_order_logics_add_info] a on a.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_order_detail] d on m.orderserial = d.orderserial "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and bm.id = " & siteBaljuid
	sqlStr = sqlStr + " 	and d.itemid <> 0 "
	sqlStr = sqlStr + " 	and d.isupchebeasong = 'N' "
    ''sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
    sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " group by "
	sqlStr = sqlStr + " 	m.orderserial, format(m.regdate,'yyyy-MM-dd HH:mm:ss'), format(m.ipkumdate,'yyyy-MM-dd HH:mm:ss') "
	sqlStr = sqlStr + " 	, (case when (m.sitename like 'its%') or (m.sitename like 'ithinksoshop%') then 'ithinkso' else '10x10' end) "
	sqlStr = sqlStr + " 	, (case when m.DlvcountryCode = 'KR' then 'CJ' else 'EMS' end) "
	sqlStr = sqlStr + " 	, m.orderserial + '_skuGroupCd' "
	sqlStr = sqlStr + " 	, (case when IsNull(a.boxType, '') in ('ABC', 'ABCD', 'A1', 'B1', 'C1', 'D1', 'AB') then 'SMALL' else 'ETC' end) "
	sqlStr = sqlStr + " 	, IsNull(a.boxType, 'NULL') "

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    i = 0
    if  not rsget.EOF  then
        do until rsget.eof
            jsonString = ""

            if (i = 0) then
                jsonString = jsonString + "     {"
                jsonString = jsonString + "      ""boxPosition"": ""<<" &  + rsget("orderserial") + "_boxPosition" & ">>"","
                jsonString = jsonString + "      ""boxTypeCd"": """ & rsget("boxTypeCd") & ""","
                jsonString = jsonString + "      ""caseFkNo"": """ & rsget("orderserial") & ""","
                jsonString = jsonString + "      ""caseName"": """ & rsget("orderserial") & ""","
                jsonString = jsonString + "      ""caseNo"": """ & rsget("orderserial") & ""","
                jsonString = jsonString + "      ""skuList"": ["
                jsonString = jsonString + "<<" + rsget("orderserial") + "_skuListJson>>"
                jsonString = jsonString + "<<" + rsget("orderserial") + "_skuGiftListJson>>"
                jsonString = jsonString + "      ]"
                jsonString = jsonString + "    }"
            else
                jsonString = jsonString + "     ,{"
                jsonString = jsonString + "      ""boxPosition"": ""<<" &  + rsget("orderserial") + "_boxPosition" & ">>"","
                jsonString = jsonString + "      ""boxTypeCd"": """ & rsget("boxTypeCd") & ""","
                jsonString = jsonString + "      ""caseFkNo"": """ & rsget("orderserial") & ""","
                jsonString = jsonString + "      ""caseName"": """ & rsget("orderserial") & ""","
                jsonString = jsonString + "      ""caseNo"": """ & rsget("orderserial") & ""","
                jsonString = jsonString + "      ""skuList"": ["
                jsonString = jsonString + "<<" + rsget("orderserial") + "_skuListJson>>"
                jsonString = jsonString + "<<" + rsget("orderserial") + "_skuGiftListJson>>"
                jsonString = jsonString + "      ]"
                jsonString = jsonString + "    }"
            end if

            orderListJson = orderListJson + jsonString

			i=i+1
			rsget.moveNext
		loop
    end if
    rsget.close

    '// 상품정보
	sqlStr = " select "
	sqlStr = sqlStr + " 	m.orderserial, d.itemgubun, d.itemid, d.itemoption, d.itemno, i.skuCd "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_Logistics_baljumaster] bm "
	sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_baljudetail] bd on bm.baljuKey = bd.baljuKey "
	sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_order_master] m on bd.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_order_detail] d on m.orderserial = d.orderserial "
	sqlStr = sqlStr + " 	left join [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] i "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and i.itemGubun = d.itemgubun "
	sqlStr = sqlStr + " 		and i.itemid = d.itemid "
	sqlStr = sqlStr + " 		and i.itemoption = d.itemoption "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and bm.baljuKey = " & baljuKey
	sqlStr = sqlStr + " 	and d.isupchebeasong = 'N' "
    sqlStr = sqlStr + " order by m.orderserial, d.itemgubun, d.itemid, d.itemoption "

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    if  not rsget_Logistics.EOF  then
        skuCnt = 0
        orderserial = ""
        currorderserial = ""
        skuListJson = ""
        jsonString = ""
        do until rsget_Logistics.eof
            currorderserial = rsget_Logistics("orderserial")

            if (currorderserial <> orderserial) then
                if (jsonString <> "") then
                    '// 생성된 데이타 있으면
                    if (skuCnt > 1) then
                        orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGroupCd>>", "COMPOUND")
                    else
                        orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGroupCd>>", "SINGLE")
                    end if
                    orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuListJson>>", jsonString)
                end if

                skuCnt = 0
                skuListJson = ""
                jsonString = ""
                orderserial = currorderserial
            end if

            if (skuCnt = 0) then
                jsonString = jsonString + "{"
                jsonString = jsonString + "  ""qty"": " & rsget_Logistics("itemno") & ","
                jsonString = jsonString + "  ""skuCd"": """ & rsget_Logistics("skuCd") & """"
                jsonString = jsonString + "}"
            else
                jsonString = jsonString + ",{"
                jsonString = jsonString + "  ""qty"": " & rsget_Logistics("itemno") & ","
                jsonString = jsonString + "  ""skuCd"": """ & rsget_Logistics("skuCd") & """"
                jsonString = jsonString + "}"
            end if

            skuCnt = skuCnt + 1

			i=i+1
			rsget_Logistics.moveNext
		loop
    end if
    rsget_Logistics.close

    if (jsonString <> "") then
        '// 생성된 데이타 있으면
        if (skuCnt > 1) then
            orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGroupCd>>", "COMPOUND")
        else
            orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGroupCd>>", "SINGLE")
        end if
        orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuListJson>>", jsonString)
    end if

    jsonString = ""

    '// 사은품정보
	sqlStr = " select o.orderserial, (k.prd_itemgubun + (case when k.prd_itemid < 1000000 then right(convert(varchar, (1000000 + k.prd_itemid)), 6) else right(convert(varchar, (100000000 + k.prd_itemid)), 8) end) +  k.prd_itemoption) as skucd, o.giftkind_cnt as itemno "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_baljumaster] bm "
    sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_baljudetail] bd on bm.id = bd.baljuid "
    sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_gift o on bd.orderserial = o.orderserial "
    sqlStr = sqlStr + " 	Join [db_event].[dbo].tbl_gift g on o.gift_code=g.gift_code "
    sqlStr = sqlStr + " 	left Join db_event.dbo.tbl_giftkind k on o.giftkind_code=k.giftkind_code "
    sqlStr = sqlStr + " 	left Join db_event.dbo.tbl_event e on g.evt_code=e.evt_code "
    sqlStr = sqlStr + " 	join [db_shop].[dbo].[tbl_shop_item] i "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and k.prd_itemgubun = i.itemgubun "
    sqlStr = sqlStr + " 		and k.prd_itemid = i.shopitemid "
    sqlStr = sqlStr + " 		and k.prd_itemoption = i.itemoption "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and bm.id = " & siteBaljuid
    sqlStr = sqlStr + " 	and o.gift_delivery = 'N' "
    sqlStr = sqlStr + " order by "
    sqlStr = sqlStr + " 	o.orderserial "

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    if  not rsget.EOF  then
        skuCnt = 0
        orderserial = ""
        currorderserial = ""
        skuListJson = ""
        jsonString = ""
        do until rsget.eof
            currorderserial = rsget("orderserial")

            if (currorderserial <> orderserial) then
                if (jsonString <> "") then
                    '// 생성된 데이타 있으면
                    orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGiftListJson>>", jsonString)
                end if

                skuCnt = 0
                skuListJson = ""
                jsonString = ""
                orderserial = currorderserial
            end if

            jsonString = jsonString + ",{"
            jsonString = jsonString + "  ""qty"": " & rsget("itemno") & ","
            jsonString = jsonString + "  ""skuCd"": """ & rsget("skuCd") & """"
            jsonString = jsonString + "}"

            skuCnt = skuCnt + 1

			i=i+1
			rsget.moveNext
		loop
    end if
    rsget.close

    if (jsonString <> "") then
        '// 생성된 데이타 있으면
        orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGiftListJson>>", jsonString)
    end if

    '// 박스위치
    sqlStr = " select orderserial, dasindex "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_baljudetail] bd "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and LocalDlvInclude = 1 "
    sqlStr = sqlStr + " 	and baljuKey = " & baljuKey

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    if  not rsget_Logistics.EOF  then
        do until rsget_Logistics.eof
            orderserial = rsget_Logistics("orderserial")
            dasindex = rsget_Logistics("dasindex")

            orderListJson = Replace(orderListJson, "<<" & orderserial & "_boxPosition>>", dasindex)

			rsget_Logistics.moveNext
		loop
    end if
    rsget_Logistics.close

    Dim regEx
    Set regEx = New RegExp

    With regEx
        .Pattern = "<<[0-9a-zA-z\s\-]+_skuGiftListJson>>"
        .IgnoreCase = True
        .Global = True
    End With
    orderListJson = regEx.Replace(orderListJson, "")

    With regEx
        .Pattern = "<<[0-9a-zA-z]+_skuGroupCd>>"
        .IgnoreCase = True
        .Global = True
    End With
    orderListJson = regEx.Replace(orderListJson, "복합")

    With regEx
        .Pattern = "<<[0-9a-zA-z]+_skuListJson>>"
        .IgnoreCase = True
        .Global = True
    End With
    orderListJson = regEx.Replace(orderListJson, "")

    With regEx
        .Pattern = "<<[0-9a-zA-z]+_boxPosition>>"
        .IgnoreCase = True
        .Global = True
    End With
    orderListJson = regEx.Replace(orderListJson, "999")

    set regEx = nothing

    jsonResult = orderListJson

end function

function fnGetOerInfoOnlineJSONBatchBaljuMulti(baljuKey, ByRef jsonResult)
    dim sqlStr, skuCnt
    dim jsonString, i, skuListJson, orderListJson, orderserial, currorderserial

    '// 주문정보
	sqlStr = " select "
	''sqlStr = sqlStr + " 	m.orderserial, format(m.regdate,'yyyy-MM-dd HH:mm:ss') as regdate, format(m.ipkumdate,'yyyy-MM-dd HH:mm:ss') as ipkumdate "
    sqlStr = sqlStr + " 	right(m.orderserial, 1) as orderserial, format(m.regdate,'yyyy-MM-dd HH:mm:ss') as regdate, format(m.ipkumdate,'yyyy-MM-dd HH:mm:ss') as ipkumdate "
	sqlStr = sqlStr + " 	, '10x10' shopGroupCd "
	sqlStr = sqlStr + " 	, 'CJ' shippingGroupCd "
	sqlStr = sqlStr + " 	, '<<' + m.orderserial + '_skuGroupCd>>' skuGroupCd "
	sqlStr = sqlStr + " 	, 'ETC' boxGroupCd "
	sqlStr = sqlStr + " 	, 'NULL' as boxTypeCd "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_Logistics_agv_baljumaster] bm "
	sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_agv_baljudetail] bd on bm.baljukey = bd.baljukey "
	sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_agv_order_master] m on bd.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_agv_order_detail] d on m.orderserial = d.orderserial "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and bm.baljukey = " & baljuKey
	sqlStr = sqlStr + " 	and d.itemid <> 0 "
	sqlStr = sqlStr + " 	and d.isupchebeasong = 'N' "
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " group by "
	sqlStr = sqlStr + " 	m.orderserial, format(m.regdate,'yyyy-MM-dd HH:mm:ss'), format(m.ipkumdate,'yyyy-MM-dd HH:mm:ss') "
	sqlStr = sqlStr + " 	, m.orderserial + '_skuGroupCd' "

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    i = 0
    if  not rsget_Logistics.EOF  then
        do until rsget_Logistics.eof
            jsonString = ""

            if (i = 0) then
                jsonString = jsonString + "     {"
                jsonString = jsonString + "      ""boxPosition"": ""<<" &  + rsget_Logistics("orderserial") + "_boxPosition" & ">>"","
                jsonString = jsonString + "      ""boxTypeCd"": """ & rsget_Logistics("boxTypeCd") & ""","
                jsonString = jsonString + "      ""caseFkNo"": """ & rsget_Logistics("orderserial") & ""","
                jsonString = jsonString + "      ""caseName"": """ & rsget_Logistics("orderserial") & ""","
                jsonString = jsonString + "      ""caseNo"": """ & rsget_Logistics("orderserial") & ""","
                jsonString = jsonString + "      ""skuList"": ["
                jsonString = jsonString + "<<" + rsget_Logistics("orderserial") + "_skuListJson>>"
                jsonString = jsonString + "<<" + rsget_Logistics("orderserial") + "_skuGiftListJson>>"
                jsonString = jsonString + "      ]"
                jsonString = jsonString + "    }"
            else
                jsonString = jsonString + "     ,{"
                jsonString = jsonString + "      ""boxPosition"": ""<<" &  + rsget_Logistics("orderserial") + "_boxPosition" & ">>"","
                jsonString = jsonString + "      ""boxTypeCd"": """ & rsget_Logistics("boxTypeCd") & ""","
                jsonString = jsonString + "      ""caseFkNo"": """ & rsget_Logistics("orderserial") & ""","
                jsonString = jsonString + "      ""caseName"": """ & rsget_Logistics("orderserial") & ""","
                jsonString = jsonString + "      ""caseNo"": """ & rsget_Logistics("orderserial") & ""","
                jsonString = jsonString + "      ""skuList"": ["
                jsonString = jsonString + "<<" + rsget_Logistics("orderserial") + "_skuListJson>>"
                jsonString = jsonString + "<<" + rsget_Logistics("orderserial") + "_skuGiftListJson>>"
                jsonString = jsonString + "      ]"
                jsonString = jsonString + "    }"
            end if

            orderListJson = orderListJson + jsonString

			i=i+1
			rsget_Logistics.moveNext
		loop
    end if
    rsget_Logistics.close

    '// 상품정보
	sqlStr = " select "
	''sqlStr = sqlStr + " 	m.orderserial, d.itemgubun, d.itemid, d.itemoption, d.itemno, i.skuCd "
    sqlStr = sqlStr + " 	right(m.orderserial, 1) as orderserial, d.itemgubun, d.itemid, d.itemoption, d.itemno, i.skuCd "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_Logistics_agv_baljumaster] bm "
	sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_agv_baljudetail] bd on bm.baljuKey = bd.baljuKey "
	sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_agv_order_master] m on bd.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_agv_order_detail] d on m.orderserial = d.orderserial "
	sqlStr = sqlStr + " 	left join [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] i "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and i.itemGubun = d.itemgubun "
	sqlStr = sqlStr + " 		and i.itemid = d.itemid "
	sqlStr = sqlStr + " 		and i.itemoption = d.itemoption "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and bm.baljuKey = " & baljuKey
	sqlStr = sqlStr + " 	and d.isupchebeasong = 'N' "
	sqlStr = sqlStr + " order by m.orderserial, d.itemgubun, d.itemid, d.itemoption "

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    if  not rsget_Logistics.EOF  then
        skuCnt = 0
        orderserial = ""
        currorderserial = ""
        skuListJson = ""
        jsonString = ""
        do until rsget_Logistics.eof
            currorderserial = rsget_Logistics("orderserial")

            if (currorderserial <> orderserial) then
                if (jsonString <> "") then
                    '// 생성된 데이타 있으면
                    if (skuCnt > 1) then
                        orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGroupCd>>", "COMPOUND")
                    else
                        orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGroupCd>>", "SINGLE")
                    end if
                    orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuListJson>>", jsonString)
                end if

                skuCnt = 0
                skuListJson = ""
                jsonString = ""
                orderserial = currorderserial
            end if

            if (skuCnt = 0) then
                jsonString = jsonString + "{"
                jsonString = jsonString + "  ""qty"": " & rsget_Logistics("itemno") & ","
                jsonString = jsonString + "  ""skuCd"": """ & rsget_Logistics("skuCd") & """"
                jsonString = jsonString + "}"
            else
                jsonString = jsonString + ",{"
                jsonString = jsonString + "  ""qty"": " & rsget_Logistics("itemno") & ","
                jsonString = jsonString + "  ""skuCd"": """ & rsget_Logistics("skuCd") & """"
                jsonString = jsonString + "}"
            end if

            skuCnt = skuCnt + 1

			i=i+1
			rsget_Logistics.moveNext
		loop
    end if
    rsget_Logistics.close

    if (jsonString <> "") then
        '// 생성된 데이타 있으면
        if (skuCnt > 1) then
            orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGroupCd>>", "COMPOUND")
        else
            orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuGroupCd>>", "SINGLE")
        end if
        orderListJson = Replace(orderListJson, "<<" & orderserial & "_skuListJson>>", jsonString)
    end if

    '// 박스위치
    sqlStr = " select right(bd.orderserial, 1) as orderserial, dasindex "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_agv_baljudetail] bd "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and LocalDlvInclude = 1 "
    sqlStr = sqlStr + " 	and baljuKey = " & baljuKey

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    if  not rsget_Logistics.EOF  then
        do until rsget_Logistics.eof
            orderserial = rsget_Logistics("orderserial")
            dasindex = rsget_Logistics("dasindex")

            orderListJson = Replace(orderListJson, "<<" & orderserial & "_boxPosition>>", dasindex)

			rsget_Logistics.moveNext
		loop
    end if
    rsget_Logistics.close

    Dim regEx
    Set regEx = New RegExp

    With regEx
        .Pattern = "<<[0-9a-zA-z\s\-]+_skuGiftListJson>>"
        .IgnoreCase = True
        .Global = True
    End With
    orderListJson = regEx.Replace(orderListJson, "")

    With regEx
        .Pattern = "<<[0-9a-zA-z\-]+_skuGroupCd>>"
        .IgnoreCase = True
        .Global = True
    End With
    orderListJson = regEx.Replace(orderListJson, "복합")

    With regEx
        .Pattern = "<<[0-9a-zA-z\-]+_skuListJson>>"
        .IgnoreCase = True
        .Global = True
    End With
    orderListJson = regEx.Replace(orderListJson, "")

    ''With regEx
    ''    .Pattern = "<<[0-9a-zA-z\-]+_boxPosition>>"
    ''    .IgnoreCase = True
    ''    .Global = True
    ''End With
    ''orderListJson = regEx.Replace(orderListJson, "999")

    set regEx = nothing

    jsonResult = orderListJson
end function

function fnGetOerInfoOfflineJSONBatchBalju(baljuKey, ByRef jsonResult)
    dim sqlStr, skuCnt
    dim jsonString, i, skuListJson, orderListJson, orderserial, currorderserial

    '// 매장 주문은 주문내역 한건으로 합친다.

    jsonString = ""
    jsonString = jsonString + "     {"
    jsonString = jsonString + "      ""boxPosition"": ""1"","
    jsonString = jsonString + "      ""boxTypeCd"": ""NULL"","
    jsonString = jsonString + "      ""caseFkNo"": """ & "ORDEROFFLINE(" & baljuKey & ")" & ""","
    jsonString = jsonString + "      ""caseName"": """ & "ORDEROFFLINE(" & baljuKey & ")" & ""","
    jsonString = jsonString + "      ""caseNo"": """ & "ORDEROFFLINE(" & baljuKey & ")" & ""","
    jsonString = jsonString + "      ""skuList"": ["
    jsonString = jsonString + "<<" & baljuKey & "_skuListJson>>"
    jsonString = jsonString + "      ]"
    jsonString = jsonString + "    }"

    sqlStr = "     select b.itemGubun, b.itemid, b.itemoption, (b.itemGubun + (case when b.itemid < 1000000 then right(convert(varchar, (1000000 + b.itemid)), 6) else right(convert(varchar, (100000000 + b.itemid)), 8) end) +  b.itemoption) as skucd, b.baljuno as itemno "
    sqlStr = sqlStr + "     from "
    sqlStr = sqlStr + "     	[db_aLogistics].[dbo].[tbl_Logistics_offline_baljuipgo] b "
    sqlStr = sqlStr + "     where b.baljuKey = " & baljuKey

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    skuListJson = ""
    if  not rsget_Logistics.EOF  then
        skuCnt = 0
        i = 0
        do until rsget_Logistics.eof
            if (skuCnt = 0) then
                skuListJson = skuListJson + "{"
                skuListJson = skuListJson + "  ""qty"": " & rsget_Logistics("itemno") & ","
                skuListJson = skuListJson + "  ""skuCd"": """ & rsget_Logistics("skuCd") & """"
                skuListJson = skuListJson + "}"
            else
                skuListJson = skuListJson + ",{"
                skuListJson = skuListJson + "  ""qty"": " & rsget_Logistics("itemno") & ","
                skuListJson = skuListJson + "  ""skuCd"": """ & rsget_Logistics("skuCd") & """"
                skuListJson = skuListJson + "}"
            end if

            skuCnt = skuCnt + 1

			i=i+1
			rsget_Logistics.moveNext
		loop
    end if
    rsget_Logistics.close

    Dim regEx
    Set regEx = New RegExp

    With regEx
        .Pattern = "<<[0-9a-zA-z]+_skuListJson>>"
        .IgnoreCase = True
        .Global = True
    End With
    orderListJson = regEx.Replace(jsonString, skuListJson)

    jsonResult = orderListJson

end function

'// ============================================================================
'// 주문정보(한건) JSON 생성함수
'// ============================================================================
function fnGetOerInfoOnlineJSON(baljuKey, orderserial, ByRef jsonResult)
    dim sqlStr, skuCnt
    dim jsonString, i, skuListJson

    '// 사용하지 말것!!
    response.end
end function

'// ============================================================================
'// 주문정보 전송함수
'// ============================================================================
function fnSendOrderInfoOnlineRequest(method, jsonString, ByRef jsonResult)
    dim url : url = "api/site/salesAdd"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// 발주정보 전송함수
'// ============================================================================
function fnSendOrderInfoOnlineBaljuRequest(method, jsonString, ByRef jsonResult)
    dim url : url = "api/site/batchAdd"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// 발주취소요청 전송함수
'// ============================================================================
function fnSendOrderInfoBaljuCancelRequest(method, jsonString, ByRef jsonResult)
    dim url : url = "api/site/batchCancel"

    Call fnSendReceiveJson(url, method, jsonString, jsonResult)
end function

'// ============================================================================
'// 주문정보
'// ============================================================================
function fnSendOrderInfoOnline(jsonString, ByRef jsonResult)
    dim sqlStr, i
    dim method
    dim resultCode, resultMessage, failCode

    method = "POST"

    Call fnSendOrderInfoOnlineRequest(method, jsonString, jsonResult)
end function

'// ============================================================================
'// 발주정보
'// ============================================================================
function fnSendOrderInfoOnlineBalju(jsonString, ByRef jsonResult)
    dim sqlStr, i
    dim method
    dim resultCode, resultMessage, failCode

    method = "POST"

    Call fnSendOrderInfoOnlineBaljuRequest(method, jsonString, jsonResult)
end function

'// ============================================================================
'// 발주취소정보
'// ============================================================================
function fnSendOrderInfoBaljuCancel(jsonString, ByRef jsonResult)
    dim sqlStr, i
    dim method
    dim resultCode, resultMessage, failCode

    method = "POST"

    Call fnSendOrderInfoBaljuCancelRequest(method, jsonString, jsonResult)
end function


select case mode
    case "agvSendErrClear":
        '// AGV 오차 초기화
        skuCdArray = Split(skuCdArray, ",")
        Call fnSendErrClear(skuCdArray, jsonResult)

        Call fnParseResultJson(jsonResult, resultCode, resultMessage, failCode)

        if (resultCode = "00") then
            responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        else
            responseJson = fnCreateCustomJsonResult("500", resultMessage, "500")
        end if

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "agvipgo":
        '// AGV 진열등록
        Call fnSendBrandItemInfo2AGV(requestMaster)

        requestNoArr = ""

        sqlStr = " select s.IDX as requestNo "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_scheduledItems] s "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and s.requestMaster = '" & requestMaster & "' "
        sqlStr = sqlStr + " 	and s.status = 0 "
        sqlStr = sqlStr + " 	and s.isUsing = 'Y' "

        rsget_Logistics.CursorLocation = adUseClient
        rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

        if  not rsget_Logistics.EOF  then
            do until rsget_Logistics.eof
                requestNoArr = requestNoArr & "|" & rsget_Logistics("requestNo")
			    rsget_Logistics.moveNext
		    loop
        end if
        rsget_Logistics.close

        responseJson = ""
        requestNoArr = Split(requestNoArr, "|")
        for i = 0 to UBound(requestNoArr)
            if Trim(requestNoArr(i)) <> "" then
                requestNo = Trim(requestNoArr(i))
                Call fnSendDisplayByRequestNo(requestNo, jsonResult)

                Call fnParseResultDisplayJson(jsonResult, resultCode, resultMessage, failCode, displayOrderNo)

                if (resultCode = "00") then
		            sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_scheduledItems] "
                    sqlStr = sqlStr + " set displayOrderNo = '" & displayOrderNo & "', status = 50 "
                    sqlStr = sqlStr + " where idx = '" & requestNo & "' and status = 0 "
                    dbget_Logistics.Execute sqlStr
                else
                    responseJson = fnCreateCustomJsonResult("500", resultMessage, "500")
                    exit for
                end if
            end if
        next

        if (responseJson = "") then
            responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        end if

        Response.ContentType = "application/json; charset=utf-8"
        response.write responseJson
    case "agvipgodel":
        '// AGV 진열등록삭제
        sqlStr = " select s.IDX as requestNo, displayOrderNo "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_scheduledItems] s "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and s.requestMaster = '" & requestMaster & "' "
        sqlStr = sqlStr + " 	and s.status = 50 "
        sqlStr = sqlStr + " 	and s.isUsing = 'Y' "
        sqlStr = sqlStr + " 	and s.displayOrderNo is not NULL "

        rsget_Logistics.CursorLocation = adUseClient
        rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

        if  not rsget_Logistics.EOF  then
            do until rsget_Logistics.eof
                requestNoArr = requestNoArr & "|" & rsget_Logistics("requestNo")
                displayOrderNoArr = displayOrderNoArr & "|" & rsget_Logistics("displayOrderNo")
			    rsget_Logistics.moveNext
		    loop
        end if
        rsget_Logistics.close

        responseJson = ""
        requestNoArr = Split(requestNoArr, "|")
        displayOrderNoArr = Split(displayOrderNoArr, "|")
        for i = 0 to UBound(requestNoArr)
            if Trim(requestNoArr(i)) <> "" then
                requestNo = Trim(requestNoArr(i))
                displayOrderNo = Trim(displayOrderNoArr(i))

                Call fnSendCancelDisplayByDisplayOrderNo(displayOrderNo, jsonResult)

                Call fnParseResultDisplayJson(jsonResult, resultCode, resultMessage, failCode, displayOrderNo)

                if (resultCode = "00") then
		            sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_scheduledItems] "
                    sqlStr = sqlStr + " set isusing = 'N', status = 10, lastupdate=getdate() "
                    sqlStr = sqlStr + " where idx = '" & requestNo & "' and status = 50 "
                    dbget_Logistics.Execute sqlStr
                else
                    responseJson = fnCreateCustomJsonResult("500", resultMessage, "500")
                    exit for
                end if
            end if
        next

        if (responseJson = "") then
            responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        end if

        Response.ContentType = "application/json; charset=utf-8"
        response.write responseJson
    case "agvpickCS":
        '// CS 피킹 지시등록
        sqlStr = " SELECT "
        sqlStr = sqlStr + " 	('10' + (case when d.itemid < 1000000 then right(convert(varchar, (1000000 + d.itemid)), 6) else right(convert(varchar, (100000000 + d.itemid)), 8) end) +  d.itemoption) as skucd "
        sqlStr = sqlStr + " 	,Sum(d.regitemno) AS totalQty "
        sqlStr = sqlStr + " FROM "
        sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list c "
        sqlStr = sqlStr + " 	LEFT JOIN [db_cs].[dbo].tbl_new_as_detail d ON c.id = d.masterid "
        sqlStr = sqlStr + " 	LEFT JOIN [db_item].[dbo].tbl_item i ON d.itemid = i.itemid "
        sqlStr = sqlStr + " WHERE "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	AND c.regdate >= DateAdd(day, -60, getdate()) "
        sqlStr = sqlStr + " 	AND c.regdate < DateAdd(day, -1, getdate()) "
        sqlStr = sqlStr + " 	AND c.deleteyn = 'N' "
        sqlStr = sqlStr + " 	AND c.currstate = 'B001' "
        sqlStr = sqlStr + " 	AND IsNull(c.makerid, '') = '' "
        sqlStr = sqlStr + " 	AND c.divcd IN ('A000','A100','A001','A002') "
        sqlStr = sqlStr + " 	AND NOT ( "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		AND (c.divcd IN ('A000','A100')) "
        sqlStr = sqlStr + " 		AND (c.gubun02 IN ('CD01','CD04','CD06','CD08')) "
        sqlStr = sqlStr + " 		AND ( "
        sqlStr = sqlStr + " 			SELECT count(*) "
        sqlStr = sqlStr + " 			FROM [db_cs].[dbo].tbl_new_as_list r "
        sqlStr = sqlStr + " 			WHERE "
        sqlStr = sqlStr + " 				1 = 1 "
        sqlStr = sqlStr + " 				AND r.refasid = c.id "
        sqlStr = sqlStr + " 				AND r.currstate = 'B007' "
        sqlStr = sqlStr + " 			) = 0 "
        sqlStr = sqlStr + " 	) "
        sqlStr = sqlStr + " 	AND IsNull(d.regitemno, 0) >= 0 "
        sqlStr = sqlStr + " GROUP BY "
        sqlStr = sqlStr + " 	d.itemid "
        sqlStr = sqlStr + " 	,d.itemoption "
        sqlStr = sqlStr + " 	,d.itemname "
        sqlStr = sqlStr + " 	,d.itemoptionname "
        sqlStr = sqlStr + " 	,i.itemrackcode "
        sqlStr = sqlStr + " 	,i.makerid "
        sqlStr = sqlStr + " 	,i.smallimage "
        sqlStr = sqlStr + " ORDER BY "
        sqlStr = sqlStr + " 	d.itemid "
        sqlStr = sqlStr + " 	,d.itemoption "

        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        skuCdArr = ""
	    if  not rsget.EOF  then
		    do until rsget.eof
                skuCdArr = skuCdArr & "|" & rsget("skuCd")
                totalQtyArr = totalQtyArr & "|" & rsget("totalQty")
			    rsget.moveNext
		    loop
	    end if
	    rsget.close

        requestNo = "PICKCS(" & getDateFormatedWithDash(Now()) & ")"
        pickingOrderTypeCd = "CS피킹"
        if (stationCd = "") then
            stationCd = "WS0001"
        end if

        dataCount = 0
        skuCdArr = Split(skuCdArr, "|")
        totalQtyArr = Split(totalQtyArr, "|")

        for i = 0 to UBound(skuCdArr)
            if Trim(skuCdArr(i)) <> "" then
                skuCd = Trim(skuCdArr(i))
                totalQty = Trim(totalQtyArr(i))

                sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_pickupItems](itemGubun, itemid, itemoption, itemno, requestNo, pickingOrderTypeCd, stationCd) "
                sqlStr = sqlStr + " select itemGubun, itemid, itemoption, " & totalQty & ", '" & requestNo & "', '" & pickingOrderTypeCd & "', '" & stationCd & "' "
                sqlStr = sqlStr + " from [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] "
                sqlStr = sqlStr + " where skuCd = '" & skuCd & "' "
                dbget_Logistics.Execute sqlStr, affectedRows

                dataCount = dataCount + affectedRows
            end if
        next

        if (dataCount > 0) then
            sqlStr = " if not exists(select idx from db_aLogistics.dbo.tbl_agv_scheduledItems where requestMaster = '" & requestNo & "' and isusing = 'Y') "
            sqlStr = sqlStr + " begin "
            sqlStr = sqlStr + "     insert into db_aLogistics.dbo.tbl_agv_scheduledItems(itemgubun, itemid, itemoption, realStock, displayOrderTypeCd, requestMaster) "
            sqlStr = sqlStr + "     select b.itemGubun, b.itemid, b.itemoption, 0, '상품정보전송', '" & requestNo & "' "
            sqlStr = sqlStr + "     from "
            sqlStr = sqlStr + "     	[db_aLogistics].[dbo].[tbl_agv_pickupItems] b "
            sqlStr = sqlStr + "     where b.requestNo = '" & requestNo & "'"
            sqlStr = sqlStr + " end "
            dbget_Logistics.Execute sqlStr, affectedRows

            if (affectedRows > 0) then
                Call fnSendBrandItemInfo2AGV(requestNo)
            end if

            Call fnSendPickingByRequestNo(requestNo, jsonResult)

            Call fnParseResultPickingJson(jsonResult, resultCode, resultMessage, failCode, pickingOrderNo)

            if (resultCode = "00") then
		        sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickupItems] "
                sqlStr = sqlStr + " set pickingOrderNo = '" & pickingOrderNo & "', status = 50 "
                sqlStr = sqlStr + " where requestNo = '" & requestNo & "' and status = 0 "
                dbget_Logistics.Execute sqlStr

                responseJson = fnCreateCustomJsonResult("200", "OK", "200")
            else
                responseJson = fnCreateCustomJsonResult(resultCode, resultMessage, failCode)
            end if
        else
            responseJson = fnCreateCustomJsonResult("500", "AGV 에 진열된 상품이 없습니다.", "500")
        end if

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "agvpickCSdel":
        '// CS 피킹 지시삭제
        response.write "작업중 : " & mode
    case "agvpickup":
        '// 피킹 지시등록
        sqlStr = " select top 1 * "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
        sqlStr = sqlStr + " where idx = " & masteridx

        rsget_Logistics.CursorLocation = adUseClient
        rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

        requestNo = ""
	    if  not rsget_Logistics.EOF  then
            status = rsget_Logistics("status")
            stationCd = rsget_Logistics("stationCd")

            requestNo = rsget_Logistics("requestNo")
            if IsNull(requestNo) then
                requestNo = "PICKUP(" & masteridx & ")"
            end if
            pickingOrderTypeCd = "INTERFACE"
	    end if
	    rsget_Logistics.close

        if (requestNo = "") then
            responseJson = fnCreateCustomJsonResult("500", "내역이 존재하지 않습니다.[" & masteridx & "]", "500")
            Response.ContentType = "application/json; charset=utf-8"
            Call Response.AddHeader("Access-Control-Allow-Origin", "*")
            response.write responseJson

            dbget.Close
            dbget_Logistics.Close
            response.end
        end if

        if Not IsNull(status) then
            if Clng(status) >= 50 then
                responseJson = fnCreateCustomJsonResult("500", "이미 전송완료된 내역입니다.[" & masteridx & "]", "500")
                Response.ContentType = "application/json; charset=utf-8"
                Call Response.AddHeader("Access-Control-Allow-Origin", "*")
                response.write responseJson

                dbget.Close
                dbget_Logistics.Close
                response.end
            end if
        end if

		sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
        sqlStr = sqlStr + " set pickingOrderTypeCd = '" & pickingOrderTypeCd & "', requestNo = '" & requestNo & "' "
        sqlStr = sqlStr + " where idx = '" & masteridx & "' and IsNull(status,0) < 50 "
        dbget_Logistics.Execute sqlStr


        '// 상품정보 전송
        sqlStr = " update db_aLogistics.dbo.tbl_agv_scheduledItems "
        sqlStr = sqlStr + " set isusing = 'N' "
        sqlStr = sqlStr + " where requestMaster = '" & requestNo & "' and isusing = 'Y' "
        dbget_Logistics.Execute sqlStr

        sqlStr = " if not exists(select idx from db_aLogistics.dbo.tbl_agv_scheduledItems where requestMaster = '" & requestNo & "' and isusing = 'Y') "
        sqlStr = sqlStr + " begin "
        sqlStr = sqlStr + "     insert into db_aLogistics.dbo.tbl_agv_scheduledItems(itemgubun, itemid, itemoption, realStock, displayOrderTypeCd, requestMaster) "
        sqlStr = sqlStr + "     select b.itemGubun, b.itemid, b.itemoption, 0, '상품정보전송', '" & requestNo & "' "
        sqlStr = sqlStr + "     from "
        sqlStr = sqlStr + "     	[db_aLogistics].[dbo].[tbl_agv_pickup_detail] b "
        sqlStr = sqlStr + "     where b.masteridx = '" & masteridx & "' and deldt is NULL "
        sqlStr = sqlStr + " end "
        dbget_Logistics.Execute sqlStr, affectedRows

        if (affectedRows > 0) then
            Call fnSendBrandItemInfo2AGV(requestNo)

            Call fnSendPickingByMasterIDX(masteridx, jsonResult)

            Call fnParseResultPickingJson(jsonResult, resultCode, resultMessage, failCode, pickingOrderNo)

            if (resultCode = "00") then
		        sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
                sqlStr = sqlStr + " set pickingOrderNo = '" & pickingOrderNo & "', status = 50 "
                sqlStr = sqlStr + " where idx = '" & masteridx & "' and IsNull(status,0) < 50 "
                dbget_Logistics.Execute sqlStr

                responseJson = fnCreateCustomJsonResult("200", "OK", "200")
            else
                responseJson = fnCreateCustomJsonResult(resultCode, resultMessage, failCode)
            end if
        else
            responseJson = fnCreateCustomJsonResult("500", "피킹지시할 상품이 없습니다.", "500")
        end if

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "agvpickupdel":
        '// masteridx
        Call fnSendPickingCancelByMasterIDX(masteridx, jsonResult)

        Call fnParseResultJson(jsonResult, resultCode, resultMessage, failCode)

        if (resultCode = "00") then
		    sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
            sqlStr = sqlStr + " set status = 10 "
            sqlStr = sqlStr + " where idx = '" & masteridx & "' "
            dbget_Logistics.Execute sqlStr

            responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        else
            responseJson = fnCreateCustomJsonResult(resultCode, resultMessage, failCode)
        end if

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "agvstockinvest":
        '// 재고조사 지시등록
        sqlStr = " select top 1 * "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_agv_stock_invest_master] "
        sqlStr = sqlStr + " where idx = " & masteridx

        rsget_Logistics.CursorLocation = adUseClient
        rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

        requestNo = ""
	    if  not rsget_Logistics.EOF  then
            status = rsget_Logistics("status")
            stationCd = rsget_Logistics("stationCd")

            requestNo = rsget_Logistics("requestNo")
            if IsNull(requestNo) then
                requestNo = "STOCKINVEST(" & masteridx & ")"
            end if
	    end if
	    rsget_Logistics.close

        if (requestNo = "") then
            responseJson = fnCreateCustomJsonResult("500", "내역이 존재하지 않습니다.[" & masteridx & "]", "500")
            Response.ContentType = "application/json; charset=utf-8"
            Call Response.AddHeader("Access-Control-Allow-Origin", "*")
            response.write responseJson

            dbget.Close
            dbget_Logistics.Close
            response.end
        end if

        if Not IsNull(status) then
            if Clng(status) >= 50 then
                responseJson = fnCreateCustomJsonResult("500", "이미 전송완료된 내역입니다.[" & masteridx & "]", "500")
                Response.ContentType = "application/json; charset=utf-8"
                Call Response.AddHeader("Access-Control-Allow-Origin", "*")
                response.write responseJson

                dbget.Close
                dbget_Logistics.Close
                response.end
            end if
        end if

		sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_stock_invest_master] "
        sqlStr = sqlStr + " set requestNo = '" & requestNo & "' "
        sqlStr = sqlStr + " where idx = '" & masteridx & "' and IsNull(status,0) < 50 "
        dbget_Logistics.Execute sqlStr

        '// 상품정보 전송
        sqlStr = " update db_aLogistics.dbo.tbl_agv_scheduledItems "
        sqlStr = sqlStr + " set isusing = 'N' "
        sqlStr = sqlStr + " where requestMaster = '" & requestNo & "' and isusing = 'Y' "
        dbget_Logistics.Execute sqlStr

        sqlStr = " if not exists(select idx from db_aLogistics.dbo.tbl_agv_scheduledItems where requestMaster = '" & requestNo & "' and isusing = 'Y') "
        sqlStr = sqlStr + " begin "
        sqlStr = sqlStr + "     insert into db_aLogistics.dbo.tbl_agv_scheduledItems(itemgubun, itemid, itemoption, realStock, displayOrderTypeCd, requestMaster) "
        sqlStr = sqlStr + "     select b.itemGubun, b.itemid, b.itemoption, 0, '상품정보전송', '" & requestNo & "' "
        sqlStr = sqlStr + "     from "
        sqlStr = sqlStr + "     	[db_aLogistics].[dbo].[tbl_agv_stock_invest_detail] b "
        sqlStr = sqlStr + "     where b.masteridx = '" & masteridx & "' and deldt is NULL "
        sqlStr = sqlStr + " end "
        dbget_Logistics.Execute sqlStr, affectedRows

        if (affectedRows > 0) then
            Call fnSendBrandItemInfo2AGV(requestNo)

            Call fnSendStockInvestByMasterIDX(masteridx, jsonResult)

            Call fnParseResultStockInvestJson(jsonResult, resultCode, resultMessage, failCode, inventorySurveyOrderId)

            if (resultCode = "00") then
		        sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_stock_invest_master] "
                sqlStr = sqlStr + " set inventorySurveyOrderId = '" & inventorySurveyOrderId & "', status = 50 "
                sqlStr = sqlStr + " where idx = '" & masteridx & "' and IsNull(status,0) < 50 "
                dbget_Logistics.Execute sqlStr

                responseJson = fnCreateCustomJsonResult("200", "OK", "200")
            else
                responseJson = fnCreateCustomJsonResult(resultCode, resultMessage, failCode)
            end if
        else
            responseJson = fnCreateCustomJsonResult("500", "재고조사지시할 상품이 없습니다.", "500")
        end if

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "agvstockinvestdel":
        '// 재고조사 지시삭제
        Call fnSendStockInvestCancelByMasterIDX(masteridx, jsonResult)

        Call fnParseResultJson(jsonResult, resultCode, resultMessage, failCode)

        if (resultCode = "00") then
		    sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_stock_invest_master] "
            sqlStr = sqlStr + " set status = 10 "
            sqlStr = sqlStr + " where idx = '" & masteridx & "' "
            dbget_Logistics.Execute sqlStr

            responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        else
            responseJson = fnCreateCustomJsonResult(resultCode, resultMessage, failCode)
        end if

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "senditeminfo":
        '// 상품정보 전송
        if (ordertype = "ipgo") then
            '// 입고처리시
            requestMaster = "IPGO(" & baljucode & ")"
            'baljucode

            sqlStr = " update db_aLogistics.dbo.tbl_agv_scheduledItems "
            sqlStr = sqlStr + " set isusing = 'N' "
            sqlStr = sqlStr + " where requestMaster = '" & requestMaster & "' and isusing = 'Y' "
            dbget_Logistics.Execute sqlStr

            sqlStr = " if not exists(select idx from db_aLogistics.dbo.tbl_agv_scheduledItems where requestMaster = '" & requestMaster & "' and isusing = 'Y') "
            sqlStr = sqlStr + " begin "
            sqlStr = sqlStr + "     insert into db_aLogistics.dbo.tbl_agv_scheduledItems(itemgubun, itemid, itemoption, realStock, displayOrderTypeCd, requestMaster) "
            sqlStr = sqlStr + "     select d.iitemgubun as itemgubun, d.itemid, d.itemoption, 0, '상품정보전송', '" & requestMaster & "' "
            sqlStr = sqlStr + "     from "
            sqlStr = sqlStr + "     	[TENDB].[db_storage].[dbo].tbl_acount_storage_detail d "
            sqlStr = sqlStr + "     where d.mastercode = '" & baljucode & "' "
            sqlStr = sqlStr + " end "
            dbget_Logistics.Execute sqlStr, affectedRows

            if (affectedRows > 0) then
                Call fnResetLastSendDate(requestMaster)
                Call fnSendBrandItemInfo2AGV(requestMaster)
            end if
        elseif (ordertype = "items") then
            requestMaster = "ITEMS(" & getDateFormatedWithDash(Now()) & ")"

            itemgubunArr = Split(itemgubun, ",")
            itemidArr = Split(itemid, ",")
            itemoptionArr = Split(itemoption, ",")

            affectedRows = 0
            for i = 0 to UBound(itemgubunArr)
                if Trim(itemgubunArr(i)) <> "" then
                    sqlStr = " if not exists(select idx from db_aLogistics.dbo.tbl_agv_scheduledItems where requestMaster = '" & requestMaster & "' and isusing = 'Y' and itemgubun = '" & Trim(itemgubunArr(i)) & "' and itemid = '" & Trim(itemidArr(i)) & "' and itemoption = '" & Trim(itemoptionArr(i)) & "') "
                    sqlStr = sqlStr + " begin "
                    sqlStr = sqlStr + "     insert into db_aLogistics.dbo.tbl_agv_scheduledItems(itemgubun, itemid, itemoption, realStock, displayOrderTypeCd, requestMaster) "
                    sqlStr = sqlStr + "     select '" & Trim(itemgubunArr(i)) & "', '" & Trim(itemidArr(i)) & "', '" & Trim(itemoptionArr(i)) & "', 0, '상품정보전송', '" & requestMaster & "' "
                    sqlStr = sqlStr + " end "
                    dbget_Logistics.Execute sqlStr

                    affectedRows = affectedRows + 1
                end if
            next

            if (affectedRows > 0) then
                Call fnResetLastSendDate(requestMaster)
                Call fnSendBrandItemInfo2AGV(requestMaster)
            end if
        end if

        responseJson = fnCreateCustomJsonResult("200", "OK", "200")

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "agvSendOrder":
        '// 주문전송
        affectedRows = 0
        if (ordertype = "onLine") then
            '// 상품정보 전송
            requestNo = "ORDERONLINE(" & baljuKey & ")"
            requestMaster = requestNo
            pickingOrderTypeCd = "온라인발주"

            sqlStr = " if not exists(select idx from db_aLogistics.dbo.tbl_agv_scheduledItems where requestMaster = '" & requestMaster & "' and isusing = 'Y') "
            sqlStr = sqlStr + " begin "
            sqlStr = sqlStr + "     insert into db_aLogistics.dbo.tbl_agv_scheduledItems(itemgubun, itemid, itemoption, realStock, displayOrderTypeCd, requestMaster) "
            sqlStr = sqlStr + "     select b.itemGubun, b.itemid, b.itemoption, 0, '상품정보전송', '" & requestMaster & "' "
            sqlStr = sqlStr + "     from "
            sqlStr = sqlStr + "     	[db_aLogistics].[dbo].[tbl_Logistics_baljuipgo] b "
            sqlStr = sqlStr + "     where b.baljuKey = " & baljuKey
            sqlStr = sqlStr + " end "
            dbget_Logistics.Execute sqlStr, affectedRows

            '// 메인디비 발주키
            sqlStr = " select top 1 siteBaljuid "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_baljumaster] "
            sqlStr = sqlStr + " where baljuKey = " & baljuKey
            rsget_Logistics.CursorLocation = adUseClient
            rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

            if  not rsget_Logistics.EOF  then
                siteBaljuid = rsget_Logistics("siteBaljuid")
            end if
            rsget_Logistics.close

            if (affectedRows > 0) then
                '// 사은품정보
                sqlStr = " select k.prd_itemgubun as itemgubun, k.prd_itemid as itemid, k.prd_itemoption as itemoption "
                sqlStr = sqlStr + " from "
                sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_baljumaster] bm "
                sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_baljudetail] bd on bm.id = bd.baljuid "
                sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_gift o on bd.orderserial = o.orderserial "
                sqlStr = sqlStr + " 	Join [db_event].[dbo].tbl_gift g on o.gift_code=g.gift_code "
                sqlStr = sqlStr + " 	left Join db_event.dbo.tbl_giftkind k on o.giftkind_code=k.giftkind_code "
                sqlStr = sqlStr + " 	left Join db_event.dbo.tbl_event e on g.evt_code=e.evt_code "
                sqlStr = sqlStr + " 	join [db_shop].[dbo].[tbl_shop_item] i "
                sqlStr = sqlStr + " 	on "
                sqlStr = sqlStr + " 		1 = 1 "
                sqlStr = sqlStr + " 		and k.prd_itemgubun = i.itemgubun "
                sqlStr = sqlStr + " 		and k.prd_itemid = i.shopitemid "
                sqlStr = sqlStr + " 		and k.prd_itemoption = i.itemoption "
                sqlStr = sqlStr + " where "
                sqlStr = sqlStr + " 	1 = 1 "
                sqlStr = sqlStr + " 	and bm.id = " & siteBaljuid
                sqlStr = sqlStr + " 	and o.gift_delivery = 'N' "
                sqlStr = sqlStr + " group by "
                sqlStr = sqlStr + " 	k.prd_itemgubun, k.prd_itemid, k.prd_itemoption "

                rsget.CursorLocation = adUseClient
                rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

                valuesArr = ""
                i = 0
                if  not rsget.EOF  then
                    do until rsget.eof
                        if (valuesArr = "") then
                            valuesArr = "values('" & rsget("itemgubun") & "', '" & rsget("itemid") & "', '" & rsget("itemoption") & "', 0, '사은품정보전송', '" & requestMaster & "')"
                        else
                            valuesArr = valuesArr + ", ('" & rsget("itemgubun") & "', '" & rsget("itemid") & "', '" & rsget("itemoption") & "', 0, '사은품정보전송', '" & requestMaster & "')"
                        end if
			            rsget.moveNext
		            loop
                end if
                rsget.close

                if (valuesArr <> "") then
                    sqlStr = " insert into db_aLogistics.dbo.tbl_agv_scheduledItems(itemgubun, itemid, itemoption, realStock, displayOrderTypeCd, requestMaster) "
                    sqlStr = sqlStr + valuesArr
                    dbget_Logistics.Execute sqlStr
                end if
            end if
        end if

        if (affectedRows > 0) then
            Call fnSendBrandItemInfo2AGV(requestMaster)
        end if

        '// orderserial, orderserialArr
        orderserialArr = ""
        if (ordertype = "onLine") then
            Call fnGetOerInfoOnlineJSONBatch(baljuKey, siteBaljuid, jsonResult)

            orderJson = jsonResult

            '''// 1. 주문번호 목록
            ''sqlStr = " select "
            ''sqlStr = sqlStr + " 	bd.orderserial "
            ''sqlStr = sqlStr + " from "
            ''sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_Logistics_baljumaster] bm "
            ''sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_baljudetail] bd on bm.baljuKey = bd.baljuKey "
            ''sqlStr = sqlStr + " where "
            ''sqlStr = sqlStr + " 	1 = 1 "
            ''sqlStr = sqlStr + " 	and bm.baljuKey = " & baljuKey

            ''rsget_Logistics.CursorLocation = adUseClient
            ''rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

            ''if  not rsget_Logistics.EOF  then
            ''    do until rsget_Logistics.eof
            ''        orderserialArr = orderserialArr & "|" & rsget_Logistics("orderserial")
			''        rsget_Logistics.moveNext
		    ''    loop
            ''end if
            ''rsget_Logistics.close

            ''orderJson = ""
            ''orderserialArr = Split(orderserialArr, "|")
            ''for i = 0 to UBound(orderserialArr)
            ''    if Trim(orderserialArr(i)) <> "" then
            ''        orderserial = Trim(orderserialArr(i))
            ''
            ''        Call fnGetOerInfoOnlineJSON(baljuKey, orderserial, jsonResult)

            ''        if (orderJson = "") then
            ''            orderJson = jsonResult
            ''        else
            ''            orderJson = orderJson & "," & jsonResult
            ''        end if
            ''    end if
            ''next
        end if

        jsonString = ""
        jsonString = jsonString + "{"
        jsonString = jsonString + "  ""salesList"": ["
        jsonString = jsonString + orderJson
        jsonString = jsonString + "  ]"
        jsonString = jsonString + "}"

        Call fnSendOrderInfoOnline(jsonString, jsonResult)

        ''responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        responseJson = jsonResult
        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "agvSendOrderCancel":
        '// 주문삭제
        responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "agvSendBalju":
        '// 발주 전송

        affectedRows = 0
        if (ordertype = "onLine") then
            '// 상품정보 전송
            if (baljuStr = "") then
                requestNo = "ORDERONLINE(" & baljuKey & ")"
            else
                baljuStr = Replace(baljuStr, "SINGLE", "단품")
                baljuStr = Replace(baljuStr, "COMPOUND", "복합")
                requestNo = baljuStr
            end if
            requestMaster = requestNo

            sqlStr = " update [db_aLogistics].[dbo].[tbl_Logistics_baljumaster] "
            sqlStr = sqlStr + " set batchNo = '" & requestNo & "' "
            sqlStr = sqlStr + " where baljuKey = '" & baljuKey & "' and IsNull(agvSendState,0) < 50 "
            dbget_Logistics.Execute sqlStr, affectedRows

            if (affectedRows < 1) then
                responseJson = fnCreateCustomJsonResult("500", "이미 전송된 발주이거나 잘못된 발주코드입니다.", "500")

                Response.ContentType = "application/json; charset=utf-8"
                Call Response.AddHeader("Access-Control-Allow-Origin", "*")
                response.write responseJson

                dbget.Close
                dbget_Logistics.Close
                response.end
            end if

            sqlStr = " update db_aLogistics.dbo.tbl_agv_scheduledItems "
            sqlStr = sqlStr + " set isusing = 'N' "
            sqlStr = sqlStr + " where requestMaster = '" & requestMaster & "' and isusing = 'Y' "
            dbget_Logistics.Execute sqlStr

            sqlStr = " if not exists(select idx from db_aLogistics.dbo.tbl_agv_scheduledItems where requestMaster = '" & requestMaster & "' and isusing = 'Y') "
            sqlStr = sqlStr + " begin "
            sqlStr = sqlStr + "     insert into db_aLogistics.dbo.tbl_agv_scheduledItems(itemgubun, itemid, itemoption, realStock, displayOrderTypeCd, requestMaster) "
            sqlStr = sqlStr + "     select b.itemGubun, b.itemid, b.itemoption, 0, '상품정보전송', '" & requestMaster & "' "
            sqlStr = sqlStr + "     from "
            sqlStr = sqlStr + "     	[db_aLogistics].[dbo].[tbl_Logistics_baljuipgo] b "
            sqlStr = sqlStr + "     where b.baljuKey = " & baljuKey
            sqlStr = sqlStr + " end "
            dbget_Logistics.Execute sqlStr, affectedRows

            '// 메인디비 발주키
            sqlStr = " select top 1 siteBaljuid "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_baljumaster] "
            sqlStr = sqlStr + " where baljuKey = " & baljuKey
            rsget_Logistics.CursorLocation = adUseClient
            rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

            if  not rsget_Logistics.EOF  then
                siteBaljuid = rsget_Logistics("siteBaljuid")
            end if
            rsget_Logistics.close

            if (affectedRows > 0) then
                '// 사은품정보
                sqlStr = " select k.prd_itemgubun as itemgubun, k.prd_itemid as itemid, k.prd_itemoption as itemoption "
                sqlStr = sqlStr + " from "
                sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_baljumaster] bm "
                sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_baljudetail] bd on bm.id = bd.baljuid "
                sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_gift o on bd.orderserial = o.orderserial "
                sqlStr = sqlStr + " 	Join [db_event].[dbo].tbl_gift g on o.gift_code=g.gift_code "
                sqlStr = sqlStr + " 	left Join db_event.dbo.tbl_giftkind k on o.giftkind_code=k.giftkind_code "
                sqlStr = sqlStr + " 	left Join db_event.dbo.tbl_event e on g.evt_code=e.evt_code "
                sqlStr = sqlStr + " 	join [db_shop].[dbo].[tbl_shop_item] i "
                sqlStr = sqlStr + " 	on "
                sqlStr = sqlStr + " 		1 = 1 "
                sqlStr = sqlStr + " 		and k.prd_itemgubun = i.itemgubun "
                sqlStr = sqlStr + " 		and k.prd_itemid = i.shopitemid "
                sqlStr = sqlStr + " 		and k.prd_itemoption = i.itemoption "
                sqlStr = sqlStr + " where "
                sqlStr = sqlStr + " 	1 = 1 "
                sqlStr = sqlStr + " 	and bm.id = " & siteBaljuid
                sqlStr = sqlStr + " 	and o.gift_delivery = 'N' "
                sqlStr = sqlStr + " group by "
                sqlStr = sqlStr + " 	k.prd_itemgubun, k.prd_itemid, k.prd_itemoption "

                rsget.CursorLocation = adUseClient
                rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

                valuesArr = ""
                i = 0
                if  not rsget.EOF  then
                    do until rsget.eof
                        if (valuesArr = "") then
                            valuesArr = "values('" & rsget("itemgubun") & "', '" & rsget("itemid") & "', '" & rsget("itemoption") & "', 0, '사은품정보전송', '" & requestMaster & "')"
                        else
                            valuesArr = valuesArr + ", ('" & rsget("itemgubun") & "', '" & rsget("itemid") & "', '" & rsget("itemoption") & "', 0, '사은품정보전송', '" & requestMaster & "')"
                        end if
			            rsget.moveNext
		            loop
                end if
                rsget.close

                if (valuesArr <> "") then
                    sqlStr = " insert into db_aLogistics.dbo.tbl_agv_scheduledItems(itemgubun, itemid, itemoption, realStock, displayOrderTypeCd, requestMaster) "
                    sqlStr = sqlStr + valuesArr
                    dbget_Logistics.Execute sqlStr
                end if
            end if

        elseif (ordertype = "offLine") then
            '// 상품정보 전송
            requestNo = "ORDEROFFLINE(" & baljuKey & ")"
            requestMaster = requestNo

            sqlStr = " update db_aLogistics.dbo.tbl_agv_scheduledItems "
            sqlStr = sqlStr + " set isusing = 'N' "
            sqlStr = sqlStr + " where requestMaster = '" & requestMaster & "' and isusing = 'Y' "
            dbget_Logistics.Execute sqlStr

            sqlStr = " if not exists(select idx from db_aLogistics.dbo.tbl_agv_scheduledItems where requestMaster = '" & requestMaster & "' and isusing = 'Y') "
            sqlStr = sqlStr + " begin "
            sqlStr = sqlStr + "     insert into db_aLogistics.dbo.tbl_agv_scheduledItems(itemgubun, itemid, itemoption, realStock, displayOrderTypeCd, requestMaster) "
            sqlStr = sqlStr + "     select b.itemGubun, b.itemid, b.itemoption, 0, '상품정보전송', '" & requestMaster & "' "
            sqlStr = sqlStr + "     from "
            sqlStr = sqlStr + "     	[db_aLogistics].[dbo].[tbl_Logistics_offline_baljuipgo] b "
            sqlStr = sqlStr + "     where b.baljuKey = " & baljuKey
            sqlStr = sqlStr + " end "
            dbget_Logistics.Execute sqlStr, affectedRows
        end if

        if (affectedRows > 0) then
            Call fnSendBrandItemInfo2AGV(requestMaster)
        end if

        if (ordertype = "onLine") then
            Call fnGetOerInfoOnlineJSONBatchBalju(baljuKey, siteBaljuid, jsonResult)

            orderJson = jsonResult

            sqlStr = " select top 1 "
            sqlStr = sqlStr + " 	bm.baljuKey as batchNo "
            sqlStr = sqlStr + " 	, 'ONLINE' as batchTypeCd "
            sqlStr = sqlStr + " 	, (case when IsNull(boxType, 'ETC') in ('ABC', 'ABCD', 'A1', 'B1', 'C1', 'D1', 'AB') then 'SMALL' else 'ETC' end) as boxGroupCd "
            sqlStr = sqlStr + " 	, (case "
            sqlStr = sqlStr + " 			when songjangdiv = '4' then 'CJ' "
            sqlStr = sqlStr + " 			when songjangdiv = '90' then 'EMS' "
            sqlStr = sqlStr + " 			else songjangdiv end) as shippingGroupCd "
            sqlStr = sqlStr + " 	, (case "
            sqlStr = sqlStr + " 	        when extSiteName = '10x10' then '10x10' "
            sqlStr = sqlStr + " 	        when extSiteName in ('ithinkso', 'itsSite') then 'ithinkso' "
            sqlStr = sqlStr + " 			else 'error' end) as shopGroupCd "
            sqlStr = sqlStr + " 	, (case when baljutype = 'S' then 'SINGLE' else 'COMPOUND' end) as skuGroupCd "
            sqlStr = sqlStr + " 	, '" & stationCd & "' as stationCd "
            sqlStr = sqlStr + " 	, * "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_baljumaster] bm "
            sqlStr = sqlStr + " where siteseq = '10' and baljuKey = " & baljuKey

            rsget_Logistics.CursorLocation = adUseClient
            rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

            if  not rsget_Logistics.EOF  then
                siteBaljuid = rsget_Logistics("siteBaljuid")

                baljuJson = ""
                baljuJson = baljuJson + "{"
                baljuJson = baljuJson + "  ""batchNo"": """ & requestNo & ""","
                baljuJson = baljuJson + "  ""orderTypeCd"": ""online"","
                baljuJson = baljuJson + "  ""distributionYn"": ""Y"","
                baljuJson = baljuJson + "  ""batchTypeCd"": """ & rsget_Logistics("batchTypeCd") & ""","
                baljuJson = baljuJson + "  ""boxGroupCd"": """ & rsget_Logistics("boxGroupCd") & ""","
                baljuJson = baljuJson + "  ""caseList"": ["
                baljuJson = baljuJson + " <<caseList>> "
                baljuJson = baljuJson + "  ],"
                baljuJson = baljuJson + "  ""shippingGroupCd"": """ & rsget_Logistics("shippingGroupCd") & ""","
                baljuJson = baljuJson + "  ""shopGroupCd"": """ & rsget_Logistics("shopGroupCd") & ""","
                baljuJson = baljuJson + "  ""skuGroupCd"": """ & rsget_Logistics("skuGroupCd") & ""","
                baljuJson = baljuJson + "  ""stationCd"": """ & rsget_Logistics("stationCd") & """"
                baljuJson = baljuJson + "}"

                baljuJson = Replace(baljuJson, "<<caseList>>", orderJson)
            end if
            rsget_Logistics.close
        elseif (ordertype = "offLine") then
            Call fnGetOerInfoOfflineJSONBatchBalju(baljuKey, jsonResult)

            orderJson = jsonResult

            sqlStr = " select top 1 "
            sqlStr = sqlStr + " 	bm.baljuKey as batchNo "
            sqlStr = sqlStr + " 	, 'OFFLINE' as batchTypeCd "
            sqlStr = sqlStr + " 	, 'ETC' as boxGroupCd "
            sqlStr = sqlStr + " 	, (case "
            sqlStr = sqlStr + " 			when songjangdiv = '4' then 'CJ' "
            sqlStr = sqlStr + " 			when songjangdiv = '90' then 'EMS' "
            sqlStr = sqlStr + " 			when songjangdiv = '91' then 'DHL' "
            sqlStr = sqlStr + " 			when songjangdiv = '98' then 'QUICK' "
            sqlStr = sqlStr + " 			when songjangdiv = '99' then 'ETC' "
            sqlStr = sqlStr + " 			else songjangdiv end) as shippingGroupCd "
            sqlStr = sqlStr + " 	, '10x10' as shopGroupCd "
            sqlStr = sqlStr + " 	, 'COMPOUND' as skuGroupCd "
            sqlStr = sqlStr + " 	, '" & stationCd & "' as stationCd "
            sqlStr = sqlStr + " 	, * "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_offline_baljumaster] bm "
            sqlStr = sqlStr + " where siteseq = '10' and baljuKey = " & baljuKey

            rsget_Logistics.CursorLocation = adUseClient
            rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

            if  not rsget_Logistics.EOF  then
                siteBaljuid = rsget_Logistics("siteBaljuid")

                baljuJson = ""
                baljuJson = baljuJson + "{"
                baljuJson = baljuJson + "  ""batchNo"": """ & requestNo & ""","
                baljuJson = baljuJson + "  ""orderTypeCd"": ""offline"","
                baljuJson = baljuJson + "  ""distributionYn"": ""N"","
                baljuJson = baljuJson + "  ""batchTypeCd"": """ & rsget_Logistics("batchTypeCd") & ""","
                baljuJson = baljuJson + "  ""boxGroupCd"": """ & rsget_Logistics("boxGroupCd") & ""","
                baljuJson = baljuJson + "  ""caseList"": ["
                baljuJson = baljuJson + " <<caseList>> "
                baljuJson = baljuJson + "  ],"
                baljuJson = baljuJson + "  ""shippingGroupCd"": """ & rsget_Logistics("shippingGroupCd") & ""","
                baljuJson = baljuJson + "  ""shopGroupCd"": """ & rsget_Logistics("shopGroupCd") & ""","
                baljuJson = baljuJson + "  ""skuGroupCd"": """ & rsget_Logistics("skuGroupCd") & ""","
                baljuJson = baljuJson + "  ""stationCd"": """ & rsget_Logistics("stationCd") & """"
                baljuJson = baljuJson + "}"

                baljuJson = Replace(baljuJson, "<<caseList>>", orderJson)
            end if
            rsget_Logistics.close
        end if

        jsonString = baljuJson

        Call fnWriteLog(mode, jsonString)

        if (Trim(orderJson) = "") then
            responseJson = fnCreateCustomJsonResult("500", "전송할 주문내역이 없습니다.(모두 취소되었습니다.)", "500")
        else
            Call fnSendOrderInfoOnlineBalju(jsonString, jsonResult)

            Call fnWriteLog(mode, jsonResult)

            Call fnParseResultSendBaljuJson(jsonResult, resultCode, resultMessage, failCode)

            if (resultCode = "00") then
                if (ordertype = "onLine") then
                    sqlStr = " update "
                    sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_baljumaster] "
                    sqlStr = sqlStr + " set agvSendState = '50' "
                    sqlStr = sqlStr + " where baljuKey = " & baljuKey
                    dbget_Logistics.Execute sqlStr
                elseif (ordertype = "offLine") then
                    sqlStr = " update "
                    sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_offline_baljumaster] "
                    sqlStr = sqlStr + " set agvSendState = '50' "
                    sqlStr = sqlStr + " where baljuKey = " & baljuKey
                    dbget_Logistics.Execute sqlStr
                end if

                responseJson = fnCreateCustomJsonResult("200", "OK", "200")
            else
                responseJson = fnCreateCustomJsonResult("500", resultMessage, "500")
            end if
        end if

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "agvSendBaljuMulti":
        '// 발주전송[멀티스테이션]
        affectedRows = 0
        sqlStr = " exec [db_aLogistics].[dbo].[usp_LogisticsItem_Balju_to_multi] " & baljuKey
        dbget_Logistics.Execute sqlStr

        sqlStr = " select baljuKey as newBaljuKey "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_agv_baljumaster] "
        sqlStr = sqlStr + " where orgBaljuKey = " & baljuKey & " and IsNull(agvSendState,0) < 50 "
        rsget_Logistics.CursorLocation = adUseClient
        rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

        newBaljuKeyArr = ""
        if  not rsget_Logistics.EOF  then
            do until rsget_Logistics.eof
                newBaljuKeyArr = newBaljuKeyArr & "," & rsget_Logistics("newBaljuKey")
			    rsget_Logistics.moveNext
		    loop
        end if
        rsget_Logistics.close

        if (newBaljuKeyArr = "") then
            responseJson = fnCreateCustomJsonResult("500", "이미 전송된 발주이거나 잘못된 발주코드입니다.", "500")

            Response.ContentType = "application/json; charset=utf-8"
            Call Response.AddHeader("Access-Control-Allow-Origin", "*")
            response.write responseJson

            dbget.Close
            dbget_Logistics.Close
            response.end
        end if

        successCnt = 0
        failCnt = 0

        newBaljuKeyArr = Split(newBaljuKeyArr, ",")
        for i = 0 to UBound(newBaljuKeyArr)
            newBaljuKey = Trim(newBaljuKeyArr(i))
            if newBaljuKey <> "" then

                sqlStr = " select top 1 batchNo "
                sqlStr = sqlStr + " from "
                sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_agv_baljumaster] "
                sqlStr = sqlStr + " where baljuKey = " & newBaljuKey
                rsget_Logistics.CursorLocation = adUseClient
                rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

                if  not rsget_Logistics.EOF  then
                    requestMaster = rsget_Logistics("batchNo")
                end if
                rsget_Logistics.close

                sqlStr = " update db_aLogistics.dbo.tbl_agv_scheduledItems "
                sqlStr = sqlStr + " set isusing = 'N' "
                sqlStr = sqlStr + " where requestMaster = '" & requestMaster & "' and isusing = 'Y' "
                dbget_Logistics.Execute sqlStr

                sqlStr = " if not exists(select idx from db_aLogistics.dbo.tbl_agv_scheduledItems where requestMaster = '" & requestMaster & "' and isusing = 'Y') "
                sqlStr = sqlStr + " begin "
                sqlStr = sqlStr + "     insert into db_aLogistics.dbo.tbl_agv_scheduledItems(itemgubun, itemid, itemoption, realStock, displayOrderTypeCd, requestMaster) "
                sqlStr = sqlStr + "     select d.itemgubun, d.itemid, d.itemoption, 0, '상품정보전송', '" & requestMaster & "' "
                sqlStr = sqlStr + " 	from "
                sqlStr = sqlStr + "         [db_aLogistics].[dbo].[tbl_Logistics_agv_baljudetail] bd "
                sqlStr = sqlStr + "         join [db_aLogistics].[dbo].[tbl_Logistics_agv_order_detail] d on bd.orderserial = d.orderserial "
                sqlStr = sqlStr + "     where bd.baljuKey = " & newBaljuKey
                sqlStr = sqlStr + "     group by d.itemgubun, d.itemid, d.itemoption "
                sqlStr = sqlStr + " end "
                dbget_Logistics.Execute sqlStr, affectedRows

                if (affectedRows > 0) then
                    Call fnSendBrandItemInfo2AGV(requestMaster)
                end if

                Call fnGetOerInfoOnlineJSONBatchBaljuMulti(newBaljuKey, jsonResult)

                orderJson = jsonResult

                sqlStr = " select top 1 "
                sqlStr = sqlStr + " 	bm.baljuKey as batchNo "
                sqlStr = sqlStr + " 	, 'ONLINE' as batchTypeCd "
                ''sqlStr = sqlStr + " 	, (case when IsNull(boxType, 'ETC') in ('ABC', 'ABCD', 'A1', 'B1', 'C1', 'D1', 'AB') then 'SMALL' else 'ETC' end) as boxGroupCd "
                sqlStr = sqlStr + " 	, 'ETC' as boxGroupCd "				'// 박스타입 ETC로 강제지정, skyer9, 2020-09-25
                sqlStr = sqlStr + " 	, (case "
                sqlStr = sqlStr + " 			when songjangdiv = '4' then 'CJ' "
                sqlStr = sqlStr + " 			when songjangdiv = '90' then 'EMS' "
                sqlStr = sqlStr + " 			else songjangdiv end) as shippingGroupCd "
                sqlStr = sqlStr + " 	, (case "
                sqlStr = sqlStr + " 	        when extSiteName = '10x10' then '10x10' "
                sqlStr = sqlStr + " 	        when extSiteName in ('ithinkso', 'itsSite') then 'ithinkso' "
                sqlStr = sqlStr + " 			else 'error' end) as shopGroupCd "
                sqlStr = sqlStr + " 	, (case when baljutype = 'S' then 'SINGLE' else 'COMPOUND' end) as skuGroupCd "
                sqlStr = sqlStr + " 	, pickingStationCd as stationCd "
                sqlStr = sqlStr + " 	, * "
                sqlStr = sqlStr + " from "
                sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_agv_baljumaster] bm "
                sqlStr = sqlStr + " where siteseq = '10' and baljuKey = " & newBaljuKey

                rsget_Logistics.CursorLocation = adUseClient
                rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

                if  not rsget_Logistics.EOF  then
                    baljuJson = ""
                    baljuJson = baljuJson + "{"
                    baljuJson = baljuJson + "  ""batchNo"": """ & requestMaster & ""","
                    baljuJson = baljuJson + "  ""orderTypeCd"": ""online"","
                    baljuJson = baljuJson + "  ""distributionYn"": ""Y"","
                    baljuJson = baljuJson + "  ""batchTypeCd"": """ & rsget_Logistics("batchTypeCd") & ""","
                    baljuJson = baljuJson + "  ""boxGroupCd"": """ & rsget_Logistics("boxGroupCd") & ""","
                    baljuJson = baljuJson + "  ""caseList"": ["
                    baljuJson = baljuJson + " <<caseList>> "
                    baljuJson = baljuJson + "  ],"
                    baljuJson = baljuJson + "  ""shippingGroupCd"": """ & rsget_Logistics("shippingGroupCd") & ""","
                    baljuJson = baljuJson + "  ""shopGroupCd"": """ & rsget_Logistics("shopGroupCd") & ""","
                    baljuJson = baljuJson + "  ""skuGroupCd"": """ & rsget_Logistics("skuGroupCd") & ""","
                    baljuJson = baljuJson + "  ""stationCd"": """ & rsget_Logistics("stationCd") & """"
                    baljuJson = baljuJson + "}"

                    baljuJson = Replace(baljuJson, "<<caseList>>", orderJson)
                end if
                rsget_Logistics.close

                jsonString = baljuJson

                Call fnWriteLog(mode, jsonString)

                if (Trim(orderJson) = "") then
                    ''responseJson = fnCreateCustomJsonResult("500", "전송할 주문내역이 없습니다.(모두 취소되었습니다.)", "500")
                    failCnt = failCnt + 1
                else
                    Call fnSendOrderInfoOnlineBalju(jsonString, jsonResult)

                    Call fnWriteLog(mode, jsonResult)

                    Call fnParseResultSendBaljuJson(jsonResult, resultCode, resultMessage, failCode)

                    if (resultCode = "00") then
                        sqlStr = " update "
                        sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_agv_baljumaster] "
                        sqlStr = sqlStr + " set agvSendState = '50' "
                        sqlStr = sqlStr + " where baljuKey = " & newBaljuKey
                        dbget_Logistics.Execute sqlStr

                        ''responseJson = fnCreateCustomJsonResult("200", "OK", "200")
                        successCnt = successCnt + 1
                    else
                        ''responseJson = fnCreateCustomJsonResult("500", resultMessage, "500")
                        failCnt = failCnt + 1
                    end if
                end if
            end if
        next

        if failCnt = 0 then
            sqlStr = " update "
            sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_baljumaster] "
            sqlStr = sqlStr + " set agvSendState = '50' "
            sqlStr = sqlStr + " where baljuKey = " & baljuKey
            dbget_Logistics.Execute sqlStr

            responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        else
            resultMessage = "성공 " & successCnt & "건 / 실패 " & failCnt & "건"
            responseJson = fnCreateCustomJsonResult("500", resultMessage, "500")
        end if

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson

    case "agvSendBaljuCancel":
        if (ordertype = "onLine") then
            sqlStr = " select top 1 batchNo "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_baljumaster] "
            sqlStr = sqlStr + " where baljuKey = " & baljuKey
            rsget_Logistics.CursorLocation = adUseClient
            rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

            if  not rsget_Logistics.EOF  then
                baljuStr = rsget_Logistics("batchNo")
            end if
            rsget_Logistics.close

            if IsNull(baljuStr) then
                baljuStr = "ORDERONLINE(" & baljuKey & ")"
            end if
            requestNo = baljuStr
        elseif (ordertype = "offLine") then
            requestNo = "ORDEROFFLINE(" & baljuKey & ")"
        end if

        jsonString = ""
        jsonString = jsonString + "{"
        jsonString = jsonString + "  ""batchNo"": """ & requestNo & """"
        jsonString = jsonString + "}"

        Call fnSendOrderInfoBaljuCancel(jsonString, jsonResult)

        ''Call fnWriteLog(mode, jsonResult)

        Call fnParseResultSendBaljuJson(jsonResult, resultCode, resultMessage, failCode)

        if (resultCode = "00") then
            if (ordertype = "onLine") then
                sqlStr = " update "
                sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_baljumaster] "
                sqlStr = sqlStr + " set agvSendState = '10' "
                sqlStr = sqlStr + " where baljuKey = " & baljuKey
                dbget_Logistics.Execute sqlStr
            elseif (ordertype = "offLine") then
                sqlStr = " update "
                sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_offline_baljumaster] "
                sqlStr = sqlStr + " set agvSendState = '10' "
                sqlStr = sqlStr + " where baljuKey = " & baljuKey
                dbget_Logistics.Execute sqlStr
            end if

            responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        else
            responseJson = fnCreateCustomJsonResult("500", resultMessage, "500")
        end if

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "agvSendBaljuCancelMulti":
        '// 전송취소[멀티스테이션]

        sqlStr = " select baljuKey as newBaljuKey "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_agv_baljumaster] "
        sqlStr = sqlStr + " where orgBaljuKey = " & baljuKey
        rsget_Logistics.CursorLocation = adUseClient
        rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

        newBaljuKeyArr = ""
        if  not rsget_Logistics.EOF  then
            do until rsget_Logistics.eof
                newBaljuKeyArr = newBaljuKeyArr & "," & rsget_Logistics("newBaljuKey")
			    rsget_Logistics.moveNext
		    loop
        end if
        rsget_Logistics.close

        if (newBaljuKeyArr = "") then
            responseJson = fnCreateCustomJsonResult("500", "잘못된 발주코드입니다.", "500")

            Response.ContentType = "application/json; charset=utf-8"
            Call Response.AddHeader("Access-Control-Allow-Origin", "*")
            response.write responseJson

            dbget.Close
            dbget_Logistics.Close
            response.end
        end if

        successCnt = 0
        failCnt = 0

        newBaljuKeyArr = Split(newBaljuKeyArr, ",")
        for i = 0 to UBound(newBaljuKeyArr)
            newBaljuKey = Trim(newBaljuKeyArr(i))
            if newBaljuKey <> "" then

                sqlStr = " select top 1 batchNo "
                sqlStr = sqlStr + " from "
                sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_agv_baljumaster] "
                sqlStr = sqlStr + " where baljuKey = " & newBaljuKey
                rsget_Logistics.CursorLocation = adUseClient
                rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

                requestNo = ""
                if  not rsget_Logistics.EOF  then
                    requestNo = rsget_Logistics("batchNo")
                end if
                rsget_Logistics.close

                jsonString = ""
                jsonString = jsonString + "{"
                jsonString = jsonString + "  ""batchNo"": """ & requestNo & """"
                jsonString = jsonString + "}"

                Call fnSendOrderInfoBaljuCancel(jsonString, jsonResult)

                ''Call fnWriteLog(mode, jsonResult)

                Call fnParseResultSendBaljuJson(jsonResult, resultCode, resultMessage, failCode)

                if (resultCode = "00") then
                    sqlStr = " update "
                    sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_agv_baljumaster] "
                    sqlStr = sqlStr + " set agvSendState = '10' "
                    sqlStr = sqlStr + " where baljuKey = " & newBaljuKey
                    dbget_Logistics.Execute sqlStr

                    ''responseJson = fnCreateCustomJsonResult("200", "OK", "200")
                    successCnt = successCnt + 1
                else
                    ''responseJson = fnCreateCustomJsonResult("500", resultMessage, "500")
                    failCnt = failCnt + 1
                end if
            end if
        next

        if failCnt = 0 then
            sqlStr = " update "
            sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_baljumaster] "
            sqlStr = sqlStr + " set agvSendState = '10' "
            sqlStr = sqlStr + " where baljuKey = " & baljuKey
            dbget_Logistics.Execute sqlStr

            responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        else
            resultMessage = "성공 " & successCnt & "건 / 실패 " & failCnt & "건"
            responseJson = fnCreateCustomJsonResult("500", resultMessage, "500")
        end if

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson

    case "currstock":
        '// AGV 현재고조회
        Call fnGetProductStockInfoBySkuCd(skuCd, jsonResult)

        '// fnGetStockByResult() 참조

        Response.ContentType = "application/json; charset=utf-8"
        response.write jsonResult
    case "currstockListView":
        '// AGV 현재고조회(복수 SKU, 복수 브랜드), 저장안함
        brandArray = Split(brandArray, ",")
        skuCdArray = Split(skuCdArray, ",")
        Call fnGetProductStockInfoByBrandBySkuCd(brandArray, skuCdArray, jsonResult)

        '// fnGetStockByResult() 참조

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write jsonResult
    case "currstockList":
        '// AGV 현재고조회(복수 SKU, 복수 브랜드) 및 업데이트
        brandArray = Split(brandArray, ",")
        skuCdArray = Split(skuCdArray, ",")
        Call fnGetProductStockInfoByBrandBySkuCd(brandArray, skuCdArray, jsonResult)

        Call fnWriteLog(mode, jsonResult)

        Call fnParseResultProductStockInfoByBrandBySkuCd(jsonResult, resultCode, resultMessage, failCode)

        if (resultCode = "00") then
            Call fnSaveStockInfo(jsonResult, True)
            responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        else
            responseJson = fnCreateCustomJsonResult("500", resultMessage, "500")
        end if

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "chgwarehouse2bulk":
        '// 상품 재고구분 벌크전환

        ''Call fnChangeWarehouseToBulk(skuCdArray)

        ''나중에 필요하다고 하면 작업하자.

    case "currstockall":
        '// AGV 현재고조회

        brandArray = Array()
        skuCdArray = Array()
        Call fnGetProductStockInfoByBrandBySkuCd(brandArray, skuCdArray, jsonResult)

        Call fnWriteLog(mode, jsonResult)

        Call fnParseResultProductStockInfoByBrandBySkuCd(jsonResult, resultCode, resultMessage, failCode)

        if (resultCode = "00") then
            Call fnSaveStockInfo(jsonResult, False)
            responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        else
            responseJson = fnCreateCustomJsonResult("500", resultMessage, "500")
        end if

        ''Call fnWriteLog(mode, jsonResult)

        '// fnGetStockByResult() 참조

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "recvagvipgo":
        If Request.TotalBytes > 0 Then
            lngBytesCount = Request.TotalBytes
            jsonString = BytesToStr(Request.BinaryRead(lngBytesCount), "UTF-8")

            Call fnWriteLog(mode, jsonString)

            Set resultJson = New aspJson
            resultJson.loadJSON(jsonString)

            displayOrderNo = resultJson.data("displayOrderNo")
            progressStatusCd = resultJson.data("progressStatusCd")

            Set resultJson = Nothing

            ''progressStatusCd
            ''READY		준비
            ''COMPLETE	완료
            ''CANCEL	취소
            select case progressStatusCd
                case "CANCEL":
		            sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_scheduledItems] "
                    sqlStr = sqlStr + " set isusing = 'N', status = 10, lastupdate=getdate() "
                    sqlStr = sqlStr + " where displayOrderNo = '" & displayOrderNo & "' and status in (50, 70) "
                    dbget_Logistics.Execute sqlStr, affectedRows
                case "READY":
		            sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_scheduledItems] "
                    sqlStr = sqlStr + " set status = 70, lastupdate=getdate() "
                    sqlStr = sqlStr + " where displayOrderNo = '" & displayOrderNo & "' and status in (50, 70) "
                    dbget_Logistics.Execute sqlStr, affectedRows
                case "COMPLETE":
                    sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_scheduledItems] "
                    sqlStr = sqlStr + " set status = 100, lastupdate=getdate() "
                    sqlStr = sqlStr + " where displayOrderNo = '" & displayOrderNo & "' and status in (50, 70) "
                    dbget_Logistics.Execute sqlStr, affectedRows
                case else
                    affectedRows = -1
            end select

            if (affectedRows > 0) then
                responseJson = fnCreateCustomJsonResult("200", "OK", "200")
            elseif (affectedRows = -1) then
                responseJson = fnCreateCustomJsonResult("500", "UNKNOWN STATE", "500")
            else
                responseJson = fnCreateCustomJsonResult("500", "NO DATA FOUND", "500")
            end if
        else
            responseJson = fnCreateCustomJsonResult("500", "NO BODY", "500")
        End If

        Response.ContentType = "application/json; charset=utf-8"
        response.write responseJson
    case "recvagvpick":
        If Request.TotalBytes > 0 Then
            lngBytesCount = Request.TotalBytes
            jsonString = BytesToStr(Request.BinaryRead(lngBytesCount), "UTF-8")

            Call fnWriteLog(mode, jsonString)

            Set resultJson = New aspJson
            resultJson.loadJSON(jsonString)

            pickingOrderNo = resultJson.data("pickingOrderNo")
            progressStatusCd = resultJson.data("progressStatusCd")

            Set resultJson = Nothing


            ''progressStatusCd
            ''READY		준비
            ''ING		진행
            ''COMPLETE	완료
            ''CANCEL	취소
            select case progressStatusCd
                case "CANCEL":
		            sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickupItems] "
                    sqlStr = sqlStr + " set isusing = 'N', status = 10, lastupdate=getdate() "
                    sqlStr = sqlStr + " where pickingOrderNo = '" & pickingOrderNo & "' and status in (50, 70, 80) "
                    dbget_Logistics.Execute sqlStr, affectedRows

                    if (affectedRows = 0) then
		                sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
                        sqlStr = sqlStr + " set status = 10,updt=getdate() "
                        sqlStr = sqlStr + " where pickingOrderNo = '" & pickingOrderNo & "' "
                        dbget_Logistics.Execute sqlStr, affectedRows
                    end if
                case "READY":
		            sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickupItems] "
                    sqlStr = sqlStr + " set status = 70, lastupdate=getdate() "
                    sqlStr = sqlStr + " where pickingOrderNo = '" & pickingOrderNo & "' and status in (50, 70, 80) "
                    dbget_Logistics.Execute sqlStr, affectedRows

                    if (affectedRows = 0) then
		                sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
                        sqlStr = sqlStr + " set status = 70,updt=getdate() "
                        sqlStr = sqlStr + " where pickingOrderNo = '" & pickingOrderNo & "' "
                        dbget_Logistics.Execute sqlStr, affectedRows
                    end if
                case "ING":
		            sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickupItems] "
                    sqlStr = sqlStr + " set status = 80, lastupdate=getdate() "
                    sqlStr = sqlStr + " where pickingOrderNo = '" & pickingOrderNo & "' and status in (50, 70, 80) "
                    dbget_Logistics.Execute sqlStr, affectedRows

                    if (affectedRows = 0) then
		                sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
                        sqlStr = sqlStr + " set status = 80,updt=getdate() "
                        sqlStr = sqlStr + " where pickingOrderNo = '" & pickingOrderNo & "' "
                        dbget_Logistics.Execute sqlStr, affectedRows
                    end if
                case "COMPLETE":
                    sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickupItems] "
                    sqlStr = sqlStr + " set status = 100 "
                    sqlStr = sqlStr + " where pickingOrderNo = '" & pickingOrderNo & "' and status in (50, 70, 80) "
                    dbget_Logistics.Execute sqlStr, affectedRows

                    if (affectedRows = 0) then
		                sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
                        sqlStr = sqlStr + " set status = 100,updt=getdate() "
                        sqlStr = sqlStr + " where pickingOrderNo = '" & pickingOrderNo & "' "
                        dbget_Logistics.Execute sqlStr, affectedRows
                    end if
                case else
                    affectedRows = -1
            end select

            if (affectedRows > 0) then
                responseJson = fnCreateCustomJsonResult("200", "OK", "200")
            elseif (affectedRows = -1) then
                responseJson = fnCreateCustomJsonResult("500", "UNKNOWN STATE", "500")
            else
                responseJson = fnCreateCustomJsonResult("500", "NO DATA FOUND", "500")
            end if
        else
            responseJson = fnCreateCustomJsonResult("500", "NO BODY", "500")
        End If

        Response.ContentType = "application/json; charset=utf-8"
        response.write responseJson
    case "recvagvpickfinish":
        If Request.TotalBytes > 0 Then
            lngBytesCount = Request.TotalBytes
            jsonString = BytesToStr(Request.BinaryRead(lngBytesCount), "UTF-8")

            Call fnWriteLog(mode, jsonString)

            Set resultJson = New aspJson
            resultJson.loadJSON(jsonString)

            pickingOrderTypeCd = resultJson.data("pickingOrderTypeCd")                   '' INTERFACE(피킹지시 인터페이스O), ORDER(발주지시 인터페이스O), MANUAL(수기 인터페이스X)
            progressStatusCd = resultJson.data("progressStatusCd")                       '' COMPLETE, FORCE_COMPLETE
            requestNo = resultJson.data("requestNo")
            Set resultData = resultJson.data("skuList")

            select case pickingOrderTypeCd
                case "INTERFACE"
                    '// 피킹지시 인터페이스O
                    if (requestNo = "") then
                        responseJson = fnCreateCustomJsonResult("500", "empty requestNo", "500")
                        Response.ContentType = "application/json; charset=utf-8"
                        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
                        response.write responseJson

                        dbget.Close : dbget_Logistics.Close : response.end
                    end if

                    sqlStr = " select top 1 * "
                    sqlStr = sqlStr + " from "
                    sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
                    sqlStr = sqlStr + " where requestNo = '" & requestNo & "' "

                    rsget_Logistics.CursorLocation = adUseClient
                    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

                    masteridx = -1
	                if  not rsget_Logistics.EOF  then
                        masteridx = rsget_Logistics("idx")
	                end if
	                rsget_Logistics.close

                    if (masteridx = -1) then
                        responseJson = fnCreateCustomJsonResult("500", "requestNo not exists", "500")
                        Response.ContentType = "application/json; charset=utf-8"
                        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
                        response.write responseJson

                        dbget.Close : dbget_Logistics.Close : response.end
                    end if

                    sqlStr = " update "
                    sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_agv_pickup_detail] "
                    sqlStr = sqlStr + " set pickupno = 0 "
                    sqlStr = sqlStr + " where masteridx = " & masteridx & " "
                    dbget_Logistics.Execute sqlStr, affectedRows

                    For Each item In resultData
                        skuCd = resultData.item(item).item("skuCd")
                        finishedSkuQty = resultData.item(item).item("finishedSkuQty")

                        sqlStr = " update "
                        sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_agv_pickup_detail] "
                        sqlStr = sqlStr + " set pickupno = " & finishedSkuQty
                        sqlStr = sqlStr + " where masteridx = " & masteridx & " and skuCd = '" & skuCd & "' "
                        dbget_Logistics.Execute sqlStr, affectedRows
                    Next

		            sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
                    sqlStr = sqlStr + " set status = 100 "
                    sqlStr = sqlStr + " where idx = '" & masteridx & "' "
                    dbget_Logistics.Execute sqlStr

                    responseJson = fnCreateCustomJsonResult("200", "OK", "200")
                case "ONLINE"
                    '// 발주지시 인터페이스
                    if (requestNo = "") then
                        responseJson = fnCreateCustomJsonResult("500", "empty requestNo", "500")
                        Response.ContentType = "application/json; charset=utf-8"
                        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
                        response.write responseJson

                        dbget.Close : dbget_Logistics.Close : response.end
                    end if


                    sqlStr = " select top 1 baljuKey "
                    sqlStr = sqlStr + " from "
                    sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_Logistics_baljumaster] "
                    sqlStr = sqlStr + " where batchNo = '" & requestNo & "' "

                    rsget_Logistics.CursorLocation = adUseClient
                    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

                    baljuKey = -1
	                if  not rsget_Logistics.EOF  then
                        baljuKey = rsget_Logistics("baljuKey")
	                end if
	                rsget_Logistics.close

                    if (baljuKey = -1) then
                        responseJson = fnCreateCustomJsonResult("500", "baljuKey not exists", "500")
                        Response.ContentType = "application/json; charset=utf-8"
                        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
                        response.write responseJson

                        dbget.Close : dbget_Logistics.Close : response.end
                    end if

                    For Each item In resultData
                        skuCd = resultData.item(item).item("skuCd")
                        finishedSkuQty = resultData.item(item).item("finishedSkuQty")

                        sqlStr = " update bi "
                        sqlStr = sqlStr + " set bi.pickupno = " & finishedSkuQty & ", bi.pickupUserID = 'AGV', bi.pickupdate = getdate() "
                        sqlStr = sqlStr + " from "
                        sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_Logistics_baljuipgo] bi "
                        sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] a "
                        sqlStr = sqlStr + " 	on "
                        sqlStr = sqlStr + " 		1 = 1 "
                        sqlStr = sqlStr + " 		and bi.baljuKey = " & baljuKey
                        sqlStr = sqlStr + " 		and a.skuCd = '" & skuCd & "' "
                        sqlStr = sqlStr + " 		and bi.itemgubun = a.itemGubun "
                        sqlStr = sqlStr + " 		and bi.itemid = a.itemid "
                        sqlStr = sqlStr + " 		and bi.itemoption = a.itemoption "
                        dbget_Logistics.Execute sqlStr, affectedRows
                    Next

                    '// 배송실 입고
                    sqlStr = " update b "
                    sqlStr = sqlStr + " set b.ipgono = b.ipgono + b.pickupno "
                    sqlStr = sqlStr + " from "
                    sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_Logistics_baljuipgo] b "
                    sqlStr = sqlStr + " where "
                    sqlStr = sqlStr + " 	1 = 1 "
                    sqlStr = sqlStr + " 	and b.baljuKey = " & baljuKey
                    sqlStr = sqlStr + " 	and b.pickupUserID = 'AGV' "
                    sqlStr = sqlStr + " 	and b.pickupno > 0 "
                    dbget_Logistics.Execute sqlStr

                    responseJson = fnCreateCustomJsonResult("200", "OK", "200")
                case "MANUAL"
                    '// 수기 인터페이스O
                    responseJson = fnCreateCustomJsonResult("200", "SKIP", "200")
                case else
                    '//
                    responseJson = fnCreateCustomJsonResult("500", "unknown pickingOrderTypeCd : " & pickingOrderTypeCd, "500")
            end select

            Set resultJson = Nothing
        else
            responseJson = fnCreateCustomJsonResult("500", "NO BODY", "500")
        End If

        Response.ContentType = "application/json; charset=utf-8"
        response.write responseJson
    case "recvagvstockchange":

        '// {
        '//   "indNo": "1879382",
        '//   "locationCd": "0041-02-02-01",
        '//   "skuCd": "10027443450012",
        '//   "workTypeCd": "ADJUST",
        '//   "workOrderNo": null,
        '//   "workOrderTypeCd": null,
        '//   "workHandlingNo": null,
        '//   "stationCd": "2300",
        '//   "deviceId": "5b73f41f9f70ab96",
        '//   "workerCd": null,
        '//   "prevQty": 4,
        '//   "nextQty": 0,
        '//   "prevStockQty": 4,
        '//   "nextStockQty": 4,
        '//   "prevAdjustQty": 0,
        '//   "nextAdjustQty": -4,
        '//   "indDt": "2021-01-11 11:49:48"
        '// }

        If Request.TotalBytes > 0 Then
            lngBytesCount = Request.TotalBytes
            jsonString = BytesToStr(Request.BinaryRead(lngBytesCount), "UTF-8")

            Call fnWriteLog(mode, jsonString)

            Set resultJson = New aspJson
            resultJson.loadJSON(jsonString)

            skuCd = resultJson.data("skuCd")

            sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_stock_change_log](skuCd)"
            sqlStr = sqlStr + "values('" & skuCd & "')"
            dbget_Logistics.Execute sqlStr

            '// AGV 입출로그 저장
            locationCd = resultJson.data("locationCd")
            itemno = resultJson.data("nextQty") - resultJson.data("prevQty")
            yyyymmdd = Left(Now(), 10)

            if (itemno <> 0) then
                sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_ipchul_log](skuCd, locationCd, itemno, yyyymmdd) "
                sqlStr = sqlStr & " values('" & skuCd & "', '" & locationCd & "', '" & itemno & "', '" & yyyymmdd & "') "
                dbget_Logistics.Execute sqlStr
            end if

            '// 오차입력
            workTypeCd = resultJson.data("workTypeCd")
            ''if (workTypeCd = "ADJUST") then
            ''    itemno = resultJson.data("nextAdjustQty") - resultJson.data("prevAdjustQty")

            ''    sqlStr = " exec [db_summary].[dbo].[sp_Ten_realchekErr_Input] '" & Left(Now(), 10) & "', '" & Left(skuCd, 2) & "', " & Mid(skuCd, 3, Len(skuCd) - 6) * 1 & ", '" & Right(skuCd, 4) & "', " & itemno & ", 'recvagvstockchange' "
            ''    dbget.Execute sqlStr
            ''end if

            Set resultJson = Nothing

            responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        else
            responseJson = fnCreateCustomJsonResult("500", "NO BODY", "500")
        End If

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "recvagvsurveychange":
        '// 재고조사 상태보고
        If Request.TotalBytes > 0 Then
            lngBytesCount = Request.TotalBytes
            jsonString = BytesToStr(Request.BinaryRead(lngBytesCount), "UTF-8")

            Call fnWriteLog(mode, jsonString)

            Set resultJson = New aspJson
            resultJson.loadJSON(jsonString)

            Set resultJson = Nothing

            responseJson = fnCreateCustomJsonResult("00", "OK", "00")
        else
            responseJson = fnCreateCustomJsonResult("99", "NO BODY", "99")
        End If

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "recvagvdistribfinish":
        '// 분배완료 상태보고
        If Request.TotalBytes > 0 Then
            lngBytesCount = Request.TotalBytes
            jsonString = BytesToStr(Request.BinaryRead(lngBytesCount), "UTF-8")

            Call fnWriteLog(mode, jsonString)

            Set resultJson = New aspJson
            resultJson.loadJSON(jsonString)

            Set resultData = resultJson.data("orderList")

            For Each item In resultData
                caseNo = resultData.item(item).item("caseNo")
                Set orderData = resultData.item(item).item("skuList")

                For Each subItem In orderData
                    skuCd = orderData.item(subItem).item("skuCd")
                    orderQty = orderData.item(subItem).item("orderQty")
                    distributionQty = orderData.item(subItem).item("distributionQty")

                    if IsNull(distributionQty) then
                        distributionQty = 0
                    end if

                    '' B : 분배완료
                    '' M : 미배

                    sqlStr = " update d "
                    sqlStr = sqlStr + " set d.distributeNo = " & distributionQty & ", d.stockoutyn = (case when d.itemno = " & distributionQty & " then 'N' else 'Y' end), d.dasstate = (case when d.itemno = " & distributionQty & " then 'B' else 'M' end) "
                    sqlStr = sqlStr + " from "
                    sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_Logistics_order_master] m "
                    sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_Logistics_order_detail] d "
                    sqlStr = sqlStr + " 	on "
                    sqlStr = sqlStr + " 		1 = 1 "
                    sqlStr = sqlStr + " 		and m.siteSeq = d.siteSeq "
                    sqlStr = sqlStr + " 		and m.orderserial = d.orderserial "
                    sqlStr = sqlStr + " 		and m.siteSeq = '10' "
                    sqlStr = sqlStr + " 		and m.orderserial = '" & caseNo & "' "
                    sqlStr = sqlStr + " 		and d.itemid not in (0, 100) "
                    sqlStr = sqlStr + " 		and d.isupchebeasong = 'N' "
                    sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] i "
                    sqlStr = sqlStr + " 	on "
                    sqlStr = sqlStr + " 		1 = 1 "
                    sqlStr = sqlStr + " 		and i.skuCd = '" & skuCd & "' "
                    sqlStr = sqlStr + " 		and i.itemGubun = d.itemgubun "
                    sqlStr = sqlStr + " 		and i.itemid = d.itemid "
                    sqlStr = sqlStr + " 		and i.itemoption = d.itemoption "
                    dbget_Logistics.Execute sqlStr, affectedRows
                Next

                Set orderData = Nothing
            Next

            Set resultJson = Nothing

            responseJson = fnCreateCustomJsonResult("00", "OK", "00")
        else
            responseJson = fnCreateCustomJsonResult("99", "NO BODY", "99")
        End If

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "updagvstockchange":
        minidx = -1
        maxidx = -1

        sqlStr = " select IsNull(min(idx), 0) as minidx, IsNull(max(idx), 0) as maxidx "
        sqlStr = sqlStr + " from [db_aLogistics].[dbo].[tbl_agv_stock_change_log] "
        sqlStr = sqlStr + " where changeApplied = 'N' "

        rsget_Logistics.CursorLocation = adUseClient
        rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

        if  not rsget_Logistics.EOF  then
            minidx = rsget_Logistics("minidx")
            maxidx = rsget_Logistics("maxidx")
        end if
        rsget_Logistics.close

        if (CDbl(minidx) > CDbl(0)) then
            sqlStr = " select skuCd "
            sqlStr = sqlStr + " from [db_aLogistics].[dbo].[tbl_agv_stock_change_log] "
            sqlStr = sqlStr + " where changeApplied = 'N' "
            sqlStr = sqlStr + " and IDX between " & minidx & " and " & maxidx
            sqlStr = sqlStr + " and IDX <= " & (CDbl(minidx) + 100)
            sqlStr = sqlStr + " group by skuCd "
            rsget_Logistics.CursorLocation = adUseClient
            rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

            brandArray = ""
            skuCdArray = ""
            if  not rsget_Logistics.EOF  then
                do until rsget_Logistics.eof
                    skuCdArray = skuCdArray & "," & rsget_Logistics("skuCd")
			        rsget_Logistics.moveNext
		        loop
            end if
            rsget_Logistics.close

            brandArray = Split(brandArray, ",")
            skuCdArray = Split(skuCdArray, ",")

            Call fnGetProductStockInfoByBrandBySkuCd(brandArray, skuCdArray, jsonResult)

            Call fnWriteLog(mode, jsonResult)

            Call fnParseResultProductStockInfoByBrandBySkuCd(jsonResult, resultCode, resultMessage, failCode)

            if (resultCode = "00") then
                Call fnSaveStockInfo(jsonResult, True)
                responseJson = fnCreateCustomJsonResult("200", "OK", "200")

                sqlStr = " update "
                sqlStr = sqlStr + " [db_aLogistics].[dbo].[tbl_agv_stock_change_log] "
                sqlStr = sqlStr + " set changeApplied = 'Y', lastupdate = getdate() "
                sqlStr = sqlStr + " where changeApplied = 'N' "
                sqlStr = sqlStr + " and IDX between " & minidx & " and " & maxidx
                sqlStr = sqlStr + " and IDX <= " & (CDbl(minidx) + 100)
                dbget_Logistics.Execute sqlStr
            else
                responseJson = fnCreateCustomJsonResult("500", resultMessage, "500")
            end if
        else
            responseJson = fnCreateCustomJsonResult("200", "OK", "200")
        end if

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case "recvagvWareHouseChg":
        '// 재고위치 변경정보 수신
        If Request.TotalBytes > 0 Then
            lngBytesCount = Request.TotalBytes
            jsonString = BytesToStr(Request.BinaryRead(lngBytesCount), "UTF-8")

            Call fnWriteLog(mode, jsonString)

            Set resultJson = New aspJson
            resultJson.loadJSON(jsonString)

            skuCd = resultJson.data("skuCd")
            warehouseCd = resultJson.data("warehouseTypeCd")

            if (warehouseCd = "BULK_WAREHOUSE") then
                warehouseCd = "BLK"
            else
                warehouseCd = "AGV"
            end if

            sqlStr = " update [db_summary].[dbo].[tbl_current_agvstock_summary] "
            sqlStr = sqlStr + " set warehouseCd = '" & warehouseCd & "', lastupdate = getdate() "
            sqlStr = sqlStr + " where skuCd = '" & skuCd & "' "
            dbget.Execute sqlStr

            Set resultJson = Nothing

            responseJson = fnCreateCustomJsonResult("00", "OK", "00")
        else
            responseJson = fnCreateCustomJsonResult("99", "NO BODY", "99")
        End If

        Response.ContentType = "application/json; charset=utf-8"
        Call Response.AddHeader("Access-Control-Allow-Origin", "*")
        response.write responseJson
    case else:
        response.write "잘못된 접근입니다..."
end select

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
