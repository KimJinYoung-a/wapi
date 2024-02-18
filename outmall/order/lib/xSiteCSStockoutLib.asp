<%

CONST ssgAPIURL = "http://eapi.ssgadm.com"
CONST ssgSSLAPIURL = "https://eapi.ssgadm.com"
CONST ssgApiKey = "18a8d870-12a7-4b36-afaf-1e9d38e2b988"
function GetCSStockout_ssg(sellsite, mode, detailidx, orderserial)
	if (mode = "stockoutOne") or (mode = "stockoutCnclOne") then
		Call GetCSStockoutOne_ssg(sellsite, mode, detailidx)
	elseif (mode = "stockoutAll") then
		Call GetCSStockoutAll_ssg(sellsite, mode, orderserial)
	else
		response.write "잘못된 접근입니다. mode = [" & mode & "]"
		dbget.close : response.end
	end if
end function

Function getApiUrl(mallid)
	Select Case mallid
		Case "lotteon"
			If application("Svr_Info") = "Dev" Then
				getApiUrl = "https://dev-openapi.lotteon.com"
			Else
				getApiUrl = "https://openapi.lotteon.com"
			End If
	End Select
End Function

Function getApiKey(mallid)
	Select Case mallid
		Case "lotteon"
			If application("Svr_Info") = "Dev" Then
				getApiKey = "5d5b2cb498f3d20001665f4e5451c4d923ac4e2c95df619996f35476"
			Else
				getApiKey = "5d5b2cb498f3d20001665f4e18a41621005d4c1ba262804ec7a10732"
			End If
	End Select
End Function

function GetCSStockout_coupang(sellsite, mode, orderserial, detailidx, itemno)
	if (mode = "cancelAll") then
		Call GetCSStockoutCancelSelected_coupang(sellsite, mode, orderserial, detailidx, itemno)
	else
		response.write "잘못된 접근입니다. mode = [" & mode & "]"
		dbget.close : response.end
	end if
end function

function GetCSStockout_interpark(sellsite, mode, orderserial, detailidx, itemno)
	if (mode = "cancelAll") then
		Call GetCSStockoutCancelSelected_interpark(sellsite, mode, orderserial, detailidx, itemno)
	else
		response.write "잘못된 접근입니다. mode = [" & mode & "]"
		dbget.close : response.end
	end if
end function

function GetCSStockout_11st1010(sellsite, mode, detailidx, orderserial)
	if (mode = "stockoutOne") then
		Call GetCSStockoutOne_11st1010(sellsite, mode, detailidx, orderserial)
	else
		response.write "잘못된 접근입니다. mode = [" & mode & "]"
		dbget.close : response.end
	end if
end function

function GetCSStockout_nvstorefarm(sellsite, mode, detailidx, orderserial)
	if (mode = "stockoutOne") then
		Call GetCSStockoutOne_nvstorefarm(sellsite, mode, detailidx, orderserial)
	else
		response.write "잘못된 접근입니다. mode = [" & mode & "]"
		dbget.close : response.end
	end if
end function

Function GetCSStockout_Lotteon(sellsite, mode, detailidx, orderserial)
	if (mode = "stockoutOne") then
		Call GetCSStockoutOne_Lotteon(sellsite, mode, detailidx, orderserial)
	else
		response.write "잘못된 접근입니다. mode = [" & mode & "]"
		dbget.close : response.end
	end if
end function

function GetCSStockout_gmarket1010(sellsite, mode, detailidx, orderserial)
	if (mode = "cancelAll") then
		Call GetCSStockoutCancelOne_gmarket1010(sellsite, mode, detailidx, orderserial)
	else
		response.write "잘못된 접근입니다. mode = [" & mode & "]"
		dbget.close : response.end
	end if
end function

function GetCSStockout_WMP(sellsite, mode, detailidx, orderserial)
	if (mode = "cancelAll") then
		Call GetCSStockoutCancelOne_WMP(sellsite, mode, detailidx, orderserial)
	else
		response.write "잘못된 접근입니다. mode = [" & mode & "]"
		dbget.close : response.end
	end if
end function

function GetCSStockout_WMPfashion(sellsite, mode, detailidx, orderserial)
	if (mode = "cancelAll") then
		Call GetCSStockoutCancelOne_WMPfashion(sellsite, mode, detailidx, orderserial)
	else
		response.write "잘못된 접근입니다. mode = [" & mode & "]"
		dbget.close : response.end
	end if
end function

const Option_TLS12 = 2048
const Option_TLS1 = 512
const Option_TLS = 128
function GetCSStockoutOne_ssg(sellsite, mode, detailidx)
	'// shppNo			배송번호
	'// shppSeq			배송순번
	'// scEvnt			등록/삭제구분(I 등록, D 삭제)
	'// shortgRsnCd		등록/삭제사유(08 상품정보오류, 09 결품)
	'// shortgProcDtlc	판매불가사유내용
	'// itemId			상품코드

    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc
    Dim errCount : errCount = 0
    Dim AssignedCNT : AssignedCNT=0
	Dim shppNo, shppSeq, scEvnt, shortgRsnCd, shortgProcDtlc, itemId
	Dim AssignedRow

	shppNo = ""
	strSql = " select top 1 T.beasongNum11st as shppNo, T.reserve01 as shppSeq, 'I' as scEvnt, '09' as shortgRsnCd, l.reqaddstr as shortgProcDtlc, T.outMallGoodsNo as itemId " & vbCrLf
	strSql = strSql + " from " & vbCrLf
	strSql = strSql + " 	[db_order].[dbo].[tbl_order_detail] d " & vbCrLf
	strSql = strSql + " 	join [db_order].[dbo].[tbl_order_master] m on d.orderserial = m.orderserial " & vbCrLf
	strSql = strSql + " 	join [db_temp].[dbo].[tbl_mibeasong_list] l on d.idx = l.detailidx " & vbCrLf
	strSql = strSql + " 	join db_temp.dbo.tbl_xSite_TMPOrder T " & vbCrLf
	strSql = strSql + " 	on " & vbCrLf
	strSql = strSql + " 		1 = 1 " & vbCrLf
	strSql = strSql + " 		and T.OrderSerial = m.orderserial " & vbCrLf
	strSql = strSql + " 		and T.matchItemID = d.itemid " & vbCrLf
	strSql = strSql + " 		and T.matchitemoption = d.itemoption " & vbCrLf
	strSql = strSql + " where " & vbCrLf
	strSql = strSql + " 	1 = 1 " & vbCrLf
	strSql = strSql + " 	and d.idx = " & detailidx & vbCrLf
	strSql = strSql + " 	and m.sitename = '" & sellsite & "' " & vbCrLf
	strSql = strSql + " 	and l.code = '05' " & vbCrLf
    rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
    If Not rsget.Eof Then
    	shppNo 		= rsget("shppNo")
		shppSeq 	= rsget("shppSeq")
		scEvnt 		= rsget("scEvnt")
		if (mode="stockoutCnclOne") then
			scEvnt = "D"
		end if
		shortgRsnCd 	= rsget("shortgRsnCd")
		shortgProcDtlc 	= rsget("shortgProcDtlc")
		itemId 			= rsget("itemId")
    End If
    rsget.Close

	if (shppNo = "") then
		response.write "에러 : 내역없음<br />"
		exit function
	end if


    Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/saveNoSellRequestRegist.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestNoSellRequestRegist>"
    requestBody = requestBoDy&"<shppNo>"&shppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&shppSeq&"</shppSeq>"
	requestBody = requestBoDy&"<scEvnt>"&scEvnt&"</scEvnt>"
	''requestBody = requestBoDy&"<shortgProcDtlc><![CDATA[" & shortgProcDtlc & "]]></shortgProcDtlc>"
	requestBody = requestBoDy&"<shortgProcDtlc><![CDATA[품절로 배송불가]]></shortgProcDtlc>"				'// 품절사유 문구 고정 : 고객센터 요청사항
	requestBody = requestBoDy&"<shortgRsnCd>"&shortgRsnCd&"</shortgRsnCd>"
	requestBody = requestBoDy&"<itemId>"&itemId&"</itemId>"
    requestBody = requestBoDy&"</requestNoSellRequestRegist>"
    ''response.write requestBoDy
	objXML.send(requestBody)
	rw objXML.status
	''response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

	Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
		xmlDOM.async = False
		xmlDOM.loadXML(objXML.responseText)

		ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
		ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
		ssgresultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text
	Set LagrgeNode = nothing
	Set xmlDOM = nothing
	Set objXML = nothing

	AssignedRow = 0
	if (ssgresultCode="00") then
        strSql = "Update [db_temp].[dbo].[tbl_mibeasong_list] "
		if (scEvnt = "I") then
			strSql = strSql & "	Set isSendAPI='Y'"
		else
			strSql = strSql & "	Set isSendAPI='N'"
		end if
        strSql = strSql & "	where detailidx='"&detailidx&"'"&VBCRLF
        dbget.Execute strSql,AssignedRow

		if (AssignedRow = 1) then
			rw "OK"
		else
			rw "에러 : 알수 없는 오류입니다.[0]"
		end if
	else
		rw 	ssgresultCode&":"&ssgresultMessage&":"&ssgresultDesc
	end if
end function

function GetCSStockoutAll_ssg(sellsite, mode, orderserial)
	'// stockoutOne
	'// stockoutAll
end function

function GetCSStockoutCancelSelected_coupang(sellsite, mode, orderserial, detailidx, itemno)
	Dim objXmlHttp, requestBody
	Dim url, oJSON, detailidxOrder
	dim strSql, i
	dim orderId, vendorItemIds, receiptCounts
	dim IsSuccess, errMsg, AssignedRow

	orderId = ""
	vendorItemIds = ""
	'// 콤마제거
	detailidx = Mid(detailidx, 2, 4000)
	itemno = Mid(itemno, 2, 4000)

	detailidx = Split(detailidx, ",")
	itemno = Split(itemno, ",")

	for i = 0 to UBound(detailidx)
		if CStr(detailidx(i)) <> "" then
			AssignedRow = 0
			IsSuccess = GetCSStockoutCancelOne_coupang(sellsite, mode, orderserial, detailidx(i), itemno(i), errMsg)
			if (IsSuccess) then
				strSql = "Update [db_temp].[dbo].[tbl_mibeasong_list] "
				strSql = strSql & "	Set isSendAPI='Y'"
				strSql = strSql & "	where detailidx='"&detailidx(i)&"'"&VBCRLF
				dbget.Execute strSql,AssignedRow

				if (AssignedRow = 1) then
					rw "OK"
				else
					rw "에러 : 알수 없는 오류입니다.[1]"
				end if
			else
				response.write "<font color='red'>주문취소 실패!!</font> : " & errMsg & "<br />"
			end if
		end if
	next
end function

function GetCSStockoutCancelOne_coupang(sellsite, mode, orderserial, detailidx, itemno, byRef errMsg)
	Dim objXmlHttp, requestBody
	Dim url, oJSON
	dim orderId, vendorItemIds, receiptCounts
	dim strSql, i

	strSql = " select top 1 T.OutMallOrderSerial, T.outMallOptionNo, d.idx as detailidx "
	strSql = strSql + " from "
	strSql = strSql + " 	[db_temp].[dbo].[tbl_xSite_TMPOrder] T "
	strSql = strSql + " 	join [db_order].[dbo].[tbl_order_detail] d "
	strSql = strSql + " 	on "
	strSql = strSql + " 		1 = 1 "
	strSql = strSql + " 		and T.OrderSerial = d.orderserial "
	strSql = strSql + " 		and T.matchItemID = d.itemid "
	strSql = strSql + " 		and T.matchitemoption = d.itemoption "
	strSql = strSql + " where "
	strSql = strSql + " 	1 = 1 "
	strSql = strSql + " 	and sellsite = '" & sellsite & "' "
	strSql = strSql + " 	and T.OrderSerial = '" & orderserial & "' "
	strSql = strSql + " 	and d.idx = " & detailidx & " "
    rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

	If Not rsget.Eof Then
		orderId 		= rsget("OutMallOrderSerial")
		vendorItemIds	= rsget("outMallOptionNo")
		receiptCounts	= itemno
	end if
    rsget.Close

	if (orderId = "") or (vendorItemIds = "") then
		errMsg = "에러 : 내역없음<br />"
		exit function
	end if

	requestBody = "{"
	requestBody = requestBody & "  ""orderId"": " & orderId & ","
	requestBody = requestBody & "  ""vendorItemIds"": ["
	requestBody = requestBody & "    " & vendorItemIds
	requestBody = requestBody & "  ],"
	requestBody = requestBody & "  ""receiptCounts"": ["
	requestBody = requestBody & "    " & receiptCounts
	requestBody = requestBody & "  ],"
	requestBody = requestBody & "  ""bigCancelCode"": ""CANERR"","
	requestBody = requestBody & "  ""middleCancelCode"": ""CCTTER"","
	requestBody = requestBody & "  ""userId"": ""XXXXXXXX"","			'// XAPI 에서 설정
	requestBody = requestBody & "  ""vendorId"": ""XXXXXXXX"""			'// XAPI 에서 설정
	requestBody = requestBody & "}"
	''response.write requestBody

	url = "http://xapi.10x10.co.kr:8080/Orders/Coupang/productcancel"

	Set objXmlHttp = Server.CreateObject("Microsoft.XMLHTTP")

	objXmlHttp.Open "POST", url, False
	objXmlHttp.SetRequestHeader "Content-Type", "application/json"
	''objXmlHttp.SetRequestHeader "User-Agent", "ASP/3.0"

	objXmlHttp.Send(requestBody)

	if (objXmlHttp.status <> "200") then
		response.write CStr(objXmlHttp.ResponseText) & "<br />"
	else
		Set oJSON = New aspJSON
		oJSON.loadJSON(CStr(objXmlHttp.ResponseText))
		select case oJSON.data("code")
			case "SUCCESS"
				GetCSStockoutCancelOne_coupang = True
				errMsg = ""
			case "PARTIAL"
				GetCSStockoutCancelOne_coupang = False
				errMsg = oJSON.data("message")
			case "FAIL"
				GetCSStockoutCancelOne_coupang = False
				errMsg = oJSON.data("message")
			case else
				'//
		end select
	end if

	Set oJSON = Nothing
	Set objXmlHttp = Nothing
end function

function GetCSStockoutCancelSelected_interpark(sellsite, mode, orderserial, detailidx, itemno)
	Dim url, errMsg
	dim strSql, i, AssignedRow, IsSuccess
	dim entrId      : entrId="10X10"                    ''고정값 : 제휴업체_ID

	'// 콤마제거
	detailidx = Mid(detailidx, 2, 4000)
	itemno = Mid(itemno, 2, 4000)

	detailidx = Split(detailidx, ",")
	itemno = Split(itemno, ",")

	for i = 0 to UBound(detailidx)
		if CStr(detailidx(i)) <> "" then
			AssignedRow = 0
			IsSuccess = GetCSStockoutCancelOne_interpark(sellsite, mode, orderserial, detailidx(i), itemno(i), errMsg)
			if (IsSuccess) then
				strSql = "Update [db_temp].[dbo].[tbl_mibeasong_list] "
				strSql = strSql & "	Set isSendAPI='Y'"
				strSql = strSql & "	where detailidx='"&detailidx(i)&"'"&VBCRLF
				dbget.Execute strSql,AssignedRow

				if (AssignedRow = 1) then
					rw "OK"
				else
					rw "에러 : 알수 없는 오류입니다.[1]"
				end if
			else
				response.write "<font color='red'>주문취소 실패!!</font> : " & errMsg & "<br />"
			end if
		end if
	next
end function

CONST interparkAPIURL = "http://ipss1.interpark.com"
function GetCSStockoutCancelOne_interpark(sellsite, mode, orderserial, detailidx, itemno, byRef errMsg)
	Dim xmlURL, objXML, objData, xmlDOM, obj
	dim strSql, i, AssignedRow, IsSuccess
	dim entrId      : entrId="10X10"                    ''고정값 : 제휴업체_ID
	dim ordclmNo, ordSeq, optPrdTp, optOrdSeqList

	strSql = " select top 1 T.OutMallOrderSerial, T.OrgDetailKey, d.idx as detailidx "
	strSql = strSql + " from "
	strSql = strSql + " 	[db_temp].[dbo].[tbl_xSite_TMPOrder] T "
	strSql = strSql + " 	join [db_order].[dbo].[tbl_order_detail] d "
	strSql = strSql + " 	on "
	strSql = strSql + " 		1 = 1 "
	strSql = strSql + " 		and T.OrderSerial = d.orderserial "
	strSql = strSql + " 		and T.matchItemID = d.itemid "
	strSql = strSql + " 		and T.matchitemoption = d.itemoption "
	strSql = strSql + " where "
	strSql = strSql + " 	1 = 1 "
	strSql = strSql + " 	and sellsite = '" & sellsite & "' "
	strSql = strSql + " 	and T.OrderSerial = '" & orderserial & "' "
	strSql = strSql + " 	and d.idx = " & detailidx
	''response.write strSql
    rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

	ordclmNo = ""
	If Not rsget.Eof Then
		ordclmNo 		= rsget("OutMallOrderSerial")
		ordSeq			= rsget("OrgDetailKey")
		optPrdTp		= "01"
		optOrdSeqList	= rsget("OrgDetailKey")
	end if
    rsget.Close

	if (ordclmNo = "") then
		errMsg = "에러 : 내역없음<br />"
		exit function
	end if


	xmlURL = "https://joinapi.interpark.com"
	xmlURL = xmlURL + "/order/OrderClmAPI.do?_method=cnclOutOfStockReqForComm&sc.entrId=" & entrId & "&sc.ordclmNo=" & ordclmNo & "&sc.ordSeq=" & ordSeq & "&sc.optPrdTp=" & optPrdTp & "&sc.optOrdSeqList=" & ordSeq

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.setTimeouts 5000,80000,80000,80000
	objXML.Send()

	if objXML.Status = "200" then
		objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		''response.write objData
		''dbget.close : response.end
	else
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML replace(objData,"&","＆")

	Set obj = xmlDOM.selectSingleNode("/ORDER_LIST/RESULT/CODE")

	if obj is Nothing then
		if IsAutoScript then
			''response.write "내역없음 : 종료"
		end if

		GetCSStockoutCancelOne_interpark = False
		errMsg = "내역없음 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	if (obj.text <> "000") then
		GetCSStockoutCancelOne_interpark = False
		errMsg = xmlDOM.selectSingleNode("/ORDER_LIST/RESULT/MESSAGE").text
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	GetCSStockoutCancelOne_interpark = True
	Set xmlDOM = Nothing
	Set objXML = Nothing
end function

CONST APIkey11st1010 = "a2319e071dbc304243ee60abd07e9664"
function GetCSStockoutOne_11st1010(sellsite, mode, detailidx, orderserial)
	Dim xmlURL, objXML, objData, xmlDOM, obj
	dim strSql, i, AssignedRow, IsSuccess
	dim ordNo, ordPrdSeq, ordCnRsnCd, ordCnDtlsRsn
	dim result_code, result_text

	'' ordCnRsnCd
	'' → 06 : 배송 지연 예상
	'' → 07 : 상품/가격 정보 잘못 입력
	'' → 08 : 상품 품절(전체옵션)
	'' → 09 : 옵션 품절(해당옵션)
	'' → 10 : 고객변심
	'' → 99 : 기타

	strSql = " select top 1 T.OutMallOrderSerial, T.OrgDetailKey, d.idx as detailidx, d.itemoption "
	strSql = strSql + " from "
	strSql = strSql + " 	[db_temp].[dbo].[tbl_xSite_TMPOrder] T "
	strSql = strSql + " 	join [db_order].[dbo].[tbl_order_detail] d "
	strSql = strSql + " 	on "
	strSql = strSql + " 		1 = 1 "
	strSql = strSql + " 		and T.OrderSerial = d.orderserial "
	strSql = strSql + " 		and T.matchItemID = d.itemid "
	strSql = strSql + " 		and T.matchitemoption = d.itemoption "
	strSql = strSql + " where "
	strSql = strSql + " 	1 = 1 "
	strSql = strSql + " 	and sellsite = '" & sellsite & "' "
	strSql = strSql + " 	and T.OrderSerial = '" & orderserial & "' "
	strSql = strSql + " 	and d.idx = " & detailidx
	''response.write strSql
    rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

	ordNo = ""
	If Not rsget.Eof Then
		ordNo 			= rsget("OutMallOrderSerial")
		ordPrdSeq		= rsget("OrgDetailKey")
		ordCnRsnCd		= "09"
		if (rsget("itemoption") = "0000") then
			ordCnRsnCd		= "08"
		end if
		ordCnDtlsRsn	= "품절로출고불가"
	end if
    rsget.Close

	if (ordNo = "") then
		response.write "<font color='red'>결품등록 실패!!</font> : 에러 : 내역없음<br />"
		exit function
	end if


	xmlURL = "https://api.11st.co.kr/rest/claimservice/reqrejectorder/" & ordNo & "/" & ordPrdSeq & "/" & ordCnRsnCd & "/" & ordCnDtlsRsn

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey11st1010&""
		objXML.send()

	if objXML.Status = "200" then
		objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		''response.write objData
		''dbget.close : response.end
	else
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML objData
		result_code = xmlDOM.getElementsByTagName("result_code").item(0).text
		result_text = xmlDOM.getElementsByTagName("result_text").item(0).text

	if (result_code <> "0") then
		'// 에러
		response.write "<font color='red'>결품등록 실패!!</font> : 에러 : " & result_text & "<br />"
	else
		strSql = "Update [db_temp].[dbo].[tbl_mibeasong_list] "
		strSql = strSql & "	Set isSendAPI='Y'"
		strSql = strSql & "	where detailidx='"&detailidx&"'"&VBCRLF
		dbget.Execute strSql,AssignedRow

		rw "OK"
	end if
end function

function GetCSStockoutOne_nvstorefarm(sellsite, mode, detailidx, orderserial)
    dim strSql
    dim OutMallOrderSerial, OrgDetailKey, resultStr

	strSql = " select top 1 T.OutMallOrderSerial, T.OrgDetailKey, d.idx as detailidx "
	strSql = strSql + " from "
	strSql = strSql + " 	[db_temp].[dbo].[tbl_xSite_TMPOrder] T "
	strSql = strSql + " 	join [db_order].[dbo].[tbl_order_detail] d "
	strSql = strSql + " 	on "
	strSql = strSql + " 		1 = 1 "
	strSql = strSql + " 		and T.OrderSerial = d.orderserial "
	strSql = strSql + " 		and T.matchItemID = d.itemid "
	strSql = strSql + " 		and T.matchitemoption = d.itemoption "
	strSql = strSql + " where "
	strSql = strSql + " 	1 = 1 "
	strSql = strSql + " 	and sellsite = '" & sellsite & "' "
	strSql = strSql + " 	and T.OrderSerial = '" & orderserial & "' "
	strSql = strSql + " 	and d.idx = " & detailidx
	''response.write strSql
    rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

	OutMallOrderSerial = ""
	If Not rsget.Eof Then
		OutMallOrderSerial 		= rsget("OutMallOrderSerial")
        OrgDetailKey 			= rsget("OrgDetailKey")
	end if
    rsget.Close

	if (OutMallOrderSerial = "") then
		response.write "<font color='red'>결품등록 실패!!</font> : 에러 : 내역없음<br />"
		exit function
	end if

    resultStr = GetOrderDetailStatus_nvstorefarm(OrgDetailKey)

    '// NO_CANCEL_INFO
    '// CANCEL_REQUEST, CANCELING, CANCEL_DONE, CANCEL_REJECT
    if resultStr = "NO_CANCEL_INFO" then
        resultStr = CancelSell_nvstorefarm(OrgDetailKey)
        response.write resultStr
    else
    	select case resultStr
            case "CANCELING"
                response.write "취소진행중입니다."
            case "CANCEL_DONE"
                response.write "취소완료 상태입니다."
            case "CANCEL_REJECT"
                response.write "취소거부 상태입니다."
            case "CANCEL_REQUEST"
                resultStr = ApproveCancel_nvstorefarm(OrgDetailKey)
                response.write resultStr
            case else
                response.write "에러 : 알 수 없는 에러입니다.(" & resultStr & ")"
        end select
    end if

    if resultStr = "OK" then
		strSql = "Update [db_temp].[dbo].[tbl_mibeasong_list] "
		strSql = strSql & "	Set isSendAPI='Y'"
		strSql = strSql & "	where detailidx='"&detailidx&"'"&VBCRLF
		dbget.Execute strSql,AssignedRow
    end if

end function

function CancelSell_nvstorefarm(OrgDetailKey)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj, objData, objDetail, objDetailArr
	dim i, j, k
	dim startdate, enddate
	dim OutMallOrderSerial, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd
	dim reqID, keyGenerated, cryptoLib
	dim ResponseType, cancelStatus

	iServ		= "SellerService41"
	iCcd		= "CancelSale"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If
	''response.write xmlURL

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		reqID = "tenten"
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">" + vbCrLf
	strRst = strRst & "	<soapenv:Header/>" + vbCrLf
	strRst = strRst & "	<soapenv:Body>" + vbCrLf
	strRst = strRst & "		<sel:CancelSale>" + vbCrLf
	strRst = strRst & "			<sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>" + vbCrLf
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>" + vbCrLf
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>" + vbCrLf
	strRst = strRst & "			</sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>" + vbCrLf
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>" + vbCrLf
	strRst = strRst & "			<sel:Version>4.1</sel:Version>" + vbCrLf
    strRst = strRst & "			<sel:ProductOrderID>" & OrgDetailKey & "</sel:ProductOrderID>" + vbCrLf
    strRst = strRst & "			<sel:CancelReasonCode>SOLD_OUT</sel:CancelReasonCode>" + vbCrLf
	strRst = strRst & "		</sel:CancelSale>" + vbCrLf
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	''response.write strRst

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
	''response.write objXML.responseText & "<br /><br />"
	''response.flush

	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		response.write "오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

    CancelSell_nvstorefarm = "OK"
end function

function ApproveCancel_nvstorefarm(OrgDetailKey)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj, objData, objDetail, objDetailArr
	dim i, j, k
	dim startdate, enddate
	dim OutMallOrderSerial, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd
	dim reqID, keyGenerated, cryptoLib
	dim ResponseType, cancelStatus

	iServ		= "SellerService41"
	iCcd		= "ApproveCancelApplication"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If
	''response.write xmlURL

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		reqID = "tenten"
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">" + vbCrLf
	strRst = strRst & "	<soapenv:Header/>" + vbCrLf
	strRst = strRst & "	<soapenv:Body>" + vbCrLf
	strRst = strRst & "		<sel:ApproveCancelApplication>" + vbCrLf
	strRst = strRst & "			<sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>" + vbCrLf
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>" + vbCrLf
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>" + vbCrLf
	strRst = strRst & "			</sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>" + vbCrLf
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>" + vbCrLf
	strRst = strRst & "			<sel:Version>4.1</sel:Version>" + vbCrLf
    strRst = strRst & "			<sel:ProductOrderID>" & OrgDetailKey & "</sel:ProductOrderID>" + vbCrLf
	strRst = strRst & "		</sel:ApproveCancelApplication>" + vbCrLf
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	''response.write strRst

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
	''response.write objXML.responseText & "<br /><br />"
	''response.flush

	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		response.write "오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

    ApproveCancel_nvstorefarm = "OK"
end function

function GetOrderDetailStatus_nvstorefarm(OrgDetailKey)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj, objData, objDetail, objDetailArr
	dim i, j, k
	dim startdate, enddate
	dim OutMallOrderSerial, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd
	dim reqID, keyGenerated, cryptoLib
	dim ResponseType, cancelStatus

	iServ		= "SellerService41"
	iCcd		= "GetProductOrderInfoList"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If
	''response.write xmlURL

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		reqID = "tenten"
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">" + vbCrLf
	strRst = strRst & "	<soapenv:Header/>" + vbCrLf
	strRst = strRst & "	<soapenv:Body>" + vbCrLf
	strRst = strRst & "		<sel:GetProductOrderInfoListRequest>" + vbCrLf
	strRst = strRst & "			<sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>" + vbCrLf
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>" + vbCrLf
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>" + vbCrLf
	strRst = strRst & "			</sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>" + vbCrLf
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>" + vbCrLf
	strRst = strRst & "			<sel:Version>4.1</sel:Version>" + vbCrLf
    strRst = strRst & "			<sel:ProductOrderIDList>" & OrgDetailKey & "</sel:ProductOrderIDList>" + vbCrLf
	strRst = strRst & "		</sel:GetProductOrderInfoListRequest>" + vbCrLf
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	''response.write strRst

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
	''response.write objXML.responseText & "<br /><br />"
	''response.flush

	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		response.write "오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	if CLng(xmlDOM.getElementsByTagName("n:ReturnedDataCount").item(0).text) <> 1 then
		response.write "건수 불일치 오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	set objArr = xmlDOM.getElementsByTagName("n:ProductOrderInfoList")

    GetOrderDetailStatus_nvstorefarm = ""
    For each obj in objArr
        if obj.selectSingleNode("n:CancelInfo") is Nothing then
            '// 취소신청 없음
            GetOrderDetailStatus_nvstorefarm = "NO_CANCEL_INFO"
        else
            '// CANCEL_REQUEST, CANCELING, CANCEL_DONE, CANCEL_REJECT
            GetOrderDetailStatus_nvstorefarm = obj.selectSingleNode("n:CancelInfo/n:ClaimStatus").text
        end if
    Next
end function

Public Function getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iserv, ioper)
	Dim cryptoLib, oLicense, osecretKey, otimeStamp, osignature
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
		If (application("Svr_Info") = "Dev") Then
			iaccessLicense = "01000100004b035a25d67f991849cad1c7042b8da528d13e9ddce6878f2e43ac88080e0a5e" 'AccessLicense Key 입력, PDF파일참조
			osecretKey = "AQABAAAWPWagCrPjFQnFEtxs5j+oyZFwuzomdNq0XZSricPuMw=="  'SecreKey 입력, PDF파일참조
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		Else
			iaccessLicense = "010001000019133c715650b9c85b820961612f2b90b431ddd8654b42c097c4df1a43d0be09" 'AccessLicense Key 입력, PDF파일참조
			osecretKey = "AQABAADX6Hz/wORFJS5pSIy4KQXkH83gC9G1aXChxBjcnUMqWw=="  'SecreKey 입력, PDF파일참조
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		End If
	Set cryptoLib = nothing
End Function

function GetCSStockoutCancelOne_gmarket1010(sellsite, mode, detailidx, orderserial)
    dim strSql
    dim OutMallOrderSerial, OrgDetailKey, resultStr, Comment, AssignedRow

	strSql = " select top 1 T.OutMallOrderSerial, T.OrgDetailKey, d.idx as detailidx "
	strSql = strSql + " from "
	strSql = strSql + " 	[db_temp].[dbo].[tbl_xSite_TMPOrder] T "
	strSql = strSql + " 	join [db_order].[dbo].[tbl_order_detail] d "
	strSql = strSql + " 	on "
	strSql = strSql + " 		1 = 1 "
	strSql = strSql + " 		and T.OrderSerial = d.orderserial "
	strSql = strSql + " 		and T.matchItemID = d.itemid "
	strSql = strSql + " 		and T.matchitemoption = d.itemoption "
	strSql = strSql + " where "
	strSql = strSql + " 	1 = 1 "
	strSql = strSql + " 	and sellsite = '" & sellsite & "' "
	strSql = strSql + " 	and T.OrderSerial = '" & orderserial & "' "
	strSql = strSql + " 	and d.idx = " & detailidx
	''response.write strSql
    rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

	OutMallOrderSerial = ""
	If Not rsget.Eof Then
		OutMallOrderSerial 		= rsget("OutMallOrderSerial")
        OrgDetailKey 			= rsget("OrgDetailKey")
	end if
    rsget.Close

	if (OutMallOrderSerial = "") then
		response.write "<font color='red'>결품등록(주문취소) 실패!!</font> : 에러 : 내역없음<br />"
		exit function
	end if

    resultStr = GetCSOrderCancelStateOne_gmarket1010(sellsite, OutMallOrderSerial, OrgDetailKey)

    select case resultStr
        case "B007"
            response.write "이미 취소된 주문입니다."
        case "B001"
            response.write "제휴상태 : 취소요청중<br />"
            resultStr = GetCSOrderCancelRequestConfirmOne_gmarket1010(sellsite, OutMallOrderSerial, OrgDetailKey)
            if resultStr = "Success" then
                response.write "취소신청 승인완료"
            end if
        case else
            response.write "제휴상태 : 취소요청 없음<br />"

            response.write "품절등록중<br />"
            resultStr = SetStockOutOne_gmarket1010(sellsite, OutMallOrderSerial, OrgDetailKey)
            if resultStr <> "Success" then
                response.write "품절등록 실패"
                dbget.close() : response.end
            end if

            response.write "취소완료 중<br />"
            resultStr = GetCSOrderCancelRequestConfirmOne_gmarket1010(sellsite, OutMallOrderSerial, OrgDetailKey)
            if resultStr <> "Success" then
                response.write "취소완료 실패"
                dbget.close() : response.end
            end if

            if resultStr = "Success" then
                response.write "주문취소 접수완료"
            end if
    end select

    if resultStr = "Success" then
		strSql = "Update [db_temp].[dbo].[tbl_mibeasong_list] "
		strSql = strSql & "	Set isSendAPI='Y'"
		strSql = strSql & "	where detailidx='"&detailidx&"'"&VBCRLF
		dbget.Execute strSql,AssignedRow
    end if

end function

function GetCSStockoutCancelOne_WMP(sellsite, mode, detailidx, orderserial)
    dim strSql
    dim OutMallOrderSerial, OrgDetailKey, resultStr, Comment, AssignedRow

	strSql = " select top 1 T.OutMallOrderSerial, T.OrgDetailKey, d.idx as detailidx "
	strSql = strSql + " from "
	strSql = strSql + " 	[db_temp].[dbo].[tbl_xSite_TMPOrder] T "
	strSql = strSql + " 	join [db_order].[dbo].[tbl_order_detail] d "
	strSql = strSql + " 	on "
	strSql = strSql + " 		1 = 1 "
	strSql = strSql + " 		and T.OrderSerial = d.orderserial "
	strSql = strSql + " 		and T.matchItemID = d.itemid "
	strSql = strSql + " 		and T.matchitemoption = d.itemoption "
	strSql = strSql + " where "
	strSql = strSql + " 	1 = 1 "
	strSql = strSql + " 	and sellsite = '" & sellsite & "' "
	strSql = strSql + " 	and T.OrderSerial = '" & orderserial & "' "
	strSql = strSql + " 	and d.idx = " & detailidx
	''response.write strSql
    rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

	OutMallOrderSerial = ""
	If Not rsget.Eof Then
		OutMallOrderSerial 		= rsget("OutMallOrderSerial")
        OrgDetailKey 			= rsget("OrgDetailKey")
	end if
    rsget.Close

	if (OutMallOrderSerial = "") then
		response.write "<font color='red'>결품등록(주문취소) 실패!!</font> : 에러 : 내역없음<br />"
		exit function
	end if

    resultStr = SetCSOrderCancelStateOne_WMP(sellsite, OutMallOrderSerial, OrgDetailKey)

    if resultStr = "Success" then
		strSql = "Update [db_temp].[dbo].[tbl_mibeasong_list] "
		strSql = strSql & "	Set isSendAPI='Y'"
		strSql = strSql & "	where detailidx='"&detailidx&"'"&VBCRLF
		dbget.Execute strSql,AssignedRow

        response.write "주문취소 접수완료"
    end if

end function

function GetCSStockoutCancelOne_WMPfashion(sellsite, mode, detailidx, orderserial)
    dim strSql
    dim OutMallOrderSerial, OrgDetailKey, resultStr, Comment, AssignedRow

	strSql = " select top 1 T.OutMallOrderSerial, T.OrgDetailKey, d.idx as detailidx "
	strSql = strSql + " from "
	strSql = strSql + " 	[db_temp].[dbo].[tbl_xSite_TMPOrder] T "
	strSql = strSql + " 	join [db_order].[dbo].[tbl_order_detail] d "
	strSql = strSql + " 	on "
	strSql = strSql + " 		1 = 1 "
	strSql = strSql + " 		and T.OrderSerial = d.orderserial "
	strSql = strSql + " 		and T.matchItemID = d.itemid "
	strSql = strSql + " 		and T.matchitemoption = d.itemoption "
	strSql = strSql + " where "
	strSql = strSql + " 	1 = 1 "
	strSql = strSql + " 	and sellsite = '" & sellsite & "' "
	strSql = strSql + " 	and T.OrderSerial = '" & orderserial & "' "
	strSql = strSql + " 	and d.idx = " & detailidx
	''response.write strSql
    rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

	OutMallOrderSerial = ""
	If Not rsget.Eof Then
		OutMallOrderSerial 		= rsget("OutMallOrderSerial")
        OrgDetailKey 			= rsget("OrgDetailKey")
	end if
    rsget.Close

	if (OutMallOrderSerial = "") then
		response.write "<font color='red'>결품등록(주문취소) 실패!!</font> : 에러 : 내역없음<br />"
		exit function
	end if

    resultStr = SetCSOrderCancelStateOne_WMPfashion(sellsite, OutMallOrderSerial, OrgDetailKey)

    if resultStr = "Success" then
		strSql = "Update [db_temp].[dbo].[tbl_mibeasong_list] "
		strSql = strSql & "	Set isSendAPI='Y'"
		strSql = strSql & "	where detailidx='"&detailidx&"'"&VBCRLF
		dbget.Execute strSql,AssignedRow

        response.write "주문취소 접수완료"
    end if

end function

CONST gmarketTicket = "0A2799EE6A1B65CC78DA96AA52C7546B2181855E48A0A31EDD4F3A77C3C61015856FE3DE5D7828B129A31AAD5914D7060556616D3AB7F2A84008A600C89F5953A0362065429900D0EB25CEBEA0E1CAF9E784FBC4F36E86608F2CF44B40113ADF"

'// 취소정보조회
function GetCSOrderCancelStateOne_gmarket1010(sellsite, OutMallOrderSerial, OrgDetailKey)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj
	dim i
	dim startdate, enddate
	dim CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql, IsDelete, ResultStr, Comment, currState
    dim OutMallOrderSerialCS, OrgDetailKeyCS

	xmlURL = "https://tpl.gmarket.co.kr/v1/OrderCancelService.asmx"

	strRst = ""
	strRst = strRst + "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst + "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst + " <soap:Header>"
	strRst = strRst + "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "			<encTicket>" & gmarketTicket & "</encTicket>"
	strRst = strRst + "		</EncTicket>"
	strRst = strRst + "	</soap:Header>"
	strRst = strRst + "	<soap:Body>"
	strRst = strRst + "		<RequestOrderCancel xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "			<RequestOrderCancel PackNo=""" & OutMallOrderSerial & """ />"
	strRst = strRst + "		</RequestOrderCancel>"
	strRst = strRst + "	</soap:Body>"
	strRst = strRst + "</soap:Envelope>"
	''response.write strRst

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(strRst)
	objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/RequestOrderCancel"
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(Replace(objXML.responseText,"soap:",""))

	currState = "BXXX"
	If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "RequestOrderCancelResponse" Then
		set objArr = xmlDOM.getElementsByTagName("RequestOrderCancelResultT")

		for i = 0 to objArr.length - 1
			set obj = objArr.item(i)

			OutMallOrderSerialCS = obj.getAttribute("PackNo")
			OrgDetailKeyCS = obj.getAttribute("ContrNo")

            if (OutMallOrderSerialCS = OutMallOrderSerial) and (OrgDetailKeyCS = OrgDetailKey) then
                '// ClaimReady : 취소신청
                '// ClaimDone : 취소완료
            	'// ClaimReject : 취소철회
            	'// ClaimDoneG : G마켓 직권 환불 건만 조회 (취소 완료 건 중 고객센터에서 환불 처리 한 Case)
                ResultStr = obj.getAttribute("ClaimStatus")

                select case ResultStr
                    case "ClaimReady"
                        currState = "B001"
                    case "ClaimDone"
                        currState = "B007"
                    case "ClaimReject"
                        currState = "B008"
                    case "ClaimDoneG"
                        currState = "B007"
                    case else
                        currState = "BERR : " & ResultStr
                end select
            end if
        next
    end if

    GetCSOrderCancelStateOne_gmarket1010 = currState
End Function

'// 취소요청 : 사용불가
'// 취소요청 대신 품절등록+취소완료 API 사용할것
function GetCSOrderCancelOne_gmarket1010(sellsite, OutMallOrderSerial, OrgDetailKey)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj
	dim i
	dim startdate, enddate
	dim CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql, IsDelete, ResultStr, Comment

	xmlURL = "https://tpl.gmarket.co.kr/v1/OrderCancelService.asmx"
    response.write xmlURL & "<br />"

	strRst = ""
	strRst = strRst + "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst + "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst + " <soap:Header>"
	strRst = strRst + "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "			<encTicket>" & gmarketTicket & "</encTicket>"
	strRst = strRst + "		</EncTicket>"
	strRst = strRst + "	</soap:Header>"
	strRst = strRst + "	<soap:Body>"
	strRst = strRst + "		<AddOrderCancel xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "			<AddOrderCancel PackNo=""" & OutMallOrderSerial & """ ContrNo=""" & OrgDetailKey & """ ClaimType=""ETC"" ClaimReason=""OUT OF STOCK"" />"
	strRst = strRst + "		</AddOrderCancel>"
	strRst = strRst + "	</soap:Body>"
	strRst = strRst + "</soap:Envelope>"
	response.write strRst

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(strRst)
	objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/AddOrderCancel"
	objXML.send(strRst)
    response.write LenB(strRst) & "<br />"

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(Replace(objXML.responseText,"soap:",""))

	GetCSOrderCancelOne_gmarket1010 = ""
	If xmlDOM.selectSingleNode("Envelope/Body/AddOrderCancelResponse").firstChild.nodeName = "AddOrderCancelResult" Then
		set obj = xmlDOM.getElementsByTagName("AddOrderCancelResult").item(0)

        ResultStr = obj.getAttribute("Result")
        Comment = obj.getAttribute("Comment")

        GetCSOrderCancelOne_gmarket1010 = ResultStr
        if (ResultStr <> "Success") then
            response.write "ERR : " & Comment
        end if
    end if
End Function

function SetStockOutOne_gmarket1010(sellsite, OutMallOrderSerial, OrgDetailKey)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj
	dim i
	dim startdate, enddate
	dim CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql, IsDelete, ResultStr, Comment

	xmlURL = "https://tpl.gmarket.co.kr/v1/ShippingService.asmx"
    response.write xmlURL & "<br />"

	strRst = ""
	strRst = strRst + "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst + "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst + " <soap:Header>"
	strRst = strRst + "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "			<encTicket>" & gmarketTicket & "</encTicket>"
	strRst = strRst + "		</EncTicket>"
	strRst = strRst + "	</soap:Header>"
	strRst = strRst + "	<soap:Body>"
	strRst = strRst + "		<AddShippingReject xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "			<AddShippingReject PackNo=""" & OutMallOrderSerial & """ ContrNo=""" & OrgDetailKey & """ RejectReaon=""OutofStock"" />"
	strRst = strRst + "		</AddShippingReject>"
	strRst = strRst + "	</soap:Body>"
	strRst = strRst + "</soap:Envelope>"
	''response.write strRst

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(strRst)
	objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/AddShippingReject"
	objXML.send(strRst)
    ''response.write LenB(strRst) & "<br />"

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(Replace(objXML.responseText,"soap:",""))

	SetStockOutOne_gmarket1010 = ""
	If xmlDOM.selectSingleNode("Envelope/Body/AddShippingRejectResponse").firstChild.nodeName = "AddShippingRejectResult" Then
		set obj = xmlDOM.getElementsByTagName("AddShippingRejectResult").item(0)

        ResultStr = obj.getAttribute("Result")
        Comment = obj.getAttribute("Comment")

        SetStockOutOne_gmarket1010 = ResultStr
        if (ResultStr <> "Success") then
            response.write "ERR : " & Comment
        end if
    end if
End Function

'// 취소요청 완료처리
function GetCSOrderCancelRequestConfirmOne_gmarket1010(sellsite, OutMallOrderSerial, OrgDetailKey)
	dim xmlURL, strRst
	dim objXML, xmlDOM, objArr, obj
	dim i
	dim startdate, enddate
	dim CSDetailKey, divcd, gubunname, OutMallRegDate, itemno
	dim iAssignedRow, iInputCnt
	dim strSql, IsDelete, ResultStr, Comment

	xmlURL = "https://tpl.gmarket.co.kr/v1/OrderCancelService.asmx"

	strRst = ""
	strRst = strRst + "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst + "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst + " <soap:Header>"
	strRst = strRst + "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "			<encTicket>" & gmarketTicket & "</encTicket>"
	strRst = strRst + "		</EncTicket>"
	strRst = strRst + "	</soap:Header>"
	strRst = strRst + "	<soap:Body>"
	strRst = strRst + "		<ConfirmOrderCancel xmlns=""http://tpl.gmarket.co.kr/"">"
	strRst = strRst + "			<ConfirmOrderCancel ContrNo=""" & OrgDetailKey & """ IsClaimConfirm=""true"" />"
	strRst = strRst + "		</ConfirmOrderCancel>"
	strRst = strRst + "	</soap:Body>"
	strRst = strRst + "</soap:Envelope>"
	''response.write strRst

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(strRst)
	objXML.setRequestHeader "SOAPAction", "http://tpl.gmarket.co.kr/ConfirmOrderCancel"
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(Replace(objXML.responseText,"soap:",""))

	GetCSOrderCancelRequestConfirmOne_gmarket1010 = ""
	If xmlDOM.selectSingleNode("Envelope/Body/ConfirmOrderCancelResponse").firstChild.nodeName = "ConfirmOrderCancelResult" Then
		set obj = xmlDOM.getElementsByTagName("ConfirmOrderCancelResult").item(0)

        ResultStr = obj.getAttribute("Result")
        Comment = obj.getAttribute("Comment")

        GetCSOrderCancelRequestConfirmOne_gmarket1010 = ResultStr
        if (ResultStr <> "Success") then
            response.write "ERR : " & Comment
        end if
    end if
End Function

function SetCSOrderCancelStateOne_WMP(sellsite, OutMallOrderSerial, OrgDetailKey)
	dim xmlURL, strRst, objXML, xmlDOM, strObj
	dim startdate, enddate
	dim retCode, retMsg, items, item, product, productOption, divcd, gubunname
	dim CSDetailKey, dlvTypeGbcd, itemno, OutMallRegDate
	dim i, j, k
	dim strSql, iAssignedRow, successCount

	xmlURL = "http://xapi.10x10.co.kr:8080/Wemake/Orders/ordercancel?OrderOptionNo=" & OrgDetailKey

	strRst = ""

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	Set strObj = JSON.parse(objXML.responseText)
	retMsg = strObj.message

	successCount = 0

    SetCSOrderCancelStateOne_WMP = ""
    if (retMsg = "성공") then
        SetCSOrderCancelStateOne_WMP = "Success"
    else
		if IsAutoScript then
			response.write "ERROR : " & retMsg
		else
			response.write "ERROR : " & objXML.Status
			response.write "<script>alert('ERROR : " & retMsg & "');</script>"
		end if
        dbget.close() : response.end
    end if
end function

function SetCSOrderCancelStateOne_WMPfashion(sellsite, OutMallOrderSerial, OrgDetailKey)
	dim xmlURL, strRst, objXML, xmlDOM, strObj
	dim startdate, enddate
	dim retCode, retMsg, items, item, product, productOption, divcd, gubunname
	dim CSDetailKey, dlvTypeGbcd, itemno, OutMallRegDate
	dim i, j, k
	dim strSql, iAssignedRow, successCount

	xmlURL = "http://110.93.128.100:8090/fwmp/Orders/ordercancel?OrderOptionNo=" & OrgDetailKey

	strRst = ""

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	Set strObj = JSON.parse(objXML.responseText)
	retMsg = strObj.message

	successCount = 0

    SetCSOrderCancelStateOne_WMPfashion = ""
    if (retMsg = "성공") then
        SetCSOrderCancelStateOne_WMPfashion = "Success"
    else
		if IsAutoScript then
			response.write "ERROR : " & retMsg
		else
			response.write "ERROR : " & objXML.Status
			response.write "<script>alert('ERROR : " & retMsg & "');</script>"
		end if
        dbget.close() : response.end
    end if
end function
%>
