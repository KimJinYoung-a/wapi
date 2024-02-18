<!-- #include virtual="/outmall/jungsan/include/dvim_brix_crypto-js-master_VB.asp"-->
<%
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

function httpBuildQueryArray(key, items, count, out)
    dim result
    for i = 0 to count - 1
        if isarray(items(i)) then
            httpBuildQueryArray (key & "[" & i & "]"), items(i), UBound(items(i)), result
        else
            result = result & Server.URLEncode(key & "[" & i & "]") & "=" & Server.URLEncode(items(i))
        end if
        if i + 1 <> count then
            result = result & "&"
        end if
    next
    out = result
end function

function httpBuildQuery(keys, items, count, out)
    dim result
    for i = 0 to count - 1
        if isarray(items(i)) then
            httpBuildQueryArray keys(i), items(i), UBound(items(i)), result
        else
            result = result & Server.URLEncode(keys(i)) & "=" & Server.URLEncode(items(i))
        end if
        if i + 1 <> count then
            result = result & "&"
        end if
    next
    out = result
end function


Function generateHmac(url, method, querystring, apikey, seckey)
    dim signedDate, signature
    dim encString
    dim requestData

    '[timestamp, httpMethod, requestPath, queryString]
    signedDate = generateSignedDate()
    requestData = signedDate & method & url & querystring
    set encString = mac256(requestData, seckey)
    signature = ", signature=" & cstr(encString)

    generateHmac = "CEA algorithm=HmacSHA256, access-key=" & apikey & ", signed-date=" & signedDate & signature
End Function

'/******  MAC Function ******/
' CEA algorithm=HmacSHA256, access-key=0a3a0f34-7852-4ad8-9368-766290b8b1ab, signed-date=190201T042152Z, signature=737374df2de01ad31cc85c14c42faa0711e7d156d8350e57c911b20916d4ba77
'Input String|WordArray , Returns WordArray
Function mac256(ent, seckey)
    Dim encWA
    Set encWA = ConvertUtf8StrToWordArray(ent)
    Dim keyWA
    Set keyWA = ConvertUtf8StrToWordArray(seckey)
    Dim resWA
    Set resWA = CryptoJS.HmacSHA256(encWA, keyWA)
    Set mac256 = resWA
End Function

'Input (Utf8)String|WordArray Returns WordArray
Function ConvertUtf8StrToWordArray(data)
    If (typename(data) = "String") Then
        Set ConvertUtf8StrToWordArray = CryptoJS.enc.Utf8.parse(data)
    Elseif (typename(data) = "JScriptTypeInfo") Then
        On error resume next
        'Set ConvertUtf8StrToWordArray = CryptoJS.enc.Utf8.parse(data.toString(CryptoJS.enc.Utf8))
        Set ConvertUtf8StrToWordArray = CryptoJS.lib.WordArray.create().concat(data) 'Just assert that data is WordArray
        If Err.number>0 Then
            Set ConvertUtf8StrToWordArray = Nothing
        End if
        On error goto 0
    Else
        Set ConvertUtf8StrToWordArray = Nothing
    End if
End Function

Function generateSignedDate()
    Dim nowDateTime

    'ISO TIMEZONE
    nowDateTime = DateAdd("H", -9, now())

    generateSignedDate = ToIsoDate(nowDateTime) & "T" & ToIsoTime(nowDateTime) & "Z"

End Function

Function ToIsoDate(datetime)
    ToIsoDate = CStr(Mid(Year(datetime), 3, 2)) & StrN2(Month(datetime)) & StrN2(Day(datetime))
End Function

Function ToIsoTime(datetime)
    ToIsoTime = StrN2(Hour(datetime)) & StrN2(Minute(datetime)) & StrN2(Second(datetime))
End Function

Function StrN2(n)
    If Len(CStr(n)) < 2 Then StrN2 = "0" & n Else StrN2 = n
End Function

Function GetJungsan_ezwel(reqDate, hasNext, page, vTotalPage)
	Dim sellsite : sellsite = "ezwel"
	Dim objXML, xmlDOM, sqlStr
	Dim retCode, iMessage, AssignedRow, accountCnt
	Dim getParam, totalPage, Nodes, SubNodes
	Dim sndNm, taxYn, goodsNm, dccpnPrice, normalSalePrice, orderQty, orderAmt, realAccountAmt, marginRate, buyPrice, payAmt, accountAmt, orderDt, cancelDt
	Dim extOrderserial, extOrderserSeq, extMeachulDate, extJungsanDate, extItemNo, extItemName, extItemOptionName, dvlprice, extItemCost, extOrgOrderserial, extVatYN, extJungsanType, extTenMeachulPrice, extOwnCouponPrice, extTenCouponPrice, extTenJungsanPrice, extReducedPrice, extCommPrice, extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice, extitemid
	Dim tenitemid, tenitemoption, extitemoption, siteNo, p_extOrderserial, p_extOrderserSeq
	Dim tmpExtMeachulDate, tmpCancelDt
	GetJungsan_ezwel = False

	If page = "1" Then
		sqlStr = ""
		sqlStr = " DELETE FROM db_temp.dbo.tbl_xSite_JungsanTmp WHERE sellsite = '" & CStr(sellsite) & "' "
		dbget.execute sqlStr
	End If

	getParam = "cspCd=10040413&crtCd=8e5a6dbdd27efb49fc600c293884ef47"
	getParam = getParam & "&accountYm=" & reqDate
	getParam = getParam & "&page=" & page
	getParam = getParam & "&rowPerPage=1000"

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://api.ezwel.com/if/api/accountListAPI.ez?" & getParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
'  response.write objXML.ResponseText
'  response.end
				retCode		= xmlDOM.getElementsByTagName("resultCode").item(0).text
				iMessage	= xmlDOM.getElementsByTagName("resultMsg").item(0).text

				If retCode = "200" Then		'성공(200)
					accountCnt = xmlDOM.getElementsByTagName("accountCnt").item(0).text
					totalPage = xmlDOM.getElementsByTagName("totalPage").item(0).text
					vTotalPage = totalPage
					Set Nodes = xmlDOM.getElementsByTagName("accountInfo")
						For each SubNodes in Nodes
							tmpExtMeachulDate = ""
							tmpCancelDt	= ""
							extMeachulDate = ""
							cancelDt = ""

							extOrderserial		= SubNodes.getElementsByTagName("orderNum")(0).Text			'## 주문번호
							extOrderserSeq		= SubNodes.getElementsByTagName("orderGoodsNum")(0).Text	'## 주문순번
							sndNm				= SubNodes.getElementsByTagName("sndNm")(0).Text			'구매자명
							taxYn				= SubNodes.getElementsByTagName("taxYn")(0).Text			'면과세여부 | Y:과세, N:면세
							extitemid			= SubNodes.getElementsByTagName("goodsCd")(0).Text			'## 상품코드
							extItemName			= SubNodes.getElementsByTagName("goodsNm")(0).Text			'## 상품명
							tmpExtMeachulDate	= SubNodes.getElementsByTagName("dlvrFinishDt")(0).Text		'## 정산확정일자 | 년월일(YYYYMMDD)
							dccpnPrice			= SubNodes.getElementsByTagName("dccpnPrice")(0).Text		'할인금액
							normalSalePrice		= SubNodes.getElementsByTagName("normalSalePrice")(0).Text	'정상(시중)가
							extItemCost			= SubNodes.getElementsByTagName("salePrice")(0).Text		'판매단가
							extItemNo			= SubNodes.getElementsByTagName("orderQty")(0).Text			'## 주문수량
							orderAmt			= SubNodes.getElementsByTagName("orderAmt")(0).Text			'원판매금액
							realAccountAmt		= SubNodes.getElementsByTagName("realAccountAmt")(0).Text	'실정산금액
							marginRate			= SubNodes.getElementsByTagName("marginRate")(0).Text		'수수료율
							buyPrice			= SubNodes.getElementsByTagName("buyPrice")(0).Text			'공급단가
							payAmt				= SubNodes.getElementsByTagName("payAmt")(0).Text			'실판매금액
							accountAmt			= SubNodes.getElementsByTagName("accountAmt")(0).Text		'매입금액
							dvlprice			= SubNodes.getElementsByTagName("dlvrAmt")(0).Text			'배송비
							orderDt				= SubNodes.getElementsByTagName("orderDt")(0).Text			'주문일 | 년월일시분초(YYYYMMDDhh24miss
							tmpCancelDt			= SubNodes.getElementsByTagName("cancelDt")(0).Text			'취소일 | 년월일시분초(YYYYMMDDhh24miss

							tmpExtMeachulDate = LEFT(tmpExtMeachulDate, 8)
							extMeachulDate = DateSerial(CInt(Mid(tmpExtMeachulDate, 1, 4)), CInt(Mid(tmpExtMeachulDate, 5, 2)), Mid(tmpExtMeachulDate, 7, 2))

							extJungsanDate = ""
							If (tmpCancelDt <> "") and (extItemNo < 0) then ''취소일 수량이 마이너스인거만
								tmpCancelDt = LEFT(tmpCancelDt, 8)
								cancelDt = DateSerial(CInt(Mid(tmpcancelDt, 1, 4)), CInt(Mid(tmpcancelDt, 5, 2)), Mid(tmpcancelDt, 7, 2))

								extMeachulDate = cancelDt
								extOrderserSeq = extOrderserSeq & "-1"
							End If
							extOrgOrderserial			= ""
							extVatYN					= "Y"
							extJungsanType				= "C"
							extTenMeachulPrice			= extItemCost ''쿠폰금액을 넣지 말자. 않맞음 2019/05/08
							extOwnCouponPrice			= Clng((dccpnPrice / extItemNo) * 100) / 100
							extTenCouponPrice			= extOwnCouponPrice * -1
							extTenJungsanPrice			= CLNG((realAccountAmt - dvlprice) / extItemNo * 100) / 100
							extReducedPrice				= CLNG(extTenMeachulPrice)
							extCommPrice				= extTenMeachulPrice - extTenJungsanPrice
							extCommSupplyPrice			= extCommPrice
							extCommSupplyVatPrice		= 0
							extTenMeachulSupplyPrice	= extTenMeachulPrice
							extTenMeachulSupplyVatPrice	= 0

							'1. 상품 정산 입력
							sqlStr = ""
							sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_xSite_JungsanTmp "
							sqlStr = sqlStr & " (sellsite, extOrderserial, extOrderserSeq "
							sqlStr = sqlStr & ", extOrgOrderserial, extItemNo, extItemCost "
							sqlStr = sqlStr & ", extReducedPrice, extOwnCouponPrice, extTenCouponPrice "
							sqlStr = sqlStr & ", extJungsanType, extCommPrice, extTenMeachulPrice "
							sqlStr = sqlStr & ", extTenJungsanPrice, extMeachulDate, extJungsanDate "
							sqlStr = sqlStr & ", extItemName, extItemOptionName, extVatYN "
							sqlStr = sqlStr & ", extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice "
							sqlStr = sqlStr & ", itemid, itemoption "
							sqlStr = sqlStr & ", extitemid, extitemoption,siteNo "
							sqlStr = sqlStr & " ) "
							sqlStr = sqlStr & " VALUES('" & CStr(sellsite) & "', '" & CStr(extOrderserial) & "', '" & CStr(extOrderserSeq) & "'"
							sqlStr = sqlStr & ", '" & CStr(extOrgOrderserial) & "', '" & CStr(extItemNo) & "', '" & CStr(extItemCost) & "'"
							sqlStr = sqlStr & ", '" & CStr(extReducedPrice) & "', '" & CStr(extOwnCouponPrice) & "', '" & CStr(extTenCouponPrice) & "'"
							sqlStr = sqlStr & ", '" & CStr(extJungsanType) & "', '" & CStr(extCommPrice) & "', '" & CStr(extTenMeachulPrice) & "'"
							sqlStr = sqlStr & ", '" & CStr(extTenJungsanPrice) & "', '" & CStr(extMeachulDate) & "', '" & CStr(extJungsanDate) & "'"
							sqlStr = sqlStr & ", '" & CStr(extItemName) & "', convert(varchar(128),'" & CStr(extItemOptionName) & "'), '" & CStr(extVatYN) & "'"
							sqlStr = sqlStr & ", '" & CStr(extCommSupplyPrice) & "', '" & CStr(extCommSupplyVatPrice) & "', '" & CStr(extTenMeachulSupplyPrice) & "', '" & CStr(extTenMeachulSupplyVatPrice) & "'"
							If (tenitemid <> "") Then
								sqlStr = sqlStr & ", '" & CStr(tenitemid) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (tenitemoption<>"") Then
								sqlStr = sqlStr & ", '" & CStr(tenitemoption) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (extitemid<>"") Then
								sqlStr = sqlStr & ", '" & CStr(extitemid) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (extitemoption<>"") Then
								sqlStr = sqlStr & ", '" & CStr(extitemoption) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (siteNo<>"") Then
								sqlStr = sqlStr & ", '" & CStr(siteNo) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							sqlStr = sqlStr & ") "

							If (extItemNo <> 0) Then
								p_extOrderserial = extOrderserial
								p_extOrderserSeq = extOrderserSeq

								on Error Resume Next
								dbget.execute sqlStr
								If Err Then
									rw "page : " & page
									rw sqlStr
									on Error Goto 0
									response.end
								End If
							End If

							'2. 만약 배송비가 0보다 크면 배송비 입력
							If (dvlprice <> 0) Then  '배송비가 한줄로 있다.
								extJungsanType = "D"
								extOrderserSeq = extOrderserSeq+"-D"
								extOwnCouponPrice		= 0
								extTenCouponPrice		= 0
								extCommPrice			= 0

								If (dvlprice < 0) Then
									extItemNo = -1
								Else
									extItemNo = 1
								End If

								extItemCost				= dvlprice/extItemNo
								extTenJungsanPrice		= dvlprice/extItemNo


								extReducedPrice			= dvlprice/extItemNo-extOwnCouponPrice
								extTenMeachulPrice      = dvlprice/extItemNo-extOwnCouponPrice

								extCommSupplyPrice		= 0
								extCommSupplyVatPrice	= 0

								extTenMeachulSupplyPrice	= 0
								extTenMeachulSupplyVatPrice	= 0
							Else
								extItemNo = 0
							End If

							If (extItemNo <> 0) Then
								sqlStr = ""
								sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_xSite_JungsanTmp "
								sqlStr = sqlStr & " (sellsite, extOrderserial, extOrderserSeq "
								sqlStr = sqlStr & ", extOrgOrderserial, extItemNo, extItemCost "
								sqlStr = sqlStr & ", extReducedPrice, extOwnCouponPrice, extTenCouponPrice "
								sqlStr = sqlStr & ", extJungsanType, extCommPrice, extTenMeachulPrice "
								sqlStr = sqlStr & ", extTenJungsanPrice, extMeachulDate, extJungsanDate "
								sqlStr = sqlStr & ", extItemName, extItemOptionName, extVatYN "
								sqlStr = sqlStr & ", extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice "
								sqlStr = sqlStr & ", itemid, itemoption "
								sqlStr = sqlStr & ", extitemid, extitemoption,siteNo"
								sqlStr = sqlStr & " ) "
								sqlStr = sqlStr & " values('" & CStr(sellsite) & "', '" & CStr(extOrderserial) & "', '" & CStr(extOrderserSeq) & "'"
								sqlStr = sqlStr & ", '" & CStr(extOrgOrderserial) & "', '" & CStr(extItemNo) & "', '" & CStr(extItemCost) & "'"
								sqlStr = sqlStr & ", '" & CStr(extReducedPrice) & "', '" & CStr(extOwnCouponPrice) & "', '" & CStr(extTenCouponPrice) & "'"
								sqlStr = sqlStr & ", '" & CStr(extJungsanType) & "', '" & CStr(extCommPrice) & "', '" & CStr(extTenMeachulPrice) & "'"
								sqlStr = sqlStr & ", '" & CStr(extTenJungsanPrice) & "', '" & CStr(extMeachulDate) & "', '" & CStr(extJungsanDate) & "'"
								sqlStr = sqlStr & ", '" & CStr(extItemName) & "', convert(varchar(128),'" & CStr(extItemOptionName) & "'), '" & CStr(extVatYN) & "'"
								sqlStr = sqlStr & ", '" & CStr(extCommSupplyPrice) & "', '" & CStr(extCommSupplyVatPrice) & "', '" & CStr(extTenMeachulSupplyPrice) & "', '" & CStr(extTenMeachulSupplyVatPrice) & "'"
								If (tenitemid <> "") Then
									sqlStr = sqlStr & ", '" & CStr(tenitemid) & "'"
								Else
									sqlStr = sqlStr & ", NULL"
								End If
								If (tenitemoption <> "") Then
									sqlStr = sqlStr & ", '" & CStr(tenitemoption) & "'"
								Else
									sqlStr = sqlStr & ", NULL"
								End If
								If (extitemid <> "") Then
									sqlStr = sqlStr & ", '" & CStr(extitemid) & "'"
								Else
									sqlStr = sqlStr & ", NULL"
								End If
								If (extitemoption <> "") Then
									sqlStr = sqlStr & ", '" & CStr(extitemoption) & "'"
								Else
									sqlStr = sqlStr & ", NULL"
								End If
								If (siteNo <> "") Then
									sqlStr = sqlStr & ", '" & CStr(siteNo) & "'"
								Else
									sqlStr = sqlStr & ", NULL"
								End If
								sqlStr = sqlStr & ") "
								dbget.execute sqlStr
							End If
						Next
					Set Nodes = nothing

					If accountCnt > 0 Then
						If CSTR(page) = CSTR(totalPage) Then
							sqlStr = " EXEC db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_ezwel] "
							dbget.Execute sqlStr
							hasnext = "N"
						Else
							hasnext = "Y"
						End If

						page = page + 1
					Else
						response.write reqDate & " JungsanData Not Exists"
						response.end
					End If
				Else
					rw iMessage
					response.end
				End If
			Set xmlDOM = nothing
		Else
			rw "오류......"
			response.end
		End If
	Set objXML = nothing
End Function

'19000091442810
Function GetJungsan_coupang(reqDate, hasnext, nextToken)
	Dim sellsite : sellsite = "coupang"
	Dim url, path, method, params, req
	Dim access_key, secret_key, vendorId
	Dim authorization

	Dim objXML, xmlDOM, sqlStr, iRbody, strObj
	Dim retCode, iMessage
	Dim datalist, i, j, itemlist

	Dim extOrderserial, extMeachulDate, extJungsanDate, extVatYN, extitemid, extitemoption, extJungsanType, extItemNo, extOwnCouponPrice, extTenCouponPrice
	Dim extOrderserSeq, tenitemid, tenitemoption, extOrgOrderserial, extCommPrice, extTenMeachulPrice, extReducedPrice, extTenJungsanPrice, extCommSupplyPrice
	Dim saleType, siteNo, extItemCost, extItemName, extItemOptionName, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice, settlementAmount

	Dim extReItemNo, extChgItemno, validitemno
	extItemNo = 0
	extReItemNo = 0
	extChgItemno = 0
	validitemno = 0

	access_key = "0af06fb7-3deb-4ac3-9a84-6d409a26d831"
	secret_key = "5474f1108ac5631e5977d4a6b7a6387426533582"
	vendorId = "A00039305"
	path = "/v2/providers/openapi/apis/api/v1/revenue-history"
	params = "vendorId="&vendorId&"&recognitionDateFrom="&reqDate&"&recognitionDateTo="&reqDate&"&token="&nextToken&"&maxPerPage=50"
	url = "https://api-gateway.coupang.com" & path
	method = "GET"
	authorization = generateHmac(path, method, params, access_key, secret_key)

	sqlStr = ""
	sqlStr = " DELETE FROM db_temp.dbo.tbl_xSite_JungsanTmp WHERE sellsite = '" & CStr(sellsite) & "' "
	dbget.execute sqlStr

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open method, url & "?" & params, false
		objXML.setRequestHeader "Authorization", authorization
		objXML.setRequestHeader "X-Requested-By", vendorId
		objXML.send()
 		If objXML.Status = "200" OR objXML.Status = "201" Then
 			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				retCode		= strObj.code			'서버 응답 코드
				iMessage	= strObj.message		'detail info
				If strObj.hasNext = "True" Then		'다음페이지에 데이터 존재 여부
					hasnext = "Y"
				Else
					hasnext = "N"
				End If
				nextToken	= strObj.nextToken	'다음 페이지를 조회하기위한 토큰 값
				If retCode = "200" Then
					Set datalist = strObj.data		'결과리스트 | 결과가 없을 때는 빈 리스트가 리턴
						For i=0 to datalist.length-1
							extOrderserial	= datalist.get(i).orderId				'주문번호
							saleType		= datalist.get(i).saleType				'항목구분 | SALE : 주문 건, REFUND : 반품 건
							extMeachulDate	= datalist.get(i).recognitionDate		'매출인식일 | 형식 : YYYY-MM-dd  배송완료 + 7day' 또는 '구매확정'
							Set itemlist = datalist.get(i).items		'주문상품별 정산금액 상세
								For j=0 to itemlist.length-1
									extJungsanType = "C"
									If itemlist.get(j).taxType = "TAX" Then
										extVatYN = "Y"
									Else
										extVatYN = "N"
									End If
									settlementAmount	= itemlist.get(j).settlementAmount	'정산금액 | 정산금액 = 매출금액 - (서비스이용료 + 서비스이용VAT)
									extitemid			= itemlist.get(j).productId			'노출상품 ID | 머지(결합)와 분리 등으로 언제든 변경될 수 있는 ID로서, 정산을 대사할때 key로 활용 불가
									extitemoption		= itemlist.get(j).vendorItemId		'옵션 ID | 쿠팡의 가장 작은 상품 단위. (상품 수정, 반품 등에서 활용 되는 단위) 변경되지 않으며, 가장 작은 단위이기 때문에 key로 사용
									extItemCost			= itemlist.get(j).salePrice			'총 판매가 | 수량이 반영된 총 판매가
									extItemNo			= itemlist.get(j).quantity			'수량

									If extItemNo <> 0 Then
										extItemCost			= extItemCost / extItemNo
										'extOwnCouponPrice	= 0
										extOwnCouponPrice	= CLNG(itemlist.get(j).coupangDiscountCoupon/extItemNo*100)/100	'쿠팡지원할인금액
										extTenCouponPrice	= CLNG(itemlist.get(j).sellerDiscountCoupon/extItemNo*100)/100	'판매자할인쿠폰
										extCommPrice		= CLNG( (itemlist.get(j).serviceFee + itemlist.get(j).serviceFeeVat) /extItemNo*100)/100
									Else
										extOwnCouponPrice	= 0
										extTenCouponPrice	= 0
										extCommPrice		= itemlist.get(j).serviceFee + itemlist.get(j).serviceFeeVat
									End If
									extItemName 		= ""
									extItemOptionName	= ""

									extTenMeachulPrice		= extItemCost - extTenCouponPrice - extOwnCouponPrice
									extReducedPrice			= CLNG(extTenMeachulPrice)
									extTenJungsanPrice      = extTenMeachulPrice-extCommPrice

									extCommSupplyPrice		= extCommPrice
									extCommSupplyVatPrice	= 0

									extTenMeachulSupplyPrice	= extTenMeachulPrice
									extTenMeachulSupplyVatPrice	= 0

									extJungsanDate = ""
									extOrderserSeq = extitemoption
									tenitemid =""
									tenitemoption =""
									If saleType = "SALE" Then
										extOrgOrderserial	= ""
									Else
										extItemNo			= extItemNo * -1
										extOrgOrderserial	= extOrderserial
										extOrderserSeq = extOrderserSeq&"-1"
									End If

									''배송비가 2개로 나눠오는 케이스가 있어서 seq 중복으로 하나만 들어감;; 2021-09-01 김진영
									If DATE() <= "2022-12-01" Then
										sqlStr = ""
										sqlStr = sqlStr & " SELECT COUNT(*) as cnt FROM db_jungsan.dbo.tbl_xSite_Jungsandata WHERE sellsite = 'coupang' "
										sqlStr = sqlStr & " and extJungsanType = 'C' "
										sqlStr = sqlStr & " and extOrderserial = '"& extOrderserial &"' "
										sqlStr = sqlStr & " and extMeachulDate < '"& extMeachulDate &"' "
										sqlStr = sqlStr & " and extOrderserSeq = '"& extOrderserSeq &"' "
										rsget.CursorLocation = adUseClient
										rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
										If not rsget.Eof Then
											If rsget("cnt") > 0 Then
												extOrderserSeq = extOrderserSeq & "-2"
											End If
										End If
										rsget.close
									End If

									'1. 상품 정산 입력
									sqlStr = ""
									sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_xSite_JungsanTmp "
									sqlStr = sqlStr & " (sellsite, extOrderserial, extOrderserSeq "
									sqlStr = sqlStr & ", extOrgOrderserial, extItemNo, extItemCost "
									sqlStr = sqlStr & ", extReducedPrice, extOwnCouponPrice, extTenCouponPrice "
									sqlStr = sqlStr & ", extJungsanType, extCommPrice, extTenMeachulPrice "
									sqlStr = sqlStr & ", extTenJungsanPrice, extMeachulDate, extJungsanDate "
									sqlStr = sqlStr & ", extItemName, extItemOptionName, extVatYN "
									sqlStr = sqlStr & ", extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice "
									sqlStr = sqlStr & ", itemid, itemoption "
									sqlStr = sqlStr & ", extitemid, extitemoption,siteNo "
									sqlStr = sqlStr & " ) "
									sqlStr = sqlStr & " VALUES('" & CStr(sellsite) & "', '" & CStr(extOrderserial) & "', '" & CStr(extOrderserSeq) & "'"
									sqlStr = sqlStr & ", '" & CStr(extOrgOrderserial) & "', '" & CStr(extItemNo) & "', '" & CStr(extItemCost) & "'"
									sqlStr = sqlStr & ", '" & CStr(extReducedPrice) & "', '" & CStr(extOwnCouponPrice) & "', '" & CStr(extTenCouponPrice) & "'"
									sqlStr = sqlStr & ", '" & CStr(extJungsanType) & "', '" & CStr(extCommPrice) & "', '" & CStr(extTenMeachulPrice) & "'"
									sqlStr = sqlStr & ", '" & CStr(extTenJungsanPrice) & "', '" & CStr(extMeachulDate) & "', '" & CStr(extJungsanDate) & "'"
									sqlStr = sqlStr & ", '" & CStr(extItemName) & "', convert(varchar(128),'" & CStr(extItemOptionName) & "'), '" & CStr(extVatYN) & "'"
									sqlStr = sqlStr & ", '" & CStr(extCommSupplyPrice) & "', '" & CStr(extCommSupplyVatPrice) & "', '" & CStr(extTenMeachulSupplyPrice) & "', '" & CStr(extTenMeachulSupplyVatPrice) & "'"
									If (tenitemid <> "") Then
										sqlStr = sqlStr & ", '" & CStr(tenitemid) & "'"
									Else
										sqlStr = sqlStr & ", NULL"
									End If
									If (tenitemoption<>"") Then
										sqlStr = sqlStr & ", '" & CStr(tenitemoption) & "'"
									Else
										sqlStr = sqlStr & ", NULL"
									End If
									If (extitemid<>"") Then
										sqlStr = sqlStr & ", '" & CStr(extitemid) & "'"
									Else
										sqlStr = sqlStr & ", NULL"
									End If
									If (extitemoption<>"") Then
										sqlStr = sqlStr & ", '" & CStr(extitemoption) & "'"
									Else
										sqlStr = sqlStr & ", NULL"
									End If
									If (siteNo<>"") Then
										sqlStr = sqlStr & ", '" & CStr(siteNo) & "'"
									Else
										sqlStr = sqlStr & ", NULL"
									End If
									sqlStr = sqlStr & ") "

									If (extItemNo <> 0) Then
										on Error Resume Next
										dbget.execute sqlStr
										'rw sqlStr
										If Err Then
											rw sqlStr
											on Error Goto 0
											response.end
										End If
									End If
								Next
							Set itemlist = nothing

							'배송비 시작
							extJungsanType = "D"
							extOrderserSeq = "D"

							sqlStr = ""
							sqlStr = sqlStr & " SELECT COUNT(*) as cnt FROM db_jungsan.dbo.tbl_xSite_Jungsandata WHERE sellsite = 'coupang' "
							sqlStr = sqlStr & " and extJungsanType = 'D' "
							sqlStr = sqlStr & " and extOrderserial = '"& extOrderserial &"' "
							sqlStr = sqlStr & " and extOrderserSeq = 'D' "
							rsget.CursorLocation = adUseClient
							rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
							If not rsget.Eof Then
								If rsget("cnt") > 0 Then
									extOrderserSeq = extOrderserSeq & "-1"
								End If
							End If
							rsget.close

							extItemCost = datalist.get(i).deliveryFee.amount		'총 배송비 | 총 배송비 : 기본배송비 + 도서산간배송비

							extItemName 		= ""
							extItemOptionName	= ""
							extOwnCouponPrice=0
							extTenCouponPrice=0
							settlementAmount = datalist.get(i).deliveryFee.settlementAmount	'배송비 정산 대상액 | 정산 대상액 : 총 배송비 - 배송비 수수료 - 배송비 부가가치세

							If extItemNo <> 0 Then
								If (extItemCost <> 0) Then
									If (extItemCost < 0) Then
										extItemCost = extItemCost * -1
										extItemNo = 1
										extOrderserSeq = extOrderserSeq & "-2"
										extCommPrice	= (datalist.get(i).deliveryFee.fee + datalist.get(i).deliveryFee.feeVat) * -1
									Else
										If saleType <> "SALE" Then
											extItemNo = -1
											extOrderserSeq = extOrderserSeq & "-1"
											extCommPrice			= (datalist.get(i).deliveryFee.fee + datalist.get(i).deliveryFee.feeVat)
										Else
											extItemNo = 1
											extCommPrice			= (datalist.get(i).deliveryFee.fee + datalist.get(i).deliveryFee.feeVat)
										End If
									End If
								Else
									extItemNo = 0
								End If
							ElseIf extItemNo = 0 AND extItemCost <> 0 Then
								If (extItemCost <> 0) Then
									If (extItemCost < 0) Then
										extItemCost = extItemCost * -1
										extItemNo = 1
										extOrderserSeq = extOrderserSeq & "-2"
										extCommPrice	= (datalist.get(i).deliveryFee.fee + datalist.get(i).deliveryFee.feeVat) * -1
									Else
										If saleType <> "SALE" Then
											extItemNo = -1
											extOrderserSeq = extOrderserSeq & "-1"
											extCommPrice			= (datalist.get(i).deliveryFee.fee + datalist.get(i).deliveryFee.feeVat)
										Else
											extItemNo = 1
											extCommPrice			= (datalist.get(i).deliveryFee.fee + datalist.get(i).deliveryFee.feeVat)
										End If
									End If
								Else
									extItemNo = 0
								End If
							End If

							extReducedPrice			= extItemCost-extOwnCouponPrice
							extTenMeachulPrice      = extItemCost-extOwnCouponPrice
							extTenJungsanPrice      = extTenMeachulPrice - extCommPrice

							extCommSupplyPrice		= 0
							extCommSupplyVatPrice	= 0

							extTenMeachulSupplyPrice	= 0
							extTenMeachulSupplyVatPrice	= 0

							sqlStr = ""
							sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_xSite_JungsanTmp "
							sqlStr = sqlStr & " (sellsite, extOrderserial, extOrderserSeq "
							sqlStr = sqlStr & ", extOrgOrderserial, extItemNo, extItemCost "
							sqlStr = sqlStr & ", extReducedPrice, extOwnCouponPrice, extTenCouponPrice "
							sqlStr = sqlStr & ", extJungsanType, extCommPrice, extTenMeachulPrice "
							sqlStr = sqlStr & ", extTenJungsanPrice, extMeachulDate, extJungsanDate "
							sqlStr = sqlStr & ", extItemName, extItemOptionName, extVatYN "
							sqlStr = sqlStr & ", extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice "
							sqlStr = sqlStr & ", itemid, itemoption "
							sqlStr = sqlStr & ", extitemid, extitemoption,siteNo"
							sqlStr = sqlStr & " ) "
							sqlStr = sqlStr & " values('" & CStr(sellsite) & "', '" & CStr(extOrderserial) & "', '" & CStr(extOrderserSeq) & "'"
							sqlStr = sqlStr & ", '" & CStr(extOrgOrderserial) & "', '" & CStr(extItemNo) & "', '" & CStr(extItemCost) & "'"
							sqlStr = sqlStr & ", '" & CStr(extReducedPrice) & "', '" & CStr(extOwnCouponPrice) & "', '" & CStr(extTenCouponPrice) & "'"
							sqlStr = sqlStr & ", '" & CStr(extJungsanType) & "', '" & CStr(extCommPrice) & "', '" & CStr(extTenMeachulPrice) & "'"
							sqlStr = sqlStr & ", '" & CStr(extTenJungsanPrice) & "', '" & CStr(extMeachulDate) & "', '" & CStr(extJungsanDate) & "'"
							sqlStr = sqlStr & ", '" & CStr(extItemName) & "', convert(varchar(128),'" & CStr(extItemOptionName) & "'), '" & CStr(extVatYN) & "'"
							sqlStr = sqlStr & ", '" & CStr(extCommSupplyPrice) & "', '" & CStr(extCommSupplyVatPrice) & "', '" & CStr(extTenMeachulSupplyPrice) & "', '" & CStr(extTenMeachulSupplyVatPrice) & "'"
							If (tenitemid <> "") Then
								sqlStr = sqlStr & ", '" & CStr(tenitemid) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (tenitemoption <> "") Then
								sqlStr = sqlStr & ", '" & CStr(tenitemoption) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (extitemid <> "") Then
								sqlStr = sqlStr & ", '" & CStr(extitemid) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (extitemoption <> "") Then
								sqlStr = sqlStr & ", '" & CStr(extitemoption) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (siteNo <> "") Then
								sqlStr = sqlStr & ", '" & CStr(siteNo) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							sqlStr = sqlStr & ") "

							If (extItemNo <> 0) Then
								on Error Resume Next
								dbget.execute sqlStr
								'rw sqlStr
								If Err Then
									rw sqlStr
									on Error Goto 0
									response.end
								End If
							End If
						Next

						sqlStr = " EXEC db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_coupang] "
						dbget.Execute sqlStr

					Set datalist = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

Function GetJungsan_WMP(reqDate)
	Dim sellsite : sellsite = "WMP"
	Dim objXML, xmlDOM, strSql, iRbody, i, datalist, strObj, sqlStr, extItemName, extItemOptionName, tenitemid, tenitemoption, extitemoption, siteNo
	Dim p_extOrderserial, p_extOrderserSeq
	Dim retCode, extOrderserial, extJungsanDate, extMeachulDate, extOrderserSeq, gubun, extOrgOrderserial, extVatYN, extItemNo, extItemCost, extReducedPrice
	Dim extOwnCouponPrice, extTenCouponPrice, extTenJungsanPrice, extCommPrice, extTenMeachulPrice, extJungsanType, extCommSupplyVatPrice, extCommSupplyPrice
	Dim extTenMeachulSupplyVatPrice, extTenMeachulSupplyPrice, extitemid

	sqlStr = ""
	sqlStr = " DELETE FROM db_temp.dbo.tbl_xSite_JungsanTmp WHERE sellsite = '" & CStr(sellsite) & "' "
	dbget.execute sqlStr

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "https://w-api.wemakeprice.com/settle/out/getSettleDailyOrderInfo?basicDt="&reqDate
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "apiKey", "b32bfc8ae3d522eb729e96a60d9d277aeb242302c8f9b33fd51dcc3ee739f19b9d974e2e0a8e1ef683ef3a76e4927378"
		objXML.send()
 		If objXML.Status = "200" OR objXML.Status = "201" Then
 			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			'rw iRbody
			Set strObj = JSON.parse(iRbody)
				retCode		= strObj.resultCode			'서버 응답 코드
				If retCode = "200" Then
					Set datalist = strObj.data		'결과리스트 | 결과가 없을 때는 빈 리스트가 리턴
						For i=0 to datalist.length-1
							extOrderserial = datalist.get(i).bundleNo			'배송번호
							extJungsanDate = ""
							extMeachulDate = datalist.get(i).basicDt			'기준일자 | 기준일자(매출일-YYYYMMDD)
							extOrderserSeq = datalist.get(i).orderNo			'주문번호
							If datalist.get(i).orderOptNo <> "0" then			'옵션주문번호
								extOrderserSeq = extOrderserSeq & "-" & datalist.get(i).orderOptNo
								If datalist.get(i).claimNo <> "0" then			'클레임번호
									extOrderserSeq = extOrderserSeq & "-" & datalist.get(i).claimNo
								End If
							End If
							gubun = datalist.get(i).gubun			'구분 | 구분(1:배송완료, 2:환불완료, 3:위메프부담환불)

							If (gubun = "2") OR (gubun = "3") Then  '' 환불
								'// 반품
								extOrgOrderserial		= extOrderserial
								extOrderserSeq			= extOrderserSeq & "-1"
								If (gubun = "3") Then
									extOrderserSeq			= extOrgOrderserial & "-2"
									extMeachulDate			= datalist.get(i).basicDt			'기준일자 | 기준일자(매출일-YYYYMMDD)
									If (extMeachulDate="") Then				''배송완료일도 없으면.  또는 환불완료일/배송완료일이 없는경우 상품번호로 찾아서 정산일을 배송완료일에 넣자.(엑셀에는없음.)
										extMeachulDate		= datalist.get(i).paymentCompleteDt		'결제완료일시 | 결제완료일시(YYYYMMDD)
									End If
								End If
							Else
								'// 정상출고
								extOrgOrderserial		= ""
							End If

							If (LEN(extMeachulDate)=8) Then
								extMeachulDate = LEFT(extMeachulDate,4)&"-"&MID(extMeachulDate,5,2)&"-"&MID(extMeachulDate,7,2)
							End If

							extVatYN 	= "Y"
							extItemNo	= replace(replace(datalist.get(i).completeQty,",","")," ","")  '완료 수량 | 완료 수량(상품단위) (누적합산의 용도가 아닌 참고용 데이터)
							extItemCost = 0
							extReducedPrice = 0
							extOwnCouponPrice = 0
							extTenCouponPrice = 0
							extTenJungsanPrice = 0
							extCommPrice = 0
							extTenMeachulPrice = 0

							If (gubun <> "1") Then
								If extItemNo = "0" Then
									If datalist.get(i).wmpChargeCouponAmt > 0 Then
										extItemNo = 1
									Else
										extItemNo = -1
									End If
								End If
							End If

							If (gubun <> "1") and (datalist.get(i).completeQty) = "0" then
								extItemCost				= 0
								'2023-07-03 김진영..cardChargeCouponAmt 추가
								extOwnCouponPrice		= CLNG((CLNG(datalist.get(i).wmpChargeCouponAmt)+CLNG(datalist.get(i).wmpCartCouponAmt) + CLNG(datalist.get(i).cardChargeCouponAmt))/extItemNo*100 / 100)'' 위메프부담 상품쿠폰 , 위메프부담 장바구니쿠폰 추가 2019/06/02
'								extTenCouponPrice		= CLNG(datalist.get(i).sellerChargeCouponAmt/extItemNo*100 / 100)  '' 판매업체부담 상품쿠폰
								extTenCouponPrice		= CLNG((CLNG(datalist.get(i).sellerChargeCouponAmt)+CLNG(datalist.get(i).sellerCartCouponAmt))/extItemNo*100 /100) '' 판매자 부담 쿠폰 할인금액 + 판매자 부담 장바구니 할인금액
							elseif extItemNo <> 0 then
								extItemCost				= CLNG((datalist.get(i).prodSaleAmt + datalist.get(i).optSaleAmt) *100)/100  ''상품판매가 + 옵션 판매가 추가..
								'2023-07-03 김진영..cardChargeCouponAmt 추가
								extOwnCouponPrice		= CLNG((CLNG(datalist.get(i).wmpChargeCouponAmt)+CLNG(datalist.get(i).wmpCartCouponAmt) + CLNG(datalist.get(i).cardChargeCouponAmt))/extItemNo*100 / 100)'' 위메프부담 상품쿠폰 , 위메프부담 장바구니쿠폰 추가 2019/06/02
'								extTenCouponPrice		= CLNG(datalist.get(i).sellerChargeCouponAmt/extItemNo*100 / 100)  '' 판매업체부담 상품쿠폰
								extTenCouponPrice		= CLNG((CLNG(datalist.get(i).sellerChargeCouponAmt)+CLNG(datalist.get(i).sellerCartCouponAmt))/extItemNo*100 /100) '' 판매자 부담 쿠폰 할인금액 + 판매자 부담 장바구니 할인금액
							end if
							extJungsanType			= "C"
							extTenMeachulPrice		= CLNG(extItemCost) - CLNG(extOwnCouponPrice) - CLNG(extTenCouponPrice)
							extReducedPrice			= CLNG(extTenMeachulPrice)
							extCommPrice			= CLNG(datalist.get(i).saleAgencyFee/extItemNo*100)/100 - extOwnCouponPrice + CLNG(datalist.get(i).wmpDiscountBurdenAmt/extItemNo*100) / 100 + CLNG(datalist.get(i).epFee / extItemNo*100) / 100 ''판매대행수수료-위메프부담쿠폰
							extTenJungsanPrice		= extTenMeachulPrice-extCommPrice
							extCommSupplyVatPrice	= 0
							extCommSupplyPrice		= 0
							extTenMeachulSupplyVatPrice	= 0
							extTenMeachulSupplyPrice	= 0
							extitemID 				= datalist.get(i).prodNo		'상품번호

							'1. 상품 정산 입력
							sqlStr = ""
							sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_xSite_JungsanTmp "
							sqlStr = sqlStr & " (sellsite, extOrderserial, extOrderserSeq "
							sqlStr = sqlStr & ", extOrgOrderserial, extItemNo, extItemCost "
							sqlStr = sqlStr & ", extReducedPrice, extOwnCouponPrice, extTenCouponPrice "
							sqlStr = sqlStr & ", extJungsanType, extCommPrice, extTenMeachulPrice "
							sqlStr = sqlStr & ", extTenJungsanPrice, extMeachulDate, extJungsanDate "
							sqlStr = sqlStr & ", extItemName, extItemOptionName, extVatYN "
							sqlStr = sqlStr & ", extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice "
							sqlStr = sqlStr & ", itemid, itemoption "
							sqlStr = sqlStr & ", extitemid, extitemoption,siteNo "
							sqlStr = sqlStr & " ) "
							sqlStr = sqlStr & " VALUES('" & CStr(sellsite) & "', '" & CStr(extOrderserial) & "', '" & CStr(extOrderserSeq) & "'"
							sqlStr = sqlStr & ", '" & CStr(extOrgOrderserial) & "', '" & CStr(extItemNo) & "', '" & CStr(extItemCost) & "'"
							sqlStr = sqlStr & ", '" & CStr(extReducedPrice) & "', '" & CStr(extOwnCouponPrice) & "', '" & CStr(extTenCouponPrice) & "'"
							sqlStr = sqlStr & ", '" & CStr(extJungsanType) & "', '" & CStr(extCommPrice) & "', '" & CStr(extTenMeachulPrice) & "'"
							sqlStr = sqlStr & ", '" & CStr(extTenJungsanPrice) & "', '" & CStr(extMeachulDate) & "', '" & CStr(extJungsanDate) & "'"
							sqlStr = sqlStr & ", '" & CStr(extItemName) & "', convert(varchar(128),'" & CStr(extItemOptionName) & "'), '" & CStr(extVatYN) & "'"
							sqlStr = sqlStr & ", '" & CStr(extCommSupplyPrice) & "', '" & CStr(extCommSupplyVatPrice) & "', '" & CStr(extTenMeachulSupplyPrice) & "', '" & CStr(extTenMeachulSupplyVatPrice) & "'"
							If (tenitemid <> "") Then
								sqlStr = sqlStr & ", '" & CStr(tenitemid) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (tenitemoption<>"") Then
								sqlStr = sqlStr & ", '" & CStr(tenitemoption) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (extitemid<>"") Then
								sqlStr = sqlStr & ", '" & CStr(extitemid) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (extitemoption<>"") Then
								sqlStr = sqlStr & ", '" & CStr(extitemoption) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (siteNo<>"") Then
								sqlStr = sqlStr & ", '" & CStr(siteNo) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							sqlStr = sqlStr & ") "

							If (extItemNo <> 0) Then
								p_extOrderserial = extOrderserial
								p_extOrderserSeq = extOrderserSeq
								on Error Resume Next
								dbget.execute sqlStr
								If Err Then
									rw sqlStr
									on Error Goto 0
									response.end
								End If
							End If
						Next
						sqlStr = " EXEC db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_WMP_API] "
						dbget.Execute sqlStr
					Set datalist = nothing
				Else
					rw "통신오류 " & reqDate
					response.end
				End If

			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

Function GetJungsan_WMPbeasongpay(reqDate)
	Dim sellsite : sellsite = "WMP"
	Dim objXML, xmlDOM, strSql, iRbody, i, datalist, strObj, sqlStr, extItemName, extItemOptionName, tenitemid, tenitemoption, extitemoption, siteNo
	Dim p_extOrderserial, p_extOrderserSeq
	Dim retCode, extOrderserial, extJungsanDate, extMeachulDate, extOrderserSeq, gubun, extOrgOrderserial, extVatYN, extItemNo, extItemCost, extReducedPrice
	Dim extOwnCouponPrice, extTenCouponPrice, extTenJungsanPrice, extCommPrice, extTenMeachulPrice, extJungsanType, extCommSupplyVatPrice, extCommSupplyPrice
	Dim extTenMeachulSupplyVatPrice, extTenMeachulSupplyPrice, extitemid

	sqlStr = ""
	sqlStr = " DELETE FROM db_temp.dbo.tbl_xSite_JungsanTmp WHERE sellsite = '" & CStr(sellsite) & "' "
	dbget.execute sqlStr

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "https://w-api.wemakeprice.com/settle/out/getSettleShipInfo?basicDt="&reqDate
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "apiKey", "b32bfc8ae3d522eb729e96a60d9d277aeb242302c8f9b33fd51dcc3ee739f19b9d974e2e0a8e1ef683ef3a76e4927378"
		objXML.send()
 		If objXML.Status = "200" OR objXML.Status = "201" Then
 			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			rw iRbody
			Set strObj = JSON.parse(iRbody)
				retCode		= strObj.resultCode			'서버 응답 코드
				If retCode = "200" Then
					Set datalist = strObj.data		'결과리스트 | 결과가 없을 때는 빈 리스트가 리턴
						For i=0 to datalist.length-1
							extOrderserial = datalist.get(i).bundleNo			'배송번호
							extJungsanDate = ""
							extMeachulDate = datalist.get(i).basicDt			'기준일자 | 기준일자(매출일-YYYYMMDD)
							extJungsanType			= "D"
							if (LEN(extMeachulDate)=8) then
								extMeachulDate = LEFT(extMeachulDate,4)&"-"&MID(extMeachulDate,5,2)&"-"&MID(extMeachulDate,7,2)
							end if

							extItemNo				= 1
							If (datalist.get(i).shipAmt	< 0 ) Then					'배송비
								extItemNo = -1
							End If
							gubun = datalist.get(i).gubun							'구분 | 구분(1:배송완료, 2:환불완료)

							If (gubun = "1") then
								'// 정상출고
								extOrderserSeq			= "D"
								extOrgOrderserial		= ""
								extItemName				= "api deliveryPay"
								extItemCost				= CLNG(datalist.get(i).shipAmt/extItemNo*100)/100
							Elseif (gubun = "2") then
								'// 반품
								extOrderserSeq			= "D-" & datalist.get(i).claimBundleNo					'클레임 배송번호
								extOrgOrderserial		= extOrderserial
								extItemName				= "api returnDeliveryPay"
								extItemCost				= CLNG( (datalist.get(i).custChargeAddShipAmt + datalist.get(i).custChargeReturnShipAmt) /extItemNo*100)/100
								If (datalist.get(i).shipAmt	< 0 ) AND (datalist.get(i).sellerChargeReturnShipAmt <> 0) Then
									extItemCost				= CLNG( (datalist.get(i).sellerChargeReturnShipAmt) /1*100)/100
								End If
							Else
								'// ??
								extOrderserSeq			= "DD-" & datalist.get(i).claimBundleNo					'클레임 배송번호
								extOrgOrderserial		= extOrderserial
								extItemName				= "api returnWMPDeliveryPay"
								extItemCost				= CLNG(datalist.get(i).shipAmt/extItemNo*100)/100
							End If
							extItemCost =  CLNG( (datalist.get(i).shipAmt + datalist.get(i).wmpChargeCouponAmt + datalist.get(i).custChargeAddShipAmt + datalist.get(i).custChargeReturnShipAmt + datalist.get(i).claimDeductAmt) /extItemNo*100)/100

							extVatYN = "Y"
							extReducedPrice			= extItemCost
							extOwnCouponPrice		= 0
							extTenCouponPrice		= 0

							extTenJungsanPrice		= extItemCost

							extCommPrice			= CLNG(extReducedPrice-extTenJungsanPrice)
							extCommSupplyVatPrice	= 0
							extCommSupplyPrice		= 0

							extTenMeachulPrice		= extReducedPrice
							extReducedPrice			= CLNG(extReducedPrice)
							extTenMeachulSupplyVatPrice	= 0
							extTenMeachulSupplyPrice	= 0

							If (extMeachulDate="") and (extReducedPrice=0) and (extTenJungsanPrice=0) then
								extItemNo = 0
							End If

							sqlStr = ""
							sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_xSite_JungsanTmp "
							sqlStr = sqlStr & " (sellsite, extOrderserial, extOrderserSeq "
							sqlStr = sqlStr & ", extOrgOrderserial, extItemNo, extItemCost "
							sqlStr = sqlStr & ", extReducedPrice, extOwnCouponPrice, extTenCouponPrice "
							sqlStr = sqlStr & ", extJungsanType, extCommPrice, extTenMeachulPrice "
							sqlStr = sqlStr & ", extTenJungsanPrice, extMeachulDate, extJungsanDate "
							sqlStr = sqlStr & ", extItemName, extItemOptionName, extVatYN "
							sqlStr = sqlStr & ", extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice "
							sqlStr = sqlStr & ", itemid, itemoption "
							sqlStr = sqlStr & ", extitemid, extitemoption,siteNo"
							sqlStr = sqlStr & " ) "
							sqlStr = sqlStr & " values('" & CStr(sellsite) & "', '" & CStr(extOrderserial) & "', '" & CStr(extOrderserSeq) & "'"
							sqlStr = sqlStr & ", '" & CStr(extOrgOrderserial) & "', '" & CStr(extItemNo) & "', '" & CStr(extItemCost) & "'"
							sqlStr = sqlStr & ", '" & CStr(extReducedPrice) & "', '" & CStr(extOwnCouponPrice) & "', '" & CStr(extTenCouponPrice) & "'"
							sqlStr = sqlStr & ", '" & CStr(extJungsanType) & "', '" & CStr(extCommPrice) & "', '" & CStr(extTenMeachulPrice) & "'"
							sqlStr = sqlStr & ", '" & CStr(extTenJungsanPrice) & "', '" & CStr(extMeachulDate) & "', '" & CStr(extJungsanDate) & "'"
							sqlStr = sqlStr & ", '" & CStr(extItemName) & "', convert(varchar(128),'" & CStr(extItemOptionName) & "'), '" & CStr(extVatYN) & "'"
							sqlStr = sqlStr & ", '" & CStr(extCommSupplyPrice) & "', '" & CStr(extCommSupplyVatPrice) & "', '" & CStr(extTenMeachulSupplyPrice) & "', '" & CStr(extTenMeachulSupplyVatPrice) & "'"
							If (tenitemid <> "") Then
								sqlStr = sqlStr & ", '" & CStr(tenitemid) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (tenitemoption <> "") Then
								sqlStr = sqlStr & ", '" & CStr(tenitemoption) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (extitemid <> "") Then
								sqlStr = sqlStr & ", '" & CStr(extitemid) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (extitemoption <> "") Then
								sqlStr = sqlStr & ", '" & CStr(extitemoption) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							If (siteNo <> "") Then
								sqlStr = sqlStr & ", '" & CStr(siteNo) & "'"
							Else
								sqlStr = sqlStr & ", NULL"
							End If
							sqlStr = sqlStr & ") "

							If (extItemNo <> 0) Then
								on Error Resume Next
								dbget.execute sqlStr
								'rw sqlStr
								If Err Then
									rw sqlStr
									on Error Goto 0
									response.end
								End If
							End If
						Next

						sqlStr = " EXEC db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_WMP_API] "
						dbget.Execute sqlStr
					Set datalist = nothing
				Else
					rw "통신오류 " & reqDate
					response.end
				End If

			Set strObj = nothing
		End If
	Set objXML = nothing
End Function

Function getRequestPage(sellsite, reqDate)
	Dim sqlStr
	sqlStr = ""
	sqlStr = sqlStr & " SELECT TOP 1 reqPage, maxPage "
	sqlStr = sqlStr & " FROM [db_temp].[dbo].[tbl_xSite_Jungsan_TimeStamp] "
	sqlStr = sqlStr & " WHERE sellsite = '"& sellsite &"'  "
	sqlStr = sqlStr & " AND reqDate = '"& reqDate &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.Eof Then
		If rsget("reqPage") >= rsget("maxPage") Then
			getRequestPage = "x"
		Else
	   		getRequestPage = rsget("reqPage") + 1
		End If
	Else
		getRequestPage = 1
	End If
	rsget.close
End Function
%>
