<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'+-------------------------------------------------------------------------------------------------------------+
'|                                  네 이 버 페 이   결 제   함 수 선 언                                       |
'+----------------------------------------------+--------------------------------------------------------------+
'|                함 수 명                      |                          기    능                            |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayReserve(                       | 결제 예약 호출                                               |
'|        주문번호,상품명,상품수,총결제금액     | - 반환값 : 예약번호 (에러시 ERR)                             |
'|        ,과세금액,배송비,주문자이름)          |                                                              |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayApply(결제번호)                | 결제 요청(승인) 호출                                         |
'|                                              | - 반환값 : 결제번호 Object (에러시 code: ERR)                |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayCheck(결제번호)                | 결제 내역 요청(확인) 호출                                    |
'|                                              | - 반환값 : 결제내역 Object (에러시 code: ERR)                |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayCashAmt(결제번호)              | 현금 영수증 발행 대상 금액 조회 호출                         |
'|                                              | - 반환값 : 결제내역 Object (에러시 code: ERR)                |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayReceipt(결제번호)              | 신용카드 매출 전표 조회 호출                                 |
'|                                              | - 반환값 : 신용카드 매출전표 조회 URL (에러시 ERR)           |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayCancel(결제번호,취소금액,사유,요청자) | 취소요청 호출                                                |
'|                                              | - 반환값 : 취소내역 Object (에러시 code: ERR)                |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayDlvFinish(결제번호)            | 거래완료 호출                                                |
'|                                              | - 반환값 : 취소내역 Object (에러시 code: ERR)                |
'+----------------------------------------------+--------------------------------------------------------------+



Dim NPay_API_URL			'// 백엔드 API 호출 URL
Dim NPay_SvcPC_URL			'// PC웹 결제 API 호출 URL
Dim NPay_SvcMobile_URL		'// 모바일 결제 API 호출 URL
Dim NPay_PartnerID			'// 네이버페이 파트너 ID
Dim NPay_ClientId, NPay_ClientKey	'// 텐바이텐 상점 ID 및 시크릿키

'네이버페이 기본 경로
if (application("Svr_Info")="Dev") then
	NPay_API_URL = "https://dev.apis.naver.com"
	NPay_SvcPC_URL = "https://alpha2-pay.naver.com"
	NPay_SvcMobile_URL = "https://alpha2-m.pay.naver.com"
else
	NPay_API_URL = "https://apis.naver.com"
	NPay_SvcPC_URL = "https://pay.naver.com"
	NPay_SvcMobile_URL = "https://m.pay.naver.com"
end if

'인증키
NPay_PartnerID = "tenbyten"
NPay_ClientId = "FyDBW8XfYK4wly9KVYVz"
NPay_ClientKey = "AzeQydNElY"

'=================================================

'// 결제예약 호출 함수
Function fnCallNaverPayReserve(ordno,itemnm,itemno,totprc,taxAmt,dlvprc,ordunm)
	dim oXML, sURL, sParam
	dim jsResult, oResult, errMsg

	'// 결제예약 (https://[API도메인]/[가맹점ID]/naverpay/payments/v1/reserve)
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/reserve"

	'// 전송 테이터 Setting
	sParam = "modelVersion=2"										'결제 연동방식 (1:즉시승인, 2:인증후승인)
	sParam = sParam & "&merchantPayKey=" & ordno							'주문번호 (임시)
	sParam = sParam & "&productName=" & server.URLEncode(itemnm)			'대표 상품명 (반드시 1개)
	sParam = sParam & "&productCount=" & itemno								'상품수량
	sParam = sParam & "&totalPayAmount=" & totprc							'총 결제 금액
	sParam = sParam & "&taxScopeAmount=" & taxAmt							'과세 금액 (과세+면세=총결제금액)
sParam = sParam & "&taxExScopeAmount=" & (totprc-taxAmt)				'면세 금액 (0이라도 전달)
	if dlvprc>0 then
		sParam = sParam & "&deliveryFee=" & dlvprc			'배송비
	end if
	sParam = sParam & "&returnUrl=" & server.URLEncode(SSLUrl & "/inipay/naverpay/naverPayResult.asp?ordsn=" & rdmSerialEnc(ordno))	'결제 완료 후 이동할 URL (임시주문번호 암호화 후 전달)
	''sParam = sParam & "&purchaserName=" & server.URLEncode(ordunm)			'구매자 성명(사용안함)

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'파라메터 전송

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		errMsg = "ERR:결제예약 통신 오류[001]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		errMsg = "ERR:결제예약 내부 오류[002]"
	end if

	On Error Goto 0

	if errMsg<>"" then
		fnCallNaverPayReserve = errMsg
		Exit Function
	end if


	on Error Resume Next

	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)

	if oResult.code="Success" then
		fnCallNaverPayReserve = oResult.body.reserveId
	else
		errMsg = "ERR:결제예약 오류. " & oResult.message
	end if

	set oResult = Nothing

	IF (Err) then
		errMsg = "ERR:예약결과 파징 오류[003]"
	end if

	On Error Goto 0

	if errMsg<>"" then
		fnCallNaverPayReserve = errMsg
	end if

End Function


'// 결제요청 호출 함수
Function fnCallNaverPayApply(npId)
	dim oXML, sURL, sParam
	dim jsResult, oResult

	'// 결제요청 (https://[API도메인]/[가맹점ID]/naverpay/payments/v1/apply/payment)
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/apply/payment"

	'// 전송 테이터 Setting
	sParam = "paymentId=" & npId										'네이버페이 결제 번호

	Dim oRstMsg
	Set oRstMsg = new cResult											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'파라메터 전송

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제인증 통신 오류[004]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제인증 내부 오류[005]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayApply = oRstMsg
		Exit Function
	end if


	on Error Resume Next

	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)

	if oResult.code="Success" then
		oRstMsg.code = "Success"
		oRstMsg.message = oResult.body.paymentId
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제인증 오류. " & oResult.message
	end if

	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "인증결과 파징 오류[006]"
	end if

	On Error Goto 0

	'// 결과 반환
	set fnCallNaverPayApply = oRstMsg
	Set oRstMsg = Nothing
End Function


'// 결제 내역 조회 호출 함수
Function fnCallNaverPayCheck(npId)
	dim oXML, sURL, sParam
	dim jsResult, oResult

	'// 결제요청 (https://[API도메인]/[가맹점ID]/naverpay/payments/v1/list-of-payment)
	'// 네이버 페이 결제 내역 조회 url 변경(2021.04.08), v1/list-of-payment-->v2.2/list/history
	'sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/list-of-payment"
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v2.2/list/history"

	'// 전송 테이터 Setting
	'sParam = "paymentId=" & npId										'네이버페이 결제 번호

	'// 전송 데이터 형태 변경
	sParam = "{"
	sParam = sParam &"""paymentId"":"""&CStr(npId)&""""
	sParam = sParam &"}"

	Dim oRstMsg
	Set oRstMsg = new cResult											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	'oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "Content-Type", "application/json"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'파라메터 전송

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제승인 통신 오류[007]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제승인 내부 오류[008]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayCheck = oRstMsg
		Exit Function
	end if


	on Error Resume Next

	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)

	if oResult.code="Success" then
		oRstMsg.code = "Success"
		''fnCallNaverPayCheck = jsResult			'Raw 결과값 반환
		Set fnCallNaverPayCheck = oResult			'파징된 정보 반환
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제승인 오류. " & oResult.message
	end if

	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "승인결과 파징 오류[009]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayCheck = oRstMsg
	end if

	Set oRstMsg = Nothing
End Function



'// 현금 영수증 발행 대상 금액 조회
Function fnCallNaverPayCashAmt(npId)
	dim oXML, sURL, sParam
	dim jsResult, oResult

	'// 현금성 결제금액 요청 (https://[API도메인]/[가맹점ID]/naverpay/payments/v1/receipt/cash-amount)
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/receipt/cash-amount"

	'// 전송 테이터 Setting
	sParam = "paymentId=" & npId										'네이버페이 결제 번호

	Dim oRstMsg
	Set oRstMsg = new cResult											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'파라메터 전송

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "현금영수증 통신 오류[010]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "현금영수증 내부 오류[011]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayCashAmt = oRstMsg
		Exit Function
	end if


	on Error Resume Next

	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)

	if oResult.code="Success" then
		oRstMsg.code = "Success"
		''fnCallNaverPayCashAmt = jsResult			'Raw 결과값 반환
		Set fnCallNaverPayCashAmt = oResult			'파징된 정보 반환
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "현금영수증 오류. " & oResult.message
	end if

	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "현금대상 확인결과 파징 오류[012]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayCashAmt = oRstMsg
	end if

	Set oRstMsg = Nothing
End Function

'// 신용카드 매출 전표 조회 호출 함수
Function fnCallNaverPayReceipt(npId)
	dim oXML, sURL, sParam
	dim jsResult, oResult, errMsg

	'// 신용카드 매출전표 (https://[API도메인]/[가맹점ID]/naverpay/payments/v1/receipt/credit-card)
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/receipt/credit-card"

	'// 전송 테이터 Setting
	sParam = "paymentId=" & npId										'네이버페이 결제 번호

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'파라메터 전송

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		errMsg = "ERR:전표조회 통신 오류[001]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		errMsg = "ERR:전표조회 내부 오류[002]"
	end if

	On Error Goto 0

	if errMsg<>"" then
		fnCallNaverPayReceipt = errMsg
		Exit Function
	end if


	on Error Resume Next

	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)

	if oResult.code="Success" then
		fnCallNaverPayReceipt = oResult.body.receiptUrl
	else
		errMsg = "ERR:전표조회 오류. " & oResult.message
	end if

	set oResult = Nothing

	IF (Err) then
		errMsg = "ERR:전표조회 파징 오류[003]"
	end if

	On Error Goto 0

	if errMsg<>"" then
		fnCallNaverPayReceipt = errMsg
	end if

End Function


'// 취소요청 호출 함수
Function fnCallNaverPayCancel(npId,cancelAmount,cancelReason,nPayCancelRequester)
	dim oXML, sURL, sParam
	dim jsResult, oResult

	'// 결제요청 (https://[API도메인]/[가맹점ID]/naverpay/payments/v1/cancel)
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/cancel"

	'// 전송 테이터 Setting
	sParam = "paymentId=" & npId										'네이버페이 결제 번호
	''sParam = sParam &"&merchantPayKey=" & merchantPayKey                '가맹점의 결제번호(선택)
    sParam = sParam &"&cancelAmount=" & cancelAmount                    '취소금액
    sParam = sParam &"&cancelReason=" & cancelReason                    '취소사유
    sParam = sParam &"&cancelRequester=" & nPayCancelRequester          '취소요청자(1:구매자,2:가맹점관리자)
    sParam = sParam &"&taxScopeAmount=" & cancelAmount                   '괴세대상금액(선택)
    sParam = sParam &"&taxExScopeAmount=" & 0                            '면세대상금액(선택)

	Dim oRstMsg
	Set oRstMsg = new cResult											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'파라메터 전송

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제인증 통신 오류[004]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제인증 내부 오류[005]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayCancel = oRstMsg
		Exit Function
	end if


	on Error Resume Next

	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)

	if oResult.code="Success" then
		oRstMsg.code = "Success"
		Set fnCallNaverPayCancel = oResult			'파징된 정보 반환
		''oRstMsg.message = oResult.body.paymentId
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제인증 오류. " & oResult.message
	end if

	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "인증결과 파징 오류[006]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallNaverPayCancel = oRstMsg
	end if

	Set oRstMsg = Nothing

End Function

'// 거래완료 호출 함수
Function fnCallNaverPayDlvFinish(npId)
    dim oXML, sURL, sParam
	dim jsResult, oResult

	'// 결제요청 (https://[API도메인]/[가맹점ID]/naverpay/payments/v1/purchase-confirm)
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/purchase-confirm"

	'// 전송 테이터 Setting
	sParam = "paymentId=" & npId										'네이버페이 결제 번호
    sParam = sParam &"&requester=2"                                     '요청자(1:구매자,2:가맹점관리자)

	Dim oRstMsg
	Set oRstMsg = new cResult											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'파라메터 전송

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제인증 통신 오류[004]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제인증 내부 오류[005]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayDlvFinish = oRstMsg
		Exit Function
	end if


	on Error Resume Next

	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)

	if oResult.code="Success" then
		oRstMsg.code = "Success"
		Set fnCallNaverPayDlvFinish = oResult			'파징된 정보 반환
		''oRstMsg.message = oResult.body.paymentId
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제인증 오류. " & oResult.message
	end if

	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "인증결과 파징 오류[006]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallNaverPayDlvFinish = oRstMsg
	end if

	Set oRstMsg = Nothing
End Function


'// 결제 내역 조회 호출 함수(정산list)
Function fnCallNaverPayList(startTime,endTime)
	dim oXML, sURL, sParam
	dim jsResult, oResult

	'// 결제요청 (https://[API도메인]/[가맹점ID]/naverpay/payments/v1/list-of-payment)
	'// 네이버 페이 결제 내역 조회 url 변경(2021.04.08), v1/list-of-payment-->v2.2/list/history
	'sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/list-of-payment"
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v2.2/list/history"

	'// 전송 테이터 Setting
	''YYYYMMDDHHMMSS
	'sParam = "startTime=" & startTime									'검색시작일
	'sParam = sParam&"&endTime=" & endTime									'검색시작일
	'// 전송 데이터 형태 변경
	sParam = "{"
	sParam = sParam &"""startTime"":"""&CStr(startTime)&""","
	sParam = sParam &"""endTime"":"""&CStr(endTime)&""""
	sParam = sParam &"}"

	Dim oRstMsg
	Set oRstMsg = new cResult											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	'oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "Content-Type", "application/json"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'파라메터 전송

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제조회 통신 오류[007]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제조회 내부 오류[008]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayList = oRstMsg
		Exit Function
	end if


	on Error Resume Next

	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)

	if oResult.code="Success" then
		oRstMsg.code = "Success"
		''fnCallNaverPayCheck = jsResult			'Raw 결과값 반환
		Set fnCallNaverPayList = oResult			'파징된 정보 반환
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제조회 오류. " & oResult.message
	end if

	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제조회 파징 오류[009]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayList = oRstMsg
	end if

	Set oRstMsg = Nothing
End Function

'// 결제 내역 조회 호출 함수(정산list)(신규 api용 페이징 코드 추가)2021.04.09
Function fnCallNaverPayListNew(startTime,endTime,t)
	dim oXML, sURL, sParam
	dim jsResult, oResult

	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v2.2/list/history"

	Dim oRstMsg
	Set oRstMsg = new cResult											'결과객체 생성

	'// 호출 처리
	on Error Resume Next


	'// 전송 데이터 setting
	sParam = "{"
	sParam = sParam &"""startTime"":"""&CStr(startTime)&""","
	sParam = sParam &"""endTime"":"""&CStr(endTime)&""","
	sParam = sParam &"""pageNumber"":"&CInt(t)&","
	sParam = sParam &"""rowsPerPage"":100"
	sParam = sParam &"}"

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	'oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "Content-Type", "application/json"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'파라메터 전송

	if oXML.status=200 then
		If i = 1 Then
			jsResult = oXML.responseText
		Else
			jsResult = jsResult & oXML.responseText
		End If
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제조회 통신 오류[007]"
	end if

	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제조회 내부 오류[008]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayListNew = oRstMsg
		Exit Function
	end if


	on Error Resume Next

	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)

	if oResult.code="Success" then
		oRstMsg.code = "Success"
		''fnCallNaverPayCheck = jsResult			'Raw 결과값 반환
		Set fnCallNaverPayListNew = oResult			'파징된 정보 반환
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제조회 오류. " & oResult.message
	end if

	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제조회 파징 오류[009]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayListNew = oRstMsg
	end if

	Set oRstMsg = Nothing
End Function

'// 결제 내역 조회 페이지 수 확인
Function fnCallNaverPayListPageCheck(startTime,endTime)
	dim oXML, sURL, sParam
	dim jsResult, oResult

	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v2.2/list/history"

	'// 전송 데이터 형태 변경
	sParam = "{"
	sParam = sParam &"""startTime"":"""&CStr(startTime)&""","
	sParam = sParam &"""endTime"":"""&CStr(endTime)&""","
	sParam = sParam &"""rowsPerPage"":100"
	sParam = sParam &"}"

	Dim oRstMsg
	Set oRstMsg = new cResult											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	'oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "Content-Type", "application/json"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'파라메터 전송

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제조회 통신 오류[007]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제조회 내부 오류[008]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayListPageCheck = oRstMsg
		Exit Function
	end if


	on Error Resume Next

	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)

	if oResult.code="Success" then
		oRstMsg.code = "Success"
		''fnCallNaverPayCheck = jsResult			'Raw 결과값 반환
		Set fnCallNaverPayListPageCheck = oResult		'파징된 정보 반환
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제조회 오류. " & oResult.message
	end if

	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제조회 파징 오류[009]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayListPageCheck = oRstMsg
	end if

	Set oRstMsg = Nothing
End Function

'==================================================
'// 결과 반환용 객체 선언
Class cResult
	public code
	public message

	Private Sub Class_Initialize()
		code = ""
		message = ""
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class
%>
