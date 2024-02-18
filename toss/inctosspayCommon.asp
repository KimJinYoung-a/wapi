
<%
	'-----------------------------------------------------------------------------
	' 토스 연동 환경설정 페이지 ( ASP )
	' incKtosspayCommon.asp
	' 2019.10.11 원승현 생성
	'-----------------------------------------------------------------------------
%>
<%
    ''Response.charset = "UTF-8"
	'---------------------------------------------------------------------------------
	' 환경변수 선언 (apiKey 값이 sk_test_M1Neq1wmNjM1NeG3Klkj 이면 테스트용)
	' 참고값
	'   - userinfo에서 사용하는 paymethod값은 980
	'   - DB에 들어가는 두글자 키값은 TS
	'-----------------------------------------------------------------------------
    Dim TossPay_Ready_Url '//결제생성요청 URL
	Dim TossPay_Payment_Approve_Url '//결제승인 URL
	Dim TossPay_Payment_Cancel '//결제취소 URL(결제 대기 중인 결제건 취소)
	Dim TossPay_Payment_Refunds '// 결제 환불 URL(결제 완료 후 결제금액 일부 또는 전부 환불)
	Dim TossPay_Payment_Status '// 결제 상태 확인(생성된 결제의 현태 상태를 조회)

    Dim TossPay_RestApi_Key '//RestAPI 키
    Dim TossPay_OrderSuccess_Url '//결제성공시이동할 URL
    Dim TossPay_OrderFail_Url '//결제실패시이동할 URL
    Dim TossPay_OrderCancel_Url '//결제취소시이동할 URL
    Dim TossPay_ApiKey '//가맹점코드
    Dim TossPay_Custom_Json '//결제화면에 보여주고 싶은 custom metadata(key, value형태){"size":"XL","color":"Red"}
	Dim TossPay_Payment_Method_Type '// 결제수단 구분변수(TOSS_MONEY, CARD 중 선택 없으면 전체)
	Dim TossPay_CashReceipt	'// 현금영수증 발급 가능 여부(true, false 기본값 true이므로 제외 시키고 싶을때만 설정)
	Dim TossPay_CashReceiptOption '// 현금영수증 발급타입(CULTURE - 문화비, GENERAL - 일반, PUBLIC_TP - 교통비) 일단 사용안함
	'// 2019-10-16일 기준 토스 지원 카드 목록
	'// 신한 - 1, 현대 - 2, 삼성 - 3, 국민(미지원) - 4, 롯데 - 5, 하나 - 6, 우리(미지원) - 7, 농협(미지원) - 8, 씨티(미지원) - 9, 비씨 - 10
	Dim TossPay_Available_Cards '// 카드사제한목록(없을경우전체){'options':[{'cardCompanyCode':3}]}
    Dim TossPay_LogUse '//로그사용여부
	Dim TossPay_Jungsan_settlement

    '// 토스 관련 각종 RestApi Url 및 Key값
    TossPay_Ready_Url                                  	= "https://pay.toss.im/api/v1/payments" '//결제생성요청 URL
    TossPay_Payment_Approve_Url                        	= "https://pay.toss.im/api/v2/execute" '//결제승인 URL
	TossPay_Payment_Cancel								= "https://pay.toss.im/api/v1/cancel" '// 결제취소 URL
	TossPay_Payment_Refunds 							= "https://pay.toss.im/api/v2/refunds" '// 결제환불 URL
	TossPay_Payment_Status 								= "https://pay.toss.im/api/v1/status" '// 결제 상태 확인 URL
	TossPay_Jungsan_settlement							= "https://pay.toss.im/pay/api/v2/settlement-details"

    if (application("Svr_Info")="Dev") then
		'개발서버
        TossPay_OrderSuccess_Url       = "http://2015www.10x10.co.kr/inipay/tosspay/ordertemp_tossresult.asp" '//결제성공시이동할url
        TossPay_OrderFail_Url          = "http://2015www.10x10.co.kr/inipay/tosspay/ordertemp_tossfail.asp" '//결제실패시이동할url
        TossPay_OrderCancel_Url        = "http://2015www.10x10.co.kr/inipay/tosspay/ordertemp_tossfail.asp" '//결제취소시이동할url
        TossPay_RestApi_Key            = "sk_test_M1Neq1wmNjM1NeG3Klkj"'//가맹점코드(테스트용)
		''TossPay_RestApi_Key            = "sk_live_3AkvOVG7263AkvlMPLN6"'//가맹점코드(실결제용)
		TossPay_Payment_Method_Type	   = ""
		TossPay_CashReceipt			   = ""
		TossPay_CashReceiptOption	   = ""
		TossPay_Available_Cards        = ""
        TossPay_LogUse                 = False

    Else
		'실서버
        TossPay_OrderSuccess_Url       = "https://www.10x10.co.kr/inipay/tosspay/ordertemp_tossresult.asp" '//결제성공시이동할url
        TossPay_OrderFail_Url          = "https://www.10x10.co.kr/inipay/tosspay/ordertemp_tossfail.asp" '//결제실패시이동할url
        TossPay_OrderCancel_Url        = "https://www.10x10.co.kr/inipay/tosspay/ordertemp_tossfail.asp" '//결제취소시이동할url
        ''TossPay_RestApi_Key            = "sk_test_M1Neq1wmNjM1NeG3Klkj"'//가맹점코드(테스트용)
        TossPay_RestApi_Key            = "sk_live_3AkvOVG7263AkvlMPLN6"'//가맹점코드(실결제용)
		TossPay_Payment_Method_Type	   = ""
		TossPay_CashReceipt			   = ""
		TossPay_CashReceiptOption	   = ""
		TossPay_Available_Cards        = ""
        TossPay_LogUse                 = False
    End If

	'---------------------------------------------------------------------------------
	' 로그 파일 선언 ( 루트경로부터 \tosspay\asp\log 폴더까지 생성을 해 놓습니다. )
	'---------------------------------------------------------------------------------
	Dim Write_LogFile
	Write_LogFile = Server.MapPath(".") + "\log\Tosspay_Log_"+Replace(FormatDateTime(Now,2),"-","")+"_asp.txt"


	'-----------------------------------------------------------------------------
	' 로그 기록 함수 ( 디버그용 )
	' 사용 방법 : Call Write_Log(Log_String)
	' Log_String : 로그 파일에 기록할 내용
	'-----------------------------------------------------------------------------
	''Const fsoForReading = 1		'- Open a file for reading. You cannot write to this file.
	''Const fsoForWriting = 2		'- Open a file for writing.
	''Const fsoForAppend = 8		'- Open a file and write to the end of the file.
	Sub Toss_Write_Log(Log_String)
		If Not Tosspay_Log_ Then Exit Sub
		'On Error Resume Next
		Dim oFSO
		Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
		Dim oTextStream
		Set oTextStream = oFSO.OpenTextFile(Write_LogFile, 8, True, 0)
		'-----------------------------------------------------------------------------
		' 내용 기록
		'-----------------------------------------------------------------------------
		oTextStream.WriteLine  CStr(FormatDateTime(Now,0)) + " " + Replace(CStr(Log_String),Chr(0),"'")
		'-----------------------------------------------------------------------------
		' 리소스 해제
		'-----------------------------------------------------------------------------
		oTextStream.Close
		Set oTextStream = Nothing
		Set oFSO = Nothing
	End Sub

	'-----------------------------------------------------------------------------
	' API 호출 함수( POST 전용 - TOSSPAY 연동은 모든 API 호출에 POST만을 사용합니다. )
	' 사용 방법 : Call_API(SiteURL, App_Mode, Param)
	' SiteURL : 호출할 API 주소
	' App_Mode : 데이터 전송 형태 ( 예: json, x-www-form-urlencoded 등 )
	' Param : 전송할 POST 데이터
	'-----------------------------------------------------------------------------
	Function Call_API(SiteURL, App_Mode, Param)
		Dim HTTP_Object

		'-----------------------------------------------------------------------------
		' WinHttpRequest 선언
		'-----------------------------------------------------------------------------
		If (application("Svr_Info")	= "Dev") Then
			set HTTP_Object = Server.CreateObject("WinHttp.WinHttpRequest.5.1")	'xmlHTTP컨퍼넌트 선언
		Else
			Set HTTP_Object = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		End If
		With HTTP_Object
			'API 통신 Timeout 을 30초로 지정
			.SetTimeouts 30000, 30000, 30000, 30000
			.Open "POST", SiteURL, False
			.SetRequestHeader "Content-Type", "application/"+CStr(App_Mode)+"; charset=UTF-8"
			'-----------------------------------------------------------------------------
			' API 전송 정보를 로그 파일에 저장
			'-----------------------------------------------------------------------------
			'Call Toss_Write_Log("Call API   "+CStr(SiteURL)+" Mode : "  + CStr(App_Mode))
			'Call Toss_Write_Log("Call API   "+CStr(SiteURL)+" Data : "  + CStr(Param))
			.Send Param
			.WaitForResponse 60
			'-----------------------------------------------------------------------------
			' 전송 결과를 리턴하기 위해 변수 선언 및 값 대입
			'-----------------------------------------------------------------------------
			Dim Result
			Set Result = New clsHTTP_Toss_Object
			Result.Status = CStr(.Status)
			Result.ResponseText = CStr(.ResponseText)
			'-----------------------------------------------------------------------------
			' API 전송 결과를 로그 파일에 저장
			'-----------------------------------------------------------------------------
			'Call Toss_Write_Log("API Result "+CStr(SiteURL) + " Status : " + CStr(.Status))
			'Call Toss_Write_Log("API Result "+CStr(SiteURL) + " ResponseText : " + CStr(.ResponseText))
		End With
		Set Call_API = Result
	End Function

	'---------------------------------------------------------------------------------
	' 주문 예약 API 호출 함수
	' 사용 방법 : Call toss_reserve(mData)
	' mData - parameter 데이터
	'---------------------------------------------------------------------------------
	Function tossapi_reserve(mData)
		Dim Result, resultValue, tmpJSON, resultJson
		Set Result = Call_API(TossPay_Ready_Url, "json", mData)
        Set resultJson = New aspJson
        resultJson.loadJSON(Result.ResponseText)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "결제 예약 도중 오류가 발생하였습니다."
						.Add "message", resultJson.data("message")
						'.Add "message_code", resultJson.data("errorCode")
						'.Add "code", resultJson.data("code")
					End With
					resultValue = tmpJSON.JSONoutput()
                    Set tmpJSON = Nothing
			End Select
		End With
        Set resultJson = Nothing
		tossapi_reserve = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' 결제 승인 API 호출 함수
	' 사용 방법 : Call tossapi_order_confirm(mData)
	' 토스 결제승인은 x-www-form-urlencoded가 아닌 json으로 보내야됨-_-
	' mData - parameter 데이터
	'---------------------------------------------------------------------------------
	Function tossapi_order_confirm(mData)
		Dim Result, resultValue, tmpJSON, resultJson
		Set Result = Call_API(TossPay_Payment_Approve_Url, "json", mData)
        Set resultJson = New aspJson
        resultJson.loadJSON(Result.ResponseText)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "결제 승인 도중 오류가 발생하였습니다."
						.Add "message", resultJson.data("message")
						'If resultJson.data("msg") <> "" Then
						'	.Add "message", resultJson.data("msg")
						'	.Add "message_code", resultJson.data("errorCode")
						'Else
						'	.Add "message", ""
						'	.Add "message_code", ""
						'End If
					End With
					resultValue = tmpJSON.JSONoutput()
                    Set tmpJSON = Nothing
			End Select
		End With
        Set resultJson = Nothing
		tossapi_order_confirm = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' 주문 접수 확인 API 호출 함수
	' 사용 방법 : Call tossapi_ordercheck(mData)
	' mData - parameter 데이터
	'---------------------------------------------------------------------------------
	Function tossapi_ordercheck(mData)
		Dim Result, resultValue, tmpJSON, resultJson
		Set Result = Call_API(TossPay_Payment_Status, "json", mData)
        Set resultJson = New aspJson
        resultJson.loadJSON(Result.ResponseText)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "결제 상세 내역 호출 도중 오류가 발생하였습니다."
						.Add "message", resultJson.data("message")
						'.Add "message_code", resultJson.data("errorCode")
					End With
					resultValue = tmpJSON.JSONoutput()
                    Set tmpJSON = Nothing
			End Select
		End With
        Set resultJson = Nothing
		tossapi_ordercheck = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' 결제 승인 전 취소 API 호출 함수(이건 결제 환불이 아닌 결제 대기 상태의 취소임)
	' 사용 방법 : Call tossapi_ordercancel(mData)
	' mData - parameter 데이터
	'---------------------------------------------------------------------------------
	Function tossapi_ordercancel(mData)
		Dim Result, resultValue, tmpJSON, resultJson
		Set Result = Call_API(TossPay_Payment_Cancel, "json", mData)
        Set resultJson = New aspJson
        resultJson.loadJSON(Result.ResponseText)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "결제 취소 중 오류가 발생하였습니다."
						.Add "message", resultJson.data("message")
						'.Add "message_code", resultJson.data("errorCode")
					End With
					resultValue = tmpJSON.JSONoutput()
                    Set tmpJSON = Nothing
			End Select
		End With
        Set resultJson = Nothing
		tossapi_ordercancel = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' 환불 API 호출 함수
	' 사용 방법 : Call tossapi_refund(mData)
	' mData - parameter 데이터
	'---------------------------------------------------------------------------------
	Function tossapi_refund(mData)
		Dim Result, resultValue, tmpJSON, resultJson
		Set Result = Call_API(TossPay_Payment_Refunds, "json", mData)
        Set resultJson = New aspJson
        resultJson.loadJSON(Result.ResponseText)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "ERR"
						.Add "message", resultJson.data("message")
						'.Add "message_code", resultJson.data("errorCode")
						'.Add "code", resultJson.data("code")
					End With
					resultValue = tmpJSON.JSONoutput()
                    Set tmpJSON = Nothing
			End Select
		End With
        Set resultJson = Nothing
		tossapi_refund = resultValue
	End Function

	'---------------------------------------------------------------------------------
	' 승인/정산 내역 API 호출 함수
	' 사용 방법 : Call tossapi_Jungsan(mData)
	' mData - parameter 데이터
	'---------------------------------------------------------------------------------
	Function tossapi_Jungsan(mData)
		Dim Result, resultValue, tmpJSON, resultJson
		Set Result = Call_API(TossPay_Jungsan_settlement, "json", mData)
        Set resultJson = New aspJson
        resultJson.loadJSON(Result.ResponseText)
		With Result
			Select Case .Status
				Case 200
					resultValue = .ResponseText
				Case Else
					Set tmpJSON = New aspJSON
					With tmpJSON.data
						.Add "result", "ERR"
						.Add "message", resultJson.data("message")
						'.Add "message_code", resultJson.data("errorCode")
						'.Add "code", resultJson.data("code")
					End With
					resultValue = tmpJSON.JSONoutput()
                    Set tmpJSON = Nothing
			End Select
		End With
        Set resultJson = Nothing
		tossapi_Jungsan = resultValue
	End Function

	Function tossapi_return_order_status_value(v)
		Select Case trim(v)
			Case "PAY_STANDBY"
				'// 결제 건이 생성되었고, 구매자의 결제 진행을 대기 중인 상태.
				'// 이 상태에서 구매자나 가맹점이 결제를 취소할 수 있습니다.
				'// 또한 설정한 '만료 기간'이 도래하면 자동으로 취소 됩니다.
				tossapi_return_order_status_value = "결제 대기 중"

			Case "PAY_APPROVED"
				'// 결제를 위한 구매자 인증 완료되고, 가맹점의 최종 승인을 기다리는 상태.
				'// (결제 생성 시 'autoExecute'를 false로 설정한 경우에만 이 단계를 거칩니다)
				'// 텐바이텐은 최종 결제 확인을 우리쪽에서 하므로 autoExecute가 false 이어야함.
				tossapi_return_order_status_value = "구매자 인증 완료"

			Case "PAY_CANCEL"
				'// 결제가 완료되기 전에 구매자나 가맹점이 결제를 취소한 상태입니다.
				'// (결론적으론 금액의 이동 없이 종료된 건)
				tossapi_return_order_status_value = "결제 취소"

			Case "PAY_PROGRESS"
				'// 구매자가 결제를 승인하여 구매자의 계좌에서 결제 금액을 출금 처리 중인 상태입니다.
				tossapi_return_order_status_value = "결제 진행 중"

			Case "PAY_COMPLETE"
				'// 구매자 및 가맹점의 결제 승인 및 출금이 정상적으로 완료된 상태입니다.
				tossapi_return_order_status_value = "결제 완료"

			Case "REFUND_PROGRESS"
				'// 전액 또는 부분 환불을 진행 중인 상태로, 완료되기 전 까지 다른 환불을 진행할 수 없습니다.
				tossapi_return_order_status_value = "환불 진행 중"

			Case "REFUND_SUCCESS"
				'// 전액 또는 부분 환불이 완료되어, 환불 처리한 금액이 구매자의 계좌로 입금 완료된 상태입니다.
				tossapi_return_order_status_value = "환불 성공"

			Case "SETTLEMENT_COMPLETE"
				'// 결제 완료된 금액에 대해 정산이 완료되어 더 이상 환불이 불가한 상태입니다.
				'// (승인일 또는 구매 확정일로부터 1년 경과)
				tossapi_return_order_status_value = "정산 완료"

			Case "SETTLEMENT_REFUND_COMPLETE"
				'// 전액 또는 부분 환불에 대한 정산이 완료되어 더 이상 환불이 불가한 상태입니다.
				'// (승인일 또는 구매 확정일로부터 1년 경과했거나 전액 환불에 대한 정산 완료된 경우)
				tossapi_return_order_status_value = "환불 정산 완료"

			Case Else
				tossapi_return_order_status_value = ""
		End Select
	End Function

	'-----------------------------------------------------------------------------
	' API 결과 전송용 데이터 구조 선언
	' Status 와 ResponseText 만을 전송한다.
	'-----------------------------------------------------------------------------
	Class clsHTTP_Toss_Object
		private m_Status
		private m_ResponseText

		public property get Status()
			Status = m_Status
		end property

		public property get ResponseText()
			ResponseText = m_ResponseText
		end property

		public property let Status(p_Status)
			m_Status = p_Status
		end property

		public property let ResponseText(p_ResponseText)
			m_ResponseText = p_ResponseText
		end property

		Private Sub Class_Initialize
			m_Status = ""
			m_ResponseText = ""
		End Sub
	End Class
%>
