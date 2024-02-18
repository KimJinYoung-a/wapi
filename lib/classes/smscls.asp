<%

CONST CNORMALCALLBAKC = "1644-6030"

Class CSMSClass
	public function CheckHpOk(byval irechp)
		CheckHpOk = false
		if Len(irechp)<3 then exit function
		if (Left(irechp,3)="013") or (Left(irechp,3)="011") or (Left(irechp,3)="016") or (Left(irechp,3)="017") or (Left(irechp,3)="018") or (Left(irechp,3)="019") or (Left(irechp,3)="010") then
			CheckHpOk = true
		end if
	end function

	public Sub SendJumunOkMsg(byval irechp, byval iorderserial)
		dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '1644-6030',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[텐바이텐]정상적으로 결제완료 되었습니다. 주문번호 : " + iorderserial + "')"
		dbget.execute sqlStr
	end Sub

	public sub SendAcctJumunOkMsg2(byval irechp, byval iorderserial, byval iacct, byval iprice)
		dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '1644-6030',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[텐바이텐]주문접수 되었습니다. 계좌:" + iacct + " 금액:" + iprice + "원')"
		dbget.execute sqlStr
	end sub

	public Sub SendAcctJumunOkMsg(byval irechp, byval iorderserial)
		dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '1644-6030',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[텐바이텐]주문접수후 입금대기중입니다.계좌안내:조흥은행534-01-016039.㈜텐바이텐')"
		dbget.execute sqlStr
	end Sub

    public Sub SendAcctIpkumCancelMsg(byval irechp, byval iorderserial)
        dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '1644-6030',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[텐바이텐]입금 후 전산상 오류로 취소 되었습니다. 계좌확인후 재 입금 해 주세요')"

		dbget.Execute sqlStr
	end Sub

	public Sub SendAcctIpkumOkMsg(byval irechp, byval iorderserial)
		dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '1644-6030',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[텐바이텐]입금확인 되었습니다. 주문번호는 " + iorderserial + "입니다.감사합니다.')"
		dbget.execute sqlStr
	end Sub


	public Sub SendAcctIpkumCancelMsgACADEMY(byval irechp, byval iorderserial)
        dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '02-741-9070',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[핑거스 아카데미]입금 후 전산상 오류로 취소 되었습니다. 계좌확인후 재 입금 해 주세요')"

		dbget.Execute sqlStr
	end Sub

	public Sub SendAcctIpkumOkMsgACADEMY(byval irechp, byval iorderserial)
		dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '02-741-9070',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[핑거스 아카데미]입금확인 되었습니다. 주문번호는 " + iorderserial + "입니다.감사합니다.')"

		dbget.Execute sqlStr
	end Sub

	public Sub SendBeaSongOkMsg(byval irechp, byval isongjangno)
		dim sqlStr
		dim delivercoper

		if Not CheckHpOk(irechp) then Exit sub

        delivercoper = "택배사 현대택배"
        if Left(isongjangno,1)="6" then
        	delivercoper = "택배사 CJ택배"
        end if

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '1644-6030',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[텐바이텐]상품이 출고되었습니다.  " + delivercoper + " 송장번호 " + isongjangno + "')"
		dbget.execute sqlStr
	end Sub

	public Sub SendJikjupWaitMsg(byval irechp)
		dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '1644-6030',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[텐바이텐]주문한 상품이 준비되었습니다.직접수령 약도는 홈페이지 를 참고해주세요.')"
		dbget.execute sqlStr
	end Sub

	public Sub SendNormalMsg(byval imsg,byval irechp)
		dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '1644-6030',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '" + imsg + "')"
		dbget.execute sqlStr
	end Sub

'' E-gift카드 전송
' 코드 수기로 박은거 모두 공용코드 변경. 필요시 /lib/email/smslib.asp 인크루드 하고 사용할것.
function sendGiftCardLMSMsg(iorderserial)
    Dim sqlStr
    Dim mmsTitle, mmsContent
    Dim sendhp, reqhp
    sendGiftCardLMSMsg = FALSE
    mmsContent = ""

    'On Error Resume Next
    sqlStr = " select mmsTitle, mmsContent"
	sqlStr = sqlStr & " , sendhp, reqhp "
	sqlStr = sqlStr & " , (substring(masterCardCode,1,4)+'-'+substring(masterCardCode,5,4)+'-'+substring(masterCardCode,9,4)+'-'+substring(masterCardCode,13,4)) as masterCardCode "
	sqlStr = sqlStr & " from db_order.dbo.tbl_giftcard_order M"
	sqlStr = sqlStr & " where M.GiftOrderSerial='"&iorderserial&"'"
'rw sqlStr
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.Eof) then
        mmsTitle	= db2html(rsget("mmsTitle"))
        sendhp		= rsget("sendhp")
        reqhp		= rsget("reqhp")

		'# 메시지 작성
		if Not(rsget("mmsContent")="" or isNull(rsget("mmsContent"))) then
			mmsContent = mmsContent & db2html(rsget("mmsContent")) & vbCrLf
			mmsContent = mmsContent & "------------------------" & vbCrLf
		end if
		mmsContent = mmsContent & "* 인증번호 : " & vbCrLf & rsget("masterCardCode") & vbCrLf & vbCrLf
		mmsContent = mmsContent & "* 오프라인 이용안내 : 인증번호 제시 후 상품 구매" & vbCrLf
		mmsContent = mmsContent & "* 온라인 이용안내 : 텐바이텐(www.10x10.co.kr) 접속→로그인→마이텐바이텐→MY쇼핑혜택>Gift카드→온라인 사용등록 및 내역→인증번호 등록→ 등록완료 후 상품 구매 시 사용 " & vbCrLf& vbCrLf
		mmsContent = mmsContent & "* 고객행복센터 : 1644-6030" & vbCrLf
		mmsContent = mmsContent & "평일 AM09:00~PM06:00/점심시간 PM12:00~01:00" & vbCrLf

    end if
    rsget.Close

    ''' 이곳에서 검증.
    IF (mmsContent="") then Exit function

    call SendNormalLMS(reqhp,mmsTitle,"1644-6030",mmsContent)

    'On Error Goto 0
    IF Err Then
        sendGiftCardLMSMsg = FALSE
    ELSE
        sendGiftCardLMSMsg = TRUE
    END IF

end function

' 코드 수기로 박은거 모두 공용코드 변경. 필요시 /lib/email/smslib.asp 인크루드 하고 사용할것.
function sendGiftCardLMSMsg2016(iorderserial)
    Dim sqlStr
    Dim mmsTitle, mmsContent
    Dim sendhp, reqhp, buyname, cardcoderdm
    sendGiftCardLMSMsg2016 = FALSE
    mmsContent = ""

    On Error Resume Next
    sqlStr = " select mmsTitle, mmsContent"
	sqlStr = sqlStr & " , sendhp, reqhp, masterCardCode "
	'sqlStr = sqlStr & " , (substring(masterCardCode,1,4)+'-'+substring(masterCardCode,5,4)+'-'+substring(masterCardCode,9,4)+'-'+substring(masterCardCode,13,4)) as masterCardCode "
	sqlStr = sqlStr & " ,buyname"
	sqlStr = sqlStr & " from db_order.dbo.tbl_giftcard_order M"
	sqlStr = sqlStr & " where M.GiftOrderSerial='"&iorderserial&"'"

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.Eof) then
        reqhp		= rsget("reqhp")
        buyname		= db2html(rsget("buyname"))
        sendhp		= rsget("sendhp")
        mmsTitle	= buyname & "님이 텐바이텐 기프트카드를 보내셨습니다."
        cardcoderdm	= rdmSerialEnc(rsget("masterCardCode"))

		mmsContent = mmsContent & "" & buyname & "님(" & sendhp & ")이 텐바이텐 기프트카드를 보내셨습니다." & vbCrLf
		mmsContent = mmsContent & "-----" & vbCrLf & vbCrLf
		mmsContent = mmsContent & "#. 온라인 등록" & vbCrLf
		mmsContent = mmsContent & "http://m.10x10.co.kr/giftcard/view.asp?gc=" & cardcoderdm & "" & vbCrLf & vbCrLf
		mmsContent = mmsContent & "-----" & vbCrLf

    end if
    rsget.Close

    ''' 이곳에서 검증.
    IF (mmsContent="") then Exit function

    call SendNormalLMS(reqhp,mmsTitle,"1644-6030",mmsContent)

    On Error Goto 0
    IF Err Then
        sendGiftCardLMSMsg2016 = FALSE
    ELSE
        sendGiftCardLMSMsg2016 = TRUE
    END IF

end function


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

%>