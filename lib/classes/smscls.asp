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
		sqlStr = sqlStr + " '[�ٹ�����]���������� �����Ϸ� �Ǿ����ϴ�. �ֹ���ȣ : " + iorderserial + "')"
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
		sqlStr = sqlStr + " '[�ٹ�����]�ֹ����� �Ǿ����ϴ�. ����:" + iacct + " �ݾ�:" + iprice + "��')"
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
		sqlStr = sqlStr + " '[�ٹ�����]�ֹ������� �Աݴ�����Դϴ�.���¾ȳ�:��������534-01-016039.���ٹ�����')"
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
		sqlStr = sqlStr + " '[�ٹ�����]�Ա� �� ����� ������ ��� �Ǿ����ϴ�. ����Ȯ���� �� �Ա� �� �ּ���')"

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
		sqlStr = sqlStr + " '[�ٹ�����]�Ա�Ȯ�� �Ǿ����ϴ�. �ֹ���ȣ�� " + iorderserial + "�Դϴ�.�����մϴ�.')"
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
		sqlStr = sqlStr + " '[�ΰŽ� ��ī����]�Ա� �� ����� ������ ��� �Ǿ����ϴ�. ����Ȯ���� �� �Ա� �� �ּ���')"

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
		sqlStr = sqlStr + " '[�ΰŽ� ��ī����]�Ա�Ȯ�� �Ǿ����ϴ�. �ֹ���ȣ�� " + iorderserial + "�Դϴ�.�����մϴ�.')"

		dbget.Execute sqlStr
	end Sub

	public Sub SendBeaSongOkMsg(byval irechp, byval isongjangno)
		dim sqlStr
		dim delivercoper

		if Not CheckHpOk(irechp) then Exit sub

        delivercoper = "�ù�� �����ù�"
        if Left(isongjangno,1)="6" then
        	delivercoper = "�ù�� CJ�ù�"
        end if

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '1644-6030',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[�ٹ�����]��ǰ�� ���Ǿ����ϴ�.  " + delivercoper + " �����ȣ " + isongjangno + "')"
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
		sqlStr = sqlStr + " '[�ٹ�����]�ֹ��� ��ǰ�� �غ�Ǿ����ϴ�.�������� �൵�� Ȩ������ �� �������ּ���.')"
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

'' E-giftī�� ����
' �ڵ� ����� ������ ��� �����ڵ� ����. �ʿ�� /lib/email/smslib.asp ��ũ��� �ϰ� ����Ұ�.
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

		'# �޽��� �ۼ�
		if Not(rsget("mmsContent")="" or isNull(rsget("mmsContent"))) then
			mmsContent = mmsContent & db2html(rsget("mmsContent")) & vbCrLf
			mmsContent = mmsContent & "------------------------" & vbCrLf
		end if
		mmsContent = mmsContent & "* ������ȣ : " & vbCrLf & rsget("masterCardCode") & vbCrLf & vbCrLf
		mmsContent = mmsContent & "* �������� �̿�ȳ� : ������ȣ ���� �� ��ǰ ����" & vbCrLf
		mmsContent = mmsContent & "* �¶��� �̿�ȳ� : �ٹ�����(www.10x10.co.kr) ���ӡ�α��Ρ渶���ٹ����١�MY��������>Giftī���¶��� ����� �� ������������ȣ ��ϡ� ��ϿϷ� �� ��ǰ ���� �� ��� " & vbCrLf& vbCrLf
		mmsContent = mmsContent & "* ���ູ���� : 1644-6030" & vbCrLf
		mmsContent = mmsContent & "���� AM09:00~PM06:00/���ɽð� PM12:00~01:00" & vbCrLf

    end if
    rsget.Close

    ''' �̰����� ����.
    IF (mmsContent="") then Exit function

    call SendNormalLMS(reqhp,mmsTitle,"1644-6030",mmsContent)

    'On Error Goto 0
    IF Err Then
        sendGiftCardLMSMsg = FALSE
    ELSE
        sendGiftCardLMSMsg = TRUE
    END IF

end function

' �ڵ� ����� ������ ��� �����ڵ� ����. �ʿ�� /lib/email/smslib.asp ��ũ��� �ϰ� ����Ұ�.
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
        mmsTitle	= buyname & "���� �ٹ����� ����Ʈī�带 �����̽��ϴ�."
        cardcoderdm	= rdmSerialEnc(rsget("masterCardCode"))

		mmsContent = mmsContent & "" & buyname & "��(" & sendhp & ")�� �ٹ����� ����Ʈī�带 �����̽��ϴ�." & vbCrLf
		mmsContent = mmsContent & "-----" & vbCrLf & vbCrLf
		mmsContent = mmsContent & "#. �¶��� ���" & vbCrLf
		mmsContent = mmsContent & "http://m.10x10.co.kr/giftcard/view.asp?gc=" & cardcoderdm & "" & vbCrLf & vbCrLf
		mmsContent = mmsContent & "-----" & vbCrLf

    end if
    rsget.Close

    ''' �̰����� ����.
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