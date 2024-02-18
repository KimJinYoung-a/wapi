<%
CONST CNORMALCALLBAKC = "1644-6030"
CONST CIPJUMSHOPCALLBAKC = "1644-6035"

function SendNormalSMS(reqhp,callback,smstext)
    dim sqlStr, RetRows
    if callback="" then callback=CNORMALCALLBAKC

    sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
	sqlStr = sqlStr + " values('" + reqhp + "',"
	sqlStr = sqlStr + " '" + callback + "',"
	sqlStr = sqlStr + " '1',"
	sqlStr = sqlStr + " getdate(),"
	sqlStr = sqlStr + " '" + html2db(smstext) + "')"

	dbget.Execute sqlStr, RetRows

	SendNormalSMS = (RetRows=1)
end function

function SendNormalSMS_LINK(reqhp,callback,smstext)  ''링크드 SMS 서버에서 발송 //2015/08/17
    dim sqlStr, RetRows
    if callback="" then callback=CNORMALCALLBAKC

    ' sqlStr = "insert into [SMSDB].[db_infoSMS].dbo.em_smt_tran (date_client_req, content, callback, service_type, broadcast_yn, msg_status,recipient_num) "
	' sqlStr = sqlStr + " values(getdate(),'"+html2db(smstext)+"','"+callback+"','0','N','1','"+reqhp+"')"

	sqlStr = "INSERT INTO SMSDB.[db_kakaoSMS].[dbo].[SMS_MSG]( REQDATE, STATUS, TYPE, PHONE, CALLBACK, MSG )"
	sqlStr = sqlStr & "		select"
	sqlStr = sqlStr & "		getdate() , '1', '0', convert(varchar(16),N'"& reqhp &"'), convert(varchar(16),N'"& callback &"'), convert(varchar(80),N'"& html2db(smstext) &"')"

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr, RetRows

	SendNormalSMS_LINK = (RetRows=1)
end function

function SendNormalLMS(reqhp, title, callback, smstext)
    dim sqlStr, RetRows
    if callback="" then callback=CNORMALCALLBAKC

    ''if LenB(smstext) > 2000 then
    ''	smstext = LeftB(smstext, 2000)
    ''end if

	' IF application("Svr_Info") = "Dev" THEN
    ' 	sqlStr = " insert into [ACADEMYDB].db_LgSMS.dbo.mms_msg( "
    ' else
    ' 	sqlStr = " insert into [LOGISTICSDB].db_LgSMS.dbo.mms_msg( "
    ' end if

	' sqlStr = sqlStr + " 	subject "
	' sqlStr = sqlStr + " 	, phone "
	' sqlStr = sqlStr + " 	, callback "
	' sqlStr = sqlStr + " 	, status "
	' sqlStr = sqlStr + " 	, reqdate "
	' sqlStr = sqlStr + " 	, msg "
	' sqlStr = sqlStr + " 	, file_cnt "
	' sqlStr = sqlStr + " 	, file_path1 "
	' sqlStr = sqlStr + " 	, expiretime) "
	' sqlStr = sqlStr + " values( "
	' sqlStr = sqlStr + " 	'" + html2db(title) + "' "
	' sqlStr = sqlStr + " 	, '" + CStr(reqhp) + "' "
	' sqlStr = sqlStr + " 	, '" + callback + "' "
	' sqlStr = sqlStr + " 	, '0' "
	' sqlStr = sqlStr + " 	, getdate() "
	' ''sqlStr = sqlStr + " 	, '" + html2db(smstext) + "' "
	' sqlStr = sqlStr + " 	, convert(varchar(4000),'" + html2db(smstext) + "') "
	' sqlStr = sqlStr + " 	, 0 "
	' sqlStr = sqlStr + " 	, null "
	' sqlStr = sqlStr + " 	, '43200' "
	' sqlStr = sqlStr + " ) "

	sqlStr = "INSERT INTO SMSDB.[db_kakaoSMS].[dbo].MMS_MSG ( REQDATE, STATUS, TYPE, PHONE, CALLBACK, SUBJECT, MSG, FILE_CNT )"
	sqlStr = sqlStr & "		select"
	sqlStr = sqlStr & "		getdate(), '1' , '0', '"& reqhp &"', '"& callback &"', convert(varchar(120),N'"& html2db(title) &"'), convert(varchar(4000),N'"& html2db(smstext) &"'), '1'"

	'response.write sqlStr &"<Br>"
	dbget.Execute sqlStr, RetRows

	SendNormalLMS = (RetRows=1)
end function

function SendNormalLMSTimeFix(reqhp, title, callback, smstext)
    dim sqlStr, RetRows
	dim hourCnt

	hourCnt = 0
	do while (Hour(DateAdd("h", hourCnt, Now())) <= 8 or Hour(DateAdd("h", hourCnt, Now())) >= 21)
		hourCnt = hourCnt + 1
	loop

    if callback="" then callback=CNORMALCALLBAKC

	' IF application("Svr_Info") = "Dev" THEN
    ' 	sqlStr = " insert into [ACADEMYDB].db_LgSMS.dbo.mms_msg( "
    ' else
    ' 	sqlStr = " insert into [LOGISTICSDB].db_LgSMS.dbo.mms_msg( "
    ' end if

	' sqlStr = sqlStr + " 	subject "
	' sqlStr = sqlStr + " 	, phone "
	' sqlStr = sqlStr + " 	, callback "
	' sqlStr = sqlStr + " 	, status "
	' sqlStr = sqlStr + " 	, reqdate "
	' sqlStr = sqlStr + " 	, msg "
	' sqlStr = sqlStr + " 	, file_cnt "
	' sqlStr = sqlStr + " 	, file_path1 "
	' sqlStr = sqlStr + " 	, expiretime) "
	' sqlStr = sqlStr + " values( "
	' sqlStr = sqlStr + " 	'" + html2db(title) + "' "
	' sqlStr = sqlStr + " 	, '" + CStr(reqhp) + "' "
	' sqlStr = sqlStr + " 	, '" + callback + "' "
	' sqlStr = sqlStr + " 	, '0' "
	' sqlStr = sqlStr + " 	, dateAdd(hour, " & hourCnt & ", getdate()) "
	' sqlStr = sqlStr + " 	, convert(varchar(4000),'" + html2db(smstext) + "') "
	' sqlStr = sqlStr + " 	, 0 "
	' sqlStr = sqlStr + " 	, null "
	' sqlStr = sqlStr + " 	, '43200' "
	' sqlStr = sqlStr + " ) "

	sqlStr = "INSERT INTO SMSDB.[db_kakaoSMS].[dbo].MMS_MSG ( REQDATE, STATUS, TYPE, PHONE, CALLBACK, SUBJECT, MSG, FILE_CNT )"
	sqlStr = sqlStr & "		select"
	sqlStr = sqlStr & "		dateAdd(hour, " & hourCnt & ", getdate()), '1' , '0', '"& reqhp &"', '"& callback &"', convert(varchar(120),N'"& html2db(title) &"'), convert(varchar(4000),N'"& html2db(smstext) &"'), '1'"

	'response.write sqlStr &"<Br>"
	dbget.Execute sqlStr, RetRows

	SendNormalLMSTimeFix = (RetRows=1)
end function

function SendOverLengthSMS(reqhp,callback,smstext)
    dim smstext1, smstext2, smstext3
    dim retVal : retVal=false
    if callback="" then callback=CNORMALCALLBAKC

    if LenB(smstext)>160 then
        smstext1 = LeftB(smstext,80)
        smstext2 = MidB(smstext,81,80)
        smstext3 = MidB(smstext,161,80)
    elseif LenB(smstext)>80 then
        smstext1 = LeftB(smstext,80)
        smstext2 = MidB(smstext,81,80)
    else
        smstext1 = smstext
    end if

    if (Trim(smstext1)<>"") then
        retVal = SendNormalSMS(reqhp,callback,smstext1)
    end if

    if (retVal) and (Trim(smstext2)<>"") then
        retVal = SendNormalSMS(reqhp,callback,smstext2)
    end if

    if (retVal) and (Trim(smstext3)<>"") then
        retVal = SendNormalSMS(reqhp,callback,smstext3)
    end if

    SendOverLengthSMS = retVal
end function

function SendMultiRowsSMS(reqhp,callback,smstext,spliter)
    dim MaxRows : MaxRows=10
    dim smstextArr, i : i=0
    dim retVal : retVal=false
    if (callback="") then callback=CNORMALCALLBAKC
    if (spliter="") then spliter=VbCrlf

    ''LMS로 변경
    if LenB(smstext)>80 then
        retVal =SendNormalLMS(reqhp, "", callback, smstext)  ''title
    else
        retVal =SendNormalSMS(reqhp,callback,smstext)
    end if
''    smstextArr = split(smstext,spliter)
''
''    if IsArray(smstextArr) then
''        for i=LBound(smstextArr) to UBound(smstextArr)
''            if (i>MaxRows) then Exit for
''            if (Trim(smstextArr(i))<>"") then
''                retVal = SendNormalSMS(reqhp,callback,smstextArr(i))
''            end if
''        next
''    else
''        retVal =SendNormalSMS(reqhp,callback,smstext)
''    end if
''    SendMultiRowsSMS = retVal
end function

function SendMiChulgoSMS(detailidx)
    dim oneMisend, smstext, buyhp
    dim maytitle, pos1,pos2
    dim IsIpjumShop		: IsIpjumShop = False
    dim CallBackNumber	: CallBackNumber = CNORMALCALLBAKC
    dim sqlStr

    set oneMisend = new COldMiSend
        oneMisend.FRectDetailIDx = detailidx
        oneMisend.getOneOldMisendItem

        smstext = oneMisend.FOneItem.getSMSText
        buyhp = oneMisend.FOneItem.FBuyHP

	sqlStr = " select top 1 accountdiv from db_order.dbo.tbl_order_master where orderserial = '" + CStr(oneMisend.FOneItem.FOrderserial) + "' "
	rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
		if (rsget("accountdiv") = "50") then
			CallBackNumber = CIPJUMSHOPCALLBAKC
		end if
	end if
	rsget.Close

    if (smstext<>"") and (buyhp<>"") then
        ''SendMiChulgoSMS = SendMultiRowsSMS(buyhp,"",smstext,vbCrlf)
        if (LenB(smstext)>80) then  ''LMS
            pos1 = InStr(smstext,"[")
            pos2 = InStr(smstext,"]")
            maytitle = ""
            if (pos1>0) and (pos2>0) and (pos2>pos1) then
                maytitle = Mid(smstext,pos1+1,pos2-pos1-1)
            end if

            SendMiChulgoSMS =SendNormalLMS(buyhp, maytitle, CallBackNumber, smstext)  ''title
        else
            SendMiChulgoSMS =SendNormalSMS(buyhp,CallBackNumber,smstext)
        end if

        call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[SMS "+ buyhp +"]" + html2db(smstext))
    end if
    set oneMisend = Nothing
end function

function SendMiChulgoSMSWithMessage(detailidx, smsmessage)
    dim oneMisend, smstext, buyhp
    dim maytitle, pos1,pos2
    dim IsIpjumShop		: IsIpjumShop = False
    dim CallBackNumber	: CallBackNumber = CNORMALCALLBAKC
    dim sqlStr

    set oneMisend = new COldMiSend
        oneMisend.FRectDetailIDx = detailidx
        oneMisend.getOneOldMisendItem

        ''smstext = oneMisend.FOneItem.getSMSText
        buyhp = oneMisend.FOneItem.FBuyHP

	sqlStr = " select top 1 accountdiv from db_order.dbo.tbl_order_master where orderserial = '" + CStr(oneMisend.FOneItem.FOrderserial) + "' "
	rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
		if (rsget("accountdiv") = "50") then
			CallBackNumber = CIPJUMSHOPCALLBAKC
		end if
	end if
	rsget.Close

	smstext = smsmessage
	smstext = Replace(smstext, "[상품명]", oneMisend.FOneItem.FItemname)
	smstext = Replace(smstext, "[상품코드]", oneMisend.FOneItem.FItemid)

    if (smstext<>"") and (buyhp<>"") then
        ''SendMiChulgoSMSWithMessage = SendMultiRowsSMS(buyhp,"",smstext,vbCrlf)
        if (LenB(smstext)>80) then  ''LMS
            pos1 = InStr(smstext,"[")
            pos2 = InStr(smstext,"]")
            maytitle = ""
            if (pos1>0) and (pos2>0) and (pos2>pos1) then
                maytitle = Mid(smstext,pos1+1,pos2-pos1-1)
            end if

            SendMiChulgoSMSWithMessage =SendNormalLMS(buyhp, maytitle, CNORMALCALLBAKC, smstext)  ''title
        else
            SendMiChulgoSMSWithMessage =SendNormalSMS(buyhp,CNORMALCALLBAKC,smstext)
        end if

        Call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[SMS "+ buyhp +"]" + html2db(smstext))
    end if
    set oneMisend = Nothing
end function

function SendMiChulgoSMS_CS(csdetailidx)
    dim oneMisend, smstext, buyhp
    dim maytitle, pos1,pos2
    set oneMisend = new CCSMifinishMaster
        oneMisend.FRectCSDetailIDx = csdetailidx
        oneMisend.getOneMifinishItem

        smstext = oneMisend.FOneItem.getSMSText
        buyhp = oneMisend.FOneItem.FBuyHP


    if (smstext<>"") and (buyhp<>"") then
        ''SendMiChulgoSMS_CS = SendMultiRowsSMS(buyhp,"",smstext,vbCrlf)
        if (LenB(smstext)>80) then  ''LMS
            pos1 = InStr(smstext,"[")
            pos2 = InStr(smstext,"]")
            maytitle = ""
            if (pos1>0) and (pos2>0) and (pos2>pos1) then
                maytitle = Mid(smstext,pos1+1,pos2-pos1-1)
            end if

            SendMiChulgoSMS_CS =SendNormalLMS(buyhp, maytitle, CNORMALCALLBAKC, smstext)  ''title
        else
            SendMiChulgoSMS_CS =SendNormalSMS(buyhp,CNORMALCALLBAKC,smstext)
        end if

        call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[SMS "+ buyhp +"]" + html2db(smstext))
    end if
    set oneMisend = Nothing
end function

function SendMiChulgoSMS_off(detailidx)
    dim oneMisend, smstext, buyhp
    dim maytitle, pos1,pos2
    set oneMisend = new cupchebeasong_list
        oneMisend.FRectDetailIDx = detailidx
        oneMisend.fOneOldMisendItem()

        smstext = oneMisend.FOneItem.getSMSText
        buyhp = oneMisend.FOneItem.FBuyHP

    if (smstext<>"") and (buyhp<>"") then
        '''SendMiChulgoSMS_off = SendMultiRowsSMS(buyhp,"",smstext,vbCrlf)
        if (LenB(smstext)>80) then  ''LMS
            pos1 = InStr(smstext,"[")
            pos2 = InStr(smstext,"]")
            maytitle = ""
            if (pos1>0) and (pos2>0) and (pos2>pos1) then
                maytitle = Mid(smstext,pos1+1,pos2-pos1-1)
            end if

            SendMiChulgoSMS_off =SendNormalLMS(buyhp, maytitle, CNORMALCALLBAKC, smstext)  ''title
        else
            SendMiChulgoSMS_off =SendNormalSMS(buyhp,CNORMALCALLBAKC,smstext)
        end if

        '//cs¸??       'call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[SMS "+ buyhp +"]" + html2db(smstext))
    end if
    set oneMisend = Nothing
end function

public Sub SendAcctCancelMsg(byval irechp, byval iorderserial)
	dim sqlStr, userid, userKey

	if Not CheckHpOk(irechp) then Exit sub

	if Not CheckSendKakaoTalk(iorderserial, userid, userKey) then
		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '1644-6030',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[텐바이텐]승인 취소 되었습니다. 주문번호 : " + iorderserial + "')"
		dbget.execute(sqlStr)
	else
		if userKey<>"" then
			sqlStr = "Insert into db_sms.dbo.tbl_kakao_tran (tr_userid, tr_kakaoUsrKey, tr_info1, tr_msg) values "
			sqlStr = sqlStr & " ('" & userid & "',"
			sqlStr = sqlStr & " '" & userKey & "',"
			sqlStr = sqlStr & " '" & iorderserial & "',"
			sqlStr = sqlStr & " '[텐바이텐] 주문이 승인취소 되었습니다." & vbCrLf & vbCrLf
			sqlStr = sqlStr & "주문번호 : " & iorderserial & vbCrLf & vbCrLf
			sqlStr = sqlStr & "앞으로도 많은 이용 바랍니다. 감사합니다.(미소)')"
			dbget.execute(sqlStr)
		end if
	end if
end Sub

public function CheckHpOk(byval irechp)
	CheckHpOk = false
	if Len(irechp)<3 then exit function
	if (Left(irechp,3)="013") or (Left(irechp,3)="011") or (Left(irechp,3)="016") or (Left(irechp,3)="017") or (Left(irechp,3)="018") or (Left(irechp,3)="019") or (Left(irechp,3)="010") then
		CheckHpOk = true
	end if
end function

'// 카카오톡 발송 여부 확인(주문건)
public function CheckSendKakaoTalk(byval iordsn, byref uid, byref ukey)
	dim sqlStr
	CheckSendKakaoTalk = false
	if Len(iordsn)<11 then exit function
	sqlStr = "Select userid From [db_sms].[dbo].tbl_kakao_chkSend Where orderserial='" & iordsn & "'"
	rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
		CheckSendKakaoTalk = true
		uid = rsget("userid")
	end if
	rsget.Close

	'카카오톡 연결정보 접수
	if uid<>"" then
		sqlStr = "select K.kakaoUserKey " &_
				" from db_sms.dbo.tbl_kakaoUser as K " &_
				"	join db_user.dbo.tbl_user_n as U " &_
				"		on K.userid=U.userid " &_
				" where U.userid='" & uid & "'"
		rsget.Open sqlStr,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			ukey = rsget(0)
		end if
		rsget.Close
	end if
end function

'// 카카오 알림톡으로 문자 발송 (2017.08.29; 허진원 - 링크드 SMS 서버에서 발송)
Sub SendKakaoMsg_LINK(reqhp,callback,tmpcd,ttext,fsendtp,ftit,ftext,btnJson)
	'알림톡 템플릿에 등록 후 승인 받은 형태로만 카카오톡으로 전송가능 (안그러면 무조건 SMS로 발송)
	'2017.11.30: v4 모듈로 판올림, Button_JSON 추가
    dim sqlStr, RetRows
    if callback="" then callback = CNORMALCALLBAKC
    if fsendtp="" then fsendtp="SMS"
    if ftext="" and ttext<>"" then ftext = ttext

    sqlStr = "INSERT INTO [SMSDB].[db_kakaomsg_v4].dbo.KKO_MSG (REQDATE, STATUS, PHONE, CALLBACK, MSG, TEMPLATE_CODE, FAILED_TYPE, FAILED_SUBJECT, FAILED_MSG, BUTTON_JSON) VALUES "
	sqlStr = sqlStr + " (getdate(),'1', "
	sqlStr = sqlStr + " '" & reqhp & "', "				'-- 수신자 휴대폰 번호
	sqlStr = sqlStr + " '" & callback & "', "			'-- 발신자 번호
	sqlStr = sqlStr + " convert(varchar(4000),N'"& html2db(ttext) &"'), "		'-- 알림톡 내용
	sqlStr = sqlStr + " '" & tmpcd & "', "				'-- 알림톡 템플릿 번호
	sqlStr = sqlStr + " '" & fsendtp & "', "			'-- 알림톡 실패시 문자 형식 > SMS / LMS
	sqlStr = sqlStr + " convert(varchar(50),N'"& html2db(ftit) &"'), "		'-- 실패시 문자 제목 (LMS 전송시에만 필요)
	sqlStr = sqlStr + " convert(varchar(4000),N'"& html2db(ftext) &"'), "		'-- 실패시 문자 내용
	sqlStr = sqlStr + " '" & html2db(btnJson) & "') "	'-- 버튼 구성 내용 (버튼타입에만 필요 / v4 메뉴얼 참고)

	dbget.Execute sqlStr
end Sub

'// 카카오톡 고객센터알림톡 발송. 링크드 SMS 서버에서 발송		' 2021.09.07 한용민 생성
Sub SendKakaoCSMsg_LINK(REQDATE, reqhp,callback,tmpcd,ttext,fsendtp,ftit,ftext,btnJson,TEMPLATE_TITLE,userid)
    dim sqlStr, RetRows

    if callback="" then callback = CNORMALCALLBAKC
    if fsendtp="" then fsendtp="SMS"
    if ftext="" and ttext<>"" then ftext = ttext
	if REQDATE="" or isnull(REQDATE) then
		REQDATE="getdate()"
	else
		REQDATE="N'"& REQDATE &"'"
	end if
	if TEMPLATE_TITLE="" or isnull(TEMPLATE_TITLE) then
		TEMPLATE_TITLE="NULL"
	else
		TEMPLATE_TITLE="N'"& TEMPLATE_TITLE &"'"
	end if
	if userid="" or isnull(userid) then
		userid="NULL"
	else
		userid="N'"& userid &"'"
	end if

    sqlStr = "INSERT INTO [SMSDB].[db_kakaomsg_v4_cs].dbo.KKO_MSG (REQDATE, STATUS, PHONE, CALLBACK, MSG, TEMPLATE_CODE, FAILED_TYPE, FAILED_SUBJECT, FAILED_MSG, BUTTON_JSON, TEMPLATE_TITLE, ETC1)"
	sqlStr = sqlStr & "		SELECT"
	sqlStr = sqlStr & "		"& REQDATE &" as REQDATE, '1' as STATUS"
	sqlStr = sqlStr & "		, '" & reqhp & "' as PHONE"		' 수신자 휴대폰 번호
	sqlStr = sqlStr & "		, '" & callback & "' as CALLBACK"	' 발신자 번호
	sqlStr = sqlStr & "		, convert(varchar(4000),N'"& html2db(ttext) &"') as MSG"	' 알림톡 내용
	sqlStr = sqlStr & "		, '" & tmpcd & "' as TEMPLATE_CODE"		' 알림톡 템플릿 번호
	sqlStr = sqlStr & "		, '" & fsendtp & "' as FAILED_TYPE"		' 알림톡 실패시 문자 형식 > SMS / LMS
	sqlStr = sqlStr & "		, convert(varchar(50),N'"& html2db(ftit) &"') as FAILED_SUBJECT"      ' 실패시 문자 제목 (LMS 전송시에만 필요)
	sqlStr = sqlStr & "		, convert(varchar(4000),N'"& html2db(ftext) &"') as FAILED_MSG"		' 실패시 문자 내용
	sqlStr = sqlStr & "		, N'" & html2db(btnJson) & "' as BUTTON_JSON"		' 버튼 구성 내용 (버튼타입에만 필요 / v4 메뉴얼 참고)
	sqlStr = sqlStr & "		, "& TEMPLATE_TITLE &" as [TEMPLATE_TITLE]"
	sqlStr = sqlStr & "		, "& userid &""

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr
end Sub

%>