<%

CLASS MailCls

	dim MailTitles		'메일 제목
	dim MailConts		'메일 내용 			(text/html)
	dim SenderMail		'메일 발송자 주소 	(customer@10x10.co.kr,mailzine@10x10.co.kr)
	dim SenderNm		'메일 발송자이름 	(텐바이텐)

	dim MailType		'템플릿 번호 		([4],5,6,7,8,9)

	dim ReceiverNm		'메일 수신자 이름 	($1)
	dim ReceiverMail	'메일 수신자 주소 	(xxxx@aaa.com..)


	dim AddrType				'메일수집 방식 (event,userid)
	dim arrUserId 				'AddrType ="userid" 일경우 사용

	dim AddrString				'메일주소 수집에 쓰일 정보
	dim EvtCode,EvtGroupCode 	'AddrType ="event" 일경우 사용
	dim MailerMailGubun		' 메일러 자동메일 번호

	dim strQuery 		'이메일 정보 수집 쿼리
	dim EmailDataType	'이메일 정보 수집 방식 (Enum : string - 직접 입력,sql - 쿼리 이용)
	Dim DB_ID 			'선더메일 디비연결 번호 - 고정 (실서버- 4 ; 테스트- 5)


	Private Sub Class_Initialize()
		EvtCode =0
		EvtGroupCode =0
		EmailDataType = "sql"
		MailType = 5
		MailerMailGubun = 2		' 메일러 자동메일 번호 기본발송 2번

		IF application("Svr_Info")="Dev" THEN
			DB_ID = "5" '//(실서버- 4 ; 테스트- 5)
		ELSE
			DB_ID = "4"
		END IF
		SenderMail	= "mailzine@10x10.co.kr"
		SenderNm	= "텐바이텐"

	End Sub

	Private Sub Class_Terminate()

	End Sub

	'// mailer 관련 주소 반환
	Public Function fnMakeMailerQuery()

		dim tmpAddrType , tmpString
		dim tmpVar

		dim tmpQuery

		On Error Resume Next

		tmpAddrType = AddrType
		tmpString = AddrString

		tmpVar = fnReArr(tmpString,",")

		IF tmpAddrType = "userid" THEN '// 아이디 배열 입력

			IF tmpVar = "" Then
				Err.Number = Err.Number - 1
			End IF

			tmpVar = replace(tmpVar,",","','")
			tmpVar = "'" & tmpVar & "'"

            ''에이메일러 변경후 더이상 사용 안함.
			''tmpQuery = " SELECT UMail, UName FROM db_user.dbo.vw_UserMailList WHERE Uid in (" & tmpVar & ")"
			Err.Number = Err.Number - 1  ''오류냄

		ELSEIF tmpAddrType ="string" Then '// 이름 주소 하나만 처리

			IF ReceiverMail="" Then
				Err.Number = Err.Number - 1
			End IF

			EmailDataType ="string"

			tmpQuery = ReceiverMail

		ELSEIF tmpAddrType ="array" Then '// 이름 & 주소 입력 규칙입력
			'// 작업안함(추후추가)
			tmpQuery=""
		ELSE
			'// 작업안함
			tmpQuery=""
		End IF

		fnMakeMailerQuery = tmpQuery

	End Function

	'// cdo 관련 주소 반환

	Public Function fnMakeCdoQuery(byref iArr)

		dim tmpAddrType , tmpString

		dim tmpVar , tmpVar2 , intLp

		dim tmpQuery,tmpArrList()

		On Error Resume Next

		tmpAddrType = AddrType
		tmpString = AddrString

		tmpVar = fnReArr(tmpString,",")

		IF tmpAddrType = "userid" THEN '// 아이디 배열 입력

			IF tmpVar = "" Then
				Err.Number = Err.Number - 1
			End IF

			tmpVar = replace(tmpVar,",","','")
			tmpVar = "'" & tmpVar & "'"

			tmpQuery = " SELECT UMail, UName FROM db_user.dbo.vw_UserMailList WHERE Uid in (" & tmpVar & ")"

			rsget.open tmpQuery , dbget , 2
			IF not rsget.eof Then
				iArr = rsget.getRows()
			End IF
			rsget.close

		ELSEIF tmpAddrType ="string" Then '// 이름 주소 하나만 처리

			IF ReceiverMail="" Then
				Err.Number = Err.Number - 1
			End IF

			Redim iArr(1,0)
			iArr(0,0) = ReceiverMail
			iArr(1,0) = ReceiverNm

		ELSEIF tmpAddrType ="array" Then '// 이름 & 주소 입력 규칙입력
			IF tmpVar = "" Then
				Err.Number = Err.Number - 1
			End IF
			tmpVar = fnReArr(tmpVar,",")
			tmpVar = Split(tmpVar,",")

			IF isArray(tmpVar) Then

				Redim iArr(1,Ubound(tmpVar))
				For intLp=0 To Ubound(tmpVar)
					tmpVar2 = tmpVar(intLp)

					IF instr(tmpVar2,"$")>0 Then
						iArr(0,intLp) = Left(tmpVar2,instr(tmpVar2,"$")-1)
						iArr(1,intLp) = Right(tmpVar2,len(tmpVar2)-instr(tmpVar2,"$"))
					ELSE
						iArr(0,intLp) = tmpVar2
						iArr(1,intLp) = ""
					End IF
				Next
			End IF
		End IF
		IF Err.Number=0 Then
			fnMakeCdoQuery = 0
		ELSE
			fnMakeCdoQuery = -1
		End IF


	End Function

	Public Function fnReArr(byval strVar,byval strChk)

		'// 구분자로 넘어온 값 strChk 체크후 반환
		'// 반환된 값은 "," 로 구분됨

		dim tmpVar , tmpArrVar , intLp

		IF strVar="" or strChk="" Then '넘어온 값 체크 (없거나 잘못된 값이 넘어오면 끝내기)
			Exit Function
		ELSE
			tmpArrVar = trim(strVar)
			tmpArrVar = split(tmpArrVar,strChk)

			IF Ubound(tmpArrVar) < 0 Then Exit Function

			For intLp=0 to Ubound(tmpArrVar)
				IF tmpArrVar(intLp)<>"" Then
					tmpVar = (tmpVar & tmpArrVar(intLp) & ",")
				END IF
			Next
			tmpVar = Left(tmpVar,Len(tmpVar)-1)
		END IF
		fnReArr = tmpVar

	End Function

	'//+++	메일 템플릿 불러오기 	+++//
	Public Function getMailTemplate()

		dim mFileNm
		dim dfPath
		dim fso,ffso,fnHTML

		'/* 파일 선택 */
		'// MailType - 5 이상 실제 사용 (관계자외 접근/수정 금지! ㅡ.ㅡㅋ )
		IF MailType ="5" Then '// 식사용자정의 양식 메일
			mFileNm =""
		ELSEIF MailType="6" Then 		'// 주문접수
			mFileNm ="mail_a01.htm"
		ELSEIF MailType ="7" Then '// 결제확인
			mFileNm ="mail_a02.htm"
		ELSEIF MailType ="8" Then '// 출고메일
			mFileNm = "mail_delivery2011.htm"
		ELSEIF MailType ="9" Then '// 무통장자동취소안내
			mFileNm ="mail_a04.htm"

		ELSEIF MailType ="10" Then '// 기타CS출고발송
			mFileNm ="mail_b01.htm"
		ELSEIF MailType ="11" Then '// 주문취소(환불안내)
			mFileNm ="mail_b02.htm"
		ELSEIF MailType ="12" Then '// 반품접수
			mFileNm ="mail_b03.htm"
		ELSEIF MailType ="13" Then '// 반품완료(환불안내)
			mFileNm ="mail_b04.htm"
		ELSEIF MailType ="14" Then '// 환불/카드취소완료
			mFileNm ="mail_b05.htm"

		ELSEIF MailType ="15" Then '// 1:1상담 답변
			mFileNm ="mail_c01.htm"
		ELSEIF MailType ="16" Then '// 상품Q&A 답변
			mFileNm ="mail_c02.htm"

		ELSEIF MailType ="17" Then '// 일반 공지 메일
			mFileNm ="mail_d01.htm"
		ELSEIF MailType ="18" Then '// 상품평작성안내
			mFileNm ="mail_d02.htm"
		ELSEIF MailType ="19" Then '// 회원등급공지
			mFileNm ="mail_d03.htm"
		ELSEIF MailType ="20" Then '// 이벤트당첨공지
			mFileNm ="mail_d06.htm"
		ELSEIF MailType ="21" Then '// 비밀번호재발송메일
			mFileNm ="mail_d07.htm"
		ELSEIF MailType ="22" Then '// 출고지연메일
			mFileNm ="mail_misend.htm"
		End IF

		IF MailType<>"5" and mFileNm="" Then
			response.write "템플릿 불러오기 실패"
			Exit Function
		End IF

		'//실섭,테섭구분
		IF application("Svr_Info")="Dev" THEN
			'dfPath = "C:\testweb\admin2009scm\lib\email\mailtemplate" 		'// 테섭(scm)
			dfPath = Server.MapPath("\lib\email\mailtemplate")
		ELSE
		    dfPath = Server.MapPath("\lib\email\mailtemplate")
			''dfPath = "E:\home\cube1010\admin2009scm\lib\email\mailtemplate" 	'// 실섭(scm)
		END IF

		'/* 파일 불러오기 */
		IF mFileNm<>"" Then
			Set fso = server.CreateObject("Scripting.FileSystemObject")
				IF fso.FileExists(dfPath & "\" & mFileNm) then
					set ffso = fso.OpenTextFile(dfPath & "\" & mFileNm,1)
					fnHTML = ffso.ReadAll
					ffso.close
					set ffso = nothing
				ELSE
					fnHTML = ""
				End IF
			Set fso = nothing
		End IF
		getMailTemplate = fnHTML

	End Function

    '//+++	TMS메일러 메일발송	' 2020.09.29 한용민 생성
    Public Function Send_TMSMailer()
        Dim sqlStr

		'response.write MailerMailGubun & "<br>"
		'response.write replace(ReceiverMail,"'","") & "<br>"
		'response.write replace(MailTitles,"'","") & "<br>"
		'response.write newhtml2db(MailConts) & "<br>"
		'response.end

        IF (AddrType<>"string") or (ReceiverMail="") Then '// 이름 주소 하나만 처리
		    Err.Number = Err.Number - 1
        ENd IF

        sqlStr =  " exec db_cs.dbo.usp_TEN_TMS_SendAutoMail '"&replace(ReceiverMail,"'","")&"','','"&replace(MailTitles,"'","")&"','"&newhtml2db(MailConts)&"',"& MailerMailGubun &""
        dbget.Execute sqlStr
    end Function

    '//+++	EMS 에이 메일러 메일발송 2014/04/28	+++//
    Public Function Send_Mailer()
        Dim sqlStr

		'response.write MailerMailGubun & "<br>"
		'response.write replace(ReceiverMail,"'","") & "<br>"
		'response.write replace(MailTitles,"'","") & "<br>"
		'response.write newhtml2db(MailConts) & "<br>"
		'response.end

        IF (AddrType<>"string") or (ReceiverMail="") Then '// 이름 주소 하나만 처리
		    Err.Number = Err.Number - 1
        ENd IF

        sqlStr =  " exec db_cs.[dbo].[sp_Ten_SendAutoMail_Amailer] '"&replace(ReceiverMail,"'","")&"','','"&replace(MailTitles,"'","")&"','"&newhtml2db(MailConts)&"',"& MailerMailGubun &""
        dbget.Execute sqlStr
    end Function

	'//+++	썬더메일 메일발송 	+++//
	Public Function Send_Mailer_OLD_THUNDER()

		On Error Resume Next

		Dim MailDbConn
		Set MailDbConn = Server.CreateObject("ADODB.Connection")
			MailDbConn.Open "DSN=ThunderDB"

		Dim strSQL

		strQuery = fnMakeMailerQuery()

		IF strQuery="" Then
			response.write "대상자가 존재하지 않습니다"
		End IF

		strQuery = replace(strQuery,"'","''")

'response.write MailConts
'dbget.close()	:	response.End
		strSQL= strSQL &_

			" INSERT INTO event_dbevent ( " &_
			" 	title, content " &_
			" 	, sender, sender_alias ,receiver_alias " &_
			"	, content_type, event_id, user_info " &_
			"	, email_insert_type, wasSended, email_data_type, email_sql, db_id) " &_
			" VALUES ( "&_
			" 	'" & MailTitles & "' , '" & newhtml2db(MailConts) & "' " &_
			" 	,'" & SenderMail & "' , '" & SenderNm & "' , '" & ReceiverNm & "' " &_
			" 	,'text/html', '" & MailType & "', '"& strQuery & "' " &_
			" 	,'new', 'X', '"& EmailDataType &"', '" & strQuery & "', '"&DB_ID&"'" &_
			" ) "
''response.write strSQL

		MailDbConn.beginTrans
		MailDbConn.execute(strSQL)


		IF Err.Number =0 THEN
			MailDbConn.CommitTrans
			response.write "메일 발송 성공_Mailer<br>"
		ELSE
			MailDbConn.RollBackTrans
			response.write "메일 발송 실패_Mailer<br>"
		END IF

		MailDbConn.close

        On Error Goto 0
	End Function

	'//+++	외부 서버 메일발송 	+++//

	Public Function Send_CDO()
		dim ArrMailList,intP,ret

		ret = fnMakeCdoQuery(ArrMailList)

		IF ret < 0 Then
			response.write "주소 처리 에러"
			Exit Function
		End IF

		dim cdoMessage,cdoConfig

		'On Error Resume Next

		IF isArray(ArrMailList) Then
			For intP=0 To Ubound(ArrMailList,2)
				Set cdoConfig = Server.CreateObject("CDO.Configuration")
				'-> 서버 접근방법을 설정합니다
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)
				'-> 서버 주소를 설정합니다(dns or ip)-(localhost or 110.93.128.94)
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="110.93.128.94"
				'-> 접근할 포트번호를 설정합니다
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'-> 접속시도할 제한시간을 설정합니다
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 5
				'-> SMTP 접속 인증방법을 설정합니다
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				'-> SMTP 서버에 인증할 ID를 입력합니다
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"
				'-> SMTP 서버에 인증할 암호를 입력합니다
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"
				cdoConfig.Fields.Update

				Set cdoMessage = CreateObject("CDO.Message")
				Set cdoMessage.Configuration = cdoConfig


				cdoMessage.To 		= ArrMailList(1,intP) &"<"& ArrMailList(0,intP) &">"
				cdoMessage.From 	= SenderNm &"<"& SenderMail &">"
				cdoMessage.SubJect 	= MailTitles

				'메일 내용이 텍스트일 경우 cdoMessage.TextBody, html일 경우 cdoMessage.HTMLBody
				cdoMessage.HTMLBody	= MailConts

				cdoMessage.BodyPart.Charset="ks_c_5601-1987"		'// 한글을 위해선 꼭 넣어 주어야 합니다.
				cdoMessage.HTMLBodyPart.Charset="ks_c_5601-1987"	'// 한글을 위해선 꼭 넣어 주어야 합니다.
				'cdoMessage.BodyPart.Charset="utf-8"		'// 한글을 위해선 꼭 넣어 주어야 합니다.
				'cdoMessage.HTMLBodyPart.Charset="utf-8"	'// 한글을 위해선 꼭 넣어 주어야 합니다.

				IF (application("Svr_Info")	= "Dev") then
				    if ((InStr(ArrMailList(0,intP),"10x10.co.kr")>0) or (ArrMailList(0,intP)="kobula@nate.com") or (ArrMailList(0,intP)="tozzinet@daum.net")) then
    		            cdoMessage.Send
                    end if
                ELSE
				    cdoMessage.Send
                END IF

				Set cdoMessage = nothing
				Set cdoConfig = nothing
			Next
		End IF

		IF Err.Number =0 THEN
			response.write "메일 발송 성공_Send_CDO<br>"
		ELSE
			response.write "메일 발송 실패_Send_CDO<br>"
		END IF

	End Function


	'//+++	내부 서버 메일발송 	+++//
	Public Function Send_CDONT()

	 	dim ArrMailList,intP,ret

		IF ReceiverMail="" THEN
			Exit Function
		END IF

		ret = fnMakeCdoQuery(ArrMailList)

		IF ret < 0 Then
			response.write "주소 처리 에러"
			Exit Function
		End IF

		dim oCDONT

        'On Error Resume Next

		IF isArray(ArrMailList) Then
			For intP=0 To Ubound(ArrMailList,2)
				Set oCDONT=Server.CreateObject("CDONTS.NewMail")
				oCDONT.to 		= ArrMailList(1,intP) &"<"& ArrMailList(0,intP) &">"
				oCDONT.from 	= SenderNm &"<"& SenderMail &">"
				oCDONT.subject 	= MailTitles
				'html style
				oCDONT.bodyformat = 0
				oCDONT.mailformat = 0

				oCDONT.body = MailConts
				oCDONT.send
				Set oCDONT = Nothing
			Next
		End IF

		IF Err.Number =0 THEN
			MailDbConn.CommitTrans
			response.write "메일 발송 성공_Send_CDONT<br>"
		ELSE
			MailDbConn.RollBackTrans
			response.write "메일 발송 실패_Send_CDONT<br>"
		END IF
	End Function

End CLASS

'// 단순 메일 발송 선택
Function fnSendMail(mailto,title,contents)

	Dim objMail

	Set objMail = New MailCls

	objMail.AddrType="string"
	objMail.ReceiverMail = mailto
	objMail.MailTitles = title
	objMail.MailConts = contents
	objMail.Send_CDO()

	Set objMail = Nothing

End Function

Function fnSendMail_Mailer(mailto,title,contents)

	Dim objMail

	Set objMail = New MailCls

	objMail.AddrType="string"
	objMail.ReceiverMail = mailto
	objMail.MailTitles = title
	objMail.MailConts = contents
	objMail.MailerMailGubun = 2		' 메일러 자동메일 번호
	objMail.Send_TMSMailer()		'TMS메일러
	'objMail.Send_Mailer()

	Set objMail = Nothing
End Function


%>
