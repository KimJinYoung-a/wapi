<%
'CI 값으로 텐바이텐 회원인지 검사
Function isMemberToCI(ici, jmember_id)
	Dim strSql, chkMember, isHoldRevive
	strSql = ""
	strSql = strSql & " SELECT TOP 1 userid FROM db_user.dbo.tbl_user_n WHERE connInfo = '"&ici&"' "
	rsget.Open strSql, dbget, 1
	If rsget.RecordCount > 0 Then
		chkMember = True
		jmember_id = rsget("userid")
	Else
		chkMember = False
	End If
	rsget.Close

	If chkMember = False Then
		'1.뷰쿼리해서 있으면 chkMember를 True로 바꾸고 jmember_id도 그 뷰의 ID를 넣자.
		strSql = ""
		strSql = strSql & " SELECT TOP 1 userid FROM [db_user_Hold].[dbo].[vw_Hold_user_BasicInfo] WHERE connInfo = '"&ici&"' "
		rsget.Open strSql, dbget, 1
		If rsget.RecordCount > 0 Then
			chkMember = True
			jmember_id = rsget("userid")
		End If
		rsget.Close
		'2.팀장님이 준 휴면고객 원복하기
		If chkMember = True AND jmember_id <> "" Then
			strSql = "db_user_hold.dbo.sp_Ten_HoldUserRevive ('" & jmember_id & "','S')"
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			if Not rsget.Eof then
				isHoldRevive = "Y"
			end if
			rsget.Close
		End If
	End If
	isMemberToCI = chkMember
End Function

'로그인 유도 JSON 데이터 생성
Function fnLoginJsonFlush(iseq_trans, icust_name, ihp_num, ierrCode)
	Dim oJson
	If ierrCode = "0000" Then
		If (iseq_trans = "") OR (icust_name = "") OR (ihp_num = "") Then
			ierrCode = "2101"
		End If
	End If

	Set oJson = jsObject()
		oJson("cd_fulltext") = "1121"
		oJson("cd_partner") = "971"
		oJson("seq_trans") = ""&iseq_trans&""
		oJson("cd_encryption") = "00"
		oJson("cd_encryption_key") = "1"
		oJson("version") = "0031"
		oJson("cd_response") = ""&ierrCode&""
		Set oJson("data") = jsObject()
			'oJson("data")("cust_name")= ""&icust_name&""
			oJson("data")("cust_name")= ""&Server.URLEncode(icust_name)&""
			oJson("data")("hp_num")= ""&ihp_num&""
			oJson("data")("cust_num")= ""
			oJson("data")("card_num")= ""
			oJson("data")("card_grade")= ""
			oJson("data")("card_pin_number")= ""
			oJson("data")("flag_payment")= "Y"				'이걸 Y로 보내야 로그인 창이 뜬다고 함
			oJson("data")("flag_complete")= "N"				'회원가입 완료여부인데 이건 시럽을 통해 회원가입하면 Y로 하라함.
			oJson("data")("coupon_count")= "0"
			Set oJson("data")("coupon_data") = jsArray()
				Set oJson("data")("coupon_data")(null) = jsObject()
					oJson("data")("coupon_data")(null)("coupon_id") = ""
					oJson("data")("coupon_data")(null)("coupon_num") = ""
					oJson("data")("coupon_data")(null)("sdate") = ""
					oJson("data")("coupon_data")(null)("edate") = ""
		oJson.Flush
	Set oJson = nothing
End Function

Function fnShopUserJsonFlushProc(iseq_trans, imember_id, icust_name, ihp_num, ici, ierrCode)
	Dim strsql, oJson
	Dim newCardNo, newUserSeq
	Dim rsUserid, rsCardNo, rsUserseq, rsUserlevel, tenMylevel, regCardCnt
	Dim regidCnt
	strsql = ""
	strsql = strsql & " SELECT TOP 1 n.userid, sc.CardNo, su.userseq, IsNULL(l.userlevel,5) as userlevel "
	strsql = strsql & " FROM db_user.dbo.tbl_user_n as n "
	strsql = strsql & " JOIN db_shop.dbo.tbl_total_shop_user as su on n.userid = su.onlineUserID "
	strsql = strsql & " JOIN db_shop.dbo.tbl_total_shop_card as sc on su.userseq = sc.userseq and sc.useYN = 'Y' "
	strsql = strsql & " JOIN [db_user].[dbo].[tbl_logindata] as l on n.userid = l.userid "
	strsql = strsql & " WHERE n.userid = '"&imember_id&"' "
	strsql = strsql & " ORDER BY sc.Regdate DESC "
	rsget.Open strSql, dbget, 1
	If rsget.RecordCount > 0 Then
		rsUserid	= rsget("userid")
		rsCardNo	= rsget("CardNo")
		rsUserseq	= rsget("userseq")
		rsUserlevel	= levelChange(rsget("userlevel"))
	Else
		rsUserid	= ""
	End If
	rsget.Close

	If rsUserid <> "" Then					'회원이면서 실제 카드가 있는 경우
		If ierrCode = "0000" Then
			If (iseq_trans = "") OR (icust_name = "") OR (ihp_num = "") OR (rsUserseq = "") OR (rsUserlevel = "") Then
				ierrCode = "2101"
			End If
		End If

		'2015-07-20 김진영 lastQueDate에 실행 날짜 로그 입력
		strsql = ""
		strsql = strsql & " UPDATE db_shop.dbo.tbl_total_shop_user SET " 
		strsql = strsql & " lastQueDate = getdate() "
		strsql = strsql & " WHERE OnlineUserID = '"&imember_id&"' "
		dbget.Execute strsql, 1

		Set oJson = jsObject()
			oJson("cd_fulltext") = "1121"
			oJson("cd_partner") = "971"
			oJson("seq_trans") = ""&iseq_trans&""
			oJson("cd_encryption") = "00"
			oJson("cd_encryption_key") = "1"
			oJson("version") = "0031"
			oJson("cd_response") = ""&ierrCode&""

			Set oJson("data") = jsObject()
				'oJson("data")("cust_name")= ""&icust_name&""
				oJson("data")("cust_name")= ""&Server.URLEncode(icust_name)&""
				oJson("data")("hp_num")= ""&ihp_num&""
				oJson("data")("cust_num")= ""&rsUserseq&""
				oJson("data")("card_num")= ""&rsCardNo&""
				oJson("data")("card_grade")= ""&rsUserlevel&""
				oJson("data")("card_pin_number")= ""
				oJson("data")("flag_payment")= "N"
				oJson("data")("flag_complete")= "N"
				oJson("data")("coupon_count")= "0"
				Set oJson("data")("coupon_data") = jsArray()
					Set oJson("data")("coupon_data")(null) = jsObject()
						oJson("data")("coupon_data")(null)("coupon_id") = ""
						oJson("data")("coupon_data")(null)("coupon_num") = ""
						oJson("data")("coupon_data")(null)("sdate") = ""
						oJson("data")("coupon_data")(null)("edate") = ""
			oJson.Flush
		Set oJson = nothing
	Else									'회원이면서 카드가 없는 경우
		On Error Resume Next
		dbget.beginTrans
			'1.카드번호 생성 프로시저를 통해 카드번호 생성
			strsql = ""
			strsql = strsql & " exec [db_shop].[dbo].[sp_ten_getSyrupCardNo] "
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.Open strsql, dbget
			If (Not rsget.Eof) then
				newCardNo = rsget("CardNo")
			End If
			rsget.close

			'2.카드 번호를 넘어온 UserID에 입력 => db_shop.dbo.tbl_total_shop_user이것은 user_n 데이터와 똑같다 생각하자
			strsql = ""
			strsql = strsql & " INSERT INTO db_shop.dbo.tbl_total_shop_user (username, jumin1, HpNo, Email, EmailYN, SMSYN, RegShopID, lastupdate, regdate, OnlineUserID, lastQueDate) " & VBCRLF
			strsql = strsql & " SELECT TOP 1 username, LEFT(juminno, 6), usercell, usermail, emailok, smsok, 'syrup', getdate(), getdate(), userid, getdate() " & VBCRLF
			strsql = strsql & " FROM db_user.dbo.tbl_user_n " & VBCRLF
			strsql = strsql & " WHERE userid = '"&imember_id&"' " & VBCRLF
			dbget.Execute strSql, 1

			'3.db_shop.dbo.tbl_total_shopcard에 카드 저장, tbl_total_card_list의 useYN을 Y로 수정
			strsql = ""
			strsql = strsql & " SELECT TOP 1 UserSeq FROM db_shop.dbo.tbl_total_shop_user WHERE OnlineUserID = '"&imember_id&"' "
			rsget.Open strsql, dbget, 1
			If Not Rsget.Eof Then
				newUserSeq = rsget("UserSeq")
			End If
			rsget.close

			If (newUserSeq <> "") Then
				strsql = ""
				strsql = strsql & " INSERT INTO db_shop.dbo.tbl_total_shop_card (UserSeq, CardNo, point, useYN, RegShopID, Regdate) VALUES " & VBCRLF
				strsql = strsql & " ('"&newUserSeq&"', '"&newCardNo&"', 0, 'Y', 'syrup', getdate()) " & VBCRLF
				dbget.Execute strSql, 1
	
				strsql = ""
				strsql = strsql & " UPDATE db_shop.dbo.tbl_total_card_list SET " & VBCRLF
				strsql = strsql & " useYN = 'Y' " & VBCRLF
				strsql = strsql & " WHERE cardNo = '"&newCardNo&"' " & VBCRLF
				dbget.Execute strSql, 1
			End If

			tenMylevel = getMyLevel(imember_id)
			tenMylevel = levelChange(tenMylevel)

			If Err.Number = 0 Then
			    dbget.CommitTrans
			Else
			    dbget.RollBackTrans
			    ierrCode = "2101"
			End If
		On error Goto 0

		'4.JSON 데이터 Flush
		Set oJson = jsObject()
			oJson("cd_fulltext") = "1121"
			oJson("cd_partner") = "971"
			oJson("seq_trans") = ""&iseq_trans&""
			oJson("cd_encryption") = "00"
			oJson("cd_encryption_key") = "1"
			oJson("version") = "0031"
			oJson("cd_response") = ""&ierrCode&""
			Set oJson("data") = jsObject()
				'oJson("data")("cust_name")= ""&icust_name&""
				oJson("data")("cust_name")= ""&Server.URLEncode(icust_name)&""
				oJson("data")("hp_num")= ""&ihp_num&""
				oJson("data")("cust_num")= ""&newUserSeq&""
				oJson("data")("card_num")= ""&newCardNo&""
				oJson("data")("card_grade")= ""&tenMylevel&""
				oJson("data")("card_pin_number")= ""
				oJson("data")("flag_payment")= "N"
				oJson("data")("flag_complete")= "N"
				oJson("data")("coupon_count")= "0"
				Set oJson("data")("coupon_data") = jsArray()
					Set oJson("data")("coupon_data")(null) = jsObject()
						oJson("data")("coupon_data")(null)("coupon_id") = ""
						oJson("data")("coupon_data")(null)("coupon_num") = ""
						oJson("data")("coupon_data")(null)("sdate") = ""
						oJson("data")("coupon_data")(null)("edate") = ""
			oJson.Flush
		Set oJson = nothing
	End If
End Function

'시럽에서 텐바이텐 신규 가입
Function fnJoin10x01Flush(iseq_trans, ibirthday, iagree_count, icd_sex, icust_name, imember_pw, ijumin_num, ihp_num, imember_id, ici, ismsok, ierrCode)
	Dim strsql, oJson
	Dim juminno, birthday, txCellNum, sitegubun, refip, txSolar, email_10x10
	Dim Enc_userpass : Enc_userpass = MD5(CStr(imember_pw))
	Dim Enc_userpass64 : Enc_userpass64 = SHA256(MD5(CStr(imember_pw)))
	Dim newCardNo, newUserSeq, exeCnt, AssignedRow, regCardCnt
	Dim sexCode

	juminno		= Left(ijumin_num, 6) & "-" & Right(ijumin_num, 1) & "000000"
	birthday	= Left(ibirthday, 4) & "-" & Mid(ibirthday, 5, 2) & "-" & Right(ibirthday, 2)
	If Len(ihp_num) < 11 Then
		txCellNum = Left(ihp_num, 3) & "-" & Mid(ihp_num, 4, 3) & "-" & Right(ihp_num, 4)
	Else
		txCellNum = Left(ihp_num, 3) & "-" & Mid(ihp_num, 4, 4) & "-" & Right(ihp_num, 4)
	End If
	sitegubun = "10x10"
	refip = Left(request.ServerVariables("REMOTE_ADDR"),32)
	txSolar = "Y"
	exeCnt = 0
	sexCode = Right(ijumin_num, 1)
	On Error Resume Next

	'################# 같은 명의의 두개의 단말기로 회원가입창까지 왔을 때 처리 ##################
	Dim regedID, isCI
	strsql = ""
	strsql = strsql& " SELECT TOP 1 userid FROM db_user.dbo.tbl_user_n WHERE connInfo = '"&ici&"'  "
	rsget.Open strSql, dbget, 1
	If rsget.RecordCount > 0 Then
		isCI = True
		regedID = rsget("userid")
	Else
		isCI = False
	End If
	rsget.Close
	'################# 같은 명의의 두개의 단말기로 회원가입창까지 왔을 때 처리 끝 ###############

	If isCI Then
		Call fnShopUserJsonFlushProc(iseq_trans, regedID, icust_name, ihp_num, ici, errCode)
	Else
		dbget.beginTrans
			strsql = ""
			strsql = strsql & " INSERT INTO [db_user].[dbo].tbl_user_n(userid, username, juminno, birthday, zipcode, useraddr, usercell, usermail, regdate, mileage,  userlogo, usercomment, emailok, eventid, sitegubun, email_10x10, email_way2way, refip, issolar, smsok, smsok_fingers, sexflag, jumin1, Enc_jumin2, userStat, isMobileChk, rdsite, realnamecheck, connInfo) " & VBCRLF
			strsql = strsql & " VALUES ('"&imember_id&"', '"&icust_name&"', '"&CStr(juminno)&"', '" + CStr(birthday) + "', '','','"&txCellNum&"','', getdate(), 0,  '', '','N','','"&sitegubun&"','N','N','"&refip&"', '"&txSolar&"', '"&ismsok&"', 'N', '"&sexCode&"', '"&LEFT(ijumin_num, 6)&"', '', 'Y', 'Y', 'syrup', 'Y', '"&ici&"')" & VBCRLF
			dbget.Execute strsql, AssignedRow
			If AssignedRow = 1 Then exeCnt = 1
	
			strsql = ""
			strsql = strsql & " INSERT INTO [db_user].[dbo].tbl_logindata(userid, userpass, userdiv, lastlogin, counter, lastrefip, Enc_userpass, Enc_userpass64) " & VBCRLF
			strsql = strsql & " VALUES ('"&imember_id&"', '', '01', getdate(), 0, '"&refip&"', '"&Enc_userpass&"', '"&Enc_userpass64&"')" & VBCRLF
			dbget.Execute strsql, AssignedRow
			If AssignedRow = 1 Then exeCnt = exeCnt + 1
	
			strsql = ""
			strsql = strsql & " INSERT INTO [db_user].[dbo].tbl_user_current_mileage(userid,bonusmileage)" & VBCRLF
			strsql = strsql & " VALUES('" & imember_id & "'," & VBCRLF
			strsql = strsql & " '0'" & VBCRLF
			strsql = strsql & ")"
			dbget.Execute strsql, AssignedRow
			If AssignedRow = 1 Then exeCnt = exeCnt + 1
	
			'사이트별 사용 구분 입력
			strsql = ""
			strsql = strsql & " INSERT INTO db_user.dbo.tbl_user_allow_site" & VBCRLF
			strsql = strsql & " (userid, sitegubun, siteusing, allowdate)" & VBCRLF
			strsql = strsql & " VALUES(" & VBCRLF
			strsql = strsql & " '" & imember_id & "'" & VBCRLF
			strsql = strsql & " ,'10x10'" & VBCRLF
			strsql = strsql & " ,'Y'" & VBCRLF
			strsql = strsql & " ,getdate()" & VBCRLF
			strsql = strsql & " )" & VBCRLF
			dbget.Execute strsql, AssignedRow
			If AssignedRow = 1 Then exeCnt = exeCnt + 1
	
			strsql = ""
			strsql = strsql & " INSERT INTO db_user.dbo.tbl_user_allow_site" & VBCRLF
			strsql = strsql & " (userid, sitegubun, siteusing, allowdate)" & VBCRLF
			strsql = strsql & " VALUES(" & VBCRLF
			strsql = strsql & " '" & imember_id & "'" & VBCRLF
			strsql = strsql & " ,'academy'" & VBCRLF
			strsql = strsql & " ,'Y'" & VBCRLF
			strsql = strsql & " ,getdate()" & VBCRLF
			strsql = strsql & " )" & VBCRLF
			dbget.Execute strsql, AssignedRow
			If AssignedRow = 1 Then exeCnt = exeCnt + 1
	
			strsql = ""
			strsql = strsql & " exec [db_shop].[dbo].[sp_ten_getSyrupCardNo] "
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.Open strsql, dbget
			If (Not rsget.Eof) then
				newCardNo = rsget("CardNo")
			End If
			rsget.close
	
			strsql = ""
			strsql = strsql & " INSERT INTO db_shop.dbo.tbl_total_shop_user (username, jumin1, HpNo, Email, EmailYN, SMSYN, RegShopID, lastupdate, regdate, OnlineUserID, connInfo, lastQueDate) " & VBCRLF
			strsql = strsql & " VALUES ('"&icust_name&"', '"&LEFT(ijumin_num, 6)&"', '"&txCellNum&"', '', 'N', '"&ismsok&"', 'syrup', getdate(), getdate(), '"&imember_id&"', '"&ici&"', getdate()) " & VBCRLF
			dbget.Execute strsql, AssignedRow
			If AssignedRow = 1 Then exeCnt = exeCnt + 1

			strsql = ""
			strsql = strsql & " SELECT TOP 1 UserSeq FROM db_shop.dbo.tbl_total_shop_user WHERE OnlineUserID = '"&imember_id&"' "
			rsget.Open strsql, dbget, 1
			If Not Rsget.Eof Then
				newUserSeq = rsget("UserSeq")
			End If
			rsget.close

			strsql = ""
			strsql = strsql & " INSERT INTO db_shop.dbo.tbl_total_shop_card (UserSeq, CardNo, point, useYN, RegShopID, Regdate) VALUES " & VBCRLF
			strsql = strsql & " ('"&newUserSeq&"', '"&newCardNo&"', 0, 'Y', 'syrup', getdate()) " & VBCRLF
			dbget.Execute strsql, AssignedRow
			If AssignedRow = 1 Then exeCnt = exeCnt + 1
	
			strsql = ""
			strsql = strsql & " UPDATE db_shop.dbo.tbl_total_card_list SET " & VBCRLF
			strsql = strsql & " useYN = 'Y' " & VBCRLF
			strsql = strsql & " WHERE cardNo = '"&newCardNo&"' " & VBCRLF
			dbget.Execute strsql, AssignedRow
			If AssignedRow = 1 Then exeCnt = exeCnt + 1

			strsql = ""
			strsql = strsql& " SELECT TOP 1 userid FROM db_user.dbo.tbl_user_n WHERE connInfo = '"&ici&"' and userid <> '"&imember_id&"' "
			rsget.Open strSql, dbget, 1
			If rsget.RecordCount > 0 Then
				exeCnt = exeCnt - 1
			End If
			rsget.Close

			If (Err.Number = 0) AND (exeCnt = 8) Then
			    dbget.CommitTrans
			Else
			    dbget.RollBackTrans
			    ierrCode = "2101"
			End If
		On error Goto 0
	
		Set oJson = jsObject()
			oJson("cd_fulltext") = "1121"
			oJson("cd_partner") = "971"
			oJson("seq_trans") = ""&iseq_trans&""
			oJson("cd_encryption") = "00"
			oJson("cd_encryption_key") = "1"
			oJson("version") = "0031"
			oJson("cd_response") = ""&ierrCode&""
			Set oJson("data") = jsObject()
				'oJson("data")("cust_name")= ""&icust_name&""
				oJson("data")("cust_name")= ""&Server.URLEncode(icust_name)&""
				oJson("data")("hp_num")= ""&ihp_num&""
				oJson("data")("cust_num")= ""&newUserSeq&""
				oJson("data")("card_num")= ""&newCardNo&""
				oJson("data")("card_grade")= "005"
				oJson("data")("card_pin_number")= ""
				oJson("data")("flag_payment")= "N"
				oJson("data")("flag_complete")= "Y"
				oJson("data")("coupon_count")= "0"
				Set oJson("data")("coupon_data") = jsArray()
					Set oJson("data")("coupon_data")(null) = jsObject()
						oJson("data")("coupon_data")(null)("coupon_id") = ""
						oJson("data")("coupon_data")(null)("coupon_num") = ""
						oJson("data")("coupon_data")(null)("sdate") = ""
						oJson("data")("coupon_data")(null)("edate") = ""
			oJson.Flush
		Set oJson = nothing
	End If
End Function

'포인트 조회
Function fnpointViewFlush(iseq_trans, icust_name, ihp_num, icard_num, icust_num, icard_grade, ierrCode)
	Dim strsql, oJson
	Dim OnLinePoint, OffLinePoint, cardnlevel, myOrgCardno, myLevel, myUserid
	Dim isCardChange : isCardChange = "N"
	Dim isGradeChange : isGradeChange = "N"
	Dim realCardno
	On Error Resume Next
		'시럽에 저장되어있는 카드인지 아니면 오프라인매장에서 카드를 발급받아서 카드 번호가 바뀐건지 검사
		myOrgCardno = tenOrgCardNum(icard_num, myLevel, myUserid)
		If CStr(icard_num) <> CStr(myOrgCardno) Then
			isCardChange = "Y"
			realCardno = myOrgCardno
		Else
			realCardno = icard_num
		End If
		If CInt(icard_grade) <> CInt(myLevel) Then isGradeChange = "Y"

		strsql = ""
		strsql = strsql & " SELECT TOP 1 su.OnlineUserid, sc.cardNo "
		strsql = strsql & " ,(m.jumunmileage +  m.flowerjumunmileage + m.bonusmileage  + m.academymileage - m.spendmileage -  IsNULL(m.expiredMile,0) - IsNULL(m.michulmile,0) - IsNULL(m.michulmileACA,0))  as OnLinePoint "
		strsql = strsql & " ,IsNull(sc.Point,0) AS OffLinePoint "
		strsql = strsql & " FROM db_shop.dbo.tbl_total_shop_user as su "
		strsql = strsql & " JOIN db_shop.dbo.tbl_total_shop_card as sc on su.userseq = sc.userseq "
		strsql = strsql & " JOIN [db_user].[dbo].tbl_user_current_mileage m on su.OnlineUserid = m.userid "
		strsql = strsql & " WHERE sc.useYN = 'Y' "
		strsql = strsql & " and sc.cardNo = '"&realCardno&"' "
		strsql = strsql & " ORDER BY sc.Regdate DESC "
		rsget.Open strsql, dbget, 1
		If (Not rsget.Eof) then
			OnLinePoint		= rsget("OnLinePoint")
			OffLinePoint	= rsget("OffLinePoint")
		End If
		rsget.close

		If OnLinePoint = "" Then OnLinePoint = 0
		If OffLinePoint = "" Then OffLinePoint = 0

		If Err.Number <> 0 Then
		    ierrCode = "3101"
		End If
	On error Goto 0

	Set oJson = jsObject()
		oJson("cd_fulltext") = "1202"
		oJson("cd_partner") = "971"
		oJson("seq_trans") = ""&iseq_trans&""
		oJson("cd_encryption") = "00"
		oJson("cd_encryption_key") = "1"
		oJson("version") = "0031"
		oJson("cd_response") = ""&ierrCode&""
		Set oJson("data") = jsObject()
			oJson("data")("card_num")= ""&icard_num&""											'카드번호
			oJson("data")("hp_num")= ""&ihp_num&""												'휴대폰번호
			'oJson("data")("cust_name")= ""&icust_name&""										'고객명
			oJson("data")("cust_name")= ""&Server.URLEncode(icust_name)&""
			oJson("data")("total_accu_point")= ""&OffLinePoint&""								'총 누적포인트
			oJson("data")("now_avail_point")= ""&OnLinePoint&""									'현재 가용포인트
			'oJson("data")("accu_point_name")= "오프라인"										'누적포인트 명
			oJson("data")("accu_point_name")= ""&Server.URLEncode("오프라인")&""
			'oJson("data")("avail_point_name")= "온라인"										'가용포인트 명
			oJson("data")("avail_point_name")= ""&Server.URLEncode("온라인")&""
			oJson("data")("accu_point_unit_name")= "p"											'누적포인트 단위명		| 기본 : '점'
			oJson("data")("avail_point_unit_name")= "p"											'가용포인트 단위명		| 기본 : '점'
			oJson("data")("flg_accu_accu_point")= "Y"											'누적포인트 사용여부
			oJson("data")("flg_avail_accu_point")= "Y"											'가용포인트 사용여부
			oJson("data")("flg_update_grade")= ""&isGradeChange&""								'등급 갱신여부			| N : 기본값, Y : 갱신 시
			oJson("data")("update_grade")= ChkIIF(isGradeChange="Y", ""&myLevel&"", "")			'카드등급 코드(갱신)	| 갱신여부 N일시 값 없음
			oJson("data")("flg_update_card_num")= ""&isCardChange&""							'카드번호 갱신여부		| N : 기본값, Y : 갱신 시
			oJson("data")("update_card_num")= ChkIIF(isCardChange="Y", ""&myOrgCardno&"", "")	'카드 번호(갱신)		| 갱신여부 N일시 값 없음
			oJson("data")("update_card_pin")= ""												'카드 Pin 번호(갱신)	| 갱신여부 N일시 값 없음
		oJson.Flush
	Set oJson = nothing
End Function

'포인트 내역조회
Function fnpointDetailViewFlush(iseq_trans, icust_name, ihp_num, icard_num, icust_num, icard_grade, ierrCode)
	Dim strsql, oJson, i
	Dim OnLinePoint, OffLinePoint, cardnlevel, myOrgCardno, myLevel, myUserid
	Dim isCardChange : isCardChange = "N"
	Dim isGradeChange : isGradeChange = "N"
	Dim realCardno, tmpjumunmileage, tmpacademymileage
	Dim mileageCnt : mileageCnt = "00"
	On Error Resume Next
		'시럽에 저장되어있는 카드인지 아니면 오프라인매장에서 카드를 발급받아서 카드 번호가 바뀐건지 검사
		myOrgCardno = tenOrgCardNum(icard_num, myLevel, myUserid)
		If CStr(icard_num) <> CStr(myOrgCardno) Then
			isCardChange = "Y"
			realCardno = myOrgCardno
		Else
			realCardno = icard_num
		End If
		If CInt(icard_grade) <> CInt(myLevel) Then isGradeChange = "Y"

		strsql = ""
		strsql = strsql & " SELECT TOP 1 su.OnlineUserid, sc.cardNo "
		strsql = strsql & " ,(m.jumunmileage +  m.flowerjumunmileage + m.bonusmileage  + m.academymileage - m.spendmileage -  IsNULL(m.expiredMile,0))  as OnLinePoint "
		strsql = strsql & " ,IsNull(sc.Point,0) AS OffLinePoint "
		strsql = strsql & " ,m.jumunmileage, m.academymileage "
		strsql = strsql & " FROM db_shop.dbo.tbl_total_shop_user as su "
		strsql = strsql & " JOIN db_shop.dbo.tbl_total_shop_card as sc on su.userseq = sc.userseq "
		strsql = strsql & " JOIN [db_user].[dbo].tbl_user_current_mileage m on su.OnlineUserid = m.userid "
		strsql = strsql & " WHERE sc.useYN = 'Y' "
		strsql = strsql & " and sc.cardNo = '"&realCardno&"' "
		strsql = strsql & " ORDER BY sc.Regdate DESC "
		rsget.Open strsql, dbget, 1
		If (Not rsget.Eof) then
			OnLinePoint			= rsget("OnLinePoint")
			OffLinePoint		= rsget("OffLinePoint")
			tmpjumunmileage		= rsget("jumunmileage")
			tmpacademymileage	= rsget("academymileage")
		End If
		rsget.close

		If OnLinePoint = "" Then OnLinePoint = 0
		If OffLinePoint = "" Then OffLinePoint = 0

		If Err.Number <> 0 Then
		    ierrCode = "3101"
		End If
	On error Goto 0
	Set oJson = jsObject()
		oJson("cd_fulltext") = "1212"
		oJson("cd_partner") = "971"
		oJson("seq_trans") = ""&iseq_trans&""
		oJson("cd_encryption") = "00"
		oJson("cd_encryption_key") = "1"
		oJson("version") = "0031"
		oJson("cd_response") = ""&ierrCode&""
		Set oJson("data") = jsObject()
			oJson("data")("card_num")= ""&icard_num&""											'카드번호
			oJson("data")("hp_num")= ""&ihp_num&""												'휴대폰번호
			'oJson("data")("cust_name")= ""&icust_name&""										'고객명
			oJson("data")("cust_name")= ""&Server.URLEncode(icust_name)&""
			oJson("data")("total_accu_point")= ""&OffLinePoint&""								'총 누적포인트
			oJson("data")("now_avail_point")= ""&OnLinePoint&""									'현재 가용포인트
			'oJson("data")("accu_point_name")= "오프라인"										'누적포인트 명
			oJson("data")("accu_point_name")= ""&Server.URLEncode("오프라인")&""
			'oJson("data")("avail_point_name")= "온라인"										'가용포인트 명
			oJson("data")("avail_point_name")= ""&Server.URLEncode("온라인")&""
			oJson("data")("accu_point_unit_name")= "p"											'누적포인트 단위명		| 기본 : '점'
			oJson("data")("avail_point_unit_name")= "p"											'가용포인트 단위명		| 기본 : '점'
			oJson("data")("flg_accu_accu_point")= "Y"											'누적포인트 사용여부
			oJson("data")("flg_avail_accu_point")= "Y"											'가용포인트 사용여부
			oJson("data")("flg_update_grade")= ""&isGradeChange&""								'등급 갱신여부			| N : 기본값, Y : 갱신 시
			oJson("data")("update_grade")= ChkIIF(isGradeChange="Y", ""&myLevel&"", "")			'카드등급 코드(갱신)	| 갱신여부 N일시 값 없음
			oJson("data")("flg_update_card_num")= ""&isCardChange&""							'카드번호 갱신여부		| N : 기본값, Y : 갱신 시
			oJson("data")("update_card_num")= ChkIIF(isCardChange="Y", ""&myOrgCardno&"", "")	'카드 번호(갱신)		| 갱신여부 N일시 값 없음
			oJson("data")("update_card_pin")= ""												'카드 Pin 번호(갱신)	| 갱신여부 N일시 값 없음

			strsql = ""
			strsql = strsql & " SELECT TOP 20 * "
			strsql = strsql & " FROM ( "
			strsql = strsql & " 	SELECT 'OF' as tg, point, convert(Varchar(50), logDesc) as logDesc, orderno, regdate "
			strsql = strsql & " 	FROM db_shop.dbo.tbl_total_shop_log "
			strsql = strsql & " 	WHERE cardno = '"&realCardno&"' "
			strsql = strsql & " 	and DATEDIFF (month ,regdate ,getdate()) <= 6 "
			strsql = strsql & " UNION ALL "
			strsql = strsql & " 	SELECT 'ON', mileage, convert(Varchar(50), jukyo) as jukyo, orderserial, regdate "
			strsql = strsql & " 	FROM [db_user].[dbo].tbl_mileagelog where userid = '"&myUserid&"' and deleteyn='N' "
			strsql = strsql & " 	and DATEDIFF (month ,regdate ,getdate()) <= 6 "
			If tmpjumunmileage > 0 Then
				strsql = strsql & " UNION ALL "
				strsql = strsql & " 	SELECT 'ONORD', totalmileage "
				strsql = strsql & " 	,CASE WHEN totalmileage >= 0 THEN '온라인 적립' ELSE '온라인 사용' END, orderserial, regdate  "
				strsql = strsql & " 	FROM [db_order].[dbo].tbl_order_master "
				strsql = strsql & " 	WHERE userid='"&myUserid&"' and ipkumdiv>3 and cancelyn='N'  and sitename='10x10' "
			End If

			If tmpacademymileage > 0 Then
				strsql = strsql & " UNION ALL "
				strsql = strsql & " 	SELECT 'ACORD', totalmileage"
				strsql = strsql & " 	,CASE WHEN totalmileage >= 0 THEN '아카데미 적립' ELSE '아카데미 사용' END, orderserial, regdate  "
				strsql = strsql & "		FROM [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_master "
				strsql = strsql & "		WHERE userid='"&myUserid&"' and ipkumdiv>3 and cancelyn='N' "
				strsql = strsql & "     and DATEDIFF (month ,regdate ,getdate()) <= 6 "
			End If
			strsql = strsql & " ) TT "
			strsql = strsql & " ORDER BY regdate DESC, point DESC "
			rsget.Open strsql, dbget, 1
			If (Not rsget.Eof) then
				mileageCnt = rsget.RecordCount
			End If
			oJson("data")("deal_count")= ""&mileageCnt&""																	'거래 개수 | 최대 20건 제한
			If mileageCnt > 0 Then
				Set oJson("data")("deal_data") = jsArray()
				For i = 0 to mileageCnt-1
					Set oJson("data")("deal_data")(null) = jsObject()
						oJson("data")("deal_data")(null)("deal_num") = ChkIIf(i+1 < 10, "0"&i+1, ""&i+1&"")					'거래순번 | 일련번호 형태(예:01,02,03, ...) => 문의결과 01이면 앱에서 제일위에, 20이면 제일 아래에 보인다고 함 => 결국 regdate desc로 해야함.
						oJson("data")("deal_data")(null)("cd_deal") = ChkIIf(Clng(rsget("point")) >= 0, "M1", "M2") 			'거래 구분 코드 | M1:적립, M2:사용
						oJson("data")("deal_data")(null)("deal_date") = FormatDate(rsget("regdate"), "00000000000000")		'거래 일시 | 형식 : YYYYMMDDhhmmss
						If rsget("tg") = "OF" Then
							oJson("data")("deal_data")(null)("deal_shop") = ""&Server.URLEncode("오프라인 "&rsget("logDesc"))&""
						ElseIf rsget("tg") = "ON" Then
							oJson("data")("deal_data")(null)("deal_shop") = ""&Server.URLEncode("온라인 "&rsget("logDesc"))&""
						Else
							oJson("data")("deal_data")(null)("deal_shop") = ""&Server.URLEncode(rsget("logDesc"))&""
						End If
						oJson("data")("deal_data")(null)("deal_point") = replace(rsget("point"), "-", "")					'거래 포인트
					rsget.moveNext
				Next
			End If
			rsget.close
		oJson.Flush
	Set oJson = nothing
End Function

'텐바이텐 level
Function getMyLevel(ilevel)
	Dim strSql, tmplevel
	strSql = ""
	strSql = strSql & " SELECT IsNULL(userlevel,5) as userlevel FROM [db_user].[dbo].[tbl_logindata] WHERE userid = '"&ilevel&"' "
	rsget.Open strsql, dbget, 1
	If (Not rsget.Eof) then
		tmplevel = rsget("userlevel")
	End If
	rsget.close
	getMyLevel = tmplevel
End Function

'실제 사용되는 카드번호와 회원등급 얻기
Function tenOrgCardNum(icardNum, byref mylevel, byref myUserid)
	Dim strsql, tmpUserseq
	strsql = ""
	strsql = strsql & " SELECT TOP 1 sc.userseq, su.OnlineUserId, IsNULL(l.userlevel,5) as userlevel "
	strsql = strsql & " FROM db_shop.dbo.tbl_total_shop_user as su "
	strsql = strsql & " JOIN db_shop.dbo.tbl_total_shop_card as sc on su.userseq = sc.userseq "
	strsql = strsql & " JOIN [db_user].[dbo].[tbl_logindata] as l on su.OnlineUserId = l.userid "
	strsql = strsql & " WHERE sc.cardNo = '"&icardNum&"' "
	rsget.Open strsql, dbget, 1
	If (Not rsget.Eof) then
		tmpUserseq		= rsget("userseq")
		myuserid		= rsget("OnlineUserId")
		mylevel			= levelChange(rsget("userlevel"))
	End If
	rsget.close

    ''카드를 변경한 CASE 검토.
	strsql = ""
	strsql = strsql & " IF NOT EXISTS(SELECT cardNo FROM db_shop.dbo.tbl_total_shop_card WHERE cardno = '"&icardNum&"' AND useYN = 'Y') "
	strsql = strsql & " 	BEGIN "
	strsql = strsql & " 		SELECT TOP 1 cardNo FROM db_shop.dbo.tbl_total_shop_card WHERE useYN = 'Y' AND userseq = '"&tmpUserseq&"' ORDER BY Regdate DESC "
	strsql = strsql & " 	END "
	strsql = strsql & " ELSE "
	strsql = strsql & " 	BEGIN "
	strsql = strsql & " 		SELECT '"&icardNum&"' as cardNo "
	strsql = strsql & " 	END "
	rsget.Open strsql, dbget, 1
	If (Not rsget.Eof) then
		tenOrgCardNum = rsget("cardNo")
	End If
	rsget.close

End Function

'신규 멤버쉽 발급 받은 자에게 멤버쉽 발급 축하 쿠폰 발급
Function isNewCardCoupon(fmember_id)
	Dim strsql
	strsql = ""
	strsql = strsql & " IF NOT EXISTS(SELECT userid FROM db_etcmall.dbo.tbl_syrup_couponLog WHERE userid = '" & fmember_id & "') "
	strsql = strsql & " BEGIN " & vbCrlf
	strsql = strsql & " 	INSERT INTO [db_user].[dbo].tbl_user_coupon " & vbCrlf
	strsql = strsql & " 	(masteridx, userid, couponvalue, coupontype, couponname, minbuyprice, " & vbCrlf
	strsql = strsql & " 	targetitemlist, startdate, expiredate) " & vbCrlf
	strsql = strsql & "	 	VALUES (744,'" & fmember_id & "',5000, '2', '6월 시럽 5,000원 쿠폰', 20000, " & vbCrlf
	strsql = strsql & " 	'','2015-06-08 00:00:00' ,'2015-07-07 23:59:59') " & vbCrlf

	strsql = strsql & " 	INSERT INTO db_etcmall.dbo.tbl_syrup_couponLog " & vbCrlf
	strsql = strsql & " 	(userid, couponidx, regdate) " & vbCrlf
	strsql = strsql & " 	VALUES ('"&fmember_id&"', '744', getdate())" & vbCrlf
	strsql = strsql & " END " & vbCrlf
	dbget.Execute strsql, 1
End Function

'6개의 등급외에 등급코드가 나왔을 시 Yellow로 치환 : YELLOW를 006으로 칭함!
Function levelChange(slevel)
	Select Case slevel
		Case "7"	levelChange = "006"			'STAFF
		Case "6"	levelChange = "009"			'VVIP	//2016-07-07 FRIEND -> VVIP & 006 -> 009로 수정..
		Case "9"	levelChange = "006"			'MANIA
		Case "8"	levelChange = "006"			'FAMILY
		Case "0"	levelChange = "006"			'YELLOW
		Case Else	levelChange = "00"&slevel	'그 외
	End Select
End Function

'우리쪽에 CI값이 없으면 업데이트, 기존에 CI값이 있으면 수정하지 않는다.
Function fnConnInfoUpdateProc(fci, fmember_id)
	Dim strSql
	strSql = ""
	strSql = strSql & " IF EXISTS (SELECT TOP 1 connInfo FROM db_user.dbo.tbl_user_n WHERE useriD = '"&fmember_id&"' AND isNull(connInfo, '') = '' ) " & VBCRLF
	strSql = strSql & " 	BEGIN " & VBCRLF
	strSql = strSql & " 		UPDATE db_user.dbo.tbl_user_n SET connInfo = '"&fci&"', realnamecheck = 'Y' " & VBCRLF
	strSql = strSql & " 		WHERE userid = '"&fmember_id&"' " & VBCRLF
	strSql = strSql & " 	END " & VBCRLF
	dbget.Execute strSql, 1

	strSql = ""
	strSql = strSql & " IF EXISTS (SELECT TOP 1 connInfo FROM db_shop.dbo.tbl_total_shop_user WHERE OnlineUserID = '"&fmember_id&"' AND isNull(connInfo, '') = '' ) " & VBCRLF
	strSql = strSql & " 	BEGIN " & VBCRLF
	strSql = strSql & "	 		UPDATE db_shop.dbo.tbl_total_shop_user SET connInfo = '"&fci&"'" & VBCRLF
	strSql = strSql & "	 		WHERE OnlineUserID = '"&fmember_id&"' " & VBCRLF
	strSql = strSql & "		END " & VBCRLF
	dbget.Execute strSql, 1
End Function
%>