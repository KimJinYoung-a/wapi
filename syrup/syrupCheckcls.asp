<%
class CTenUserItem
	public FUserID
	public FUserDiv
	public FUserLevel
	public FUserName
	public FUserEmail
	public FUserIcon
	public FLoginTime
	public FCouponCnt
	public FCurrentMileage
	public FCurrentTenCash		'// 예치금
	public FCurrentTenGiftCard	'// 기프트카드
	public FRealNameCheck

	''200907추가
	public FSexFlag
    public FAge

    public FBaguniCount		''201004추가
	public ForderCount		''201409추가

	Private Sub Class_Initialize()
        FBaguniCount = 0
        ForderCount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CTenUser
	public FOneUser

	public FRectUserID
	public FRectPassWord
	public FRectEnc

	private FPassOk
    private FNotUsingSite

	public FConfirmUser
	Public FRectUserEmail
	public function IsPassOk()
		IsPassOk = FPassOk
	end function

    public function IsRequireUsingSite
        IsRequireUsingSite = (FNotUsingSite=true)
    end function

	Public bIsExist
	Public bIsMailExist
	public Sub LoginProc()
		dim sqlStr
		dim tmpuserpass
		dim tmpuserdiv, tmplevel, tmpuserid, tmpuserlevel, tmplogintime
        dim tmpEnc_userpass, tmpEnc_userpass64, EncedPassWord, EncedPassWord64
        dim tmpsexflag, tmpage

		FPassOk = false
		FConfirmUser = "Y"		'(Y: 승인회원, N:승인대기, E:기간만료, O:기존회원, X:정지회원)

		if (FRectUserID="") or (FRectPassWord="") then Exit Sub

        if FRectEnc then
        	EncedPassWord = FRectPassWord
        	EncedPassWord64 = FRectPassWord
        else
        	EncedPassWord = Md5(FRectPassWord)
        	EncedPassWord64 = SHA256(Md5(FRectPassWord))
        end if

		sqlStr = " select top 1 userid, userdiv, IsNULL(userlevel,5) as userlevel," + VbCrlf
		sqlStr = sqlStr + " userpass, Enc_userpass, Enc_userpass64, "
		sqlStr = sqlStr + " convert(varchar(19),getdate(),20) as logintime " + VbCrlf
		sqlStr = sqlStr + " from [db_user].[dbo].[tbl_logindata]" + vbCrlf
		sqlStr = sqlStr + " where userid='" + FRectUserID + "'" + vbCrlf
		sqlStr = sqlStr + " and userid<>''" + vbCrlf

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			tmpuserid = rsget("userid")
			tmpuserdiv = rsget("userdiv")
			tmpuserpass = rsget("userpass")
			tmpuserlevel = rsget("userlevel")
			tmplogintime = rsget("logintime")
			tmpEnc_userpass = rsget("Enc_userpass")
			tmpEnc_userpass64 = rsget("Enc_userpass64")

			FPassOk    = true
		end if
		rsget.Close
        ''암호화 사용(SHA256)
        FPassOk = FPassOk and (UCASE(EncedPassWord64)=UCASE(tmpEnc_userpass64))

		if Not FPassOk then Exit Sub
        ''#### 사용 사이트 Check ( 텐바이텐 이용안함 원하는 고객.) ######
        sqlStr = " select count(*) as cnt "
        sqlStr = sqlStr & " from db_user.dbo.tbl_user_allow_site"
        sqlStr = sqlStr & " where userid='" & FRectUserID & "'"
        sqlStr = sqlStr & " and userid<>''"
        sqlStr = sqlStr & " and sitegubun='10x10'"
        sqlStr = sqlStr & " and siteusing='N'"

        rsget.Open sqlStr,dbget,1
            FPassOk = rsget("cnt")=0
        rsget.Close

        if (Not FPassOk) then
            FNotUsingSite = true
            Exit Sub
        end if

		'// 휴면 회원 확인 (일반회원만 적용 2015.08.13; 허진원)
		dim chkHoldUser: chkHoldUser=false
		if (tmpuserdiv="01") or (tmpuserdiv="99") then
			sqlStr = " select count(*) " + VbCrlf
			sqlStr = sqlStr + " from [db_user].[dbo].[tbl_user_n]" + vbCrlf
			sqlStr = sqlStr + " where userid='" + FRectUserID + "'"
			
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
			if rsget(0)<=0 then
				chkHoldUser = true
			end if
			rsget.Close

			if chkHoldUser then
				'회원정보가 없으면 휴면 복구 처리
				sqlStr = "db_user_hold.dbo.sp_Ten_HoldUserRevive ('" & FRectUserID & "','S')"
	            ''Response.Write sqlStr: response.End
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
				if Not rsget.Eof then
					FPassOk = rsget(0)>0
				end if
				rsget.Close	
				if Not FPassOk then Exit Sub
			end if
		end if
        ''###############################################################

		set FOneUser = new CTenUserItem
		FOneUser.FUserID = tmpuserid
		FOneUser.FUserDiv = tmpuserdiv
		FOneUser.FUserLevel = tmpuserlevel
		FOneUser.FLoginTime = tmplogintime

        FOneUser.FSexFlag = tmpsexflag
        FOneUser.FAge     = tmpage

		'FOneUser.FUserLevel = tmplevel
		if (tmpuserdiv="01") or (tmpuserdiv="99") then
            sqlStr = " exec [db_user].[dbo].[sp_Ten_User_Login_AddInfo] '"&FRectUserID&"'"  ''2014/12/23 변경
			rsget.Open sqlStr,dbget,1
				if Not rsget.Eof then
					FOneUser.FUserName = db2html(rsget("username"))
					FOneUser.FUserEmail = db2html(rsget("usermail"))
					FOneUser.FUserIcon = rsget("usericon")
					IF isNull(FOneUser.FUserIcon) then FOneUser.FUserIcon = ""
					FOneUser.FCouponCnt = rsget("couponCnt")

					FOneUser.FCurrentMileage 		= rsget("currentmileage")
					FOneUser.FCurrentTenCash 		= rsget("currenttencash")
					FOneUser.FCurrentTenGiftCard 	= rsget("currenttengiftcard")

				    '// 가입대기 확인
				    if isNull(rsget("userStat")) then
				    	FConfirmUser = "O"
				    elseif (rsget("userStat")="N") then
				    	FPassOk = false
				    	if (datediff("h",rsget("regdate"),now())<=12) then
			    			FConfirmUser = "N"
			    		else
			    			FConfirmUser = "E"
			    		end if
				    end if

				end if
			rsget.Close
		ElseIf (tmpuserdiv="95") or (tmpuserdiv="96") then
			'### 일시정지 회원 (사용안함, 외부 기타요인에 의한 회원 정지처리)
			FPassOk = false
			FConfirmUser = "X"
			Exit Sub
		end if

	    '// 로그인 카운터 및 기타 정보 저장
	    sqlStr = ""
		sqlStr = sqlStr & " UPDATE [db_user].dbo.[tbl_logindata]" & vbCrlf
		sqlStr = sqlStr & " SET lastlogin=getdate()," & vbCrlf
		sqlStr = sqlStr & " counter = counter+1," & vbCrlf
		sqlStr = sqlStr & " lastrefip='" &  Left(request.ServerVariables("REMOTE_ADDR"),32) & "'"  & vbCrlf
		sqlStr = sqlStr & " WHERE userid='" & FRectUserID & "'" & vbCrlf
		dbget.Execute sqlStr, 1
	end Sub

	Public Sub DuplicateUserIDProc()
		Dim strSql
		Dim vLeftUserIDCheck
		strSql = "select userid from [db_user].[dbo].tbl_logindata where userid = '" & FRectUserID & "' "
		rsget.Open strSql, dbget, 1
		If rsget.EOF = True Then
			bIsExist = False
		Else
			bIsExist = True
		End If
		rsget.Close
	
		strSql = "select userid from [db_user].[dbo].tbl_deluser where userid = '" + FRectUserID + "'"
		rsget.Open strSql, dbget, 1
		bIsExist = bIsExist or (Not rsget.Eof)
		rsget.Close
	End Sub

	Public Sub DuplicateUserEmailProc()
		Dim strSql
		'// 회원정보에서 인증기록이 있는 정보만 확인(userStat N:인증전, Y:인증완료, Null:기존고객)
		strSql = "select top 1 userid from [db_user].[dbo].tbl_user_n " &_
				" where usermail='" & FRectUserEmail & "' " &_
				" and (userStat='Y' or (userStat='N' and datediff(hh,regdate,getdate())<12)) "
		rsget.Open strSql, dbget, 1
	
		'동일한 이메일 없음
		If rsget.EOF = True Then
			bIsMailExist = False
		'동일한 이메일 존재
		Else
			bIsMailExist = True
		End If
		rsget.Close
	End Sub

	Private Sub Class_Initialize()
        FNotUsingSite = false
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

'//로그인 오류 메시지 출력
Function getErrMsg(sCd,byRef fDesc)
	Select Case sCd
		Case "1000"
			getErrMsg = "Y"
			fDesc = "로그인에 성공하였습니다."
		Case "1102"
			getErrMsg = "N"
			fDesc = "아이디 또는 비밀번호가 맞지 않습니다."
		Case "1103"
			getErrMsg = "N"
			fDesc = "아이디 또는 비밀번호가 맞지 않습니다."
		Case "1201"
			getErrMsg = "N"
			fDesc = "텐바이텐 이용를 허용하지 않으셨습니다."
		Case "1202"
			getErrMsg = "N"
			fDesc = "죄송합니다. 사용이 불가한 회원입니다."
		Case "1301"
			getErrMsg = "N"
			fDesc = "아직 인증을 받지 않은 회원입니다."
		Case "9999"
			getErrMsg = "N"
		Case Else
			getErrMsg = "N"
			fDesc = "로그인에 실패하였습니다."
	End Select
End Function

'//ID 중복체크 오류 메시지 출력
Function getIDCheckErrMsg(sCd,byRef fDesc)
	Select Case sCd
		Case "3000"
			getIDCheckErrMsg = "Y"
			fDesc = "사용 가능한 아이디 입니다."
		Case "3102"
			getIDCheckErrMsg = "N"
			fDesc = "이미 사용중인 아이디 입니다."
		Case "3103"
			getIDCheckErrMsg = "N"
			fDesc = "아이디를 정확히 입력하세요.(4~32 자리로 입력)"
		Case "3201"
			getIDCheckErrMsg = "N"
			fDesc = "특수문자나 한글/한문은 사용불가능합니다."
		Case Else
			getIDCheckErrMsg = "N"
			fDesc = "아이디 중복체크에 실패하였습니다."
	End Select
End Function

'//이메일 중복체크 오류 메시지 출력
Function getEmailCheckErrMsg(sCd,byRef fDesc)
	Select Case sCd
		Case "5000"
			getEmailCheckErrMsg = "Y"
			fDesc = "사용 가능한 이메일 입니다."
		Case "5102"
			getEmailCheckErrMsg = "N"
			fDesc = "이미 사용중인 이메일 입니다."
		Case "5103"
			getEmailCheckErrMsg = "N"
			fDesc = "이메일 주소를 정확히 입력하세요."
		Case Else
			getEmailCheckErrMsg = "N"
			fDesc = "이메일 중복체크에 실패하였습니다."
	End Select
End Function

Function ls10x10(pGamepopId)
	Dim i, MyArray, Check
	i=1
	DO until i>len( pGamepopId)
		MyArray=mid(pGamepopid,i,cint(1))
		If MyArray >= "a" and MyArray <= "z" Then
			ls10x10=False
		Elseif MyArray >= "A" and MyArray <= "Z" Then
			ls10x10=False
		ElseIf  MyArray >= "0" and MyArray <= "9" Then
			ls10x10=False
		Else
			ls10x10=True
			Exit function
		End If
		i = i + 1
	Loop
End Function

Function chkEmailForm(strEmail)
	dim isValidE, regEx

	isValidE = True
	set regEx = New RegExp

	regEx.IgnoreCase = False

	regEx.Pattern = "^[a-zA-Z0-9][\w\.-_]*[a-zA-Z0-9][\w\.-_]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
	isValidE = regEx.Test(strEmail)

	chkEmailForm = isValidE
end Function
%>