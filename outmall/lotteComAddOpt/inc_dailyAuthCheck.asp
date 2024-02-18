<%
''//https://partner.lotte.com/main/Login.lotte 로그인정책과 같음 // 이곳도 같이 변경해야 하는듯.
'// 롯데닷컴API 연동서버 URL
dim lotteAPIURL, lotteAuthNo, lottenTenID, tenBrandCd, tenDlvCd, CertPasswd
Dim iisql
IF application("Svr_Info")="Dev" THEN
	'lotteAPIURL = "http://openapidev.lotte.com"	'' 테스트서버
	lotteAPIURL = "http://openapitest.lotte.com"	'' 테스트서버
	tenBrandCd = "14846"	'텐바(임시)
	tenDlvCd = "706460"		'배송정책코드
	CertPasswd = "1234"		'Dev는 비번 : 1234
Else
	lotteAPIURL = "https://openapi.lotte.com"		'' 실서버
	tenBrandCd = "155112"	'텐바이텐
	tenDlvCd = "706460"
	CertPasswd = "cube1010!"
End if
lottenTenID = "124072"					'텐바이텐ID

Dim updateAuth, dbAuthNo
iisql = "select top 1 isnull(iniVal, '') as iniVal, lastupdate "&VbCRLF
iisql = iisql & " from db_etcmall.dbo.tbl_outmall_ini"&VbCRLF
iisql = iisql & " where mallid='lotteCom'"&VbCRLF
iisql = iisql & " and inikey='auth'"
rsget.CursorLocation = adUseClient
rsget.Open iisql, dbget, adOpenForwardOnly, adLockReadOnly
if not rsget.Eof then
    dbAuthNo	= rsget("iniVal")
    updateAuth	= rsget("lastupdate")
end if
rsget.close

If DateDiff("h", updateAuth, now()) > 12 OR dbAuthNo = "" then
	dim objAuthXML, xmlAuthDOM
	Set objAuthXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objAuthXML.Open "POST", lotteAPIURL & "/openapi/createCertification.lotte?strUserId=" & lottenTenID & "&strPassWd="&CertPasswd&"", false
	objAuthXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objAuthXML.Send()
	If objAuthXML.Status = "200" Then
		Set xmlAuthDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlAuthDOM.async = False
		xmlAuthDOM.LoadXML BinaryToText(objAuthXML.ResponseBody, "euc-kr")

		on Error Resume Next
			lotteAuthNo = xmlAuthDOM.getElementsByTagName("SubscriptionId").item(0).text		'인증번호 저장
			if Err<>0 then
				Response.Write "<script language=javascript>alert('Lotte.com인증에 오류가 발생했습니다.\n나중에 다시 시도해보세요');history.back();</script>"
				Response.End
			end if

			iisql = "update db_etcmall.dbo.tbl_outmall_ini "&VbCRLF
			iisql = iisql & " set iniVal='"&lotteAuthNo&"'"&VbCRLF
			iisql = iisql & " ,lastupdate=getdate()"&VbCRLF
			iisql = iisql & " where mallid='lotteCom'"&VbCRLF
			iisql = iisql & " and inikey='auth'"
			dbget.Execute iisql
		on Error Goto 0

		Set xmlAuthDOM = Nothing
	else
		Response.Write "<script language=javascript>alert('Lotte.com과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');history.back();</script>"
		Response.End
	end if
	Set objAuthXML = Nothing
Else
	lotteAuthNo = dbAuthNo
End If


'############## 이하는 예전 소스 ################
'	'// 롯데닷컴 인증코드 확인(매일 업데이트; 어플리케이션변수에 저장)
'	'Application("lotteAuthDate")="2012-01-01"
'
'	if Application("lotteAuthDate")="" or datediff("n",Application("lotteAuthDate"),now())>10 then
'		''94서버에서 디비로 저장된 값 가져옴.
'		dim iisql
'		iisql = "select top 1 iniVal "&VbCRLF
'		iisql = iisql & " from db_etcmall.dbo.tbl_outmall_ini"&VbCRLF
'		iisql = iisql & " where mallid='lotteCom'"&VbCRLF
'		iisql = iisql & " and inikey='auth'"
'
'		rsget.Open iisql, dbget, 1
'		if not rsget.Eof then
'		    lotteAuthNo = rsget("iniVal")
'		end if
'		rsget.close
'
'		on Error Resume Next
'    		Application("lotteAuthNo") = lotteAuthNo		'인증번호 저장
'    		Application("lotteAuthDate") = now()			'인증시간 기록
'    		if Err<>0 then
'    			Response.Write "<script language=javascript>alert('Lotte.com인증에 오류가 발생했습니다.\n나중에 다시 시도해보세요');history.back();</script>"
'    			Response.End
'    		end if
'    	on Error Goto 0
'
'	end if
'lotteAuthNo = Application("lotteAuthNo")
'##################################################
%>