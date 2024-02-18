<%
'// 롯데아이몰API 연동서버 URL

'배송비코드 23725 => 640839
'기본출고지 44765 => 525712
'기본반품지 44764 => 525713

Dim ltiMallAPIURL, ltiMallAuthNo, ltiMallTenID, tenBrandCd, tenDlvCd, tenDlvFreeCd
IF application("Svr_Info") = "Dev" THEN
	'ltiMallAPIURL = "http://openapidev.lotteimall.com"	'' 테스트서버
	ltiMallAPIURL = "http://openapitst.lotteimall.com"	'' 테스트서버
	tenDlvCd = "640839"
	tenDlvFreeCd = "577045"
Else
	ltiMallAPIURL = "https://openapi.lotteimall.com"		'' 실서버
	tenDlvCd = "640839"
	tenDlvFreeCd = "577045"
End if
ltiMallTenID = "011799LT"

Dim updateAuth, dbAuthNo
dim iisql
iisql = "select top 1 isnull(iniVal, '') as iniVal, lastupdate "&VbCRLF
iisql = iisql & " from db_etcmall.dbo.tbl_outmall_ini"&VbCRLF
iisql = iisql & " where mallid='lotteimall'"&VbCRLF
iisql = iisql & " and inikey='auth'"
rsget.CursorLocation = adUseClient
rsget.Open iisql, dbget, adOpenForwardOnly, adLockReadOnly
if not rsget.Eof then
    dbAuthNo	= rsget("iniVal")
    updateAuth	= rsget("lastupdate")
end if
rsget.close

'// 롯데아이몰 인증코드 확인(매일 업데이트; 어플리케이션변수에 저장)
If DateDiff("h", updateAuth, now()) > 12 OR dbAuthNo = "" then
	Dim objXML, xmlDOM
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objXML.Open "GET", ltiMallAPIURL & "/openapi/createCertification.lotte?strUserId=" & ltiMallTenID & "&strPassWd=cube101010!*", False
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
			On Error Resume Next
				ltiMallAuthNo = xmlDOM.getElementsByTagName("SubscriptionId").item(0).text		'인증번호 저장
				If Err <> 0 then
					ltiMallAuthNo = ""
				End If
				iisql = "update db_etcmall.dbo.tbl_outmall_ini "&VbCRLF
				iisql = iisql & " set iniVal='"&ltiMallAuthNo&"'"&VbCRLF
				iisql = iisql & " ,lastupdate=getdate()"&VbCRLF
				iisql = iisql & " where mallid='lotteimall'"&VbCRLF
				iisql = iisql & " and inikey='auth'"
				dbget.Execute iisql
				If Err <> 0 then
					Response.Write "<script language=javascript>alert('Lotteimall.com인증에 오류가 발생했습니다.\n나중에 다시 시도해보세요');history.back();</script>"
					Response.End
				End If
			On Error Goto 0
			Set xmlDOM = Nothing
		Else
			Response.Write "<script language=javascript>alert('Lotteimall.com과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');history.back();</script>"
			Response.End
		End If
	Set objXML = Nothing
Else
	ltiMallAuthNo = dbAuthNo
End If
%>