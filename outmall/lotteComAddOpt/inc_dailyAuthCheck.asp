<%
''//https://partner.lotte.com/main/Login.lotte �α�����å�� ���� // �̰��� ���� �����ؾ� �ϴµ�.
'// �Ե�����API �������� URL
dim lotteAPIURL, lotteAuthNo, lottenTenID, tenBrandCd, tenDlvCd, CertPasswd
Dim iisql
IF application("Svr_Info")="Dev" THEN
	'lotteAPIURL = "http://openapidev.lotte.com"	'' �׽�Ʈ����
	lotteAPIURL = "http://openapitest.lotte.com"	'' �׽�Ʈ����
	tenBrandCd = "14846"	'�ٹ�(�ӽ�)
	tenDlvCd = "706460"		'�����å�ڵ�
	CertPasswd = "1234"		'Dev�� ��� : 1234
Else
	lotteAPIURL = "https://openapi.lotte.com"		'' �Ǽ���
	tenBrandCd = "155112"	'�ٹ�����
	tenDlvCd = "706460"
	CertPasswd = "cube1010!"
End if
lottenTenID = "124072"					'�ٹ�����ID

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
			lotteAuthNo = xmlAuthDOM.getElementsByTagName("SubscriptionId").item(0).text		'������ȣ ����
			if Err<>0 then
				Response.Write "<script language=javascript>alert('Lotte.com������ ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');history.back();</script>"
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
		Response.Write "<script language=javascript>alert('Lotte.com�� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');history.back();</script>"
		Response.End
	end if
	Set objAuthXML = Nothing
Else
	lotteAuthNo = dbAuthNo
End If


'############## ���ϴ� ���� �ҽ� ################
'	'// �Ե����� �����ڵ� Ȯ��(���� ������Ʈ; ���ø����̼Ǻ����� ����)
'	'Application("lotteAuthDate")="2012-01-01"
'
'	if Application("lotteAuthDate")="" or datediff("n",Application("lotteAuthDate"),now())>10 then
'		''94�������� ���� ����� �� ������.
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
'    		Application("lotteAuthNo") = lotteAuthNo		'������ȣ ����
'    		Application("lotteAuthDate") = now()			'�����ð� ���
'    		if Err<>0 then
'    			Response.Write "<script language=javascript>alert('Lotte.com������ ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');history.back();</script>"
'    			Response.End
'    		end if
'    	on Error Goto 0
'
'	end if
'lotteAuthNo = Application("lotteAuthNo")
'##################################################
%>