<%
'// ���� ����

Dim PayInfoHTML

Dim ReqInfoHTML

Dim MailTo , MailTo_Nm

Function getInfo(vOrderSerial)

	dim strInfo_Html , strSQL

	dim PayMethod		'�������
	dim PayMethodName 	'���������
	dim PayStatus		'��������
	dim SpendMileage 	'���ϸ��� ����
	dim	TenCardSpend	'���α� ����
	dim AllAtDisPrice	'��Ÿ���ξ�(�ÿ�)
	dim TotalPayPrice	'�� �����ݾ�
	dim AccountNo		'�Աݰ��� ����

	dim ReqName		'�����ôº�
	dim ReqPhone	'��ȭ��ȣ
	dim ReqHp		'�ڵ���
	dim ReqZipCode	'������ȣ
	dim ReqAddress	'����ּ�
	dim ReqComment	'��۸޸�

	PayInfoHTML = ""
	ReqInfoHTML = ""
	MailTo = ""
	MailTo_Nm = ""

	strSQL =" SELECT Top 1 BuyName , BuyEmail , AccountDiv,AccountNo,SubTotalPrice " &_
			" , IsNULL(miletotalprice,0) as SpendMileage , IsNULL(tencardspend,0) as TenCardSpend , IsNULL(allatdiscountprice,0) as AllAtDiscountPrice " &_
			" , ReqName , ReqPhone , ReqHp , ReqZipCode , (ReqZipAddr + ' ' + ReqAddress) as ReqAllAddress , Comment " &_
			" FROM [db_order].[dbo].tbl_order_master " &_
			" WHERE cancelyn='N' and orderserial = '"& vOrderSerial &"' "

	rsget.open strSQL, dbget,2

	IF not rsget.eof THEN

		MailTo_Nm  	= db2html(rsget("BuyName"))
		MailTo 		= db2html(rsget("BuyEmail"))

		PayMethod 		= CStr(rsget("AccountDiv"))
		AccountNo 		= rsget("AccountNo")
		SpendMileage 	= FormatNumber(rsget("SpendMileage"),0)
		TenCardSpend 	= FormatNumber(rsget("TenCardSpend"),0)
		AllAtDisPrice 	= FormatNumber(rsget("AllAtDiscountPrice"),0)
		TotalPayPrice 	= FormatNumber(rsget("SubTotalPrice"),0)

		ReqName 	= rsget("ReqName")
		ReqPhone 	= rsget("ReqPhone")
		ReqHp 		= rsget("ReqHp")
		ReqZipCode 	= rsget("ReqZipCode")
		ReqAddress 	= rsget("ReqAllAddress")
		ReqComment 	= rsget("Comment")

		getInfo 	= 0 '����

	ELSE
		getInfo 	= -1 '����
		PayInfoHTML		=""
		ReqInfoHTML		=""

		rsget.Close
		Exit Function

	End IF

	rsget.Close

	'//=============  ���� ���� ���� ================//

	SELECT CASE PayMethod
		CASE "100" '�ſ�ī��
			PayMethodName="�ſ�ī��"
			PayStatus	="�����Ϸ�"
		CASE "80" ' �ÿ�ī��
			PayMethodName="�ÿ�ī��"
			PayStatus	="�����Ϸ�"
		CASE "20" ' �ǽð� ������ü
			PayMethodName="�ǽð� ������ü"
			PayStatus	="�����Ϸ�"
		CASE "7" ' ������ �Ա�
			PayMethodName="������ �Ա�"
			PayStatus	="�Ա��� ����"
		CASE ELSE
			PayMethodName=""
			PayStatus	="�Ա��� ����"
	END SELECT

	PayInfoHTML= ""&_
		" <table width=""550"" border=""0"" cellspacing=""0"" cellpadding=""0""> "&_
		" <tr> "&_
		" 	<td style=""padding:0 0 7 0;""><img src=""http://fiximage.10x10.co.kr/web2008/mail/a01_text02.gif"" width=""60"" height=""18""></td> "&_
		" </tr> "&_
		" <tr> "&_
		" 	<td> "&_
		" 		<table width=""548""  border=""0"" cellspacing=""0"" cellpadding=""0""> "&_
		" 		<tr> "&_
		" 			<td align=""center""> "&_
		" 				<table width=""548"" height=""92""  border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-top:1px solid #dddddd""> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> ������� </td> "&_
		" 					<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& PayMethodName &" </td> "&_
		"					<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">��������</td> "&_
		" 					<td valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& PayStatus &"</td> "&_
		" 				</tr> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">���ϸ�������</td> "&_
		" 					<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& SpendMileage &" P </td> "&_
		" 					<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">���αǻ���</td> "&_
		" 					<td valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& TenCardSpend &" ��</td> "&_
		" 				</tr> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">��Ÿ ���ξ�</td> "&_
		" 					<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& AllAtDisPrice &" ��</td> "&_
		" 					<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">�� �����ݾ� </td> "&_
		" 					<td valign=""bottom"" class=""price"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""><strong> "& TotalPayPrice &" ��</strong></td> "&_
		" 				</tr> "&_
		" 				</table> "&_
		" 			</td> "&_
		" 		</tr> "


	IF PayMethod = "7" THEN '������ �Ա�
		PayInfoHTML= 	PayInfoHTML &_
		" 		 <!-- �������Ա� --> "&_
		" 		<tr> "&_
		" 			<td align=""center""  style=""padding:5 0 0 0 ""> "&_
		" 				<table width=""548"" height=""31""  border=""0"" cellpadding=""0"" cellspacing=""0""style=""border-top:1px solid #dddddd""> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> �Ա� ���� ���� </td> "&_
		" 					<td valign=""bottom"" class=""BIG_Black"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""><strong>&nbsp;"& AccountNo &" </strong> (��)�ٹ�����</td> "&_
		" 				</tr> "&_
		" 				</table> "&_
		" 			</td> "&_
		" 		</tr> "&_
		" 		<tr> "&_
		" 			<td align=""left"" class=""black11px"" style=""padding:10 15 0 15"">* �������Ա� Ȯ���� ���� ���� ���� 10��, ���� 3�� �ι� �̷������ �Ա�Ȯ�ν� ����� �̷�����ϴ�.<br> "&_
		" 			* �������ֹ� �� 7���� ���������� �Ա��� �ȵǸ� �ֹ��� �ڵ����� ��ҵ˴ϴ�. �Ϻ� ������ǰ �ֹ��� �����Ͽ� �ֽñ� �ٶ��ϴ�.</td> "&_
		" 		</tr> "
	END IF

		PayInfoHTML= 	PayInfoHTML &_
		" 		</table> "&_
		" 	</td> "&_
		" </tr> "&_
		" </table>"

	PayInfoHTML = PayInfoHTML

	'//=============  ���� ���� �� ================//

	'//=============  ����� ���� ���� =================//

	ReqInfoHTML= ""&_
		"<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" style=""border-top:3px solid #be0808;font-family:Dotum; font-size:11px; color:#888; padding-top:3px"">"&_
		"<tr>"&_
		"	<td height=""30"" width=""120"" style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;"">�����ôº�</span></td>"&_
		"	<td colspan=""3"" style=""border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;""> "& ReqName &" &nbsp;</span></td>"&_
		"</tr>"&_
		"<tr>"&_
		"	<td height=""30"" style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;"">�޴�����ȣ</span></td>"&_
		"	<td style=""border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;""> "& ReqHp &" &nbsp;</span></td>"&_
		"	<td width=""120"" style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;"">��ȭ��ȣ</span></td>"&_
		"	<td width=""205"" style=""border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;""> "& ReqPhone &" &nbsp;</span></td>"&_
		"</tr>"&_
		"<tr>"&_
		"	<td height=""30"" style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;"">�ּ�</span></td>"&_
		"	<td colspan=""3"" style=""border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;""> ["& ReqZipCode &"]" & ReqAddress &" &nbsp;</span></td>"&_
		"</tr>"&_
		"<tr>"&_
		"	<td height=""30"" style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;"">��� ���ǻ���</span></td>"&_
		"	<td colspan=""3"" style=""border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;""> "& ReqComment &" &nbsp;</span></td>"&_
		"</tr>"&_
		"</table>"		
	ReqInfoHTML = ReqInfoHTML

	'//=============  ����� ���� �� =================//

End Function

Function getInfo_off(vmasteridx)
	dim strInfo_Html , strSQL
	dim PayMethod		'�������
	dim PayMethodName 	'���������
	dim PayStatus		'��������
	dim SpendMileage 	'���ϸ��� ����
	dim	TenCardSpend	'���α� ����
	dim AllAtDisPrice	'��Ÿ���ξ�(�ÿ�)
	dim TotalPayPrice	'�� �����ݾ�
	dim AccountNo		'�Աݰ��� ����
	dim ReqName		'�����ôº�
	dim ReqPhone	'��ȭ��ȣ
	dim ReqHp		'�ڵ���
	dim ReqZipCode	'������ȣ
	dim ReqAddress	'����ּ�
	dim ReqComment	'��۸޸�

	PayInfoHTML = ""
	ReqInfoHTML = ""
	MailTo = ""
	MailTo_Nm = ""

	strSQL =" SELECT Top 1" &_
			" BuyName ,BuyEmail ,Comment" &_
			" , ReqName , ReqPhone , ReqHp , ReqZipCode , (ReqZipAddr + ' ' + ReqAddress) as ReqAllAddress" &_
			" FROM db_shop.dbo.tbl_shopbeasong_order_master" &_
			" WHERE cancelyn='N' and masteridx = '"& vmasteridx &"' "

	'response.write strSQL &"<Br>"
	rsget.open strSQL, dbget,2

	IF not rsget.eof THEN

		MailTo_Nm  	= db2html(rsget("BuyName"))
		MailTo 		= db2html(rsget("BuyEmail"))
		ReqName 	= rsget("ReqName")
		ReqPhone 	= rsget("ReqPhone")
		ReqHp 		= rsget("ReqHp")
		ReqZipCode 	= rsget("ReqZipCode")
		ReqAddress 	= rsget("ReqAllAddress")
		ReqComment 	= rsget("Comment")

		getInfo_off 	= 0 '����

	ELSE
		getInfo_off 	= -1 '����
		PayInfoHTML		=""
		ReqInfoHTML		=""

		rsget.Close
		Exit Function

	End IF

	rsget.Close

	'//=============  ����� ���� ���� =================//
	ReqInfoHTML= ""&_
	" <table width=""550"" border=""0"" cellspacing=""0"" cellpadding=""0""> "&_
	" <tr> "&_
	" 	<td style=""padding:0 0 7 0;""><img src=""http://fiximage.10x10.co.kr/web2008/mail/a01_text03.gif"" width=""330"" height=""18""></td> "&_
	" </tr> "&_
	" <tr> "&_
	" 	<td> "&_
	" 		<table width=""548"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-top:1px solid #dddddd""> "&_
	" 		<tr> "&_
	" 			<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">�����ô� �� </td> "&_
	" 			<td width=""438"" colspan=""4"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& ReqName &" &nbsp;</td> "&_
	" 		</tr> "&_
	" 		<tr> "&_
	" 			<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">��ȭ��ȣ</td> "&_
	" 			<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& ReqPhone &" &nbsp;</td> "&_
	" 			<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">�޴�����ȣ</td> "&_
	" 			<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& ReqHp &" &nbsp;</td> "&_
	" 		</tr> "&_
	" 		<tr> "&_
	" 			<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">����ּ�</td> "&_
	" 			<td width=""438"" colspan=""3"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> ["& ReqZipCode &"]" & ReqAddress &" &nbsp;</td> "&_
	" 		</tr> "&_
	" 		<tr> "&_
	" 			<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">���ǻ���</td> "&_
	" 			<td width=""438"" colspan=""3"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& ReqComment &" &nbsp;</td> "&_
	" 		</tr> "&_
	" 		</table> "&_
	" 	</td> "&_
	" </tr> "&_
	" </table> "
	ReqInfoHTML = ReqInfoHTML
	'//=============  ����� ���� �� =================//
End Function

%>