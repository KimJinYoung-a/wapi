<%
'// 결제 정보

Dim PayInfoHTML

Dim ReqInfoHTML

Dim MailTo , MailTo_Nm

Function getInfo(vOrderSerial)

	dim strInfo_Html , strSQL

	dim PayMethod		'결제방법
	dim PayMethodName 	'결제방법명
	dim PayStatus		'결제상태
	dim SpendMileage 	'마일리지 사용액
	dim	TenCardSpend	'할인권 사용액
	dim AllAtDisPrice	'기타할인액(올엣)
	dim TotalPayPrice	'총 결제금액
	dim AccountNo		'입금계좌 정보

	dim ReqName		'받으시는분
	dim ReqPhone	'전화번호
	dim ReqHp		'핸드폰
	dim ReqZipCode	'우편번호
	dim ReqAddress	'배송주소
	dim ReqComment	'배송메모

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

		getInfo 	= 0 '정상

	ELSE
		getInfo 	= -1 '오류
		PayInfoHTML		=""
		ReqInfoHTML		=""

		rsget.Close
		Exit Function

	End IF

	rsget.Close

	'//=============  결제 정보 시작 ================//

	SELECT CASE PayMethod
		CASE "100" '신용카드
			PayMethodName="신용카드"
			PayStatus	="결제완료"
		CASE "80" ' 올엣카드
			PayMethodName="올엣카드"
			PayStatus	="결제완료"
		CASE "20" ' 실시간 계좌이체
			PayMethodName="실시간 계좌이체"
			PayStatus	="결제완료"
		CASE "7" ' 무통장 입금
			PayMethodName="무통장 입금"
			PayStatus	="입금전 상태"
		CASE ELSE
			PayMethodName=""
			PayStatus	="입금전 상태"
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
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> 결제방법 </td> "&_
		" 					<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& PayMethodName &" </td> "&_
		"					<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">결제상태</td> "&_
		" 					<td valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& PayStatus &"</td> "&_
		" 				</tr> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">마일리지사용액</td> "&_
		" 					<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& SpendMileage &" P </td> "&_
		" 					<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">할인권사용액</td> "&_
		" 					<td valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& TenCardSpend &" 원</td> "&_
		" 				</tr> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">기타 할인액</td> "&_
		" 					<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& AllAtDisPrice &" 원</td> "&_
		" 					<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">총 결제금액 </td> "&_
		" 					<td valign=""bottom"" class=""price"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""><strong> "& TotalPayPrice &" 원</strong></td> "&_
		" 				</tr> "&_
		" 				</table> "&_
		" 			</td> "&_
		" 		</tr> "


	IF PayMethod = "7" THEN '무통장 입금
		PayInfoHTML= 	PayInfoHTML &_
		" 		 <!-- 무통장입금 --> "&_
		" 		<tr> "&_
		" 			<td align=""center""  style=""padding:5 0 0 0 ""> "&_
		" 				<table width=""548"" height=""31""  border=""0"" cellpadding=""0"" cellspacing=""0""style=""border-top:1px solid #dddddd""> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> 입금 은행 정보 </td> "&_
		" 					<td valign=""bottom"" class=""BIG_Black"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""><strong>&nbsp;"& AccountNo &" </strong> (주)텐바이텐</td> "&_
		" 				</tr> "&_
		" 				</table> "&_
		" 			</td> "&_
		" 		</tr> "&_
		" 		<tr> "&_
		" 			<td align=""left"" class=""black11px"" style=""padding:10 15 0 15"">* 무통장입금 확인은 평일 매일 오전 10시, 오후 3시 두번 이루어지며 입금확인시 배송이 이루어집니다.<br> "&_
		" 			* 무통장주문 후 7일이 지날때까지 입금이 안되면 주문은 자동으로 취소됩니다. 일부 한정상품 주문시 유의하여 주시기 바랍니다.</td> "&_
		" 		</tr> "
	END IF

		PayInfoHTML= 	PayInfoHTML &_
		" 		</table> "&_
		" 	</td> "&_
		" </tr> "&_
		" </table>"

	PayInfoHTML = PayInfoHTML

	'//=============  결제 정보 끝 ================//

	'//=============  배송지 정보 시작 =================//

	ReqInfoHTML= ""&_
		"<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" style=""border-top:3px solid #be0808;font-family:Dotum; font-size:11px; color:#888; padding-top:3px"">"&_
		"<tr>"&_
		"	<td height=""30"" width=""120"" style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;"">받으시는분</span></td>"&_
		"	<td colspan=""3"" style=""border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;""> "& ReqName &" &nbsp;</span></td>"&_
		"</tr>"&_
		"<tr>"&_
		"	<td height=""30"" style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;"">휴대폰번호</span></td>"&_
		"	<td style=""border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;""> "& ReqHp &" &nbsp;</span></td>"&_
		"	<td width=""120"" style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;"">전화번호</span></td>"&_
		"	<td width=""205"" style=""border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;""> "& ReqPhone &" &nbsp;</span></td>"&_
		"</tr>"&_
		"<tr>"&_
		"	<td height=""30"" style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;"">주소</span></td>"&_
		"	<td colspan=""3"" style=""border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;""> ["& ReqZipCode &"]" & ReqAddress &" &nbsp;</span></td>"&_
		"</tr>"&_
		"<tr>"&_
		"	<td height=""30"" style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;"">배송 유의사항</span></td>"&_
		"	<td colspan=""3"" style=""border-bottom:1px solid #eaeaea;margin-left:15px;""><span style=""margin-left:15px;""> "& ReqComment &" &nbsp;</span></td>"&_
		"</tr>"&_
		"</table>"		
	ReqInfoHTML = ReqInfoHTML

	'//=============  배송지 정보 끝 =================//

End Function

Function getInfo_off(vmasteridx)
	dim strInfo_Html , strSQL
	dim PayMethod		'결제방법
	dim PayMethodName 	'결제방법명
	dim PayStatus		'결제상태
	dim SpendMileage 	'마일리지 사용액
	dim	TenCardSpend	'할인권 사용액
	dim AllAtDisPrice	'기타할인액(올엣)
	dim TotalPayPrice	'총 결제금액
	dim AccountNo		'입금계좌 정보
	dim ReqName		'받으시는분
	dim ReqPhone	'전화번호
	dim ReqHp		'핸드폰
	dim ReqZipCode	'우편번호
	dim ReqAddress	'배송주소
	dim ReqComment	'배송메모

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

		getInfo_off 	= 0 '정상

	ELSE
		getInfo_off 	= -1 '오류
		PayInfoHTML		=""
		ReqInfoHTML		=""

		rsget.Close
		Exit Function

	End IF

	rsget.Close

	'//=============  배송지 정보 시작 =================//
	ReqInfoHTML= ""&_
	" <table width=""550"" border=""0"" cellspacing=""0"" cellpadding=""0""> "&_
	" <tr> "&_
	" 	<td style=""padding:0 0 7 0;""><img src=""http://fiximage.10x10.co.kr/web2008/mail/a01_text03.gif"" width=""330"" height=""18""></td> "&_
	" </tr> "&_
	" <tr> "&_
	" 	<td> "&_
	" 		<table width=""548"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-top:1px solid #dddddd""> "&_
	" 		<tr> "&_
	" 			<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">받으시는 분 </td> "&_
	" 			<td width=""438"" colspan=""4"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& ReqName &" &nbsp;</td> "&_
	" 		</tr> "&_
	" 		<tr> "&_
	" 			<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">전화번호</td> "&_
	" 			<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& ReqPhone &" &nbsp;</td> "&_
	" 			<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">휴대폰번호</td> "&_
	" 			<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& ReqHp &" &nbsp;</td> "&_
	" 		</tr> "&_
	" 		<tr> "&_
	" 			<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">배송주소</td> "&_
	" 			<td width=""438"" colspan=""3"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> ["& ReqZipCode &"]" & ReqAddress &" &nbsp;</td> "&_
	" 		</tr> "&_
	" 		<tr> "&_
	" 			<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">유의사항</td> "&_
	" 			<td width=""438"" colspan=""3"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& ReqComment &" &nbsp;</td> "&_
	" 		</tr> "&_
	" 		</table> "&_
	" 	</td> "&_
	" </tr> "&_
	" </table> "
	ReqInfoHTML = ReqInfoHTML
	'//=============  배송지 정보 끝 =================//
End Function

%>