<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

'' Local 134
sub SendMail(mailfrom, mailto, mailtitle, mailcontent)
        dim mailobject


        IF Not (application("Svr_Info")	= "Dev") then
            '' Real Svr - 2000/ Cdonts : Local SEND..
            On Error Resume Next

            set mailobject=server.createobject("CDONTS.NewMail")
            mailobject.from = mailfrom
            mailobject.to = mailto
            mailobject.subject = mailtitle

            'html style
            mailobject.bodyformat = 0
            mailobject.mailformat = 0

            mailobject.body = mailcontent
            mailobject.send
            set mailobject = nothing

            On Error Goto 0
        Else
            '' TestSvr -2003 .. Local..?
            dim cdoMessage,cdoConfig
            Set cdoConfig = CreateObject("CDO.Configuration")

    		'-> 서버 접근방법을 설정합니다
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)

    		'-> 서버 주소를 설정합니다
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "110.93.128.94"

    		'-> 접근할 포트번호를 설정합니다
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

    		'-> 접속시도할 제한시간을 설정합니다
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10

    		'-> SMTP 접속 인증방법을 설정합니다
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

    		'-> SMTP 서버에 인증할 ID를 입력합니다
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"

    		'-> SMTP 서버에 인증할 암호를 입력합니다
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"

    		cdoConfig.Fields.Update

    		Set cdoMessage = CreateObject("CDO.Message")

    		''Set cdoMessage.Configuration = cdoConfig

    		cdoMessage.To 				= mailto
    		cdoMessage.From 			= "test@testsvr-am1vgl5" ''mailfrom
    		cdoMessage.SubJect 	= mailtitle
    		'메일 내용이 텍스트일 경우 cdoMessage.TextBody, html일 경우 cdoMessage.HTMLBody
    		cdoMessage.HTMLBody	= mailcontent

    		''테스트 환경
    		if ((InStr(mailto,"10x10.co.kr")>0) or (mailto="archilee@shinbiro.com")) then
    		    cdoMessage.Send
            end if

    		Set cdoMessage = nothing
    		Set cdoConfig = nothing



        End IF

end sub


''외부 서버로 보내기
'sub SendMail(mailfrom, mailto, mailtitle, mailcontent)
'
'		dim cdoMessage,cdoConfig
'        On Error Resume Next
'		Set cdoConfig = CreateObject("CDO.Configuration")
'
'		'-> 서버 접근방법을 설정합니다
'		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)
'
'		'-> 서버 주소를 설정합니다
'		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="mailzine.10x10.co.kr"
'
'		'-> 접근할 포트번호를 설정합니다
'		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'
'		'-> 접속시도할 제한시간을 설정합니다
'		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 5
'
'		'-> SMTP 접속 인증방법을 설정합니다
'		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'
'		'-> SMTP 서버에 인증할 ID를 입력합니다
'		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"
'
'		'-> SMTP 서버에 인증할 암호를 입력합니다
'		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"
'
'		cdoConfig.Fields.Update
'
'		Set cdoMessage = CreateObject("CDO.Message")
'
'		Set cdoMessage.Configuration = cdoConfig
'
'		cdoMessage.To 				= mailto
'		cdoMessage.From 			= mailfrom
'		cdoMessage.SubJect 	= mailtitle
'		'메일 내용이 텍스트일 경우 cdoMessage.TextBody, html일 경우 cdoMessage.HTMLBody
'		cdoMessage.HTMLBody	= mailcontent
'		cdoMessage.Send
'
'		Set cdoMessage = nothing
'		Set cdoConfig = nothing
'        On Error Goto 0
'end sub

function SendMailPayDelay(orderserial,mailfrom)
        dim sql,discountrate,paymethod, i
        dim mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal, ttlsumHTML

        mailtitle = "[텐바이텐] 주문에 대한 입금확인(미입금) 안내메일입니다"

        dim myorder
        set myorder = new COrderMaster
        myorder.FRectOrderserial = orderserial
        myorder.QuickSearchOrderMaster

        if (myorder.FOneItem.IsForeignDeliver) then
            myorder.getEmsOrderInfo
        end if

        dim myorderdetail
        set myorderdetail = new COrderMaster
        myorderdetail.FRectOrderserial = orderserial
        myorderdetail.QuickSearchOrderDetail

        if (myorder.FResultCount<1) then Exit function

        ' 파일을 불러와서 ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_pay_delay.htm"


        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
'        mailcontent = replace(mailcontent,":USERNAME:",userName)


		dim SpendMile, tencardspend
		dim IsForeighDeliver : IsForeighDeliver = false
        '주문정보 확인.---------------------------------------------------------------------------


        mailto = myorder.FOneItem.Fbuyemail
        paymethod = trim(myorder.FOneItem.Faccountdiv)


        if paymethod = "7" then    ' 무통장
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "무통장입금")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "입금전 상태")
        elseif paymethod = "100" then   ' 신용카드
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "신용카드")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "결제완료")
        elseif paymethod = "20" then   ' 실시간이체
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "실시간이체")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "결제완료")
        elseif paymethod = "80" then   ' 올앳
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "올앳카드")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "결제완료")
        elseif paymethod = "110" then   ' OKCashbag+신용카드
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "OKCashbag+신용카드")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "결제완료")
        elseif paymethod = "400" then   ' 핸드폰결제
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "핸드폰")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "결제완료")
        else
        	mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "")
        end if

        if (paymethod<>"7") then
            mailcontent = ReplaceText(mailcontent,"(<!-----bankinfo------>)[\s\S]*(<!-----/bankinfo------>)","")
            mailcontent = ReplaceText(mailcontent,"(<!-----banknotiinfo------>)[\s\S]*(<!-----/banknotiinfo------>)","")
        end if

        IsForeighDeliver = myorder.FOneItem.IsForeignDeliver

        if (IsForeighDeliver) then
            mailcontent = replace(mailcontent,":REQHPORREQEMAIL:", "이메일") ' 수령인 이메일
            mailcontent = replace(mailcontent,":REQHP:", myorder.FOneItem.Freqemail) ' 수령인 전화번호=>이메일로
            mailcontent = replace(mailcontent,":COUNTRYNAME:", myorder.FOneItem.FcountryNameEn) ' 국가.
            mailcontent = replace(mailcontent,":REQZIPCODE:", myorder.FOneItem.FemsZipCode) ' 배송우편번호
        else
            mailcontent = replace(mailcontent,":REQHPORREQEMAIL:", "휴대폰번호") ' 휴대폰번호
            mailcontent = replace(mailcontent,":REQHP:", myorder.FOneItem.Freqhp) ' 수령인 전화번호
            mailcontent = replace(mailcontent,":REQZIPCODE:", myorder.FOneItem.Freqzipcode) ' 배송우편번호
            mailcontent = ReplaceText(mailcontent,"(<!-- foreigndelivery -->)[\s\S]*(<!--/foreigndelivery -->)","")
        end if

        mailcontent = replace(mailcontent,":BUYNAME:", myorder.FOneItem.Fbuyname) ' 주문자 이름
        mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' 주문번호
        mailcontent = replace(mailcontent,":REQNAME:", myorder.FOneItem.Freqname) ' 수령인 이름
        mailcontent = replace(mailcontent,":REQALLADDRESS:", myorder.FOneItem.FreqZipaddr + " " + myorder.FOneItem.Freqaddress) ' 배송주소
        mailcontent = replace(mailcontent,":REQPHONE:", myorder.FOneItem.Freqphone) ' 수령인 전화번호

        mailcontent = replace(mailcontent,":BEASONGMEMO:", myorder.FOneItem.Fcomment) ' 배송메모


    	if (paymethod="110") then
    	    mailcontent = replace(mailcontent,":MAJORTOTALPRICE:", formatNumber(myorder.FOneItem.TotalMajorPaymentPrice,0) & " (신용카드:" &FormatNumber(myorder.FOneItem.TotalMajorPaymentPrice-myorder.FOneItem.FokcashbagSpend,0)& ",  OKCashbag:" &FormatNumber(myorder.FOneItem.FokcashbagSpend,0) &")") ' 결제총액
    	else
    	    mailcontent = replace(mailcontent,":MAJORTOTALPRICE:", formatNumber(myorder.FOneItem.TotalMajorPaymentPrice,0)) ' 결제총액
        end if

        mailcontent = replace(mailcontent,":ACCOUNTNO:", myorder.FOneItem.Faccountno) ' 입금계좌

        if (myorder.FOneItem.FsumPaymentEtc<>0) then
            mailcontent = replace(mailcontent,":SPENDTENCASH:", FormatNumber(myorder.FOneItem.FsumPaymentEtc,0))
        else
            mailcontent = ReplaceText(mailcontent,"(<!-----spendtencash------>)[\s\S]*(<!-----/spendtencash------>)","")
        end if


		'주문아이템 정보 확인.-----------------------------------------------------------------------------


itemHtml = itemHtml + "<table width='100%' border='0' cellspacing='0' cellpadding='0' style='border-top:3px solid #be0808;'>"
itemHtml = itemHtml + "  <tr>"
itemHtml = itemHtml + "    <td height='30' style='background:#fcf6f6; border-bottom:1px solid #eaeaea;'>"
itemHtml = itemHtml + "    <table width='100%' border='0' cellspacing='0' cellpadding='0' style='font-family:Dotum; font-size:11px; color:#888;padding-top:3px; '>"
itemHtml = itemHtml + "        <tr align='center'>"
itemHtml = itemHtml + "          <td width='60'>상품</td>"
itemHtml = itemHtml + "          <td width='60'>상품코드</td>"
itemHtml = itemHtml + "          <td>상품명 [옵션]</td>"
itemHtml = itemHtml + "          <td width='90'>판매가격</td>"
itemHtml = itemHtml + "          <td width='40'>수량</td>"
itemHtml = itemHtml + "          <td width='90'>주문금액</td>"
itemHtml = itemHtml + "        </tr>"
itemHtml = itemHtml + "      </table></td>"
itemHtml = itemHtml + "  </tr>"


        for i=0 to myorderdetail.FResultCount-1
        	if myorderdetail.FItemList(i).FItemID <> 0 then

itemHtml = itemHtml + "  <tr>"
itemHtml = itemHtml + "    <td height='80' style='border-bottom:1px solid #eaeaea;'><table width='100%' border='0' cellspacing='0' cellpadding='0' style='font-family:Dotum; font-size:11px; color:#888; padding-top:3px;'>"
itemHtml = itemHtml + "        <tr align='center'>"
itemHtml = itemHtml + "          <td width='60'><img src='" &  myorderdetail.FItemList(i).FSmallImage & "' width='50' height='50'></td>"
itemHtml = itemHtml + "          <td width='60' style='font-family:Dotum; font-size:11px; color:#888; text-decoration:none;'>"& myorderdetail.FItemList(i).FItemID&"</td>"
itemHtml = itemHtml + "          <td style='text-align:left; line-height:16px; padding-left:5px;'><span style='font-family: Verdana; font-size: 11px; color: #aaaaaa; text-decoration:none;'>["&myorderdetail.FItemList(i).Fmakerid& "]</span><br>"
itemHtml = itemHtml + "            <span style='font-family:Dotum; font-size:11px; color:#888; text-decoration:none;'>" + myorderdetail.FItemList(i).FItemName + "</span>"
if ( myorderdetail.FItemList(i).FItemOptionName <>"") then
itemHtml = itemHtml + "            [<span style='color:#1545f9'>"&myorderdetail.FItemList(i).FItemOptionName&"</span>]<br>"
end if
itemHtml = itemHtml + "          </td>"
itemHtml = itemHtml + "          <td width='90'>"

if (myorderdetail.FItemList(i).Fissailitem = "Y") then
    itemHtml = itemHtml + "          <strike>"&FormatNumber(myorderdetail.FItemList(i).Forgitemcost,0)&"</strike><br>"
    itemHtml = itemHtml + "          <span style='font-family: 돋움; FONT-SIZE: 11px; COLOR: #c20a0a; font-weight:bold;'><img src='http://fiximage.10x10.co.kr/web2008/shoppingbag/icon_sale.gif' width=14>"&FormatNumber(myorderdetail.FItemList(i).getItemcostCouponNotApplied,0)&"</span>"
else
    if (Not IsNull(myorderdetail.FItemList(i).Fitemcouponidx)) then
    itemHtml = itemHtml + "          <strike>"&FormatNumber(myorderdetail.FItemList(i).FitemcostCouponNotApplied,0)&"</strike>"
    else
    itemHtml = itemHtml + "          "&FormatNumber(myorderdetail.FItemList(i).FitemcostCouponNotApplied,0)
    end if
end if
    itemHtml = itemHtml + "          "&"원"

if (Not IsNull(myorderdetail.FItemList(i).Fitemcouponidx)) then
    itemHtml = itemHtml + "          <br><span style='font-family: 돋움; FONT-SIZE: 11px; font-weight:bold ; COLOR: #438938;' >"&FormatNumber(myorderdetail.FItemList(i).FItemCost,0)& "원" &"</span>"
end if

itemHtml = itemHtml + "          </td>"
itemHtml = itemHtml + "          <td width='40'>" &myorderdetail.FItemList(i).FItemNo& "</td>"
itemHtml = itemHtml + "          <td width='90'>" &FormatNumber(myorderdetail.FItemList(i).FItemCost*myorderdetail.FItemList(i).FItemNo,0) & "원" & "</td>"
itemHtml = itemHtml + "        </tr>"
itemHtml = itemHtml + "      </table></td>"
itemHtml = itemHtml + "  </tr>"
			end if
        next

itemHtml = itemHtml + "</table>"


		itemHtmlTotal = replace(mailcontent,":INNERORDERTABLE:", itemHtml) ' 주문정보테이블 넣기

        mailcontent = itemHtmlTotal

ttlsumHTML = ""
ttlsumHTML = ttlsumHTML + "<table width='100%' height='50' border='0' cellpadding='0' cellspacing='1' bgcolor='#eaeaea'>"
ttlsumHTML = ttlsumHTML + "  <tr>"
ttlsumHTML = ttlsumHTML + "    <td bgcolor='#FFFFFF' style='border:3px solid #f3f3f3;'>"
ttlsumHTML = ttlsumHTML + "    <table width='100%' border='0' cellspacing='0' cellpadding='0'>"
ttlsumHTML = ttlsumHTML + "      <tr>"
ttlsumHTML = ttlsumHTML + "        <td style='padding-right:10px;padding-left:10px;'>"
ttlsumHTML = ttlsumHTML + "        <table width='100%' border='0' cellspacing='0' cellpadding='0'>"
ttlsumHTML = ttlsumHTML + "          <tr>"
ttlsumHTML = ttlsumHTML + "            <td align='right' style='padding:10px 5px 10px 5px;'>"
ttlsumHTML = ttlsumHTML + "            <table border='0' cellspacing='0' cellpadding='0'>"
ttlsumHTML = ttlsumHTML + "              <tr>"
ttlsumHTML = ttlsumHTML + "                <td style='padding-left:30px;' style='font-family: 돋움; FONT-SIZE: 12px; COLOR: #888888;'></td>"
ttlsumHTML = ttlsumHTML + "                <td style='padding-left:30px;' style='font-family: 돋움; FONT-SIZE: 12px; COLOR: #888888;'>적립 마일리지 <strong>"& FormatNumber(myorder.FOneItem.Ftotalmileage,0) &"</strong> Point</td>"
ttlsumHTML = ttlsumHTML + "                <td style='padding-left:30px;' style='font-family: 돋움; FONT-SIZE: 12px; COLOR: #888888;'>상품구매총액 <strong>"& FormatNumber(myorder.FOneItem.FTotalSum-myorderdetail.BeasongPay,0) &"</strong> 원</td>"
ttlsumHTML = ttlsumHTML + "              </tr>"
ttlsumHTML = ttlsumHTML + "            </table></td>"
ttlsumHTML = ttlsumHTML + "          </tr>"
ttlsumHTML = ttlsumHTML + "          <tr height='1'>"
ttlsumHTML = ttlsumHTML + "            <td height='1' bgcolor='#eaeaea'></td>"
ttlsumHTML = ttlsumHTML + "          </tr>"
ttlsumHTML = ttlsumHTML + "          <tr height='30'>"
ttlsumHTML = ttlsumHTML + "            <td align='right' style='padding:10px 5px 10px 5px;' style='font-family: 돋움; FONT-SIZE: 12px; COLOR: #000000;'>"
ttlsumHTML = ttlsumHTML + "         총결제금액 : "
ttlsumHTML = ttlsumHTML + "			상품구매총액 "& FormatNumber((myorder.FOneItem.FTotalSum-myorderdetail.BeasongPay),0) &"원 "
ttlsumHTML = ttlsumHTML + "			+ 배송비 "& FormatNumber(myorderdetail.BeasongPay,0) &"원"

IF (myorder.FOneItem.Fmiletotalprice<>0) then
    ttlsumHTML = ttlsumHTML + "			- 마일리지 "& FormatNumber(myorder.FOneItem.Fmiletotalprice,0) &"원 "
end if
IF (myorder.FOneItem.Ftencardspend<>0) then
    ttlsumHTML = ttlsumHTML + "			- 보너스쿠폰할인 "&FormatNumber(myorder.FOneItem.Ftencardspend,0) &"원 "
end if

if (myorder.FOneItem.Fallatdiscountprice + myorder.FOneItem.Fspendmembership<>0) then
    ttlsumHTML = ttlsumHTML + "			- 기타할인 "& FormatNumber((myorder.FOneItem.Fallatdiscountprice + myorder.FOneItem.Fspendmembership),0) &"원 "
end if
ttlsumHTML = ttlsumHTML + "			= "
ttlsumHTML = ttlsumHTML + "			<span style='font-family: 돋움; FONT-SIZE: 12px; COLOR: #c20a0a; font-weight:bold;'>"& FormatNumber(myorder.FOneItem.FsubtotalPrice,0) &"</span> 원"
ttlsumHTML = ttlsumHTML + "            </td>"
ttlsumHTML = ttlsumHTML + "          </tr>"
ttlsumHTML = ttlsumHTML + "        </table></td>"
ttlsumHTML = ttlsumHTML + "      </tr>"
ttlsumHTML = ttlsumHTML + "    </table></td>"
ttlsumHTML = ttlsumHTML + "  </tr>"
ttlsumHTML = ttlsumHTML + "</table>"

        mailcontent = replace(mailcontent,":ORDERPRICESUMMARY:", ttlsumHTML) ' 주문 합계금액



        set myorder = Nothing
        set myorderDetail = Nothing

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
end function

sub dsendmail(mailfrom, mailto, mailtitle, mailcontent)
        dim mailobject

        set mailobject=server.createobject("CDONTS.NewMail")
        mailobject.from = mailfrom
        mailobject.to = mailto
        mailobject.subject = mailtitle

        'html style
        mailobject.bodyformat = 0
        mailobject.mailformat = 0

        mailobject.body = mailcontent
        mailobject.send
        set mailobject = nothing
end sub

function sendmailCS(mailto, title, contents)
        dim mailfrom, mailtitle, mailcontent,dirPath,fileName
        dim fs,objFile

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "[10x10] " + title

        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_cs.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":CONTENTS:",contents)

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)

end function

function sendmailnewuser2(mailto,userName) ' 가입메일파일을 읽어들이는 방식으로 전환
        dim mailfrom, mailtitle, mailcontent,dirPath,fileName
        dim fs,objFile

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "10x10 사이트 가입을 축하 드립니다."

        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_join.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":USERNAME:",userName)

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailnewuser2 = mailcontent
end function

sub sendmailnewuser(mailto) ' 위 function으로 전환함.20020329/
        dim mailfrom, mailtitle, mailcontent

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "10x10 사이트 가입을 축하 드립니다."

        '이부분이 메일 html 임. 따로 프로그램 들어가는것은 없는상태...
        mailcontent	= "<HTML>																													"	_
        +"	<HEAD><TITLE>Thank you for Join at Member of 10X10 Design Group</TITLE>															"	_
        +"<link rel=stylesheet href=http://www.10x10.co.kr/css/main.css type=text/css>														"	_
        +"</HEAD>																															"	_
        +"<body bgcolor=#FFFFFF text=#000000 leftmargin=0 topmargin=0 marginwidth=0 marginheight=0>											"	_
        +"<table width=100% border=0 background=http://www.10x10.co.kr/images/emailtop_bg.gif height=220>											"	_
        +"  <tr>																															"	_
        +"	<td height=75 valign=top align=left width=500><img src=http://www.10x10.co.kr/images/top_sitelogo.gif width=282 height=145></td>	"	_
        +"    <td valign=top rowspan=2 width=80><img src=http://www.10x10.co.kr/images/top_people.gif width=80 height=217></td>				"	_
        +"    <td rowspan=2 align=right valign=top width=49><img src=http://www.10x10.co.kr/images/top_flower.gif width=152 height=197></td>	"	_
        +"  </tr>																															"	_
        +"  <tr>																															"	_
        +"    <td valign=top align=right width=500><img src=http://www.10x10.co.kr/images/1_1_white.gif width=150 height=1><img src=http://www.10x10.co.kr/images/join_ment.gif width=350 height=50></td>"	_
        +"  </tr>"	_
        +" </table>"	_
        +"<div align=center><br>"	_
        +"  <table width=646 border=0 cellpadding=0 cellspacing=0>"	_
        +"    <tr> "	_
        +"      <td> <img src=http://www.10x10.co.kr/images/slice01_01.gif width=20 height=19></td>"	_
        +"      <td bgcolor=F1F1F1>&nbsp; </td>"	_
        +"      <td> <img src=http://www.10x10.co.kr/images/slice01_03.gif width=26 height=19></td>"	_
        +"    </tr>"	_
        +"    <tr> "	_
        +"      <td rowspan=3 bgcolor=F1F1F1>&nbsp; </td>"	_
        +"      <td bgcolor=F1F1F1> "	_
        +"        <p><font face=verdana size=1><img src=http://www.10x10.co.kr/images/icon_basic.gif width=20 height=20><b>tenbyten</b> "	_
        +"          since 2001.10.10</font></p>"	_
        +"			<p>디자인 전문 사이트 10X10.co.kr (텐바이텐) 에 가입해 주셔서 진심으로 감사드립니다.<br><br>"	_
        +"			   저희 10X10 은 쇼핑몰과 커뮤니티가 결합된 디자인 채널로서 <br><br>"	_
        +"			   디자인을 쉽고 재밌게 즐길수 있는 사이트 입니다.<br><br>"	_
        +"          항상 행복한 일들이 회원 여러분께 가득하길 바랍니다... : )</p>"	_
        +"        <p><br>"	_
        +"        </p>"	_
        +"      </td>"	_
        +"      <td rowspan=3 bgcolor=F1F1F1>&nbsp; </td>"	_
        +"    </tr>"	_
        +"    <tr> "	_
        +"      <td> <img src=http://www.10x10.co.kr/images/slice01_07.gif width=600 height=4></td>"	_
        +"    </tr>"	_
        +"    <tr> "	_
        +"      <td bgcolor=F1F1F1><img src=http://www.10x10.co.kr/images/slice01_08.gif width=367 height=53> "	_
        +"      </td>"	_
        +"    </tr>"	_
        +"    <tr> "	_
        +"      <td bgcolor=F1F1F1> <img src=http://www.10x10.co.kr/images/slice01_09.gif width=20 height=23></td>"	_
        +"      <td bgcolor=F1F1F1>&nbsp; </td>"	_
        +"      <td> <img src=http://www.10x10.co.kr/images/slice01_11.gif width=26 height=23></td>"	_
        +"    </tr>"	_
        +"  </table>"	_
        +"  <br>"	_
        +"  <table width=646 border=0>"	_
        +"    <tr> "	_
        +"      <td align=right valign=top>(주)큐브 커뮤니모우션<img src=http://www.10x10.co.kr/images/cube_ci.gif width=210 height=52 hspace=15></td>"	_
        +"    </tr>"	_
        +"  </table>"	_
        +"</div>"	_
        +"</body>"	_
        +"</html>	"

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
end sub

sub sendmailorder(orderserial)
        dim sql,discountrate
        dim mailfrom, mailto, mailtitle, mailcontent

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "주문이 정상적으로 접수되었습니다!"

        '주문자 메일주소 확인.
        sql = "select buyemail from tbl_order_master where orderserial = '" + orderserial + "'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                rsfunc.Movefirst
                mailto = rsfunc("buyemail")
        else
                exit sub
        end if
        rsfunc.close

        mailcontent = "<HTML> " + vbcr
        mailcontent = mailcontent + "<HEAD><TITLE>Thank you for Join at Member of 10X10 Design Group</TITLE> " + vbcr
        mailcontent = mailcontent + "<link rel=stylesheet href=http://www.10x10.co.kr/css/main.css type=text/css> " + vbcr
        mailcontent = mailcontent + "</HEAD> " + vbcr
        mailcontent = mailcontent + "<body bgcolor=#FFFFFF text=#000000 leftmargin=0 topmargin=0 marginwidth=0 marginheight=0> " + vbcr
        mailcontent = mailcontent + "<table width=100% border=0 background=http://www.10x10.co.kr/images/emailtop_bg.gif height=220> " + vbcr
        mailcontent = mailcontent + "  <tr> " + vbcr
        mailcontent = mailcontent + "	<td height=75 valign=top align=left width=500><img src=http://www.10x10.co.kr/images/top_sitelogo.gif width=282 height=145></td> " + vbcr
        mailcontent = mailcontent + "    <td valign=top rowspan=2 width=80><img src=http://www.10x10.co.kr/images/top_people.gif width=80 height=217></td> " + vbcr
        mailcontent = mailcontent + "    <td rowspan=2 align=right valign=top width=49><img src=http://www.10x10.co.kr/images/top_flower.gif width=152 height=197></td> " + vbcr
        mailcontent = mailcontent + "  </tr> " + vbcr
        mailcontent = mailcontent + "  <tr> " + vbcr
        mailcontent = mailcontent + "    <td valign=top align=right width=500><img src=http://www.10x10.co.kr/images/1_1_white.gif width=150 height=1><img src=http://www.10x10.co.kr/images/order_ment.gif width=350 height=50></td> " + vbcr
        mailcontent = mailcontent + "  </tr> " + vbcr
        mailcontent = mailcontent + " </table> " + vbcr
        mailcontent = mailcontent + "<div align=center><br> " + vbcr
        mailcontent = mailcontent + "  <table width=646 border=0 cellpadding=0 cellspacing=0> " + vbcr
        mailcontent = mailcontent + "    <tr> " + vbcr
        mailcontent = mailcontent + "      <td> <img src=http://www.10x10.co.kr/images/slice01_01.gif width=20 height=19></td> " + vbcr
        mailcontent = mailcontent + "      <td bgcolor=F1F1F1>&nbsp; </td> " + vbcr
        mailcontent = mailcontent + "      <td> <img src=http://www.10x10.co.kr/images/slice01_03.gif width=26 height=19></td> " + vbcr
        mailcontent = mailcontent + "    </tr> " + vbcr
        mailcontent = mailcontent + "    <tr> " + vbcr
        mailcontent = mailcontent + "      <td rowspan=5 bgcolor=F1F1F1>&nbsp; </td> " + vbcr
        mailcontent = mailcontent + "      <td bgcolor=F1F1F1>저희 텐바이텐 사이트를 이용해 주셔서 진심으로 감사드리며 다음 주문이 정상적으로 접수되어 처리중입니다. " + vbcr
        mailcontent = mailcontent + "        <br> 신용카드 결제는 주문접수 후 곧바로 배송에 들어가면 온라인입금 주문은 입금확인 후 배송이 이루어 집니다." + vbcr
        mailcontent = mailcontent + "        <br> (온라인 입금하실 곳은 조흥은행 / 534-01-016039 / (주)큐브 커뮤니모우션 입니다.)" + vbcr
        mailcontent = mailcontent + "        배송은 약 2일에서 4일 가량이 소요되며, 주문정보에 대한 변동사항이나 문의사항은 <br> " + vbcr
        mailcontent = mailcontent + "        이메일(<a href=mailto:customer@10X10.co.kr>customer@10X10.co.kr</a>)이나 02-515-5945로 " + vbcr
        mailcontent = mailcontent + "        연락주시기 바랍니다.<br> " + vbcr
        mailcontent = mailcontent + "        <br> " + vbcr
        mailcontent = mailcontent + "      </td> " + vbcr
        mailcontent = mailcontent + "      <td rowspan=5 bgcolor=F1F1F1></td> " + vbcr
        mailcontent = mailcontent + "    </tr> " + vbcr
        mailcontent = mailcontent + "    <tr> " + vbcr
        mailcontent = mailcontent + "      <td bgcolor=F1F1F1><img src=http://www.10x10.co.kr/images/slice01_07.gif width=600 height=4></td> " + vbcr
        mailcontent = mailcontent + "    </tr> " + vbcr
        mailcontent = mailcontent + "    <tr> " + vbcr
        mailcontent = mailcontent + "      <td bgcolor=F1F1F1><br> " + vbcr

        '주문정보 확인.
        sql = "select regdate, reqzipcode, (b.addr010_si + ' ' + b.addr010_gu + ' ' + a.reqaddress) as reqalladdress, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice from tbl_order_master a, addr010tl b, tbl_order_detail c"
        sql = sql + " where b.addr010_zip1 = left(a.reqzipcode,3) and b.addr010_zip2 = right(a.reqzipcode,3) and a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                discountrate = rsfunc("discountrate")
                rsfunc.Movefirst
                mailcontent = mailcontent + "  <table width=600 border=0> " + vbcr
                mailcontent = mailcontent + "    <tr> " + vbcr
                mailcontent = mailcontent + "      <td><img src=http://www.10x10.co.kr/images/order_ment02.gif width=150 height=35 vspace=5 hspace=0></td> " + vbcr
                mailcontent = mailcontent + "            <td><font color=990000>주문 번호 : " + orderserial + " &nbsp;&nbsp;|&nbsp;주문 일자 : " + cStr(year(rsfunc("regdate"))) + "년 " + cStr(month(rsfunc("regdate"))) + "월 " + cStr(day(rsfunc("regdate"))) + "일<br> " + vbcr
                mailcontent = mailcontent + "              배 송 지 : [" + rsfunc("reqzipcode") + "] " + rsfunc("reqalladdress") + "<br> " + vbcr
                mailcontent = mailcontent + "        주문 총액 : " + cstr(rsfunc("subtotalprice")) + "원 = 소계 : " + cstr(rsfunc("subtotalprice") - rsfunc("itemcost")) + "원 (" + cstr(rsfunc("totalmileage")) + "포인트) + 배송비 : " + cstr(rsfunc("itemcost")) + "원</font> </td> " + vbcr
                mailcontent = mailcontent + "    </tr> " + vbcr
                mailcontent = mailcontent + "  </table> " + vbcr
        else
                exit sub
        end if
        rsfunc.close

        '주문아이템 정보 확인.
        dim itemserial
        sql = " select a.itemid, b.imglist, c.itemname, (c.cate_large + c.cate_mid + c.cate_small) as itemserial, c.sellcash, a.itemno from tbl_order_detail a, tbl_item_image b, tbl_item c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and a.itemid <> '0' and a.itemid = b.itemid and c.itemid = a.itemid"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                rsfunc.Movefirst
                do until rsfunc.eof
                        itemserial = rsfunc("itemserial") + "-" + FormatCode(rsfunc("itemid"))

                        mailcontent	= mailcontent + "        <table width=300 border=0> " + vbcr
                        mailcontent	= mailcontent + "          <tr> " + vbcr
                        mailcontent	= mailcontent + "            <td width=100><img src=http://www.10x10.co.kr/image/list/" + rsfunc("imglist") + " width=100 height=100></td> " + vbcr
                        mailcontent	= mailcontent + "            <td> " + vbcr
                        mailcontent	= mailcontent + "              <table border=0 cellspacing=0 cellpadding=3> " + vbcr
                        mailcontent	= mailcontent + "                <tr>  " + vbcr
                        mailcontent	= mailcontent + "                  <td class=ggg width=60><font face='Verdana, Arial, Helvetica, sans-serif' size=1>Product</font></td> " + vbcr
                        mailcontent	= mailcontent + "                  <td width=120 class=text1>" + rsfunc("itemname") + "</td> " + vbcr
                        mailcontent	= mailcontent + "                </tr> " + vbcr
                        mailcontent	= mailcontent + "                <tr>  " + vbcr
                        mailcontent	= mailcontent + "                  <td class=ggg height=2><font face='Verdana, Arial, Helvetica, sans-serif' size=1>Code</font></td> " + vbcr
                        mailcontent	= mailcontent + "                  <td class=ggg height=2><font size=1 face='Verdana, Arial, Helvetica, sans-serif'>" + vbcr
                        mailcontent	= mailcontent + itemserial + "</font></td> " + vbcr
                        mailcontent	= mailcontent + "                </tr> " + vbcr
                        mailcontent	= mailcontent + "                <tr>  " + vbcr
                        mailcontent	= mailcontent + "                  <td class=ggg><font face='Verdana, Arial, Helvetica, sans-serif' size=1>Price</font></td> " + vbcr
                        mailcontent	= mailcontent + "                  <td class=ggg><font size=1 face='Verdana, Arial, Helvetica, sans-serif'>" + cstr(rsfunc("sellcash")*cdbl(discountrate)) + "won</font></td> " + vbcr
                        mailcontent	= mailcontent + "                </tr> " + vbcr
                        'mailcontent	= mailcontent + "                <tr>  " + vbcr
                        'mailcontent	= mailcontent + "                  <td class=ggg><font face='Verdana, Arial, Helvetica, sans-serif' size=1>Option</font></td> " + vbcr
                        'mailcontent	= mailcontent + "                  <td class=ggg> <font size=1 face='Verdana, Arial, Helvetica, sans-serif'>" + vbcr

                        '옵션 표시부분. 일단 생략.

                        'mailcontent	= mailcontent + "                    </font></td> " + vbcr
                        'mailcontent	= mailcontent + "                </tr> " + vbcr
                        mailcontent	= mailcontent + "                <tr>  " + vbcr
                        mailcontent	= mailcontent + "                  <td class=ggg><font face='Verdana, Arial, Helvetica, sans-serif' size=1>Quantity</font></td> " + vbcr
                        mailcontent	= mailcontent + "                  <td class=ggg> <font size=1 face='Verdana, Arial, Helvetica, sans-serif'> " + cstr(rsfunc("itemno")) + " EA </font></td> " + vbcr
                        mailcontent	= mailcontent + "                </tr> " + vbcr
                        mailcontent	= mailcontent + "              </table> " + vbcr
                        mailcontent	= mailcontent + "            </td> " + vbcr
                        mailcontent	= mailcontent + "          </tr> " + vbcr
                        mailcontent	= mailcontent + "        </table> " + vbcr
                rsfunc.movenext
                loop
        else
                exit sub
        end if
        rsfunc.close

        mailcontent = mailcontent + "      </td> "
        mailcontent = mailcontent + "    </tr> "
        mailcontent = mailcontent + "    <tr> "
        mailcontent = mailcontent + "      <td> <img src=http://www.10x10.co.kr/images/slice01_07.gif width=600 height=4></td> "
        mailcontent = mailcontent + "    </tr> "
        mailcontent = mailcontent + "    <tr> "
        mailcontent = mailcontent + "      <td bgcolor=F1F1F1><img src=http://www.10x10.co.kr/images/slice01_08.gif width=367 height=53> "
        mailcontent = mailcontent + "      </td> "
        mailcontent = mailcontent + "    </tr> "
        mailcontent = mailcontent + "    <tr> "
        mailcontent = mailcontent + "      <td bgcolor=F1F1F1> <img src=http://www.10x10.co.kr/images/slice01_09.gif width=20 height=23></td> "
        mailcontent = mailcontent + "      <td bgcolor=F1F1F1>&nbsp; </td> "
        mailcontent = mailcontent + "      <td> <img src=http://www.10x10.co.kr/images/slice01_11.gif width=26 height=23></td> "
        mailcontent = mailcontent + "    </tr> "
        mailcontent = mailcontent + "  </table> "
        mailcontent = mailcontent + "  <br> "
        mailcontent = mailcontent + "  <table width=646 border=0> "
        mailcontent = mailcontent + "    <tr> "
        mailcontent = mailcontent + "      <td align=right valign=top>(주)큐브 커뮤니모우션<img src=http://www.10x10.co.kr/images/cube_ci.gif width=210 height=52 hspace=15></td> "
        mailcontent = mailcontent + "    </tr> "
        mailcontent = mailcontent + "  </table> "
        mailcontent = mailcontent + "</div> "
        mailcontent = mailcontent + "</body> "
        mailcontent = mailcontent + "</html> "

        'response.write mailcontent
        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
end sub

function sendmailorder2(orderserial)
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "주문이 정상적으로 접수되었습니다!"

        '주문자 메일주소 확인,주문거래종류 선택
        sql = "select buyemail,accountdiv from tbl_order_master where orderserial = '" + orderserial + "'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                rsfunc.Movefirst
                mailto = rsfunc("buyemail")
                paymethod = trim(rsfunc("accountdiv"))
        else
                exit function
        end if
        rsfunc.close

        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        if paymethod = "7" then    ' 무통장
            fileName = dirPath&"\\email_bank1.htm"
        elseif paymethod = "100" then   ' 신용카드
            fileName = dirPath&"\\email_card1.htm"
        end if

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
'        mailcontent = replace(mailcontent,":USERNAME:",userName)



        '주문정보 확인.
        sql = "select buyname,regdate, reqzipcode, (b.addr010_si + ' ' + b.addr010_gu + ' ' + a.reqaddress) as reqalladdress, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice from tbl_order_master a, addr010tl b, tbl_order_detail c"
        sql = sql + " where b.addr010_zip1 = left(a.reqzipcode,3) and b.addr010_zip2 = right(a.reqzipcode,3) and a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                discountrate = rsfunc("discountrate")
                rsfunc.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", FormatCurrency(cstr(rsfunc("subtotalprice")))) ' 주문총액
                mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(rsfunc("subtotalprice") - rsfunc("itemcost"))) ) ' 주문한 총item  가격
                mailcontent = replace(mailcontent,":DELIVERYFEE:",  FormatCurrency(cstr(rsfunc("itemcost"))) ) ' 배송금액
                mailcontent = replace(mailcontent,":BUYNAME:", rsfunc("buyname")) ' 주문자 이름
                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' 주문번호
                mailcontent = replace(mailcontent,":REQZIPCODE:", rsfunc("reqzipcode")) ' 배송우편번호
                mailcontent = replace(mailcontent,":REQALLADDRESS:", rsfunc("reqalladdress")) ' 배송주소
        else
                exit function
        end if
        rsfunc.close

        'item 루프 앞뒤부분 짜르기
        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item 루프를 돌릴부분 자르기
        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)

        '주문아이템 정보 확인.
        dim itemserial,inx
        sql = " select a.itemid, b.imglist, c.itemname, (c.cate_large + c.cate_mid + c.cate_small) as itemserial, c.sellcash, a.itemno from tbl_order_detail a, tbl_item_image b, tbl_item c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and a.itemid <> '0' and a.itemid = b.itemid and c.itemid = a.itemid"
        inx = 1
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                rsfunc.Movefirst
                do until rsfunc.eof
                        itemserial = rsfunc("itemserial") + "-" + FormatCode(rsfunc("itemid"))
                        itemHtml = replace(itemHtmlOri,":ITEMSERIAL:", itemserial) ' 상품코드
                        itemHtml = replace(itemHtml,":ITEMNAME:", rsfunc("itemname")) ' 상품이름
                        if discountrate=1 then
                        	itemHtml = replace(itemHtml,":ITEMPRICE:",  CStr(rsfunc("sellcash"))) ' 상품가격
                        else
                        	itemHtml = replace(itemHtml,":ITEMPRICE:",  cstr(round(rsfunc("sellcash")*cdbl(discountrate)/100)*100) ) ' 상품가격
                    	end if
                        itemHtml = replace(itemHtml,":ITEMNO:", cstr(rsfunc("itemno"))) ' 수량
                        itemHtml = replace(itemHtml,":IMGLIST:", cstr(rsfunc("imglist"))) ' 수량
                        if  inx mod 3 = 0 then
                            itemHtml = itemHtml + vbcr + "<tr></tr>"
                        end if
                        itemHtmlTotal = itemHtmlTotal & itemHtml

                inx = inx + 1
                rsfunc.movenext
                loop
        else
                exit function
        end if
        rsfunc.close

        mailcontent = beforeItemHtml & itemHtmlTotal & afterItemHtml



        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailorder2 = mailcontent
end function


function sendmailorder3(orderserial,mailfrom)
        dim sql,discountrate,paymethod
        dim mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal


        mailtitle = "주문이 정상적으로 접수되었습니다!"

        '주문자 메일주소 확인,주문거래종류 선택---------------------------------------------------------------------------
        sql = "select buyemail,accountdiv from tbl_order_master where orderserial = '" + orderserial + "'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                rsfunc.Movefirst
                mailto = rsfunc("buyemail")
                paymethod = trim(rsfunc("accountdiv"))
        else
                exit function
        end if
        rsfunc.close

        ' 파일을 불러와서 ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        if paymethod = "7" then    ' 무통장
            fileName = dirPath&"\\email_bank1.htm"
        elseif paymethod = "100" then   ' 신용카드
            fileName = dirPath&"\\email_card1.htm"
        end if

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
'        mailcontent = replace(mailcontent,":USERNAME:",userName)


		dim SpendMile
        '주문정보 확인.---------------------------------------------------------------------------
        sql = "select buyname,regdate, reqzipcode, (b.addr010_si + ' ' + b.addr010_gu + ' ' + a.reqaddress) as reqalladdress, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice, a.miletotalprice from tbl_order_master a, addr010tl b, tbl_order_detail c"
        sql = sql + " where b.addr010_zip1 = left(a.reqzipcode,3) and b.addr010_zip2 = right(a.reqzipcode,3) and a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                discountrate = rsfunc("discountrate")
                rsfunc.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", FormatCurrency(cstr(rsfunc("subtotalprice")))) ' 주문총액
                'mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(Titemcost - rsfunc("itemcost"))) ) ' 주문한 총item  가격
                mailcontent = replace(mailcontent,":DELIVERYFEE:",  FormatCurrency(cstr(rsfunc("itemcost"))) ) ' 배송금액
                mailcontent = replace(mailcontent,":BUYNAME:", rsfunc("buyname")) ' 주문자 이름
                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' 주문번호
                mailcontent = replace(mailcontent,":REQZIPCODE:", rsfunc("reqzipcode")) ' 배송우편번호
                mailcontent = replace(mailcontent,":REQALLADDRESS:", rsfunc("reqalladdress")) ' 배송주소

                if IsNull(rsfunc("miletotalprice")) then
                	SpendMile =""
                else
                	SpendMile = rsfunc("miletotalprice")
                	SpendMile = "(마일리지사용: " + formatNumber(FormatCurrency(SpendMile),0) + " )"
            	end if
            	mailcontent = replace(mailcontent,":SPENDMILEAGE:", SpendMile) ' 마일리지
        else
                exit function
        end if
        rsfunc.close

        'item 루프 앞뒤부분 짜르기
        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item 루프를 돌릴부분 자르기
        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)




		'주문아이템 정보 확인.-----------------------------------------------------------------------------
        dim itemserial,inx
        dim Titemcost,BufCost

        Titemcost = 0
        sql = " select a.itemid, b.imglist, c.itemname, (c.cate_large + c.cate_mid + c.cate_small) as itemserial, c.sellcash, a.itemno from tbl_order_detail a, tbl_item_image b, tbl_item c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and a.itemid <> '0' and a.itemid = b.itemid and c.itemid = a.itemid"
        sql = sql + " and (a.cancelyn='N' or a.cancelyn='A')"
        inx = 1
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                rsfunc.Movefirst
                do until rsfunc.eof
                        itemserial = rsfunc("itemserial") + "-" + FormatCode(rsfunc("itemid"))
                        itemHtml = replace(itemHtmlOri,":ITEMSERIAL:", itemserial) ' 상품코드
                        itemHtml = replace(itemHtml,":ITEMNAME:", rsfunc("itemname")) ' 상품이름

                        if CDbl(discountrate)=1 then
                        	BufCost = rsfunc("sellcash") * rsfunc("itemno")
                        	Titemcost = Titemcost + BufCost
                        	itemHtml = replace(itemHtml,":ITEMPRICE:", CStr(BufCost) ) ' 상품가격
                        else
                        	BufCost = round(rsfunc("sellcash")*cdbl(discountrate)/100)*100 * rsfunc("itemno")
                        	Titemcost = Titemcost + BufCost
                        	itemHtml = replace(itemHtml,":ITEMPRICE:", CStr(BufCost) ) ' 상품가격
                    	end if
                        itemHtml = replace(itemHtml,":ITEMNO:", cstr(rsfunc("itemno"))) ' 수량
                        itemHtml = replace(itemHtml,":IMGLIST:", cstr(rsfunc("imglist"))) ' 수량
                        if  inx mod 3 = 0 then
                            itemHtml = itemHtml + vbcr + "<tr></tr>"
                        end if
                        itemHtmlTotal = itemHtmlTotal & itemHtml

                inx = inx + 1
                rsfunc.movenext
                loop
        else
                exit function
        end if
        rsfunc.close

        mailcontent = beforeItemHtml & itemHtmlTotal & afterItemHtml

		mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(Titemcost)) ) ' 주문한 총item  가격

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailorder3 = mailcontent
end function

function ReSendmailorder(orderserial,mailfrom)
        sendmailorder3 orderserial,mailfrom
end function

function sendmailcome(orderserial) ' 직접수령시 메일 보내기
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "10X10 수령 안내 메일입니다!"

        '주문자 메일주소 확인,주문거래종류 선택
        sql = "select buyemail,accountdiv from tbl_order_master where orderserial = '" + orderserial + "'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                rsfunc.Movefirst
                mailto = rsfunc("buyemail")
        else
                exit function
        end if
        rsfunc.close

        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_come.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall

        '주문정보 확인.
        sql = "select buyname,regdate, reqzipcode, (b.addr010_si + ' ' + b.addr010_gu + ' ' + a.reqaddress) as reqalladdress, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice from tbl_order_master a, addr010tl b, tbl_order_detail c"
        sql = sql + " where b.addr010_zip1 = left(a.reqzipcode,3) and b.addr010_zip2 = right(a.reqzipcode,3) and a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                discountrate = rsfunc("discountrate")
                rsfunc.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", FormatCurrency(cstr(rsfunc("subtotalprice")))) ' 주문총액
                mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(rsfunc("subtotalprice") - rsfunc("itemcost"))) ) ' 주문한 총item  가격
                mailcontent = replace(mailcontent,":DELIVERYFEE:",  FormatCurrency(cstr(rsfunc("itemcost"))) ) ' 배송금액
                mailcontent = replace(mailcontent,":BUYNAME:", rsfunc("buyname")) ' 주문자 이름
                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' 주문번호
                mailcontent = replace(mailcontent,":REQZIPCODE:", rsfunc("reqzipcode")) ' 배송우편번호
                mailcontent = replace(mailcontent,":REQALLADDRESS:", rsfunc("reqalladdress")) ' 배송주소
        else
                exit function
        end if
        rsfunc.close

        'item 루프 앞뒤부분 짜르기
        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item 루프를 돌릴부분 자르기
        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)

        '주문아이템 정보 확인.
        dim itemserial,inx
        sql = " select a.itemid, b.imglist, c.itemname, (c.cate_large + c.cate_mid + c.cate_small) as itemserial, c.sellcash, a.itemno from tbl_order_detail a, tbl_item_image b, tbl_item c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and a.itemid <> '0' and a.itemid = b.itemid and c.itemid = a.itemid"
        sql = sql + " and (a.cancelyn='N' or a.cancelyn='A')"
        inx = 1
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                rsfunc.Movefirst
                do until rsfunc.eof
                        itemserial = rsfunc("itemserial") + "-" + FormatCode(rsfunc("itemid"))
                        itemHtml = replace(itemHtmlOri,":ITEMSERIAL:", itemserial) ' 상품코드
                        itemHtml = replace(itemHtml,":ITEMNAME:", rsfunc("itemname")) ' 상품이름
                        itemHtml = replace(itemHtml,":ITEMPRICE:",  cstr(rsfunc("sellcash")*cdbl(discountrate)) ) ' 상품가격
                        itemHtml = replace(itemHtml,":ITEMNO:", cstr(rsfunc("itemno"))) ' 수량
                        itemHtml = replace(itemHtml,":IMGLIST:", cstr(rsfunc("imglist"))) ' 상품이미지
                        itemHtmlTotal = itemHtmlTotal & itemHtml

                inx = inx + 1
                rsfunc.movenext
                loop
        else
                exit function
        end if
        rsfunc.close

        mailcontent = beforeItemHtml & itemHtmlTotal & afterItemHtml

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailcome = mailcontent
end function

function sendmailbankok(mailto,userName,orderserial) ' 입금확인메일
        dim sql,discountrate
        dim mailfrom, mailtitle, mailcontent
        dim fs,objFile,dirPath,fileName

        mailfrom = "텐바이텐<customer@10x10.co.kr>"
        mailtitle = "무통장 입금이 정상적으로 처리 되었습니다!"

        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_bank2011.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":USERNAME:",userName)
        mailcontent = replace(mailcontent,":ORDERSERIAL:",orderserial)

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
end function

function sendmailbankokNoDLV(mailto,userName,orderserial) ' 입금확인메일
        dim sql,discountrate
        dim mailfrom, mailtitle, mailcontent
        dim fs,objFile,dirPath,fileName

        mailfrom = "텐바이텐<customer@10x10.co.kr>"
        mailtitle = "무통장 입금이 정상적으로 처리 되었습니다!"

        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_bank2011.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":USERNAME:",userName)
        mailcontent = replace(mailcontent,":ORDERSERIAL:",orderserial)
        mailcontent = replace(mailcontent,"빠른 시일내에 배송이 이루어 질 수 있도록 노력하겠습니다.","")

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
end function

function sendmailbankok_GIFTCard(mailto,userName,orderserial) ' 입금확인메일
        dim sql,discountrate
        dim mailfrom, mailtitle, mailcontent
        dim fs,objFile,dirPath,fileName

        mailfrom = "텐바이텐<customer@10x10.co.kr>"
        mailtitle = "무통장 입금이 정상적으로 처리 되었습니다!"

        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_bank2011_GiftCard.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":USERNAME:",userName)
        mailcontent = replace(mailcontent,":ORDERSERIAL:",orderserial)

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
end function

function sendmailfinish(orderserial,deliverno)
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal
        dim subtotalprice
        mailfrom = "customer@10x10.co.kr"
        mailtitle = "주문하신 상품에 대한 텐바이텐 배송안내입니다!"
        '주문자 메일주소 확인,주문거래종류 선택
        sql = "select buyemail,discountrate,subtotalprice from tbl_order_master where orderserial = '" + orderserial + "'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                rsfunc.Movefirst
                mailto = rsfunc("buyemail")
                discountrate = rsfunc("discountrate")
                subtotalprice = rsfunc("subtotalprice")
        else
                exit function
        end if
        rsfunc.close

        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_finish.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall

        '주문정보 확인.
        sql = "select buyname,regdate, reqzipcode, (b.addr010_si + ' ' + b.addr010_gu + ' ' + a.reqaddress) as reqalladdress, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice from tbl_order_master a, addr010tl b, tbl_order_detail c"
        sql = sql + " where b.addr010_zip1 = left(a.reqzipcode,3) and b.addr010_zip2 = right(a.reqzipcode,3) and a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                discountrate = rsfunc("discountrate")
                rsfunc.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", FormatCurrency(cstr(rsfunc("subtotalprice")))) ' 주문총액
                mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(rsfunc("subtotalprice") - rsfunc("itemcost"))) ) ' 주문한 총item  가격
                mailcontent = replace(mailcontent,":DELIVERYFEE:",  FormatCurrency(cstr(rsfunc("itemcost"))) ) ' 배송금액
                mailcontent = replace(mailcontent,":DELIVERNO:",  deliverno ) ' 운송장번호
                mailcontent = replace(mailcontent,":BUYNAME:", rsfunc("buyname")) ' 주문자 이름
                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' 주문번호
                mailcontent = replace(mailcontent,":REQZIPCODE:", rsfunc("reqzipcode")) ' 배송우편번호
                mailcontent = replace(mailcontent,":REQALLADDRESS:", rsfunc("reqalladdress")) ' 배송주소
        else
                exit function
        end if
        rsfunc.close

        'item 루프 앞뒤부분 짜르기
        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item 루프를 돌릴부분 자르기
        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)

        '주문아이템 정보 확인.
        dim itemserial,inx
        sql = " select a.itemid, b.imglist, c.itemname, (c.cate_large + c.cate_mid + c.cate_small) as itemserial, c.sellcash, a.itemno from tbl_order_detail a, tbl_item_image b, tbl_item c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and a.itemid <> '0' and a.itemid = b.itemid and c.itemid = a.itemid"
        sql = sql + " and (a.cancelyn='N' or a.cancelyn='A')"
        inx = 1
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                rsfunc.Movefirst
                do until rsfunc.eof
                        itemserial = rsfunc("itemserial") + "-" + FormatCode(rsfunc("itemid"))
                        itemHtml = replace(itemHtmlOri,":ITEMSERIAL:", itemserial) ' 상품코드
                        itemHtml = replace(itemHtml,":ITEMNAME:", rsfunc("itemname")) ' 상품이름
                        itemHtml = replace(itemHtml,":ITEMPRICE:",  cstr(rsfunc("sellcash")*cdbl(discountrate)) ) ' 상품가격
                        itemHtml = replace(itemHtml,":ITEMNO:", cstr(rsfunc("itemno"))) ' 수량
                        itemHtml = replace(itemHtml,":IMGLIST:", cstr(rsfunc("imglist"))) ' 상품이미지
                        itemHtmlTotal = itemHtmlTotal & itemHtml

                inx = inx + 1
                rsfunc.movenext
                loop
        else
                exit function
        end if
        rsfunc.close

        mailcontent = beforeItemHtml & itemHtmlTotal & afterItemHtml

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailfinish = mailcontent
end function




function SendMailBaeSongFinish(orderserial,designerid)

		  dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal
        dim subtotalprice, tensongjangno, ipkumdiv, IpkumDivName

		  mailfrom = "customer@10x10.co.kr"
        mailtitle = "상품이 출고되었습니다!"

        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_upche_finish.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall

		sql = "select ipkumdiv,buyname,buyemail,subtotalprice,deliverno from [db_order].[dbo].tbl_order_master"
		sql = sql + " where orderserial = '" + orderserial + "'"
		rsget.Open sql,dbget,1
		if  not rsget.EOF  then
			mailto = rsget("buyemail")
			subtotalprice = rsget("subtotalprice")
			mailcontent = replace(mailcontent,":BUYNAME:", db2html(rsget("buyname"))) ' 주문자 이름

			mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' 주문번호

			tensongjangno = rsget("deliverno")
			ipkumdiv = rsget("ipkumdiv")
		else
			exit function
		end if
		rsget.close

'텐텐배송상태 - 사용안함.

		if ipkumdiv="0" then
			IpkumDivName="주문대기"
		elseif ipkumdiv="1" then
			IpkumDivName="주문실패"
		elseif ipkumdiv="2" then
			IpkumDivName="주문접수"
		elseif ipkumdiv="3" then
			IpkumDivName="주문접수"
		elseif ipkumdiv="4" then
			IpkumDivName="결제완료"
		elseif ipkumdiv="5" then
			IpkumDivName="배송대기"
		elseif ipkumdiv="6" then
			IpkumDivName="배송대기"
		elseif ipkumdiv="7" then
			IpkumDivName="상품출고"
		elseif ipkumdiv="8" then
			IpkumDivName="상품출고"
		end if

        dim itemserial,inx,sinx,einx
		  dim BaesongState
		  dim transco,transurl,songjangstr


				sql = " SELECT a.makerid,a.itemid, a.itemoptionname, c.smallimage, c.itemname, " &_
							" (c.cate_large + c.cate_mid + c.cate_small) as itemserial, " &_
							" a.itemcost as sellcash, a.itemno, c.deliverytype, a.songjangdiv, replace(a.songjangno,'-','') as songjangno, a.currstate " &_
							" ,s.divname,s.findurl " &_
							" FROM [db_order].[dbo].tbl_order_detail a " &_
							" JOIN [db_item].[dbo].tbl_item c " &_
							" 	on a.itemid=c.itemid " &_
							" LEFT JOIN db_order.[dbo].tbl_songjang_div s " &_
							" 	on a.songjangdiv=s.divcd " &_
							" WHERE a.orderserial = '" & Cstr(orderserial) & "' " &_
							" and a.itemid <> '0' " &_
							" and (a.cancelyn='N' or a.cancelyn='A') " &_
							" ORDER BY ( " &_
							" 	case a.makerid  " &_
							" 		when '" & designerid & "' then replace(a.makerid,a.makerid,1) " &_
							" 		else 2 " &_
							" 	end) asc, currstate desc "

        'sql = "select a.makerid,a.itemid, a.itemoptionname, b.imgsmall, c.itemname," + vbcrlf
        'sql = sql + " (c.cate_large + c.cate_mid + c.cate_small) as itemserial," + vbcrlf
        'sql = sql + " a.itemcost as sellcash, a.itemno, c.deliverytype, a.songjangdiv, a.songjangno, a.currstate" + vbcrlf
        'sql = sql + " from [db_order].[dbo].tbl_order_detail a," + vbcrlf
        'sql = sql + " [db_item].[dbo].tbl_item_image b, [db_item].[dbo].tbl_item c" + vbcrlf
        'sql = sql + " where a.orderserial = '" + Cstr(orderserial) + "'" + vbcrlf
        'sql = sql + " and a.itemid <> '0'" + vbcrlf
        'sql = sql + " and a.itemid = b.itemid" + vbcrlf
        'sql = sql + " and c.itemid = a.itemid" + vbcrlf
        'sql = sql + " and (a.cancelyn='N' or a.cancelyn='A')" + vbcrlf
        'sql = sql + " order by (case a.makerid when '" + designerid + "' then" + vbcrlf
        'sql = sql + " replace(a.makerid,a.makerid,1)" + vbcrlf
        'sql = sql + " else" + vbcrlf
        'sql = sql + " 2" + vbcrlf
        'sql = sql + " end) asc, currstate desc"
'response.write sql
'dbget.close()	:	response.End
        inx = 0
		  sinx = 1
		  einx = 0

itemHtml = "<table border='0' cellpadding='0' cellspacing='0'>"

        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                rsget.Movefirst
                do until rsget.eof

						  if inx = 0 then
								if rsget("makerid") = designerid and rsget("currstate") = 7 then
									sinx = 0 '소속업체 처음 실행
									einx = 1
								end if
						  elseif inx <> 0 and rsget("makerid") = designerid and rsget("currstate") <> 7 then
									einx = 0
									sinx = 0 '소속업체지만 미발송 상품 첫 실행
						  elseif einx = 1 and rsget("makerid") <> designerid then
									einx = 0
									sinx = 0 '소속업체이외 상품 첫 실행
						  end if
'

if sinx = 0 then
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td>"
itemHtml = itemHtml + "<table width='550' border='0' cellspacing='1' cellpadding='0'>"
itemHtml = itemHtml + "<tr>"
if rsget("makerid") = designerid then
itemHtml = itemHtml + "<td align='left' valign='top'><img src='http://www.10x10.co.kr/lib/email/images/order01.gif' width='121' height='30'></td>"
else
itemHtml = itemHtml + "<td align='left' valign='top'><img src='http://www.10x10.co.kr/lib/email/images/order02.gif' width='200' height='30'></td>"
end if
itemHtml = itemHtml + "<td>&nbsp;</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td>"
itemHtml = itemHtml + "<table style='border-top: 1px solid #aaaaaa' border='0' cellpadding='0' cellspacing='0' height='4' bgcolor='ECECEC'width='550'>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td><img src='http://www.10x10.co.kr/lib/email/images/spacer.gif' width='550' height='4' align='center'></td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "<table style='border-bottom: 1px solid #555555;'width='550' border='0' height='23' cellpadding='0' cellspacing='0'>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td width='50' class='p11' align='center'>상품</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' class='p11' align='center'>상품명</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>옵션</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='30' class='p11' align='center'>수량</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>배송현황</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='100' class='p11' align='center'>택배/송장</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "</td>"
itemHtml = itemHtml + "</tr>"
end if


'배송상태 지정
if rsget("deliverytype") = 1 or rsget("deliverytype") = 4 then
    if rsget("currstate") = 7 then
		 BaesongState = "<font color='red'>출고완료</font>"
	 else
		 BaesongState = "<font color='#004080'>상품준비중</font>"
	 end if

    ''BaesongState = IpkumDivName '텐텐배송상태
else
	 if rsget("currstate") = 7 then
		 BaesongState = "<font color='red'>출고완료</font>"
	 else
		 BaesongState = "<font color='#004080'>상품준비중</font>"
	 end if
end if


'택배/송장 설정

if ((Not isnull(rsget("songjangno"))) and  (rsget("songjangno")<>"") ) then
	songjangstr = db2html(rsget("divname")) & "<br />( <a href='" & db2html(rsget("findurl")) & rsget("songjangno") & "' target='_blank'>" & rsget("songjangno") & "</a> )"
else
	songjangstr="-"
end if

itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td>"
itemHtml = itemHtml + "<table style='border-bottom: 1px solid #c8c8c8' width='550' border='0' height='57' cellpadding='0' cellspacing='0'>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td width='50'><img src='http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallimage") & "' width='50' height='50'></td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6'>" + db2html(rsget("itemname")) + "</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' align='center'>" + rsget("itemoptionname") + "</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='30' align='center'>" + Cstr(rsget("itemno")) + "ea</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' align='center'>" + BaesongState + "</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='100' align='center'>" & songjangstr & "</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "</td>"
itemHtml = itemHtml + "</tr>"



                inx = inx + 1
                sinx = sinx + 1
                rsget.movenext
                loop
        else
                exit function
        end if
        rsget.close

		itemHtml = itemHtml + "</table>"

		itemHtmlTotal = replace(mailcontent,":INNERORDERTABLE:", itemHtml) ' 주문정보테이블 넣기

      mailcontent = itemHtmlTotal

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)

        SendMailBaeSongFinish = mailcontent
'response.write mailcontent
end function

function SendMailFleaMarketEnd(idx,itemname,buyer,icon1,itemcontents,usermail)

        dim mailfrom, mailtitle, mailcontent
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal
        dim subtotalprice
        mailfrom = "guide@way2way.com"
        mailtitle = "여행자 장터 안내 메일 입니다.!"

        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_fleamarket_end.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall

		mailcontent = replace(mailcontent,"$IDX$", idx )
		mailcontent = replace(mailcontent,"$ITEMNAME$", itemname)
		mailcontent = replace(mailcontent,"$BUYER$", buyer)
		mailcontent = replace(mailcontent,"$ICON$", icon1)
		mailcontent = replace(mailcontent,"$ITEMCONTENTS$", itemcontents)

        call sendmail(mailfrom, usermail, mailtitle, mailcontent)
        SendMailFleaMarketEnd = mailcontent
end function



function SendMailUpCheBaeSongFinish(orderserial)
        dim sql,discountrate,paymethod
        dim mailto, mailtitle, mailcontent,mailfrom
        dim subtotalprice,itemcost,buyname,reqname,reqzipcode,reqalladdress
		dim reqphone,comment

		mailfrom = "customer@10x10.co.kr"
        mailtitle = "상품이 출고되었습니다!"

        '주문자 메일주소 확인,주문거래종류 선택---------------------------------------------------------------------------
        sql = "select buyemail,accountdiv from [db_order].[dbo].tbl_order_master where orderserial = '" + orderserial + "'"
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                rsget.Movefirst
                mailto = rsget("buyemail")
                paymethod = trim(rsget("accountdiv"))
        else
                exit function
        end if
        rsget.close


		dim SpendMile, tencardspend
        '주문정보 확인.---------------------------------------------------------------------------
        sql = "select buyname,regdate, reqname, reqzipcode, (a.reqzipaddr + ' ' + a.reqaddress) as reqalladdress, a.reqphone, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice, a.miletotalprice ,a.tencardspend, a.comment from [db_order].[dbo].tbl_order_master a, [db_order].[dbo].tbl_order_detail c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"

		rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                discountrate = rsget("discountrate")
                tencardspend = rsget("tencardspend")
                rsget.Movefirst
                subtotalprice = formatNumber(FormatCurrency(rsget("subtotalprice")),0) ' 주문총액
                itemcost = formatNumber(FormatCurrency(rsget("itemcost")),0) ' 배송금액
                buyname = rsget("buyname") ' 주문자 이름
                reqname = rsget("reqname") ' 수령인 이름
                reqalladdress = rsget("reqalladdress") ' 배송주소
                reqphone = rsget("reqphone") ' 주문자 전화번호
                comment = rsget("comment") ' 배송메모
                if IsNull(rsget("miletotalprice")) then
                	SpendMile =""
                else
                	SpendMile = rsget("miletotalprice") + tencardspend
                	SpendMile = formatNumber(FormatCurrency(SpendMile),0)
            	end if

		else
                exit function
        end if
        rsget.close


mailcontent ="<html>"
mailcontent = mailcontent + "<head>"
mailcontent = mailcontent + "<title>[텐바이텐] 즐거움이 가득한 쇼핑몰 10x10 = tenbyten</title>"
mailcontent = mailcontent + "<link rel=stylesheet type='text/css' href='http://www.10x10.co.kr/css/tenten.css'>"
mailcontent = mailcontent + "</head>"
mailcontent = mailcontent + "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0' rightmargin='0' bottommargin='0' bgcolor=#ffffff>"
mailcontent = mailcontent + "<table style='padding:3 6 3 6;border: 7px solid #eeeeee' width='355' border='0' cellpadding='0' cellspacing='0' align='center'>"
mailcontent = mailcontent + "<tr>"
mailcontent = mailcontent + "<td>"
mailcontent = mailcontent + "<table width='600' border='0' cellpadding='0' cellspacing='0'>"
mailcontent = mailcontent + "<tr valign='top'>"
mailcontent = mailcontent + "<td width='39' height='57'><img src='http://www.10x10.co.kr/lib/email/images/main_10x10_logo.gif' width='222' height='56'></td>"
mailcontent = mailcontent + "<td width='561' height='57'>"
mailcontent = mailcontent + "<div align='right'><img src='http://www.10x10.co.kr/lib/email/images/mail_order_ok.gif' width='127' height='45'></div>"
mailcontent = mailcontent + "</td>"
mailcontent = mailcontent + "</tr>"
mailcontent = mailcontent + "</table>"
mailcontent = mailcontent + "</td>"
mailcontent = mailcontent + "</tr>"
mailcontent = mailcontent + "<tr>"
mailcontent = mailcontent + "<td><img src='http://www.10x10.co.kr/lib/email/images/mail_finish_title.gif' width='600' height='160'></td>"
mailcontent = mailcontent + "</tr>"
mailcontent = mailcontent + "<tr>"
mailcontent = mailcontent + "<td height='30' bgcolor='f7f7f7'>"
mailcontent = mailcontent + "<div align='center'>"
mailcontent = mailcontent + "<table width='580' border='0' cellpadding='0' cellspacing='5'>"
mailcontent = mailcontent + "<tr>"
mailcontent = mailcontent + "<td><b>[" + buyname + "]님의 주문내역입니다 </b></td>"
mailcontent = mailcontent + "<td>"
mailcontent = mailcontent + "<div align='right'><b>주문번호 : <font color='#CC3300'><span class='verdana-mid'>" + orderserial + "</span></font></b></div>"
mailcontent = mailcontent + "</td>"
mailcontent = mailcontent + "</tr>"
mailcontent = mailcontent + "</table>"
mailcontent = mailcontent + "</div>"
mailcontent = mailcontent + "</td>"
mailcontent = mailcontent + "</tr>"


		'주문아이템 정보 확인.-----------------------------------------------------------------------------

		dim itemserial,inx,inx2,tdata,tdata2
      dim Titemcost,BufCost
		dim upchebaesong
		dim currstate
		dim transco,transurl
        Titemcost = 0

		'업체배송 출고된 상품 가져오기
        sql = "select a.itemid, a.itemoptionname, a.currstate, a.itemname, a.songjangno, a.songjangdiv," + vbcrlf
        sql = sql + " a.itemcost as sellcash, a.itemno, b.imgsmall" + vbcrlf
        sql = sql + " from [db_order].[dbo].tbl_order_detail a," + vbcrlf
        sql = sql + " [db_item].[dbo].tbl_item_image b" + vbcrlf
        sql = sql + " where a.orderserial = '" + orderserial + "'" + vbcrlf
        sql = sql + " and a.itemid <> '0'" + vbcrlf
        sql = sql + " and a.itemid = b.itemid" + vbcrlf
        sql = sql + " and a.currstate >= 7" + vbcrlf
        sql = sql + " and a.isupchebeasong = 'Y'" + vbcrlf
        sql = sql + " and (a.cancelyn='N' or a.cancelyn='A')" + vbcrlf

		inx = 1

        rsget.Open sql,dbget,1

		tdata = rsget.RecordCount

        if  not rsget.EOF  then
                rsget.Movefirst
                do until rsget.eof


					if CDbl(discountrate)=1 then
						BufCost = rsget("sellcash") * rsget("itemno")
						Titemcost = Titemcost + BufCost
					else
						BufCost = round(rsget("sellcash")*cdbl(discountrate)/100)*100 * rsget("itemno")
						Titemcost = Titemcost + BufCost
					end if

					if rsget("currstate") = 3 then
					currstate = "<font color='#46A3FF'>상품준비중</font>"
					elseif rsget("currstate") = 7 then
					currstate = "<font color='#FF6060'>출고완료</font>"
					else
					currstate = "<font color='#939300'>상품준비중</font>"
					end if

					if rsget("songjangdiv") = "1" then
					transco = "한진택배"
					transurl = "http://www.hanjin.co.kr/transmission/main.htm"
					elseif rsget("songjangdiv") = "2" then
					transco = "현대택배"
					transurl = "http://www.hyundaiexpress.com/hydex/jsp/support/search/re_03.jsp"
					elseif rsget("songjangdiv") = "3" then
					transco = "대한통운"
					transurl = "http://doortodoor.korex.co.kr/jsp/cmn/index.jsp"
					elseif rsget("songjangdiv") = "4" then
					transco = "CJ GLS"
					transurl = "http://www.cjgls.com/contents/gls/gls004/gls004_06.asp"
					elseif rsget("songjangdiv") = "5" then
					transco = "이클라인"
					transurl = "http://www.ecline.net/tracking/customer02.html#t01"
					elseif rsget("songjangdiv") = "6" then
					transco = "HTH"
					transurl = "https://samsunghth.com/homepage/searchTraceGoods/SearchTraceInput.jhtml?mc=5"
					elseif rsget("songjangdiv") = "7" then
					transco = "훼미리택배"
					transurl = "http://www.e-family.co.kr/"
					elseif rsget("songjangdiv") = "8" then
					transco = "우체국"
					transurl = "http://service.epost.go.kr/kps_index.html"
					elseif rsget("songjangdiv") = "9" then
					transco = "KGB"
					transurl = "http://www.kgbl.co.kr/"
					elseif rsget("songjangdiv") = "10" then
					transco = "아주택배"
					transurl = "http://www.ajulogis.co.kr/"
					elseif rsget("songjangdiv") = "11" then
					transco = "오렌지택배"
					transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
					elseif rsget("songjangdiv") = "12" then
					transco = "한국택배"
					transurl = "http://www.kls.co.kr/"
					elseif rsget("songjangdiv") = "13" then
					transco = "옐로우캡"
					transurl = "http://www.yellowcap.co.kr/"
					elseif rsget("songjangdiv") = "14" then
					transco = "나이스택배"
					transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
					elseif rsget("songjangdiv") = "15" then
					transco = "중앙택배"
					transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
					elseif rsget("songjangdiv") = "16" then
					transco = "주코택배"
					transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
					elseif rsget("songjangdiv") = "17" then
					transco = "트라넷택배"
					transurl = "http://www.transclub.com/"
					else
					transco = "기타"
					transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
					end if

					if inx = 1 then
						mailcontent = mailcontent + "<tr>"
						mailcontent = mailcontent + "<td align='center'>"
						mailcontent = mailcontent + "<table width='550' border='0' align='center' cellpadding='0' cellspacing='1'>"
						mailcontent = mailcontent + "<tr>"
						mailcontent = mailcontent + "<td  align='center'>"
						mailcontent = mailcontent + "<table width='550' border='0' cellspacing='1' cellpadding='0'>"
						mailcontent = mailcontent + "<tr>"
						mailcontent = mailcontent + "<td height='5'></td>"
						mailcontent = mailcontent + "</tr>"
						mailcontent = mailcontent + "<tr>"
						mailcontent = mailcontent + "<td valign='top' align='center'>"
						mailcontent = mailcontent + "<table width='550' border='0' cellspacing='1' cellpadding='0'>"
						mailcontent = mailcontent + "<tr>"
						mailcontent = mailcontent + "<td align='left' valign='top'><img src='http://www.10x10.co.kr/lib/email/images/order01.gif' width='121' height='30'></td>"
						mailcontent = mailcontent + "<td>&nbsp;</td>"
						mailcontent = mailcontent + "</tr>"
						mailcontent = mailcontent + "</table>"
						mailcontent = mailcontent + "</td>"
						mailcontent = mailcontent + "</tr>"
						mailcontent = mailcontent + "</table>"
						mailcontent = mailcontent + "</td>"
						mailcontent = mailcontent + "</tr>"
						mailcontent = mailcontent + "<tr>"
						mailcontent = mailcontent + "<td valign='top'  align='center'>"
						mailcontent = mailcontent + "<table style='border-top: 1px solid #aaaaaa' border='0' cellpadding='0' cellspacing='0' height='4' bgcolor='ECECEC'width='550'>"
						mailcontent = mailcontent + "<tr>"
						mailcontent = mailcontent + "<td><img src='http://www.10x10.co.kr/lib/email/images/spacer.gif' width='550' height='4' align='center'></td>"
						mailcontent = mailcontent + "</tr>"
						mailcontent = mailcontent + "</table>"
						mailcontent = mailcontent + "<table width='550' border='0' cellspacing='1' cellpadding='0'>"
						mailcontent = mailcontent + "<tr>"
						mailcontent = mailcontent + "<td valign='top'>"
						mailcontent = mailcontent + "<table  width='270' border='0' cellpadding='0' cellspacing='0'>"
						mailcontent = mailcontent + "<tr>"
						mailcontent = mailcontent + "<td style='' valign='top'>"
						mailcontent = mailcontent + "<table style='border-bottom: 1px solid #555555;'width='550' border='0' height='23' cellpadding='0' cellspacing='0'>"
						mailcontent = mailcontent + "<tr>"
						mailcontent = mailcontent + "<td width='50' class='p11' align='center'>상품</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' class='p11' align='center'>상품명</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='50' class='p11' align='center'>옵션</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='50' class='p11' align='center'>수량</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>가격</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>배송현황</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>택배/송장</td>"
						mailcontent = mailcontent + "</tr>"
						mailcontent = mailcontent + "</table>"
						mailcontent = mailcontent + "</td>"
						mailcontent = mailcontent + "</tr>"
					end if

						mailcontent = mailcontent + "<tr>"
						mailcontent = mailcontent + "<td align='center'>"
						mailcontent = mailcontent + "<table style='border-bottom: 1px solid #c8c8c8' width='550' border='0' height='57' cellpadding='0' cellspacing='0'>"
						mailcontent = mailcontent + "<tr>"
						mailcontent = mailcontent + "<td width='50'><img src='http://webimage.10x10.co.kr/image/small/" +  cstr( "0" + CStr(Clng(rsget("itemid")\10000)) + "/" + rsget("imgsmall")) + "' width='50' height='50'></td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6'>" + rsget("itemname") + "</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='50' align='center'>" + rsget("itemoptionname") + "</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='50' align='center'>" + Cstr(rsget("itemno")) + "ea</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='80' align='center'>" + Cstr(BufCost) + "won</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='80' align='center'>" +  currstate  + "</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='80' align='center'>" +  transco + "<br>(<a href='" + transurl + "' target='_blank'>" + rsget("songjangno") + "</a>)</td>"
						mailcontent = mailcontent + "</tr>"
						mailcontent = mailcontent + "</table>"
						mailcontent = mailcontent + "</td>"
						mailcontent = mailcontent + "</tr>"


					if tdata = inx then
						mailcontent = mailcontent + "</table>"
						mailcontent = mailcontent + "</td>"
						mailcontent = mailcontent + "</tr>"
						mailcontent = mailcontent + "</table>"
						mailcontent = mailcontent + "</td>"
						mailcontent = mailcontent + "</tr>"
					end if
				inx = inx + 1

				rsget.movenext
                loop
        end if
        rsget.close


mailcontent = mailcontent + "<tr>"
mailcontent = mailcontent + "<td><img src='http://www.10x10.co.kr/lib/email/images/main_footer.gif' width='600' height='80'></td>"
mailcontent = mailcontent + "</tr>"
mailcontent = mailcontent + "</table>"
mailcontent = mailcontent + "</body>"
mailcontent = mailcontent + "</html>"





        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        SendMailUpCheBaeSongFinish = mailcontent
'response.write mailcontent
end function



'' E-gift카드 전송
function sendGiftCardEmail_SMTP(iorderserial)
    Dim sqlStr
    Dim emailTitle, mailcontents
    Dim sendemail, sender_alias, reqemail, receiver_alias, SendDiv
    sendGiftCardEmail_SMTP = FALSE

    On Error Resume Next
    sqlStr = " select emailTitle"
	sqlStr = sqlStr & " , sendemail"
	sqlStr = sqlStr & " , buyname as sender_alias"
	sqlStr = sqlStr & " , reqemail"
	sqlStr = sqlStr & " , reqemail as receiver_alias"
	sqlStr = sqlStr & " , SendDiv"
	sqlStr = sqlStr & " , db_order.dbo.[sp_Ten_Make_GiftCardEmailMSG]('"&iorderserial&"') as mailcontents"
	sqlStr = sqlStr & " from db_order.dbo.tbl_giftcard_order M"
	sqlStr = sqlStr & " where M.GiftOrderSerial='"&iorderserial&"'"

    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
        emailTitle      = rsget("emailTitle")
        mailcontents    = rsget("mailcontents")
        sendemail       = rsget("sendemail")
        sender_alias    = rsget("sender_alias")
        reqemail        = rsget("reqemail")
        receiver_alias  = rsget("receiver_alias")
        SendDiv         = rsget("SendDiv")
    end if
    rsget.Close

    ''' 이곳에서 검증.
    IF (mailcontents="") then Exit function
    IF (SendDiv<>"E") then Exit function

    call SendMail(sender_alias&"<"&sendemail&">", receiver_alias&"<"&reqemail&">", emailTitle, mailcontents)

    On Error Goto 0
    IF Err Then
        sendGiftCardEmail_SMTP = FALSE
    ELSE
        sendGiftCardEmail_SMTP = TRUE
    END IF

end function

function sendmailStockOutAlarm(orderserial) ' 품절안내 메일
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal, accountdiv, ipkumdate, totItemPrice
		dim sqlStr

        if isnull(orderserial) or orderserial="" then exit function

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "[텐바이텐] 품절안내 메일입니다."

        '주문자 메일주소 확인,주문거래종류 선택
        sql = "select buyemail,accountdiv, ipkumdate from db_order.dbo.tbl_order_master where orderserial = '" + orderserial + "'"
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
            rsget.Movefirst
            mailto = rsget("buyemail")
			ipkumdate = rsget("ipkumdate")
			accountdiv = Trim(rsget("accountdiv"))
        else
			exit function
        end if
        rsget.close

        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_stockoutalarm.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall


		mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' 주문번호
		if isnull(ipkumdate) or ipkumdate="" then
			mailcontent = replace(mailcontent,":IPKUMDATE:", "")
		else
			mailcontent = replace(mailcontent,":IPKUMDATE:", ipkumdate)
		end if

        if accountdiv = "7" then    ' 무통장
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "무통장입금")
        elseif accountdiv = "100" then   ' 신용카드
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "신용카드")
        elseif accountdiv = "20" then   ' 실시간이체
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "실시간이체")
        elseif accountdiv = "80" then   ' 올앳
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "올앳카드")
        elseif accountdiv = "110" then   ' OKCashbag+신용카드
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "OKCashbag+신용카드")
        elseif accountdiv = "400" then   ' 핸드폰결제
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "핸드폰")
        else
        	mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "")
        end if

        '주문아이템 정보 확인.
        dim itemserial,inx
		sqlStr = " select T.mibeasongidx, m.buyname, m.buyhp, m.buyemail, T.mibeasongidx, d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.itemcostCouponNotApplied, d.reducedPrice, T.itemlackno as itemno, i.smallimage as imgsmall "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + "		( "
		sqlStr = sqlStr + "			select l.orderserial, l.detailidx, l.idx as mibeasongidx, l.itemlackno "
		sqlStr = sqlStr + "			from "
		sqlStr = sqlStr + "			db_temp.dbo.tbl_mibeasong_list l "
		sqlStr = sqlStr + "			where "
		sqlStr = sqlStr + "				1 = 1 "
		sqlStr = sqlStr + "				and l.code = '05' "
		sqlStr = sqlStr + "				and l.state < '4' "
		sqlStr = sqlStr + "				and l.isSendSMS = 'N' "
		sqlStr = sqlStr + "				and l.isSendEmail = 'N' "
		sqlStr = sqlStr + "				and l.orderserial = '" & orderserial & "' "
		sqlStr = sqlStr + "		) T "
		sqlStr = sqlStr + "		join [db_order].[dbo].[tbl_order_master] m "
		sqlStr = sqlStr + "		on "
		sqlStr = sqlStr + "			T.orderserial = m.orderserial "
		sqlStr = sqlStr + "		join [db_order].[dbo].[tbl_order_detail] d "
		sqlStr = sqlStr + "		on "
		sqlStr = sqlStr + "			1 = 1 "
		sqlStr = sqlStr + "			and T.detailidx = d.idx "
		sqlStr = sqlStr + "			and T.orderserial = d.orderserial "
		sqlStr = sqlStr + "		join [db_item].[dbo].[tbl_item] i "
		sqlStr = sqlStr + "		on "
		sqlStr = sqlStr + "			1 = 1 "
		sqlStr = sqlStr + "			and i.itemid = d.itemid "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + "		1 = 1 "
		sqlStr = sqlStr + "		and d.cancelyn <> 'Y' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + "		d.itemid, d.itemoption "
        rsget.Open sqlStr,dbget,1



		itemHtmlTotal = ""

		itemHtmlOri = "											<tr>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:50px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea;""><img src="":ITEMIMAGE:"" alt="""" /></td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:100px;margin:0;  padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-size:11px; line-height:11px; font-family:dotum, '돋움', sans-serif; color:#707070;"">:ITEMID:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:295px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:left; font-size:11px; line-height:17px; font-family:dotum, '돋움', sans-serif; color:#707070;"">:ITEMNAME::ITEMOPTIONNAME:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:12px; text-align:right; font-family:dotum, '돋움', sans-serif;"">" & vbCrLf
		itemHtmlOri = itemHtmlOri & "													<span style=""margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, '돋움', sans-serif; text-align:right;"">:ITEMCOST:</span><br />" & vbCrLf
		itemHtmlOri = itemHtmlOri & "													:REDUCEDPRICE:" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:25px; padding:6px 0; border-bottom:solid 1px #eaeaea;""></td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px; line-height:12px;"">:ITEMNO:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "											</tr>" & vbCrLf

        if  not rsget.EOF  then
            rsget.Movefirst
            do until rsget.eof
				itemHtml = replace(itemHtmlOri,":ITEMIMAGE:", webImgUrl & "/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("imgsmall"))
				itemHtml = replace(itemHtml,":ITEMID:", rsget("itemid")) ' 상품코드
				itemHtml = replace(itemHtml,":ITEMNAME:", db2html(rsget("itemname")))
				if (rsget("itemoption") = "0000") then
					itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "")
				else
					itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "<br />" & db2html(rsget("itemoptionname")))
				end if

				itemHtml = replace(itemHtml,":ITEMCOST:", FormatNumber(rsget("itemcostCouponNotApplied"), 0) & "원")
				if (rsget("reducedPrice") < rsget("itemcostCouponNotApplied")) then
					itemHtml = replace(itemHtml,":REDUCEDPRICE:", "<span style=""margin:0; padding:6px 0; color:#dd5555; font-size:11px; line-height:17px; text-align:center""><img src=""http://mailzine.10x10.co.kr/2017/ico_coupon.png"" alt=""쿠폰적용"" style=""vertical-align:-2px; padding-right:2px;""/>" & FormatNumber(rsget("reducedPrice"), 0) & "원" & "</span>")
				else
					itemHtml = replace(itemHtml,":REDUCEDPRICE:", "")
				end if

				itemHtml = replace(itemHtml,":ITEMNO:", rsget("itemno"))

				itemHtmlTotal = itemHtmlTotal & itemHtml
				totItemPrice = totItemPrice + (rsget("reducedPrice") * rsget("itemno"))

                rsget.movenext
            loop
		else
			rsget.close
            exit function
		end if
        rsget.close

        mailcontent = replace(mailcontent,":ITEMLIST:", itemHtmlTotal)
        mailcontent = replace(mailcontent,":TOTITEMPRICE:", FormatNumber(totItemPrice, 0) & "원")

	'//=======  메일 발송 =========/
	dim oMail
	set oMail = New MailCls

	IF replace(mailto,"'","")<>"" THEN

		oMail.MailTitles	= replace(mailtitle,"'","")
		oMail.SenderNm		= "텐바이텐"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= replace(mailto,"'","")
		oMail.ReceiverMail	= replace(mailto,"'","")
		oMail.MailConts 	= newhtml2db(mailcontent)
		oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
		oMail.Send_TMSMailer()		'TMS메일러
		'oMail.Send_Mailer()
		''oMail.Send_CDO()

	End IF

	SET oMail = nothing

        sendmailStockOutAlarm = mailcontent
end function

' 택배파업 안내 메일
function sendmailDeliverystrikeAlarm(orderserial)
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal, accountdiv, ipkumdate, totItemPrice
		dim sqlStr

        if isnull(orderserial) or orderserial="" then exit function

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "[텐바이텐] 택배파업 안내 메일입니다."

        '주문자 메일주소 확인,주문거래종류 선택
        sql = "select buyemail,accountdiv, ipkumdate from db_order.dbo.tbl_order_master where orderserial = '" + orderserial + "'"
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
            rsget.Movefirst
            mailto = rsget("buyemail")
			ipkumdate = rsget("ipkumdate")
			accountdiv = Trim(rsget("accountdiv"))
        else
			exit function
        end if
        rsget.close

        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_Deliverystrikealarm.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall


		mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' 주문번호
		if isnull(ipkumdate) or ipkumdate="" then
			mailcontent = replace(mailcontent,":IPKUMDATE:", "")
		else
			mailcontent = replace(mailcontent,":IPKUMDATE:", ipkumdate)
		end if

        if accountdiv = "7" then    ' 무통장
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "무통장입금")
        elseif accountdiv = "100" then   ' 신용카드
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "신용카드")
        elseif accountdiv = "20" then   ' 실시간이체
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "실시간이체")
        elseif accountdiv = "80" then   ' 올앳
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "올앳카드")
        elseif accountdiv = "110" then   ' OKCashbag+신용카드
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "OKCashbag+신용카드")
        elseif accountdiv = "400" then   ' 핸드폰결제
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "핸드폰")
        else
        	mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "")
        end if

        '주문아이템 정보 확인.
        dim itemserial,inx
		sqlStr = " select T.mibeasongidx, m.buyname, m.buyhp, m.buyemail, T.mibeasongidx, d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.itemcostCouponNotApplied, d.reducedPrice, T.itemlackno as itemno, i.smallimage as imgsmall "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + "		( "
		sqlStr = sqlStr + "			select l.orderserial, l.detailidx, l.idx as mibeasongidx, l.itemlackno "
		sqlStr = sqlStr + "			from "
		sqlStr = sqlStr + "			db_temp.dbo.tbl_mibeasong_list l "
		sqlStr = sqlStr + "			where "
		sqlStr = sqlStr + "				1 = 1 "
		sqlStr = sqlStr + "				and l.code = '06' "
		sqlStr = sqlStr + "				and l.state <= '4' "
		sqlStr = sqlStr + "				and l.isSendSMS = 'N' "
		sqlStr = sqlStr + "				and l.isSendEmail = 'N' "
		sqlStr = sqlStr + "				and l.orderserial = '" & orderserial & "' "
		sqlStr = sqlStr + "		) T "
		sqlStr = sqlStr + "		join [db_order].[dbo].[tbl_order_master] m "
		sqlStr = sqlStr + "		on "
		sqlStr = sqlStr + "			T.orderserial = m.orderserial "
		sqlStr = sqlStr + "		join [db_order].[dbo].[tbl_order_detail] d "
		sqlStr = sqlStr + "		on "
		sqlStr = sqlStr + "			1 = 1 "
		sqlStr = sqlStr + "			and T.detailidx = d.idx "
		sqlStr = sqlStr + "			and T.orderserial = d.orderserial "
		sqlStr = sqlStr + "		join [db_item].[dbo].[tbl_item] i "
		sqlStr = sqlStr + "		on "
		sqlStr = sqlStr + "			1 = 1 "
		sqlStr = sqlStr + "			and i.itemid = d.itemid "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + "		1 = 1 "
		sqlStr = sqlStr + "		and d.cancelyn <> 'Y' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + "		d.itemid, d.itemoption "
        rsget.Open sqlStr,dbget,1



		itemHtmlTotal = ""

		itemHtmlOri = "											<tr>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:50px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea;""><img src="":ITEMIMAGE:"" alt="""" /></td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:100px;margin:0;  padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-size:11px; line-height:11px; font-family:dotum, '돋움', sans-serif; color:#707070;"">:ITEMID:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:295px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:left; font-size:11px; line-height:17px; font-family:dotum, '돋움', sans-serif; color:#707070;"">:ITEMNAME::ITEMOPTIONNAME:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:12px; text-align:right; font-family:dotum, '돋움', sans-serif;"">" & vbCrLf
		itemHtmlOri = itemHtmlOri & "													<span style=""margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, '돋움', sans-serif; text-align:right;"">:ITEMCOST:</span><br />" & vbCrLf
		itemHtmlOri = itemHtmlOri & "													:REDUCEDPRICE:" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:25px; padding:6px 0; border-bottom:solid 1px #eaeaea;""></td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px; line-height:12px;"">:ITEMNO:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "											</tr>" & vbCrLf

        if  not rsget.EOF  then
            rsget.Movefirst
            do until rsget.eof
				itemHtml = replace(itemHtmlOri,":ITEMIMAGE:", webImgUrl & "/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("imgsmall"))
				itemHtml = replace(itemHtml,":ITEMID:", rsget("itemid")) ' 상품코드
				itemHtml = replace(itemHtml,":ITEMNAME:", db2html(rsget("itemname")))
				if (rsget("itemoption") = "0000") then
					itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "")
				else
					itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "<br />" & db2html(rsget("itemoptionname")))
				end if

				itemHtml = replace(itemHtml,":ITEMCOST:", FormatNumber(rsget("itemcostCouponNotApplied"), 0) & "원")
				if (rsget("reducedPrice") < rsget("itemcostCouponNotApplied")) then
					itemHtml = replace(itemHtml,":REDUCEDPRICE:", "<span style=""margin:0; padding:6px 0; color:#dd5555; font-size:11px; line-height:17px; text-align:center""><img src=""http://mailzine.10x10.co.kr/2017/ico_coupon.png"" alt=""쿠폰적용"" style=""vertical-align:-2px; padding-right:2px;""/>" & FormatNumber(rsget("reducedPrice"), 0) & "원" & "</span>")
				else
					itemHtml = replace(itemHtml,":REDUCEDPRICE:", "")
				end if

				itemHtml = replace(itemHtml,":ITEMNO:", rsget("itemno"))

				itemHtmlTotal = itemHtmlTotal & itemHtml
				totItemPrice = totItemPrice + (rsget("reducedPrice") * rsget("itemno"))

                rsget.movenext
            loop
		else
			rsget.close
            exit function
		end if
        rsget.close

        mailcontent = replace(mailcontent,":ITEMLIST:", itemHtmlTotal)
        mailcontent = replace(mailcontent,":TOTITEMPRICE:", FormatNumber(totItemPrice, 0) & "원")

	'//=======  메일 발송 =========/
	dim oMail
	set oMail = New MailCls

	IF replace(mailto,"'","")<>"" THEN

		oMail.MailTitles	= replace(mailtitle,"'","")
		oMail.SenderNm		= "텐바이텐"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= replace(mailto,"'","")
		oMail.ReceiverMail	= replace(mailto,"'","")
		oMail.MailConts 	= newhtml2db(mailcontent)
		oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
		oMail.Send_TMSMailer()		'TMS메일러
		'oMail.Send_Mailer()
		''oMail.Send_CDO()

	End IF

	SET oMail = nothing

        sendmailDeliverystrikeAlarm = mailcontent
end function

' 주문지연안내메일      ' 2020.10.27 한용민
function sendmaildelayAlarm(orderserial)
        dim sql,discountrate,paymethod, sqlStr, mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal, accountdiv, ipkumdate, totItemPrice

        if isnull(orderserial) or orderserial="" then exit function

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "[텐바이텐] 발송지연 안내 메일입니다."

        '주문자 메일주소 확인,주문거래종류 선택
        sql = "select buyemail,accountdiv, ipkumdate from db_order.dbo.tbl_order_master with (nolock)"
		sql = sql & " where orderserial = '" + orderserial + "'"

		'response.write sql & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
        if  not rsget.EOF  then
            rsget.Movefirst
            mailto = rsget("buyemail")
			ipkumdate = rsget("ipkumdate")
			accountdiv = Trim(rsget("accountdiv"))
        else
			exit function
        end if
        rsget.close

        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_stockoutalarm_customer.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall


		mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' 주문번호
		if isnull(ipkumdate) or ipkumdate="" then
			mailcontent = replace(mailcontent,":IPKUMDATE:", "")
		else
			mailcontent = replace(mailcontent,":IPKUMDATE:", ipkumdate)
		end if

        if accountdiv = "7" then    ' 무통장
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "무통장입금")
        elseif accountdiv = "100" then   ' 신용카드
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "신용카드")
        elseif accountdiv = "20" then   ' 실시간이체
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "실시간이체")
        elseif accountdiv = "80" then   ' 올앳
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "올앳카드")
        elseif accountdiv = "110" then   ' OKCashbag+신용카드
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "OKCashbag+신용카드")
        elseif accountdiv = "400" then   ' 핸드폰결제
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "핸드폰")
        else
        	mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "")
        end if

        '주문아이템 정보 확인.
        dim itemserial,inx
		sqlStr = " select T.mibeasongidx, m.buyname, m.buyhp, m.buyemail, T.mibeasongidx, d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.itemcostCouponNotApplied, d.reducedPrice, d.itemno, i.smallimage as imgsmall "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + "		( "
		sqlStr = sqlStr + "			select l.orderserial, l.detailidx, l.idx as mibeasongidx "
		sqlStr = sqlStr + "			from "
		sqlStr = sqlStr + "			db_temp.dbo.tbl_mibeasong_list l with (nolock) "
		sqlStr = sqlStr + "			where l.code = '03' "
		sqlStr = sqlStr + "				and l.state < '4' "
		sqlStr = sqlStr + "				and l.isSendSMS = 'N' "
		sqlStr = sqlStr + "				and l.isSendEmail = 'N' "
		sqlStr = sqlStr + "				and l.orderserial = '" & orderserial & "' "
		sqlStr = sqlStr + "		) T "
		sqlStr = sqlStr + "		join [db_order].[dbo].[tbl_order_master] m with (nolock) "
		sqlStr = sqlStr + "		on T.orderserial = m.orderserial "
		sqlStr = sqlStr + "		join [db_order].[dbo].[tbl_order_detail] d with (nolock) "
		sqlStr = sqlStr + "		on T.detailidx = d.idx "
		sqlStr = sqlStr + "			and T.orderserial = d.orderserial "
		sqlStr = sqlStr + "		join [db_item].[dbo].[tbl_item] i with (nolock) "
		sqlStr = sqlStr + "		on i.itemid = d.itemid "
		sqlStr = sqlStr + " left join db_temp.dbo.tbl_mibeasong_list bl with (nolock)"
		sqlStr = sqlStr + " 	on t.orderserial = bl.orderserial"
		sqlStr = sqlStr + " 	and bl.code = '05'"
		sqlStr = sqlStr + " where d.cancelyn <> 'Y' "
		sqlStr = sqlStr + " and bl.orderserial is null"		' 품절문자 발송이 있는경우, 출고지연 문자는 보내지 않는다.
		sqlStr = sqlStr + " order by d.itemid, d.itemoption "

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		itemHtmlTotal = ""

		itemHtmlOri = "											<tr>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:50px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea;""><img src="":ITEMIMAGE:"" alt="""" /></td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:100px;margin:0;  padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-size:11px; line-height:11px; font-family:dotum, '돋움', sans-serif; color:#707070;"">:ITEMID:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:295px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:left; font-size:11px; line-height:17px; font-family:dotum, '돋움', sans-serif; color:#707070;"">:ITEMNAME::ITEMOPTIONNAME:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:12px; text-align:right; font-family:dotum, '돋움', sans-serif;"">" & vbCrLf
		itemHtmlOri = itemHtmlOri & "													<span style=""margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, '돋움', sans-serif; text-align:right;"">:ITEMCOST:</span><br />" & vbCrLf
		itemHtmlOri = itemHtmlOri & "													:REDUCEDPRICE:" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:25px; padding:6px 0; border-bottom:solid 1px #eaeaea;""></td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px; line-height:12px;"">:ITEMNO:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "											</tr>" & vbCrLf

        if  not rsget.EOF  then
            rsget.Movefirst
            do until rsget.eof
				itemHtml = replace(itemHtmlOri,":ITEMIMAGE:", webImgUrl & "/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("imgsmall"))
				itemHtml = replace(itemHtml,":ITEMID:", rsget("itemid")) ' 상품코드
				itemHtml = replace(itemHtml,":ITEMNAME:", db2html(rsget("itemname")))
				if (rsget("itemoption") = "0000") then
					itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "")
				else
					itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "<br />" & db2html(rsget("itemoptionname")))
				end if

				itemHtml = replace(itemHtml,":ITEMCOST:", FormatNumber(rsget("itemcostCouponNotApplied"), 0) & "원")
				if (rsget("reducedPrice") < rsget("itemcostCouponNotApplied")) then
					itemHtml = replace(itemHtml,":REDUCEDPRICE:", "<span style=""margin:0; padding:6px 0; color:#dd5555; font-size:11px; line-height:17px; text-align:center""><img src=""http://mailzine.10x10.co.kr/2017/ico_coupon.png"" alt=""쿠폰적용"" style=""vertical-align:-2px; padding-right:2px;""/>" & FormatNumber(rsget("reducedPrice"), 0) & "원" & "</span>")
				else
					itemHtml = replace(itemHtml,":REDUCEDPRICE:", "")
				end if

				itemHtml = replace(itemHtml,":ITEMNO:", rsget("itemno"))

				itemHtmlTotal = itemHtmlTotal & itemHtml
				totItemPrice = totItemPrice + (rsget("reducedPrice") * rsget("itemno"))

                rsget.movenext
            loop
		else
			rsget.close
            exit function
		end if
        rsget.close

        mailcontent = replace(mailcontent,":ITEMLIST:", itemHtmlTotal)
        mailcontent = replace(mailcontent,":TOTITEMPRICE:", FormatNumber(totItemPrice, 0) & "원")

	'//=======  메일 발송 =========/
	dim oMail
	set oMail = New MailCls

	IF replace(mailto,"'","")<>"" THEN

		oMail.MailTitles	= replace(mailtitle,"'","")
		oMail.SenderNm		= "텐바이텐"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= replace(mailto,"'","")
		oMail.ReceiverMail	= replace(mailto,"'","")
		oMail.MailConts 	= newhtml2db(mailcontent)
		oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
		oMail.Send_TMSMailer()		'TMS메일러
		'oMail.Send_Mailer()
		''oMail.Send_CDO()

	End IF

	SET oMail = nothing

        sendmaildelayAlarm = mailcontent
end function
%>
