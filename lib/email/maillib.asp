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

    		'-> ���� ���ٹ���� �����մϴ�
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)

    		'-> ���� �ּҸ� �����մϴ�
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "110.93.128.94"

    		'-> ������ ��Ʈ��ȣ�� �����մϴ�
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

    		'-> ���ӽõ��� ���ѽð��� �����մϴ�
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10

    		'-> SMTP ���� ��������� �����մϴ�
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

    		'-> SMTP ������ ������ ID�� �Է��մϴ�
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"

    		'-> SMTP ������ ������ ��ȣ�� �Է��մϴ�
    		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"

    		cdoConfig.Fields.Update

    		Set cdoMessage = CreateObject("CDO.Message")

    		''Set cdoMessage.Configuration = cdoConfig

    		cdoMessage.To 				= mailto
    		cdoMessage.From 			= "test@testsvr-am1vgl5" ''mailfrom
    		cdoMessage.SubJect 	= mailtitle
    		'���� ������ �ؽ�Ʈ�� ��� cdoMessage.TextBody, html�� ��� cdoMessage.HTMLBody
    		cdoMessage.HTMLBody	= mailcontent

    		''�׽�Ʈ ȯ��
    		if ((InStr(mailto,"10x10.co.kr")>0) or (mailto="archilee@shinbiro.com")) then
    		    cdoMessage.Send
            end if

    		Set cdoMessage = nothing
    		Set cdoConfig = nothing



        End IF

end sub


''�ܺ� ������ ������
'sub SendMail(mailfrom, mailto, mailtitle, mailcontent)
'
'		dim cdoMessage,cdoConfig
'        On Error Resume Next
'		Set cdoConfig = CreateObject("CDO.Configuration")
'
'		'-> ���� ���ٹ���� �����մϴ�
'		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)
'
'		'-> ���� �ּҸ� �����մϴ�
'		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="mailzine.10x10.co.kr"
'
'		'-> ������ ��Ʈ��ȣ�� �����մϴ�
'		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'
'		'-> ���ӽõ��� ���ѽð��� �����մϴ�
'		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 5
'
'		'-> SMTP ���� ��������� �����մϴ�
'		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'
'		'-> SMTP ������ ������ ID�� �Է��մϴ�
'		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"
'
'		'-> SMTP ������ ������ ��ȣ�� �Է��մϴ�
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
'		'���� ������ �ؽ�Ʈ�� ��� cdoMessage.TextBody, html�� ��� cdoMessage.HTMLBody
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

        mailtitle = "[�ٹ�����] �ֹ��� ���� �Ա�Ȯ��(���Ա�) �ȳ������Դϴ�"

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

        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_pay_delay.htm"


        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
'        mailcontent = replace(mailcontent,":USERNAME:",userName)


		dim SpendMile, tencardspend
		dim IsForeighDeliver : IsForeighDeliver = false
        '�ֹ����� Ȯ��.---------------------------------------------------------------------------


        mailto = myorder.FOneItem.Fbuyemail
        paymethod = trim(myorder.FOneItem.Faccountdiv)


        if paymethod = "7" then    ' ������
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�������Ա�")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "�Ա��� ����")
        elseif paymethod = "100" then   ' �ſ�ī��
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ſ�ī��")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "�����Ϸ�")
        elseif paymethod = "20" then   ' �ǽð���ü
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ǽð���ü")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "�����Ϸ�")
        elseif paymethod = "80" then   ' �þ�
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�þ�ī��")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "�����Ϸ�")
        elseif paymethod = "110" then   ' OKCashbag+�ſ�ī��
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "OKCashbag+�ſ�ī��")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "�����Ϸ�")
        elseif paymethod = "400" then   ' �ڵ�������
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ڵ���")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "�����Ϸ�")
        else
        	mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "")
        end if

        if (paymethod<>"7") then
            mailcontent = ReplaceText(mailcontent,"(<!-----bankinfo------>)[\s\S]*(<!-----/bankinfo------>)","")
            mailcontent = ReplaceText(mailcontent,"(<!-----banknotiinfo------>)[\s\S]*(<!-----/banknotiinfo------>)","")
        end if

        IsForeighDeliver = myorder.FOneItem.IsForeignDeliver

        if (IsForeighDeliver) then
            mailcontent = replace(mailcontent,":REQHPORREQEMAIL:", "�̸���") ' ������ �̸���
            mailcontent = replace(mailcontent,":REQHP:", myorder.FOneItem.Freqemail) ' ������ ��ȭ��ȣ=>�̸��Ϸ�
            mailcontent = replace(mailcontent,":COUNTRYNAME:", myorder.FOneItem.FcountryNameEn) ' ����.
            mailcontent = replace(mailcontent,":REQZIPCODE:", myorder.FOneItem.FemsZipCode) ' ��ۿ����ȣ
        else
            mailcontent = replace(mailcontent,":REQHPORREQEMAIL:", "�޴�����ȣ") ' �޴�����ȣ
            mailcontent = replace(mailcontent,":REQHP:", myorder.FOneItem.Freqhp) ' ������ ��ȭ��ȣ
            mailcontent = replace(mailcontent,":REQZIPCODE:", myorder.FOneItem.Freqzipcode) ' ��ۿ����ȣ
            mailcontent = ReplaceText(mailcontent,"(<!-- foreigndelivery -->)[\s\S]*(<!--/foreigndelivery -->)","")
        end if

        mailcontent = replace(mailcontent,":BUYNAME:", myorder.FOneItem.Fbuyname) ' �ֹ��� �̸�
        mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
        mailcontent = replace(mailcontent,":REQNAME:", myorder.FOneItem.Freqname) ' ������ �̸�
        mailcontent = replace(mailcontent,":REQALLADDRESS:", myorder.FOneItem.FreqZipaddr + " " + myorder.FOneItem.Freqaddress) ' ����ּ�
        mailcontent = replace(mailcontent,":REQPHONE:", myorder.FOneItem.Freqphone) ' ������ ��ȭ��ȣ

        mailcontent = replace(mailcontent,":BEASONGMEMO:", myorder.FOneItem.Fcomment) ' ��۸޸�


    	if (paymethod="110") then
    	    mailcontent = replace(mailcontent,":MAJORTOTALPRICE:", formatNumber(myorder.FOneItem.TotalMajorPaymentPrice,0) & " (�ſ�ī��:" &FormatNumber(myorder.FOneItem.TotalMajorPaymentPrice-myorder.FOneItem.FokcashbagSpend,0)& ",  OKCashbag:" &FormatNumber(myorder.FOneItem.FokcashbagSpend,0) &")") ' �����Ѿ�
    	else
    	    mailcontent = replace(mailcontent,":MAJORTOTALPRICE:", formatNumber(myorder.FOneItem.TotalMajorPaymentPrice,0)) ' �����Ѿ�
        end if

        mailcontent = replace(mailcontent,":ACCOUNTNO:", myorder.FOneItem.Faccountno) ' �Աݰ���

        if (myorder.FOneItem.FsumPaymentEtc<>0) then
            mailcontent = replace(mailcontent,":SPENDTENCASH:", FormatNumber(myorder.FOneItem.FsumPaymentEtc,0))
        else
            mailcontent = ReplaceText(mailcontent,"(<!-----spendtencash------>)[\s\S]*(<!-----/spendtencash------>)","")
        end if


		'�ֹ������� ���� Ȯ��.-----------------------------------------------------------------------------


itemHtml = itemHtml + "<table width='100%' border='0' cellspacing='0' cellpadding='0' style='border-top:3px solid #be0808;'>"
itemHtml = itemHtml + "  <tr>"
itemHtml = itemHtml + "    <td height='30' style='background:#fcf6f6; border-bottom:1px solid #eaeaea;'>"
itemHtml = itemHtml + "    <table width='100%' border='0' cellspacing='0' cellpadding='0' style='font-family:Dotum; font-size:11px; color:#888;padding-top:3px; '>"
itemHtml = itemHtml + "        <tr align='center'>"
itemHtml = itemHtml + "          <td width='60'>��ǰ</td>"
itemHtml = itemHtml + "          <td width='60'>��ǰ�ڵ�</td>"
itemHtml = itemHtml + "          <td>��ǰ�� [�ɼ�]</td>"
itemHtml = itemHtml + "          <td width='90'>�ǸŰ���</td>"
itemHtml = itemHtml + "          <td width='40'>����</td>"
itemHtml = itemHtml + "          <td width='90'>�ֹ��ݾ�</td>"
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
    itemHtml = itemHtml + "          <span style='font-family: ����; FONT-SIZE: 11px; COLOR: #c20a0a; font-weight:bold;'><img src='http://fiximage.10x10.co.kr/web2008/shoppingbag/icon_sale.gif' width=14>"&FormatNumber(myorderdetail.FItemList(i).getItemcostCouponNotApplied,0)&"</span>"
else
    if (Not IsNull(myorderdetail.FItemList(i).Fitemcouponidx)) then
    itemHtml = itemHtml + "          <strike>"&FormatNumber(myorderdetail.FItemList(i).FitemcostCouponNotApplied,0)&"</strike>"
    else
    itemHtml = itemHtml + "          "&FormatNumber(myorderdetail.FItemList(i).FitemcostCouponNotApplied,0)
    end if
end if
    itemHtml = itemHtml + "          "&"��"

if (Not IsNull(myorderdetail.FItemList(i).Fitemcouponidx)) then
    itemHtml = itemHtml + "          <br><span style='font-family: ����; FONT-SIZE: 11px; font-weight:bold ; COLOR: #438938;' >"&FormatNumber(myorderdetail.FItemList(i).FItemCost,0)& "��" &"</span>"
end if

itemHtml = itemHtml + "          </td>"
itemHtml = itemHtml + "          <td width='40'>" &myorderdetail.FItemList(i).FItemNo& "</td>"
itemHtml = itemHtml + "          <td width='90'>" &FormatNumber(myorderdetail.FItemList(i).FItemCost*myorderdetail.FItemList(i).FItemNo,0) & "��" & "</td>"
itemHtml = itemHtml + "        </tr>"
itemHtml = itemHtml + "      </table></td>"
itemHtml = itemHtml + "  </tr>"
			end if
        next

itemHtml = itemHtml + "</table>"


		itemHtmlTotal = replace(mailcontent,":INNERORDERTABLE:", itemHtml) ' �ֹ��������̺� �ֱ�

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
ttlsumHTML = ttlsumHTML + "                <td style='padding-left:30px;' style='font-family: ����; FONT-SIZE: 12px; COLOR: #888888;'></td>"
ttlsumHTML = ttlsumHTML + "                <td style='padding-left:30px;' style='font-family: ����; FONT-SIZE: 12px; COLOR: #888888;'>���� ���ϸ��� <strong>"& FormatNumber(myorder.FOneItem.Ftotalmileage,0) &"</strong> Point</td>"
ttlsumHTML = ttlsumHTML + "                <td style='padding-left:30px;' style='font-family: ����; FONT-SIZE: 12px; COLOR: #888888;'>��ǰ�����Ѿ� <strong>"& FormatNumber(myorder.FOneItem.FTotalSum-myorderdetail.BeasongPay,0) &"</strong> ��</td>"
ttlsumHTML = ttlsumHTML + "              </tr>"
ttlsumHTML = ttlsumHTML + "            </table></td>"
ttlsumHTML = ttlsumHTML + "          </tr>"
ttlsumHTML = ttlsumHTML + "          <tr height='1'>"
ttlsumHTML = ttlsumHTML + "            <td height='1' bgcolor='#eaeaea'></td>"
ttlsumHTML = ttlsumHTML + "          </tr>"
ttlsumHTML = ttlsumHTML + "          <tr height='30'>"
ttlsumHTML = ttlsumHTML + "            <td align='right' style='padding:10px 5px 10px 5px;' style='font-family: ����; FONT-SIZE: 12px; COLOR: #000000;'>"
ttlsumHTML = ttlsumHTML + "         �Ѱ����ݾ� : "
ttlsumHTML = ttlsumHTML + "			��ǰ�����Ѿ� "& FormatNumber((myorder.FOneItem.FTotalSum-myorderdetail.BeasongPay),0) &"�� "
ttlsumHTML = ttlsumHTML + "			+ ��ۺ� "& FormatNumber(myorderdetail.BeasongPay,0) &"��"

IF (myorder.FOneItem.Fmiletotalprice<>0) then
    ttlsumHTML = ttlsumHTML + "			- ���ϸ��� "& FormatNumber(myorder.FOneItem.Fmiletotalprice,0) &"�� "
end if
IF (myorder.FOneItem.Ftencardspend<>0) then
    ttlsumHTML = ttlsumHTML + "			- ���ʽ��������� "&FormatNumber(myorder.FOneItem.Ftencardspend,0) &"�� "
end if

if (myorder.FOneItem.Fallatdiscountprice + myorder.FOneItem.Fspendmembership<>0) then
    ttlsumHTML = ttlsumHTML + "			- ��Ÿ���� "& FormatNumber((myorder.FOneItem.Fallatdiscountprice + myorder.FOneItem.Fspendmembership),0) &"�� "
end if
ttlsumHTML = ttlsumHTML + "			= "
ttlsumHTML = ttlsumHTML + "			<span style='font-family: ����; FONT-SIZE: 12px; COLOR: #c20a0a; font-weight:bold;'>"& FormatNumber(myorder.FOneItem.FsubtotalPrice,0) &"</span> ��"
ttlsumHTML = ttlsumHTML + "            </td>"
ttlsumHTML = ttlsumHTML + "          </tr>"
ttlsumHTML = ttlsumHTML + "        </table></td>"
ttlsumHTML = ttlsumHTML + "      </tr>"
ttlsumHTML = ttlsumHTML + "    </table></td>"
ttlsumHTML = ttlsumHTML + "  </tr>"
ttlsumHTML = ttlsumHTML + "</table>"

        mailcontent = replace(mailcontent,":ORDERPRICESUMMARY:", ttlsumHTML) ' �ֹ� �հ�ݾ�



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

function sendmailnewuser2(mailto,userName) ' ���Ը��������� �о���̴� ������� ��ȯ
        dim mailfrom, mailtitle, mailcontent,dirPath,fileName
        dim fs,objFile

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "10x10 ����Ʈ ������ ���� �帳�ϴ�."

        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_join.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":USERNAME:",userName)

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailnewuser2 = mailcontent
end function

sub sendmailnewuser(mailto) ' �� function���� ��ȯ��.20020329/
        dim mailfrom, mailtitle, mailcontent

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "10x10 ����Ʈ ������ ���� �帳�ϴ�."

        '�̺κ��� ���� html ��. ���� ���α׷� ���°��� ���»���...
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
        +"			<p>������ ���� ����Ʈ 10X10.co.kr (�ٹ�����) �� ������ �ּż� �������� ����帳�ϴ�.<br><br>"	_
        +"			   ���� 10X10 �� ���θ��� Ŀ�´�Ƽ�� ���յ� ������ ä�ημ� <br><br>"	_
        +"			   �������� ���� ��հ� ���� �ִ� ����Ʈ �Դϴ�.<br><br>"	_
        +"          �׻� �ູ�� �ϵ��� ȸ�� �����в� �����ϱ� �ٶ��ϴ�... : )</p>"	_
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
        +"      <td align=right valign=top>(��)ť�� Ŀ�´ϸ���<img src=http://www.10x10.co.kr/images/cube_ci.gif width=210 height=52 hspace=15></td>"	_
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
        mailtitle = "�ֹ��� ���������� �����Ǿ����ϴ�!"

        '�ֹ��� �����ּ� Ȯ��.
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
        mailcontent = mailcontent + "      <td bgcolor=F1F1F1>���� �ٹ����� ����Ʈ�� �̿��� �ּż� �������� ����帮�� ���� �ֹ��� ���������� �����Ǿ� ó�����Դϴ�. " + vbcr
        mailcontent = mailcontent + "        <br> �ſ�ī�� ������ �ֹ����� �� ��ٷ� ��ۿ� ���� �¶����Ա� �ֹ��� �Ա�Ȯ�� �� ����� �̷�� ���ϴ�." + vbcr
        mailcontent = mailcontent + "        <br> (�¶��� �Ա��Ͻ� ���� �������� / 534-01-016039 / (��)ť�� Ŀ�´ϸ��� �Դϴ�.)" + vbcr
        mailcontent = mailcontent + "        ����� �� 2�Ͽ��� 4�� ������ �ҿ�Ǹ�, �ֹ������� ���� ���������̳� ���ǻ����� <br> " + vbcr
        mailcontent = mailcontent + "        �̸���(<a href=mailto:customer@10X10.co.kr>customer@10X10.co.kr</a>)�̳� 02-515-5945�� " + vbcr
        mailcontent = mailcontent + "        �����ֽñ� �ٶ��ϴ�.<br> " + vbcr
        mailcontent = mailcontent + "        <br> " + vbcr
        mailcontent = mailcontent + "      </td> " + vbcr
        mailcontent = mailcontent + "      <td rowspan=5 bgcolor=F1F1F1></td> " + vbcr
        mailcontent = mailcontent + "    </tr> " + vbcr
        mailcontent = mailcontent + "    <tr> " + vbcr
        mailcontent = mailcontent + "      <td bgcolor=F1F1F1><img src=http://www.10x10.co.kr/images/slice01_07.gif width=600 height=4></td> " + vbcr
        mailcontent = mailcontent + "    </tr> " + vbcr
        mailcontent = mailcontent + "    <tr> " + vbcr
        mailcontent = mailcontent + "      <td bgcolor=F1F1F1><br> " + vbcr

        '�ֹ����� Ȯ��.
        sql = "select regdate, reqzipcode, (b.addr010_si + ' ' + b.addr010_gu + ' ' + a.reqaddress) as reqalladdress, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice from tbl_order_master a, addr010tl b, tbl_order_detail c"
        sql = sql + " where b.addr010_zip1 = left(a.reqzipcode,3) and b.addr010_zip2 = right(a.reqzipcode,3) and a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                discountrate = rsfunc("discountrate")
                rsfunc.Movefirst
                mailcontent = mailcontent + "  <table width=600 border=0> " + vbcr
                mailcontent = mailcontent + "    <tr> " + vbcr
                mailcontent = mailcontent + "      <td><img src=http://www.10x10.co.kr/images/order_ment02.gif width=150 height=35 vspace=5 hspace=0></td> " + vbcr
                mailcontent = mailcontent + "            <td><font color=990000>�ֹ� ��ȣ : " + orderserial + " &nbsp;&nbsp;|&nbsp;�ֹ� ���� : " + cStr(year(rsfunc("regdate"))) + "�� " + cStr(month(rsfunc("regdate"))) + "�� " + cStr(day(rsfunc("regdate"))) + "��<br> " + vbcr
                mailcontent = mailcontent + "              �� �� �� : [" + rsfunc("reqzipcode") + "] " + rsfunc("reqalladdress") + "<br> " + vbcr
                mailcontent = mailcontent + "        �ֹ� �Ѿ� : " + cstr(rsfunc("subtotalprice")) + "�� = �Ұ� : " + cstr(rsfunc("subtotalprice") - rsfunc("itemcost")) + "�� (" + cstr(rsfunc("totalmileage")) + "����Ʈ) + ��ۺ� : " + cstr(rsfunc("itemcost")) + "��</font> </td> " + vbcr
                mailcontent = mailcontent + "    </tr> " + vbcr
                mailcontent = mailcontent + "  </table> " + vbcr
        else
                exit sub
        end if
        rsfunc.close

        '�ֹ������� ���� Ȯ��.
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

                        '�ɼ� ǥ�úκ�. �ϴ� ����.

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
        mailcontent = mailcontent + "      <td align=right valign=top>(��)ť�� Ŀ�´ϸ���<img src=http://www.10x10.co.kr/images/cube_ci.gif width=210 height=52 hspace=15></td> "
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
        mailtitle = "�ֹ��� ���������� �����Ǿ����ϴ�!"

        '�ֹ��� �����ּ� Ȯ��,�ֹ��ŷ����� ����
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

        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        if paymethod = "7" then    ' ������
            fileName = dirPath&"\\email_bank1.htm"
        elseif paymethod = "100" then   ' �ſ�ī��
            fileName = dirPath&"\\email_card1.htm"
        end if

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
'        mailcontent = replace(mailcontent,":USERNAME:",userName)



        '�ֹ����� Ȯ��.
        sql = "select buyname,regdate, reqzipcode, (b.addr010_si + ' ' + b.addr010_gu + ' ' + a.reqaddress) as reqalladdress, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice from tbl_order_master a, addr010tl b, tbl_order_detail c"
        sql = sql + " where b.addr010_zip1 = left(a.reqzipcode,3) and b.addr010_zip2 = right(a.reqzipcode,3) and a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                discountrate = rsfunc("discountrate")
                rsfunc.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", FormatCurrency(cstr(rsfunc("subtotalprice")))) ' �ֹ��Ѿ�
                mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(rsfunc("subtotalprice") - rsfunc("itemcost"))) ) ' �ֹ��� ��item  ����
                mailcontent = replace(mailcontent,":DELIVERYFEE:",  FormatCurrency(cstr(rsfunc("itemcost"))) ) ' ��۱ݾ�
                mailcontent = replace(mailcontent,":BUYNAME:", rsfunc("buyname")) ' �ֹ��� �̸�
                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
                mailcontent = replace(mailcontent,":REQZIPCODE:", rsfunc("reqzipcode")) ' ��ۿ����ȣ
                mailcontent = replace(mailcontent,":REQALLADDRESS:", rsfunc("reqalladdress")) ' ����ּ�
        else
                exit function
        end if
        rsfunc.close

        'item ���� �յںκ� ¥����
        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item ������ �����κ� �ڸ���
        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)

        '�ֹ������� ���� Ȯ��.
        dim itemserial,inx
        sql = " select a.itemid, b.imglist, c.itemname, (c.cate_large + c.cate_mid + c.cate_small) as itemserial, c.sellcash, a.itemno from tbl_order_detail a, tbl_item_image b, tbl_item c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and a.itemid <> '0' and a.itemid = b.itemid and c.itemid = a.itemid"
        inx = 1
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                rsfunc.Movefirst
                do until rsfunc.eof
                        itemserial = rsfunc("itemserial") + "-" + FormatCode(rsfunc("itemid"))
                        itemHtml = replace(itemHtmlOri,":ITEMSERIAL:", itemserial) ' ��ǰ�ڵ�
                        itemHtml = replace(itemHtml,":ITEMNAME:", rsfunc("itemname")) ' ��ǰ�̸�
                        if discountrate=1 then
                        	itemHtml = replace(itemHtml,":ITEMPRICE:",  CStr(rsfunc("sellcash"))) ' ��ǰ����
                        else
                        	itemHtml = replace(itemHtml,":ITEMPRICE:",  cstr(round(rsfunc("sellcash")*cdbl(discountrate)/100)*100) ) ' ��ǰ����
                    	end if
                        itemHtml = replace(itemHtml,":ITEMNO:", cstr(rsfunc("itemno"))) ' ����
                        itemHtml = replace(itemHtml,":IMGLIST:", cstr(rsfunc("imglist"))) ' ����
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


        mailtitle = "�ֹ��� ���������� �����Ǿ����ϴ�!"

        '�ֹ��� �����ּ� Ȯ��,�ֹ��ŷ����� ����---------------------------------------------------------------------------
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

        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        if paymethod = "7" then    ' ������
            fileName = dirPath&"\\email_bank1.htm"
        elseif paymethod = "100" then   ' �ſ�ī��
            fileName = dirPath&"\\email_card1.htm"
        end if

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
'        mailcontent = replace(mailcontent,":USERNAME:",userName)


		dim SpendMile
        '�ֹ����� Ȯ��.---------------------------------------------------------------------------
        sql = "select buyname,regdate, reqzipcode, (b.addr010_si + ' ' + b.addr010_gu + ' ' + a.reqaddress) as reqalladdress, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice, a.miletotalprice from tbl_order_master a, addr010tl b, tbl_order_detail c"
        sql = sql + " where b.addr010_zip1 = left(a.reqzipcode,3) and b.addr010_zip2 = right(a.reqzipcode,3) and a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                discountrate = rsfunc("discountrate")
                rsfunc.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", FormatCurrency(cstr(rsfunc("subtotalprice")))) ' �ֹ��Ѿ�
                'mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(Titemcost - rsfunc("itemcost"))) ) ' �ֹ��� ��item  ����
                mailcontent = replace(mailcontent,":DELIVERYFEE:",  FormatCurrency(cstr(rsfunc("itemcost"))) ) ' ��۱ݾ�
                mailcontent = replace(mailcontent,":BUYNAME:", rsfunc("buyname")) ' �ֹ��� �̸�
                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
                mailcontent = replace(mailcontent,":REQZIPCODE:", rsfunc("reqzipcode")) ' ��ۿ����ȣ
                mailcontent = replace(mailcontent,":REQALLADDRESS:", rsfunc("reqalladdress")) ' ����ּ�

                if IsNull(rsfunc("miletotalprice")) then
                	SpendMile =""
                else
                	SpendMile = rsfunc("miletotalprice")
                	SpendMile = "(���ϸ������: " + formatNumber(FormatCurrency(SpendMile),0) + " )"
            	end if
            	mailcontent = replace(mailcontent,":SPENDMILEAGE:", SpendMile) ' ���ϸ���
        else
                exit function
        end if
        rsfunc.close

        'item ���� �յںκ� ¥����
        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item ������ �����κ� �ڸ���
        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)




		'�ֹ������� ���� Ȯ��.-----------------------------------------------------------------------------
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
                        itemHtml = replace(itemHtmlOri,":ITEMSERIAL:", itemserial) ' ��ǰ�ڵ�
                        itemHtml = replace(itemHtml,":ITEMNAME:", rsfunc("itemname")) ' ��ǰ�̸�

                        if CDbl(discountrate)=1 then
                        	BufCost = rsfunc("sellcash") * rsfunc("itemno")
                        	Titemcost = Titemcost + BufCost
                        	itemHtml = replace(itemHtml,":ITEMPRICE:", CStr(BufCost) ) ' ��ǰ����
                        else
                        	BufCost = round(rsfunc("sellcash")*cdbl(discountrate)/100)*100 * rsfunc("itemno")
                        	Titemcost = Titemcost + BufCost
                        	itemHtml = replace(itemHtml,":ITEMPRICE:", CStr(BufCost) ) ' ��ǰ����
                    	end if
                        itemHtml = replace(itemHtml,":ITEMNO:", cstr(rsfunc("itemno"))) ' ����
                        itemHtml = replace(itemHtml,":IMGLIST:", cstr(rsfunc("imglist"))) ' ����
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

		mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(Titemcost)) ) ' �ֹ��� ��item  ����

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailorder3 = mailcontent
end function

function ReSendmailorder(orderserial,mailfrom)
        sendmailorder3 orderserial,mailfrom
end function

function sendmailcome(orderserial) ' �������ɽ� ���� ������
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "10X10 ���� �ȳ� �����Դϴ�!"

        '�ֹ��� �����ּ� Ȯ��,�ֹ��ŷ����� ����
        sql = "select buyemail,accountdiv from tbl_order_master where orderserial = '" + orderserial + "'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                rsfunc.Movefirst
                mailto = rsfunc("buyemail")
        else
                exit function
        end if
        rsfunc.close

        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_come.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall

        '�ֹ����� Ȯ��.
        sql = "select buyname,regdate, reqzipcode, (b.addr010_si + ' ' + b.addr010_gu + ' ' + a.reqaddress) as reqalladdress, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice from tbl_order_master a, addr010tl b, tbl_order_detail c"
        sql = sql + " where b.addr010_zip1 = left(a.reqzipcode,3) and b.addr010_zip2 = right(a.reqzipcode,3) and a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                discountrate = rsfunc("discountrate")
                rsfunc.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", FormatCurrency(cstr(rsfunc("subtotalprice")))) ' �ֹ��Ѿ�
                mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(rsfunc("subtotalprice") - rsfunc("itemcost"))) ) ' �ֹ��� ��item  ����
                mailcontent = replace(mailcontent,":DELIVERYFEE:",  FormatCurrency(cstr(rsfunc("itemcost"))) ) ' ��۱ݾ�
                mailcontent = replace(mailcontent,":BUYNAME:", rsfunc("buyname")) ' �ֹ��� �̸�
                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
                mailcontent = replace(mailcontent,":REQZIPCODE:", rsfunc("reqzipcode")) ' ��ۿ����ȣ
                mailcontent = replace(mailcontent,":REQALLADDRESS:", rsfunc("reqalladdress")) ' ����ּ�
        else
                exit function
        end if
        rsfunc.close

        'item ���� �յںκ� ¥����
        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item ������ �����κ� �ڸ���
        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)

        '�ֹ������� ���� Ȯ��.
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
                        itemHtml = replace(itemHtmlOri,":ITEMSERIAL:", itemserial) ' ��ǰ�ڵ�
                        itemHtml = replace(itemHtml,":ITEMNAME:", rsfunc("itemname")) ' ��ǰ�̸�
                        itemHtml = replace(itemHtml,":ITEMPRICE:",  cstr(rsfunc("sellcash")*cdbl(discountrate)) ) ' ��ǰ����
                        itemHtml = replace(itemHtml,":ITEMNO:", cstr(rsfunc("itemno"))) ' ����
                        itemHtml = replace(itemHtml,":IMGLIST:", cstr(rsfunc("imglist"))) ' ��ǰ�̹���
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

function sendmailbankok(mailto,userName,orderserial) ' �Ա�Ȯ�θ���
        dim sql,discountrate
        dim mailfrom, mailtitle, mailcontent
        dim fs,objFile,dirPath,fileName

        mailfrom = "�ٹ�����<customer@10x10.co.kr>"
        mailtitle = "������ �Ա��� ���������� ó�� �Ǿ����ϴ�!"

        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_bank2011.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":USERNAME:",userName)
        mailcontent = replace(mailcontent,":ORDERSERIAL:",orderserial)

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
end function

function sendmailbankokNoDLV(mailto,userName,orderserial) ' �Ա�Ȯ�θ���
        dim sql,discountrate
        dim mailfrom, mailtitle, mailcontent
        dim fs,objFile,dirPath,fileName

        mailfrom = "�ٹ�����<customer@10x10.co.kr>"
        mailtitle = "������ �Ա��� ���������� ó�� �Ǿ����ϴ�!"

        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_bank2011.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":USERNAME:",userName)
        mailcontent = replace(mailcontent,":ORDERSERIAL:",orderserial)
        mailcontent = replace(mailcontent,"���� ���ϳ��� ����� �̷�� �� �� �ֵ��� ����ϰڽ��ϴ�.","")

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
end function

function sendmailbankok_GIFTCard(mailto,userName,orderserial) ' �Ա�Ȯ�θ���
        dim sql,discountrate
        dim mailfrom, mailtitle, mailcontent
        dim fs,objFile,dirPath,fileName

        mailfrom = "�ٹ�����<customer@10x10.co.kr>"
        mailtitle = "������ �Ա��� ���������� ó�� �Ǿ����ϴ�!"

        ' ������ �ҷ��ͼ�
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
        mailtitle = "�ֹ��Ͻ� ��ǰ�� ���� �ٹ����� ��۾ȳ��Դϴ�!"
        '�ֹ��� �����ּ� Ȯ��,�ֹ��ŷ����� ����
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

        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_finish.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall

        '�ֹ����� Ȯ��.
        sql = "select buyname,regdate, reqzipcode, (b.addr010_si + ' ' + b.addr010_gu + ' ' + a.reqaddress) as reqalladdress, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice from tbl_order_master a, addr010tl b, tbl_order_detail c"
        sql = sql + " where b.addr010_zip1 = left(a.reqzipcode,3) and b.addr010_zip2 = right(a.reqzipcode,3) and a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsfunc.Open sql,dbfunc,1
        if  not rsfunc.EOF  then
                discountrate = rsfunc("discountrate")
                rsfunc.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", FormatCurrency(cstr(rsfunc("subtotalprice")))) ' �ֹ��Ѿ�
                mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(rsfunc("subtotalprice") - rsfunc("itemcost"))) ) ' �ֹ��� ��item  ����
                mailcontent = replace(mailcontent,":DELIVERYFEE:",  FormatCurrency(cstr(rsfunc("itemcost"))) ) ' ��۱ݾ�
                mailcontent = replace(mailcontent,":DELIVERNO:",  deliverno ) ' ������ȣ
                mailcontent = replace(mailcontent,":BUYNAME:", rsfunc("buyname")) ' �ֹ��� �̸�
                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
                mailcontent = replace(mailcontent,":REQZIPCODE:", rsfunc("reqzipcode")) ' ��ۿ����ȣ
                mailcontent = replace(mailcontent,":REQALLADDRESS:", rsfunc("reqalladdress")) ' ����ּ�
        else
                exit function
        end if
        rsfunc.close

        'item ���� �յںκ� ¥����
        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item ������ �����κ� �ڸ���
        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)

        '�ֹ������� ���� Ȯ��.
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
                        itemHtml = replace(itemHtmlOri,":ITEMSERIAL:", itemserial) ' ��ǰ�ڵ�
                        itemHtml = replace(itemHtml,":ITEMNAME:", rsfunc("itemname")) ' ��ǰ�̸�
                        itemHtml = replace(itemHtml,":ITEMPRICE:",  cstr(rsfunc("sellcash")*cdbl(discountrate)) ) ' ��ǰ����
                        itemHtml = replace(itemHtml,":ITEMNO:", cstr(rsfunc("itemno"))) ' ����
                        itemHtml = replace(itemHtml,":IMGLIST:", cstr(rsfunc("imglist"))) ' ��ǰ�̹���
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
        mailtitle = "��ǰ�� ���Ǿ����ϴ�!"

        ' ������ �ҷ��ͼ�
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
			mailcontent = replace(mailcontent,":BUYNAME:", db2html(rsget("buyname"))) ' �ֹ��� �̸�

			mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ

			tensongjangno = rsget("deliverno")
			ipkumdiv = rsget("ipkumdiv")
		else
			exit function
		end if
		rsget.close

'���ٹ�ۻ��� - ������.

		if ipkumdiv="0" then
			IpkumDivName="�ֹ����"
		elseif ipkumdiv="1" then
			IpkumDivName="�ֹ�����"
		elseif ipkumdiv="2" then
			IpkumDivName="�ֹ�����"
		elseif ipkumdiv="3" then
			IpkumDivName="�ֹ�����"
		elseif ipkumdiv="4" then
			IpkumDivName="�����Ϸ�"
		elseif ipkumdiv="5" then
			IpkumDivName="��۴��"
		elseif ipkumdiv="6" then
			IpkumDivName="��۴��"
		elseif ipkumdiv="7" then
			IpkumDivName="��ǰ���"
		elseif ipkumdiv="8" then
			IpkumDivName="��ǰ���"
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
									sinx = 0 '�ҼӾ�ü ó�� ����
									einx = 1
								end if
						  elseif inx <> 0 and rsget("makerid") = designerid and rsget("currstate") <> 7 then
									einx = 0
									sinx = 0 '�ҼӾ�ü���� �̹߼� ��ǰ ù ����
						  elseif einx = 1 and rsget("makerid") <> designerid then
									einx = 0
									sinx = 0 '�ҼӾ�ü�̿� ��ǰ ù ����
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
itemHtml = itemHtml + "<td width='50' class='p11' align='center'>��ǰ</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' class='p11' align='center'>��ǰ��</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>�ɼ�</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='30' class='p11' align='center'>����</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>�����Ȳ</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='100' class='p11' align='center'>�ù�/����</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "</td>"
itemHtml = itemHtml + "</tr>"
end if


'��ۻ��� ����
if rsget("deliverytype") = 1 or rsget("deliverytype") = 4 then
    if rsget("currstate") = 7 then
		 BaesongState = "<font color='red'>���Ϸ�</font>"
	 else
		 BaesongState = "<font color='#004080'>��ǰ�غ���</font>"
	 end if

    ''BaesongState = IpkumDivName '���ٹ�ۻ���
else
	 if rsget("currstate") = 7 then
		 BaesongState = "<font color='red'>���Ϸ�</font>"
	 else
		 BaesongState = "<font color='#004080'>��ǰ�غ���</font>"
	 end if
end if


'�ù�/���� ����

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

		itemHtmlTotal = replace(mailcontent,":INNERORDERTABLE:", itemHtml) ' �ֹ��������̺� �ֱ�

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
        mailtitle = "������ ���� �ȳ� ���� �Դϴ�.!"

        ' ������ �ҷ��ͼ�
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
        mailtitle = "��ǰ�� ���Ǿ����ϴ�!"

        '�ֹ��� �����ּ� Ȯ��,�ֹ��ŷ����� ����---------------------------------------------------------------------------
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
        '�ֹ����� Ȯ��.---------------------------------------------------------------------------
        sql = "select buyname,regdate, reqname, reqzipcode, (a.reqzipaddr + ' ' + a.reqaddress) as reqalladdress, a.reqphone, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice, a.miletotalprice ,a.tencardspend, a.comment from [db_order].[dbo].tbl_order_master a, [db_order].[dbo].tbl_order_detail c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"

		rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                discountrate = rsget("discountrate")
                tencardspend = rsget("tencardspend")
                rsget.Movefirst
                subtotalprice = formatNumber(FormatCurrency(rsget("subtotalprice")),0) ' �ֹ��Ѿ�
                itemcost = formatNumber(FormatCurrency(rsget("itemcost")),0) ' ��۱ݾ�
                buyname = rsget("buyname") ' �ֹ��� �̸�
                reqname = rsget("reqname") ' ������ �̸�
                reqalladdress = rsget("reqalladdress") ' ����ּ�
                reqphone = rsget("reqphone") ' �ֹ��� ��ȭ��ȣ
                comment = rsget("comment") ' ��۸޸�
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
mailcontent = mailcontent + "<title>[�ٹ�����] ��ſ��� ������ ���θ� 10x10 = tenbyten</title>"
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
mailcontent = mailcontent + "<td><b>[" + buyname + "]���� �ֹ������Դϴ� </b></td>"
mailcontent = mailcontent + "<td>"
mailcontent = mailcontent + "<div align='right'><b>�ֹ���ȣ : <font color='#CC3300'><span class='verdana-mid'>" + orderserial + "</span></font></b></div>"
mailcontent = mailcontent + "</td>"
mailcontent = mailcontent + "</tr>"
mailcontent = mailcontent + "</table>"
mailcontent = mailcontent + "</div>"
mailcontent = mailcontent + "</td>"
mailcontent = mailcontent + "</tr>"


		'�ֹ������� ���� Ȯ��.-----------------------------------------------------------------------------

		dim itemserial,inx,inx2,tdata,tdata2
      dim Titemcost,BufCost
		dim upchebaesong
		dim currstate
		dim transco,transurl
        Titemcost = 0

		'��ü��� ���� ��ǰ ��������
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
					currstate = "<font color='#46A3FF'>��ǰ�غ���</font>"
					elseif rsget("currstate") = 7 then
					currstate = "<font color='#FF6060'>���Ϸ�</font>"
					else
					currstate = "<font color='#939300'>��ǰ�غ���</font>"
					end if

					if rsget("songjangdiv") = "1" then
					transco = "�����ù�"
					transurl = "http://www.hanjin.co.kr/transmission/main.htm"
					elseif rsget("songjangdiv") = "2" then
					transco = "�����ù�"
					transurl = "http://www.hyundaiexpress.com/hydex/jsp/support/search/re_03.jsp"
					elseif rsget("songjangdiv") = "3" then
					transco = "�������"
					transurl = "http://doortodoor.korex.co.kr/jsp/cmn/index.jsp"
					elseif rsget("songjangdiv") = "4" then
					transco = "CJ GLS"
					transurl = "http://www.cjgls.com/contents/gls/gls004/gls004_06.asp"
					elseif rsget("songjangdiv") = "5" then
					transco = "��Ŭ����"
					transurl = "http://www.ecline.net/tracking/customer02.html#t01"
					elseif rsget("songjangdiv") = "6" then
					transco = "HTH"
					transurl = "https://samsunghth.com/homepage/searchTraceGoods/SearchTraceInput.jhtml?mc=5"
					elseif rsget("songjangdiv") = "7" then
					transco = "�ѹ̸��ù�"
					transurl = "http://www.e-family.co.kr/"
					elseif rsget("songjangdiv") = "8" then
					transco = "��ü��"
					transurl = "http://service.epost.go.kr/kps_index.html"
					elseif rsget("songjangdiv") = "9" then
					transco = "KGB"
					transurl = "http://www.kgbl.co.kr/"
					elseif rsget("songjangdiv") = "10" then
					transco = "�����ù�"
					transurl = "http://www.ajulogis.co.kr/"
					elseif rsget("songjangdiv") = "11" then
					transco = "�������ù�"
					transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
					elseif rsget("songjangdiv") = "12" then
					transco = "�ѱ��ù�"
					transurl = "http://www.kls.co.kr/"
					elseif rsget("songjangdiv") = "13" then
					transco = "���ο�ĸ"
					transurl = "http://www.yellowcap.co.kr/"
					elseif rsget("songjangdiv") = "14" then
					transco = "���̽��ù�"
					transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
					elseif rsget("songjangdiv") = "15" then
					transco = "�߾��ù�"
					transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
					elseif rsget("songjangdiv") = "16" then
					transco = "�����ù�"
					transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
					elseif rsget("songjangdiv") = "17" then
					transco = "Ʈ����ù�"
					transurl = "http://www.transclub.com/"
					else
					transco = "��Ÿ"
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
						mailcontent = mailcontent + "<td width='50' class='p11' align='center'>��ǰ</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' class='p11' align='center'>��ǰ��</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='50' class='p11' align='center'>�ɼ�</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='50' class='p11' align='center'>����</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>����</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>�����Ȳ</td>"
						mailcontent = mailcontent + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>�ù�/����</td>"
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



'' E-giftī�� ����
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

    ''' �̰����� ����.
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

function sendmailStockOutAlarm(orderserial) ' ǰ���ȳ� ����
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal, accountdiv, ipkumdate, totItemPrice
		dim sqlStr

        if isnull(orderserial) or orderserial="" then exit function

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "[�ٹ�����] ǰ���ȳ� �����Դϴ�."

        '�ֹ��� �����ּ� Ȯ��,�ֹ��ŷ����� ����
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

        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_stockoutalarm.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall


		mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
		if isnull(ipkumdate) or ipkumdate="" then
			mailcontent = replace(mailcontent,":IPKUMDATE:", "")
		else
			mailcontent = replace(mailcontent,":IPKUMDATE:", ipkumdate)
		end if

        if accountdiv = "7" then    ' ������
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�������Ա�")
        elseif accountdiv = "100" then   ' �ſ�ī��
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ſ�ī��")
        elseif accountdiv = "20" then   ' �ǽð���ü
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ǽð���ü")
        elseif accountdiv = "80" then   ' �þ�
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�þ�ī��")
        elseif accountdiv = "110" then   ' OKCashbag+�ſ�ī��
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "OKCashbag+�ſ�ī��")
        elseif accountdiv = "400" then   ' �ڵ�������
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ڵ���")
        else
        	mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "")
        end if

        '�ֹ������� ���� Ȯ��.
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
		itemHtmlOri = itemHtmlOri & "												<td style=""width:100px;margin:0;  padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-size:11px; line-height:11px; font-family:dotum, '����', sans-serif; color:#707070;"">:ITEMID:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:295px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:left; font-size:11px; line-height:17px; font-family:dotum, '����', sans-serif; color:#707070;"">:ITEMNAME::ITEMOPTIONNAME:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:12px; text-align:right; font-family:dotum, '����', sans-serif;"">" & vbCrLf
		itemHtmlOri = itemHtmlOri & "													<span style=""margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, '����', sans-serif; text-align:right;"">:ITEMCOST:</span><br />" & vbCrLf
		itemHtmlOri = itemHtmlOri & "													:REDUCEDPRICE:" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:25px; padding:6px 0; border-bottom:solid 1px #eaeaea;""></td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px; line-height:12px;"">:ITEMNO:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "											</tr>" & vbCrLf

        if  not rsget.EOF  then
            rsget.Movefirst
            do until rsget.eof
				itemHtml = replace(itemHtmlOri,":ITEMIMAGE:", webImgUrl & "/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("imgsmall"))
				itemHtml = replace(itemHtml,":ITEMID:", rsget("itemid")) ' ��ǰ�ڵ�
				itemHtml = replace(itemHtml,":ITEMNAME:", db2html(rsget("itemname")))
				if (rsget("itemoption") = "0000") then
					itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "")
				else
					itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "<br />" & db2html(rsget("itemoptionname")))
				end if

				itemHtml = replace(itemHtml,":ITEMCOST:", FormatNumber(rsget("itemcostCouponNotApplied"), 0) & "��")
				if (rsget("reducedPrice") < rsget("itemcostCouponNotApplied")) then
					itemHtml = replace(itemHtml,":REDUCEDPRICE:", "<span style=""margin:0; padding:6px 0; color:#dd5555; font-size:11px; line-height:17px; text-align:center""><img src=""http://mailzine.10x10.co.kr/2017/ico_coupon.png"" alt=""��������"" style=""vertical-align:-2px; padding-right:2px;""/>" & FormatNumber(rsget("reducedPrice"), 0) & "��" & "</span>")
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
        mailcontent = replace(mailcontent,":TOTITEMPRICE:", FormatNumber(totItemPrice, 0) & "��")

	'//=======  ���� �߼� =========/
	dim oMail
	set oMail = New MailCls

	IF replace(mailto,"'","")<>"" THEN

		oMail.MailTitles	= replace(mailtitle,"'","")
		oMail.SenderNm		= "�ٹ�����"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= replace(mailto,"'","")
		oMail.ReceiverMail	= replace(mailto,"'","")
		oMail.MailConts 	= newhtml2db(mailcontent)
		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
		oMail.Send_TMSMailer()		'TMS���Ϸ�
		'oMail.Send_Mailer()
		''oMail.Send_CDO()

	End IF

	SET oMail = nothing

        sendmailStockOutAlarm = mailcontent
end function

' �ù��ľ� �ȳ� ����
function sendmailDeliverystrikeAlarm(orderserial)
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal, accountdiv, ipkumdate, totItemPrice
		dim sqlStr

        if isnull(orderserial) or orderserial="" then exit function

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "[�ٹ�����] �ù��ľ� �ȳ� �����Դϴ�."

        '�ֹ��� �����ּ� Ȯ��,�ֹ��ŷ����� ����
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

        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_Deliverystrikealarm.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall


		mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
		if isnull(ipkumdate) or ipkumdate="" then
			mailcontent = replace(mailcontent,":IPKUMDATE:", "")
		else
			mailcontent = replace(mailcontent,":IPKUMDATE:", ipkumdate)
		end if

        if accountdiv = "7" then    ' ������
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�������Ա�")
        elseif accountdiv = "100" then   ' �ſ�ī��
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ſ�ī��")
        elseif accountdiv = "20" then   ' �ǽð���ü
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ǽð���ü")
        elseif accountdiv = "80" then   ' �þ�
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�þ�ī��")
        elseif accountdiv = "110" then   ' OKCashbag+�ſ�ī��
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "OKCashbag+�ſ�ī��")
        elseif accountdiv = "400" then   ' �ڵ�������
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ڵ���")
        else
        	mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "")
        end if

        '�ֹ������� ���� Ȯ��.
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
		itemHtmlOri = itemHtmlOri & "												<td style=""width:100px;margin:0;  padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-size:11px; line-height:11px; font-family:dotum, '����', sans-serif; color:#707070;"">:ITEMID:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:295px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:left; font-size:11px; line-height:17px; font-family:dotum, '����', sans-serif; color:#707070;"">:ITEMNAME::ITEMOPTIONNAME:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:12px; text-align:right; font-family:dotum, '����', sans-serif;"">" & vbCrLf
		itemHtmlOri = itemHtmlOri & "													<span style=""margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, '����', sans-serif; text-align:right;"">:ITEMCOST:</span><br />" & vbCrLf
		itemHtmlOri = itemHtmlOri & "													:REDUCEDPRICE:" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:25px; padding:6px 0; border-bottom:solid 1px #eaeaea;""></td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px; line-height:12px;"">:ITEMNO:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "											</tr>" & vbCrLf

        if  not rsget.EOF  then
            rsget.Movefirst
            do until rsget.eof
				itemHtml = replace(itemHtmlOri,":ITEMIMAGE:", webImgUrl & "/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("imgsmall"))
				itemHtml = replace(itemHtml,":ITEMID:", rsget("itemid")) ' ��ǰ�ڵ�
				itemHtml = replace(itemHtml,":ITEMNAME:", db2html(rsget("itemname")))
				if (rsget("itemoption") = "0000") then
					itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "")
				else
					itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "<br />" & db2html(rsget("itemoptionname")))
				end if

				itemHtml = replace(itemHtml,":ITEMCOST:", FormatNumber(rsget("itemcostCouponNotApplied"), 0) & "��")
				if (rsget("reducedPrice") < rsget("itemcostCouponNotApplied")) then
					itemHtml = replace(itemHtml,":REDUCEDPRICE:", "<span style=""margin:0; padding:6px 0; color:#dd5555; font-size:11px; line-height:17px; text-align:center""><img src=""http://mailzine.10x10.co.kr/2017/ico_coupon.png"" alt=""��������"" style=""vertical-align:-2px; padding-right:2px;""/>" & FormatNumber(rsget("reducedPrice"), 0) & "��" & "</span>")
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
        mailcontent = replace(mailcontent,":TOTITEMPRICE:", FormatNumber(totItemPrice, 0) & "��")

	'//=======  ���� �߼� =========/
	dim oMail
	set oMail = New MailCls

	IF replace(mailto,"'","")<>"" THEN

		oMail.MailTitles	= replace(mailtitle,"'","")
		oMail.SenderNm		= "�ٹ�����"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= replace(mailto,"'","")
		oMail.ReceiverMail	= replace(mailto,"'","")
		oMail.MailConts 	= newhtml2db(mailcontent)
		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
		oMail.Send_TMSMailer()		'TMS���Ϸ�
		'oMail.Send_Mailer()
		''oMail.Send_CDO()

	End IF

	SET oMail = nothing

        sendmailDeliverystrikeAlarm = mailcontent
end function

' �ֹ������ȳ�����      ' 2020.10.27 �ѿ��
function sendmaildelayAlarm(orderserial)
        dim sql,discountrate,paymethod, sqlStr, mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal, accountdiv, ipkumdate, totItemPrice

        if isnull(orderserial) or orderserial="" then exit function

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "[�ٹ�����] �߼����� �ȳ� �����Դϴ�."

        '�ֹ��� �����ּ� Ȯ��,�ֹ��ŷ����� ����
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

        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_stockoutalarm_customer.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall


		mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
		if isnull(ipkumdate) or ipkumdate="" then
			mailcontent = replace(mailcontent,":IPKUMDATE:", "")
		else
			mailcontent = replace(mailcontent,":IPKUMDATE:", ipkumdate)
		end if

        if accountdiv = "7" then    ' ������
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�������Ա�")
        elseif accountdiv = "100" then   ' �ſ�ī��
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ſ�ī��")
        elseif accountdiv = "20" then   ' �ǽð���ü
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ǽð���ü")
        elseif accountdiv = "80" then   ' �þ�
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�þ�ī��")
        elseif accountdiv = "110" then   ' OKCashbag+�ſ�ī��
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "OKCashbag+�ſ�ī��")
        elseif accountdiv = "400" then   ' �ڵ�������
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ڵ���")
        else
        	mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "")
        end if

        '�ֹ������� ���� Ȯ��.
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
		sqlStr = sqlStr + " and bl.orderserial is null"		' ǰ������ �߼��� �ִ°��, ������� ���ڴ� ������ �ʴ´�.
		sqlStr = sqlStr + " order by d.itemid, d.itemoption "

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		itemHtmlTotal = ""

		itemHtmlOri = "											<tr>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:50px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea;""><img src="":ITEMIMAGE:"" alt="""" /></td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:100px;margin:0;  padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-size:11px; line-height:11px; font-family:dotum, '����', sans-serif; color:#707070;"">:ITEMID:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:295px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:left; font-size:11px; line-height:17px; font-family:dotum, '����', sans-serif; color:#707070;"">:ITEMNAME::ITEMOPTIONNAME:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:12px; text-align:right; font-family:dotum, '����', sans-serif;"">" & vbCrLf
		itemHtmlOri = itemHtmlOri & "													<span style=""margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, '����', sans-serif; text-align:right;"">:ITEMCOST:</span><br />" & vbCrLf
		itemHtmlOri = itemHtmlOri & "													:REDUCEDPRICE:" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:25px; padding:6px 0; border-bottom:solid 1px #eaeaea;""></td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "												<td style=""width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px; line-height:12px;"">:ITEMNO:</td>" & vbCrLf
		itemHtmlOri = itemHtmlOri & "											</tr>" & vbCrLf

        if  not rsget.EOF  then
            rsget.Movefirst
            do until rsget.eof
				itemHtml = replace(itemHtmlOri,":ITEMIMAGE:", webImgUrl & "/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("imgsmall"))
				itemHtml = replace(itemHtml,":ITEMID:", rsget("itemid")) ' ��ǰ�ڵ�
				itemHtml = replace(itemHtml,":ITEMNAME:", db2html(rsget("itemname")))
				if (rsget("itemoption") = "0000") then
					itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "")
				else
					itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "<br />" & db2html(rsget("itemoptionname")))
				end if

				itemHtml = replace(itemHtml,":ITEMCOST:", FormatNumber(rsget("itemcostCouponNotApplied"), 0) & "��")
				if (rsget("reducedPrice") < rsget("itemcostCouponNotApplied")) then
					itemHtml = replace(itemHtml,":REDUCEDPRICE:", "<span style=""margin:0; padding:6px 0; color:#dd5555; font-size:11px; line-height:17px; text-align:center""><img src=""http://mailzine.10x10.co.kr/2017/ico_coupon.png"" alt=""��������"" style=""vertical-align:-2px; padding-right:2px;""/>" & FormatNumber(rsget("reducedPrice"), 0) & "��" & "</span>")
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
        mailcontent = replace(mailcontent,":TOTITEMPRICE:", FormatNumber(totItemPrice, 0) & "��")

	'//=======  ���� �߼� =========/
	dim oMail
	set oMail = New MailCls

	IF replace(mailto,"'","")<>"" THEN

		oMail.MailTitles	= replace(mailtitle,"'","")
		oMail.SenderNm		= "�ٹ�����"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= replace(mailto,"'","")
		oMail.ReceiverMail	= replace(mailto,"'","")
		oMail.MailConts 	= newhtml2db(mailcontent)
		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
		oMail.Send_TMSMailer()		'TMS���Ϸ�
		'oMail.Send_Mailer()
		''oMail.Send_CDO()

	End IF

	SET oMail = nothing

        sendmaildelayAlarm = mailcontent
end function
%>
