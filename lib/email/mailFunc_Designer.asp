<!-- #include virtual="/lib/email/mailFunction.asp" -->
<%
'+--------------------------------------------------------------------------------------------------------------------------------+
'|                                        ��ü ��� ��ǰ ���� �߼�                                                                |
'+----------------------------------------------------+---------------------------------------------------------------------------+
'|             �� �� ��                               |                          ��    ��                                         |
'+----------------------------------------------------+---------------------------------------------------------------------------+
'| fcSendMailFinish_Dlv_Designer(orderserial,makerid) | ��� ���� �߼�(��ü��� ���)                                             |
'|                                                    | ��뿹 : fcSendMailFinish_Dlv_Designer('012012304','1293495006')          |
'+----------------------------------------------------+---------------------------------------------------------------------------+
'| fcSendMailFinish_Dlv_Designer_off(detailidx,makerid)  �������� ��� ���� �߼�(��ü��� ���)                                   |
'|                                                    | ��뿹 : fcSendMailFinish_Dlv_Designer_off('012012304','1293495006')      |
'+----------------------------------------------------+---------------------------------------------------------------------------+

'' �ش�귣�� ��ü ���� �߼�(�ߺ� �߼� ����) ''//2014/03/31 �߰�
function isDlvFinishedByBrand(vOrderSerial,vMakerid)
    dim strSQL, targetCNT, DLVCNT
    targetCNT = 0
    DLVCNT    = 0

    strSQL = " select count(*) as targetCNT"
    strSQL = strSQL & " , sum(CASE WHEN d.currstate=7 and beasongdate is Not NULL THEN 1 ELSE 0 END) as DLVCNT" &VbCRLF
    strSQL = strSQL & " from [db_order].[dbo].tbl_order_master m" &VbCRLF
    strSQL = strSQL & " 	Join [db_order].[dbo].tbl_order_detail d" &VbCRLF
    strSQL = strSQL & " 	on m.orderserial=d.orderserial" &VbCRLF
    strSQL = strSQL & " where d.itemid<>0" &VbCRLF
    strSQL = strSQL & " and d.orderserial='"&vOrderSerial&"'" &VbCRLF
    strSQL = strSQL & " and d.makerid='"&vMakerid&"'" &VbCRLF
    strSQL = strSQL & " and d.cancelyn<>'Y'" &VbCRLF
    strSQL = strSQL & " and m.cancelyn='N'" &VbCRLF
    rsget.Open strSQL,dbget,1
    IF  not rsget.Eof  THEN
        targetCNT = rsget("targetCNT")
        DLVCNT    = rsget("DLVCNT")
    END IF
    rsget.CLOSE

    isDlvFinishedByBrand = false
    if (DLVCNT<1) or (targetCNT<>DLVCNT) then EXIT Function

    isDlvFinishedByBrand = true
end function

Function fcSendMailFinish_Dlv_Designer(vOrderSerial,vMakerid)	'/2011.04.21 �ѿ�� ����

	IF trim(vOrderSerial) ="" or vMakerid="" then EXIT Function

	dim strHTML_MAIN,strHTML_Sub ,strHTML_MAINother
	' ��� ��ü�� HTML
	strHTML_MAIN ="" &_
		"<tr>" &_
		"	<td style=""padding-bottom:5px;"">[$DELIVERY_HOST_IMG$]</td>" &_
		"</tr>" &_
		"<tr>" &_
		"	<td>" &_
		"		<!--�߼۵� ��ǰ ����Ʈ-->" &_
		"		<table width=""100%"" border=0 cellspacing=0 cellpadding=0 style=""border-top:3px solid #be0808;"">" &_
		"		<tr>" &_
		"			<td height=30 style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;"">" &_
		"				<table width=""100%"" border=0 cellspacing=0 cellpadding=0 style=""font-family:Dotum; font-size:11px; color:#888; padding-top:3px;"">" &_
		"				<tr align=""center"">" &_
		"				<td width=70>��ǰ</td>" &_
		"				<td width=60>��ǰ�ڵ�</td>" &_
		"				<td>��ǰ��(�ɼ�)</td>" &_
		"				<td width=40>����</td>" &_
		"				<td width=80>�ֹ�����</td>" &_
		"				<td width=120>�ù�����</td>" &_
		"				</tr>" &_
		"				</table>" &_
		"			</td>" &_
		"		</tr>" &_
		"		[$ITEMHTMLTABLE$]" &_
		"		</table>" &_
		"	</td>" &_
		"</tr>"

	strHTML_MAINother ="" &_
		"<tr>" &_
		"	<td style=""padding:35px 0 5px 0;"">[$DELIVERY_HOST_IMG$]</td>" &_
		"</tr>" &_
		"<tr>" &_
		"	<td>" &_
		"		<!--�߼۵� ��ǰ ����Ʈ-->" &_
		"		<table width=""100%"" border=0 cellspacing=0 cellpadding=0 style=""border-top:3px solid #be0808;"">" &_
		"		<tr>" &_
		"			<td height=30 style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;"">" &_
		"				<table width=""100%"" border=0 cellspacing=0 cellpadding=0 style=""font-family:Dotum; font-size:11px; color:#888; padding-top:3px;"">" &_
		"				<tr align=""center"">" &_
		"				<td width=70>��ǰ</td>" &_
		"				<td width=60>��ǰ�ڵ�</td>" &_
		"				<td>��ǰ��(�ɼ�)</td>" &_
		"				<td width=40>����</td>" &_
		"				<td width=80>�ֹ�����</td>" &_
		"				<td width=120>�ù�����</td>" &_
		"				</tr>" &_
		"				</table>" &_
		"			</td>" &_
		"		</tr>" &_
		"		[$ITEMHTMLTABLE$]" &_
		"		</table>" &_
		"	</td>" &_
		"</tr>"

	' �⺻ ��ǰ ����κ� HTML
	strHTML_Sub ="" &_
		"<tr>" &_
		"	<td height=80 style=""border-bottom:1px solid #eaeaea;"">" &_
		"		<table width=""100%"" border=0 cellspacing=0 cellpadding=0 style=""font-family:Dotum; font-size:11px; color:#888; padding-top:3px;"">" &_
		"		<tr align=""center"">" &_
		"			<td width=70>" &_
		"				<a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=[$ITEM_ID$]"" onFocus=""blur()""><img src=""[$ITEM_IMAGE_URL$]"" width=50 height=50 border=0></a></td>" &_
		"			<td width=60>" &_
		"				<a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=[$ITEM_ID$]"" style=""font-family:Dotum; font-size:11px; color:#888; text-decoration:none;"">[$ITEM_ID$]</a></td>" &_
		"			<td style=""text-align:left; line-height:16px; padding-left:5px;"">" &_
		"				<a href=""http://www.10x10.co.kr/street/street_brand.asp?makerid=[$ITEM_makerid$]"" style=""font-family: Verdana; font-size: 11px; color: #aaaaaa; text-decoration:none;"">" &_
		"				[[$ITEM_brandName$]]</a>" &_
		"				<br><a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=[$ITEM_ID$]"" style=""font-family:Dotum; font-size:11px; color:#888; text-decoration:none;"">" &_
		"				[$ITEM_NAME$]</a>" &_
		"				</td>" &_
		"			<td width=40>[$ITEM_QUANTITY$]</td>" &_
		"			<td width=80><span style=""color:#c01b1f; font-weight:bold;"">[$ITEM_DLV_STATUS$]</span></td>" &_
		"			<td width=120 style=""line-height:15px;"">[$ITEM_DELIVERY_LINK$]</td>" &_
		"		</tr>" &_
		"		</table>" &_
		"	</td>" &_
		"</tr>"

    '�ֹ� ��ǰ ����
	dim strSQL
	dim ITIMG , ITNM , ITID , ITOPNM , ITNO , ITbrandName ,ITmakerid
	dim DLVSTS, DLVLKTXT
	dim tmpHTML,NowHTML,OtherHTML,ITTITLEIMG
	dim isNowDLV,isOtherDLV '���� ���,�����ֹ��� ��ǰ

	tmpHTML="":NowHTML="":OtherHTML=""

	strSQL =" SELECT a.itemid, a.itemoptionname, c.smallimage, c.itemname,c.makerid ," &_
			" (c.cate_large + c.cate_mid + c.cate_small) as itemserial," &_
			" a.itemcost as sellcash, a.itemno, a.isupchebeasong, a.songjangdiv, replace(isnull(a.songjangno,''),'-','') as songjangno, a.currstate" &_
			" ,s.divname,s.findurl ,c.brandName" &_
			" FROM [db_order].[dbo].tbl_order_detail a" &_
			" JOIN [db_item].[dbo].tbl_item c" &_
			" 	on c.itemid = a.itemid" &_
			" LEFT JOIN db_order.[dbo].tbl_songjang_div s" &_
			" 	on a.songjangdiv=s.divcd" &_
			" WHERE a.orderserial = '" & vOrderSerial & "'" &_
			" and a.itemid <> '0'" &_
			" and (a.cancelyn<>'Y')"

	'response.write strSQL
	rsget.Open strSQL,dbget,1

	IF  not rsget.Eof  THEN
		rsget.Movefirst

		DO UNTIL rsget.eof

			'-- �귣��
			ITmakerid = db2html(rsget("makerid"))

			'-- �귣���
			ITbrandName = db2html(rsget("brandName"))

			'--- ��ǰ�̹���
			ITIMG = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallimage")
			' ��ǰ �ڵ�
			ITID = rsget("itemid")
			'--- ��ǰ��
			ITNM = db2html(rsget("itemname"))
			'--- ��ǰ�ɼǸ�
			ITOPNM = db2html(rsget("itemoptionname"))

			IF ITOPNM<>"" then
				ITNM = ITNM & "<br>[<span style=""color:#1545f9"">" & ITOPNM & "</span>]<br>"
			END IF
			'--- ��ǰ���� -- ������ style
			ITNO = Cstr(rsget("itemno"))
			IF rsget("itemno")>1 THEN
				ITNO = "<strong>" & Cstr(rsget("itemno")) & "</strong>"
			END IF

			'--- ��ۻ��� ����
				IF rsget("currstate") = 7 THEN
					 DLVSTS = "<span class=""black12px"">���Ϸ�</span>"
				 ELSE
					 DLVSTS = "��ǰ�غ���"
				 END IF
			'--- �ù�/���� ����
			IF ((Not isnull(rsget("songjangno"))) and  (rsget("songjangno")<>"") ) THEN

				DLVLKTXT ="<strong>" & db2html(rsget("divname")) & "</strong><br><a href=""" & db2html(rsget("findurl")) & rsget("songjangno") & """ style=""font-family:Dotum; font-size:11px; color:#888; text-decoration:none;"">" & rsget("songjangno") & "</a>"
			else
				DLVLKTXT ="-"
			end if
			tmpHTML = strHTML_Sub
			tmpHTML = replace(tmpHTML,"[$ITEM_makerid$]",ITmakerid)
			tmpHTML = replace(tmpHTML,"[$ITEM_brandName$]",ITbrandName)
			tmpHTML = replace(tmpHTML,"[$ITEM_IMAGE_URL$]",ITIMG)
			tmpHTML = replace(tmpHTML,"[$ITEM_ID$]",ITID)
			tmpHTML = replace(tmpHTML,"[$ITEM_NAME$]",ITNM)
			tmpHTML = replace(tmpHTML,"[$ITEM_QUANTITY$]",ITNO)
			tmpHTML = replace(tmpHTML,"[$ITEM_DLV_STATUS$]",DLVSTS)
			tmpHTML = replace(tmpHTML,"[$ITEM_DELIVERY_LINK$]",DLVLKTXT)

			IF rsget("isupchebeasong") = "Y" and rsget("makerid")=vMakerid and rsget("songjangno")<>"" THEN
				NowHTML= NowHTML & tmpHTML
				isNowDLV= true
			ELSE
				OtherHTML = OtherHTML & tmpHTML
				isOtherDLV= true
			END IF

			tmpHTML ="":ITIMG="":ITID="":ITNM="":ITOPNM="":ITNO="":DLVSTS="":DLVLKTXT=""

			rsget.movenext
		LOOP
    ELSE
    	rsget.close
		EXIT FUNCTION

    END IF
    rsget.close

	IF NowHTML<>"" and isNowDLV THEN
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2011/mail/tit_shiped.gif"" alt=""���� ��ǰ�� �������"">"
		NowHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",NowHTML)
		NowHTML = replace(NowHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		NowHTML= ""
	END IF

	IF OtherHTML<>"" and isOtherDLV THEN
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2011/mail/tit_otherpd.gif"" alt="" ���� �ֹ��Ͻ� ��ǰ �����Ȳ"">"
		OtherHTML = replace(strHTML_MAINother,"[$ITEMHTMLTABLE$]",OtherHTML)
		OtherHTML = replace(OtherHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		OtherHTML=""
	END IF


	'//=======  �������� & ������� , �������� �ҷ����� =========/
	call getInfo(vOrderSerial)

	IF MailTo ="" Then
		Exit Function
	End IF

	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	oMail.MailType		 = 8 '���� ������ ������ (mailLib2.asp ����)
	oMail.MailTitles	 = "[�ٹ�����]�ֹ��Ͻ� ��ǰ�� ���� �ٹ����� ��۾ȳ��Դϴ�!"
	'oMail.SenderNm		 = "�ٹ�����"
	'oMail.SenderMail	 = "customer@10x10.co.kr"
	oMail.AddrType		 = "string"
	oMail.ReceiverNm	 = MailTo_Nm
	oMail.ReceiverMail	 = MailTo

	MailHTML= oMail.getMailTemplate()

	IF MailHTML="" Then
		SET oMail = nothing
		response.write "<script>alert('���Ϲ߼��� ���� �Ͽ����ϴ�.');</script>"
		Exit Function
    End IF

	'// ���� ���Ͽ� ���� ġȯ
	MailHTML = replace(MailHTML,"[$USER_NAME$]", MailTo_Nm) ' �ֹ��� �̸�
	MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' �ֹ���ȣ
	MailHTML = replace(MailHTML,"[$$DELIVERY_ITEM_INFO$$]",NowHTML) '���� ��ǰ HTML
	MailHTML = replace(MailHTML,"[$$DELIVERY_OTHER_ITEM_INFO$$]",OtherHTML)	'���� �ֹ��ѻ�ǰ HTML
	MailHTML = replace(MailHTML,"[$$REQ_INFO_HTML$$]",ReqInfoHTML)	'����� ���� HTML

	oMail.MailConts = MailHTML

	'response.write MailHTML
	'response.end
	oMail.MailerMailGubun = 4		' ���Ϸ� �ڵ����� ��ȣ
	oMail.Send_TMSMailer()		'TMS���Ϸ�
	'oMail.Send_Mailer()
	''oMail.Send_CDO()
	'oMail.Send_CDONT()

	SET oMail = nothing

End Function

Function fcSendMailFinish_Dlv_Designer_off(vmasteridx,vMakerid)
	IF trim(vmasteridx) ="" or vMakerid="" then EXIT Function

	dim strHTML_MAIN,strHTML_Sub

	dim vOrderSerial

	' ��� ��ü�� HTML
	strHTML_MAIN ="" &_
		"<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" &_
		"<tr>" &_
		"	<td style=""padding-bottom:7px;"">[$DELIVERY_HOST_IMG$]</td>" &_
		"</tr>" &_
		"<tr>" &_
		"	<td>" &_
		"		<table width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""0"" style=""border-bottom:1px solid #dddddd"">" &_
		"		[$ITEMHTMLTABLE$]" &_
		"		</table>" &_
		"	</td>" &_
		"</tr>" &_
		"</table>"

	' �⺻ ��ǰ ����κ� HTML '�̹���"<td><img src=""[$ITEM_IMAGE_URL$]"" width=""50"" height=""50""></td>" &_
	strHTML_Sub ="" &_
			"<tr>" &_
			"	<td>" &_
			"		<table width=""548"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-top:1px solid #dddddd"">" &_
			"		<tr>" &_
			"			<td width=""260"" align=""right"" style=""border-right: 1px solid #dddddd"">" &_
			"				<table width=""255"" height=""50""  border=""0"" cellpadding=""0"" cellspacing=""0"">" &_
			"				<tr>" &_
			"					<td width=""50"" valign=""bottom"">" &_
			"						<table width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""0"">" &_
			"						<tr>" &_

			"						</tr>" &_
			"						</table>" &_
			"					</td>" &_
			"					<td  style=""padding:5"">[$ITEM_ID$]<br>[$ITEM_NAME$] </td>" &_
			"				</tr>" &_
			"				</table>" &_
			"			</td>" &_
			"			<td align=""center"">" &_
			"				<table width=""100%"" height=""70""  border=""0"" cellpadding=""0"" cellspacing=""0"" bgcolor=""#eeeeee"">" &_
			"				<tr>" &_
			"					<td width=""60"" height=""35"" align=""center"">�� ��</td>" &_
			"					<td width=""60"" style=""padding:0 5 0 5;"" bgcolor=""#FFFFFF"">[$ITEM_QUANTITY$]</td>" &_
			"					<td width=""60"" align=""center"" style=""padding:0 5 0 5;"">�����Ȳ</td>" &_
			"					<td class=""black12px"" style=""padding:0 5 0 5;"" bgcolor=""#FFFFFF""> [$ITEM_DLV_STATUS$]</td>" &_
			"				</tr>" &_
			"				<tr height=""1"">" &_
			"					<td colspan=""4"" align=""center"" bgcolor=""#dddddd""></td>" &_
			"				</tr>" &_
			"				<tr>" &_
			"					<td align=""center"">�����</td>" &_
			"					<td colspan=""3"" style=""padding:5"" bgcolor=""#FFFFFF""><strong class=""Information_font"">[$ITEM_DELIVERY_LINK$]</strong></td>" &_
			"				</tr>" &_
			"				</table>" &_
			"			</td>" &_
			"		</tr>" &_
			"		</table>" &_
			"	</td>" &_
			"</tr>"

    '�ֹ� ��ǰ ����
	dim strSQL, ITIMG , ITNM , ITID , ITOPNM , ITNO ,DLVSTS, DLVLKTXT
	dim tmpHTML,NowHTML,OtherHTML,ITTITLEIMG
	dim isNowDLV,isOtherDLV '���� ���,�����ֹ��� ��ǰ

	tmpHTML="":NowHTML="":OtherHTML=""

	strSQL =" SELECT" &_
			" d.itemid, d.itemgubun,d.itemoption,d.makerid, d.itemno, d.isupchebeasong" &_
			" ,replace(isnull(d.songjangno,''),'-','') as songjangno, d.currstate, d.songjangdiv" &_
			" ,od.sellprice as sellcash,od.itemoptionname, od.itemname" &_
			" ,s.divname,s.findurl, m.orderno " &_
			" from db_shop.dbo.tbl_shopbeasong_order_master m" &_
			" join db_shop.dbo.tbl_shopbeasong_order_detail d" &_
			" on m.masteridx=d.masteridx" &_
			" left join [db_shop].[dbo].tbl_shopjumun_detail od" &_
			" on d.orgdetailidx = od.idx" &_
			" LEFT JOIN db_order.[dbo].tbl_songjang_div s" &_
			" 	on d.songjangdiv=s.divcd" &_
			" WHERE d.masteridx = " & vmasteridx & "" &_
			" and d.itemid <> '0'" &_
			" and (d.cancelyn<>'Y')"

	'response.write strSQL &"<br>"
	rsget.Open strSQL,dbget,1
	IF  not rsget.Eof  THEN
		rsget.Movefirst

		vOrderSerial = rsget("orderno")

		DO UNTIL rsget.eof

			'--- ��ǰ�̹���
			ITIMG = ""
			' ��ǰ �ڵ�
			ITID = rsget("itemgubun")&Format00(6,rsget("itemid"))&rsget("itemoption")
			'--- ��ǰ��
			ITNM = db2html(rsget("itemname"))
			'--- ��ǰ�ɼǸ�
			ITOPNM = db2html(rsget("itemoptionname"))

			IF ITOPNM<>"" then
				ITNM = ITNM & "<br><font color=""blue"">[" & ITOPNM & "]</font>"
			END IF
			'--- ��ǰ���� -- ������ style
			ITNO = Cstr(rsget("itemno"))
			IF rsget("itemno")>1 THEN
				ITNO = "<strong>" & Cstr(rsget("itemno")) & "</strong>"
			END IF

			'--- ��ۻ��� ����
				IF rsget("currstate") = 7 THEN
					 DLVSTS = "<span class=""black12px"">���Ϸ�</span>"
				 ELSE
					 DLVSTS = "��ǰ�غ���"
				 END IF
			'--- �ù�/���� ����
			IF ((Not isnull(rsget("songjangno"))) and  (rsget("songjangno")<>"") ) THEN
				DLVLKTXT ="<a href=""" & db2html(rsget("findurl")) & rsget("songjangno") & """ target=""_blank""  class=""link_title"">" & db2html(rsget("divname")) & " " & rsget("songjangno") & "</a>"
			else
				DLVLKTXT ="-"
			end if
			tmpHTML = strHTML_Sub
			'tmpHTML = replace(tmpHTML,"[$ITEM_IMAGE_URL$]",ITIMG)
			tmpHTML = replace(tmpHTML,"[$ITEM_ID$]",ITID)
			tmpHTML = replace(tmpHTML,"[$ITEM_NAME$]",ITNM)
			tmpHTML = replace(tmpHTML,"[$ITEM_QUANTITY$]",ITNO)
			tmpHTML = replace(tmpHTML,"[$ITEM_DLV_STATUS$]",DLVSTS)
			tmpHTML = replace(tmpHTML,"[$ITEM_DELIVERY_LINK$]",DLVLKTXT)

			IF rsget("isupchebeasong") = "Y" and rsget("makerid")=vMakerid and rsget("songjangno")<>"" THEN
				NowHTML= NowHTML & tmpHTML
				isNowDLV= true
			ELSE
				OtherHTML = OtherHTML & tmpHTML
				isOtherDLV= true
			END IF

			tmpHTML ="":ITIMG="":ITID="":ITNM="":ITOPNM="":ITNO="":DLVSTS="":DLVLKTXT=""

			rsget.movenext
		LOOP
    ELSE

    	rsget.close
		EXIT FUNCTION

    END IF
    rsget.close

	IF NowHTML<>"" and isNowDLV THEN
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2008/mail/a03_text01.gif"" width=""79"" height=""18"" alt=""���� ��ǰ�� �������"">"
		NowHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",NowHTML)
		NowHTML = replace(NowHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		NowHTML= ""
	END IF

	IF OtherHTML<>"" and isOtherDLV THEN
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2008/mail/a03_text02.gif"" width=""193"" height=""18"" alt="" ���� �ֹ��Ͻ� ��ǰ �����Ȳ"">"
		OtherHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",OtherHTML)
		OtherHTML = replace(OtherHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		OtherHTML=""
	END IF

	'//=======  �������� & ������� , �������� �ҷ����� =========/
	call getInfo_off(vmasteridx)

	IF MailTo ="" Then
		Exit Function
	End IF

	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	oMail.MailType		 = 8 '���� ������ ������ (mailLib2.asp ����)
	oMail.MailTitles	 = "[�ٹ����ټ�]�ֹ��Ͻ� ��ǰ�� ���� �ٹ����� ��۾ȳ��Դϴ�!"
	'oMail.SenderNm		 = "�ٹ�����"
	'oMail.SenderMail	 = "customer@10x10.co.kr"
	oMail.AddrType		 = "string"
	oMail.ReceiverNm	 = MailTo_Nm
	oMail.ReceiverMail	 = MailTo

	MailHTML= oMail.getMailTemplate()

	IF MailHTML="" Then
		SET oMail = nothing
		response.write "<script>alert('���Ϲ߼��� ���� �Ͽ����ϴ�.');</script>"
		Exit Function
    End IF

	'// ���� ���Ͽ� ���� ġȯ
	MailHTML = replace(MailHTML,"[$USER_NAME$]", MailTo_Nm) ' �ֹ��� �̸�
	MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' �ֹ���ȣ
	MailHTML = replace(MailHTML,"[$$DELIVERY_ITEM_INFO$$]",NowHTML) '���� ��ǰ HTML
	MailHTML = replace(MailHTML,"[$$DELIVERY_OTHER_ITEM_INFO$$]",OtherHTML)	'���� �ֹ��ѻ�ǰ HTML
	MailHTML = replace(MailHTML,"[$$REQ_INFO_HTML$$]",ReqInfoHTML)	'����� ���� HTML

	oMail.MailConts = MailHTML
	oMail.MailerMailGubun = 4		' ���Ϸ� �ڵ����� ��ȣ
	oMail.Send_TMSMailer()		'TMS���Ϸ�
	'oMail.Send_Mailer()
	'oMail.Send_CDO()
	'oMail.Send_CDONT()

	SET oMail = nothing
End Function

%>
