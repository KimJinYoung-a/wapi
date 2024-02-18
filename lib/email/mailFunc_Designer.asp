<!-- #include virtual="/lib/email/mailFunction.asp" -->
<%
'+--------------------------------------------------------------------------------------------------------------------------------+
'|                                        업체 배송 상품 메일 발송                                                                |
'+----------------------------------------------------+---------------------------------------------------------------------------+
'|             함 수 명                               |                          기    능                                         |
'+----------------------------------------------------+---------------------------------------------------------------------------+
'| fcSendMailFinish_Dlv_Designer(orderserial,makerid) | 출고 메일 발송(업체배송 출고)                                             |
'|                                                    | 사용예 : fcSendMailFinish_Dlv_Designer('012012304','1293495006')          |
'+----------------------------------------------------+---------------------------------------------------------------------------+
'| fcSendMailFinish_Dlv_Designer_off(detailidx,makerid)  오프라인 출고 메일 발송(업체배송 출고)                                   |
'|                                                    | 사용예 : fcSendMailFinish_Dlv_Designer_off('012012304','1293495006')      |
'+----------------------------------------------------+---------------------------------------------------------------------------+

'' 해당브랜드 전체 출고시 발송(중복 발송 제거) ''//2014/03/31 추가
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

Function fcSendMailFinish_Dlv_Designer(vOrderSerial,vMakerid)	'/2011.04.21 한용민 수정

	IF trim(vOrderSerial) ="" or vMakerid="" then EXIT Function

	dim strHTML_MAIN,strHTML_Sub ,strHTML_MAINother
	' 배송 주체별 HTML
	strHTML_MAIN ="" &_
		"<tr>" &_
		"	<td style=""padding-bottom:5px;"">[$DELIVERY_HOST_IMG$]</td>" &_
		"</tr>" &_
		"<tr>" &_
		"	<td>" &_
		"		<!--발송된 상품 리스트-->" &_
		"		<table width=""100%"" border=0 cellspacing=0 cellpadding=0 style=""border-top:3px solid #be0808;"">" &_
		"		<tr>" &_
		"			<td height=30 style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;"">" &_
		"				<table width=""100%"" border=0 cellspacing=0 cellpadding=0 style=""font-family:Dotum; font-size:11px; color:#888; padding-top:3px;"">" &_
		"				<tr align=""center"">" &_
		"				<td width=70>상품</td>" &_
		"				<td width=60>상품코드</td>" &_
		"				<td>상품명(옵션)</td>" &_
		"				<td width=40>수량</td>" &_
		"				<td width=80>주문상태</td>" &_
		"				<td width=120>택배정보</td>" &_
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
		"		<!--발송된 상품 리스트-->" &_
		"		<table width=""100%"" border=0 cellspacing=0 cellpadding=0 style=""border-top:3px solid #be0808;"">" &_
		"		<tr>" &_
		"			<td height=30 style=""background:#fcf6f6; border-bottom:1px solid #eaeaea;"">" &_
		"				<table width=""100%"" border=0 cellspacing=0 cellpadding=0 style=""font-family:Dotum; font-size:11px; color:#888; padding-top:3px;"">" &_
		"				<tr align=""center"">" &_
		"				<td width=70>상품</td>" &_
		"				<td width=60>상품코드</td>" &_
		"				<td>상품명(옵션)</td>" &_
		"				<td width=40>수량</td>" &_
		"				<td width=80>주문상태</td>" &_
		"				<td width=120>택배정보</td>" &_
		"				</tr>" &_
		"				</table>" &_
		"			</td>" &_
		"		</tr>" &_
		"		[$ITEMHTMLTABLE$]" &_
		"		</table>" &_
		"	</td>" &_
		"</tr>"

	' 기본 상품 설명부분 HTML
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

    '주문 상품 정보
	dim strSQL
	dim ITIMG , ITNM , ITID , ITOPNM , ITNO , ITbrandName ,ITmakerid
	dim DLVSTS, DLVLKTXT
	dim tmpHTML,NowHTML,OtherHTML,ITTITLEIMG
	dim isNowDLV,isOtherDLV '지금 배송,같이주문한 상품

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

			'-- 브랜드
			ITmakerid = db2html(rsget("makerid"))

			'-- 브랜드명
			ITbrandName = db2html(rsget("brandName"))

			'--- 상품이미지
			ITIMG = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallimage")
			' 상품 코드
			ITID = rsget("itemid")
			'--- 상품명
			ITNM = db2html(rsget("itemname"))
			'--- 상품옵션명
			ITOPNM = db2html(rsget("itemoptionname"))

			IF ITOPNM<>"" then
				ITNM = ITNM & "<br>[<span style=""color:#1545f9"">" & ITOPNM & "</span>]<br>"
			END IF
			'--- 상품수량 -- 수량별 style
			ITNO = Cstr(rsget("itemno"))
			IF rsget("itemno")>1 THEN
				ITNO = "<strong>" & Cstr(rsget("itemno")) & "</strong>"
			END IF

			'--- 배송상태 지정
				IF rsget("currstate") = 7 THEN
					 DLVSTS = "<span class=""black12px"">출고완료</span>"
				 ELSE
					 DLVSTS = "상품준비중"
				 END IF
			'--- 택배/송장 설정
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
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2011/mail/tit_shiped.gif"" alt=""출고된 상품의 배송정보"">"
		NowHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",NowHTML)
		NowHTML = replace(NowHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		NowHTML= ""
	END IF

	IF OtherHTML<>"" and isOtherDLV THEN
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2011/mail/tit_otherpd.gif"" alt="" 같이 주문하신 상품 배송현황"">"
		OtherHTML = replace(strHTML_MAINother,"[$ITEMHTMLTABLE$]",OtherHTML)
		OtherHTML = replace(OtherHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		OtherHTML=""
	END IF


	'//=======  메일정보 & 배송정보 , 결제정보 불러오기 =========/
	call getInfo(vOrderSerial)

	IF MailTo ="" Then
		Exit Function
	End IF

	'//=======  메일 발송 =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	oMail.MailType		 = 8 '메일 종류별 고정값 (mailLib2.asp 참고)
	oMail.MailTitles	 = "[텐바이텐]주문하신 상품에 대한 텐바이텐 배송안내입니다!"
	'oMail.SenderNm		 = "텐바이텐"
	'oMail.SenderMail	 = "customer@10x10.co.kr"
	oMail.AddrType		 = "string"
	oMail.ReceiverNm	 = MailTo_Nm
	oMail.ReceiverMail	 = MailTo

	MailHTML= oMail.getMailTemplate()

	IF MailHTML="" Then
		SET oMail = nothing
		response.write "<script>alert('메일발송이 실패 하였습니다.');</script>"
		Exit Function
    End IF

	'// 실제 메일에 정보 치환
	MailHTML = replace(MailHTML,"[$USER_NAME$]", MailTo_Nm) ' 주문자 이름
	MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' 주문번호
	MailHTML = replace(MailHTML,"[$$DELIVERY_ITEM_INFO$$]",NowHTML) '출고된 상품 HTML
	MailHTML = replace(MailHTML,"[$$DELIVERY_OTHER_ITEM_INFO$$]",OtherHTML)	'같이 주문한상품 HTML
	MailHTML = replace(MailHTML,"[$$REQ_INFO_HTML$$]",ReqInfoHTML)	'배송지 정보 HTML

	oMail.MailConts = MailHTML

	'response.write MailHTML
	'response.end
	oMail.MailerMailGubun = 4		' 메일러 자동메일 번호
	oMail.Send_TMSMailer()		'TMS메일러
	'oMail.Send_Mailer()
	''oMail.Send_CDO()
	'oMail.Send_CDONT()

	SET oMail = nothing

End Function

Function fcSendMailFinish_Dlv_Designer_off(vmasteridx,vMakerid)
	IF trim(vmasteridx) ="" or vMakerid="" then EXIT Function

	dim strHTML_MAIN,strHTML_Sub

	dim vOrderSerial

	' 배송 주체별 HTML
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

	' 기본 상품 설명부분 HTML '이미지"<td><img src=""[$ITEM_IMAGE_URL$]"" width=""50"" height=""50""></td>" &_
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
			"					<td width=""60"" height=""35"" align=""center"">수 량</td>" &_
			"					<td width=""60"" style=""padding:0 5 0 5;"" bgcolor=""#FFFFFF"">[$ITEM_QUANTITY$]</td>" &_
			"					<td width=""60"" align=""center"" style=""padding:0 5 0 5;"">배송현황</td>" &_
			"					<td class=""black12px"" style=""padding:0 5 0 5;"" bgcolor=""#FFFFFF""> [$ITEM_DLV_STATUS$]</td>" &_
			"				</tr>" &_
			"				<tr height=""1"">" &_
			"					<td colspan=""4"" align=""center"" bgcolor=""#dddddd""></td>" &_
			"				</tr>" &_
			"				<tr>" &_
			"					<td align=""center"">운송장</td>" &_
			"					<td colspan=""3"" style=""padding:5"" bgcolor=""#FFFFFF""><strong class=""Information_font"">[$ITEM_DELIVERY_LINK$]</strong></td>" &_
			"				</tr>" &_
			"				</table>" &_
			"			</td>" &_
			"		</tr>" &_
			"		</table>" &_
			"	</td>" &_
			"</tr>"

    '주문 상품 정보
	dim strSQL, ITIMG , ITNM , ITID , ITOPNM , ITNO ,DLVSTS, DLVLKTXT
	dim tmpHTML,NowHTML,OtherHTML,ITTITLEIMG
	dim isNowDLV,isOtherDLV '지금 배송,같이주문한 상품

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

			'--- 상품이미지
			ITIMG = ""
			' 상품 코드
			ITID = rsget("itemgubun")&Format00(6,rsget("itemid"))&rsget("itemoption")
			'--- 상품명
			ITNM = db2html(rsget("itemname"))
			'--- 상품옵션명
			ITOPNM = db2html(rsget("itemoptionname"))

			IF ITOPNM<>"" then
				ITNM = ITNM & "<br><font color=""blue"">[" & ITOPNM & "]</font>"
			END IF
			'--- 상품수량 -- 수량별 style
			ITNO = Cstr(rsget("itemno"))
			IF rsget("itemno")>1 THEN
				ITNO = "<strong>" & Cstr(rsget("itemno")) & "</strong>"
			END IF

			'--- 배송상태 지정
				IF rsget("currstate") = 7 THEN
					 DLVSTS = "<span class=""black12px"">출고완료</span>"
				 ELSE
					 DLVSTS = "상품준비중"
				 END IF
			'--- 택배/송장 설정
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
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2008/mail/a03_text01.gif"" width=""79"" height=""18"" alt=""출고된 상품의 배송정보"">"
		NowHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",NowHTML)
		NowHTML = replace(NowHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		NowHTML= ""
	END IF

	IF OtherHTML<>"" and isOtherDLV THEN
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2008/mail/a03_text02.gif"" width=""193"" height=""18"" alt="" 같이 주문하신 상품 배송현황"">"
		OtherHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",OtherHTML)
		OtherHTML = replace(OtherHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		OtherHTML=""
	END IF

	'//=======  메일정보 & 배송정보 , 결제정보 불러오기 =========/
	call getInfo_off(vmasteridx)

	IF MailTo ="" Then
		Exit Function
	End IF

	'//=======  메일 발송 =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	oMail.MailType		 = 8 '메일 종류별 고정값 (mailLib2.asp 참고)
	oMail.MailTitles	 = "[텐바이텐샵]주문하신 상품에 대한 텐바이텐 배송안내입니다!"
	'oMail.SenderNm		 = "텐바이텐"
	'oMail.SenderMail	 = "customer@10x10.co.kr"
	oMail.AddrType		 = "string"
	oMail.ReceiverNm	 = MailTo_Nm
	oMail.ReceiverMail	 = MailTo

	MailHTML= oMail.getMailTemplate()

	IF MailHTML="" Then
		SET oMail = nothing
		response.write "<script>alert('메일발송이 실패 하였습니다.');</script>"
		Exit Function
    End IF

	'// 실제 메일에 정보 치환
	MailHTML = replace(MailHTML,"[$USER_NAME$]", MailTo_Nm) ' 주문자 이름
	MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' 주문번호
	MailHTML = replace(MailHTML,"[$$DELIVERY_ITEM_INFO$$]",NowHTML) '출고된 상품 HTML
	MailHTML = replace(MailHTML,"[$$DELIVERY_OTHER_ITEM_INFO$$]",OtherHTML)	'같이 주문한상품 HTML
	MailHTML = replace(MailHTML,"[$$REQ_INFO_HTML$$]",ReqInfoHTML)	'배송지 정보 HTML

	oMail.MailConts = MailHTML
	oMail.MailerMailGubun = 4		' 메일러 자동메일 번호
	oMail.Send_TMSMailer()		'TMS메일러
	'oMail.Send_Mailer()
	'oMail.Send_CDO()
	'oMail.Send_CDONT()

	SET oMail = nothing
End Function

%>
