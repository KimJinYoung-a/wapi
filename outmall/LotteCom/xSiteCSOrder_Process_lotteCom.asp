<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 CS 주문처리
'###########################################################

%>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteCSOrderCls.asp"-->
<!-- include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/outmall/LotteCom/inc_dailyAuthCheck.asp" -->
<!-- #include virtual="/outmall/LotteCom/lotteitemcls.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<%

'// 2014-08-27, skyer9
''Server.ScriptTimeout = 60

if application("Svr_Info")="Dev" then
	lotteAPIURL = "http://openapi.lotte.com"
	lotteAuthNo = "afc92a6024a23c9ae7c6e8fa3647c9fc0de8384e2b7798af0961e8a127d30516efd5a556fd6008b89630b3cf2b40b09b7e4a7a5f1ebd67a6d29446a381ed803c"
end if

'' response.write lotteAuthNo
'' response.end

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, buf
dim i, j, k

'response.write lotteAuthNo
'response.end

'' -- 교환/반품
'' http://openapi.lotte.com/openapi/searchReturnList.lotte?subscriptionId=[subscriptionId]&start_date=20130415&end_date=20130416&ord_dtl_stat_cd=20

'' -- 취소
'' http://openapi.lotte.com/openapi/searchCnclList.lotte?subscriptionId=[subscriptionId]&start_date=20130416&end_date=20130416&selCval=17

'' -- 신규주문
'' http://openapi.lotte.com/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=[subscriptionId]&start_date=20130416&end_date=20130416&SelOption=01

'' -- 발주확인주문
'' http://openapi.lotte.com/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=[subscriptionId]&start_date=20130416&end_date=20130416&SelOption=02


'// ============================================================================
'// [divcd]
'// ============================================================================
'A008			주문취소
'
'A004			반품접수(업체배송)
'A010			회수신청(텐바이텐배송)
'
'A001			누락재발송
'A002			서비스발송
'
'A000			맞교환출고
'A100			상품변경 맞교환출고
'
'A009			기타사항
'A006			출고시유의사항
'A700			업체기타정산
'
'A003			환불
'A005			외부몰환불요청
'A007			카드,이체,휴대폰취소요청
'
'A011			맞교환회수(텐바이텐배송)
'A012			맞교환반품(업체배송)

'A111			상품변경 맞교환회수(텐바이텐배송)
'A112			상품변경 맞교환반품(업체배송)
'// ============================================================================

dim mode
dim sellsite
dim reguserid
Dim AssignedRow
Dim ErrMsg

dim resultCount

dim divcd, yyyymmdd, idx, finUserid

mode = requestCheckVar(html2db(request("mode")),32)
sellsite = requestCheckVar(html2db(request("sellsite")),32)
idx = requestCheckVar(html2db(request("idx")),32)
finUserid = session("ssBctID")
If finUserid = "" Then
	finUserid = "system"
End If

dim oCxSiteCSOrderXML
Set oCxSiteCSOrderXML = new CxSiteCSOrderXML

dim getDivCD
if (mode = "getxsitecslist") then
    IF (sellsite="lotteCom") then
    	ErrMsg = ""

		getDivCD = Trim(application("xSiteGetLotteComCS_DIVCD"))
		if (getDivCD = "") then
			getDivCD = "A008"
		elseif (getDivCD = "A004") then
			getDivCD = "A008"
		else
			getDivCD = "A004"
		end if

		'// ========================================================================
		'// 취소 or 반품
		oCxSiteCSOrderXML.FRectSellSite = sellsite
		oCxSiteCSOrderXML.FRectDivCD = getDivCD
		if Hour(Now()) < 6 then
			'// 새벽시간은 어제내역
			oCxSiteCSOrderXML.FRectStartYYYYMMDD = Left(DateAdd("d", -1, now), 10)				'// 2013-01-01
			oCxSiteCSOrderXML.FRectEndYYYYMMDD = Left(DateAdd("d", -1, now), 10)
		else
			oCxSiteCSOrderXML.FRectStartYYYYMMDD = Left(now, 10)
			oCxSiteCSOrderXML.FRectEndYYYYMMDD = Left(now, 10)
		end if

		Call oCxSiteCSOrderXML.SavexSiteCSOrderListtoDB

		Call oCxSiteCSOrderXML.ResetXML()

		if (getDivCD = "A004") then
			'// 반품완료내역
			oCxSiteCSOrderXML.FRectDivCD = "A004_B007"
			Call oCxSiteCSOrderXML.SavexSiteCSOrderListtoDB
			Call oCxSiteCSOrderXML.ResetXML()
		end if

		if (getDivCD <> Trim(application("xSiteGetLotteComCS_DIVCD"))) then
			application("xSiteGetLotteComCS_DIVCD") = getDivCD
		end if
    else
        rw "미지정 sellsite:"&sellsite
        dbget.Close : response.end
    end if
elseif (mode = "setfinish") then

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPCS "
	sqlStr = sqlStr + " set currstate = 'B007' "
	sqlStr = sqlStr + " ,finUserId = '"&finUserId&"' "			'2015-07-06 김진영 추가
	sqlStr = sqlStr + " where idx = " + CStr(idx) + " and currstate = 'B001' "
	''rw strSql
	dbget.Execute(sqlStr)
elseif (mode = "delfinish") then

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPCS "
	sqlStr = sqlStr + " set currstate = 'B001' "
	sqlStr = sqlStr + " where idx = " + CStr(idx) + " and currstate = 'B007' "
	''rw strSql
	dbget.Execute(sqlStr)

else

end if

%>

<% if  (IsAutoScript) then  %>
<% rw "OK" %>
<% else %>
<script>alert('저장되었습니다.(<%= application("xSiteGetLotteComCS_DIVCD") %>)');</script>
<script>window.close();</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
