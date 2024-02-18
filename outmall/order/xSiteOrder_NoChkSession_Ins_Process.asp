<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 240 ''초단위
%>
<%
'###########################################################
' Description : 제휴몰 API 주문입력
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/auction/auctionItemcls.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%

dim IS_TEST_MODE : IS_TEST_MODE = False

Dim sqlStr, sellsite, selldate, selldateStr, mode
dim isSuccess
dim i, j, k
dim orderObjArr, tmpObjArr
dim nowdate, fromdate, todate, currdate

sellsite	= requestCheckVar(html2db(request("sellsite")),32)
selldate	= requestCheckVar(html2db(request("selldate")),32)
mode		= requestCheckVar(html2db(request("mode")),32)

dim IS_SELLDATE_FIXED : IS_SELLDATE_FIXED = False
if (selldate = "") then
	'// 오늘까지 일괄로 가져오기
	Call GetCheckStatus(sellsite, selldate, isSuccess)
	fromdate = selldate
	todate = Left(Now, 10)
	if (fromdate = todate) then
		selldateStr = fromdate
	else
		selldateStr = fromdate & " ~ " & todate
	end if
else
	fromdate = selldate
	todate = selldate
	selldateStr = fromdate
	IS_SELLDATE_FIXED = True
end if

select case sellsite
	case "gseshopNew"
		currdate = fromdate
		do while (currdate <= todate)
			response.write "gseshop : " & currdate & "<br />"
			Call GetOrderFromExtSite(sellsite, currdate)
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop

		''하루씩만 하자.
		'response.write "<script>var popwin = window.open('http://ecb2b.gsshop.com/SupSendOrderInfo.gs?supCd=1003890&sdDt="&replace(currdate,"-","")&"&tnsType=S','popGsOrdReceiv','width=300,height=300,scrollbars=yes,resizable=yes');popwin.focus();</script>"
		'sellsite = "gseshop"
	case else
		response.write "잘못된 접근입니다."
		dbget.close : response.end
end select

if (IS_TEST_MODE = False) and (IS_SELLDATE_FIXED = False) then
	if (selldate < Left(Now(), 10)) then
		Call SetCheckStatus(sellsite, Left(DateAdd("d", 1, CDate(selldate)), 10), "N")
	elseif (selldate = Left(Now(), 10)) then
		Call SetCheckStatus(sellsite, selldate, "Y")
	end if
end if

''품절/가격 오류체크
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
<% rw "OK" %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
