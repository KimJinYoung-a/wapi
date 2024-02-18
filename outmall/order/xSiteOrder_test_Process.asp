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
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteOrdertest.asp"-->
<!-- #include virtual="/outmall/auction/auctionItemcls.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%

dim IS_TEST_MODE : IS_TEST_MODE = True

Dim sqlStr, sellsite, selldate, selldateStr, mode
dim isSuccess
dim i, j, k
dim orderObjArr, tmpObjArr
dim nowdate, fromdate, todate, currdate, hasMoreData, nvCount, doff
Dim chgCode, lastOrderNo, lastTime
Dim isOrderComplete : isOrderComplete = "N"
hasMoreData = "N"
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
selldate	= requestCheckVar(html2db(request("selldate")),32)
mode		= requestCheckVar(html2db(request("mode")),32)
chgCode		= requestCheckVar(html2db(request("chgCode")),32)
doff		= requestCheckVar(html2db(request("doff")),32)
If chgCode = "" Then
	chgCode = "DELIVERED"
End If

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
	case "nvstorefarm"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			If (sellsite = "nvstorefarm") Then
				response.write  sellsite & " : " & currdate & "<br />"
				'1. 최초 테이블 비운다
				If doff <> "Y" Then
					sqlStr = ""
					sqlStr = sqlStr & " DELETE FROM db_temp.[dbo].[tbl_xSite_TMPOrder_storefarm] WHERE sellsite = '"& sellsite &"' and regdate = '"& currdate &"' "
					dbget.Execute sqlStr
				End If

				Do Until isOrderComplete = "Y"
					Call GetOrder_nvstorefarm(sellsite, currdate, hasMoreData, chgCode, lastOrderNo, lastTime)
					If hasMoreData = "N" Then
						isOrderComplete = "Y"
					End If
					response.flush
				Loop
				'2. 위 call해서 얻은 주문번호의 상세내역을 다시 콜한다.
				Call GetOrderFrom_NewCall_nvstorefarm(sellsite, currdate)
			End If
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
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
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
