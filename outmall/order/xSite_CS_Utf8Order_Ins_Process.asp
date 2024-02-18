<%@ language=vbscript %>
<% option explicit %>
<%
Response.CharSet="utf-8"
Session.codepage="65001"
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"
Server.ScriptTimeOut = 60 * 15
%>
<%
'###########################################################
' Description : 제휴몰 API CS입력
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteCSUtf8OrderLib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%

dim IS_TEST_MODE : IS_TEST_MODE = False

dim sellsite, selldate, csGubun, isSuccess
dim i, j, k
dim nowdate, fromdate, todate, currdate
dim sqlStr, msg

sellsite	= requestCheckVar(html2db(request("sellsite")),32)
selldate	= requestCheckVar(html2db(request("selldate")),32)
csGubun		= requestCheckVar(html2db(request("mode")),32)

if (sellsite = "") or (csGubun = "") then
	response.write "잘못된 접근입니다."
	dbget.close : response.end
end if


dim IS_SELLDATE_FIXED : IS_SELLDATE_FIXED = False
if (selldate = "") then
	'// 오늘까지 일괄로 가져오기
	Call GetCSCheckStatus(sellsite, csGubun, selldate, isSuccess)
	fromdate = selldate
	todate = Left(Now, 10)
else
	fromdate = selldate
	todate = selldate
	IS_SELLDATE_FIXED = True
end if

select case csGubun
    case "matchcs"
        '// 제휴 CS내역 어드민 등록IDX 매칭
        Call MatchTenCSAsid(sellsite)
        rw "OK"
        dbget.close : response.end
    case "chkextcs"
        '// 제휴 CS내역 상태체크
        Call CheckExtCsState(sellsite)
        rw "OK"
        dbget.close : response.end
end select

select case sellsite
	case "lfmall"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			if (csGubun = "all") then
				Call GetCSOrderCS_lfmall(sellsite, currdate)
			else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	Case Else
		response.write "잘못된 접근입니다."
		dbget.close : response.end
End Select

if (IS_TEST_MODE = False) and (IS_SELLDATE_FIXED = False) then
	if (selldate < Left(Now(), 10)) then
		Call SetCSCheckStatus(sellsite, csGubun, Left(DateAdd("d", 1, CDate(selldate)), 10), "N")
	elseif (selldate = Left(Now(), 10)) then
		Call SetCSCheckStatus(sellsite, csGubun, selldate, "Y")
	end if

	'// 제휴몰 취소건 어드민 주문취소 : 전체취소만
	for i = 0 to 20
		msg = ""
		sqlStr = " exec [db_cs].[dbo].[usp_Ten_CheckCancelExtOrder] "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			msg = rsget("msg")
		rsget.Close

		response.write msg & "<br />"
		if (msg = "NO ORDER") then
			exit for
		end if
	next
end if

%>
<% rw "OK" %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
