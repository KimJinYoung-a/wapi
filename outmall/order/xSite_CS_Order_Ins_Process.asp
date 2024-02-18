<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 180 ''초단위
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
<!-- #include virtual="/outmall/shintvshopping/inc_authCheck.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteCSOrderLib.asp"-->
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
	case "gmarket1010"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			if (csGubun = "ordercancel") then
				''ClaimReady : 취소신청
				''CalimDone : 취소완료
				''ClaimReject : 취소철회
				''ClaimDoneG : G마켓 직권 환불 건만 조회 (취소 완료 건 중 고객센터에서 환불 처리 한 Case)
				Call GetCSOrderCancel_gmarket(sellsite, csGubun, "ClaimReady", currdate)
				Call GetCSOrderCancel_gmarket(sellsite, csGubun, "ClaimDone", currdate)
				Call GetCSOrderCancel_gmarket(sellsite, csGubun, "ClaimReject", currdate)
				Call GetCSOrderCancel_gmarket(sellsite, csGubun, "ClaimDoneG", currdate)
			elseif (csGubun = "return") then
				''ClaimReady
				''ClaimDone
				''ClaimReject
				''ClaimDoneG
				Call GetCSOrderReturn_gmarket(sellsite, csGubun, "ClaimReady", currdate)
				Call GetCSOrderReturn_gmarket(sellsite, csGubun, "ClaimDone", currdate)
				Call GetCSOrderReturn_gmarket(sellsite, csGubun, "ClaimReject", currdate)
				Call GetCSOrderReturn_gmarket(sellsite, csGubun, "ClaimDoneG", currdate)
			elseif (csGubun = "exchange") then
				''ClaimReady
				''ClaimDone
				''ClaimReject
				Call GetCSOrderExchange_gmarket(sellsite, csGubun, "ClaimReady", currdate)
				Call GetCSOrderExchange_gmarket(sellsite, csGubun, "ClaimDone", currdate)
				Call GetCSOrderExchange_gmarket(sellsite, csGubun, "ClaimReject", currdate)
			elseif (csGubun = "all") then
				Call GetCSOrderCancel_gmarket(sellsite, csGubun, "ClaimReady", currdate)
				Call GetCSOrderCancel_gmarket(sellsite, csGubun, "ClaimDone", currdate)
				Call GetCSOrderCancel_gmarket(sellsite, csGubun, "ClaimReject", currdate)
				Call GetCSOrderCancel_gmarket(sellsite, csGubun, "ClaimDoneG", currdate)

				Call GetCSOrderReturn_gmarket(sellsite, csGubun, "ClaimReady", currdate)
				Call GetCSOrderReturn_gmarket(sellsite, csGubun, "ClaimDone", currdate)
				Call GetCSOrderReturn_gmarket(sellsite, csGubun, "ClaimReject", currdate)
				Call GetCSOrderReturn_gmarket(sellsite, csGubun, "ClaimDoneG", currdate)

				Call GetCSOrderExchange_gmarket(sellsite, csGubun, "ClaimReady", currdate)
				Call GetCSOrderExchange_gmarket(sellsite, csGubun, "ClaimDone", currdate)
				Call GetCSOrderExchange_gmarket(sellsite, csGubun, "ClaimReject", currdate)
			else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "interpark"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			if (csGubun = "all") then
				Call GetCSOrderAll_interpark(sellsite, csGubun, "", currdate)
				Call GetCSOrderChgRet_interpark(sellsite, csGubun, "", currdate)
			else
				response.write "잘못된 접근입니다.2"
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "nvstorefarm", "nvstoregift", "Mylittlewhoopee"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			if (csGubun = "all") then
				Call GetCSOrderAll_nvstorefarm(sellsite, "CANCELED", "", currdate)
				'Call GetCSOrderAll_nvstorefarm(sellsite, "RETURNED", "", currdate)
				Call GetCSOrderAll_nvstorefarm(sellsite, "CANCEL_REQUESTED", "", currdate)	'2019-07-23 김진영 CANCEL_REQUESTED로 변경
				Call GetCSOrderAll_nvstorefarm(sellsite, "RETURN_REQUESTED", "", currdate)	'2019-07-23 김진영 RETURN_REQUESTED로 변경
				Call GetCSOrderAll_nvstorefarm(sellsite, "EXCHANGE_REQUESTED", "", currdate)	'2022-03-03 김진영 추가
				'// 교환건 없는 듯...
				''Call GetCSOrderAll_nvstorefarm(sellsite, "EXCHANGED", "", currdate)
			else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "11st1010"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			if (csGubun = "all") then
				Call GetCSOrderCancel_11st1010(sellsite, "ClaimReady", "", currdate)
				Call GetCSOrderCancel_11st1010(sellsite, "ClaimDone", "", currdate)

				Call GetCSOrderExchange_11st1010(sellsite, "EXCHANGED", "", currdate)
				Call GetCSOrderReturn_11st1010(sellsite, "RETURNED", "", currdate)
			else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "gseshop"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -1, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			if (csGubun = "ordercancel") then
				Call GetCSOrderCancel_gseshop(sellsite, "", "", currdate)
			ElseIf (csGubun = "orderNewcancel") then
				Call GetCSOrderNewCancel_gseshop(sellsite, "", "", currdate)
			else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "halfclub"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			if (csGubun = "all") then
				Call GetCSOrderCancel_halfclub(sellsite, "CANCELED", "", currdate)
				Call GetCSOrderReturn_halfclub(sellsite, "RETURNED", "", currdate)
			else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "coupang"
        'rw "TEST중"
        'response.end
        select case csGubun
            case "all"
                '// 취소/반품 내역 가져오기
		        if (selldate = Left(Now(), 10)) then
			        fromdate = Left(DateAdd("d", -5, Now()), 10)
		        end if

		        currdate = fromdate
		        do while (currdate <= todate)
			        response.write "<br />" & sellsite & " : " & currdate & "<br />"
				    Call GetCSOrderCancel_coupang(sellsite, "CANCEL", "", currdate)
					Call GetCSOrderCancel_coupang(sellsite, "RETURN", "", currdate)
				    Call GetCSOrderReturn_coupang(sellsite, "EXCHANGE", "", currdate)

			        selldate = currdate
			        currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		        loop
            case "matchcs"
                '// 제휴 CS내역 어드민 등록IDX 매칭
                Call MatchTenCSAsid(sellsite)
            case "chkextcs"
                '// 제휴 CS내역 상태체크
                Call CheckExtCsState(sellsite)
            case else
                response.write "잘못된 접근입니다."
				dbget.close : response.end
        end select


	case "hmall1010"
'################################### 2022-05-31 김진영 주석처리..쿼리 느려짐 ############################################
'		Dim selldate2, fromdate2, currdate2, todate2
'		selldate2 = selldate
'		todate2 = todate
'		if (selldate2 = Left(Now(), 10)) then
'			fromdate2 = Left(DateAdd("d", -14, Now()), 10)
'		end if
'
'		currdate2 = fromdate2
'		todate2 = Left(DateAdd("d", 7, fromdate2), 10)
'
'
'		do while (currdate2 <= todate2)
'			response.write "<br />1. " & sellsite & " : " & currdate2 & "<br />"
'			if (csGubun = "all") then
'				Call GetCSOrderCancel_hmall(sellsite, "CANCELED", "", currdate2)
'				Call GetCSOrderReturn_hmall(sellsite, "RETURNED", "", currdate2)
'			else
'				response.write "잘못된 접근입니다."
'				dbget.close : response.end
'			end if
'
'			selldate2 = currdate2
'			currdate2 = Left(DateAdd("d", 1, CDate(currdate2)), 10)
'		loop
'#######################################################################################################################
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -6, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "<br />2. " & sellsite & " : " & currdate & "<br />"
			if (csGubun = "all") then
				Call GetCSOrderCancel_hmall(sellsite, "CANCELED", "", currdate)
				Call GetCSOrderReturn_hmall(sellsite, "RETURNED", "", currdate)
			else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "WMP"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			if (csGubun = "all") then
				Call GetCSOrderCS_WMP(sellsite, "CANCEL", "", currdate)
				Call GetCSOrderCS_WMP(sellsite, "CANCELDONE", "", currdate)
				Call GetCSOrderCS_WMP(sellsite, "EXCHANGE", "", currdate)
				Call GetCSOrderCS_WMP(sellsite, "RETURN", "", currdate)
			else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "wmpfashion"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			if (csGubun = "all") then
				Call GetCSOrderCS_wmpfashion(sellsite, "CANCEL", "", currdate)
				Call GetCSOrderCS_wmpfashion(sellsite, "EXCHANGE", "", currdate)
				Call GetCSOrderCS_wmpfashion(sellsite, "RETURN", "", currdate)
			else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	Case "shintvshopping"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -1, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			if (csGubun = "ordercancel") then
				Call GetCSOrderCancel_shintvshopping(sellsite, csGubun, "", currdate)
				'https://wapi.10x10.co.kr/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=shintvshopping&mode=ordercancel
			ElseIf (csGubun = "returnchange") then
				Call GetCSOrderReturnExchange_shintvshopping(sellsite, csGubun, "", currdate)
				'https://wapi.10x10.co.kr/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=shintvshopping&mode=returnchange
			ElseIf (csGubun = "all") then
				Call GetCSOrderCancel_shintvshopping(sellsite, csGubun, "", currdate)
				rw "======================Cancel End======================"
				Call GetCSOrderReturnExchange_shintvshopping(sellsite, csGubun, "", currdate)
			else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	Case "skstoa"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			if (csGubun = "ordercancel") then
				Call GetCSOrderCancel_skstoa(sellsite, csGubun, "", currdate)
				'https://wapi.10x10.co.kr/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=skstoa&mode=ordercancel
			ElseIf (csGubun = "returnchange") then
				Call GetCSOrderReturnExchange_skstoa(sellsite, csGubun, "", currdate)
				'https://wapi.10x10.co.kr/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=skstoa&mode=returnchange
			ElseIf (csGubun = "all") then
				Call GetCSOrderCancel_skstoa(sellsite, csGubun, "", currdate)
				rw "======================Cancel End======================"
				Call GetCSOrderReturnExchange_skstoa(sellsite, csGubun, "", currdate)
			else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	Case "wetoo1300k"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			if (csGubun = "ordercancel") then
				Call GetCSOrderCancel_wetoo1300k(sellsite, csGubun, "", currdate)
				Call GetCSOrderCancel2_wetoo1300k(sellsite, csGubun, "", currdate)
				'https://wapi.10x10.co.kr/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=wetoo1300k&mode=ordercancel
			ElseIf (csGubun = "return") then
				Call GetCSOrderReturn_wetoo1300k(sellsite, csGubun, "", currdate)
				'https://wapi.10x10.co.kr/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=wetoo1300k&mode=returnchange
			ElseIf (csGubun = "all") then
				Call GetCSOrderCancel_wetoo1300k(sellsite, csGubun, "", currdate)
				Call GetCSOrderCancel2_wetoo1300k(sellsite, csGubun, "", currdate)
				rw "======================Cancel End======================"
				Call GetCSOrderReturn_wetoo1300k(sellsite, csGubun, "", currdate)
			else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	Case "lotteon"
		If (selldate = Left(Now(), 10)) Then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		End If

		currdate = fromdate
		Do While (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			If (csGubun = "ordercancel") Then
				Call GetCSOrderCancel_lotteon(sellsite, csGubun, "", currdate)
			ElseIf (csGubun = "return") Then
				Call GetCSOrderReturn_lotteon(sellsite, csGubun, "", currdate)
				Call GetCSOrderReturnReject_lotteon(sellsite, csGubun, "", currdate)
			ElseIf (csGubun = "exchange") then
				Call GetCSOrderExchange_lotteon(sellsite, csGubun, "", currdate)
				Call GetCSOrderExchangeReject_lotteon(sellsite, csGubun, "", currdate)
			ElseIf (csGubun = "all") then
				Call GetCSOrderCancel_lotteon(sellsite, csGubun, "", currdate)

				Call GetCSOrderReturn_lotteon(sellsite, csGubun, "", currdate)
				Call GetCSOrderReturnReject_lotteon(sellsite, csGubun, "", currdate)

				Call GetCSOrderExchange_lotteon(sellsite, csGubun, "", currdate)
				Call GetCSOrderExchangeReject_lotteon(sellsite, csGubun, "", currdate)
			Else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			End If

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		Loop
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
