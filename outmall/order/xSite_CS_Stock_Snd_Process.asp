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
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteCSStockoutLib.asp"-->
<!-- #include virtual="/lib/util/aspJSON1.17.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%

dim mode
dim sellsite, itemno, detailidx, orderserial

mode		= requestCheckVar(html2db(request("mode")),32)
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
detailidx	= requestCheckVar(html2db(request("detailidx")),3200)
itemno		= requestCheckVar(html2db(request("itemno")),3200)
orderserial	= requestCheckVar(html2db(request("orderserial")),32)

select case sellsite
	case "ssg"
		'// stockoutOne
		'// stockoutAll
		'// stockoutCnclOne
		Call GetCSStockout_ssg(sellsite, mode, detailidx, orderserial)
	case "coupang"
		'// cancelAll
		Call GetCSStockout_coupang(sellsite, mode, orderserial, detailidx, itemno)
	case "interpark"
		'// cancelAll
		Call GetCSStockout_interpark(sellsite, mode, orderserial, detailidx, itemno)
	case "11st1010"
		'// stockoutOne
		Call GetCSStockout_11st1010(sellsite, mode, detailidx, orderserial)
	case "nvstorefarm"
		'// stockoutOne
		Call GetCSStockout_nvstorefarm(sellsite, mode, detailidx, orderserial)
	case "gmarket1010"
		'// cancelAll
		Call GetCSStockout_gmarket1010(sellsite, mode, detailidx, orderserial)
	case "WMP"
		'// cancelAll
		Call GetCSStockout_WMP(sellsite, mode, detailidx, orderserial)
	case "wmpfashion"
		'// cancelAll
		Call GetCSStockout_WMPfashion(sellsite, mode, detailidx, orderserial)
	' case "lotteon"
	' 	Call GetCSStockout_Lotteon(sellsite, mode, detailidx, orderserial)
	case else
		response.write "잘못된 접근입니다."
		dbget.close : response.end
end select

%>
<% if  (IsAutoScript) then  %>
<% rw "OK" %>
<% else %>
<script>alert('저장되었습니다.');</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
