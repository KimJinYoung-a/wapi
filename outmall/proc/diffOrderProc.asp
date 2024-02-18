<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, mallid, orderState, paramData, action, retVal, isSuccessYn, idx
Dim sqlStr, procStr
itemid			= request("itemid")
mallid			= request("mallid")
orderState		= request("orderState")
idx				= request("idx")

If mallid = "gseshop" Then mallid = "gsshop"
action = "EDIT"
If mallid = "coupang" Then
	If orderState = "P" Then
		action = "PRICE"
	ElseIf orderState = "S" Then
		action = "SOLDOUT"
	End If
End If

Select Case mallid
	Case "11st1010"			procStr = "http://wapi.10x10.co.kr/outmall/proc/11stProc.asp"
	Case "auction1010"		procStr = "http://wapi.10x10.co.kr/outmall/proc/AuctionProc.asp"
	Case "cjmall"			procStr = "http://wapi.10x10.co.kr/outmall/proc/CJMallProc.asp"
	Case "ezwel"			procStr = "http://wapi.10x10.co.kr/outmall/proc/EzwelProc.asp"
	Case "gmarket1010"		procStr = "http://wapi.10x10.co.kr/outmall/proc/GmarketProc.asp"
	Case "gsshop"			procStr = "http://wapi.10x10.co.kr/outmall/proc/GSShopProc.asp"
	Case "interpark"		procStr = "http://wapi.10x10.co.kr/outmall/proc/InterparkProc.asp"
	Case "nvstorefarm"		procStr = "http://wapi.10x10.co.kr/outmall/proc/NvstorefarmProc.asp"
	Case "nvstorefarmclass"	procStr = "http://wapi.10x10.co.kr/outmall/proc/NvClassProc.asp"
	Case "nvstoremoonbangu"	procStr = "http://wapi.10x10.co.kr/outmall/proc/nvstoremoonbanguProc.asp"
	Case "lotteimall"		procStr = "http://wapi.10x10.co.kr/outmall/proc/LotteimallProc.asp"
	Case "ssg"				procStr = "http://wapi.10x10.co.kr/outmall/proc/ssgProc.asp"
	Case "coupang"			procStr = "http://wapi.10x10.co.kr/outmall/proc/coupangProc.asp"
	Case "hmall1010"		procStr = "http://wapi.10x10.co.kr/outmall/proc/HmallProc.asp"
	Case "WMP"				procStr = "http://wapi.10x10.co.kr/outmall/proc/wmpProc.asp"
	Case "wmpfashion"		procStr = "http://wapi.10x10.co.kr/outmall/proc/wmpfashionProc.asp"
	Case "lotteon"			procStr = "http://wapi.10x10.co.kr/outmall/proc/lotteonProc.asp"
	Case "lfmall"			procStr = "http://wapi.10x10.co.kr/outmall/proc/lfmallProc.asp"
	Case "shintvshopping"	procStr = "http://wapi.10x10.co.kr/outmall/proc/shintvshoppingProc.asp"
	Case "wetoo1300k"		procStr = "http://wapi.10x10.co.kr/outmall/proc/wetoo1300kProc.asp"
	Case "skstoa"			procStr = "http://wapi.10x10.co.kr/outmall/proc/skstoaProc.asp"
End Select

paramData = "redSsnKey=system&itemid="&itemid&"&mallid="&mallid&"&action="&action&"&idx="&idx

On Error Resume Next
	retVal = SendReq(procStr, paramData)
	Call SugiQueLogInsert(mallid, action, itemid, Split(retVal,"||")(0), retVal, "system")
	If LEFT(retVal, 2) = "OK" Then
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE [db_etcmall].[dbo].[tbl_outmall_diffOrder] SET " & vbcrlf
		sqlStr = sqlStr & " isOk = 'Y' "	 & vbcrlf
		sqlStr = sqlStr & " WHERE idx = '"&idx&"'  "
		dbget.Execute sqlStr
	End If
On Error Goto 0
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->