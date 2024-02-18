<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 롯데On 주문 확인처리
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteOrderXMLCls.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/lotteon/lotteonItemcls.asp"-->
<!-- #include virtual="/outmall/lotteon/inclotteonFunction.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim i, j, sellsite
Dim idx
Dim jenkinsBatchYn, lastErrStr
Dim strsql, succCNT, failCNT, readySeq
Dim arrList, lp, ret1
Dim OKcnt, NOcnt
sellsite		= requestCheckVar(html2db(request("sellsite")),32)
idx				= requestCheckVar(html2db(request("idx")),32)
jenkinsBatchYn	= request("jenkinsBatchYn")
readySeq		= request("readySeq")
'readySeq = 1 : 내부처리 완료 (구매자가 취소처리 못 함)
'readySeq = 2 : 배송상태 변경 (주문 확인 처리)

If sellsite = "lotteon" and jenkinsBatchYn = "Y" AND readySeq <> "" Then
	If readySeq = "1" Then
		On Error Resume Next
			OKcnt = 0
			NOcnt = 0

			strsql = ""
			strsql = strsql & " SELECT TOP 100 outmallorderserial as odNo, OrgDetailKey as odSeq, beasongNum11st as procSeq "
			strsql = strsql & " FROM db_temp.[dbo].[tbl_xSite_TMP11stOrder] "
			strsql = strsql & " WHERE isbaljuConfirmSend = 'N' "
			strsql = strsql & " and isTenConfirmSend = 'N' "
			strsql = strsql & " and mallid ='lotteon' "
			rsget.CursorLocation = adUseClient
			rsget.Open strsql, dbget, adOpenForwardOnly, adLockReadOnly
			if not rsget.Eof then
				arrList = rsget.getRows()
			end if
			rsget.close

			If isArray(arrList) then
				For lp = 0 To Ubound(arrList, 2)
					ret1 = fnlotteonTenConfirmOrder(arrList(0, lp), arrList(1, lp), arrList(2, lp))
					If (ret1) then
						OKcnt = OKcnt + 1
					Else
						NOcnt = NOcnt + 1
					End If
				Next

				If OKcnt <> 0 then
					rw "["&OKcnt&"] 건 성공(내부확인)"
				End If

				If NOcnt <> 0 then
					rw "["&NOcnt&"] 건 실패(내부확인)"
				End If
		'		response.end
			Else
				rw "내부확인건 없음"
			End If
		On Error Goto 0
	Else
		On Error Resume Next

			OKcnt = 0
			NOcnt = 0

			strsql = ""
			strsql = strsql & " UPDATE T "
			strsql = strsql & " SET T.isbaljuConfirmSend='Y' "
			strsql = strsql & " FROM db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T "
			strsql = strsql & " JOIN db_temp.dbo.tbl_xsite_tmporder as O on T.outmallorderserial = O.OutMallOrderSerial and T.OrgDetailKey = O.OrgDetailKey "
			strsql = strsql & " WHERE T.isbaljuConfirmSend <> 'Y' "
			strsql = strsql & " and O.sendState = 1 "
			strsql = strsql & " and O.matchstate in ('O') "
			strsql = strsql & " and T.mallid = 'lotteon' "
			dbget.Execute strsql

			strsql = ""
			strsql = strsql & " UPDATE T "
			strsql = strsql & " SET T.isbaljuConfirmSend='Y' "
			strsql = strsql & " FROM db_order.dbo.tbl_order_master as M "
			strsql = strsql & " JOIN db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T on M.authcode = T.outmallorderserial "
			strsql = strsql & " WHERE M.cancelyn ='Y' "
			strsql = strsql & " and T.isbaljuConfirmSend <> 'Y' "
			strsql = strsql & " and T.mallid = 'lotteon' "
			dbget.Execute strsql

			strsql = ""
			strsql = strsql & " SELECT TOP 100 outmallorderserial as odNo, OrgDetailKey as odSeq, beasongNum11st as procSeq, outMallGoodsNo as spdNo, outMallOptionNo as sitmNo, ItemOrderCount as slQty "
			strsql = strsql & " FROM db_temp.[dbo].[tbl_xSite_TMP11stOrder] "
			strsql = strsql & " WHERE isbaljuConfirmSend = 'N' "
			strsql = strsql & " and isTenConfirmSend = 'Y' "
			strsql = strsql & " and mallid ='lotteon' "
			rsget.CursorLocation = adUseClient
			rsget.Open strsql, dbget, adOpenForwardOnly, adLockReadOnly
			if not rsget.Eof then
				arrList = rsget.getRows()
			end if
			rsget.close

			If isArray(arrList) then
				For lp = 0 To Ubound(arrList, 2)
					ret1 = fnlotteonConfirmOrder(arrList(0, lp), arrList(1, lp), arrList(2, lp), arrList(3, lp), arrList(4, lp), arrList(5, lp))
					If (ret1) then
						OKcnt = OKcnt + 1
					Else
						NOcnt = NOcnt + 1
					End If
				Next

				If OKcnt <> 0 then
					rw "["&OKcnt&"] 건 성공(주문확인)"
				End If

				If NOcnt <> 0 then
					rw "["&NOcnt&"] 건 실패(주문확인)"
				End If
		'		response.end
			Else
				rw "주문확인건 없음"
			End If
		On Error Goto 0
	End If
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->