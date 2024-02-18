<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 쿠팡 주문 확인처리
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
<!-- #include virtual="/outmall/coupang/coupangItemcls.asp"-->
<!-- #include virtual="/outmall/coupang/incCoupangFunction.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Function fnCoupangConfirmOrder(vOrderserial, vOutMallOptionNo, vBeasongNum11st)
	Dim objXML, xmlDOM, iRbody, strSql, istrParam
	istrParam = "lstshipmentboxIds="&vBeasongNum11st
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "http://xapi.10x10.co.kr:8080/Orders/Coupang/ready", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If Err.number <> 0 Then
			iErrStr = ivendorItemId
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			strSql = ""
			strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] SET "
			strSql = strSql & " isbaljuConfirmSend = 'Y' "
			strSql = strSql & " , lastUpdate = getdate() "
			strSql = strSql & " WHERE outmallorderserial = '"&vOrderserial&"'  "
			strSql = strSql & " and beasongNum11st = '"&vBeasongNum11st&"' "
			strSql = strSql & " and OrgDetailKey = '"&vOutMallOptionNo&"' "
			strSql = strSql & " and mallid = 'coupang' "
			dbget.Execute strSql
			fnCoupangConfirmOrder= true
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

' Function fnCoupangConfirmOrder(vOrderserial, vOutMallOptionNo, vBeasongNum11st)
' 	Dim objXML, xmlDOM, iRbody, strSql, istrParam
' 	Dim strObj, strObjitems, i, succeed
' 	istrParam = "lstshipmentboxIds="&vBeasongNum11st
' 	On Error Resume Next
' 	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
' 		objXML.open "POST", "http://xapi.10x10.co.kr:8080/Orders/Coupang/ready", false
' 		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
' 		objXML.Send(istrParam)

' 		If Err.number <> 0 Then
' 			iErrStr = ivendorItemId
' 			Exit Function
' 		End If
' 		rw objXML.Status
' 		rw BinaryToText(objXML.ResponseBody,"utf-8")

' 		If objXML.Status = "200" OR objXML.Status = "201" Then
' 			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
' 			Set strObj = JSON.parse(iRbody)
' 				set strObjitems = strObj.data.responseList
' 					For i=0 to strObjitems.length-1
' 						succeed = strObjitems.get(i).succeed
' 					Next
' 				set strObjitems = nothing
' 			Set strObj = nothing

' 			If succeed <> "False" Then
' 				strSql = ""
' 				strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] SET "
' 				strSql = strSql & " isbaljuConfirmSend = 'Y' "
' 				strSql = strSql & " , lastUpdate = getdate() "
' 				strSql = strSql & " WHERE outmallorderserial = '"&vOrderserial&"'  "
' 				strSql = strSql & " and beasongNum11st = '"&vBeasongNum11st&"' "
' 				strSql = strSql & " and OrgDetailKey = '"&vOutMallOptionNo&"' "
' 				strSql = strSql & " and mallid = 'coupang' "
' 				dbget.Execute strSql
' 				fnCoupangConfirmOrder= true
' 			End If
' 		End If
' 	Set objXML = nothing
' 	On Error Goto 0
' End Function


Dim i, j, sellsite
Dim idx, objXML, xmlDOM
Dim jenkinsBatchYn, lastErrStr
sellsite		= requestCheckVar(html2db(request("sellsite")),32)
idx				= requestCheckVar(html2db(request("idx")),32)
jenkinsBatchYn	= request("jenkinsBatchYn")

Dim tmpxml, strsql, retVal, errStr, succCNT, failCNT
Dim OrgDetailKey
Dim OutMallOrderSerial, SellDate, outMallGoodsNo, matchItemID, beasongNum11st
Dim strObj, iRbody
If sellsite = "coupang" and jenkinsBatchYn = "Y" Then
	On Error Resume Next
		Dim arrList, lp, ret1
		Dim OKcnt, NOcnt
		OKcnt = 0
		NOcnt = 0

		strsql = ""
		strsql = strsql & " update T "
		strsql = strsql & " set T.isbaljuConfirmSend='Y' "
		strsql = strsql & " From db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T "
		strsql = strsql & " JOIN db_temp.dbo.tbl_xsite_tmporder as O on T.outmallorderserial = O.OutMallOrderSerial and T.OrgDetailKey = O.outMallOptionNo "
		strsql = strsql & " where T.isbaljuConfirmSend <> 'Y' "
		strsql = strsql & " and O.sendState = 1 "
		strsql = strsql & " and O.matchstate in ('O') "
		strsql = strsql & " and T.mallid = 'coupang' "
		dbget.Execute strsql

		strsql = ""
		strsql = strsql & " update T "
		strsql = strsql & " set T.isbaljuConfirmSend='Y' "
		strsql = strsql & " FROM db_order.dbo.tbl_order_master as M "
		strsql = strsql & " JOIN db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T on M.authcode = T.outmallorderserial "
		strsql = strsql & " WHERE M.cancelyn ='Y' "
		strsql = strsql & " and T.isbaljuConfirmSend <> 'Y' "
		strsql = strsql & " and T.mallid = 'coupang' "
		dbget.Execute strsql

		strsql = ""
		strsql = strsql & " SELECT TOP 100 outmallorderserial, OrgDetailKey, beasongNum11st FROM db_temp.[dbo].[tbl_xSite_TMP11stOrder] "
		strsql = strsql & " WHERE isbaljuConfirmSend = 'N' "
		strsql = strsql & " and mallid = 'coupang' "
		rsget.Open strsql,dbget,1
		if not rsget.Eof then
			arrList = rsget.getRows()
		end if
		rsget.close

		If isArray(arrList) then
			For lp = 0 To Ubound(arrList, 2)
				ret1 = fnCoupangConfirmOrder(arrList(0, lp), arrList(1, lp), arrList(2, lp))

				If (ret1) then
					OKcnt = OKcnt + 1
				Else
					NOcnt = NOcnt + 1
				End If
			Next

			If OKcnt <> 0 then
				rw "["&OKcnt&"] 건 성공(발주확인)"
			End If

			If NOcnt <> 0 then
				rw "["&NOcnt&"] 건 실패(발주확인)"
			End If
	'		response.end
		Else
			rw "발주확인건 없음"
		End If
	On Error Goto 0
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->