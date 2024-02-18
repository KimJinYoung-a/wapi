<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 XML 주문처리
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteOrderXMLCls.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/lfmall/lfmallItemcls.asp"-->
<!-- #include virtual="/outmall/lfmall/inclfmallFunction.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim sqlStr, buf, i, j, mode, sellsite
Dim divcd
Dim objXML, xmlDOM, retCode, iMessage
mode		= requestCheckVar(html2db(request("mode")),32)
sellsite	= requestCheckVar(html2db(request("sellsite")),32)

Dim strsql, retVal, deliverymemo, orderCsGbn, errCode, errStr, succCNT, failCNT
Dim OrgDetailKey, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, OrderName, OrderTelNo, OrderHpNo
Dim OutMallOrderSerial, SellDate, outMallGoodsNo, matchItemID, partnerItemName, SellPrice, RealSellPrice, ItemOrderCount, orderDlvPay, requireDetail, matchItemOption, outMallOptionNo
Dim partnerOptionName, SalePrice, beasongNum11st, reserve01
Dim regOrderCnt, strObj, iRbody
Dim iSellDate, iIsSuccess, fromDate, nowDate, searchDate, orderCount

Dim dlvstNo, dlvstPtcSeq, ordNo, lastDlvstPrgrGbcd, dlvTypeGbcd, POS1, POS2, POS3, ReceiveAddr, dlvCnclYn

Call GetCheckStatus("LFmall", iSellDate, iIsSuccess)
searchDate = replace(iSellDate, "-", "")
rw searchDate & " Order START"
'searchDate = "20181119"
If sellsite = "LFmall" Then
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://110.93.128.100:8090/lfmall/order/ordermanage?startdate="&searchDate&"&enddate="&searchDate&"&ordergubun=30", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||" & Err.Description
		End If
'		rw BinaryToText(objXML.ResponseBody,"utf-8")

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			response.write iRbody

			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success
				If isSuccess = true Then
					If (iSelldate < Left(Now(), 10)) then
						Call SetCheckStatus(sellsite, Left(DateAdd("d", 1, CDate(iSellDate)), 10), "N")
					ElseIf (iSellDate = Left(Now(), 10)) then
						Call SetCheckStatus(sellsite, iSellDate, "Y")
					End If
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success
				iMessage		= strObj.message
				rw "주문연동 실패..잠시 후 시도 요망"
				rw iMessage
			Set strObj = nothing
		End If
	On Error Goto 0
	Set objXML = nothing
End If
response.write "<br />"
rw searchDate & " Order End"

''품절/가격 오류체크
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->