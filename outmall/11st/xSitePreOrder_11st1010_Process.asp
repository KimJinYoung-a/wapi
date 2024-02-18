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
<!-- #include virtual="/outmall/11st/11stItemcls.asp"-->
<!-- #include virtual="/outmall/11st/inc11stFunction.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<%
Function fn11stPreConfirmOrder(vOrderserial, vOrgDetailKey)
	Dim objXML, xmlDOM, iRbody, strSql
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & APISSLURL&"/ordservices/saleconfirm/" & vOrderserial & "/" & vOrgDetailKey
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				If xmlDOM.getElementsByTagName("result_code").item(0).text = "0" Then
					strSql = ""
					strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] SET "
					strSql = strSql & " isbaljuConfirmSend = 'Y' "
					strSql = strSql & " , lastUpdate = getdate() "
					strSql = strSql & " WHERE outmallorderserial = '"&vOrderserial&"'  "
					strSql = strSql & " and orgDetailKey = '"&vOrgDetailKey&"' "
					strSql = strSql & " and mallid = 'pre11st1010' "
					dbget.Execute strSql
					fn11stPreConfirmOrder= true
				Else
					fn11stPreConfirmOrder= false
				End If
			Set xmlDOM = Nothing
		Else
			fn11stPreConfirmOrder= false
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

function getLastOrderInputDT()
    dim sqlStr
    sqlStr = "select top 1 convert(varchar(10),selldate,21) as lastOrdInputDt"
    sqlStr = sqlStr&" from db_temp.dbo.tbl_XSite_TMpOrder"
    sqlStr = sqlStr&" where sellsite='11st1010'"
    sqlStr = sqlStr&" order by selldate desc"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
		getLastOrderInputDT = rsget("lastOrdInputDt")
	end if
	rsget.Close
end function

Dim sqlStr, buf, i, mode, sellsite
Dim divcd, yyyymmdd, idx, Nodes, Nodes2, SubNodes, SubNodes2, vOrder
Dim objXML, xmlDOM, retCode, iMessage, reqOrderdate
mode		= requestCheckVar(html2db(request("mode")),32)
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
idx			= requestCheckVar(html2db(request("idx")),32)
reqOrderdate = request("reqOrderdate")

Dim strsql, retVal, deliverymemo, orderCsGbn, errCode, errStr, succCNT
Dim dueDate, iRbody, result_text
Dim orderDlvPay, beasongNum11st, sellsiteUserID, ordAmt, SellDate, OrderName, outMallGoodsNo, ordOptWonStl, ordPayAmt, OrgDetailKey, OrderHpNo, ItemOrderCount, OrderTelNo, partnerItemName
Dim OutMallOrderSerial, prdStckNo, ReceiveAddr1, ReceiveAddr2, ReceiveZipCode, ReceiveName, ReceiveHpNo, ReceiveTelNo, selPrc, sellerDscPrc, matchItemID, partnerOptionName, tmallDscPrc, lstTmallDscPrc, lstSellerDscPrc, sellerStockCd
Dim requireDetail, matchItemOption, SellPrice, RealSellPrice
Dim prev7Day, nowDay, lastOrderDate, resultNode, AssignedRow

If reqOrderdate = "" Then
	lastOrderDate = getLastOrderInputDT
Else
	lastOrderDate = reqOrderdate
End If
'lastOrderDate = "2017-11-13"
prev7Day = CStr(Replace((lastOrderDate), "-", ""))&"0000"
nowDay	 = CStr(Replace(Date(), "-", ""))&"2359"

If (CDate(lastOrderDate) > date()) Then
	response.write "날짜 오류 입니다."
	response.end
End If

dueDate = prev7Day &"/"& nowDay
Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "GET", "" & APISSLURL&"/ordservices/reservatecomplete/"&dueDate
	objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
	objXML.setRequestHeader "openapikey",""&APIkey&""
	objXML.send()
	succCNT = 0
	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
			xmlDOM.LoadXML iRbody

			rw "REQ : " & APISSLURL&"/ordservices/reservatecomplete/"&dueDate
			rw "RES : <textarea cols=40 rows=10>"&BinaryToText(objXML.ResponseBody, "euc-kr")&"</textarea>"
		
			Set vOrder = xmlDOM.getElementsByTagName("ns2:order")
				For each SubNodes in vOrder
					orderCsGbn			= 0
					OutMallOrderSerial	= Trim(SubNodes.getElementsByTagName("ordNo").item(0).text)					'11번가 주문번호
					OrgDetailKey		= Trim(SubNodes.getElementsByTagName("ordPrdSeq").item(0).text)				'주문순번
					SellDate			= Trim(SubNodes.getElementsByTagName("ordDt").item(0).text)					'주문일시

					strsql = ""
					strsql = strsql & " INSERT INTO db_temp.[dbo].[tbl_xSite_TMP11stOrder] (outmallorderserial, OrgDetailKey, isbaljuConfirmSend, regdate, deliveryDate, mallid, beasongNum11st) "
					strsql = strsql & " VALUES ('"&OutMallOrderSerial&"', '"&OrgDetailKey&"', 'N', getdate(), '"& SellDate &"', 'pre11st1010', '')"
					dbget.Execute strSql, AssignedRow
					succCNT = succCNT + 1
				Next
			Set vOrder = nothing
		Set xmlDOM = nothing

		If (succCNT <> 0) then
		    rw "["&succCNT&"] 건 성공(예약주문조회)"
			Dim arrList, lp, ret1
			Dim OKcnt, NOcnt
			OKcnt = 0
			NOcnt = 0
			strsql = ""
			strsql = strsql & " SELECT TOP 3000 outmallorderserial, OrgDetailKey FROM db_temp.[dbo].[tbl_xSite_TMP11stOrder] "
			strsql = strsql & " WHERE isbaljuConfirmSend = 'N' "
			strsql = strsql & " and mallid = 'pre11st1010' "
			strsql = strsql & " ORDER BY regdate DESC "
			rsget.CursorLocation = adUseClient
			rsget.Open strsql, dbget, adOpenForwardOnly, adLockReadOnly
			if not rsget.Eof then
				arrList = rsget.getRows()
			end if
			rsget.close

			For lp = 0 To Ubound(arrList, 2)
				ret1 = fn11stPreConfirmOrder(arrList(0, lp), arrList(1, lp))
				If (ret1) then
					OKcnt = OKcnt + 1
				End If
			Next

			If OKcnt <> 0 then
				rw "["&OKcnt&"] 건 성공(입고완료처리)"
			End If
		End If
	Else
		rw "주문연동 실패..잠시 후 시도 요망"
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
			xmlDOM.LoadXML iRbody
			Set resultNode = xmlDOM.getElementsByTagName("resultCode")
				If NOT (resultNode Is Nothing)  Then
					result_text = xmlDOM.getElementsByTagName("resultMessage").item(0).text
				End If
			Set resultNode = nothing
		Set xmlDOM = nothing

		rw "REQ : " & APISSLURL&"/ordservices/reservatecomplete/"&dueDate
		rw "RES : <textarea cols=40 rows=10>"&BinaryToText(objXML.ResponseBody, "euc-kr")&"</textarea>"
	End If
Set objXML = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->