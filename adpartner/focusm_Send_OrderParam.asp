<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/commlib.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
	dim lp, strSql, arrRst

	'// 접수된 주문중 출고 완료건 취합 (실제 출고수량으로 numOfItem 계산)
	strSql = "select top 50 f.orderserial, f.itemid, f.transactionId, f.mediaCode, f.adCode, f.gadid, sum(d.itemno) as numOfItem " & vbCrLf &_
			" from db_order.dbo.tbl_order_master as m with (noLock) " & vbCrLf &_
			" 	join db_order.dbo.tbl_order_detail as d with (noLock) " & vbCrLf &_
			" 		on m.orderserial=d.orderserial " & vbCrLf &_
			" 	join db_temp.dbo.tbl_focusm_orderInfo as f with (noLock) " & vbCrLf &_
			" 		on m.orderserial=f.orderserial " & vbCrLf &_
			" 			and d.itemid=f.itemid " & vbCrLf &_
			" where m.ipkumdiv>3 " & vbCrLf &_
			" 	and m.jumundiv not in ('6','9') " & vbCrLf &_
			" 	and m.cancelyn='N' " & vbCrLf &_
			" 	and d.cancelyn='N' " & vbCrLf &_
			" 	and d.currstate=7 " & vbCrLf &_	
			" 	and f.status=1" & vbCrLf &_
			" group by f.orderserial, f.itemid, f.transactionId, f.mediaCode, f.adCode, f.gadid"
	rsget.Open strSql,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		arrRst = rsget.getRows()
	end if
	
	rsget.Close

	if isArray(arrRst) then
		for lp=0 to ubound(arrRst,2)
			Call FnSendComplete(arrRst(2,lp), arrRst(3,lp), arrRst(4,lp), arrRst(5,lp), arrRst(6,lp), arrRst(0,lp), arrRst(1,lp))
		next
	end if

	'// 출고완료 정보 포커스엠에 전송
	Sub FnSendComplete(transactionId, mediaCode, adCode, gadid, numOfItem, orderserial, itemid)
		Dim oXML
		set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
		oXML.open "POST", "http://ad.focusm.kr/receive/event/tenByTen.php", false			'동기 처리
		oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		oXML.send "eventType=31&transactionId=" & transactionId & "&mediaCode=" & mediaCode & "&adCode=" & adCode & "&gadid=" & gadid & "&numOfItem=" & numOfItem
		'통신 결과 출력
		if oXML.status=200 then
			dim oRstJS, rstCd
			set oRstJS = JSON.parse(oXML.responseText)
			rstCd = oRstJS.result
			set oRstJS = Nothing

			'상태값 저장(Status - 0:대기, 1:주문완료, 2:배송완료, 8:주문정보 전송오류, 9:배송정보 전송오류)
			strSql = "UPDATE db_temp.dbo.tbl_focusm_orderInfo SET status=" & chkIIF(cStr(rstCd)="1","2","9") & " where orderserial='" & orderserial & "' and itemid='" & itemid & "'"
			dbget.execute(strSql)
		end if

		Set oXML = Nothing	'컨퍼넌트 해제
	end Sub
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->