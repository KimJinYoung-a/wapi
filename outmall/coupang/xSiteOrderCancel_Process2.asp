<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 3000 ''초단위
'###########################################################
' Description : 제휴몰 주문 취소
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/outmall/jungsan/lib/xSiteJungsanLib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim i
Dim reqDate : reqDate = request("reqDate")
Dim isCancelComplete, hasnext, nextToken
isCancelComplete = "N"

Do Until isCancelComplete = "Y"
	Call GetCance_Coupang(reqDate, hasnext, nextToken)
	If hasnext = "N" Then
		isCancelComplete = "Y"
		rw "complete"
	Else
		rw "API Calling"
		rw "------->> " & nextToken
	End If
	response.flush
Loop

Function GetCance_Coupang(reqDate, ihasnext, inextToken)
	Dim sellsite : sellsite = "coupang"
	Dim url, path, method, params
	Dim access_key, secret_key, vendorId
	Dim authorization

	Dim objXML, xmlDOM, sqlStr, iRbody, strObj
	Dim retCode, iMessage
	Dim datalist, i, j, itemlist, datalistReturnItems


	dim startdate, enddate
	dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, divcd, gubunname, OutMallRegDate, itemno, OrderName, OrderHpNo
	dim iAssignedRow, iInputCnt
	dim strSql, ResultCode, ResultMsg, cancelOrderList, SubNodes
	Dim ClaimSeq, ClaimMemo, RegYMD
    dim OrgOutMallOrderSerial

	access_key = "0af06fb7-3deb-4ac3-9a84-6d409a26d831"
	secret_key = "5474f1108ac5631e5977d4a6b7a6387426533582"
	vendorId = "A00039305"
	path = "/v2/providers/openapi/apis/api/v4/vendors/"&vendorId&"/returnRequests"
	params = "searchType=&createdAtFrom="&reqDate&"&createdAtTo="&reqDate&"&cancelType=CANCEL&nextToken="&inextToken
'	rw params

	url = "https://api-gateway.coupang.com" & path
	method = "GET"
	authorization = generateHmac(path, method, params, access_key, secret_key)
	iInputCnt = 0

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open method, url & "?" & params, false
		objXML.setRequestHeader "Authorization", authorization
		objXML.setRequestHeader "X-Requested-By", vendorId
		objXML.setRequestHeader "X-Extended-Timeout", "60000"
		objXML.send()
 		If objXML.Status = "200" OR objXML.Status = "201" Then
 			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				retCode		= strObj.code			'서버 응답 코드
				iMessage	= strObj.message		'detail info

				If strObj.nextToken <> "" Then		'다음페이지에 데이터 존재 여부
					hasnext = "Y"
					inextToken = strObj.nextToken
				Else
					hasnext = "N"
				End If

				If retCode = "200" Then
					Set datalist = strObj.data		'결과리스트 | 결과가 없을 때는 빈 리스트가 리턴
						For i=0 to datalist.length-1
                            OrgOutMallOrderSerial	= datalist.get(i).orderId				'주문번호
							CSDetailKey				= datalist.get(i).receiptId				'취소(반품)접수번호
							OutMallOrderSerial		= datalist.get(i).orderId				'주문번호
                            if (datalist.get(i).receiptType = "CANCEL") then
                                divcd = "A008"
                            elseif (datalist.get(i).receiptType = "RETURN") then
                                divcd = "A004"
                            else
                                divcd = "AXXX"
                            end if

							OutMallRegDate		= datalist.get(i).createdAt				'취소(반품) 접수시간
							OutMallRegDate		= Replace(OutMallRegDate, "T", " ")
							OrderName			= datalist.get(i).requesterName				'반품 신청인 이름
							OrderHpNo			= datalist.get(i).requesterPhoneNumber		'반품 신청인 전화번호
							gubunname			= datalist.get(i).cancelReasonCategory1	'반품 사유 카테고리1

							set datalistReturnItems = datalist.get(i).returnItems
								For j=0 to datalistReturnItems.length-1
									OrgDetailKey = datalistReturnItems.get(j).vendorItemId	'벤더아이템번호
									itemno = datalistReturnItems.get(j).purchaseCount		'취소 수량

						            strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
									strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
									strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno, OrgOutMallOrderSerial) VALUES "
									strSql = strSql & " 	('" & divcd & "', '" & gubunname & "', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
									strSql = strSql & "		'', '', '', '', '', '' "
									strSql = strSql & "		, '" & OutMallRegDate & "', '" & CStr(OrgDetailKey) & "', '" & CStr(CSDetailKey) & "', " & itemno & ", '" & OrgOutMallOrderSerial & "') "
									strSql = strSql & " END "
									strSql = strSql & " ELSE "
									strSql = strSql & " BEGIN "
									strSql = strSql & " 	update db_temp.dbo.tbl_xSite_TMPCS "
									strSql = strSql & " 	set divcd = '" & divcd & "', OutMallRegDate = '" & OutMallRegDate & "', currstate = 'B001' "
									strSql = strSql & " 	WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and divcd <> '" & divcd & "' "
									strSql = strSql & " END "
									'rw strSql
									dbget.Execute strSql,iAssignedRow

									if (iAssignedRow > 0) then
										iInputCnt = iInputCnt+iAssignedRow

										'' CS 마스터정보. 업데이트
										strSql = " update c "
										strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
										strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
										strSql = strSql + " , c.OrderName = o.OrderName "
										strSql = strSql + " from "
										strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
										strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
										strSql = strSql + " on "
										strSql = strSql + " 	1 = 1 "
										strSql = strSql + " 	and c.SellSite = o.SellSite "
										strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
										strSql = strSql + " 	and c.OrgDetailKey = o.OutMallOptionNo "
										strSql = strSql + " where "
										strSql = strSql + " 	1 = 1 "
										strSql = strSql + " 	and c.orderserial is NULL "
										strSql = strSql + " 	and o.orderserial is not NULL "
										strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
										'response.write strSql & "<br />"
										dbget.Execute strSql
									end if
								Next
							set datalistReturnItems = nothing
						Next
					Set datalist = nothing
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
End Function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->