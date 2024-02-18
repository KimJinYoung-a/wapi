<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteCSOrderCls.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/outmall/ezwel/ezwelItemcls.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<%
'// 2014-08-27, skyer9
''Server.ScriptTimeout = 60
'' response.write lotteAuthNo
'' response.end
Dim refer
refer = request.ServerVariables("HTTP_REFERER")

Dim sqlStr, buf
Dim i, j, k

'주문완료 (1001) /출고준비중 (1002) /배송중 (1003) /수취완료 (1004) /주문취소 (1005) /반품요청 (1007)
'반품완료 (1008) /교환요청 (1011) /교환완료 (1012) /반품후 주문취소 (1009) /오류 (1010)/품절취소요청 (1013)/품절취소 (1014)

'// ============================================================================
'// [divcd]
'// ============================================================================
'A008			주문취소
'
'A004			반품접수(업체배송)
'A010			회수신청(텐바이텐배송)
'
'A001			누락재발송
'A002			서비스발송
'
'A000			맞교환출고
'A100			상품변경 맞교환출고
'
'A009			기타사항
'A006			출고시유의사항
'A700			업체기타정산
'
'A003			환불
'A005			외부몰환불요청
'A007			카드,이체,휴대폰취소요청
'
'A011			맞교환회수(텐바이텐배송)
'A012			맞교환반품(업체배송)

'A111			상품변경 맞교환회수(텐바이텐배송)
'A112			상품변경 맞교환반품(업체배송)
'// ============================================================================

Dim mode
Dim sellsite
Dim reguserid
Dim AssignedRow
Dim ErrMsg

Dim resultCount

Dim divcd, yyyymmdd, idx, finUserid
Dim getDivCD, sDate, eDate

Dim postParam
Dim objXML, xmlDOM, strSql
Dim retCode, goodsCd, iMessage, oMsg, ocount, stdt, eddt
Dim parentNodes, parentSubNodes, Nodes, masterSubNodes
mode		= requestCheckVar(html2db(request("mode")),32)
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
idx			= requestCheckVar(html2db(request("idx")),32)
finUserid	= session("ssBctID")
If finUserid = "" Then
	finUserid = "system"
End If

If (mode = "getxsitecslist") Then
    If (sellsite="ezwel") Then
    	ErrMsg = ""
		getDivCD = Trim(application("xSiteGetEzwelCS_DIVCD"))
		If (getDivCD = "") Then
			getDivCD = "A008"
		ElseIf (getDivCD = "A004") Then
			getDivCD = "A008"
		Else
			getDivCD = "A004"
		End If

		postParam = "cspCd="&cspCd&"&crtCd="&crtCd
		If getDivCD = "A008" Then				'주문취소
			stdt = getLastCSInputDT("ordercancel")
		Else									'반품
			stdt = getLastCSInputDT("return")
		End If
		eddt = Replace(Date,"-","") & Replace(FormatDateTime(Now,4),":","") & Right(Now,2)
		postParam = postParam & "&startDate="&stdt&"&endDate="&eddt

'		If Hour(Now()) < 6 then
'			postParam = postParam & "&startDate="&Replace(Date-1,"-","") & "000000"&"&endDate="&Replace(Date-1,"-","") & "235959"
'		Else
'			postParam = postParam & "&startDate="&Replace(Date,"-","") & "000000"&"&endDate="&Replace(Date,"-","") & Replace(FormatDateTime(Now,4),":","") & Right(Now,2)
'		End If

		If getDivCD = "A008" Then				'주문취소
			postParam = postParam & "&orderStatus=1005"
		Else									'반품
			postParam = postParam & "&orderStatus=1007"
		End If

'		On Error Resume Next
		Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.open "POST", "http://api.ezwel.com/if/api/orderListAPI.ez", false
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
			objXML.send(postParam)
			If objXML.Status = "200" Then
				Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
					xmlDOM.async = False
					xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
					If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
						'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
					End If
					retCode = xmlDOM.getElementsByTagName("resultCode").item(0).text
					If retCode = "200" Then		'성공(200)
						Dim retVal, succCnt, failCnt
						Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, rcvrNm, rcvrTelNum, rcvrMobile, rcvrPost, rcvrAddr1, rcvrAddr2, orderDt, orderQty, sndNm, sndTelNum, sndMobile, orderReqContent
						succCnt = 0
						failCnt = 0
						Set parentNodes = xmlDOM.getElementsByTagName("arrOrderList")
							For each parentSubNodes in parentNodes
								OutMallOrderSerial = parentSubNodes.getElementsByTagName("orderNum").item(0).text				'주문번호
								CSDetailKey = parentSubNodes.getElementsByTagName("aspOrderNum").item(0).text					'(CS)주문번호
								sndNm = parentSubNodes.getElementsByTagName("sndNm").item(0).text								'구매자명
								sndTelNum = parentSubNodes.getElementsByTagName("sndTelNum").item(0).text						'구매자 전화번호
								sndMobile = parentSubNodes.getElementsByTagName("sndMobile").item(0).text						'구매자 휴대폰
								rcvrNm =  parentSubNodes.getElementsByTagName("rcvrNm").item(0).text							'수취인명
								rcvrTelNum = parentSubNodes.getElementsByTagName("rcvrTelNum").item(0).text						'수취인 전화번호
								rcvrMobile = parentSubNodes.getElementsByTagName("rcvrMobile").item(0).text						'수취인 휴대폰
								rcvrPost = parentSubNodes.getElementsByTagName("rcvrPost").item(0).text							'우편번호
								rcvrAddr1 = parentSubNodes.getElementsByTagName("rcvrAddr1").item(0).text						'주소
								rcvrAddr2 = parentSubNodes.getElementsByTagName("rcvrAddr2").item(0).text						'상세주소
								orderDt = LEFT(parentSubNodes.getElementsByTagName("orderDt").item(0).text, 8)					'주문일 | 년월일시분초(YYYYMMDDhh24miss)
								orderDt = LEFT(orderDt, 4) &"-"& MID(orderDt, 5,2) &"-"& right(orderDt,2)
'								rw "배송희망일 : " & parentSubNodes.getElementsByTagName("dlvrHopeDt").item(0).text				'배송희망일 | 년월일시분초(YYYYMMDDhh24miss)
								orderReqContent = parentSubNodes.getElementsByTagName("orderReqContent").item(0).text			'배송요청사항
'								rw "###############################################################"
								strSql = " select idx from db_temp.dbo.tbl_xSite_TMPCS where SellSite = 'ezwel' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "' "
								rsget.CursorLocation = adUseClient
								rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
								If (Not rsget.Eof) then
									retVal = false
								Else
									retVal = true
								End if
								rsget.Close

								Set Nodes = parentSubNodes.getElementsByTagName("arrOrderGoods")
									For each masterSubNodes in Nodes
										OrgDetailKey = masterSubNodes.getElementsByTagName("orderGoodsNum")(0).Text				'주문순번 / 장바구니 순번
'										rw "출고지 ID : " & masterSubNodes.getElementsByTagName("cspDlvrId")(0).Text			'출고지 ID
'										rw "CP업체 코드 : " & masterSubNodes.getElementsByTagName("cspCd")(0).Text				'CP업체 코드
'										rw "상품코드 : " & masterSubNodes.getElementsByTagName("goodsCd")(0).Text				'상품코드
'										rw "업체상품코드 : " & masterSubNodes.getElementsByTagName("cspGoodsCd")(0).Text		'업체상품코드
'										rw "상품명 : " & masterSubNodes.getElementsByTagName("goodsNm")(0).Text					'상품명
										orderQty = masterSubNodes.getElementsByTagName("orderQty")(0).Text						'주문수량
'										rw "상품옵션 : " & masterSubNodes.getElementsByTagName("optionContent")(0).Text			'상품옵션(구분자 ^) Ex) 색상 및 용량 선택:500GB 레드^'추가 구성상품:선택없음^
'										rw "매입가 : " & masterSubNodes.getElementsByTagName("buyPrice")(0).Text				'상세주소 | 매입가
'										rw "판매가 : " & masterSubNodes.getElementsByTagName("salePrice")(0).Text				'판매가 | 옵션가격포함
'										rw "할인액 : " & masterSubNodes.getElementsByTagName("dccpnPrice")(0).Text				'할인액 | 쿠폰적용시할인된금액
'										rw "택배 운송장 번호 : " & masterSubNodes.getElementsByTagName("dlvrNo")(0).Text		'택배 운송장 번호
'										rw "배송업체 코드 : " & masterSubNodes.getElementsByTagName("dlvrCd")(0).Text			'배송업체 코드 | 별도첨부
'										rw "배송일시 : " & masterSubNodes.getElementsByTagName("dlvrDt")(0).Text				'배송일시 | 년월일시분초(YYYYMMDDhh24miss)
'										rw "옵션가격 : " & masterSubNodes.getElementsByTagName("optionAddPrice")(0).Text		'옵션가격
'										rw "주문상태 : " & masterSubNodes.getElementsByTagName("orderStatus")(0).Text			'주문상태
'										rw "배송비 결제방식 코드 : " & masterSubNodes.getElementsByTagName("dlvrPayCd")(0).Text	'배송비 결제방식 코드 | 선결제(1001), 착불(1002)
'										rw "배송비 : " & masterSubNodes.getElementsByTagName("dlvrPrice")(0).Text				'배송비
'										rw "배송완료일 : " & masterSubNodes.getElementsByTagName("dlvrFinishDt")(0).Text		'배송완료일 | 년월일(YYYYMMDD)
'										cancelDt = masterSubNodes.getElementsByTagName("cancelDt")(0).Text						'주문취소일 | 년월일시분초(YYYYMMDDhh24miss)
'										rw "======================================================"
										strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = 'ezwel' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "') "
										strSql = strSql & " BEGIN "
										strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
										strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
										strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
										strSql = strSql & " 	('" & CStr(getDivCD) & "', '단순변심', 'ezwel', '" & html2db(CStr(OutMallOrderSerial)) & "', '"&CStr(sndNm)&"', '', '"&html2db(CStr(sndTelNum))&"', '"&html2db(CStr(sndMobile))&"', '" & html2db(CStr(rcvrNm)) & "', "
										strSql = strSql & "		'" & html2db(CStr(rcvrTelNum)) & "', '" & html2db(CStr(rcvrMobile)) & "', '" & html2db(CStr(rcvrPost)) & "', '" & html2db(CStr(rcvrAddr1)) & "', '" & html2db(CStr(rcvrAddr2)) & "', '"&html2db(CStr(orderReqContent))&"' "
										strSql = strSql & "		, '" & html2db(CStr(orderDt)) & "', '" & html2db(CStr(OrgDetailKey)) & "', '" & html2db(CStr(CSDetailKey)) & "', " & CStr(orderQty) & ") "
										strSql = strSql & " END "
										''rw strSql
										dbget.Execute(strSql)
				                        If (retVal) Then
				                            succCnt = succCnt + 1
				                        Else
											failCnt = failCnt + 1
				                        End If
									Next
								Set Nodes = nothing
							Next
						Set parentNodes = nothing
					    rw succCnt & "건 입력"
					    rw failCnt & "건 실패"

						If (succCnt > 0) then
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
							strSql = strSql + " 	and c.OutMallOrderSerial = Replace(o.OutMallOrderSerial, '-', '') "
							strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
							strSql = strSql + " where "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.orderserial is NULL "
							strSql = strSql + " 	and o.orderserial is not NULL "
							strSql = strSql + " 	and c.sellsite = 'ezwel' "
							''rw strSql
							dbget.Execute(strSql)

							If getDivCD = "A008" Then
								strSql = " update c "
								strSql = strSql + " set c.currstate = 'B007' "
								strSql = strSql + " from "
								strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
								strSql = strSql + " left join db_temp.dbo.tbl_xSite_TMPOrder o "
								strSql = strSql + " on "
								strSql = strSql + " 	1 = 1 "
								strSql = strSql + " 	and c.SellSite = o.SellSite "
								strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
								strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
								strSql = strSql + " where "
								strSql = strSql + " 	1 = 1 "
								strSql = strSql + " 	and c.orderserial is NULL "
								strSql = strSql + " 	and o.SellSite is NULL "
								strSql = strSql + " 	and c.sellsite = 'ezwel' "
								strSql = strSql + " 	and c.currstate = 'B001' "
								strSql = strSql + " 	and c.divcd = 'A008' "
								''rw strSql
								dbget.Execute(strSql)
							end if
						End If

						If getDivCD = "A008" Then				'주문취소
							Call UpdateLastCSInputDT("ordercancel", date())
						Else									'반품
							Call UpdateLastCSInputDT("return", date())
						End If

						If (getDivCD <> Trim(application("xSiteGetEzwelCS_DIVCD"))) then
							application("xSiteGetEzwelCS_DIVCD") = getDivCD
						End If
					End If
				Set objXML = Nothing
			End If
		Set xmlDOM = Nothing
'		On Error Goto 0
    End If
End If

Function getLastCSInputDT(mode)
	Dim sqlStr
	sqlStr = "select top 1 convert(varchar(10),LastCheckDate,21) as lastCSInputDt"
	sqlStr = sqlStr&" from db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	sqlStr = sqlStr&" where sellsite = 'ezwel' and csGubun = '" & CStr(mode) & "' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If (Not rsget.Eof) Then
		getLastCSInputDT = replace(rsget("lastCSInputDt"), "-", "") & "000000"
	Else
		getLastCSInputDT = "20161020000000"
	End If
	rsget.Close
End Function

Function UpdateLastCSInputDT(mode, dt)
	Dim sqlStr
	sqlStr = " UPDATE db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	sqlStr = sqlStr & " SET LastCheckDate = '" & CStr(dt) & "' "
	sqlStr = sqlStr & " WHERE sellsite = 'ezwel' and csGubun = '" & CStr(mode) & "' "
	dbget.Execute sqlStr
End Function
%>
<% rw "OK" %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
