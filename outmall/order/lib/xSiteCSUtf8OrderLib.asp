<%
Function GetCSOrderCS_lfmall(sellsite, selldate)
	dim xmlURL, strRst, objXML, xmlDOM, strObj
	dim startdate, enddate
	dim retCode, retMsg, items, item, product, productOption, divcd, gubunname, iRbody, SubNodes, Nodes
	dim OutMallOrderSerial, CSDetailKey, OrgDetailKey, dlvTypeGbcd, itemno, OutMallRegDate, iInputCnt
	dim i, j, k, canceDoneStr
	dim strSql, iAssignedRow, successCount
	Dim REQUEST_XML

	startdate = selldate
	enddate = selldate
	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"&VBCRLF
	strRst = strRst & "<OrderInfo>"&VBCRLF
	strRst = strRst & "	<Header>"&VBCRLF
	strRst = strRst & "		<AuthId><![CDATA[tenten]]></AuthId>"&VBCRLF
	strRst = strRst & "		<AuthKey><![CDATA[Ten1010*!!]]></AuthKey>"&VBCRLF
	strRst = strRst & "		<Format>XML</Format>"&VBCRLF
	strRst = strRst & "		<Charset>UTF-8</Charset>"&VBCRLF
	strRst = strRst & "	</Header>"&VBCRLF
	strRst = strRst & "	<Body>"&VBCRLF
	strRst = strRst & "		<Request>"&VBCRLF
'	strRst = strRst & "			<OrdNo>24074677</OrdNo>"&VBCRLF								'주문번호
'	strRst = strRst & "			<OrdDtlNo>29833070</OrdDtlNo>"&VbCRLF						'주문상세일련번호
'	strRst = strRst & "			<OrdererId>orderId</OrdererId>"&VbCRLF						'주문자ID
'	strRst = strRst & "			<RequestGubun>20</RequestGubun>"&VBCRLF						'주문변경요청구분코드 | 20(취소요청), 30(반품요청), 40(교환요청), 50(옵션변경)
'	strRst = strRst & "			<ProcStatusCode>30</ProcStatusCode>"&VbCRLF					'처리상태코드
	strRst = strRst & "			<RequestStartDate>"&Replace(startdate, "-", "")&"</RequestStartDate>"&VbCRLF		'#요청일자(검색)
	strRst = strRst & "			<RequestEndDate>"&Replace(enddate, "-", "")&"</RequestEndDate>"&VbCRLF			'#요청일자(검색)
'	strRst = strRst & "			<CompleteStartDate>20171105</CompleteStartDate>"&VBCRLF		'처리완료일시(검색)
'	strRst = strRst & "			<CompleteEndDate>20171105</CompleteEndDate>"&VbCRLF			'처리완료일시(검색)
	strRst = strRst & "		</Request>"&VBCRLF
	strRst = strRst & "	</Body>"&VBCRLF
	strRst = strRst & "</OrderInfo>"&VBCRLF
	REQUEST_XML = "REQUEST_XML=" & Server.URLEncode(strRst)

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "https://b2b.lfmall.co.kr/interface.do?cmd=getOrderRequestList", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(REQUEST_XML)
		' If session("ssBctID")="kjy8517" Then
		' 	response.write "<textarea cols=100 rows=30>"&strRst&"</textarea>"
		' End If

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
				xmlDOM.LoadXML iRbody
				If session("ssBctID")="kjy8517" Then
					response.write "<textarea cols=100 rows=30>"&iRbody&"</textarea>"
				End If

				retCode = xmlDOM.getElementsByTagName("RequestInfo/Header/ResultCode").item(0).text
				If retCode = "SUCCESS" Then
					If xmlDOM.getElementsByTagName("RequestInfo/Body/Request").length > 0 Then
						Set Nodes = xmlDOM.getElementsByTagName("RequestInfo/Body/Request")
						i = 0
						For each SubNodes in Nodes
							Select Case SubNodes.getElementsByTagName("RequestGubun")(0).Text					'주문변경요청구분코드
								Case "20" 	divcd = "A008"
								Case "30"	divcd = "A004"
								Case "40"	divcd = "A000"
								Case "50"	divcd = "A000"
							End Select
							CSDetailKey			= ""
							OutMallOrderSerial	= SubNodes.getElementsByTagName("OrdNo")(0).Text				'주문번호
							OrgDetailKey		= SubNodes.getElementsByTagName("OrdDtlNo")(0).Text				'주문상세일련번호
							OutMallRegDate		= SubNodes.getElementsByTagName("RequestDate")(0).Text			'요청일자
							OutMallRegDate		= Left(OutMallRegDate,4)&"-"&Mid(OutMallRegDate,5,2)&"-"&Mid(OutMallRegDate,7,2)&" "&Mid(OutMallRegDate,9,2)&":"&Mid(OutMallRegDate,11,2)&":"&Mid(OutMallRegDate,13,2)
							itemno				= SubNodes.getElementsByTagName("RequestQty")(0).text			'요청/처리완료 상품 수량

							If (SubNodes.getElementsByTagName("RequestReasonName").length > 0) Then
								gubunname	= SubNodes.getElementsByTagName("RequestReasonName")(0).Text
							Else
								gubunname	= ""
							End If

							strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' ) "
							strSql = strSql & " BEGIN "
							strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
							strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
							strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
							strSql = strSql & " 	('" & divcd & "', '"& html2db(gubunname) &"', '" & sellsite & "', '" & CStr(OutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
							strSql = strSql & "		'', '', '', '', '', '' "
							strSql = strSql & "		, '"& Replace(OutMallRegDate, "/", "-") &"', '" & CStr(OrgDetailKey) & "', '', " & itemno & ") "
							strSql = strSql & " END "
							dbget.Execute strSql, iAssignedRow


							if (iAssignedRow > 0) then
								iInputCnt = iInputCnt+iAssignedRow
								If divcd = "A008" Then
									''주문 입력 이전 내역은 삭제 하자
									strSql = ""
									strSql = strSql & " UPDATE c "
									strSql = strSql & " SET matchState = 'D'"
									strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder c "
									strSql = strSql & " WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
									strSql = strSql & " and orderserial is NULL"
									dbget.Execute strSql
								End If

								'' CS 마스터정보. 업데이트
								strSql = ""
								strSql = strSql & " UPDATE c "
								strSql = strSql & " SET c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
								strSql = strSql & " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
								strSql = strSql & " , c.OrderName = o.OrderName, c.itemno = o.itemOrderCount "
								strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
								strSql = strSql & " JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
								strSql = strSql & " ON "
								strSql = strSql & " 	1 = 1 "
								strSql = strSql & " 	and c.SellSite = o.SellSite "
								strSql = strSql & " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
								strSql = strSql & " 	and c.OrgDetailKey = o.OrgDetailKey "
								strSql = strSql & " WHERE "
								strSql = strSql & " 	1 = 1 "
								strSql = strSql & " 	and c.orderserial is NULL "
								strSql = strSql & " 	and o.orderserial is not NULL "
								strSql = strSql & " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
								''response.write strSql & "<br />"
								dbget.Execute strSql

								If divcd = "A008" Then
									strSql = ""
									strSql = strSql & " UPDATE c "
									strSql = strSql & " SET c.currstate = 'B007' "
									strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS c "
									strSql = strSql & " LEFT JOIN db_temp.dbo.tbl_xSite_TMPOrder o "
									strSql = strSql & " ON "
									strSql = strSql & "		1 = 1 "
									strSql = strSql & "		and c.SellSite = o.SellSite "
									strSql = strSql & "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
									strSql = strSql & "		and c.OrgDetailKey = o.OrgDetailKey "
									strSql = strSql & " WHERE "
									strSql = strSql & "		1 = 1 "
									strSql = strSql & "		and c.orderserial is NULL "
									strSql = strSql & "		and o.SellSite is NULL "
									strSql = strSql & "		and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(OrgDetailKey) & "' "
									strSql = strSql & "		and c.currstate = 'B001' "
									strSql = strSql & "		and c.divcd = 'A008' "
									''rw strSql
									dbget.execute strSql
								End If
							End If

'							rw SubNodes.getElementsByTagName("OrdererId")(0).Text				'주문자ID
'							rw SubNodes.getElementsByTagName("OrdererName")(0).Text				'주문자이름
'							rw SubNodes.getElementsByTagName("RequestReasonCode")(0).Text		'변경사유코드
'							rw SubNodes.getElementsByTagName("ProcStatusCode")(0).Text			'처리상태코드
							i = i + 1
						Next
					End If
				End If
			Set xmlDOM = nothing
		End If
	Set objXML = nothing
End Function

function GetCSCheckStatus(byVal sellsite, byVal csGubun, byRef LastCheckDate, byRef isSuccess)
	dim strSql

    strSql = " IF NOT Exists("
    strSql = strSql + " 	select LastcheckDate"
    strSql = strSql + " 	from db_temp.[dbo].[tbl_xSite_TMPCS_timestamp]"
    strSql = strSql + " 	where sellsite='" + CStr(sellsite) + "' and csGubun='" + CStr(csGubun) + "'"
	strSql = strSql + " )"
	strSql = strSql + " BEGIN"
	strSql = strSql + "		insert into db_temp.[dbo].[tbl_xSite_TMPCS_timestamp](sellsite, csGubun, lastcheckdate, issuccess, LastUpdate) "
	strSql = strSql + "		values('" & sellsite & "', '" & csGubun & "', '" & Left(DateAdd("d", -1, Now()), 10) & "', 'N', getdate()) "
	strSql = strSql + " END"
	dbget.Execute strSql

	strSql = " select convert(varchar(10), LastCheckDate, 121) as LastCheckDate, isSuccess from db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	strSql = strSql + " 	where sellsite='" + CStr(sellsite) + "' and csGubun='" + CStr(csGubun) + "'"

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		LastCheckDate = rsget("LastCheckDate")
		isSuccess = rsget("isSuccess")
	rsget.Close
end function

function SetCSCheckStatus(sellsite, csGubun, LastCheckDate, isSuccess)
	dim strSql

	strSql = " update db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	strSql = strSql + " set lastcheckdate = '" & LastCheckDate & "', issuccess = '" & isSuccess & "', LastUpdate = getdate() "
	strSql = strSql + " where sellsite = '" + CStr(sellsite) + "' and csGubun='" + CStr(csGubun) + "' "
	''response.write strSql
	dbget.Execute strSql
end function

function fnMatchCs(sellsite, ioutmallorderserial)
    dim affectedRow, strSql

    strSql = " update T "
	strSql = strSql & " set T.asid = a.id "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_list] a "
	strSql = strSql & " 	on "
	strSql = strSql & " 		1 = 1 "
	strSql = strSql & " 		and a.orderserial = T.OrderSerial "
	strSql = strSql & " 		and a.deleteyn = 'N' "
	strSql = strSql & " 		and ( "
	strSql = strSql & " 			(T.divcd = 'A004' and a.divcd in ('A004', 'A010', 'A008', 'A011', 'A012', 'A112', 'A112')) "
	strSql = strSql & " 			or "
	strSql = strSql & " 			(T.divcd = 'A011' and a.divcd in ('A011', 'A012', 'A112', 'A112')) "
    strSql = strSql & " 			or "
    strSql = strSql & " 			(T.divcd = 'A000' and a.divcd in ('A000', 'A100')) "
    strSql = strSql & " 			or "
    strSql = strSql & " 			(T.divcd = 'A008' and a.divcd in ('A008', 'A004', 'A010')) "
	strSql = strSql & " 		) "
	strSql = strSql & " 		and a.id not in ( "
	strSql = strSql & " 			select T.asid "
	strSql = strSql & " 			from "
	strSql = strSql & " 				[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " 			where "
	strSql = strSql & " 				1 = 1 "
	strSql = strSql & " 				and T.SellSite = '" & sellsite & "' "
	strSql = strSql & " 				and T.OutMallOrderSerial = '" & ioutmallorderserial & "' "
	strSql = strSql & " 				and T.asid is not NULL "
	strSql = strSql & " 		) "
	strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_detail] d "
	strSql = strSql & " 	on "
	strSql = strSql & " 		1 = 1 "
	strSql = strSql & " 		and a.id = d.masterid "
	strSql = strSql & " 		and d.itemid = T.ItemID "
	strSql = strSql & " 		and d.itemoption = T.itemoption "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
	strSql = strSql & " 	and T.OutMallOrderSerial = '" & ioutmallorderserial & "' "
	strSql = strSql & " 	and T.asid is NULL "
    strSql = strSql & " 	and IsNull(T.outmallCurrState, 'B001') <> 'B008' "
    dbget.Execute strSql, affectedRow

    '// 이중매칭 삭제
    if (sellsite = "coupang") then
        strSql = " update A "
        strSql = strSql & " set A.asid = NULL "
        strSql = strSql & " from "
        strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] A "
        strSql = strSql & " 	join [db_temp].[dbo].[tbl_xSite_TMPCS] B "
        strSql = strSql & " 	on "
        strSql = strSql & " 		1 = 1 "
        strSql = strSql & " 		and A.SellSite = B.SellSite "
        strSql = strSql & " 		and A.OutMallOrderSerial = B.OutMallOrderSerial "
        strSql = strSql & " 		and A.OrgDetailKey = B.OrgDetailKey "
        strSql = strSql & " 		and A.CSDetailKey <> B.CSDetailKey "
        strSql = strSql & " 		and A.asid = B.asid "
        strSql = strSql & " where "
        strSql = strSql & " 	1 = 1 "
        strSql = strSql & " 	and A.SellSite = '" & sellsite & "' "
        strSql = strSql & " 	and A.OutMallOrderSerial = '" & ioutmallorderserial & "' "
        strSql = strSql & " 	and A.OutMallCurrState = 'B008' "
        dbget.Execute strSql
    end if

    fnMatchCs = affectedRow
end function

function fnUnmatchDeletedCS(sellsite, ioutmallorderserial)
    dim affectedRow, strSql

    strSql = " update T "
	strSql = strSql & " set T.asid = NULL "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_list] a on T.asid = a.id and a.deleteyn = 'Y' "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
	strSql = strSql & " 	and T.OutMallOrderSerial = '" & ioutmallorderserial & "' "
	strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008', 'A000') "
	strSql = strSql & " 	and T.orderserial is not NULL "
    dbget.Execute strSql

    strSql = " update T "
	strSql = strSql & " set T.asid = NULL "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
	strSql = strSql & " 	and T.OutMallOrderSerial = '" & ioutmallorderserial & "' "
	strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008', 'A000') "
	strSql = strSql & " 	and T.orderserial is not NULL "
    strSql = strSql & " 	and T.asid is not NULL "
    strSql = strSql & " 	and T.outmallCurrState = 'B008' "
    dbget.Execute strSql
end function

function MatchTenCSAsid(sellsite)
    dim strSql, affectedRows
    dim OutMallOrderSerialArr, OutMallOrderSerial

    strSql = " update T "
	strSql = strSql & " set T.asid = NULL "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
	''strSql = strSql & " 	and T.OutMallOrderSerial = '" & ioutmallorderserial & "' "
	strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008', 'A000') "
	strSql = strSql & " 	and T.orderserial is not NULL "
    strSql = strSql & " 	and T.asid is not NULL "
    strSql = strSql & " 	and T.outmallCurrState = 'B008' "
    dbget.Execute strSql

    strSql = " select distinct top 100 T.OutMallOrderSerial "
	strSql = strSql & " from "
	strSql = strSql & " [db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
	strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008', 'A000') "
    strSql = strSql & " 	and T.orderserial is not NULL "
	strSql = strSql & " 	and T.asid is NULL "
    ''strSql = strSql & " 	and T.regdate < convert(varchar(10), getdate(), 121) "
    strSql = strSql & " 	and T.regdate >= DateAdd(day, -90, getdate()) "
    strSql = strSql & " 	and IsNull(T.asidCheckDT, DateAdd(day, -1, getdate())) < DateAdd(hour, -1, getdate()) "
    ''strSql = strSql & " order by newid() "
    ''rw strSql

    OutMallOrderSerialArr = ""

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
        do until rsget.eof
            OutMallOrderSerialArr = OutMallOrderSerialArr + "," + rsget("OutMallOrderSerial")
            rsget.moveNext
        loop
	end if
	rsget.Close

    Response.CharSet="euc-kr"
    Session.codepage="949"
    Response.codepage="949"
    Response.ContentType="text/html;charset=euc-kr"

    '// git 업로드 확인
    if OutMallOrderSerialArr = "" then
        rw "내역없음"
        dbget.close() : response.end
    end if

    affectedRows = 0
    OutMallOrderSerialArr = Split(OutMallOrderSerialArr, ",")
    for i = 0 to UBound(OutMallOrderSerialArr)
        OutMallOrderSerial = OutMallOrderSerialArr(i)
        if OutMallOrderSerial <> "" then
            affectedRows = fnMatchCs(sellsite, OutMallOrderSerial)
            Call fnUnmatchDeletedCS(sellsite, OutMallOrderSerial)

            rw OutMallOrderSerial & " : " & affectedRows & " 건 반영됨"

            if affectedRows = 0 then
                strSql = " update T "
                strSql = strSql & " set T.asidCheckDT = getdate() "
	            strSql = strSql & " from "
	            strSql = strSql & " [db_temp].[dbo].[tbl_xSite_TMPCS] T "
	            strSql = strSql & " where "
	            strSql = strSql & " 	1 = 1 "
	            strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
                strSql = strSql & " 	and T.OutMallOrderSerial = '" & OutMallOrderSerial & "' "
	            strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008', 'A000') "
	            strSql = strSql & " 	and T.asid is NULL "
                dbget.Execute strSql
            end if
        end if
    next
end function

function CheckExtCsState(sellsite)
    dim strSql, affectedRows
    dim divcd, csdetailkey1, csdetailkey2
    dim divcdArr, csdetailkey1Arr, csdetailkey2Arr

    strSql = " select distinct top 100 "
    select case sellsite
        case "coupang"
            strSql = strSql & " T.divcd, T.OutMallOrderSerial as csdetailkey1, T.CSDetailKey as csdetailkey2 "
        case else
            strSql = strSql & " T.divcd, T.OutMallOrderSerial as csdetailkey1, '' as csdetailkey2 "
    end select

	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	''strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_list] a on T.asid = a.id and a.currstate = 'B007' and a.deleteyn = 'N' "		'// 접수이전 내역도 체크
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "

    if (sellsite = "coupang") then
        strSql = strSql & " 	and T.divcd in ('A004') "
    else
        strSql = strSql & " 	and T.divcd in ('A004', 'A011') "
    end if

	strSql = strSql & " 	and T.orderserial is not NULL "
	''strSql = strSql & " 	and T.regdate < convert(varchar(10), DateAdd(day, -0, getdate()), 121) "
	strSql = strSql & " 	and T.regdate >= convert(varchar(10), DateAdd(day, -80, getdate()), 121) "
	strSql = strSql & " 	and IsNull(T.outmallCheckDT, DateAdd(day, -1, getdate())) < DateAdd(hour, -1, getdate()) "
	strSql = strSql & " 	and IsNull(T.OutMallCurrState, 'B001') < 'B007' "
    ''rw strSql

    csdetailkey1Arr = ""
    csdetailkey2Arr = ""
    divcdArr = ""

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
        do until rsget.eof
            divcdArr = divcdArr & rsget("divcd") & ","
            csdetailkey1Arr = csdetailkey1Arr & rsget("csdetailkey1") & ","
            csdetailkey2Arr = csdetailkey2Arr & rsget("csdetailkey2") & ","
            rsget.moveNext
        loop
	end if
	rsget.Close

    Response.CharSet="euc-kr"
    Session.codepage="949"
    Response.codepage="949"
    Response.ContentType="text/html;charset=euc-kr"

    ''Call GetCSOrderCancel_One_coupang("coupang", "10000060357304", "184259720")

    '// git 업로드 확인
    if divcdArr = "" then
        rw "내역없음"
        dbget.close() : response.end
    end if

    affectedRows = 0
    csdetailkey1Arr = Split(csdetailkey1Arr, ",")
    csdetailkey2Arr = Split(csdetailkey2Arr, ",")
    divcdArr = Split(divcdArr, ",")

    for i = 0 to UBound(divcdArr)
        csdetailkey1 = csdetailkey1Arr(i)
        csdetailkey2 = csdetailkey2Arr(i)
        divcd = divcdArr(i)

        if Trim(divcd) <> "" then
        	select case sellsite
                case "coupang"
                    if divcd = "A004" then
                        affectedRows = GetCSOrderCancel_One_coupang(sellsite, csdetailkey1, csdetailkey2)
                        rw csdetailkey1 & " : " & affectedRows & " 건 반영됨"

                        if affectedRows = 0 then
                            strSql = " update T "
                            strSql = strSql & " set T.outmallCheckDT = getdate() "
	                        strSql = strSql & " from "
	                        strSql = strSql & " [db_temp].[dbo].[tbl_xSite_TMPCS] T "
	                        strSql = strSql & " where "
	                        strSql = strSql & " 	1 = 1 "
	                        strSql = strSql & " 	and T.SellSite = '" & sellsite & "' "
                            strSql = strSql & " 	and T.OutMallOrderSerial = '" & csdetailkey1 & "' "
	                        strSql = strSql & " 	and T.divcd = '" & divcd & "' "
	                        strSql = strSql & " 	and T.asid is not NULL "
                            dbget.Execute strSql
                        end if
                    end if
                case else
                    response.write "TEST2<br />"
            end select
        end if
    next
end function

%>
