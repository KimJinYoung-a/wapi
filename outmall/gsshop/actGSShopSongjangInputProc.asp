<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<%

''dbget.close() : response.end

''Response.CharSet = "UTF-8"
Response.CharSet = "euc-kr"

Function XMLSend(url, xmlStr)
	Dim poster, SendDoc, retDoc, buf, retXML, objLst, i
	'Set SendDoc = server.createobject("MSXML2.DomDocument.3.0")
	'	SendDoc.async = False
	'	SendDoc.LoadXML(xmlStr)

	Set poster = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		poster.open "POST", url, false
		''poster.setRequestHeader "CONTENT_TYPE", "text/xml"

		''poster.setRequestHeader "Content-Type", "application/xml; charset=utf-8"
		''poster.setRequestHeader "Accept", "application/xml"

		''poster.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
		poster.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"

		''poster.setRequestHeader "Content-Type", "application/xml; charset=utf-8"
		''poster.SetRequestHeader "Accept", "application/xml; charset=utf-8"

		''poster.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"

		''poster.setTimeouts 5000,90000,90000,90000  ''2013/07/22 추가
		''poster.send SendDoc
		poster.send xmlStr
		''poster.send Server.URLEncode(xmlStr)

		If poster.Status = "200" Then
			''XMLSend = poster.ResponseBody
			XMLSend = poster.ResponseText
		else
			XMLSend = "22"
		end if


	''XMLSend = poster.responseTEXT

	Set SendDoc = Nothing
	Set poster = Nothing
End Function

function Format00(val)
	Format00 = Right("00" & val,2)
end Function

function getGSShopSongjangXMLStr(ordclmNo, ordSeq, delvEntrNo, invoNo)
    Dim strRst

	dim currDateStr
	dim yyyy, mm, dd, hh, mi, ss

	dim ordNo, ordItemNo
	'2015-09-17 김진영 하단 If문 추가
	If Ubound(Split(ordclmNo,"_")) > 0 Then
		ordNo = Split(ordclmNo,"_")(0)
		ordNo = Right(("0000000000" & ordNo), 10)
	Else
		ordNo = Right(("0000000000" & ordclmNo), 10)
	End If
	'ordNo = Right(("0000000000" & ordclmNo), 10)
	'ordNo = Right(("0000000000" & ordNo), 10)		''2015-09-17 김진영 ordclmNo -> ordNo로 교체

	ordItemNo = Right(("000000" & ordSeq), 5) + "0"

	yyyy = Year(Now())
	mm = Format00(Month(Now()))
	dd = Format00(Day(Now()))
	hh = Format00(Hour(Now()))
	mi = Format00(Minute(Now()))
	ss = Format00(Second(Now()))

	currDateStr = yyyy & mm & dd & hh & mi & ss

	strRst = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
	strRst = strRst &"<DeliveryStatus_V01_00>" & vbCrLf
	strRst = strRst &"<MessageHeader>" & vbCrLf
	strRst = strRst &"        <Sender>10X10</Sender>" & vbCrLf
	strRst = strRst &"        <Receiver>GS SHOP</Receiver>" & vbCrLf
	strRst = strRst &"        <MessageID>ORDINFRST-" + CStr(currDateStr) + "</MessageID>" & vbCrLf
	strRst = strRst &"        <DateTime>" + CStr(currDateStr) + "</DateTime>" & vbCrLf
	strRst = strRst &"        <ProcessType>C</ProcessType>" & vbCrLf
	strRst = strRst &"        <DocumentID>DLVINF</DocumentID>" & vbCrLf
	strRst = strRst &"        <UniqueID>DLVINF-" + CStr(currDateStr) + "</UniqueID>" & vbCrLf
	strRst = strRst &"        <ErrorOccur></ErrorOccur>" & vbCrLf
	strRst = strRst &"        <ErrorMessage></ErrorMessage>" & vbCrLf
	strRst = strRst &"</MessageHeader>" & vbCrLf
	strRst = strRst &"<MessageBody>" & vbCrLf
	strRst = strRst &"        <OrderStatus>" & vbCrLf
	strRst = strRst &"                <ordNo>" + CStr(ordNo) + "</ordNo>" & vbCrLf
	strRst = strRst &"                <ordItemNo>" + CStr(ordItemNo) + "</ordItemNo>" & vbCrLf
	strRst = strRst &"                <deliveryCd>" + CStr(delvEntrNo) + "</deliveryCd>" & vbCrLf
	strRst = strRst &"                <deliveryNo>" + CStr(invoNo) + "</deliveryNo>" & vbCrLf
	strRst = strRst &"                <cmpulDlv></cmpulDlv>" & vbCrLf
	strRst = strRst &"        </OrderStatus>" & vbCrLf
	strRst = strRst &"</MessageBody>" & vbCrLf
	strRst = strRst &"</DeliveryStatus_V01_00>"

	''response.write strRst

    getGSShopSongjangXMLStr = strRst
end function

dim entrId      : entrId="10X10"                    ''고정값 : 제휴업체_ID
dim ordclmNo    : ordclmNo=request("ordclmNo")      ''인터파크 주문번호
dim ordSeq      : ordSeq=request("ordSeq")          ''인터파크 주문순번
dim delvDt      : delvDt=request("delvDt")          ''YYYYMMDD 출고완료일자
dim delvEntrNo  : delvEntrNo=request("delvEntrNo")  ''택배사코드
dim invoNo      : invoNo=request("invoNo")          ''운송장번호 숫자만 가능함.
dim reqXML
dim errCount

invoNo= trim(replace(invoNo,"-",""))


'2013/02/28 진영추가
dim mode      : mode=request("mode")
If mode = "updateSendState" Then
	Dim sqlStr, AssignedRow
	sqlStr = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr & "	Set sendState='"&request("updateSendState")&"'"
	sqlStr = sqlStr & "	,sendReqCnt=sendReqCnt+1"
	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ordclmNo")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ordSeq")&"'"
	sqlStr = sqlStr & "	and sellsite='gseshop'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If
'2013/02/28 진영추가 끝

reqXML = getGSShopSongjangXMLStr(ordclmNo, ordSeq, delvEntrNo, invoNo)
''response.write "<html><body><textarea rows=20 cols=80>"&reqXML&"</textarea></body></html>"
''response.end


'// 변수선언
dim strSql, actCnt, lp
dim AssignedCNT

actCnt = 0			'실갱신건수

Dim iResult, iMessage
Dim iParkURL, replyXML, ErrMsg


''[운  영] http://ecb2b.gsshop.com/aliaSupCommonReceiveOrderInfo.gs
''[테스트] http://test1.gsshop.com/aliaSupCommonReceiveOrderInfo.gs
iParkURL = "http://ecb2b.gsshop.com/aliaSupCommonReceiveOrderInfo.gs"
''iParkURL = "http://test1.gsshop.com/aliaSupCommonReceiveOrderInfo.gs"


''response.write "aa" & iParkURL
''response.write "bb" & reqXML
dim retDoc
retDoc = xmlSend(iParkURL, reqXML)


''response.write CStr(retDoc)
''response.end


    Select Case CStr(retDoc)
    	Case "E", "S", "Y"
    		ErrMsg = ""
    	Case Else
    		ErrMsg = "ERROR"
    end Select

    if (ErrMsg<>"") then
        if (IsAutoScript) then
            rw "GS샵 송장입력중  오류가 발생했습니다. "&ordclmNo&" "&ordclmNo&"_"&ordSeq
        else
    	    Response.Write "<script language=javascript>alert('GS샵 송장입력중  오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
			rw ErrMsg
    	ENd IF

    	'' 시도 회수 추가 sendReqCnt
            strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
            strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
            strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
            strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
            strSql = strSql & "	and sellsite='gseshop'"
            strSql = strSql & "	and matchstate in ('O','C','Q','A')"
            dbget.Execute strSql

			strSql = ""
			strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
			strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
			strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
			strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not rsget.Eof Then
				errCount = rsget("cnt")
			End If
			rsget.Close

			If errCount > 0 Then
				response.write  "<select name='updateSendState' id=""updateSendState"">" &_
				"	<option value=''>선택</option>" &_
				"	<option value='901'>발송처리누락 수기등록건</option>" &_
				"	<option value='902'>취소후 제결제건</option>" &_
				"	<option value='903'>반품처리건</option>" &_
				"</select>&nbsp;&nbsp;"
				response.write "<input type='button' value='완료처리' onClick=""finCancelOrd2('"&ordclmNo&"','"&ordSeq&"',document.getElementById('updateSendState').value)"">"
				response.write "<script language='javascript'>"&VbCRLF
				response.write "function finCancelOrd2(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
				response.write "    if(selectValue == ''){"&VbCRLF
				response.write "    	alert('선택해주세요');"&VbCRLF
				response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
				response.write "    	return;"&VbCRLF
				response.write "    }"&VbCRLF
				response.write "    var uri = 'actGSShopSongjangInputProc.asp?mode=updateSendState&ordclmNo='+ORG_ord_no+'&ordSeq='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
				response.write "    location.replace(uri);"&VbCRLF
				response.write "}"&VbCRLF
				response.write "</script>"&VbCRLF
			end if

    	''dbget.Close: Response.End

    else

    	'XML을 담을 DOM 객체 생성
    	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
    	xmlDOM.async = False
    	'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
''rw "objXML.ResponseBody="&BinaryToText(objXML.ResponseBody, "euc-kr")
 ''response.end

 		'// S 는 성공이 아닌듯
        IF (CStr(retDoc) = "Y") THEN
            rw "성공"

        	strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
        	strSql = strSql & "	Set sendState=1"
        	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
            strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
            strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
            strSql = strSql & "	and sellsite='gseshop'"
            strSql = strSql & "	and matchstate in ('O')"
       		''rw strSql
        	dbget.Execute strSql,AssignedCNT
        	actCnt = actCnt+AssignedCNT

			iMessage = "성공"
        ELSE
            '' 시도 회수 추가 sendReqCnt
            strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
            strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
            strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
            strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
            strSql = strSql & "	and sellsite='gseshop'"
            strSql = strSql & "	and matchstate in ('O','C','Q','A')"
            dbget.Execute strSql

            iMessage = "<font color=red>ERROR</font>"

			strSql = ""
			strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
			strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
			strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
			strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not rsget.Eof Then
				errCount = rsget("cnt")
			End If
			rsget.Close

			If errCount > 0 Then
				response.write  "<select name='updateSendState' id=""updateSendState"">" &_
								"	<option value=''>선택</option>" &_
								"	<option value='901'>발송처리누락 수기등록건</option>" &_
								"	<option value='902'>취소후 제결제건</option>" &_
								"	<option value='903'>반품처리건</option>" &_
								"</select>&nbsp;&nbsp;"
				response.write "<input type='button' value='완료처리' onClick=""finCancelOrd2('"&ordclmNo&"','"&ordSeq&"',document.getElementById('updateSendState').value)"">"
				response.write "<script language='javascript'>"&VbCRLF
				response.write "function finCancelOrd2(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
				response.write "    if(selectValue == ''){"&VbCRLF
				response.write "    	alert('선택해주세요');"&VbCRLF
				response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
				response.write "    	return;"&VbCRLF
				response.write "    }"&VbCRLF
				response.write "    var uri = 'actGSShopSongjangInputProc.asp?mode=updateSendState&ordclmNo='+ORG_ord_no+'&ordSeq='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
				response.write "    location.replace(uri);"&VbCRLF
				response.write "}"&VbCRLF
				response.write "</script>"&VbCRLF
			end if
        END IF

        if (IsAutoScript) then
            rw "iMessage="&iMessage&":"&ordclmNo&" "&ordclmNo&"_"&ordSeq
        else
            rw "iMessage="&iMessage
        ENd IF


    end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
