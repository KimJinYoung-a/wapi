<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/lotteCom/inc_dailyAuthCheck.asp"-->
<% if NOT (IsAutoScript) then %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body bgcolor="#F4F4F4" >
<% end if %>
<%


	'// 변수선언
	dim LotteGoodNo, LotteStatCd
	dim strSql, actCnt, lp
    dim AssignedCNT

	actCnt = 0			'실갱신건수

Dim proc_gubun : proc_gubun = "sfin"
Dim ord_no     : ord_no     = request("ord_no")
Dim ord_dtl_sn : ord_dtl_sn = request("ord_dtl_sn")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = request("inv_no")
Dim paramAdd
Dim tenOrderSerial,orgsubtotalprice ,minusOrderSerial ,minusSubtotalprice


Dim ORG_ord_no : ORG_ord_no = ord_no

ord_no = replace(ord_no,":01","")  '''2011-11-30-4100282:01
ord_no = replace(ord_no,"_1","")  ''2012-03-11-8159343_1
ord_no = replace(ord_no,"_2","")
ord_no = replace(ord_no,"_3","")
ord_no = replace(ord_no,"_4","")

if (inv_no="보코통 순면 오가닉 원형패드") then inv_no="기타"
paramAdd = "&ord_no="+Replace(ord_no,"-","")
paramAdd = paramAdd + "&ord_dtl_sn="+ord_dtl_sn
paramAdd = paramAdd + "&proc_gubun="+proc_gubun
paramAdd = paramAdd + "&hdc_cd="+hdc_cd
paramAdd = paramAdd + "&inv_no="+server.UrlEncode(replace(inv_no,"-",""))

''rw lotteAPIURL & "/openapi/registDeliver.lotte?subscriptionId=" & lotteAuthNo & paramAdd
'response.end

Dim iResult, iMessage
Dim objXML, xmlDOM
    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
    objXML.Open "GET", lotteAPIURL & "/openapi/registDeliver.lotte?subscriptionId=" & lotteAuthNo & paramAdd, false
    objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objXML.Send()
    If objXML.Status = "200" Then

    	'//전달받은 내용 확인
    	'Response.contentType = "text/xml; charset=euc-kr"
    	'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
    	'response.End

    	'XML을 담을 DOM 객체 생성
    	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
    	xmlDOM.async = False
    	'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
''rw "objXML.ResponseBody="&BinaryToText(objXML.ResponseBody, "euc-kr")
 ''response.end

        ''결과가 빈값 (20120702) 으로 왔음.. :: 성공시?
        IF (Trim(objXML.ResponseBody)="") THEN
            rw "결과 빈값"
            iResult="1"
        ELSE
        	xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

        	if Err<>0 then
        		Response.Write "<script language=javascript>alert('롯데닷컴 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
        		dbget.Close: Response.End
        	end if
            On Error Resume next
        	iResult		= Trim(xmlDOM.getElementsByTagName("Result").item(0).text)			'결과
            	If ERR THEN
            	    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
            	    iMessage2 = xmlDOM.getElementsByTagName("Message").Item(0).Text
            	ENd IF
        	On Error Goto 0
        END IF
    '	'// 수정

        IF (iResult="1") THEN
        	strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
        	strSql = strSql & "	Set sendState=1"
        	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
            strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
            strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
            strSql = strSql & "	and matchstate in ('O')"
       ''rw strSql
        	dbget.Execute strSql,AssignedCNT
        	actCnt = actCnt+AssignedCNT

        	IF (AssignedCNT>0) then
        	    if (IsAutoScript) then
        	        rw "OK|"&ord_no&" "&ord_dtl_sn
        	    ELSE
            	    response.write "OK"
            	ENd IF
            ENd IF
        ELSE
            ''이미 송장이 입력된경우.// 무조건 N번 재시도로 변경.
            ''송장 늦게 입력 : 주문일로부터 30일이 경과하여 발송완료 불가합니다. 고객센터에 문의하시기 바랍니다. 문의처) 02-3289-3236
            ''취소된경우 : 발송완료 처리할 주문이 존재하지 않습니다.
            ''Maybe 고객수령 : 현재 주문내역상태가 발송완료를 처리할 수 없습니다
            IF LEft(iMessage,Len("현재 주문내역상태가 발송완료를 처리할 수 없습니다"))="현재 주문내역상태가 발송완료를 처리할 수 없습니다" THEN
                strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
            	strSql = strSql & "	Set sendState=9"
            	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
                strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
                strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
                strSql = strSql & "	and matchstate in ('O','C','Q','A')"
''rw  strSql
            	dbget.Execute strSql,AssignedCNT

            	if (IsAutoScript) then
            	    rw "SKIP 처리 "&ord_no&" "&ord_dtl_sn
            	ELSE
            	    rw "SKIP 처리"
            	ENd IF

            ELSE
                '' 시도 회수 추가 sendReqCnt
                strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
            	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
                strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
                strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
                strSql = strSql & "	and matchstate in ('O','C','Q','A')"

            	dbget.Execute strSql

                iMessage = "<font color=red>"&iMessage&"</font>"
            END IF

            if (IsAutoScript) then
                rw "iMessage="&iMessage&":"&ord_no&" "&ord_dtl_sn
            else
                rw "iMessage="&iMessage
            ENd IF

        END IF
    	Set xmlDOM = Nothing



    	IF (iResult<>"1") and (Not IsAutoScript) THEN
    	    ''
    	    strSql = "select * from db_temp.dbo.tbl_xSite_TMPOrder "
    	    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
            strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
            strSql = strSql & "	and sendState=0"

            rsget.CursorLocation = adUseClient
            rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
            if Not rsget.Eof then
                tenOrderSerial = rsget("Orderserial")
            end if
            rsget.Close

            if (tenOrderSerial<>"") then
                strSql = "select top 10 m1.orderserial,m1.subtotalprice,m2.orderserial as minusOrderSerial,m2.subtotalprice as minusSubtotalprice"
                strSql = strSql & "	from db_order.dbo.tbl_order_master m1"
                strSql = strSql & "		left join db_order.dbo.tbl_order_master m2"
                strSql = strSql & "		on m2.linkorderserial='"&tenOrderSerial&"'"
                strSql = strSql & "		and m2.jumundiv='9'"
                strSql = strSql & "		and m2.cancelyn='N'"
                strSql = strSql & "	where m1.orderserial='"&tenOrderSerial&"'"
                rsget.CursorLocation = adUseClient
	            rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
                if Not rsget.Eof then
                    orgsubtotalprice    = rsget("subtotalprice")
                    minusOrderSerial    = rsget("minusOrderSerial")
                    minusSubtotalprice  = rsget("minusSubtotalprice")
                end if
                rsget.Close
            end if

            if (minusOrderSerial<>"") then

                rw "결제 금액 : " & FormatNumber(orgsubtotalprice,0) & " / " & FormatNumber(minusSubtotalprice,0)
                rw "<br>주문번호 :" & tenOrderSerial & " / " & minusOrderSerial
                rw "<input type='button' value='완료처리' onClick=""finCancelOrd('"&ORG_ord_no&"','"&ord_dtl_sn&"')"">"
                response.write VbCRLF
                response.write "<script language='javascript'>"&VbCRLF
                response.write "function finCancelOrd(ORG_ord_no,ord_dtl_sn){"&VbCRLF
                response.write "    var uri = 'LotteComProc.asp?mode=etcSongjangFin&ORG_ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn;"&VbCRLF
                response.write "    var popwin = window.open(uri,'finCancelOrd','width=200,height=200');"&VbCRLF
                response.write "    popwin.focus()"&VbCRLF
                response.write "}"&VbCRLF
                response.write "</script>"&VbCRLF
'2013/02/28 김진영 Else부분 추가
'만약 에러횟수가 3회가 넘으면서 minusOrderSerial이 공백일 때 해당
'updateSendState = 901		발송처리누락 수기등록건
'updateSendState = 902		취소후 제결제건
'updateSendState = 903		반품처리건
			Else
				Dim errCount
				strSql = ""
				strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
				strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
				strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
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
					response.write "<input type='button' value='완료처리' onClick=""finCancelOrd2('"&ORG_ord_no&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
					response.write "<script language='javascript'>"&VbCRLF
					response.write "function finCancelOrd2(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
					response.write "    if(selectValue == ''){"&VbCRLF
					response.write "    	alert('선택해주세요');"&VbCRLF
					response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
					response.write "    	return;"&VbCRLF
					response.write "    }"&VbCRLF
					response.write "    var uri = 'LotteComProc.asp?mode=updateSendState&ORG_ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
	                response.write "    var popwin = window.open(uri,'finCancelOrd2','width=200,height=200');"&VbCRLF
	                response.write "    popwin.focus()"&VbCRLF
					response.write "}"&VbCRLF
					response.write "</script>"&VbCRLF
				End If
'2013/02/28 김진영 else부분 추가 끝
            end if
    	end if
    else
        if (IsAutoScript) then
            rw "롯데닷컴과 통신중에 오류가 발생했습니다. "&ord_no&" "&ord_dtl_sn
        else
    	    Response.Write "<script language=javascript>alert('롯데닷컴과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
    	ENd IF
    	dbget.Close: Response.End
    end if
    Set objXML = Nothing
%>
<% if NOT (IsAutoScript) then %>
</body>
</html>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->