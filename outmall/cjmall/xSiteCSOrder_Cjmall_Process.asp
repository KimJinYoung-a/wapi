<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 900 %>
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/outmall/cjmall/cjmallitemcls_forcsorder.asp"-->
<!-- #include virtual="/outmall/cjmall/incCJmallFunction_CsOrder.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),16)
Dim sday : sday = request("sday")
Dim cksel : cksel = request("cksel")
Dim subcmd : subcmd = request("subcmd")
Dim iitemid, ret, sqlStr, AssignedRow, i
Dim alertMsg, ierrStr
Dim SuccCNT, FailCNT
dim todate, stdt, maxloop


If (cmdparam="cjmallOrdreg") Then ''주문목록 조회

    todate = LEFT(CStr(now()),10)
    maxloop = 10
    stdt = getLastOrderInputDT()
    sday = stdt
    for i=0 to maxloop
        rw sday & "주문건 등록시작 ======================================"
    	call getCjOrderList("ORDLIST", sday)
    	rw sday & "주문취소건 등록시작 ======================================"
    	call getCjOrderList("ORDCANCELLIST", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next
ElseIf (cmdparam="cjmallCsreg1") Then ''CS목록 조회(반품)
    todate = LEFT(CStr(now()),10)
    maxloop = 15					'// TODO : 15일 이상 CS가 없으면 그 이후 CS내역을 가져오지 못한다.

	'// ========================================================================
    stdt = getLastCSInputDT("return")
	''rw stdt
    sday = stdt
    for i=0 to maxloop
        rw sday & " CS건[회수내역] 조회 등록시작 ======================================"
    	call getCjCsList("CSLIST", sday)
		Call UpdateLastCSInputDT("return", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
		rw ""
    next

ElseIf (cmdparam="cjmallCsreg2") Then ''CS목록 조회(주문취소)
    todate = LEFT(CStr(now()),10)
    maxloop = 15					'// TODO : 15일 이상 CS가 없으면 그 이후 CS내역을 가져오지 못한다.

	'// ========================================================================

    stdt = getLastCSInputDT("ordercancel")
    sday = stdt
    for i=0 to maxloop
		rw sday & " CS건[주문내역:취소] 조회 등록시작 ======================================"
    	call getCjCsListInOrder("CSORDCANCELLIST", sday)
		Call UpdateLastCSInputDT("ordercancel", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next

ElseIf (cmdparam="cjmallCsreg3") Then ''CS목록 조회(CS출고 : 교환출고 등)
    todate = LEFT(CStr(now()),10)
    maxloop = 15					'// TODO : 15일 이상 CS가 없으면 그 이후 CS내역을 가져오지 못한다.

	'// ========================================================================
    stdt = getLastCSInputDT("order")
    sday = stdt
    for i=0 to maxloop
        rw sday & " CS건[주문내역:출고,기출하] 조회 등록시작 ======================================"
    	call getCjCsListInOrder("CSORDLIST", sday)
		Call UpdateLastCSInputDT("order", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next

ElseIf (cmdparam="cjmallCancelreg") Then ''주문취소목록 조회
    sday = LEFT(CStr(now()),10)
	call getCjOrderCancelList(sday)

	rw sday
ElseIf (cmdparam="cjmallCommonCode") Then ''공통코드 조회
	Dim ccd
	ccd = request("CommCD")
	call getcjCommonCodeList(ccd)
Else
	rw "미지정 ["&cmdparam&"]"
End If

If (alertMsg <> "") Then
	rw "msg : " & alertMsg
End if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
