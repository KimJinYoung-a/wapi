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


If (cmdparam="cjmallOrdreg") Then ''�ֹ���� ��ȸ

    todate = LEFT(CStr(now()),10)
    maxloop = 10
    stdt = getLastOrderInputDT()
    sday = stdt
    for i=0 to maxloop
        rw sday & "�ֹ��� ��Ͻ��� ======================================"
    	call getCjOrderList("ORDLIST", sday)
    	rw sday & "�ֹ���Ұ� ��Ͻ��� ======================================"
    	call getCjOrderList("ORDCANCELLIST", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next
ElseIf (cmdparam="cjmallCsreg1") Then ''CS��� ��ȸ(��ǰ)
    todate = LEFT(CStr(now()),10)
    maxloop = 15					'// TODO : 15�� �̻� CS�� ������ �� ���� CS������ �������� ���Ѵ�.

	'// ========================================================================
    stdt = getLastCSInputDT("return")
	''rw stdt
    sday = stdt
    for i=0 to maxloop
        rw sday & " CS��[ȸ������] ��ȸ ��Ͻ��� ======================================"
    	call getCjCsList("CSLIST", sday)
		Call UpdateLastCSInputDT("return", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
		rw ""
    next

ElseIf (cmdparam="cjmallCsreg2") Then ''CS��� ��ȸ(�ֹ����)
    todate = LEFT(CStr(now()),10)
    maxloop = 15					'// TODO : 15�� �̻� CS�� ������ �� ���� CS������ �������� ���Ѵ�.

	'// ========================================================================

    stdt = getLastCSInputDT("ordercancel")
    sday = stdt
    for i=0 to maxloop
		rw sday & " CS��[�ֹ�����:���] ��ȸ ��Ͻ��� ======================================"
    	call getCjCsListInOrder("CSORDCANCELLIST", sday)
		Call UpdateLastCSInputDT("ordercancel", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next

ElseIf (cmdparam="cjmallCsreg3") Then ''CS��� ��ȸ(CS��� : ��ȯ��� ��)
    todate = LEFT(CStr(now()),10)
    maxloop = 15					'// TODO : 15�� �̻� CS�� ������ �� ���� CS������ �������� ���Ѵ�.

	'// ========================================================================
    stdt = getLastCSInputDT("order")
    sday = stdt
    for i=0 to maxloop
        rw sday & " CS��[�ֹ�����:���,������] ��ȸ ��Ͻ��� ======================================"
    	call getCjCsListInOrder("CSORDLIST", sday)
		Call UpdateLastCSInputDT("order", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next

ElseIf (cmdparam="cjmallCancelreg") Then ''�ֹ���Ҹ�� ��ȸ
    sday = LEFT(CStr(now()),10)
	call getCjOrderCancelList(sday)

	rw sday
ElseIf (cmdparam="cjmallCommonCode") Then ''�����ڵ� ��ȸ
	Dim ccd
	ccd = request("CommCD")
	call getcjCommonCodeList(ccd)
Else
	rw "������ ["&cmdparam&"]"
End If

If (alertMsg <> "") Then
	rw "msg : " & alertMsg
End if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
