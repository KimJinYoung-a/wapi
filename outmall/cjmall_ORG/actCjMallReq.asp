<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 600 %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/outmall/cjmall/incCJmallFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/cjmall/cjmallitemcls.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),30)
Dim sday : sday = request("sday")
Dim cksel : cksel = request("cksel")
Dim subcmd : subcmd = request("subcmd")
Dim retFlag : retFlag = request("retFlag")
Dim iitemid, ret, sqlStr, AssignedRow
Dim alertMsg, ierrStr
Dim SuccCNT : SuccCNT = 0
Dim FailCNT : FailCNT = 0
Dim i
Dim todate, stdt, maxloop
Dim ArrRows
Dim sellOK

If (cmdparam="RegSelect") Then					'상품 등록
	cksel = split(cksel, ",")
	For i = 0 To UBound(cksel)
		iitemid = Trim(cksel(i))
		ret = regCjMallOneItem(iitemid, ierrStr)
		If (Not ret) Then
			FailCNT = FailCNT + 1
			rw ierrStr
		Else
			SuccCNT = SuccCNT + 1
		End If
	Next
	alertMsg = ""&SuccCNT&"건 성공 "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"건 실패 "
	End If
ElseIf (cmdparam="confirmItem") Then			'상품 조회
    cksel = split(cksel,",")
	For i = 0 To UBound(cksel)
		iitemid = Trim(cksel(i))
		ret = oneCjMallItemConfirm(iitemid, ierrStr)
		If (Not ret) Then
			FailCNT = FailCNT + 1
			rw ierrStr
		Else
			SuccCNT = SuccCNT + 1
		End If
	Next
	alertMsg = ""&SuccCNT&"건 성공 "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"건 실패 "
	End If
ElseIf (cmdparam="EditSellYn") Then '			'상품 상태 수정
	cksel = split(cksel,",")
	For i = 0 To UBound(cksel)
		iitemid=Trim(cksel(i))
		ret = editSellStatusCjmallOneItem(iitemid, ierrStr, subcmd)
		If (Not ret) Then
			FailCNT = FailCNT + 1
			rw ierrStr
		Else
			SuccCNT = SuccCNT + 1
		End If
	Next
	alertMsg = ""&SuccCNT&"건 성공 "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"건 실패 "
	End If
ElseIf (cmdparam="EditSelect") Then				'상품 정보 수정.
	cksel = split(cksel, ",")
	For i = 0 To UBound(cksel)
		iitemid = Trim(cksel(i))
		ret = editCjmallOneItem(iitemid, ierrStr)
		If (Not ret) Then
			FailCNT = FailCNT + 1
			rw ierrStr
		Else
			SuccCNT = SuccCNT + 1
		End If
	Next
	alertMsg = ""&SuccCNT&"건 성공 "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"건 실패 "
	End If
ElseIf (cmdparam = "EdSaleDTSel") Then			'단품 상태 수정.
	cksel = split(cksel, ",")
	For i = 0 To UBound(cksel)
		iitemid = Trim(cksel(i))
		ret = editDTCjmallOneItem(iitemid, ierrStr)
		If (Not ret) Then
			FailCNT = FailCNT + 1
			rw ierrStr
		Else
			SuccCNT = SuccCNT + 1
		End If
	Next
	alertMsg = ""&SuccCNT&"건 성공 "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"건 실패 "
	End If

'	if (retFlag<>"") then
'		Response.Write "<script language=javascript>parent."&retFlag&";</script>"
'		response.end
'	end if
ElseIf (cmdparam="EditQty") Then				'단품 수량 수정
	cksel = split(cksel,",")
	For i=0 To UBound(cksel)
		iitemid=Trim(cksel(i))
		ret = editqtyCjmallOneItem(iitemid, ierrStr)
		If (Not ret) Then
			FailCNT = FailCNT + 1
			rw ierrStr
		Else
			SuccCNT = SuccCNT + 1
		End If
	Next
	alertMsg = ""&SuccCNT&"건 성공 "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"건 실패 "
	End If
ElseIf (cmdparam="EditPriceSelect2") Then		'상품 가격 수정(관리자만)
	cksel = split(cksel, ",")
	For i=0 To UBound(cksel)
		iitemid=Trim(cksel(i))
		ret = editSellPriceCjmallOneItem(iitemid, ierrStr)
		If (Not ret) Then
			FailCNT = FailCNT + 1
			rw ierrStr
		Else
			SuccCNT = SuccCNT + 1
		End If
	Next
	alertMsg = ""&SuccCNT&"건 성공 "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"건 실패 "
	End If
ElseIf (cmdparam="EditPriceSelect") Then		'단품 가격 수정
	cksel = split(cksel, ",")
	For i=0 To UBound(cksel)
		iitemid=Trim(cksel(i))
		ret = editPriceCjmallOneItem(iitemid, ierrStr)
		If (Not ret) Then
			FailCNT = FailCNT + 1
			rw ierrStr
		Else
			SuccCNT = SuccCNT + 1
		End If
	Next
	alertMsg = ""&SuccCNT&"건 성공 "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"건 실패 "
	End If
ElseIf (cmdparam="cjmallCommonCode") Then		'공통코드 조회
	Dim ccd
	ccd = request("CommCD")
	Call getcjCommonCodeList(ccd)
ElseIf (cmdparam="EditSelect2") Then							'정보 + 단품 수정.
	cksel = split(cksel, ",")
	For i=0 To UBound(cksel)
		iitemid=Trim(cksel(i))
		ierrStr = ""
		ret = oneCjMallItemConfirm(iitemid, ierrStr)			'상품 조회
		If (Not ret) Then
			rw ierrStr
		End If
		rw "<strong>------------------------------------------------</strong>"

		ierrStr = ""
		ret = editCjmallOneItem(iitemid, ierrStr)				'상품 정보 수정
		If (Not ret) Then
			rw ierrStr
		End If
		rw "<strong>------------------------------------------------</strong>"

		ierrStr = ""
		ret = editqtyCjmallOneItem(iitemid, ierrStr)			'단품 수량 수정
		If (Not ret) Then
			rw ierrStr
		End If
		rw "<strong>------------------------------------------------</strong>"

		ierrStr = ""
		ret = editDTCjmallOneItem(iitemid, ierrStr)				'단품 상태 수정.
		If (Not ret) Then
			rw ierrStr
		End If
		rw "<strong>------------------------------------------------</strong>"

		ierrStr = ""
		ret = editPriceCjmallOneItem(iitemid, ierrStr)			'단품 가격 수정.
		If (Not ret) Then
			rw ierrStr
		End If
		rw "<strong>------------------------------------------------</strong>"
	Next
ElseIf (cmdparam="confirmItemAuto") Then						'판매상태Check(관리자)
	cksel = ""
	If (subcmd = "1") Then
		sqlStr = "select top 15 r.itemid "
		sqlStr = sqlStr & "	from db_outmall.dbo.tbl_cjmall_regitem r"
		sqlStr = sqlStr & "	Join db_AppWish.dbo.tbl_item i"
		sqlStr = sqlStr & "	on r.itemid=i.itemid"
		sqlStr = sqlStr & "	where r.cjMallStatcd=3" ''-1: 등록실패 , 0: 등록예정, 1: 전송시도 , 3:승인대기
		sqlStr = sqlStr & "	order by r.lastStatCheckDate, (CASE WHEN r.cjmallsellyn='X' THEN '0' ELSE r.cjmallsellyn END), r.cjmallLastUpdate , r.itemid desc"
	Else
		sqlStr = "select top 15 r.itemid "
		sqlStr = sqlStr & "	from db_outmall.dbo.tbl_cjmall_regitem r"
		sqlStr = sqlStr & "	where cjMallStatcd>0" ''-1: 등록실패 , 0: 등록예정, 1: 전송시도
		sqlStr = sqlStr & "	and isnull(cjMallPrdno, '') <> '' "
		sqlStr = sqlStr & "	order by r.lastStatCheckDate, (CASE WHEN r.cjmallsellyn='X' THEN '0' ELSE r.cjmallsellyn END), r.cjmallLastUpdate , r.itemid desc"
	End If
    rsCTget.Open sqlStr ,dbCTget , 1
    If not rsCTget.Eof Then
        ArrRows = rsCTget.getRows()
    End If
    rsCTget.close
	If isArray(ArrRows) then
		For i =0 To UBound(ArrRows,2)
		    cksel = cksel + CStr(ArrRows(0,i)) + ","
		Next
	Else
		rw "S_NONE"
		dbCTget.Close() : response.end
	End if

	cksel = split(cksel,",")
	For i=0 To UBound(cksel)
		iitemid = Trim(cksel(i))
		If (iitemid <> "") Then
			ret = oneCjMallItemConfirm(iitemid, ierrStr)
			If (Not ret) Then
				rw ierrStr
			End If
		End If
	Next
ElseIf (cmdparam="LIST") Then  		''승인된 상품인지 총 기간동안 검색		http://testscm.10x10.co.kr/admin/etc/cjmall/actCjMallreq.asp?cmdparam=LIST
	listCjMallItem(sday)
ElseIf (cmdparam="DayLIST") Then	''승인된 상품인지 일정기간동안 검색		http://testscm.10x10.co.kr/admin/etc/cjmall/actCjMallreq.asp?cmdparam=DayLIST&sday=0
	daylistCjMallItem(sday)
ElseIf (cmdparam="cjmallOrdreg") Then ''주문목록 조회

    todate = LEFT(CStr(now()),10)
    maxloop = 4 ''추석때 발주 안함 추석이후 영업일에 주문 통보됨// 연휴시 이값을 길게 할것. (일반적으로 3,4일 적당)
    stdt = getLastOrderInputDT()
    sday = stdt
    for i=0 to maxloop
        rw sday & "주문건 등록시작 ======================================"
    	call getCjOrderList("ORDLIST", sday)

    	'' 취소쪽은 CS건에서함. 주석처리 2013/08/05
    	''rw sday & "주문취소건 등록시작 ======================================"
    	''call getCjOrderList("ORDCANCELLIST", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next
ElseIf (cmdparam="cjmallOrdUp") Then ''주문목록 실판매가 업데이트

    todate = LEFT(CStr(now()),10)
    maxloop = 1
    stdt = getLastOrderInputDTUp()
    if (request("stdt")<>"") then stdt=request("stdt")
    rw stdt
    if stdt>"2014-11-27" then 
        response.write "TT"
        response.end
    end if

    sday = stdt
    for i=0 to maxloop-1
        rw sday & "주문건 등록시작 ======================================"
    	call getCjOrderList("ORDLISTUP", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next
    rw "<form name=frmR method=post action=''><input type='hidden' name='cmdparam' value='cjmallOrdUp'><input type='hidden' name='stdt' value='"&sday&"'><input type='button' name='reloadBtn' value='reload' onClick='document.frmR.submit();'></form>"

    if (sday<"2014-11-27") then
    response.write "<script>"
    response.write "setTimeout(function(){document.frmR.submit();},2000);"
    response.write "</script>"
    end if

ElseIf (cmdparam="cjmallCsreg") Then ''CS목록 조회
    todate = LEFT(CStr(now()),10)
    maxloop = 10
    stdt = LEFT(CStr(DATEADD("d",-1,now())),10)
    sday = stdt
    for i=0 to maxloop
        rw sday & "CS건 조회 등록시작 ======================================"
    	call getCjCsList("CSLIST", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next
ElseIf (cmdparam="cjmallCancelreg") Then ''주문취소목록 조회
    sday = LEFT(CStr(now()),10)
	call getCjOrderCancelList(sday)

	rw sday
Else
	rw "미지정 ["&cmdparam&"]"
End If

If (alertMsg <> "") Then
	IF (IsAutoScript) Then
		rw alertMsg
	Else
		response.write "<script>alert('"&alertMsg&"');</script>"
	End If
End if
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->