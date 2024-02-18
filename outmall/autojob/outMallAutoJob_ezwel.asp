<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/outmall/ezwel/ezwelItemcls.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<%
Function CheckVaildIP(ref)
	CheckVaildIP = false
	Dim VaildIP
	VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.9","61.252.133.10","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
	Dim i
	For i=0 to UBound(VaildIP)
		If (VaildIP(i)=ref) then
			CheckVaildIP = true
			Exit Function
		End If
	Next
End Function

Dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

' If (Not CheckVaildIP(ref)) Then
'     dbget.Close()
'     response.end
' End If

Dim act     : act = requestCheckVar(request("act"),32)
Dim param1  : param1 = requestCheckVar(request("param1"),32)
Dim param2  : param2 = requestCheckVar(request("param2"),32)
Dim param3  : param3 = requestCheckVar(request("param3"),32)
Dim param4  : param4 = requestCheckVar(request("param4"),32)
Dim param5  : param5 = requestCheckVar(request("param5"),32)
Dim sqlStr, i, paramData, retVal
Dim retCnt : retCnt = 0

Dim cnt
Dim OutMallOrderSerialArr
Dim OrgDetailKeyArr
Dim songjangDivArr
Dim songjangNoArr, sendReqCntArr, beasongdateArr, outmallGoodsIDArr
Dim oEzwel, itemidArr

Select Case act
	Case "outmallSongJangIp"
    'response.end

		sqlStr = "select top 30 T.orderserial, T.OutMallOrderSerial"
		sqlStr = sqlStr & " ,T.OrgDetailKey, IsNULL(T.sendState,0) as sendState"
		sqlStr = sqlStr & " ,D.songjangDiv, D.songjangNo, D.itemNo, D.beasongdate, T.outMallGoodsNo"
		sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T"
		sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master M"
		sqlStr = sqlStr & " 	on T.orderserial=M.orderserial"
		sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_detail D"
		sqlStr = sqlStr & " 	on T.orderserial=D.orderserial"
		sqlStr = sqlStr & " 	and IsNull(T.changeitemid, T.matchitemid)=D.itemid"					'// 기존 주문에 합쳐진 경우(빨강1개,파랑1개 -> 파랑2개)
		sqlStr = sqlStr & " 	and IsNull(T.changeitemoption, T.matchitemoption)=D.itemoption"
		sqlStr = sqlStr & " 	and D.currstate=7"
		sqlStr = sqlStr & " 	left join db_order.dbo.tbl_songjang_div V"
		sqlStr = sqlStr & " 	on D.songjangDiv=V.divcd"
'        sqlStr = sqlStr & " where datediff(m,T.regdate,getdate())<7"    ''20130304 추가
        sqlStr = sqlStr & " where T.regdate > dateadd(month, -2, getdate()) "    ''7개월 -> 2개월로 변경..2021-11-18 김진영
		sqlStr = sqlStr & " and T.sellsite='"&param1&"'"
		sqlStr = sqlStr & " and T.OrgDetailKey is Not NULL"             ''디테일키 입력 주문건만..
		sqlStr = sqlStr & " and IsNULL(T.sendState,0)=0"
		sqlStr = sqlStr & " and T.sendReqCnt<3"                         ''여러번 시도 안되도록. 추가.
		sqlStr = sqlStr & " and T.matchState not in ('R','D','B')"      ''교환 취소 반품 제외.
		sqlStr = sqlStr & " order by D.beasongdate desc"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		cnt = rsget.RecordCount
		ReDim TenOrderserial(cnt)
		ReDim OutMallOrderSerialArr(cnt)
		ReDim OrgDetailKeyArr(cnt)
		ReDim songjangDivArr(cnt)
		ReDim songjangNoArr(cnt)
		Redim sendReqCntArr(cnt)
		Redim beasongdateArr(cnt)
		Redim outmallGoodsIDArr(cnt)
		i = 0
		if Not rsget.Eof then
			do until rsget.eof
			TenOrderserial(i) = rsget("orderserial")
			OutMallOrderSerialArr(i) = rsget("OutMallOrderSerial")
			OrgDetailKeyArr(i) = rsget("OrgDetailKey")
			songjangDivArr(i) = rsget("songjangDiv")
			songjangNoArr(i) = rsget("songjangNo")
			sendReqCntArr(i) = rsget("itemNo")
			beasongdateArr(i) = rsget("beasongdate")
			outmallGoodsIDArr(i) = rsget("outMallGoodsNo")
			i=i+1
			rsget.MoveNext
			loop
		end if
		rsget.close

		if (cnt<1) then
			response.Write "S_NONE.."
			dbget.Close() : response.end
		else
			rw "CNT="&CNT
			for i=LBound(OutMallOrderSerialArr) to UBound(OutMallOrderSerialArr)
				if (OutMallOrderSerialArr(i)<>"") then
				    IF (LCASE(param1)="ezwel") then
				        paramData = "redSsnKey=system&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode2EzwelDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)
				        if (application("Svr_Info")<>"Dev") then
							'retVal = SendReq("http://scm.10x10.co.kr/admin/etc/ezwel/actEzwelSongjangInputProc.asp",paramData)
							retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/Ezwel_SongjangProc.asp",paramData)
							rw retVal
				        else
							retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/Ezwel_SongjangProc.asp",paramData)
							rw retVal
				        end if
				    end if
				end if
			next
        end if
    Case Else
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->