<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<%
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.9","61.252.133.10","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72", "61.252.133.67", "61.252.133.70")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function


dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

' if (Not CheckVaildIP(ref)) then
'     'rw ref
'     dbget.Close()
'     response.end
' end if

dim act     : act = requestCheckVar(request("act"),32)
dim param1  : param1 = requestCheckVar(request("param1"),32)
dim param2  : param2 = requestCheckVar(request("param2"),32)
dim param3  : param3 = requestCheckVar(request("param3"),32)
dim param4  : param4 = requestCheckVar(request("param4"),32)
dim sqlStr, i, paramData, retVal
dim retCnt : retCnt = 0

Dim cnt
Dim OutMallOrderSerialArr
Dim OrgDetailKeyArr
Dim songjangDivArr
Dim songjangNoArr, sendReqCntArr, beasongdateArr, outmallGoodsIDArr, orderItemOptionArr, orderItemOptionNameArr, outmalloptionnoArr, requireDetail11stYNArr, beasongNum11stArr, reserve01Arr
dim oLotteitem, itemidArr, orgsendReqCntArr

select Case act

    Case "outmallSongJangIp" ''제휴사 송장입력
    'response.end

        sqlStr = "select top 10 T.OutMallOrderSerial, D.songjangDiv, D.songjangNo "
        sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T"
        sqlStr = sqlStr & " Join db_order.dbo.tbl_order_master M on T.orderserial=M.orderserial"
        sqlStr = sqlStr & " Join db_order.dbo.tbl_order_detail D on T.orderserial=D.orderserial"
		sqlStr = sqlStr & " 	and IsNull(T.changeitemid, T.matchitemid)=D.itemid"					'// 기존 주문에 합쳐진 경우(빨강1개,파랑1개 -> 파랑2개)
		sqlStr = sqlStr & " 	and IsNull(T.changeitemoption, T.matchitemoption)=D.itemoption"
        sqlStr = sqlStr & " 	and D.currstate=7"
        sqlStr = sqlStr & " left join db_order.dbo.tbl_songjang_div V on D.songjangDiv=V.divcd"
'        sqlStr = sqlStr & " where datediff(m,T.regdate,getdate())<7"    ''20130304 추가
        sqlStr = sqlStr & " where T.regdate > dateadd(month, -2, getdate()) "    ''7개월 -> 2개월로 변경..2021-11-18 김진영
        sqlStr = sqlStr & " and T.sellsite='"&param1&"'"
        sqlStr = sqlStr & " and T.OrgDetailKey is Not NULL"             ''디테일키 입력 주문건만..
        sqlStr = sqlStr & " and IsNULL(T.sendState,0)=0"
        sqlStr = sqlStr & " and T.sendReqCnt<3"                         ''여러번 시도 안되도록. 추가.
        sqlStr = sqlStr & " and T.matchState not in ('R','D','B')"      ''교환 취소 반품 제외.
        sqlStr = sqlStr & " GROUP BY T.OutMallOrderSerial, D.songjangDiv, D.songjangNo "
        sqlStr = sqlStr & " order by T.OutMallOrderSerial desc"

        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        cnt = rsget.RecordCount
        ReDim OutMallOrderSerialArr(cnt)
        ReDim songjangDivArr(cnt)
        ReDim songjangNoArr(cnt)

        i = 0
        if Not rsget.Eof then
            do until rsget.eof
            OutMallOrderSerialArr(i) = rsget("OutMallOrderSerial")
			songjangDivArr(i) = rsget("songjangDiv")
			songjangNoArr(i) = Trim(replace(rsget("songjangNo"),"-",""))
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
                    IF (LCASE(param1)="wmp") then
                        paramData = "redSsnKey=system&OutMallOrderSerial="&OutMallOrderSerialArr(i)&"&hdc_cd="&TenDlvCode2WMPDlvCode(songjangDivArr(i))&"&songjangNo="&songjangNoArr(i)
                        ''response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/wmp_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                    ElseIF (LCASE(param1)="wmpfashion") then
                        paramData = "redSsnKey=system&OutMallOrderSerial="&OutMallOrderSerialArr(i)&"&hdc_cd="&TenDlvCode2WMPDlvCode(songjangDivArr(i))&"&songjangNo="&songjangNoArr(i)
                        ''response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/wmpfashion_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                    End If
                end if
            next
        end if

    Case ELSE
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
