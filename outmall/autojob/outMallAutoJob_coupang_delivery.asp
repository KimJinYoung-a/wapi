<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/coupang/coupangItemcls.asp"-->
<!-- #include virtual="/outmall/coupang/incCoupangFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
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

if (Not CheckVaildIP(ref)) then
    'rw ref
    dbget.Close()
    response.end
end if

dim act     : act = requestCheckVar(request("act"),32)
dim param1  : param1 = requestCheckVar(request("param1"),32)
dim sqlStr, i, paramData, retVal
dim retCnt : retCnt = 0
Dim cnt
Dim makeridArr

select Case act
    Case "delivery" 'ÄíÆÎ ¹è¼ÛÁö
        sqlStr = ""
        sqlStr = sqlStr & " SELECT TOP 30 m.makerid "
        sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_coupang_branddelivery_mapping as m "
        sqlStr = sqlStr & " LEFT JOIN db_etcmall.[dbo].[tbl_coupang_delivery_LOG] as l on m.makerid = l.makerid "
        sqlStr = sqlStr & " WHERE isnull(m.outboundShippingPlaceCode, '') = '' "
        sqlStr = sqlStr & " and isnull(l.failCnt, 0) < 5 "
        sqlStr = sqlStr & " ORDER BY IsNull(l.failcnt, 0), m.makerid ASC "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        cnt = rsget.RecordCount
        ReDim makeridArr(cnt)
        i = 0
        if Not rsget.Eof then
            do until rsget.eof
            makeridArr(i) = rsget("makerid")
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
            for i=LBound(makeridArr) to UBound(makeridArr)
                if (makeridArr(i)<>"") then
                    IF (LCASE(param1)="coupang") then
                        paramData = "redSsnKey=system&makerid="&makeridArr(i)
                        IF application("Svr_Info")="Dev" THEN
                            retVal = SendReq("http://localhost:11117/outmall/proc/coupang_DeliveryProc.asp",paramData)
                        ELSE
                            retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/coupang_DeliveryProc.asp",paramData)
                        End IF
                        response.write retVal
                    End If
                end if
            next
        end if
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
