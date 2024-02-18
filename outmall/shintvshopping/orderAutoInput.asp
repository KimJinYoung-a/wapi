<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 60 * 15
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<%
public function getXSiteTmpOrderBatchTargetValidList(iSellSite)
    Dim sqlStr
    sqlStr = "db_temp.dbo.usp_TEN_xSiteTmpOrderBatchInputTarget '"&iSellSite&"',1000,1"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        getXSiteTmpOrderBatchTargetValidList = rsget.GetRows()
    end if
    rsget.close
end function

Dim ArrRows, sellsite
Dim i, bufStr, paramData, retVal
sellsite = request("sellsite")

ArrRows = getXSiteTmpOrderBatchTargetValidList(sellsite)
If Not IsArray(ArrRows) Then
    response.Write "S_NONE|Nop"
    dbget.Close() : response.end
End If

bufStr =""
For i=0 To UBound(ArrRows,2)
    If (NOT ISNULL(ArrRows(0,i)) and ArrRows(0,i) <> "") and (NOT ISNULL(ArrRows(3,i)) and ArrRows(3,i) <> "") Then
        paramData = "mode=add&xtype=batch&oseq="&ArrRows(0,i)&"&cksel="&ArrRows(3,i)&"&redSsnKey=system"
        if (application("Svr_Info")	= "Dev") then
            retVal = SendReq("http://testscm.10x10.co.kr/admin/etc/orderInput/OrderInput_Process.asp",paramData)
        else
            retVal = SendReq("http://stscm.10x10.co.kr/admin/etc/orderInput/OrderInput_Process.asp",paramData)
        end if

        if (LEFT(retVal,2)="S_") then
            response.write retVal
        else
            response.write "S_OK|<textarea cols=100 rows=10>"&retVal&"</textarea>"
        end if
    End If
Next
'http://localhost:11117/outmall/shintvshopping/orderAutoInput.asp?sellsite=shintvshopping
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
