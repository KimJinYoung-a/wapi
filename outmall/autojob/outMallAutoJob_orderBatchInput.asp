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

public function getXSiteTmpOrderBatchTargetValidList(iSellSite)
    Dim sqlStr
    sqlStr = "db_temp.dbo.usp_TEN_xSiteTmpOrderBatchInputTarget '"&iSellSite&"',500,1"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        getXSiteTmpOrderBatchTargetValidList = rsget.GetRows()
    end if
    rsget.close

end function

public function getXSiteTmpOrderBatchTargetValidYYYYMMDD(iSellSite)
    Dim sqlStr
    sqlStr = "db_temp.dbo.usp_TEN_xSiteTmpOrderLastYYYYMM '"&iSellSite&"'"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        getXSiteTmpOrderBatchTargetValidYYYYMMDD = rsget.GetRows()
    end if
    rsget.close

end function


public function setXSiteTmpOrderBatchTargetValidYYYYMMDD(iSellSite,iyyyymmdd)
    Dim sqlStr
    sqlStr = "db_temp.dbo.usp_TEN_xSiteTmpOrderLastYYYYMM_SET '"&iSellSite&"','"&iyyyymmdd&"'"
    dbget.Execute sqlStr
end function


dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

IF NOT (application("Svr_Info")	= "Dev") Then
    ' if (Not CheckVaildIP(ref)) then
    '     'rw ref
    '     dbget.Close()
    '     response.end
    ' end if
end IF

Dim act             : act = requestCheckVar(request("act"),32)
Dim sellsite        : sellsite = requestCheckVar(request("sellsite"),32)
Dim oseq            : oseq = requestCheckVar(request("oseq"),10)
Dim xOrderSerial    : xOrderSerial = requestCheckVar(request("xOrderSerial"),32)
Dim yyyymmdd        : yyyymmdd = requestCheckVar(request("yyyymmdd"),10)

dim sqlStr, i
dim isValidSiteOrTime : isValidSiteOrTime = FALSE
Dim bufStr, ArrRows

Dim retVal, paramData

''/outmall/autojob/outMallAutoJob_orderBatchInput.asp?act=G&sellsite=cjmall
''/outmall/autojob/outMallAutoJob_orderBatchInput.asp?act=S&sellsite=cjmall
''/outmall/autojob/outMallAutoJob_orderBatchInput.asp?act=I&oseq=11111&xOrderSerial=123123123
if (act="G") then
    '' 입력 가능한시각인지 검토
    sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderBatch_AvailTime] '"&sellsite&"'"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        isValidSiteOrTime = rsget("isValid")>0
    end if
    rsget.Close

    '' STEP1 - CHECK valid
    if NOT (isValidSiteOrTime) then
        response.Write "S_ERR|Not Valid Time or Site :"&sellsite
        dbget.Close() : response.end
    end if

    paramData = "redSsnKey=system&mode=getxsiteorderlist&sellsite="&sellsite

    '' STEP2 - Receive API
    if (application("Svr_Info")	= "Dev") then
        retVal = "S_ERR|Dev not valid"
    elseif (LCASE(sellsite)="cjmall") then
        paramData = paramData&"&cmdparam=cjmallOrdreg"
        retVal = SendReq("http://scm.10x10.co.kr/admin/etc/cjMall/actCjMallReq.asp",paramData)
    elseif (LCASE(sellsite)="lotteimall") then
        retVal = SendReq("http://stscm.10x10.co.kr/admin/etc/orderinput/xSiteOrder_lotteimall_Process.asp",paramData)
    elseif (LCASE(sellsite)="gmarket1010") then

        retVal = SendReq("http://wapi.10x10.co.kr/outmall/gmarket/xSiteOrder_gmarket_Process.asp",paramData)
    elseif (LCASE(sellsite)="11st1010") then

        retVal = SendReq("http://wapi.10x10.co.kr/outmall/11st/xSiteOrder_11st1010_Process.asp",paramData)
    elseif (LCASE(sellsite)="interpark") then

		retVal = SendReq("http://wapi.10x10.co.kr/outmall/order/xSiteOrder_Ins_Process.asp",paramData)
	elseif (LCASE(sellsite)="auction1010") then

		retVal = SendReq("http://wapi.10x10.co.kr/outmall/order/xSiteOrder_Ins_Process.asp",paramData)
	elseif (LCASE(sellsite)="nvstorefarm") then

		retVal = SendReq("http://wapi.10x10.co.kr/outmall/order/xSiteOrder_Ins_Process.asp",paramData)
	elseif (LCASE(sellsite)="nvstoremoonbangu") then

		retVal = SendReq("http://wapi.10x10.co.kr/outmall/order/xSiteOrder_Ins_Process.asp",paramData)
	elseif (LCASE(sellsite)="mylittlewhoopee") then

		retVal = SendReq("http://wapi.10x10.co.kr/outmall/order/xSiteOrder_Ins_Process.asp",paramData)

	elseif (LCASE(sellsite)="nvstoregift") then

		retVal = SendReq("http://wapi.10x10.co.kr/outmall/order/xSiteOrder_Ins_Process.asp",paramData)
	elseif (LCASE(sellsite)="ezwel") then

		retVal = SendReq("http://wapi.10x10.co.kr/outmall/order/xSiteOrder_Ins_Process.asp",paramData)
	elseif (LCASE(sellsite)="lottecom") then

		retVal = SendReq("http://wapi.10x10.co.kr/outmall/order/xSiteOrder_Ins_Process.asp",paramData)
    elseif (LCASE(sellsite)="lotteon") then

        retVal = SendReq("http://wapi.10x10.co.kr/outmall/order/xSiteOrder_Ins_Process.asp",paramData)
    elseif (LCASE(sellsite)="ssg") then
        paramData = paramData & "&rcvtp="&request("rcvtp")
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/ssg/xSiteOrder_ssg_Process.asp",paramData)
	elseif (LCASE(sellsite)="halfclub") then

		retVal = SendReq("http://wapi.10x10.co.kr/outmall/halfclub/xSiteOrder_halfclub_Process.asp",paramData)
	elseif (LCASE(sellsite)="coupang") then

		retVal = SendReq("http://wapi.10x10.co.kr/outmall/coupang/xSiteOrder_coupang_Process.asp",paramData)
	elseif (LCASE(sellsite)="hmall1010") then

		'retVal = SendReq("http://wapi.10x10.co.kr/outmall/hmall/xSiteOrder_hmall_Process.asp",paramData)
        retVal = SendReq("http://wapi.10x10.co.kr/outmall/hmall/xSiteOrder_hmall_new_Process.asp",paramData)
	elseif (LCASE(sellsite)="wmp") then

		retVal = SendReq("http://wapi.10x10.co.kr/outmall/wmp/xSiteOrder_wmp_Process.asp",paramData)
    ' elseif (LCASE(sellsite)="sabangnet") then
    '     paramData = "redSsnKey=system&sellsite=sabangnet"
	' 	retVal = "http://wapi.10x10.co.kr/outmall/order/xSiteOrder_Ins_Process.asp",paramData)
	elseif (LCASE(sellsite)="wmpfashion") then

		retVal = SendReq("http://wapi.10x10.co.kr/outmall/wmpfashion/xSiteOrder_wmpfashion_Process.asp",paramData)
    else
        retVal = "S_ERR|Dev not valid Site : "&sellsite
        response.write retVal
        dbget.Close() : response.end
    end if

    if (LEFT(retVal,2)="S_") then
        response.write retVal
    else
        response.write "S_OK|<textarea cols=100 rows=10>"&retVal&"</textarea>"
    end if
elseif (act="S") and (sellsite="gseshop") then
    '' gsshop은 우리가 호출하면 gs에서 다시 쏜다
    '' outmall\order\xSiteOrder_GSShop_recv_Process.asp
    ArrRows = getXSiteTmpOrderBatchTargetValidYYYYMMDD(sellsite)

    if Not IsArray(ArrRows) then
        response.Write "S_NONE|Nop"
        dbget.Close() : response.end
    end if

    bufStr =""
    For i=0 To UBound(ArrRows,2)
        bufStr = bufStr&ArrRows(0,i)&"||"&sellsite&vBCRLF
    Next
    bufStr = "S_OK"&vBCRLF&bufStr

    response.write bufStr
elseif (act="F") and (sellsite="gseshop") then
    Call setXSiteTmpOrderBatchTargetValidYYYYMMDD(sellsite,yyyymmdd)
    response.write "S_OK"

elseif (act="S") and (sellsite<>"gseshop") then
    '' STEP3 - Input Avail Order
    ArrRows = getXSiteTmpOrderBatchTargetValidList(sellsite)

    if Not IsArray(ArrRows) then
        response.Write "S_NONE|Nop"
        dbget.Close() : response.end
    end if

    bufStr =""
    For i=0 To UBound(ArrRows,2)
        bufStr = bufStr&ArrRows(0,i)&"||"&sellsite&"||"&ArrRows(3,i)&vBCRLF  'OutMallOrderSeq||OutMallOrderSerial
    Next
    bufStr = "S_OK"&vBCRLF&bufStr

    response.write bufStr
elseif (act="I") and (xOrderSerial<>"") and (oseq<>"") then
    '' STEP4 - InputOrder from tmp to Real
    paramData = "mode=add&xtype=batch&oseq="&oseq&"&cksel="&xOrderSerial&"&redSsnKey=system"
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
else
    response.Write "S_ERR|Not Valid"
    dbget.Close() : response.end
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
