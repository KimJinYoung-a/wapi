<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Response.CharSet="UTF-8"
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("222.109.123.95","211.206.236.117","115.94.163.42","61.252.133.88","192.168.1.70","192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next

    dim validToken : validToken = Array("bd2acd564c264459908cd0d744986ea0","554740c796aa47b1aae8ee9bacd2643c")
    dim authtkn : authtkn = LCASE(request("authtkn"))
    for i=0 to UBound(validToken)
        if (validToken(i)=authtkn) then
            CheckVaildIP = true
            exit function
        end if
    next

end function


dim ref : ref = Request.ServerVariables("REMOTE_ADDR")
Dim oJson, oDataArr, oDataItem
Set oJson = jsObject()

if (Not CheckVaildIP(ref)) then
    oJson("resultCode") = "S_ERR" ''데이타 없음.
    oJson("resultMessage") = ref
    oJson("resultCount") = 0

    oJson.flush
    Set oJson = Nothing
    dbget.Close() :     response.end
end if

dim apiAction   : apiAction=requestCheckVar(request("apiAction"),32)
dim topN        : topN=requestCheckVar(request("topN"),10)
dim yyyymm      : yyyymm=requestCheckVar(request("yyyymm"),7)
dim importance  : importance=requestCheckVar(request("importance"),6)


if (topN="") then topN="5"
if (yyyymm="") then yyyymm=LEFT(dateadd("d",-10,NOW()),7)
if (importance="") then importance="0"

dim sqlStr
dim ArrRows, i,cnt
dim bufStr



if (apiAction="getCompanyState")  then
    sqlStr = "db_partner.[dbo].[usp_Ten_Company_State_API_S]"
    rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    IF not rsget.EOF THEN
		ArrRows = rsget.getRows()
	END IF
	rsget.close

    if IsArray(ArrRows) then
        oJson("resultCode") = "S_OK"
        oJson("resultMessage") = ""
        oJson("resultCount") = UBound(ArrRows,2)+1

        Set oDataArr = jsArray()			'배열구조로 선언
        For i=0 To UBound(ArrRows,2)
            Set oDataItem = jsObject()

            oDataItem("companyno") = ArrRows(0,i)

			set oDataArr(null) = oDataItem
			SET oDataItem = Nothing
		Next

		Set oJson("retDatas") = oDataArr
		Set oDataArr = Nothing
    else
        oJson("resultCode") = "S_NONE" ''데이타 없음.
        oJson("resultMessage") = ""
        oJson("resultCount") = 0
    end if
elseif (apiAction="getdlvtrc")  then
    sqlStr = "db_order.[dbo].[usp_Ten_Delivery_Trace_Queue_API_S] "
    rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    IF not rsget.EOF THEN
        ArrRows = rsget.getRows()
    END IF
    rsget.close

    if isArray(ArrRows) then
        oJson("resultCode") = "S_OK"
        oJson("resultMessage") = ""
        oJson("resultCount") = UBound(ArrRows,2)+1

        Set oDataArr = jsArray()			'배열구조로 선언
        For i=0 To UBound(ArrRows,2)
            Set oDataItem = jsObject()

            oDataItem("songjangno")     = ArrRows(0,i)
            oDataItem("songjangdiv")    = ArrRows(1,i)
            oDataItem("gettype")        = ArrRows(2,i)


            set oDataArr(null) = oDataItem
            Set oDataItem = Nothing
        Next
        Set oJson("retDatas") = oDataArr
        Set oDataArr = Nothing
    else
        oJson("resultCode") = "S_NONE"
        oJson("resultMessage") = ""
        oJson("resultCount") = 0
    end if
elseif (apiAction="getnvlowest")  then
    sqlStr = "db_temp.[dbo].[usp_ten_nv_crawl_matchid_get] (" & importance & ")"
    rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    IF not rsget.EOF THEN
        ArrRows = rsget.getRows()
    END IF
    rsget.close

    if isArray(ArrRows) then
        oJson("resultCode") = "S_OK"
        oJson("resultMessage") = ""
        oJson("resultCount") = UBound(ArrRows,2)+1

        Set oDataArr = jsArray()			'배열구조로 선언
        For i=0 To UBound(ArrRows,2)
            Set oDataItem = jsObject()

            oDataItem("idx")            = ArrRows(0,i)
            oDataItem("matchNvMid")     = ArrRows(1,i)

            set oDataArr(null) = oDataItem
            Set oDataItem = Nothing
        Next
        Set oJson("retDatas") = oDataArr
        Set oDataArr = Nothing
    else
        oJson("resultCode") = "S_NONE"
        oJson("resultMessage") = ""
        oJson("resultCount") = 0
    end if
elseif (apiAction="YYYYYYYYYYYYYYYYYY")  then
    sqlStr = "db_etcmall.[dbo].[usp_Ten_OutMall_API_Json_Data] ("&itemid&",'"&apiAction&"','"&mallid&"')"
    rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    IF not rsget.EOF THEN
        ArrRows = rsget.getRows()
    END IF
    rsget.close

    if isArray(ArrRows) then
        oJson("resultCode") = "S_OK"
        oJson("resultMessage") = ""
        oJson("resultCount") = UBound(ArrRows,2)+1
            Set oDataItem = jsObject()

            oDataItem("itemid")         = ArrRows(0,i)
            oDataItem("mallitemid")     = ArrRows(1,i)
            oDataItem("saleStat")       = ArrRows(2,i)
            oDataItem("salePrice")      = ArrRows(3,i)
            oDataItem("originPrice")    = ArrRows(4,i)
            oDataItem("itemname")       = ArrRows(5,i)
            oDataItem("itemlimitStockno")   = ArrRows(6,i)
            oDataItem("ttlOptCnt")      = ArrRows(7,i)
            oDataItem("sellOptCnt")     = ArrRows(8,i)

            Set oJson("retData") = oDataItem
            Set oDataItem = Nothing
    else
        oJson("resultCode") = "S_NONE"
        oJson("resultMessage") = ""
        oJson("resultCount") = 0
    end if
elseif (apiAction="ZZZZZZZZZZZZZZZZZZZZZZZ")  then
    sqlStr = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_READ_coupang_WithInfo] ('"&apiAction&"','"&mallid&"',"&topN&")"
    rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    IF not rsget.EOF THEN
		ArrRows = rsget.getRows()
	END IF
	rsget.close

    if IsArray(ArrRows) then
        oJson("resultCode") = "S_OK"
        oJson("resultMessage") = ""
        oJson("resultCount") = UBound(ArrRows,2)+1

        Set oDataArr = jsArray()			'배열구조로 선언
        For i=0 To UBound(ArrRows,2)
            Set oDataItem = jsObject()
            oDataItem("queidx") = ArrRows(0,i)
            oDataItem("mallid") = ArrRows(1,i)
            oDataItem("itemid") = ArrRows(2,i)
            oDataItem("apiAction") = ArrRows(3,i)
            oDataItem("goodsno") = ArrRows(4,i)


			set oDataArr(null) = oDataItem
			SET oDataItem = Nothing
		Next

		Set oJson("retDatas") = oDataArr
		Set oDataArr = Nothing
    else
        oJson("resultCode") = "S_NONE" ''데이타 없음.
        oJson("resultMessage") = ""
        oJson("resultCount") = 0
    end if
else
    oJson("resultCode") = "ERR"
    oJson("resultMessage") = "UnDefained"
end if

oJson.flush
Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->