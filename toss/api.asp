<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% Server.ScriptTimeOut = 900 %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Response.CharSet="UTF-8"
%>
<!-- #include virtual="/lib/util/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib_UTF8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/aspJSON1.17.asp" -->
<!-- #include virtual="/toss/inctosspayCommon.asp" -->
<script language="jscript" runat="server">
    function jsURLDecode(v){ return decodeURI(v); }
    function jsURLEncode(v){ return encodeURI(v); }
</script>
<%

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("192.168.1.67","192.168.1.70","192.168.1.71","192.168.1.72","192.168.1.73","110.93.128.107","121.78.103.60","110.93.128.114","110.93.128.113", "::1", "192.168.5.200", "112.218.65.244")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

function JungsanTossPay(yyyymmdd, dateType, iNextCursor)
	dim jungsanData, conResult, rstJson
	dim item, itm

    dim appDivCode, PGkey, PGCSkey, appDate, cancelDate, appMethod
    dim appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, etcPoint
    dim sqlStr, sitename, PGgubun, PGuserid, retLength
    dim i

	jungsanData = "{"
	jungsanData = jungsanData &"""apiKey"":"""&CStr(TossPay_RestApi_Key)&""""
	jungsanData = jungsanData &",""dateType"":"""&CStr(dateType)&""""
	jungsanData = jungsanData &",""baseDate"":"""&CStr(yyyymmdd)&""""
	if (iNextCursor <> "") then
		jungsanData = jungsanData &",""nextCursor"":"""&CStr(iNextCursor)&""""
	end if
	jungsanData = jungsanData &"}"

	conResult = tossapi_Jungsan(jungsanData)

	Set rstJson = new aspJson
	rstJson.loadJson(conResult)
    ''response.write conResult

	JungsanTossPay = ""
	if rstJson.data("result") = "ERR" then
		'// 에러
		response.write "S_ERR|UNKNOWN"
		dbget.close()	:	response.End
	else
		PGgubun = "toss"
		PGuserid = "toss"
		sitename = "10x10"
		etcPoint = 0

		sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
		sqlStr = sqlStr & " where PGgubun = '" & PGgubun & "' " & VbCRLF
		''response.write sqlStr
		dbget.execute sqlStr

		for each item in rstJson.data("transactionList")
			Set itm = rstJson.data("transactionList").item(item)

			select case itm.item("transactionType")
				case "PAY"
					'// 결제
    				PGkey			= itm.item("payToken")
    				PGCSkey			= ""
    				appDivCode 		= "A"
    				appDate 		= itm.item("settleDate")
    				appDate 		= "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + "'"
    				cancelDate		= "NULL"
				case "REFUND"
					'// 환불
    				PGkey			= itm.item("payToken")
    				PGCSkey			= itm.item("transactionId")
    				appDivCode 		= "R"
    				appDate			= "NULL"
    				cancelDate 		= itm.item("settleDate")
    				cancelDate 		= "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + "'"
				case "FEE"
					'// 수수료 기타
					appDivCode		= "E"
				case else
					'//
					appDivCode		= "E"
			end select

			if appDivCode = "A" or appDivCode = "R" then
				select case itm.item("payMethod")
					case "TOSS_MONEY"
						appMethod 		= "20"
					case "CARD"
						appMethod 		= "100"
					case else
						appMethod 		= "ERR"
				end select

				appPrice 		= itm.item("amount")
    			commPrice		= itm.item("feeVatSum")
    			commVatPrice	= itm.item("vat")
    			commPrice 		= commPrice - commVatPrice
    			jungsanPrice 	= appPrice + commPrice + commVatPrice
				ipkumdate		= itm.item("dueDate")
    			ipkumdate 		= Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(Left(ipkumdate, 8), 2)

    			sqlStr = " if NOT Exists( select top 1 * from db_temp.dbo.tbl_onlineApp_log_tmp where PGgubun='" + CStr(PGgubun) + "' and PGkey='" + CStr(PGkey) + "' and PGCSkey='" + CStr(PGCSkey) + "')"&vbCRLF
    			sqlStr = sqlStr + " BEGIN"&vbCRLF
    			sqlStr = sqlStr + " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, etcPoint) "&vbCRLF
    			sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "', " & etcPoint & ") "&vbCRLF
    			sqlStr = sqlStr + " END"&vbCRLF

    			''response.write sqlStr + "<br>"
    			dbget.execute sqlStr
			end if
		next

		sqlStr = " select count(*) as cnt from db_temp.dbo.tbl_onlineApp_log_tmp "
		sqlStr = sqlStr & " where PGgubun = '" & PGgubun & "' " & VbCRLF

        ''response.write sqlStr
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
        IF not rsget.EOF THEN
    		retLength = rsget("cnt")
    	END IF
    	rsget.close

		if (retLength<1) then
			response.write "S_ERR|NO_DATA"
			dbget.close()	:	response.End
		else
			sqlStr = " update t "
			sqlStr = sqlStr & " set t.orderserial = m.orderserial, t.sitename = (case when m.rdsite = 'mobile' or m.rdsite = 'app_wish2' then '10x10mobile' else '10x10' end) "
			sqlStr = sqlStr & " from "
			sqlStr = sqlStr & " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
			sqlStr = sqlStr & " 	join [db_order].[dbo].[tbl_order_temp] m on t.pgkey = m.P_TID "
			sqlStr = sqlStr & " where t.pggubun = '" & PGgubun & "' "
			dbget.execute sqlStr

			sqlStr = " update t "
			sqlStr = sqlStr & " set t.appDate = m.regdate "
			sqlStr = sqlStr & " from "
			sqlStr = sqlStr & " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
			sqlStr = sqlStr & " 	join [db_order].[dbo].[tbl_order_temp] m on t.pgkey = m.P_TID "
			sqlStr = sqlStr & " where t.pggubun = '" & PGgubun & "' and t.appDivCode = 'A' "
			dbget.execute sqlStr

			sqlStr = " update T "
			sqlStr = sqlStr & " set T.csasid = a.id, T.cancelDate = a.finishdate "
			sqlStr = sqlStr & " from "
			sqlStr = sqlStr & " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr & " 	join [db_cs].[dbo].[tbl_new_as_list] a "
			sqlStr = sqlStr & " 	on "
			sqlStr = sqlStr & " 		1 = 1 "
			sqlStr = sqlStr & " 		and T.orderserial = a.orderserial "
			sqlStr = sqlStr & " 		and T.PGCSkey = a.orderserial + '_' + convert(varchar, a.id) "
			sqlStr = sqlStr & " where t.pggubun = '" & PGgubun & "' "
			dbget.execute sqlStr

			sqlStr = " update r "
			sqlStr = sqlStr & " set r.PGCSkey = 'CANCELALL', r.appDivCode = 'C' "
			sqlStr = sqlStr & " from "
			sqlStr = sqlStr & " 	db_temp.dbo.tbl_onlineApp_log_tmp a "
			sqlStr = sqlStr & " 	join db_temp.dbo.tbl_onlineApp_log_tmp r "
			sqlStr = sqlStr & " 	on "
			sqlStr = sqlStr & " 		1 = 1 "
			sqlStr = sqlStr & " 		and a.PGgubun = r.PGgubun "
			sqlStr = sqlStr & " 		and a.PGkey = r.PGkey "
			sqlStr = sqlStr & " 		and a.appdivcode = 'A' "
			sqlStr = sqlStr & " 		and r.appdivcode = 'R' "
			sqlStr = sqlStr & " 		and a.appPrice = r.appPrice*-1 "
			sqlStr = sqlStr & " where a.pggubun = '" & PGgubun & "' "
			dbget.execute sqlStr

			sqlStr = " update r "
			sqlStr = sqlStr & " set r.PGCSkey = 'CANCELALL', r.appDivCode = 'C' "
			sqlStr = sqlStr & " from "
			sqlStr = sqlStr & " 	[db_order].[dbo].[tbl_onlineApp_log] a "
			sqlStr = sqlStr & " 	join db_temp.dbo.tbl_onlineApp_log_tmp r "
			sqlStr = sqlStr & " 	on "
			sqlStr = sqlStr & " 		1 = 1 "
			sqlStr = sqlStr & " 		and a.PGgubun = r.PGgubun "
			sqlStr = sqlStr & " 		and a.PGkey = r.PGkey "
			sqlStr = sqlStr & " 		and a.appdivcode = 'A' "
			sqlStr = sqlStr & " 		and r.appdivcode = 'R' "
			sqlStr = sqlStr & " 		and a.appPrice = r.appPrice*-1 "
			sqlStr = sqlStr & " where a.pggubun = '" & PGgubun & "' "
			dbget.execute sqlStr

			sqlStr = " update m "
			sqlStr = sqlStr & " set m.ipkumdate = T.appDate "
			sqlStr = sqlStr & " from "
			sqlStr = sqlStr & " 	[db_temp].[dbo].[tbl_onlineApp_log_tmp] T "
			sqlStr = sqlStr & " 	join [db_order].[dbo].[tbl_order_master] m on T.orderserial = m.orderserial "
			sqlStr = sqlStr & " where "
			sqlStr = sqlStr & " 	1 = 1 "
			sqlStr = sqlStr & " 	and T.PGgubun = '" & PGgubun & "' "
			sqlStr = sqlStr & " 	and T.appDivCode = 'A' "
			sqlStr = sqlStr & " 	and DateDiff(day, T.appDate, m.ipkumdate) <> 0 "
			dbget.execute sqlStr

			sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate, orderserial, etcPoint, csasid) "
			sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121), t.orderserial, t.etcPoint, t.csasid "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
			sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
			sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
			sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and l.idx is NULL "
			sqlStr = sqlStr + " 	and t.PGgubun = '" + CStr(PGgubun) + "' "
			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
			''response.write sqlStr + "<br>"
			dbget.execute sqlStr

			sqlStr = " update l "
			sqlStr = sqlStr + " set l.ipkumdate = t.ipkumdate, l.commPrice = t.commPrice, l.commVatPrice = t.commVatPrice, l.jungsanPrice = t.jungsanPrice, l.PGmeachulDate = convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
			sqlStr = sqlStr + " 	join db_temp.dbo.tbl_onlineApp_log_tmp t "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
			sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
			sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and l.commPrice = 0 "
			sqlStr = sqlStr + " 	and t.PGgubun = '" + CStr(PGgubun) + "' "
            dbget.execute sqlStr

			response.write "S_OK"
		end if

		JungsanTossPay = rstJson.data("nextCursor")
	end if
end function


'// ============================================================================
dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    response.write ref
    response.end
end if


'// ============================================================================
Dim currDT : currDT = Left(dateAdd("d",-1,Now()), 10)
dim yyyymmdd, mode, nextCursor

mode  = requestCheckVar(request("mode"), 30)
yyyymmdd  = requestCheckVar(request("yyyymmdd"), 30)

if (yyyymmdd = "") then
	yyyymmdd = currDT
end if

yyyymmdd = Replace(yyyymmdd, "-", "")
nextCursor = ""

select case mode
	case "settle"
		do
			nextCursor = JungsanTossPay(yyyymmdd, "SETTLE", nextCursor)
		loop while (nextCursor <> "")
	case "due"
		do
			nextCursor = JungsanTossPay(yyyymmdd, "DUE", nextCursor)
		loop while (nextCursor <> "")
	case else
		response.write "잘못된 접근입니다."
end select

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
