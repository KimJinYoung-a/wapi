<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbDatamartopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/util/JSON_UTIL_0.1.1.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/aspJSON1.17.asp"-->
<%

dim iAnalCon : set iAnalCon = CreateObject("ADODB.Connection")

'// ===========================================================================
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.70","61.252.133.10","61.252.133.80","110.93.128.114","110.93.128.113","61.252.133.67","192.168.1.67", "61.252.133.18")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

dim ref : ref = Request.ServerVariables("REMOTE_ADDR")
''response.write ref

if (Not CheckVaildIP(ref)) then
    dbDatamart_dbget.Close()
    response.end
end if

'// ===========================================================================
dim mode, host, gubun, data
dim i, j, k, sqlStr, affectedRows

mode = Left(request("mode"), 32)
host = Left(request("host"), 32)
gubun = Left(request("gubun"), 32)
data = request("data")

'response.write mode & vbcrlf
'response.write host & vbcrlf
'response.write gubun & vbcrlf
'response.write data & vbcrlf

select case mode
	case "savesrchuniq"
		'// 저장하기
		sqlStr = ""
		data = Split(data, ",")
		if (gubun = "searchcnt") and UBound(data) = 3 then
			sqlStr = " exec [db_analyze_data_raw].[dbo].[sp_TEN_ELK_Get_UniqueSearchUser] '" & data(0) & "', '" & host & "', '" & gubun & "', " & data(1) & ", " & data(2) & ", " & data(3) & " "
			iAnalCon.Open Application("db_analyze")
			iAnalCon.Execute sqlStr
			iAnalCon.Close
		end if

		if (gubun = "itemviewcnt") and UBound(data) = 3 then
			sqlStr = " exec [db_analyze_data_raw].[dbo].[sp_TEN_ELK_Get_UniqueSearchUser] '" & data(0) & "', '" & host & "', '" & gubun & "', " & data(1) & ", " & data(2) & ", " & data(3) & " "
			iAnalCon.Open Application("db_analyze")
			iAnalCon.Execute sqlStr
			iAnalCon.Close
		end if

		response.write sqlStr
	case else
		response.write "ERROR"
end select

%>