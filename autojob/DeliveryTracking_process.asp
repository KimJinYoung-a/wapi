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

'// ===========================================================================
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.70","61.252.133.10","61.252.133.80","110.93.128.114","110.93.128.113","61.252.133.67","192.168.1.67", "52.79.95.197")
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
    dbDatamart_dbget.Close()
    response.end
end if


'// ===========================================================================
dim HTTP_MODE, mode

HTTP_MODE = "GET"
mode = Left(request("mode"), 32)
if (request.Form("mode") <> "") then
	HTTP_MODE = "POST"
	mode = Left(request.Form("mode"), 32)
	data = request.Form("data")
end if


'// ===========================================================================
dim i, j, k, sqlStr, affectedRows
dim json, order, orderlist()
dim OrderArr, data, item

select case mode
	case "getorder"

		sqlStr = " select top 300 orderserial, songjangdiv, songjangno "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_datamart].[dbo].[tbl_DeliveryTrackingList] "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and realDeliveryDate is NULL "
		sqlStr = sqlStr + " 	and checkCnt < 5 "
		sqlStr = sqlStr + " 	and songjangdiv in (4, 2, 18, 8, 3, 1, 39, 41) "		'// 21, 26, 28, 29, 31, 33, 34, 35, 37
		sqlStr = sqlStr + " 	and songjangdiv <> 2 "											'// 롯대택배 제외
		sqlStr = sqlStr + " 	and DateDiff(day, regdate, getdate()) >= checkCnt "
		sqlStr = sqlStr + " 	and beasongdate >= Convert(varchar(10), DateAdd(day, -14, getdate()), 121) "
		sqlStr = sqlStr + " order by checkCnt, idx "
		''response.write sqlStr
        dbDatamart_rsget.CursorLocation = adUseClient
        dbDatamart_rsget.Open sqlStr,dbDatamart_dbget,adOpenForwardOnly, adLockReadOnly

		if Not dbDatamart_rsget.Eof then
			if dbDatamart_rsget.RecordCount > 20 then
				'// 5개 미만이면 오류내역일 수 있음, 롯데택배로 넘어가기.
				OrderArr = dbDatamart_rsget.GetRows
			end if
		end if
		dbDatamart_rsget.Close

		if Not isarray(OrderArr) then
			sqlStr = " select top 5 orderserial, songjangdiv, songjangno "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " [db_datamart].[dbo].[tbl_DeliveryTrackingList] "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and realDeliveryDate is NULL "
			sqlStr = sqlStr + " 	and checkCnt < 5 "
			sqlStr = sqlStr + " 	and songjangdiv in (4, 2, 18, 8, 3, 1, 39, 41, 21, 26, 28, 29, 31, 33, 34, 35, 37) "
			sqlStr = sqlStr + " 	and songjangdiv = 2 "											'// 롯대택배만(5개씩)
			sqlStr = sqlStr + " 	and DateDiff(day, regdate, getdate()) >= checkCnt "
			sqlStr = sqlStr + " 	and beasongdate >= Convert(varchar(10), DateAdd(day, -14, getdate()), 121) "
			sqlStr = sqlStr + " order by checkCnt, idx "
			''response.write sqlStr
	        dbDatamart_rsget.CursorLocation = adUseClient
	        dbDatamart_rsget.Open sqlStr,dbDatamart_dbget,adOpenForwardOnly, adLockReadOnly

			if Not dbDatamart_rsget.Eof then
				OrderArr = dbDatamart_rsget.GetRows
			end if
			dbDatamart_rsget.Close
		end if

		if Not isarray(OrderArr) then
			Set json = jsObject()
			json("count") = 0
			Response.Write toJSON(json)
			dbDatamart_dbget.Close : Response.end
		end if

		redim orderlist(UBound(OrderArr, 2))

		for i = 0 to UBound(OrderArr, 2)
			Set order = jsObject()
			order("orderserial") = OrderArr(0, i)
			order("songjangdiv") = OrderArr(1, i)
			order("songjangno") = OrderArr(2, i)
			Set orderlist(i) = order
		next

		Set json = jsObject()
		if (UBound(orderlist) >= 0) then
			json("count") = UBound(orderlist) + 1
		else
			json("count") = 0
		end if

		json("data") = orderlist
		Response.Write toJSON(json)
	case "savedeliveryinfo"
		Set json = new aspJSON
		json.loadJSON(data)

		if (json.data("count") > 0) then
			for each order in json.data("data")
				set order = json.data("data").item(order)

				sqlStr = " update [db_datamart].[dbo].[tbl_DeliveryTrackingList] "
				sqlStr = sqlStr + " set realDeliveryDate = (case when '" & order.item("beasongStartDate") & "' <> '' then '" & order.item("beasongStartDate") & "' else NULL end), findDate = (case when '" & order.item("beasongStartDate") & "' <> '' then Convert(varchar(10), getdate(), 121) else NULL end), checkCnt = checkCnt + 1, lastupdate = getdate() "
				sqlStr = sqlStr + " where "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and orderserial = '" & order.item("orderserial") & "' "
				sqlStr = sqlStr + " 	and songjangdiv = " & order.item("songjangdiv") & " "
				sqlStr = sqlStr + " 	and RTrim(LTrim(songjangno)) = '" & order.item("songjangno") & "' "
				sqlStr = sqlStr + " 	and realDeliveryDate is NULL "
				dbDatamart_dbget.Execute sqlStr


				''Response.Write  order.item("orderserial")

			next

			Response.Write "{""status"": ""OK"",""HTTP_MODE"": """ & HTTP_MODE & """}"
		else
			Response.Write "{""status"": ""FAIL"",""HTTP_MODE"": """ & HTTP_MODE & """}"
		end if
	case else
		response.write "Hello, World! ERR"
end select

%>
<!-- #include virtual="/lib/db/dbDatamartclose.asp" -->