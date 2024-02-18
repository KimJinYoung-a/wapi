<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/ssg/ssgItemcls.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->

<%
'' TLS 1.2를 지원하지 않는 서버가 있는듯함..
const Option_TLS12 = 2048
const Option_TLS1 = 512
const Option_TLS = 128

'' 1. 배송지시목록조회
'' 2. 주문 확인 처리
'' 3. 촐고 대상목록 조회

'CONST ssgAPIURL = "http://eapi.ssgadm.com"
'CONST ssgSSLAPIURL = "https://eapi.ssgadm.com"
'CONST ssgApiKey = "18a8d870-12a7-4b36-afaf-1e9d38e2b988"

dim mode, yyyymmdd
dim submode, tempcsidx, CSDetailKey, songjangdiv, songjangno, resultCode, ishppNo, ishppSeq, iwblNo, idelicoVenId, ishppTypeCd, ishppTypeDtlCd, extsongjangdiv, iresultDesc, itemno
dim OutMallCurrState, outMallorderSerial, OutMallOrderSerialArr
dim i, strSql, affectedRows, divcd, orderserial, asid, tmpasid, resellPsblYn, retImptMainCd
mode	= requestCheckVar(html2db(request("mode")),32)
yyyymmdd	= requestCheckVar(html2db(request("yyyymmdd")),32)
tempcsidx	= requestCheckVar(html2db(request("tempcsidx")),32)
outMallorderSerial	= requestCheckVar(html2db(request("outMallorderSerial")),32)
divcd	= requestCheckVar(html2db(request("divcd")),32)
orderserial	= requestCheckVar(html2db(request("orderserial")),32)
asid	= requestCheckVar(html2db(request("asid")),32)

function fnMatchCs(ioutmallorderserial)
    dim affectedRow, strSql

    strSql = " update T "
	strSql = strSql & " set T.asid = a.id "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_list] a "
	strSql = strSql & " 	on "
	strSql = strSql & " 		1 = 1 "
	strSql = strSql & " 		and a.orderserial = T.OrderSerial "
	strSql = strSql & " 		and a.deleteyn = 'N' "
	strSql = strSql & " 		and ( "
	strSql = strSql & " 			(T.divcd = 'A004' and a.divcd in ('A004', 'A010', 'A008')) "
	strSql = strSql & " 			or "
	strSql = strSql & " 			(T.divcd = 'A011' and a.divcd in ('A011', 'A012', 'A112', 'A112')) "
    strSql = strSql & " 			or "
    strSql = strSql & " 			(T.divcd = 'A008' and a.divcd in ('A008', 'A004', 'A010')) "
	strSql = strSql & " 		) "
	strSql = strSql & " 		and a.id not in ( "
	strSql = strSql & " 			select T.asid "
	strSql = strSql & " 			from "
	strSql = strSql & " 				[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " 			where "
	strSql = strSql & " 				1 = 1 "
	strSql = strSql & " 				and T.SellSite = 'ssg' "
	strSql = strSql & " 				and T.OutMallOrderSerial = '" & ioutmallorderserial & "' "
	strSql = strSql & " 				and T.asid is not NULL "
	strSql = strSql & " 		) "
	strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_detail] d "
	strSql = strSql & " 	on "
	strSql = strSql & " 		1 = 1 "
	strSql = strSql & " 		and a.id = d.masterid "
	strSql = strSql & " 		and d.itemid = T.ItemID "
	strSql = strSql & " 		and d.itemoption = T.itemoption "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = 'ssg' "
	strSql = strSql & " 	and T.OutMallOrderSerial = '" & ioutmallorderserial & "' "
	strSql = strSql & " 	and T.asid is NULL "
    strSql = strSql & " 	and IsNull(T.outmallCurrState, 'B001') <> 'B008' "
    dbget.Execute strSql, affectedRow

    fnMatchCs = affectedRow
end function

function fnUnmatchDeletedCS(ioutmallorderserial)
    strSql = " update T "
	strSql = strSql & " set T.asid = NULL "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_list] a on T.asid = a.id and a.deleteyn = 'Y' "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = 'ssg' "
	strSql = strSql & " 	and T.OutMallOrderSerial = '" & OutMallOrderSerial & "' "
	strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008') "
	strSql = strSql & " 	and T.orderserial is not NULL "
    dbget.Execute strSql

    strSql = " update T "
	strSql = strSql & " set T.asid = NULL "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = 'ssg' "
	strSql = strSql & " 	and T.OutMallOrderSerial = '" & OutMallOrderSerial & "' "
	strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008') "
	strSql = strSql & " 	and T.orderserial is not NULL "
    strSql = strSql & " 	and T.asid is not NULL "
    strSql = strSql & " 	and T.outmallCurrState = 'B008' "
    dbget.Execute strSql
end function

Dim istyyyymmdd, iedyyyymmdd
    iedyyyymmdd = replace(LEFT(now(),10),"-","")
    istyyyymmdd = replace(dateadd("d",-5,LEFT(iedyyyymmdd,4)&"-"&Mid(iedyyyymmdd,5,2)&"-"&Mid(iedyyyymmdd,7,2)),"-","")

if (mode = "oldexchangechulgo") then
	istyyyymmdd = yyyymmdd

	for i = 0 to 5
		iedyyyymmdd = Left(DateAdd("d", 6, istyyyymmdd), 10)
		response.write istyyyymmdd & " ~ " & iedyyyymmdd & "<br />"
		call getSsgExchangeChulgoList(Replace(istyyyymmdd, "-", ""),Replace(iedyyyymmdd, "-", ""))
		istyyyymmdd = Left(DateAdd("d", 7, istyyyymmdd), 10)
	next

	dbget.close() : response.end

elseif (mode = "matchCsAs") then

	'// 교환출고 CS내역 매핑
	strSql = " update c "
	strSql = strSql & " set c.asid = T.asid "
	strSql = strSql & " from "
	strSql = strSql & " 	db_temp.dbo.tbl_xSite_TMPCS c "
	strSql = strSql & " 	join ( "
	strSql = strSql & " 		 select c.idx, max(a.id) as asid, count(a.id) as cnt "
	strSql = strSql & " 		 from "
	strSql = strSql & " 		 	db_temp.dbo.tbl_xSite_TMPCS c WITH(NOLOCK) "
	strSql = strSql & " 			join [db_cs].[dbo].[tbl_new_as_list] a WITH(NOLOCK) "
	strSql = strSql & " 			on "
	strSql = strSql & " 				1 = 1 "
	strSql = strSql & " 				and a.divcd = 'A000' "
	strSql = strSql & " 				and a.currstate = 'B007' "
	strSql = strSql & " 				and a.deleteyn = 'N' "
	strSql = strSql & " 				and a.songjangdiv <> '0' "
	strSql = strSql & " 				and a.songjangno <> '' "
	strSql = strSql & " 			join [db_cs].[dbo].[tbl_new_as_detail] d WITH(NOLOCK) "
	strSql = strSql & " 			on "
	strSql = strSql & " 				1 = 1 "
	strSql = strSql & " 				and a.id = d.masterid "
	strSql = strSql & " 				and d.itemid = c.itemid "
	strSql = strSql & " 				and d.itemoption = c.itemoption "
	strSql = strSql & " 		where "
	strSql = strSql & " 			1 = 1 "
	strSql = strSql & " 			and c.SellSite = '" & CMALLNAME & "' "
	strSql = strSql & " 			and c.divcd = 'A000' "
	strSql = strSql & " 			and c.orderserial = a.orderserial "
	strSql = strSql & " 			and c.deleteyn = 'N' "
	strSql = strSql & " 			and c.currstate = 'B001' "
	strSql = strSql & " 			and c.OutMallCurrState <> 'B007' "
	strSql = strSql & " 			and c.asid is NULL "
	strSql = strSql & " 			and a.id not in ( "				'// 기존 매칭된 CS건 제외
	strSql = strSql & " 				select c.asid "
	strSql = strSql & " 				from "
	strSql = strSql & " 					db_temp.dbo.tbl_xSite_TMPCS c WITH(NOLOCK) "
	strSql = strSql & " 					join ( "
	strSql = strSql & " 						select distinct c.SellSite, c.OrderSerial "
	strSql = strSql & " 						from "
	strSql = strSql & " 							db_temp.dbo.tbl_xSite_TMPCS c WITH(NOLOCK) "
	strSql = strSql & " 						where "
	strSql = strSql & " 							1 = 1 "
	strSql = strSql & " 							and c.SellSite = '" & CMALLNAME & "' "
	strSql = strSql & " 							and c.divcd = 'A000' "
	strSql = strSql & " 							and c.asid is NULL "
	strSql = strSql & " 							and c.deleteyn = 'N' "
	strSql = strSql & " 							and c.currstate = 'B001' "
	strSql = strSql & " 							and c.OutMallCurrState <> 'B007' "
	strSql = strSql & " 					) T "
	strSql = strSql & " 					on "
	strSql = strSql & " 						1 = 1 "
	strSql = strSql & " 						and c.SellSite = T.SellSite "
	strSql = strSql & " 						and c.OrderSerial = T.OrderSerial "
	strSql = strSql & " 				where "
	strSql = strSql & " 					c.asid is not NULL "
	strSql = strSql & " 			) "
	strSql = strSql & " 		group by "
	strSql = strSql & " 			c.idx "
	strSql = strSql & " 	) T on T.idx = c.idx and T.cnt = 1 and c.asid is NULL "
	dbget.Execute strSql

elseif (mode = "sendSongJang") then

	strSql = " select top 1 c.idx as tempcsidx, c.CSDetailKey, a.songjangdiv, a.songjangno, c.OutMallCurrState, d.confirmitemno as itemno "
	strSql = strSql & " 	from db_temp.dbo.tbl_xSite_TMPCS c "
	strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_list] a "
	strSql = strSql & " 	on "
	strSql = strSql & " 		c.asid = a.id "
	strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_detail] d WITH(NOLOCK) "
	strSql = strSql & " 	on "
	strSql = strSql & " 		1 = 1 "
	strSql = strSql & " 		and a.id = d.masterid "
	strSql = strSql & " 		and d.itemid = c.itemid "
	strSql = strSql & " 		and d.itemoption = c.itemoption "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and c.SellSite = '" & CMALLNAME & "' "
	strSql = strSql & " 	and c.divcd = 'A000' "
	strSql = strSql & " 	and c.orderserial = a.orderserial "
	strSql = strSql & " 	and (c.currstate = 'B001' or c.idx = 155316) "
	strSql = strSql & " 	and c.OutMallCurrState in ('B001', 'B004') "
	strSql = strSql & " 	and c.deleteyn = 'N' "
	strSql = strSql & " order by "
	strSql = strSql & " 	IsNull(c.checkCount,0) "

	tempcsidx = -1
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
		tempcsidx 	= rsget("tempcsidx")
		CSDetailKey = rsget("CSDetailKey")
		songjangdiv = rsget("songjangdiv")
		songjangno 	= rsget("songjangno")
		OutMallCurrState 	= rsget("OutMallCurrState")
		itemno 		= rsget("itemno")
	end if
	rsget.Close

	if tempcsidx = -1 then
		rw "OK : NONE"
		dbget.close : response.end
	end if

	ishppNo = ""
	CSDetailKey = Split(CSDetailKey, "_")
	if UBound(CSDetailKey) = 1 then
		ishppNo = CSDetailKey(0)
		ishppSeq = CSDetailKey(1)
	end if

	if ishppNo = "" then
		rw "ERR : CS Detail Key None"
		dbget.close : response.end
	end if

	if (OutMallCurrState = "B001") then
		'// 송장전송
		extsongjangdiv = TEN_TenDlvCode2CommonDlvCode(CMALLNAME, songjangdiv)

		iwblNo = songjangno
		idelicoVenId = extsongjangdiv
		if (songjangdiv = "98") then
			'퀵서비스->직배송
			ishppTypeCd = "10"
			ishppTypeDtlCd = "14"
		else
			ishppTypeCd = "20"
			ishppTypeDtlCd = "22"
		end if
		resultCode = sendSsgSendSongjang(ishppNo, ishppSeq, iwblNo, idelicoVenId, ishppTypeCd, ishppTypeDtlCd)

		if (resultCode = "00") then
			strSql = " update "
			strSql = strSql & " db_temp.dbo.tbl_xSite_TMPCS "
			strSql = strSql & " set OutMallCurrState = 'B004' "
			strSql = strSql & " where idx = " & tempcsidx & " and OutMallCurrState = 'B001' "
			dbget.Execute strSql

			OutMallCurrState = "B004"
		end if
	end if

	if (OutMallCurrState = "B004") then
		resultCode = sendSsgChulgoFinish(ishppNo, ishppSeq, itemno, iresultDesc)

		if (resultCode = "00") then
			strSql = " update "
			strSql = strSql & " db_temp.dbo.tbl_xSite_TMPCS "
			strSql = strSql & " set OutMallCurrState = 'B006', outMallFinishDate = getdate() "
			strSql = strSql & " where idx = " & tempcsidx & " and OutMallCurrState < 'B006' "
			dbget.Execute strSql

			OutMallCurrState = "B006"
		end if
	end if

	if (OutMallCurrState = "B006") then
		rw "OK"
	else
		rw "ERR : " & resultCode
		rw "ERR MSG : " & iresultDesc
		rw "ishppTypeCd : " & ishppTypeCd
	end if

elseif (mode = "finishChulgo") then

	strSql = " select top 1 c.idx as tempcsidx, c.CSDetailKey, a.songjangdiv, a.songjangno, c.OutMallCurrState "
	strSql = strSql & " 	from db_temp.dbo.tbl_xSite_TMPCS c "
	strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_list] a "
	strSql = strSql & " 	on "
	strSql = strSql & " 		c.asid = a.id "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and c.SellSite = '" & CMALLNAME & "' "
	strSql = strSql & " 	and c.divcd = 'A000' "
	strSql = strSql & " 	and c.orderserial = a.orderserial "
	strSql = strSql & " 	and (c.currstate = 'B001' or c.idx = 155316) "
	strSql = strSql & " 	and c.OutMallCurrState = 'B006' "
	strSql = strSql & " 	and DateDiff(day, c.outMallFinishDate, getdate()) > 0 "
	strSql = strSql & " 	and c.deleteyn = 'N' "
	strSql = strSql & " order by "
	strSql = strSql & " 	IsNull(c.checkCount,0) "

	tempcsidx = -1
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
		tempcsidx 	= rsget("tempcsidx")
		CSDetailKey = rsget("CSDetailKey")
	end if
	rsget.Close

	if tempcsidx = -1 then
		rw "OK : NONE"
		dbget.close : response.end
	end if

	ishppNo = ""
	CSDetailKey = Split(CSDetailKey, "_")
	if UBound(CSDetailKey) = 1 then
		ishppNo = CSDetailKey(0)
		ishppSeq = CSDetailKey(1)
	end if

	if ishppNo = "" then
		rw "ERR : CS Detail Key None"
		dbget.close : response.end
	end if

	resultCode = sendSsgBeasongFinish(ishppNo, ishppSeq, iresultDesc)

	if (resultCode = "00") then
		strSql = " update "
		strSql = strSql & " db_temp.dbo.tbl_xSite_TMPCS "
		strSql = strSql & " set OutMallCurrState = 'B007' "
		strSql = strSql & " where idx = " & tempcsidx & " and OutMallCurrState < 'B007' "
		dbget.Execute strSql

		OutMallCurrState = "B007"
	end if

	if (OutMallCurrState = "B006") then
		rw "OK"
	else
		rw "ERR : " & resultCode
		rw "ERR MSG : " & iresultDesc
	end if

elseif (mode = "chkBatchMatchCS") then
    '// 접수CS 체크 배치

    strSql = " update T "
	strSql = strSql & " set T.asid = NULL "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = 'ssg' "
	''strSql = strSql & " 	and T.OutMallOrderSerial = '" & ioutmallorderserial & "' "
	strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008') "
	strSql = strSql & " 	and T.orderserial is not NULL "
    strSql = strSql & " 	and T.asid is not NULL "
    strSql = strSql & " 	and T.outmallCurrState = 'B008' "
    dbget.Execute strSql

    strSql = " select distinct top 100 T.OutMallOrderSerial "
	strSql = strSql & " from "
	strSql = strSql & " [db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = 'ssg' "
	strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008') "
    strSql = strSql & " 	and T.orderserial is not NULL "
	strSql = strSql & " 	and T.asid is NULL "
    ''strSql = strSql & " 	and T.regdate < convert(varchar(10), getdate(), 121) "
    strSql = strSql & " 	and IsNull(T.asidCheckDT, DateAdd(day, -1, getdate())) < DateAdd(hour, -1, getdate()) "
    ''strSql = strSql & " order by newid() "
    ''rw strSql

    OutMallOrderSerialArr = ""

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
        do until rsget.eof
            OutMallOrderSerialArr = OutMallOrderSerialArr + "," + rsget("OutMallOrderSerial")
            rsget.moveNext
        loop
	end if
	rsget.Close

    Response.CharSet="euc-kr"
    Session.codepage="949"
    Response.codepage="949"
    Response.ContentType="text/html;charset=euc-kr"

    '// git 업로드 확인
    if OutMallOrderSerialArr = "" then
        rw "내역없음"
        dbget.close() : response.end
    end if

    affectedRows = 0
    OutMallOrderSerialArr = Split(OutMallOrderSerialArr, ",")
    for i = 0 to UBound(OutMallOrderSerialArr)
        OutMallOrderSerial = OutMallOrderSerialArr(i)
        if OutMallOrderSerial <> "" then
            affectedRows = fnMatchCs(OutMallOrderSerial)
            rw OutMallOrderSerial & " : " & affectedRows & " 건 반영됨"

            if affectedRows = 0 then
                strSql = " update T "
                strSql = strSql & " set T.asidCheckDT = getdate() "
	            strSql = strSql & " from "
	            strSql = strSql & " [db_temp].[dbo].[tbl_xSite_TMPCS] T "
	            strSql = strSql & " where "
	            strSql = strSql & " 	1 = 1 "
	            strSql = strSql & " 	and T.SellSite = 'ssg' "
                strSql = strSql & " 	and T.OutMallOrderSerial = '" & OutMallOrderSerial & "' "
	            strSql = strSql & " 	and T.divcd in ('A004', 'A011', 'A008') "
	            strSql = strSql & " 	and T.asid is NULL "
                dbget.Execute strSql
            end if
        end if
    next

    dbget.close() : response.end

elseif (mode = "chkBatchExtCsState") then
    '// 제휴CS 상태 체크(배치)
    dim divcdArr

    strSql = " select distinct top 100 T.divcd, T.OutMallOrderSerial "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	''strSql = strSql & " 	join [db_cs].[dbo].[tbl_new_as_list] a on T.asid = a.id and a.currstate = 'B007' and a.deleteyn = 'N' "		'// 접수이전 내역도 체크
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = 'ssg' "
	strSql = strSql & " 	and T.divcd in ('A004', 'A011') "
	strSql = strSql & " 	and T.orderserial is not NULL "
	''strSql = strSql & " 	and T.regdate < convert(varchar(10), DateAdd(day, -1, getdate()), 121) "
	strSql = strSql & " 	and T.regdate >= convert(varchar(10), DateAdd(day, -80, getdate()), 121) "
    strSql = strSql & " 	and IsNull(T.outmallCheckDT, DateAdd(day, -1, getdate())) < DateAdd(hour, -1, getdate()) "
	strSql = strSql & " 	and IsNull(T.OutMallCurrState, 'B001') < 'B007' "
    ''rw strSql

    OutMallOrderSerialArr = ""
    divcdArr = ""

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
        do until rsget.eof
            OutMallOrderSerialArr = OutMallOrderSerialArr + "," + rsget("OutMallOrderSerial")
            divcdArr = divcdArr & rsget("divcd") & ","
            rsget.moveNext
        loop
	end if
	rsget.Close

    Response.CharSet="euc-kr"
    Session.codepage="949"
    Response.codepage="949"
    Response.ContentType="text/html;charset=euc-kr"

    '// git 업로드 확인
    if OutMallOrderSerialArr = "" then
        rw "내역없음"
        dbget.close() : response.end
    end if

    affectedRows = 0
    OutMallOrderSerialArr = Split(OutMallOrderSerialArr, ",")
    divcdArr = Split(divcdArr, ",")
    for i = 0 to UBound(OutMallOrderSerialArr)
        OutMallOrderSerial = OutMallOrderSerialArr(i)
        divcd = divcdArr(i)
        if OutMallOrderSerial <> "" then
            affectedRows = getSsgExchangeReturnOne(divcd, OutMallOrderSerial, "")
            rw OutMallOrderSerial & " : " & affectedRows & " 건 반영됨"

            if affectedRows = 0 then
                strSql = " update T "
                strSql = strSql & " set T.outmallCheckDT = getdate() "
	            strSql = strSql & " from "
	            strSql = strSql & " [db_temp].[dbo].[tbl_xSite_TMPCS] T "
	            strSql = strSql & " where "
	            strSql = strSql & " 	1 = 1 "
	            strSql = strSql & " 	and T.SellSite = 'ssg' "
                strSql = strSql & " 	and T.OutMallOrderSerial = '" & OutMallOrderSerial & "' "
	            strSql = strSql & " 	and T.divcd = '" & divcd & "' "
	            strSql = strSql & " 	and T.asid is not NULL "
                dbget.Execute strSql
            end if
        end if
    next

    dbget.close() : response.end

elseif (mode = "chkExtCsStateOne") then
    '// 제휴CS 상태 체크
    Call getSsgExchangeReturnOne(divcd, OutMallOrderSerial, "")
    rw "OK"

elseif (mode = "chkMatchCS") then
    '// 접수CS 체크

    Call fnUnmatchDeletedCS(OutMallOrderSerial)

    affectedRows = fnMatchCs(OutMallOrderSerial)

    rw affectedRows & " 건 반영됨"
    dbget.close() : response.end

elseif (mode = "sendReturnConfirm") then
    '// 회수확인처리
    ''resultCode = sendSsgReturnConfirm(shppNo, shppSeq, procItemQty, resellPsblYn, retImptMainCd, ByRef resultDesc)
    '//<option value="0000033028">기타택배사</option>
    '// sendSsgReturnConfirm(shppNo, shppSeq, procItemQty, resellPsblYn, retImptMainCd, ByRef resultDesc)
elseif (mode = "sendReturnFinish") then
    '// 회수완료처리

    '// ========================================================================
    '// 1. 제휴 주문번호 구하기
    '// ========================================================================
    OutMallOrderSerial = ""

    strSql = " select top 1 T.OutMallOrderSerial "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = 'ssg' "
    strSql = strSql & " 	and T.orderserial = '" & orderserial & "' "
    ''rw strSql

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
        OutMallOrderSerial = rsget("OutMallOrderSerial")
	end if
	rsget.Close

    if (OutMallOrderSerial = "") then
        rw "에러 : 제휴 주문번호 없음"
        dbget.close() : response.end
    end if

    '// ========================================================================
    '// 2. ASID 매칭상태 확인
    '// ========================================================================
    tmpasid = ""

    strSql = " select top 1 T.asid as tmpasid, IsNull(T.OutMallCurrState, 'B001') as OutMallCurrState, d.confirmitemno as itemno, shppNo, shppSeq, a.gubun01, a.gubun02 "
	strSql = strSql & " from "
	strSql = strSql & " 	[db_temp].[dbo].[tbl_xSite_TMPCS] T "
    strSql = strSql & " 	left join [db_cs].[dbo].[tbl_new_as_list] a on T.asid = a.id and a.deleteyn = 'N' and a.id = " & asid
	strSql = strSql & " 	left join [db_cs].[dbo].[tbl_new_as_detail] d WITH(NOLOCK) "
	strSql = strSql & " 	on "
	strSql = strSql & " 		1 = 1 "
	strSql = strSql & " 		and a.id = d.masterid "
	strSql = strSql & " 		and d.itemid = T.itemid "
	strSql = strSql & " 		and d.itemoption = T.itemoption "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and T.SellSite = 'ssg' "
    strSql = strSql & " 	and T.orderserial = '" & orderserial & "' "
	strSql = strSql & " 	and T.divcd = '" & divcd & "' "
    strSql = strSql & " 	and T.asid = '" & asid & "' "
    strSql = strSql & " order by IsNull(T.OutMallCurrState, 'B001') "
    ''rw strSql

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
        tmpasid = rsget("tmpasid")
	end if
	rsget.Close

    if tmpasid = "" then
        Call fnUnmatchDeletedCS(OutMallOrderSerial)
        affectedRows = fnMatchCs(OutMallOrderSerial)
    end if


    if tmpasid = "" then
        rw "에러 : 매칭실패"
        dbget.close() : response.end
    else
	    rsget.CursorLocation = adUseClient
	    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	    if (Not rsget.Eof) then
            tmpasid = rsget("tmpasid")
            OutMallCurrState = rsget("OutMallCurrState")

            itemno = rsget("itemno")
            ishppNo = rsget("shppNo")
            ishppSeq = rsget("shppSeq")

            if rsget("gubun01") = "C004" and rsget("gubun02") = "CD01" then
                '// 고객변심
                resellPsblYn = "Y"
                retImptMainCd = "10"
            elseif rsget("gubun01") = "C005" and rsget("gubun02") = "CE01" then
                '// 상품불량
                resellPsblYn = "N"
                retImptMainCd = "20"
            else
                resellPsblYn = "Y"
                retImptMainCd = "20"
            end if
	    end if
	    rsget.Close
    end if

    if OutMallCurrState >= "B007" then
        rw "OK"
        dbget.close() : response.end
    end if

    if OutMallCurrState = "B001" then
        resultCode = sendSsgReturnConfirm(ishppNo, ishppSeq, itemno, resellPsblYn, retImptMainCd, iresultDesc)

        if (resultCode = "00") then
            OutMallCurrState = "B006"

		    strSql = " update "
		    strSql = strSql & " db_temp.dbo.tbl_xSite_TMPCS "
		    strSql = strSql & " set OutMallCurrState = 'B006' "
	        strSql = strSql & " where "
	        strSql = strSql & " 	1 = 1 "
	        strSql = strSql & " 	and SellSite = 'ssg' "
			strSql = strSql & "		and orderserial = '" & orderserial & "' "
			strSql = strSql & "		and divcd = '" & divcd & "' "
			strSql = strSql & "		and asid = '" & asid & "' "
			strSql = strSql & "		and shppNo = '" & ishppNo & "' "
			strSql = strSql & "		and shppSeq = '" & ishppSeq & "' "
		    dbget.Execute strSql
        else
            rw "ERR : " & resultCode
		    rw "ERR MSG : " & iresultDesc
            dbget.close() : response.end
        end if
    end if

    if OutMallCurrState = "B006" then
        resultCode = sendSsgReturnFinish(ishppNo, ishppSeq, itemno, resellPsblYn, retImptMainCd, iresultDesc)

        if (resultCode = "00") then
            OutMallCurrState = "B007"

		    strSql = " update "
		    strSql = strSql & " db_temp.dbo.tbl_xSite_TMPCS "
		    strSql = strSql & " set OutMallCurrState = 'B007' "
	        strSql = strSql & " where "
	        strSql = strSql & " 	1 = 1 "
	        strSql = strSql & " 	and SellSite = 'ssg' "
			strSql = strSql & "		and orderserial = '" & orderserial & "' "
			strSql = strSql & "		and divcd = '" & divcd & "' "
			strSql = strSql & "		and asid = '" & asid & "' "
			strSql = strSql & "		and shppNo = '" & ishppNo & "' "
			strSql = strSql & "		and shppSeq = '" & ishppSeq & "' "
		    dbget.Execute strSql
        else
            rw "ERR : " & resultCode
		    rw "ERR MSG : " & iresultDesc
            dbget.close() : response.end
        end if
    end if

    rw "OK"
    dbget.close() : response.end

else

	''istyyyymmdd = "20180211"
	''iedyyyymmdd = "20180215"
	'// 5일 전부터 오늘까지
	'call getSsgCancelList(istyyyymmdd,iedyyyymmdd)
	Call getNewSsgCancelList(istyyyymmdd,iedyyyymmdd)
	call getSsgExchangeList(istyyyymmdd,iedyyyymmdd)
	call getSsgExchangeChulgoList(istyyyymmdd,iedyyyymmdd)
    call getSsgExchangeChulgoList(replace(dateadd("d",1,Now()),"-",""),replace(dateadd("d",6,Now()),"-",""))		'// 출고예정일은 내일이후가 될수 있다.

	'' 조회일의 기준이 출고 요청일 인듯함. 함더쿼리 2018/02/20
	'// 10일 전부터 6일전까지
	iedyyyymmdd = replace(dateadd("d",-6,LEFT(iedyyyymmdd,4)&"-"&Mid(iedyyyymmdd,5,2)&"-"&Mid(iedyyyymmdd,7,2)),"-","")
	istyyyymmdd = replace(dateadd("d",-6,LEFT(iedyyyymmdd,4)&"-"&Mid(iedyyyymmdd,5,2)&"-"&Mid(iedyyyymmdd,7,2)),"-","")
	'call getSsgCancelList(istyyyymmdd,iedyyyymmdd)
	Call getNewSsgCancelList(istyyyymmdd,iedyyyymmdd)

	response.flush

	'' 조회일의 기준이 출고 요청일 인듯함. 함더쿼리 2018/02/20
	'// 15일 전부터 11일까지
	iedyyyymmdd = replace(dateadd("d",-6,LEFT(iedyyyymmdd,4)&"-"&Mid(iedyyyymmdd,5,2)&"-"&Mid(iedyyyymmdd,7,2)),"-","")
	istyyyymmdd = replace(dateadd("d",-6,LEFT(iedyyyymmdd,4)&"-"&Mid(iedyyyymmdd,5,2)&"-"&Mid(iedyyyymmdd,7,2)),"-","")
	'call getSsgCancelList(istyyyymmdd,iedyyyymmdd)
	Call getNewSsgCancelList(istyyyymmdd,iedyyyymmdd)

	response.flush

	'' 조회일의 기준이 출고 요청일 인듯함. 함더쿼리 2018/02/20
	'// 20일 전부터 16일까지
	iedyyyymmdd = replace(dateadd("d",-6,LEFT(iedyyyymmdd,4)&"-"&Mid(iedyyyymmdd,5,2)&"-"&Mid(iedyyyymmdd,7,2)),"-","")
	istyyyymmdd = replace(dateadd("d",-6,LEFT(iedyyyymmdd,4)&"-"&Mid(iedyyyymmdd,5,2)&"-"&Mid(iedyyyymmdd,7,2)),"-","")
	'call getSsgCancelList(istyyyymmdd,iedyyyymmdd)
	Call getNewSsgCancelList(istyyyymmdd,iedyyyymmdd)

	response.flush

    ''Call chkSsgExchangeNotFinished()

end if


''/ 반품/ 교환회수 목록 조회
public function getSsgExchangeList(styyyymmdd,edyyyymmdd)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim shppNo, shppSeq, ordNo , ordItemSeq , orordNo , orordItemSeq
    Dim shppStatCd , shppStatNm, itemId , itemNm , shppcstCodYn , ordpeNm , rcptpeNm , rcptpeHpno , rcptpeTelno
    Dim shppDivDtlCd , shppDivDtlNm , ordQty ,dircItemQty , procItemQty , shppMainCd
    Dim iDivCD

    Dim rcovDircDt , ordStatNm , ordItemStatNm
    Dim ordItemStatCd , shppMainNm , ordRcpDts

    Dim oMaster, oDetailArr(0)
    Dim ttlCNT : ttlCNT=0
    Dim iDontCareCnt : iDontCareCnt=0
    Dim iCareCnt : iCareCnt=0
    Dim iInputCnt : iInputCnt=0
    Dim iAssignedRow : iAssignedRow=0

    dim rcovMthdNm, lastShppProgStatDtlNm, retProcStatNm, delicoVenNm, wblNo, OutMallCurrState

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listExchangeTarget.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestExchangeTarget>"
	requestBody = requestBoDy&"<perdType>01</perdType>"  '' 회수지시일
    requestBody = requestBoDy&"<perdStrDts>"&styyyymmdd&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&edyyyymmdd&"</perdEndDts>"
    requestBody = requestBoDy&"</requestExchangeTarget>"

	objXML.send(requestBody)
	''rw objXML.status
	response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text

			Set LagrgeNode = xmlDOM.SelectNodes("/result/exchangeTargets/exchangeTarget")
			If Not (LagrgeNode Is Nothing) Then
			    For i = 0 To LagrgeNode.length - 1
			        shppNo ="": shppSeq = ""
			        ordNo ="": ordItemSeq =""
			        orordNo ="": orordItemSeq =""
			        shppStatCd =""          '' 10정상, [20취소], 30 대기
					shppStatNm=""
			        itemId ="" : itemNm =""
			        shppcstCodYn ="":       ''착불여부(Y?)
			        ordpeNm ="" : rcptpeNm ="" : rcptpeHpno ="" : rcptpeTelno=""  ''주문자 수령인, 수령인HP, 수령인전화
			        shppDivDtlCd =""        ''11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고
			        shppDivDtlNm =""
			        ordQty ="" : dircItemQty ="" : procItemQty =""  ''주문수량, 지시수량, 처리수량
			        shppMainCd ="" ''배송주체 (32 업체창고 [41] 협력업체 42 브랜드직배)

			        rcovDircDt =""  ''반품접수일
			        ordStatNm ="" '' [취소]
			        ordItemStatNm ="" '' 피킹완료 / 배송지시
			        ordItemStatCd ="" ''140(피킹완료) /130(배송지시) '' 배송지시 상태는 필요 없음.
			        shppMainNm =""  ''[협력업체]
			        ordRcpDts  =""  ''주문접수일시

                    rcovMthdNm = ""
                    lastShppProgStatDtlNm = ""
                    retProcStatNm = ""
                    delicoVenNm = ""
                    wblNo = ""
                    OutMallCurrState = ""

			        shppNo              = LagrgeNode(i).SelectSingleNode("shppNo").Text                 ''*배송번호
                    shppSeq             = LagrgeNode(i).SelectSingleNode("shppSeq").Text                ''*배송순번
                    ordNo               = LagrgeNode(i).SelectSingleNode("ordNo").Text                  ''*주문번호 [20171123128379]
                    ordItemSeq          = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text             ''*주문순번

                    if NOT (LagrgeNode(i).SelectSingleNode("orordNo") is Nothing) then
                        orordNo             = LagrgeNode(i).SelectSingleNode("orordNo").Text            ''원주문번호 [20171123128379]  ''대소문자 유의
                        ordNo = orordNo ''2017/12/27 추가
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("orordItemSeq") is Nothing) then
                        orordItemSeq        = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''원주문순번 [2]
                        ordItemSeq = orordItemSeq ''2017/12/27 추가
                    end if

                    shppStatCd          = LagrgeNode(i).SelectSingleNode("shppStatCd").Text         ''*배송상태코드 10 정상 30 대기 20취소
					shppStatNm          = LagrgeNode(i).SelectSingleNode("shppStatNm").Text

                    itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*상품번호 [1000024811163]
                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*상품명

                    ordpeNm             = LEFT(LagrgeNode(i).SelectSingleNode("ordpeNm").Text, 15)            ''*주문자
                    rcptpeNm            = LEFT(LagrgeNode(i).SelectSingleNode("rcptpeNm").Text, 15)           ''*수취인
                    rcptpeHpno          = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text         ''*수취인 휴대폰번호
                    if NOT (LagrgeNode(i).SelectSingleNode("rcptpeTelno") is Nothing) then
                        rcptpeTelno         = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text        ''*수취인 집전화번호
                    end if
                    shppDivDtlCd        = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text           ''*배송구분상세코드 11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고 21 반품회수
                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''배송구분상세명
                    dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''지시수량 [2]
                    procItemQty         = LagrgeNode(i).SelectSingleNode("procItemQty").Text        ''처리수량 [0]

                    if NOT (LagrgeNode(i).SelectSingleNode("rcovDircDt") is Nothing) then
                        rcovDircDt          = Left(LagrgeNode(i).SelectSingleNode("rcovDircDt").Text,19)  ''없는경우있음 20180218490344
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlNm") is Nothing) then
                        lastShppProgStatDtlNm = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlNm").Text
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("retProcStatNm") is Nothing) then
                        retProcStatNm = LagrgeNode(i).SelectSingleNode("retProcStatNm").Text
                    end if

                    ttlCNT = ttlCNT+1

                    if (lastShppProgStatDtlNm = "회수지시") then
                        OutMallCurrState = "B001"
                    elseif (lastShppProgStatDtlNm = "회수철회") then
                        OutMallCurrState = "B008"
                    elseif (lastShppProgStatDtlNm = "회수확인") then
                        OutMallCurrState = "B006"
                    elseif (retProcStatNm = "반품완료") then
                        OutMallCurrState = "B007"
                    else
                        OutMallCurrState = "BXXX"
                    end if

					if (shppStatCd="10") or (shppStatCd="20") then
						if (shppDivDtlCd = "21") then
							iDivCD = "A004"
						elseif (shppDivDtlCd = "22") then
							'// 교환회수
							iDivCD = "A011"
						else
							iDivCD = shppDivDtlCd
						end if

                        strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '"&CMALLNAME&"' and OutMallOrderSerial = '" & CStr(ordNo) & "' and OrgDetailKey = '" & CStr(ordItemSeq) & "' and divcd = '" & CStr(iDivCD) & "' and shppNo = '" & CStr(shppNo) & "' and shppSeq = '" & CStr(shppSeq) & "' ) "
    					strSql = strSql & " BEGIN "
    					strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
    					strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
    					strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno, shppNo, shppSeq, OutMallCurrState) VALUES "
    					strSql = strSql & " 	('"&iDivCD&"', '" & shppDivDtlNm & "', '"&CMALLNAME&"', '" & CStr(ordNo) & "', '"& ordpeNm &"', '', '"& "" &"', '"& "" &"', '"& rcptpeNm &"', "
    					strSql = strSql & "		'"&rcptpeTelno&"', '"&rcptpeHpno&"', '', '', '', '' "
    					strSql = strSql & "		, '" & html2db(CStr(rcovDircDt)) & "', '" & CStr(ordItemSeq) & "', '" & CStr(iDivCD) & "', '"&dircItemQty&"', '"&shppNo&"', '"&shppSeq&"', '"&OutMallCurrState&"') "
    					strSql = strSql & " END "
                        strSql = strSql & " ELSE "
    					strSql = strSql & " BEGIN "
    					strSql = strSql & " 	UPDATE db_temp.dbo.tbl_xSite_TMPCS SET OutMallCurrState = '"&OutMallCurrState&"' WHERE SellSite = '"&CMALLNAME&"' and OutMallOrderSerial = '" & CStr(ordNo) & "' and OrgDetailKey = '" & CStr(ordItemSeq) & "' and divcd = '" & CStr(iDivCD) & "' and shppNo = '" & CStr(shppNo) & "' and shppSeq = '" & CStr(shppSeq) & "' "
    					strSql = strSql & " END "
                        ''rw strSql
    					dbget.Execute strSql,iAssignedRow

    					if (iAssignedRow>0) then
    					    iInputCnt = iInputCnt+iAssignedRow
    				    end if

    					iCareCnt = iCareCnt+1
					else
					    '' New CASE
					    ''Dim TTT : TTT=1/0  '' raseErr
					    rw "shppStatCd:"&shppStatCd
					    rw "shppStatNm:"&shppStatNm
						response.end
					end if
			    Next
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing

	rw "======================================"
	rw "총 회수 건수:"&ttlCNT
end function

'// 반품/ 교환회수 미완료건 상태조회
public function chkSsgExchangeNotFinished()
    dim sqlStr, i
    dim shppNoArr, shppNo, OutMallOrderSerial, OutMallOrderSerialArr, divcd, divcdArr

    sqlStr = " select OutMallOrderSerial, IsNull(shppNo, '') as shppNo, divcd "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	db_temp.dbo.tbl_xSite_TMPCS "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + "     1 = 1 "
    sqlStr = sqlStr + "     and SellSite = 'ssg' "
    sqlStr = sqlStr + " 	and divcd in ('A011', 'A004') "
    sqlStr = sqlStr + " 	and currstate <> 'B007' "
    sqlStr = sqlStr + " 	and shppNo is not NULL "
    sqlStr = sqlStr + " order by idx desc "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

    shppNoArr = ""
    OutMallOrderSerialArr = ""
	if Not rsget.Eof then
		do until rsget.eof
		    shppNoArr = shppNoArr & rsget("shppNo") & ","
            OutMallOrderSerialArr = OutMallOrderSerialArr & rsget("OutMallOrderSerial") & ","
            divcdArr = divcdArr & rsget("divcd") & ","
		    rsget.MoveNext
    	loop
	end if
	rsget.close

    if (shppNoArr = "") then
        exit function
    end if

    shppNoArr = Split(shppNoArr)
    OutMallOrderSerialArr = Split(OutMallOrderSerialArr)
    divcdArr = Split(divcdArr)
    for i = 0 to UBound(shppNoArr)
        shppNo = Trim(shppNoArr(i))
        OutMallOrderSerial = Trim(OutMallOrderSerialArr(i))
        divcd = Trim(divcdArr(i))
        if shppNo <> "" then
            Call getSsgExchangeReturnOne(divcd, OutMallOrderSerial, shppNo)
        end if
    next
end function

''/ 반품/ 교환회수 한건 조회
public function getSsgExchangeReturnOne(divcd, ordNo, shppNo)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim shppSeq, ordItemSeq , orordNo , orordItemSeq
    Dim shppStatCd , shppStatNm, itemId , itemNm , shppcstCodYn , ordpeNm , rcptpeNm , rcptpeHpno , rcptpeTelno
    Dim shppDivDtlCd , shppDivDtlNm , ordQty ,dircItemQty , procItemQty , shppMainCd
    Dim iDivCD

    Dim rcovDircDt , ordStatNm , ordItemStatNm
    Dim ordItemStatCd , shppMainNm , ordRcpDts

    Dim oMaster, oDetailArr(0)
    Dim ttlCNT : ttlCNT=0
    Dim iDontCareCnt : iDontCareCnt=0
    Dim iCareCnt : iCareCnt=0
    Dim iInputCnt : iInputCnt=0
    Dim iAssignedRow : iAssignedRow=0

    dim rcovMthdNm, lastShppProgStatDtlNm, retProcStatNm, delicoVenNm, wblNo, OutMallCurrState

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listExchangeTarget.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestExchangeTarget>"
    if (shppNo <> "") then
	    requestBody = requestBoDy&"<commType>03</commType>"  '' 배송번호
        requestBody = requestBoDy&"<commValue>"&shppNo&"</commValue>"
    else
	    requestBody = requestBoDy&"<commType>02</commType>"  '' 주문번호
        requestBody = requestBoDy&"<commValue>"&ordNo&"</commValue>"
    end if
    requestBody = requestBoDy&"</requestExchangeTarget>"

	objXML.send(requestBody)

    Response.CharSet="euc-kr"
    Session.codepage="949"
    Response.codepage="949"
    Response.ContentType="text/html;charset=euc-kr"

    getSsgExchangeReturnOne = 0

	''rw objXML.status
	''response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text

			Set LagrgeNode = xmlDOM.SelectNodes("/result/exchangeTargets/exchangeTarget")
			If Not (LagrgeNode Is Nothing) Then
			    For i = 0 To LagrgeNode.length - 1
			        shppNo ="": shppSeq = ""
			        ordNo ="": ordItemSeq =""
			        orordNo ="": orordItemSeq =""
			        shppStatCd =""          '' 10정상, [20취소], 30 대기
					shppStatNm=""
			        itemId ="" : itemNm =""
			        shppcstCodYn ="":       ''착불여부(Y?)
			        ordpeNm ="" : rcptpeNm ="" : rcptpeHpno ="" : rcptpeTelno=""  ''주문자 수령인, 수령인HP, 수령인전화
			        shppDivDtlCd =""        ''11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고
			        shppDivDtlNm =""
			        ordQty ="" : dircItemQty ="" : procItemQty =""  ''주문수량, 지시수량, 처리수량
			        shppMainCd ="" ''배송주체 (32 업체창고 [41] 협력업체 42 브랜드직배)

			        rcovDircDt =""  ''반품접수일
			        ordStatNm ="" '' [취소]
			        ordItemStatNm ="" '' 피킹완료 / 배송지시
			        ordItemStatCd ="" ''140(피킹완료) /130(배송지시) '' 배송지시 상태는 필요 없음.
			        shppMainNm =""  ''[협력업체]
			        ordRcpDts  =""  ''주문접수일시

                    rcovMthdNm = ""
                    lastShppProgStatDtlNm = ""
                    retProcStatNm = ""
                    delicoVenNm = ""
                    wblNo = ""
                    OutMallCurrState = ""

			        shppNo              = LagrgeNode(i).SelectSingleNode("shppNo").Text                 ''*배송번호
                    shppSeq             = LagrgeNode(i).SelectSingleNode("shppSeq").Text                ''*배송순번
                    ordNo               = LagrgeNode(i).SelectSingleNode("ordNo").Text                  ''*주문번호 [20171123128379]
                    ordItemSeq          = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text             ''*주문순번

                    if NOT (LagrgeNode(i).SelectSingleNode("orOrdNo") is Nothing) then
                        orordNo             = LagrgeNode(i).SelectSingleNode("orOrdNo").Text            ''원주문번호 [20171123128379]  ''대소문자 유의
                        ordNo = orordNo ''2017/12/27 추가
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("orordItemSeq") is Nothing) then
                        orordItemSeq        = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''원주문순번 [2]
                        ordItemSeq = orordItemSeq ''2017/12/27 추가
                    end if

                    shppStatCd          = LagrgeNode(i).SelectSingleNode("shppStatCd").Text         ''*배송상태코드 10 정상 30 대기 20취소
					shppStatNm          = LagrgeNode(i).SelectSingleNode("shppStatNm").Text

                    itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*상품번호 [1000024811163]
                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*상품명

                    ordpeNm             = LEFT(LagrgeNode(i).SelectSingleNode("ordpeNm").Text, 15)            ''*주문자
                    rcptpeNm            = LEFT(LagrgeNode(i).SelectSingleNode("rcptpeNm").Text, 15)           ''*수취인
                    rcptpeHpno          = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text         ''*수취인 휴대폰번호
                    if NOT (LagrgeNode(i).SelectSingleNode("rcptpeTelno") is Nothing) then
                        rcptpeTelno         = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text        ''*수취인 집전화번호
                    end if
                    shppDivDtlCd        = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text           ''*배송구분상세코드 11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고 21 반품회수
                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''배송구분상세명
                    dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''지시수량 [2]
                    procItemQty         = LagrgeNode(i).SelectSingleNode("procItemQty").Text        ''처리수량 [0]

                    if NOT (LagrgeNode(i).SelectSingleNode("rcovMthdNm") is Nothing) then
                        rcovMthdNm = LagrgeNode(i).SelectSingleNode("rcovMthdNm").Text
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlNm") is Nothing) then
                        lastShppProgStatDtlNm = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlNm").Text
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("retProcStatNm") is Nothing) then
                        retProcStatNm = LagrgeNode(i).SelectSingleNode("retProcStatNm").Text
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("delicoVenNm") is Nothing) then
                        delicoVenNm = LagrgeNode(i).SelectSingleNode("delicoVenNm").Text
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("wblNo") is Nothing) then
                        wblNo = LagrgeNode(i).SelectSingleNode("wblNo").Text
                    end if

                    if (lastShppProgStatDtlNm = "회수지시") then
                        OutMallCurrState = "B001"
                    elseif (lastShppProgStatDtlNm = "회수철회") then
                        OutMallCurrState = "B008"
                    elseif (lastShppProgStatDtlNm = "회수확인") then
                        OutMallCurrState = "B006"
                    elseif (retProcStatNm = "반품완료") then
                        OutMallCurrState = "B007"
                    else
                        OutMallCurrState = "BXXX"
                    end if

					rw "shppNo:"&shppNo
                    rw "rcovMthdNm:"&rcovMthdNm
                    rw "lastShppProgStatDtlNm:"&lastShppProgStatDtlNm
                    rw "retProcStatNm:"&retProcStatNm
                    rw "delicoVenNm:"&delicoVenNm
                    rw "wblNo:"&wblNo

                    strSql = " update db_temp.dbo.tbl_xSite_TMPCS "
                    strSql = strSql + " set OutMallCurrState = '" & OutMallCurrState & "' "
                    strSql = strSql + " where "
                    strSql = strSql + " 	1 = 1 "
                    strSql = strSql + " 	and SellSite = 'ssg' "
                    strSql = strSql + " 	and OutMallOrderSerial = '" & ordNo & "' "
                    strSql = strSql + " 	and OrgDetailKey = '" & ordItemSeq & "' "
                    strSql = strSql + " 	and shppNo = '" & shppNo & "' "
                    strSql = strSql + " 	and shppSeq = '" & shppSeq & "' "
                    dbget.Execute strSql,iAssignedRow

                    if iAssignedRow = 0 then
                        strSql = " update db_temp.dbo.tbl_xSite_TMPCS "
                        strSql = strSql + " set OutMallCurrState = '" & OutMallCurrState & "', shppNo = '" & shppNo & "', shppSeq = '" & shppSeq & "' "
                        strSql = strSql + " where "
                        strSql = strSql + " 	1 = 1 "
                        strSql = strSql + " 	and SellSite = 'ssg' "
                        strSql = strSql + " 	and OutMallOrderSerial = '" & ordNo & "' "
                        strSql = strSql + " 	and OrgDetailKey = '" & ordItemSeq & "' "
                        strSql = strSql + " 	and divcd = '" & divcd & "' "
                        strSql = strSql + " 	and shppNo is NULL "
                        ''strSql = strSql + " 	and shppSeq = '" & shppSeq & "' "
                        dbget.Execute strSql,iAssignedRow

                        if iAssignedRow > 1 then
                            rw ordNo
                            rw ordItemSeq
                            dbget.close() : response.end
                        end if

                        getSsgExchangeReturnOne = iAssignedRow
                    end if

                    ttlCNT = ttlCNT+1
			    Next
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing

	rw "======================================"
end function

'// 일반출고, 교환출고
public function getSsgExchangeChulgoList(styyyymmdd,edyyyymmdd)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim shppNo, shppSeq, ordNo , ordItemSeq , orordNo , orordItemSeq
    Dim shppStatCd , shppStatNm, itemId , itemNm , shppcstCodYn , ordpeNm , rcptpeNm , rcptpeHpno , rcptpeTelno
    Dim shppDivDtlCd , shppDivDtlNm , ordQty ,dircItemQty , procItemQty , shppMainCd
    Dim iDivCD

    Dim rcovDircDt , ordStatNm , ordItemStatNm
    Dim ordItemStatCd , shppMainNm , ordRcpDts

    Dim oMaster, oDetailArr(0)
    Dim ttlCNT : ttlCNT=0
    Dim iDontCareCnt : iDontCareCnt=0
    Dim iCareCnt : iCareCnt=0
    Dim iInputCnt : iInputCnt=0
    Dim iAssignedRow : iAssignedRow=0

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listShppDirection.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestShppDirection>"
	requestBody = requestBoDy&"<perdType>01</perdType>"  '' 배송지시일
    requestBody = requestBoDy&"<perdStrDts>"&styyyymmdd&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&edyyyymmdd&"</perdEndDts>"		'// 기간은 7일이내
	requestBody = requestBoDy&"<shppDivDtlCd>15</shppDivDtlCd>"		'// 11 일반출고, 12 부분출고, 14 재배송, 15 교환출고, 16 AS출고
    requestBody = requestBoDy&"</requestShppDirection>"

	objXML.send(requestBody)
	''rw objXML.status
	''response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text

			Set LagrgeNode = xmlDOM.SelectNodes("/result/shppDirections/shppDirection")
			If Not (LagrgeNode Is Nothing) Then
			    For i = 0 To LagrgeNode.length - 1
			        shppNo ="": shppSeq = ""
			        ordNo ="": ordItemSeq =""
			        orordNo ="": orordItemSeq =""
			        shppStatCd =""          '' 10정상, [20취소], 30 대기
					shppStatNm=""
			        itemId ="" : itemNm =""
			        shppcstCodYn ="":       ''착불여부(Y?)
			        ordpeNm ="" : rcptpeNm ="" : rcptpeHpno ="" : rcptpeTelno=""  ''주문자 수령인, 수령인HP, 수령인전화
			        shppDivDtlCd =""        ''11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고
			        shppDivDtlNm =""
			        ordQty ="" : dircItemQty ="" : procItemQty =""  ''주문수량, 지시수량, 처리수량
			        shppMainCd ="" ''배송주체 (32 업체창고 [41] 협력업체 42 브랜드직배)

			        rcovDircDt =""  ''반품접수일
			        ordStatNm ="" '' [취소]
			        ordItemStatNm ="" '' 피킹완료 / 배송지시
			        ordItemStatCd ="" ''140(피킹완료) /130(배송지시) '' 배송지시 상태는 필요 없음.
			        shppMainNm =""  ''[협력업체]
			        ordRcpDts  =""  ''주문접수일시

			        shppNo              = LagrgeNode(i).SelectSingleNode("shppNo").Text                 ''*배송번호
                    shppSeq             = LagrgeNode(i).SelectSingleNode("shppSeq").Text                ''*배송순번
                    ordNo               = LagrgeNode(i).SelectSingleNode("ordNo").Text                  ''*주문번호 [20171123128379]
                    ordItemSeq          = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text             ''*주문순번

                    if NOT (LagrgeNode(i).SelectSingleNode("orOrdNo") is Nothing) then
                        orordNo             = LagrgeNode(i).SelectSingleNode("orOrdNo").Text            ''원주문번호 [20171123128379]  ''대소문자 유의
                        ordNo = orordNo ''2017/12/27 추가
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("orordItemSeq") is Nothing) then
                        orordItemSeq        = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''원주문순번 [2]
                        ordItemSeq = orordItemSeq ''2017/12/27 추가
                    end if

                    shppStatCd          = LagrgeNode(i).SelectSingleNode("shppStatCd").Text         ''*배송상태코드 10 정상 30 대기 20취소
					shppStatNm          = LagrgeNode(i).SelectSingleNode("shppStatNm").Text

                    itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*상품번호 [1000024811163]
                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*상품명

                    ordpeNm             = LEFT(LagrgeNode(i).SelectSingleNode("ordpeNm").Text, 15)            ''*주문자
                    rcptpeNm            = LEFT(LagrgeNode(i).SelectSingleNode("rcptpeNm").Text, 15)           ''*수취인
                    rcptpeHpno          = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text         ''*수취인 휴대폰번호
                    if NOT (LagrgeNode(i).SelectSingleNode("rcptpeTelno") is Nothing) then
                        rcptpeTelno         = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text        ''*수취인 집전화번호
                    end if
                    shppDivDtlCd        = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text           ''*배송구분상세코드 11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고 21 반품회수
                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''배송구분상세명
                    dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''지시수량 [2]
                    ''procItemQty         = LagrgeNode(i).SelectSingleNode("procItemQty").Text        ''처리수량 [0]

                    if NOT (LagrgeNode(i).SelectSingleNode("rcovDircDt") is Nothing) then
                        rcovDircDt          = Left(LagrgeNode(i).SelectSingleNode("rcovDircDt").Text,19)  ''없는경우있음 20180218490344
                    end if

                    ttlCNT = ttlCNT+1

                    rw ordNo
                    rw shppDivDtlCd

					if (shppStatCd="10") then
						if (shppDivDtlCd = "15") then
							iDivCD = "A000"
						else
							iDivCD = shppDivDtlCd
						end if

                        strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '"&CMALLNAME&"' and OutMallOrderSerial = '" & CStr(ordNo) & "' and OrgDetailKey = '" & CStr(shppNo) & "_" & CStr(shppSeq) & "' and divcd = '" & CStr(iDivCD) & "' ) "
    					strSql = strSql & " BEGIN "
    					strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
    					strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
    					strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno, OutMallCurrState) VALUES "
    					strSql = strSql & " 	('"&iDivCD&"', '" & shppDivDtlNm & "', '"&CMALLNAME&"', '" & CStr(ordNo) & "', '"& ordpeNm &"', '', '"& "" &"', '"& "" &"', '"& rcptpeNm &"', "
    					strSql = strSql & "		'"&rcptpeTelno&"', '"&rcptpeHpno&"', '', '', '', '' "
    					strSql = strSql & "		, '" & html2db(CStr(rcovDircDt)) & "', '" & CStr(ordItemSeq) & "', '" & CStr(shppNo) & "_" & CStr(shppSeq) & "', '"&dircItemQty&"', 'B001') "
    					strSql = strSql & " END "
    					dbget.Execute strSql,iAssignedRow

    					if (iAssignedRow>0) then
    					    iInputCnt = iInputCnt+iAssignedRow

							'// 주문확인
							Call getSsgExchangeChulgoConfirm(shppNo, shppSeq)
    				    end if

    					iCareCnt = iCareCnt+1
					elseif (shppStatCd="30") then
						'// 대기
						'// skip
					else
					    '' New CASE
					    ''Dim TTT : TTT=1/0  '' raseErr
					    rw "shppStatCd:"&shppStatCd
					    rw "shppStatNm:"&shppStatNm
						response.end
					end if
			    Next

			    '' CS 마스터정보. 업데이트?  위치변경 2018/07/20 eastone
				strSql = " update c "
				strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
				strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
				strSql = strSql + " , c.OrderName = o.OrderName "
				strSql = strSql + " from "
				strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
				strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
				strSql = strSql + " on "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.SellSite = o.SellSite "
				strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
				strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
				strSql = strSql + " where "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.orderserial is NULL "
				strSql = strSql + " 	and o.orderserial is not NULL "
				strSql = strSql + " 	and c.sellsite = '"&CMALLNAME&"' "
				dbget.Execute strSql
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing

	rw "======================================"
	rw "총 CS출고 건수:"&iCareCnt
end function

'// 주문확인
public function getSsgExchangeChulgoConfirm(shppNo, shppSeq)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim ordNo , ordItemSeq , orordNo , orordItemSeq
    Dim shppStatCd , shppStatNm, itemId , itemNm , shppcstCodYn , ordpeNm , rcptpeNm , rcptpeHpno , rcptpeTelno
    Dim shppDivDtlCd , shppDivDtlNm , ordQty ,dircItemQty , procItemQty , shppMainCd
    Dim iDivCD

    Dim rcovDircDt , ordStatNm , ordItemStatNm
    Dim ordItemStatCd , shppMainNm , ordRcpDts

    Dim oMaster, oDetailArr(0)
    Dim ttlCNT : ttlCNT=0
    Dim iDontCareCnt : iDontCareCnt=0
    Dim iCareCnt : iCareCnt=0
    Dim iInputCnt : iInputCnt=0
    Dim iAssignedRow : iAssignedRow=0

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/updateOrderSubjectManage.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestOrderSubjectManage>"
    requestBody = requestBoDy&"<shppNo>"&shppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&shppSeq&"</shppSeq>"
    requestBody = requestBoDy&"</requestOrderSubjectManage>"

	objXML.send(requestBody)
	''rw objXML.status
	''response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text

	Set objXML = nothing

	rw "======================================"
	rw CStr(shppNo) & "_" & CStr(shppSeq) & " CS교환출고 확인처리:" & ssgresultCode & "<br />"
end function

'// 송장전송
'// 참조 : http://eapi.ssgadm.com/info/shpp/saveWblNo.ssg
''wblNo : 운송장번호, delicoVenId : 택배사, shppTypeCd :배송유형, shppTypeDtlCd : 배송유형디테일
''shppTypeCd
''10 자사배송
''20 택배배송
''30 매장방문
''40 등기
''50 미배송
''60 미발송
''shppTypeDtlCd
''14 업체자사배송
''22 업체택배배송
''25 해외택배배송
''31 매장방문
''41 등기
''51 SMS
''52 EMAIL
''61 미발송
public function sendSsgSendSongjang(shppNo, shppSeq, wblNo, delicoVenId, shppTypeCd, shppTypeDtlCd)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim ordNo , ordItemSeq , orordNo , orordItemSeq
    Dim shppStatCd , shppStatNm, itemId , itemNm , shppcstCodYn , ordpeNm , rcptpeNm , rcptpeHpno , rcptpeTelno
    Dim shppDivDtlCd , shppDivDtlNm , ordQty ,dircItemQty , procItemQty , shppMainCd
    Dim iDivCD

    Dim rcovDircDt , ordStatNm , ordItemStatNm
    Dim ordItemStatCd , shppMainNm , ordRcpDts

    Dim oMaster, oDetailArr(0)
    Dim ttlCNT : ttlCNT=0
    Dim iDontCareCnt : iDontCareCnt=0
    Dim iCareCnt : iCareCnt=0
    Dim iInputCnt : iInputCnt=0
    Dim iAssignedRow : iAssignedRow=0

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/saveWblNo.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestWhOutCompleteProcess>"
    requestBody = requestBoDy&"<shppNo>"&shppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&shppSeq&"</shppSeq>"
	if (shppTypeCd <> "10") then
		'// 자사배송 아닐때만
		requestBody = requestBoDy&"<wblNo>"&wblNo&"</wblNo>"
		requestBody = requestBoDy&"<delicoVenId>"&delicoVenId&"</delicoVenId>"
	end if
	requestBody = requestBoDy&"<shppTypeCd>"&shppTypeCd&"</shppTypeCd>"
	requestBody = requestBoDy&"<shppTypeDtlCd>"&shppTypeDtlCd&"</shppTypeDtlCd>"
    requestBody = requestBoDy&"</requestWhOutCompleteProcess>"

	objXML.send(requestBody)
	''rw objXML.status
	''response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text

	Set objXML = nothing

	sendSsgSendSongjang = ssgresultCode
end function

public function sendSsgChulgoFinish(shppNo, shppSeq, procItemQty, ByRef resultDesc)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim ordNo , ordItemSeq , orordNo , orordItemSeq
    Dim shppStatCd , shppStatNm, itemId , itemNm , shppcstCodYn , ordpeNm , rcptpeNm , rcptpeHpno , rcptpeTelno
    Dim shppDivDtlCd , shppDivDtlNm , ordQty ,dircItemQty , shppMainCd
    Dim iDivCD

    Dim rcovDircDt , ordStatNm , ordItemStatNm
    Dim ordItemStatCd , shppMainNm , ordRcpDts

    Dim oMaster, oDetailArr(0)
    Dim ttlCNT : ttlCNT=0
    Dim iDontCareCnt : iDontCareCnt=0
    Dim iCareCnt : iCareCnt=0
    Dim iInputCnt : iInputCnt=0
    Dim iAssignedRow : iAssignedRow=0

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/saveWhOutCompleteProcess.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestWhOutCompleteProcess>"
    requestBody = requestBoDy&"<shppNo>"&shppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&shppSeq&"</shppSeq>"
	requestBody = requestBoDy&"<procItemQty>"&procItemQty&"</procItemQty>"
    requestBody = requestBoDy&"</requestWhOutCompleteProcess>"

	objXML.send(requestBody)
	''rw objXML.status
	''response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			resultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text

	Set objXML = nothing

	sendSsgChulgoFinish = ssgresultCode
end function

public function sendSsgBeasongFinish(shppNo, shppSeq, ByRef resultDesc)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim ordNo , ordItemSeq , orordNo , orordItemSeq
    Dim shppStatCd , shppStatNm, itemId , itemNm , shppcstCodYn , ordpeNm , rcptpeNm , rcptpeHpno , rcptpeTelno
    Dim shppDivDtlCd , shppDivDtlNm , ordQty ,dircItemQty , procItemQty , shppMainCd
    Dim iDivCD

    Dim rcovDircDt , ordStatNm , ordItemStatNm
    Dim ordItemStatCd , shppMainNm , ordRcpDts

    Dim oMaster, oDetailArr(0)
    Dim ttlCNT : ttlCNT=0
    Dim iDontCareCnt : iDontCareCnt=0
    Dim iCareCnt : iCareCnt=0
    Dim iInputCnt : iInputCnt=0
    Dim iAssignedRow : iAssignedRow=0

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/saveDeliveryEnd.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestDeliveryEnd>"
    requestBody = requestBoDy&"<shppNo>"&shppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&shppSeq&"</shppSeq>"
    requestBody = requestBoDy&"</requestDeliveryEnd>"

	objXML.send(requestBody)
	''rw objXML.status
	''response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			resultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text

	Set objXML = nothing

	sendSsgBeasongFinish = ssgresultCode
end function

''취소
public function getSsgCancelList(styyyymmdd,edyyyymmdd)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim shppNo, shppSeq, ordNo , ordItemSeq , orordNo , orordItemSeq
    Dim shppStatCd , itemId , itemNm , shppcstCodYn , ordpeNm , rcptpeNm , rcptpeHpno , rcptpeTelno
    Dim shppDivDtlCd , shppDivDtlNm , ordQty ,dircItemQty , cnclItemQty , shppMainCd
    Dim iDivCD

    Dim ordCnclDts , ordStatNm , ordItemStatNm
    Dim ordItemStatCd , shppMainNm , ordRcpDts

    Dim oMaster, oDetailArr(0)
    Dim ttlCNT : ttlCNT=0
    Dim iDontCareCnt : iDontCareCnt=0
    Dim iCareCnt : iCareCnt=0
    Dim iInputCnt : iInputCnt=0
    Dim iAssignedRow : iAssignedRow=0

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listOrdCancel.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestShppDirection>"
    requestBody = requestBoDy&"<perdStrDts>"&styyyymmdd&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&edyyyymmdd&"</perdEndDts>"
    requestBody = requestBoDy&"</requestShppDirection>"

	objXML.send(requestBody)
'rw objXML.status
'response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"
'exit function
'response.end

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text

			Set LagrgeNode = xmlDOM.SelectNodes("/result/shppDirections/shppDirection")
			If Not (LagrgeNode Is Nothing) Then
			    For i = 0 To LagrgeNode.length - 1
			        ''변수초기화.
			        shppNo ="": shppSeq = ""
			        ordNo ="": ordItemSeq =""
			        orordNo ="": orordItemSeq =""
			        shppStatCd =""          '' 10정상, [20취소], 30 대기
			        itemId ="" : itemNm =""
			        shppcstCodYn ="":       ''착불여부(Y?)
			        ordpeNm ="" : rcptpeNm ="" : rcptpeHpno ="" : rcptpeTelno=""  ''주문자 수령인, 수령인HP, 수령인전화
			        shppDivDtlCd =""        ''11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고
			        shppDivDtlNm =""
			        ordQty ="" : dircItemQty ="" : cnclItemQty =""  ''주문수량, 지시수량, 취소수량
			        shppMainCd ="" ''배송주체 (32 업체창고 [41] 협력업체 42 브랜드직배)

			        ordCnclDts =""  ''취소일
			        ordStatNm ="" '' [취소]
			        ordItemStatNm ="" '' 피킹완료 / 배송지시
			        ordItemStatCd ="" ''140(피킹완료) /130(배송지시) '' 배송지시 상태는 필요 없음.
			        shppMainNm =""  ''[협력업체]
			        ordRcpDts  =""  ''주문접수일시


			        shppNo              = LagrgeNode(i).SelectSingleNode("shppNo").Text                 ''*배송번호
                    shppSeq             = LagrgeNode(i).SelectSingleNode("shppSeq").Text                ''*배송순번
                    ordNo               = LagrgeNode(i).SelectSingleNode("ordNo").Text                  ''*주문번호 [20171123128379]
                    ordItemSeq          = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text             ''*주문순번

                    if NOT (LagrgeNode(i).SelectSingleNode("orOrdNo") is Nothing) then
                        orordNo             = LagrgeNode(i).SelectSingleNode("orOrdNo").Text            ''원주문번호 [20171123128379]  ''대소문자 유의
                        ordNo = orordNo ''2017/12/27 추가
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("orordItemSeq") is Nothing) then
                        orordItemSeq        = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''원주문순번 [2]
                        ordItemSeq = orordItemSeq ''2017/12/27 추가
                    end if

                    shppStatCd          = LagrgeNode(i).SelectSingleNode("shppStatCd").Text         ''*배송상태코드 10 정상 30 대기 20취소
                    itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*상품번호 [1000024811163]
                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*상품명
                    shppcstCodYn        = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text   ''*배송비 착불여부 Y: 착불 N: 선불

                    ordpeNm             = LEFT(LagrgeNode(i).SelectSingleNode("ordpeNm").Text, 15)            ''*주문자
                    rcptpeNm            = LEFT(LagrgeNode(i).SelectSingleNode("rcptpeNm").Text, 15)           ''*수취인
                    rcptpeHpno          = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text         ''*수취인 휴대폰번호
                    if NOT (LagrgeNode(i).SelectSingleNode("rcptpeTelno") is Nothing) then
                        rcptpeTelno         = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text        ''*수취인 집전화번호
                    end if
                    shppDivDtlCd        = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text           ''*배송구분상세코드 11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고
                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''배송구분상세명
                    ordQty              = LagrgeNode(i).SelectSingleNode("ordQty").Text             ''주문수량 [2]
                    dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''지시수량 [2]
                    cnclItemQty         = LagrgeNode(i).SelectSingleNode("cnclItemQty").Text        ''취소수량 [0]
                    shppMainCd          = LagrgeNode(i).SelectSingleNode("shppMainCd").Text         ''배송주체코드 32 업체창고 41 협력업체 42 브랜드직배  [41]

                    if NOT (LagrgeNode(i).SelectSingleNode("ordCnclDts") is Nothing) then
                        ordCnclDts          = Left(LagrgeNode(i).SelectSingleNode("ordCnclDts").Text,19)  ''없는경우있음 20180218490344
                    end if
                    ordStatNm           = LagrgeNode(i).SelectSingleNode("ordStatNm").Text
                    ordItemStatNm       = LagrgeNode(i).SelectSingleNode("ordItemStatNm").Text
                    ordItemStatCd       = LagrgeNode(i).SelectSingleNode("ordItemStatCd").Text
                    shppMainNm          = LagrgeNode(i).SelectSingleNode("shppMainNm").Text
                    if NOT (LagrgeNode(i).SelectSingleNode("ordRcpDts") is Nothing) then
                        ordRcpDts           = Left(LagrgeNode(i).SelectSingleNode("ordRcpDts").Text,19)  ''원주문일시  ''없는경우있음 20180218490344
                    end if

                    ttlCNT = ttlCNT+1

                    if (shppStatCd="20") and ((ordItemStatCd="130") or (ordItemStatCd="110"))  then  ''배송지시상태중 취소 : idont care => 입력한다.
                        ''iDontCareCnt = iDontCareCnt+1
                        iDivCD = "A008"  ''취소

                        strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '"&CMALLNAME&"' and OutMallOrderSerial in ('" & CStr(ordNo) & "', '" & CStr(ordNo) & "_1', '" & CStr(ordNo) & "_2', '" & CStr(ordNo) & "_3') and OrgDetailKey = '" & CStr(ordItemSeq) & "' ) "
    					strSql = strSql & " BEGIN "

    					strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
    					strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
    					strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
    					strSql = strSql & " 	('"&iDivCD&"', '단순변심', '"&CMALLNAME&"', '" & CStr(ordNo) & "', '"& ordpeNm &"', '', '"& "" &"', '"& "" &"', '"& rcptpeNm &"', "
    					strSql = strSql & "		'"&rcptpeTelno&"', '"&rcptpeHpno&"', '', '', '', '' "
    					strSql = strSql & "		, '" & html2db(CStr(ordCnclDts)) & "', '" & CStr(ordItemSeq) & "', '', '"&cnclItemQty&"'); "

						strSql = strSql & "	update c "
    					strSql = strSql & "	set c.OutMallOrderSerial = o.OutMallOrderSerial "
    					strSql = strSql & "	from "
    					strSql = strSql & "		db_temp.dbo.tbl_xSite_TMPCS c "
    					strSql = strSql & "		join db_temp.dbo.tbl_xSite_TMPOrder o "
    					strSql = strSql & "		on "
    					strSql = strSql & "			1 = 1 "
    					strSql = strSql & "			and c.SellSite = o.SellSite "
    					strSql = strSql & "			and o.OutMallOrderSerial in (c.OutMallOrderSerial, c.OutMallOrderSerial + '_1', c.OutMallOrderSerial + '_2', c.OutMallOrderSerial + '_3') "
    					strSql = strSql & "			and c.orgdetailkey = o.orgdetailkey "
    					strSql = strSql & "	where "
    					strSql = strSql & "		1 = 1 "
    					strSql = strSql & "		and c.SellSite = '"&CMALLNAME&"' "
    					strSql = strSql & "		and c.OutMallOrderSerial = '" & CStr(ordNo) & "' "
    					strSql = strSql & "		and c.orgdetailkey = '" & CStr(ordItemSeq) & "' "
    					strSql = strSql & "		and c.OutMallOrderSerial <> o.OutMallOrderSerial "

						strSql = strSql & " END "
    					dbget.Execute strSql,iAssignedRow

    					if (iAssignedRow>0) then
    					    iInputCnt = iInputCnt+iAssignedRow
    				    end if

    					''주문 입력 이전 내역은 삭제 하자
    					strSql = " update c "
    					strSql = strSql + " set matchState='D'"
    					strSql = strSql + " from db_temp.dbo.tbl_xSite_TMPOrder c "
    					strSql = strSql + " WHERE SellSite = '"&CMALLNAME&"' and OutMallOrderSerial in ('" & CStr(ordNo) & "', '" & CStr(ordNo) & "_1', '" & CStr(ordNo) & "_2', '" & CStr(ordNo) & "_3') and OrgDetailKey = '" & CStr(ordItemSeq) & "'"
    					strSql = strSql + " and orderserial is NULL"
    					dbget.Execute strSql

    					iCareCnt = iCareCnt+1
                    elseif (shppStatCd="20") and (ordItemStatCd="140") then ''피킹완료상태중 취소.
                        iDivCD = "A008"  ''취소

                        strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '"&CMALLNAME&"' and OutMallOrderSerial in ('" & CStr(ordNo) & "', '" & CStr(ordNo) & "_1', '" & CStr(ordNo) & "_2', '" & CStr(ordNo) & "_3') and OrgDetailKey = '" & CStr(ordItemSeq) & "' ) "
    					strSql = strSql & " BEGIN "
    					strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
    					strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
    					strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
    					strSql = strSql & " 	('"&iDivCD&"', '단순변심', '"&CMALLNAME&"', '" & CStr(ordNo) & "', '"& ordpeNm &"', '', '"& "" &"', '"& "" &"', '"& rcptpeNm &"', "
    					strSql = strSql & "		'"&rcptpeTelno&"', '"&rcptpeHpno&"', '', '', '', '' "
    					strSql = strSql & "		, '" & html2db(CStr(ordCnclDts)) & "', '" & CStr(ordItemSeq) & "', '', '"&cnclItemQty&"'); "

						strSql = strSql & "	update c "
    					strSql = strSql & "	set c.OutMallOrderSerial = o.OutMallOrderSerial "
    					strSql = strSql & "	from "
    					strSql = strSql & "		db_temp.dbo.tbl_xSite_TMPCS c "
    					strSql = strSql & "		join db_temp.dbo.tbl_xSite_TMPOrder o "
    					strSql = strSql & "		on "
    					strSql = strSql & "			1 = 1 "
    					strSql = strSql & "			and c.SellSite = o.SellSite "
    					strSql = strSql & "			and o.OutMallOrderSerial in (c.OutMallOrderSerial, c.OutMallOrderSerial + '_1', c.OutMallOrderSerial + '_2', c.OutMallOrderSerial + '_3') "
    					strSql = strSql & "			and c.orgdetailkey = o.orgdetailkey "
    					strSql = strSql & "	where "
    					strSql = strSql & "		1 = 1 "
    					strSql = strSql & "		and c.SellSite = '"&CMALLNAME&"' "
    					strSql = strSql & "		and c.OutMallOrderSerial = '" & CStr(ordNo) & "' "
    					strSql = strSql & "		and c.orgdetailkey = '" & CStr(ordItemSeq) & "' "
    					strSql = strSql & "		and c.OutMallOrderSerial <> o.OutMallOrderSerial "

    					strSql = strSql & " END "
    					dbget.Execute strSql,iAssignedRow

    					if (iAssignedRow>0) then
    					    iInputCnt = iInputCnt+iAssignedRow

    					    '' ToDo : 실주문 입력 되었으나 취소되는 케이스
    					    '' 여기서 상품 준비중인 케이스가 있으면 먼가 노티나 취소처리를 해야한다. ex: 17120454071
    					    '' 텐배만 있는경우. (현재 SSG는 텐배만 있다.)
    				    end if

    					''주문 입력 이전 내역은 삭제 하자
    					strSql = " update c "
    					strSql = strSql + " set matchState='D'"
    					strSql = strSql + " from db_temp.dbo.tbl_xSite_TMPOrder c "
    					strSql = strSql + " WHERE SellSite = '"&CMALLNAME&"' and OutMallOrderSerial in ('" & CStr(ordNo) & "', '" & CStr(ordNo) & "_1', '" & CStr(ordNo) & "_2', '" & CStr(ordNo) & "_3') and OrgDetailKey = '" & CStr(ordItemSeq) & "'"
    					strSql = strSql + " and orderserial is NULL"
    					dbget.Execute strSql

    					iCareCnt = iCareCnt+1



    				elseif (shppStatCd="20") and (ordItemStatCd="170") then ''배송완료 상품 오배송.	CASE 20180226526399  //2018/02/26 회수 철회?
    				    iDontCareCnt = iDontCareCnt+1
    				    '''

					else

					    '' New CASE
					    Dim TTT : TTT=1/0  '' raseErr
					    rw "shppStatCd:"&shppStatCd
					    rw "ordItemStatCd:"&ordItemStatCd

					response.end
					end if


			    Next

			    '' CS 마스터정보. 업데이트?  위치변경 2018/07/20 eastone
				strSql = " update c "
				strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
				strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
				strSql = strSql + " , c.OrderName = o.OrderName "
				strSql = strSql + " from "
				strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
				strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
				strSql = strSql + " on "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.SellSite = o.SellSite "
				strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
				strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
				strSql = strSql + " where "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.orderserial is NULL "
				strSql = strSql + " 	and o.orderserial is not NULL "
				strSql = strSql + " 	and c.sellsite = '"&CMALLNAME&"' "
				dbget.Execute strSql

				strSql = " update c "
				strSql = strSql + " set c.currstate = 'B007' "
				strSql = strSql + " from "
				strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
				strSql = strSql + " left join db_temp.dbo.tbl_xSite_TMPOrder o "
				strSql = strSql + " on "
				strSql = strSql + "		1 = 1 "
				strSql = strSql + "		and c.SellSite = o.SellSite "
				strSql = strSql + "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
				strSql = strSql + "		and c.OrgDetailKey = o.OrgDetailKey "
				strSql = strSql + " where "
				strSql = strSql + "		1 = 1 "
				strSql = strSql + "		and c.orderserial is NULL "
				strSql = strSql + "		and o.SellSite is NULL "
				strSql = strSql + "		and c.SellSite = '" & CMALLNAME & "' "
				strSql = strSql + "		and c.currstate = 'B001' "
				strSql = strSql + "		and c.divcd = 'A008' "
				''rw strSql
				dbget.Execute strSql

			    '2018-05-09 김진영 수정..중복건 발생..'18050927907', '18050497596', '18041886778', '18041125387', '18040902758'
				strSql = strSql + " exec [db_etcmall].[dbo].[usp_Ten_Outmall_Cs_OrderCancelProc] 'ssg'"
				dbget.Execute strSql
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing

	rw "======================================"
	rw "styyyymmdd:"&styyyymmdd
	rw "edyyyymmdd:"&edyyyymmdd
	rw "총CS건수:"&ttlCNT
	rw "지시상태중취소:"&iDontCareCnt
	rw "피킹상태중취소:"&iCareCnt
	rw "CS입력건수:"&iInputCnt
end function

''취소 개편 버전..2019-11-04 김진영 추가..위 getSsgCancelList는 2020. 1. 15에 닫는다 함
public function getNewSsgCancelList(styyyymmdd,edyyyymmdd)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim shppNo, shppSeq, ordNo , ordItemSeq , orordNo , orordItemSeq
    Dim shppStatCd , itemId , itemNm , shppcstCodYn , ordpeNm , rcptpeNm ,rcptpeTelno
    Dim shppDivDtlCd , shppDivDtlNm , ordQty ,dircItemQty , cnclItemQty , shppMainCd
    Dim iDivCD

    Dim ordCnclDts , ordStatNm , ordItemStatNm
    Dim ordItemStatCd , shppMainNm , ordRcpDts

    Dim oMaster, oDetailArr(0)
    Dim ttlCNT : ttlCNT=0
    Dim iDontCareCnt : iDontCareCnt=0
    Dim iCareCnt : iCareCnt=0
    Dim iInputCnt : iInputCnt=0
    Dim iAssignedRow : iAssignedRow=0
	Dim existsCnt

	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
		objXML.Option(9) = Option_TLS
		objXML.open "POST", "" & ssgSSLAPIURL&"/api/clm/cncl/ord/inquiry.ssg"
		objXML.setRequestHeader "Authorization", ssgApiKey
		objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
		objXML.setRequestHeader "Content-Type", "application/xml"

		requestBody = ""
		requestBody = requestBody & "<request>"
		requestBody = requestBody & "	<perdStrDts>"&styyyymmdd&"</perdStrDts>"
		requestBody = requestBody & "	<perdEndDts>"&edyyyymmdd&"</perdEndDts>"
		requestBody = requestBody & "</request>"
		objXML.send(requestBody)

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
' response.write objXML.responseText
' response.end
			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultMessage").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text

			Set LagrgeNode = xmlDOM.SelectNodes("/result/data")

			If Not (LagrgeNode Is Nothing) Then
			    For i = 0 To LagrgeNode.length - 1
					ordpeNm = ""
					orordNo ="": orordItemSeq =""
					ordNo ="": ordItemSeq =""
					ordItemStatCd = ""
					itemNm = ""
					itemId = ""
					ordCnclDts = ""
					cnclItemQty = ""
					dircItemQty = ""
					existsCnt = 0

					ordpeNm 		= LEFT(LagrgeNode(i).SelectSingleNode("ordpeNm").Text, 15)		'주문자명
					'orordNo			= LagrgeNode(i).SelectSingleNode("orordNo").Text			'원주문번호
					'orordItemSeq	= LagrgeNode(i).SelectSingleNode("orordItemSeq").Text			'원주문상품순번
					ordNo			= LagrgeNode(i).SelectSingleNode("ordNo").Text					'주문번호

                    if NOT (LagrgeNode(i).SelectSingleNode("orordNo") is Nothing) then
                        orordNo             = LagrgeNode(i).SelectSingleNode("orordNo").Text		'원주문번호
                        ordNo = orordNo ''2017/12/27 추가
                    end if

					ordItemSeq		= LagrgeNode(i).SelectSingleNode("ordItemSeq").Text				'주문상품순번
                    if NOT (LagrgeNode(i).SelectSingleNode("orordItemSeq") is Nothing) then
                        orordItemSeq        = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''원주문순번 [2]
                        ordItemSeq = orordItemSeq ''2017/12/27 추가
                    end if

					ordItemStatCd	= LagrgeNode(i).SelectSingleNode("ordItemStatCd").Text			'주문상품상태코드 | 180 주문취소
					itemNm			= LagrgeNode(i).SelectSingleNode("itemNm").Text					'상품명
					itemId			= LagrgeNode(i).SelectSingleNode("itemId").Text					'상품ID
					If NOT (LagrgeNode(i).SelectSingleNode("ordCnclDts") is Nothing) then			'주문취소일시
						ordCnclDts	= Left(LagrgeNode(i).SelectSingleNode("ordCnclDts").Text,19)	'없는경우있음 20180218490344
					End If
					cnclItemQty		= LagrgeNode(i).SelectSingleNode("cnclItemQty").Text			'취소수량(클레임단위)
					dircItemQty		= LagrgeNode(i).SelectSingleNode("dircItemQty").Text        	'지시수량(원주문단위)

					ttlCNT = ttlCNT+1

					iDivCD = "A008"  ''취소

					strSql = " SELECT COUNT(*) as cnt "
					strSql = strSql & "FROM db_temp.dbo.tbl_xSite_TMPCS "
					strSql = strSql & "WHERE SellSite = '"&CMALLNAME&"' and OutMallOrderSerial in ('" & CStr(ordNo) & "', '" & CStr(ordNo) & "_1', '" & CStr(ordNo) & "_2', '" & CStr(ordNo) & "_3') and OrgDetailKey = '" & CStr(ordItemSeq) & "' "
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					if (Not rsget.Eof) then
						existsCnt 	= rsget("cnt")
					end if
					rsget.Close

					If existsCnt = 0 Then
						strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '"&CMALLNAME&"' and OutMallOrderSerial in ('" & CStr(ordNo) & "', '" & CStr(ordNo) & "_1', '" & CStr(ordNo) & "_2', '" & CStr(ordNo) & "_3') and OrgDetailKey = '" & CStr(ordItemSeq) & "' ) "
						strSql = strSql & " BEGIN "
						strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
						strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
						strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
						strSql = strSql & " 	('"&iDivCD&"', '단순변심', '"&CMALLNAME&"', '" & CStr(ordNo) & "', '"& ordpeNm &"', '', '"& "" &"', '"& "" &"', '"& ordpeNm &"', "
						strSql = strSql & "		'', '', '', '', '', '' "
						strSql = strSql & "		, '" & html2db(CStr(ordCnclDts)) & "', '" & CStr(ordItemSeq) & "', '', '"&cnclItemQty&"'); "

						strSql = strSql & "	update c "
						strSql = strSql & "	set c.OutMallOrderSerial = o.OutMallOrderSerial "
						strSql = strSql & "	from "
						strSql = strSql & "		db_temp.dbo.tbl_xSite_TMPCS c "
						strSql = strSql & "		join db_temp.dbo.tbl_xSite_TMPOrder o "
						strSql = strSql & "		on "
						strSql = strSql & "			1 = 1 "
						strSql = strSql & "			and c.SellSite = o.SellSite "
						strSql = strSql & "			and o.OutMallOrderSerial in (c.OutMallOrderSerial, c.OutMallOrderSerial + '_1', c.OutMallOrderSerial + '_2', c.OutMallOrderSerial + '_3') "
						strSql = strSql & "			and c.orgdetailkey = o.orgdetailkey "
						strSql = strSql & "	where "
						strSql = strSql & "		1 = 1 "
						strSql = strSql & "		and c.SellSite = '"&CMALLNAME&"' "
						strSql = strSql & "		and c.OutMallOrderSerial = '" & CStr(ordNo) & "' "
						strSql = strSql & "		and c.orgdetailkey = '" & CStr(ordItemSeq) & "' "
						strSql = strSql & "		and c.OutMallOrderSerial <> o.OutMallOrderSerial "
						strSql = strSql & " END "
						dbget.Execute strSql,iAssignedRow
					End If

					If (iAssignedRow > 0) Then
						iInputCnt = iInputCnt + iAssignedRow
					End If

					''주문 입력 이전 내역은 삭제 하자
					strSql = " update c "
					strSql = strSql + " set matchState='D'"
					strSql = strSql + " from db_temp.dbo.tbl_xSite_TMPOrder c "
					strSql = strSql + " WHERE SellSite = '"&CMALLNAME&"' and OutMallOrderSerial in ('" & CStr(ordNo) & "', '" & CStr(ordNo) & "_1', '" & CStr(ordNo) & "_2', '" & CStr(ordNo) & "_3') and OrgDetailKey = '" & CStr(ordItemSeq) & "'"
					strSql = strSql + " and orderserial is NULL"
					dbget.Execute strSql
					iCareCnt = iCareCnt + 1

					' rw LagrgeNode(i).SelectSingleNode("ordpeNm").Text				'주문자명
					' rw LagrgeNode(i).SelectSingleNode("ordRcpMediaCd").Text		'접수매체코드 | 10 PC웹, 20 모바일웹, 30 모바일앱(iOS), 40 모바일앱(안드로이드), 50 패드앱(iOS), 60 패드앱(안드로이드), 70 스마트기기, 80 제휴채널, 90 점포, 99 콜센터
					' rw LagrgeNode(i).SelectSingleNode("ordRcpMediaNm").Text		'접수매체명
					' rw LagrgeNode(i).SelectSingleNode("orordNo").Text				'원주문번호
					' rw LagrgeNode(i).SelectSingleNode("orordItemSeq").Text		'원주문상품순번
					' rw LagrgeNode(i).SelectSingleNode("ordNo").Text				'주문번호
					' rw LagrgeNode(i).SelectSingleNode("ordItemSeq").Text			'주문상품순번
					' rw LagrgeNode(i).SelectSingleNode("ordItemStatCd").Text		'주문상품상태코드 | 180 주문취소
					' rw LagrgeNode(i).SelectSingleNode("ordItemStatNm").Text		'주문상품상태명
					' rw LagrgeNode(i).SelectSingleNode("itemNm").Text				'상품명
					' rw LagrgeNode(i).SelectSingleNode("itemId").Text				'상품ID
					' rw LagrgeNode(i).SelectSingleNode("clmRsnCd").Text			'클레임사유코드 | 251 단순취소, 252 사이즈/색상 옵션 변경, 253 재주문, 258 배송지 변경, 259 최저가아님, 260 배송시간변경
					' rw LagrgeNode(i).SelectSingleNode("clmRsnNm").Text			'클레임사유명
					' If NOT (LagrgeNode(i).SelectSingleNode("clmRsnCntt") is Nothing) then
					' 	rw LagrgeNode(i).SelectSingleNode("clmRsnCntt").Text & "!!!!!!!!!!!!!!!!!!!!"	'클레임사유내용
					' End If
					' rw LagrgeNode(i).SelectSingleNode("imptDivCd").Text			'귀책구분코드 | 10 고객, 20 업체
					' rw LagrgeNode(i).SelectSingleNode("ordCnclDts").Text		'주문취소일시
					' rw LagrgeNode(i).SelectSingleNode("cnclItemQty").Text		'취소수량(클레임단위)
					' rw LagrgeNode(i).SelectSingleNode("ordRcpDts").Text			'주문접수일시
					' rw LagrgeNode(i).SelectSingleNode("dircItemQty").Text		'지시수량(원주문단위)
				Next
			    '' CS 마스터정보. 업데이트?  위치변경 2018/07/20 eastone
				strSql = " update c "
				strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
				strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
				strSql = strSql + " , c.OrderName = o.OrderName "
				strSql = strSql + " , c.ReceiveName = o.ReceiveName "
				strSql = strSql + " from "
				strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
				strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
				strSql = strSql + " on "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.SellSite = o.SellSite "
				strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
				strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
				strSql = strSql + " where "
				strSql = strSql + " 	1 = 1 "
				strSql = strSql + " 	and c.orderserial is NULL "
				strSql = strSql + " 	and o.orderserial is not NULL "
				strSql = strSql + " 	and c.sellsite = '"&CMALLNAME&"' "
				dbget.Execute strSql

				strSql = " update c "
				strSql = strSql + " set c.currstate = 'B007' "
				strSql = strSql + " from "
				strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
				strSql = strSql + " left join db_temp.dbo.tbl_xSite_TMPOrder o "
				strSql = strSql + " on "
				strSql = strSql + "		1 = 1 "
				strSql = strSql + "		and c.SellSite = o.SellSite "
				strSql = strSql + "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
				strSql = strSql + "		and c.OrgDetailKey = o.OrgDetailKey "
				strSql = strSql + " where "
				strSql = strSql + "		1 = 1 "
				strSql = strSql + "		and c.orderserial is NULL "
				strSql = strSql + "		and o.SellSite is NULL "
				strSql = strSql + "		and c.SellSite = '" & CMALLNAME & "' "
				strSql = strSql + "		and c.currstate = 'B001' "
				strSql = strSql + "		and c.divcd = 'A008' "
				''rw strSql
				dbget.Execute strSql

			    '2018-05-09 김진영 수정..중복건 발생..'18050927907', '18050497596', '18041886778', '18041125387', '18040902758'
				strSql = strSql + " exec [db_etcmall].[dbo].[usp_Ten_Outmall_Cs_OrderCancelProc] 'ssg'"
				dbget.Execute strSql
			End If
			Set LagrgeNode = nothing
	    Set xmlDOM = nothing
	Set objXML = nothing
	rw "======================================"
	rw "styyyymmdd:"&styyyymmdd
	rw "edyyyymmdd:"&edyyyymmdd
	rw "총CS건수:"&ttlCNT
	rw "CS입력건수:"&iInputCnt
End Function

public function sendSsgReturnConfirm(shppNo, shppSeq, procItemQty, resellPsblYn, retImptMainCd, ByRef resultDesc)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim ordNo , ordItemSeq , orordNo , orordItemSeq
    Dim shppStatCd , shppStatNm, itemId , itemNm , shppcstCodYn , ordpeNm , rcptpeNm , rcptpeHpno , rcptpeTelno
    Dim shppDivDtlCd , shppDivDtlNm , ordQty ,dircItemQty , shppMainCd
    Dim iDivCD

    Dim rcovDircDt , ordStatNm , ordItemStatNm
    Dim ordItemStatCd , shppMainNm , ordRcpDts

    Dim oMaster, oDetailArr(0)
    Dim ttlCNT : ttlCNT=0
    Dim iDontCareCnt : iDontCareCnt=0
    Dim iCareCnt : iCareCnt=0
    Dim iInputCnt : iInputCnt=0
    Dim iAssignedRow : iAssignedRow=0

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/saveConfirmRcov.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestConfirmRcov>"
    requestBody = requestBoDy&"<shppNo>"&shppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&shppSeq&"</shppSeq>"
	requestBody = requestBoDy&"<procItemQty>"&procItemQty&"</procItemQty>"
    requestBody = requestBoDy&"<shppTypeDtlCd>22</shppTypeDtlCd>"				'// 22 업체택배배송
    requestBody = requestBoDy&"<delicoVenId>0000033028</delicoVenId>"
    requestBody = requestBoDy&"<wblNo>9999</wblNo>"
    requestBody = requestBoDy&"<resellPsblYn>"&resellPsblYn&"</resellPsblYn>"			'// 재판매가능 구분
    requestBody = requestBoDy&"<retImptMainCd>"&retImptMainCd&"</retImptMainCd>"		'// 귀책사유주체 : 10 고객 20 판매자 30 택배사
    requestBody = requestBoDy&"</requestConfirmRcov>"

	objXML.send(requestBody)
	''rw objXML.status
	''response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			resultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text

	Set objXML = nothing

	sendSsgReturnConfirm = ssgresultCode
end function

public function sendSsgReturnFinish(shppNo, shppSeq, procItemQty, resellPsblYn, retImptMainCd, ByRef resultDesc)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim ordNo , ordItemSeq , orordNo , orordItemSeq
    Dim shppStatCd , shppStatNm, itemId , itemNm , shppcstCodYn , ordpeNm , rcptpeNm , rcptpeHpno , rcptpeTelno
    Dim shppDivDtlCd , shppDivDtlNm , ordQty ,dircItemQty , shppMainCd
    Dim iDivCD

    Dim rcovDircDt , ordStatNm , ordItemStatNm
    Dim ordItemStatCd , shppMainNm , ordRcpDts

    Dim oMaster, oDetailArr(0)
    Dim ttlCNT : ttlCNT=0
    Dim iDontCareCnt : iDontCareCnt=0
    Dim iCareCnt : iCareCnt=0
    Dim iInputCnt : iInputCnt=0
    Dim iAssignedRow : iAssignedRow=0

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/saveCompleteRcov.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestConfirmRcov>"
    requestBody = requestBoDy&"<shppNo>"&shppNo&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&shppSeq&"</shppSeq>"
	requestBody = requestBoDy&"<procItemQty>"&procItemQty&"</procItemQty>"
    requestBody = requestBoDy&"<shppTypeDtlCd>22</shppTypeDtlCd>"				'// 22 업체택배배송
    requestBody = requestBoDy&"<delicoVenId>0000033028</delicoVenId>"
    requestBody = requestBoDy&"<wblNo>9999</wblNo>"
    requestBody = requestBoDy&"<resellPsblYn>"&resellPsblYn&"</resellPsblYn>"			'// 재판매가능 구분
    requestBody = requestBoDy&"<retImptMainCd>"&retImptMainCd&"</retImptMainCd>"		'// 귀책사유주체 : 10 고객 20 판매자 30 택배사
    requestBody = requestBoDy&"</requestConfirmRcov>"

	objXML.send(requestBody)
	''rw objXML.status
	''response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			resultDesc = xmlDOM.getElementsByTagName("resultDesc").Item(0).Text

	Set objXML = nothing

	sendSsgReturnFinish = ssgresultCode
end function

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
