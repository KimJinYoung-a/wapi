<%
function getCardRibonName(cardribbon)
    if IsNULL(cardribbon) then Exit Function

    if (cardribbon="1") then
        getCardRibonName  = "카드"
    elseif (cardribbon="2") then
        getCardRibonName  = "리본"
    elseif (cardribbon="3") then
        getCardRibonName  = "없음"
    end if
end function

function FinishCSMaster(iAsid, finishuser, contents_finish)
    dim sqlStr

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"                      + VbCrlf
    sqlStr = sqlStr + " set finishuser='" + finishuser + "'"            + VbCrlf
    sqlStr = sqlStr + " , contents_finish='" + contents_finish + "'"    + VbCrlf
    sqlStr = sqlStr + " , finishdate=getdate()"                         + VbCrlf
    sqlStr = sqlStr + " , currstate='B007'"                             + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(iAsid)

    dbget.Execute sqlStr

    ''웹처리 인경우 완료내역을 OpenContents에 저장.
    Call AddCustomerOpenContents(iAsid, contents_finish)
end function

function SetStockOutByCsAs(iAsid)
    dim sqlStr
    dim resultCount	: resultCount = 0
    dim arrItemID

	'// 업배상품만 품절 등록

	'// =======================================================================
	sqlStr = " select IsNull(count(i.itemid), 0) as cnt " + VbCrLf
    sqlStr = sqlStr + " from " + VbCrLf
    sqlStr = sqlStr + " db_cs.dbo.tbl_new_as_list m " + VbCrLf
    sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_detail d " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	m.id = d.masterid " + VbCrLf
    sqlStr = sqlStr + " join db_item.dbo.tbl_item i " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	d.itemid = i.itemid " + VbCrLf
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// 미출고사유 품절일 경우만, 2022-02-24, skyer9
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and T.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	and T.detailidx = d.orderdetailidx "
    sqlStr = sqlStr + " 	and T.code = '05' "
    sqlStr = sqlStr + " where " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and m.id = " + CStr(iAsid) + " " + VbCrLf
    sqlStr = sqlStr + " 	and ((m.divcd in ('A008', 'A112')) or ((m.divcd = 'A900') and (d.confirmitemno < 0))) " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun01 = 'C004' " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun02 = 'CD05' " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid <> 0 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption = '0000' " + VbCrLf
    sqlStr = sqlStr + " 	and d.isupchebeasong = 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and m.deleteyn <> 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and i.sellyn = 'Y' " + VbCrLf
    rsget.Open sqlStr,dbget,1
        resultCount = resultCount + rsget("cnt")
    rsget.Close

	sqlStr = " select IsNull(count(o.itemid), 0) as cnt " + VbCrLf
    sqlStr = sqlStr + " from " + VbCrLf
    sqlStr = sqlStr + " db_cs.dbo.tbl_new_as_list m " + VbCrLf
    sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_detail d " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	m.id = d.masterid " + VbCrLf
    sqlStr = sqlStr + " join [db_item].[dbo].tbl_item_option o " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid = o.itemid " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption = o.itemoption " + VbCrLf
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// 미출고사유 품절일 경우만, 2022-02-24, skyer9
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and T.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	and T.detailidx = d.orderdetailidx "
    sqlStr = sqlStr + " 	and T.code = '05' "
    sqlStr = sqlStr + " where " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and m.id = " + CStr(iAsid) + " " + VbCrLf
    sqlStr = sqlStr + " 	and ((m.divcd in ('A008', 'A112')) or ((m.divcd = 'A900') and (d.confirmitemno < 0))) " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun01 = 'C004' " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun02 = 'CD05' " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid <> 0 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption <> '0000' " + VbCrLf
    sqlStr = sqlStr + " 	and d.isupchebeasong = 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and m.deleteyn <> 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and o.optsellyn = 'Y' " + VbCrLf
    rsget.Open sqlStr,dbget,1
        resultCount = resultCount + rsget("cnt")
    rsget.Close

    '// =======================================================================
    SetStockOutByCsAs = resultCount
    if (resultCount < 1) then
        exit function
    end if

    '// =======================================================================
    '// 1. 옵션 없는 상품(일시품절 전환)
    sqlStr = " update i " + VbCrLf
    sqlStr = sqlStr + " set i.sellyn = 'S', i.lastupdate = getdate() " + VbCrLf
    sqlStr = sqlStr + " from " + VbCrLf
    sqlStr = sqlStr + " db_cs.dbo.tbl_new_as_list m " + VbCrLf
    sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_detail d " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	m.id = d.masterid " + VbCrLf
    sqlStr = sqlStr + " join db_item.dbo.tbl_item i " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	d.itemid = i.itemid " + VbCrLf
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// 미출고사유 품절일 경우만, 2022-02-24, skyer9
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and T.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	and T.detailidx = d.orderdetailidx "
    sqlStr = sqlStr + " 	and T.code = '05' "
    sqlStr = sqlStr + " where " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and m.id = " + CStr(iAsid) + " " + VbCrLf
    sqlStr = sqlStr + " 	and ((m.divcd in ('A008', 'A112')) or ((m.divcd = 'A900') and (d.confirmitemno < 0))) " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun01 = 'C004' " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun02 = 'CD05' " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid <> 0 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption = '0000' " + VbCrLf
    sqlStr = sqlStr + " 	and d.isupchebeasong = 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and m.deleteyn <> 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and i.sellyn = 'Y' " + VbCrLf
    'response.write sqlStr
	rsget.Open sqlStr,dbget

    '// =======================================================================
	'// 2-1. 옵션 있는 상품(상품코드목록)
	sqlStr = " select distinct o.itemid " + VbCrLf
    sqlStr = sqlStr + " from " + VbCrLf
    sqlStr = sqlStr + " db_cs.dbo.tbl_new_as_list m " + VbCrLf
    sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_detail d " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	m.id = d.masterid " + VbCrLf
    sqlStr = sqlStr + " join [db_item].[dbo].tbl_item_option o " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid = o.itemid " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption = o.itemoption " + VbCrLf
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// 미출고사유 품절일 경우만, 2022-02-24, skyer9
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and T.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	and T.detailidx = d.orderdetailidx "
    sqlStr = sqlStr + " 	and T.code = '05' "
    sqlStr = sqlStr + " where " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and m.id = " + CStr(iAsid) + " " + VbCrLf
    sqlStr = sqlStr + " 	and ((m.divcd in ('A008', 'A112')) or ((m.divcd = 'A900') and (d.confirmitemno < 0))) " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun01 = 'C004' " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun02 = 'CD05' " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid <> 0 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption <> '0000' " + VbCrLf
    sqlStr = sqlStr + " 	and d.isupchebeasong = 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and m.deleteyn <> 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and o.optsellyn = 'Y' " + VbCrLf
    'response.write sqlStr
    rsget.Open sqlStr,dbget,1

    arrItemID = "-1"
	do until rsget.Eof
		arrItemID = arrItemID + "," + CStr(rsget("itemid"))
		rsget.MoveNext
	loop
	rsget.Close

	'// 2-2. 옵션 있는 상품(품절전환)
	sqlStr = " update o " + VbCrLf
	sqlStr = sqlStr + " set o.isusing = 'N', o.optsellyn = 'N' " + VbCrLf
    sqlStr = sqlStr + " from " + VbCrLf
    sqlStr = sqlStr + " db_cs.dbo.tbl_new_as_list m " + VbCrLf
    sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_detail d " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	m.id = d.masterid " + VbCrLf
    sqlStr = sqlStr + " join [db_item].[dbo].tbl_item_option o " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid = o.itemid " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption = o.itemoption " + VbCrLf
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// 미출고사유 품절일 경우만, 2022-02-24, skyer9
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and T.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	and T.detailidx = d.orderdetailidx "
    sqlStr = sqlStr + " 	and T.code = '05' "
    sqlStr = sqlStr + " where " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and m.id = " + CStr(iAsid) + " " + VbCrLf
    sqlStr = sqlStr + " 	and ((m.divcd in ('A008', 'A112')) or ((m.divcd = 'A900') and (d.confirmitemno < 0))) " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun01 = 'C004' " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun02 = 'CD05' " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid <> 0 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption <> '0000' " + VbCrLf
    sqlStr = sqlStr + " 	and d.isupchebeasong = 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and m.deleteyn <> 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and o.optsellyn = 'Y' " + VbCrLf
    'response.write sqlStr
	rsget.Open sqlStr,dbget

	'// 2-3. 옵션 있는 상품(옵션갯수)
	sqlStr = " update i " + VbCrLf
	sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0), lastupdate = getdate() " + VbCrLf
	sqlStr = sqlStr + " from " + VbCrLf
	sqlStr = sqlStr + " 	[db_item].[dbo].tbl_item i " + VbCrLf
	sqlStr = sqlStr + " 	join ( " + VbCrLf
	sqlStr = sqlStr + " 		select itemid, sum(case when isusing = 'Y' then 1 else 0 end) optioncnt " + VbCrLf
	sqlStr = sqlStr + " 		from [db_item].[dbo].tbl_item_option " + VbCrLf
	sqlStr = sqlStr + " 		where itemid in ( " + VbCrLf
	sqlStr = sqlStr + " 			" + CStr(arrItemID) + " " + VbCrLf
	sqlStr = sqlStr + " 		) " + VbCrLf
	''sqlStr = sqlStr + " 		and isusing='Y'" + VBCrlf
	sqlStr = sqlStr + " 		group by itemid " + VbCrLf
	sqlStr = sqlStr + " 	) T " + VbCrLf
	sqlStr = sqlStr + " 	on " + VbCrLf
	sqlStr = sqlStr + " 		i.itemid = T.itemid " + VbCrLf
	'response.write sqlStr
	dbget.Execute sqlStr

	'// 2-4. 옵션 있는 상품(판매중인 옵션이 없으면 품절처리)
    sqlStr = " update [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " set sellyn='N'"
	sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
	sqlStr = sqlStr + " where itemid in (" + CStr(arrItemID) + ") "
	sqlStr = sqlStr + " and optioncnt=0"
	'response.write sqlStr
    dbget.Execute sqlStr

end function

function GetDefaultTitle(divcd, id, orderserial)
    dim ipkumdiv, accountdiv, cancelyn, comm_name, ipkumdivName, accountdivName, pggubun, comm_cd
    dim sqlStr

    sqlStr = " select m.ipkumdiv, m.accountdiv, m.cancelyn, C.comm_name, isNULL(m.pggubun,'') as pggubun, C.comm_cd"
    sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
    sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_list A"
    sqlStr = sqlStr + "     on A.orderserial='" + orderserial + "'"
    if (id<>"") then
        sqlStr = sqlStr + " and A.id=" + CStr(id)
    end if
    sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_cs_comm_code C"
    sqlStr = sqlStr + " on C.comm_cd='" + divcd + "'"

    sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        ipkumdiv    = rsget("ipkumdiv")
        cancelyn    = rsget("cancelyn")
        comm_name   = rsget("comm_name")
        accountdiv  = Trim(rsget("accountdiv"))
        pggubun     = rsget("pggubun")
        comm_cd     = rsget("comm_cd")
    end if
    rsget.close


    if (ipkumdiv="2") then
        ipkumdivName = "입금 대기"
    elseif (ipkumdiv="4") then
        ipkumdivName = "결제 완료"
    elseif (ipkumdiv="5") then
        ipkumdivName = "상품 준비"
    elseif (ipkumdiv="6") then
        ipkumdivName = "출고 준비"
    elseif (ipkumdiv="7") then
        ipkumdivName = "일부 출고"
    elseif (ipkumdiv="8") then
        ipkumdivName = "출고 완료"
    end if

    if (accountdiv="7") then
        accountdivName = "무통장"
    elseif (accountdiv="14") then
        accountdivName = "편의점결제"
    elseif (accountdiv="100") then
        accountdivName = "신용카드"
    elseif (accountdiv="80") then
        accountdivName = "올엣카드"
    elseif (accountdiv="50") then
        accountdivName = "제휴몰결제"
    elseif (accountdiv="20") then
        accountdivName = "실시간이체"
    elseif (accountdiv="400") then
        accountdivName = "핸드폰"
    elseif (accountdiv="150") then
        accountdivName = "이니렌탈"
    end if

    ''2016/08/04
    if (pggubun="NP") then
        accountdivName = "네이버페이"
        if (comm_cd="A007") then
            comm_name = "네이버페이 취소요청"
        end if
    end if

    GetDefaultTitle = accountdivName + " " + ipkumdivName + " 상태 중 " + comm_name
end function


function SetCustomerOpenMsg(id, opentitle, opencontents)
    dim sqlStr

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"        + VbCrlf
    sqlStr = sqlStr + " set opentitle='" + opentitle + "'"  + VbCrlf
    sqlStr = sqlStr + " , opencontents='" + opencontents + "'" + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)

    dbget.Execute sqlStr

end function

function AddCSMasterRefundInfo(asid, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

    dim sqlStr

    sqlStr = " update "
    sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_as_refund_info "
    sqlStr = sqlStr + " set "
    sqlStr = sqlStr + " 	orggiftcardsum = " & CStr(orggiftcardsum) & " "
    sqlStr = sqlStr + " 	, orgdepositsum = " & CStr(orgdepositsum) & " "
    sqlStr = sqlStr + " 	, refundgiftcardsum = " & CStr(refundgiftcardsum) & " "
    sqlStr = sqlStr + " 	, refunddepositsum = " & CStr(refunddepositsum) & " "
    sqlStr = sqlStr + " where asid = " & CStr(asid) & " "

	'response.write "aaaaaaaaaaa" & sqlStr
    dbget.Execute sqlStr

end function

function EditCSMasterRefundEncInfo(asid, encmethod, bnkaccount)
    dim sqlStr
    ''2017/10/02 암호화 방식 변경
    sqlStr = "exec db_cs.[dbo].[sp_Ten_EditCSMasterRefundEncInfo] "&CStr(asid)&",'"&encmethod&"','"&bnkaccount&"'"
    dbget.Execute sqlStr
    exit function

    IF (encmethod="PH1") then
        IF (bnkaccount="") then
            sqlStr = " update [db_cs].[dbo].tbl_as_refund_info " & VbCRLF
            sqlStr = sqlStr + " set encmethod = '' " & VbCRLF
            sqlStr = sqlStr + " 	, encaccount = NULL" & VbCRLF
            sqlStr = sqlStr + " 	, rebankaccount=''" & VbCRLF
            sqlStr = sqlStr + " where asid = " & CStr(asid) & " " & VbCRLF

            dbget.Execute sqlStr
        ELSE
            sqlStr = " update [db_cs].[dbo].tbl_as_refund_info " & VbCRLF
            sqlStr = sqlStr + " set encmethod = '" & Left(CStr(encmethod), 8) & "' " & VbCRLF
            sqlStr = sqlStr + " 	, encaccount = db_cs.dbo.uf_EncAcctPH1('"&bnkaccount&"')" & VbCRLF
            sqlStr = sqlStr + " 	, rebankaccount=''" & VbCRLF
            sqlStr = sqlStr + " where asid = " & CStr(asid) & " " & VbCRLF

            dbget.Execute sqlStr
        END IF
    end IF

end function

function AddCustomerOpenContents(id, addcontents)
    dim sqlStr

    if ((addcontents="") or (id="")) then Exit Function

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"        + VbCrlf
    sqlStr = sqlStr + " set opencontents=IsNULL(opencontents,'') + (Case When (IsNULL(opencontents,'')='') then '" & addcontents & "' else '" & VbCrlf & addcontents + "' End )" + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)

    dbget.Execute sqlStr

end function

function RegCSMasterAddUpche(id, imakerid)
    dim sqlStr
    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"    + VbCrlf
    sqlStr = sqlStr + " set makerid='" + imakerid + "'"   + VbCrlf
    sqlStr = sqlStr + " , requireupche='Y'"               + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)

    dbget.Execute sqlStr
end function

function RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    '' CS Master 저장
    dim sqlStr, InsertedId
    sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
    rsget.Open sqlStr,dbget,1,3
    rsget.AddNew
        rsget("divcd")          = divcd
    	rsget("orderserial")    = orderserial
    	rsget("customername")   = ""
    	rsget("userid")         = ""
    	rsget("writeuser")      = reguserid
    	rsget("title")          = title
    	rsget("contents_jupsu") = server.htmlencode(contents_jupsu)
    	rsget("gubun01")        = gubun01
    	rsget("gubun02")        = gubun02

    	rsget("currstate")      = "B001"
    	rsget("deleteyn")       = "N"

        ''''''''''''''''''''''''''''''''''
    	''rsget("requireupche")   = "N"
    	''rsget("makerid")        = ""
    	''''''''''''''''''''''''''''''''''

    rsget.update
	    InsertedId = rsget("id")
	rsget.close

	dim opentitle, opencontents
	opentitle = GetDefaultTitle(divcd, InsertedId, orderserial)

	opencontents = ""


	''set Default openContents
	sqlStr = " update [db_cs].[dbo].tbl_new_as_list"  + VbCrlf
	sqlStr = sqlStr + " set userid=T.userid"        + VbCrlf
	sqlStr = sqlStr + " , customername=T.buyname"   + VbCrlf
	sqlStr = sqlStr + " , opentitle='" + html2db(opentitle) + "'" + VbCrlf
	sqlStr = sqlStr + " , opencontents='" + html2db(opencontents) + "'" + VbCrlf
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master T" + VbCrlf
	sqlStr = sqlStr + " where T.orderserial='" + orderserial + "'"  + VbCrlf
	sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_list.id=" + CStr(InsertedId)
	dbget.Execute sqlStr

	dim IsUpdateSuccess
	IsUpdateSuccess = False
	sqlStr = " select @@rowcount as cnt "
	'response.write sqlStr

    rsget.Open sqlStr,dbget,1
        IsUpdateSuccess = (rsget("cnt") > 0)
    rsget.Close

	if (Not IsNumeric(orderserial)) and (IsUpdateSuccess = False) then
		'Gift카드 주문인지 확인한다
		sqlStr = " update [db_cs].[dbo].tbl_new_as_list"  + VbCrlf
		sqlStr = sqlStr + " set userid=T.userid"        + VbCrlf
		sqlStr = sqlStr + " , customername=T.buyname"   + VbCrlf
		sqlStr = sqlStr + " , opentitle='" + title + "'" + VbCrlf
		sqlStr = sqlStr + " , opencontents=''" + VbCrlf
		sqlStr = sqlStr + " , extsitename='giftcard' "   + VbCrlf
    	sqlStr = sqlStr + " from [db_order].[dbo].tbl_giftcard_order T" + VbCrlf
		sqlStr = sqlStr + " where T.giftorderserial='" + orderserial + "'"  + VbCrlf
		sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_list.id=" + CStr(InsertedId)
		dbget.Execute sqlStr
	end if

	RegCSMaster = InsertedId
end function


function RegWebCSDetailAllCancel(byval CsId, orderserial)
	dim sqlStr

	sqlStr = " Insert into [db_cs].[dbo].tbl_new_as_detail"
    sqlStr = sqlStr + " (masterid, orderdetailidx, gubun01,gubun02"
    sqlStr = sqlStr + " ,orderserial, itemid, itemoption, makerid, itemname, itemoptionname, regitemno, confirmitemno,orderitemno, itemcost, buycash, isupchebeasong,regdetailstate) "
    sqlStr = sqlStr + " select " + CStr(CsId) + ", d.idx, c.gubun01, c.gubun02"
    sqlStr = sqlStr + " , d.orderserial, d.itemid, d.itemoption, d.makerid, d.itemname, d.itemoptionname, d.itemno, d.itemno, d.itemno, d.itemcost, d.buycash, d.isupchebeasong,d.currstate"
    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list c"
    sqlStr = sqlStr + " ,[db_order].[dbo].tbl_order_detail d"
    sqlStr = sqlStr + " where c.id=" + CStr(CsId)
    sqlStr = sqlStr + " and d.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and c.orderserial=d.orderserial"
    'sqlStr = sqlStr + " and d.itemid<>0"								'배송비도 넣는다.
    sqlStr = sqlStr + " and d.cancelyn <> 'Y' "							'CS에서 일부취소후 프런트에서 전부취소하는 경우

    dbget.Execute sqlStr
end function

function RegWebCSDetailStockoutCancel(byval CsId, orderserial)
	dim sqlStr

	sqlStr = " Insert into [db_cs].[dbo].tbl_new_as_detail"
    sqlStr = sqlStr + " (masterid, orderdetailidx, gubun01,gubun02"
    sqlStr = sqlStr + " ,orderserial, itemid, itemoption, makerid, itemname, itemoptionname, regitemno, confirmitemno,orderitemno, itemcost, buycash, isupchebeasong,regdetailstate) "
    sqlStr = sqlStr + " select " + CStr(CsId) + ", d.idx, c.gubun01, c.gubun02"
    sqlStr = sqlStr + " , d.orderserial, d.itemid, d.itemoption, d.makerid, d.itemname, d.itemoptionname, (case when d.itemid = 0 then d.itemno else IsNull(m.itemlackno,0) end), (case when d.itemid = 0 then d.itemno else IsNull(m.itemlackno,0) end), d.itemno, d.itemcost, d.buycash, d.isupchebeasong,d.currstate"
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list c "
    sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and c.id = " & CStr(CsId)
    sqlStr = sqlStr + " 		and c.orderserial=d.orderserial "
    sqlStr = sqlStr + " 	left join db_temp.dbo.tbl_mibeasong_list m "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		d.idx = m.detailidx "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and d.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 	and IsNull(d.currstate, '0') < '7' "
    sqlStr = sqlStr + " 	and ((IsNull(m.itemlackno,0) > 0) or (d.itemid = 0)) "
    sqlStr = sqlStr + " 	and ( "
    sqlStr = sqlStr + " 		((d.itemid <> 0) and (IsNull(m.code, '') in ('05','06')))"
    'sqlStr = sqlStr + " 		((d.itemid <> 0) and (IsNull(m.code, '') in ('05') or (IsNull(m.code, '') in ('03') and d.isupchebeasong='N'))) "
    sqlStr = sqlStr + " 		or "
    sqlStr = sqlStr + " 		((d.itemid = 0) and (d.makerid in ( "
    sqlStr = sqlStr + " 			select "
    sqlStr = sqlStr + " 				(case when d.isupchebeasong = 'Y' or d.itemid = 0 then d.makerid else '' end) as makerid "
    sqlStr = sqlStr + " 			from "
    sqlStr = sqlStr + " 			[db_order].[dbo].[tbl_order_detail] d "
    sqlStr = sqlStr + " 			left join db_temp.dbo.tbl_mibeasong_list m "
    sqlStr = sqlStr + " 			on "
    sqlStr = sqlStr + " 				d.idx = m.detailidx "
    sqlStr = sqlStr + " 			where "
    sqlStr = sqlStr + " 				1 = 1 "
    sqlStr = sqlStr + " 				and d.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr + " 				and d.cancelyn <> 'Y' "
    sqlStr = sqlStr + " 			group by "
    sqlStr = sqlStr + " 				(case when d.isupchebeasong = 'Y' or d.itemid = 0 then d.makerid else '' end) "
    sqlStr = sqlStr + " 			having "
    sqlStr = sqlStr + " 				sum(case when d.itemid <> 0 then d.itemno else 0 end) = sum(case when d.itemid <> 0 and IsNull(m.code, '') in ('05','06') then IsNull(m.itemlackno,0) else 0 end) "
    'sqlStr = sqlStr + " 				sum(case when d.itemid <> 0 then d.itemno else 0 end) = sum(case when d.itemid <> 0 and (IsNull(m.code, '') in ('05') or (IsNull(m.code, '') in ('03') and d.isupchebeasong='N')) then IsNull(m.itemlackno,0) else 0 end) "
    sqlStr = sqlStr + " 		))) "
    sqlStr = sqlStr + " 	) "

	dbget.Execute sqlStr

    ' 자동취소날짜를 넣는다.
    RegmibesongCanceldate(orderserial)

	'// 품절 잔여상품 출고지연 전환
	sqlStr = " update T "
    sqlStr = sqlStr + " set T.code = '03', T.itemno = (ad.orderitemno - ad.regitemno), T.itemlackno = (T.itemlackno - ad.regitemno), T.state = 0, T.reqaddstr = '품절상품자동취소' "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list a "
    sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_detail d on a.orderserial = d.orderserial "
    sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_detail ad on a.id = ad.masterid and ad.orderdetailidx = d.idx "
    sqlStr = sqlStr + " 	join db_temp.dbo.tbl_mibeasong_list T on d.idx = T.detailidx "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and a.divcd = 'A008' "
    sqlStr = sqlStr + " 	and a.id = " & CsId
    sqlStr = sqlStr + " 	and T.code = '05' "
    'sqlStr = sqlStr + " 	and (T.code in ('05') or (T.code in ('03') and d.isupchebeasong='N'))"

	dbget.Execute sqlStr

	'// 택배파업 잔여상품 출고지연 전환
	sqlStr = " update T "
    sqlStr = sqlStr + " set T.code = '03', T.itemno = (ad.orderitemno - ad.regitemno), T.itemlackno = (T.itemlackno - ad.regitemno), T.state = 0, T.reqaddstr = '택배파업자동취소' "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list a "
    sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_detail d on a.orderserial = d.orderserial "
    sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_detail ad on a.id = ad.masterid and ad.orderdetailidx = d.idx "
    sqlStr = sqlStr + " 	join db_temp.dbo.tbl_mibeasong_list T on d.idx = T.detailidx "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and a.divcd = 'A008' "
    sqlStr = sqlStr + " 	and a.id = " & CsId
    sqlStr = sqlStr + " 	and T.code = '06' "

	dbget.Execute sqlStr
end function

' 자동취소날짜를 넣는다.
' 요 쿼리 수정시. \autojob\cs_cancel_autojob.asp 에 mode cssoldoutitemcancel 같이 수정해 주세요. 한쪽이 누락될경우 자동취소완료 처리가 되지 않습니다.
function RegmibesongCanceldate(orderserial)
    dim sqlStr

    if orderserial="" or isnull(orderserial) then exit function

    sqlStr = " update l"
    sqlStr = sqlStr + " set l.isautocanceldate = getdate()"
    sqlStr = sqlStr + " from db_temp.dbo.tbl_mibeasong_list l with (nolock)"
    sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d with (nolock) on d.idx = l.detailidx "
    sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] m with (nolock) on d.orderserial = m.orderserial "
    sqlStr = sqlStr + " where l.code in ('05','06') "
    sqlStr = sqlStr + " 	and l.state <= '4' "
    sqlStr = sqlStr + " 	and l.isSendSMS = 'Y' "
    sqlStr = sqlStr + " 	and l.isSendEmail = 'Y' "
    sqlStr = sqlStr + " 	and l.sendCount > 0 "
    sqlStr = sqlStr + " 	and l.itemno>0 and l.itemlackno>0 "
    sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
    sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
    sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// 출고완료 제외
    sqlStr = sqlStr + " 	and ( "
    sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
    sqlStr = sqlStr + " 		or "
    'sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','10x10_cs')) "		'// 인터파크,LF,지마켓,옥션,11번가,쿠팡,네이버 스토어팜은 우리가 먼저 SMS 발송후 고객이 동의해야 취소진행 가능
    sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('10x10_cs')) "
    sqlStr = sqlStr + " 	) "
    'sqlStr = sqlStr + " 	and m.sitename not in ('10x10_cs')"
	sqlStr = sqlStr + " 	and l.isSendSMSdate is not null"
    sqlStr = sqlStr + " 	and datediff(hour,l.isSendSMSdate,getdate()) > 24"	' 문자발송된지24시간지난거
	'sqlStr = sqlStr + " 	and datediff(hour,l.isSendSMSdate,getdate()) < 72"	' 문자발송된지 3일 지난건은 보내지 않는다
    'sqlStr = sqlStr + " 	and d.isupchebeasong='N'"
    sqlStr = sqlStr + " 	and l.isautocanceldate is null"		' 자동취소안된거
    sqlStr = sqlStr + " 	and m.orderserial in ('" & orderserial & "') "

    'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr
end function

function RegWebCSDetailReturn(CsId, orderserial, detailidx, regitemno, gubun01, gubun02)
    dim sqlStr, i
    dim detailidxArr, regitemnoArr

    detailidxArr = split(detailidx, ",")
    regitemnoArr = split(regitemno, ",")

    for i = 0 to UBound(detailidxArr)
		if (TRIM(detailidxArr(i)) <> "") and (TRIM(regitemnoArr(i))<>"") then
	        call AddOneCSDetail(CsId, detailidxArr(i), gubun01, gubun02, orderserial, regitemnoArr(i))
		end if
	next
	sqlStr = " update [db_cs].[dbo].tbl_new_as_detail"
	sqlStr = sqlStr + " set itemid=T.itemid"
	sqlStr = sqlStr + " , itemoption=T.itemoption"
	sqlStr = sqlStr + " , makerid=T.makerid"
	sqlStr = sqlStr + " , itemname=T.itemname"
	sqlStr = sqlStr + " , itemoptionname=T.itemoptionname"
	sqlStr = sqlStr + " , itemcost=T.itemcost"
	sqlStr = sqlStr + " , orderitemno=T.itemno"
	sqlStr = sqlStr + " , isupchebeasong=T.isupchebeasong"
	sqlStr = sqlStr + " , regdetailstate=T.currstate"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail T"
	sqlStr = sqlStr + " where T.orderserial='" + orderserial + "'"
	sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_detail.masterid=" + CStr(CsId)
	sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_detail.orderdetailidx=T.idx"

	dbget.Execute sqlStr
end function


function GetWebCSDetailReturnBeasongPay(orderserial, ReturnMakerid)
	dim sqlStr

    sqlStr = " select d.idx as detailidx "
    sqlStr = sqlStr + " from [db_order].dbo.tbl_order_detail d " + VbCrlf
    sqlStr = sqlStr + " where d.orderserial='" + orderserial + "'" + VbCrlf
    sqlStr = sqlStr + " and d.makerid='" + ReturnMakerid + "'" + VbCrlf
    sqlStr = sqlStr + " and d.itemid=0 " + VbCrlf
	sqlStr = sqlStr + " and d.cancelyn <> 'Y' " + VbCrlf
	'response.write sqlStr

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        GetWebCSDetailReturnBeasongPay = rsget("detailidx")
    else
    	GetWebCSDetailReturnBeasongPay = 0
    end if
    rsget.close
end function


function AddOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno)
    dim sqlStr

    sqlStr = " insert into [db_cs].[dbo].tbl_new_as_detail"
    sqlStr = sqlStr + " (masterid, orderdetailidx, gubun01,gubun02"
    sqlStr = sqlStr + " ,orderserial, itemid, itemoption, makerid, itemname, itemoptionname, regitemno, confirmitemno,orderitemno) "
    sqlStr = sqlStr + " values(" + CStr(id) + ""
    sqlStr = sqlStr + " ," + CStr(dorderdetailidx) + ""
    sqlStr = sqlStr + " ,'" + CStr(dgubun01) + "'"
    sqlStr = sqlStr + " ,'" + CStr(dgubun02) + "'"
    sqlStr = sqlStr + " ,'" + CStr(orderserial) + "'"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " )"

    dbget.Execute sqlStr
end function

function AddCSDetailByArrStr(byval detailitemlist, id, orderserial)
    dim sqlStr, tmp, buf, i
    dim dorderdetailidx, dgubun01, dgubun02, dregitemno

    buf = split(detailitemlist, "|")

    for i = 0 to UBound(buf)
		if (TRIM(buf(i)) <> "") then
			tmp = split(buf(i), Chr(9))

			dorderdetailidx = tmp(0)
			dgubun01        = tmp(1)
			dgubun02        = tmp(2)
			dregitemno      = tmp(3)

	        call AddOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno)
		end if
	next
	sqlStr = " update [db_cs].[dbo].tbl_new_as_detail"
	sqlStr = sqlStr + " set itemid=T.itemid"
	sqlStr = sqlStr + " , itemoption=T.itemoption"
	sqlStr = sqlStr + " , makerid=T.makerid"
	sqlStr = sqlStr + " , itemname=T.itemname"
	sqlStr = sqlStr + " , itemoptionname=T.itemoptionname"
	sqlStr = sqlStr + " , itemcost=T.itemcost"
	sqlStr = sqlStr + " , buycash=T.buycash"
	sqlStr = sqlStr + " , orderitemno=(CASE WHEN T.cancelyn='Y' THEN 0 ELSE T.itemno END)"
	sqlStr = sqlStr + " , isupchebeasong=T.isupchebeasong"
	sqlStr = sqlStr + " , regdetailstate=T.currstate"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail T"
	sqlStr = sqlStr + " where T.orderserial='" + orderserial + "'"
	sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_detail.masterid=" + CStr(id)
	sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_detail.orderdetailidx=T.idx"

	dbget.Execute sqlStr

end function

function RegWebCancelRefundInfo(CsId, orderserial, returnmethod, refundrequire , rebankname, rebankaccount, rebankownername)
    dim sqlStr
    ''전체 취소 환불정보

    sqlStr = " insert into [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " (asid"
    sqlStr = sqlStr + " ,returnmethod"
    sqlStr = sqlStr + " ,refundrequire"
    sqlStr = sqlStr + " ,orgsubtotalprice"
    sqlStr = sqlStr + " ,orgitemcostsum"
    sqlStr = sqlStr + " ,orgbeasongpay"
    sqlStr = sqlStr + " ,orgmileagesum"
    sqlStr = sqlStr + " ,orgcouponsum"
    sqlStr = sqlStr + " ,orgallatdiscountsum"

    ''취소 관련정보
    sqlStr = sqlStr + " ,canceltotal"
    sqlStr = sqlStr + " ,refunditemcostsum"
    sqlStr = sqlStr + " ,refundmileagesum"
    sqlStr = sqlStr + " ,refundcouponsum"
    sqlStr = sqlStr + " ,allatsubtractsum"
    sqlStr = sqlStr + " ,refundbeasongpay"
    sqlStr = sqlStr + " ,refunddeliverypay"
    sqlStr = sqlStr + " ,refundadjustpay"
    sqlStr = sqlStr + " ,rebankname"
    sqlStr = sqlStr + " ,rebankaccount"
    sqlStr = sqlStr + " ,rebankownername"
    sqlStr = sqlStr + " ,paygateTid"
    sqlStr = sqlStr + " ,orggiftcardsum"
    sqlStr = sqlStr + " ,orgdepositsum"
    sqlStr = sqlStr + " ,refundgiftcardsum"
    sqlStr = sqlStr + " ,refunddepositsum"
    sqlStr = sqlStr + " )"

    sqlStr = sqlStr + " select " + CStr(CsId)
    sqlStr = sqlStr + " ,'" + returnmethod + "'"
    sqlStr = sqlStr + " ," + CStr(refundrequire)
    sqlStr = sqlStr + " ,m.subtotalprice"
    sqlStr = sqlStr + " ,m.totalsum-IsNULL(d.itemcost,0)"
    sqlStr = sqlStr + " ,IsNULL(d.itemcost,0)"
    sqlStr = sqlStr + " ,m.miletotalprice"
    sqlStr = sqlStr + " ,m.tencardspend"
    sqlStr = sqlStr + " ,m.allatdiscountprice"

    sqlStr = sqlStr + " ,m.subtotalprice"
    sqlStr = sqlStr + " ,m.totalsum-IsNULL(d.itemcost,0)"
    sqlStr = sqlStr + " ,m.miletotalprice*-1"
    sqlStr = sqlStr + " ,m.tencardspend*-1"
    sqlStr = sqlStr + " ,m.allatdiscountprice*-1"
    sqlStr = sqlStr + " ,IsNULL(d.itemcost,0)"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,'" + rebankname + "'"
    sqlStr = sqlStr + " ,'" + rebankaccount + "'"
    sqlStr = sqlStr + " ,'" + rebankownername + "'"
    sqlStr = sqlStr + " ,m.paygatetid "

    sqlStr = sqlStr + " ,IsNull(p900.realPayedSum, 0) "
    sqlStr = sqlStr + " ,IsNull(p200.realPayedSum, 0) "
    sqlStr = sqlStr + " ,IsNull(p900.realPayedSum, 0) * -1 "
    sqlStr = sqlStr + " ,IsNull(p200.realPayedSum, 0) * -1 "

    sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
    sqlStr = sqlStr + "     left join (select orderserial, sum(itemcost) as itemcost from [db_order].[dbo].tbl_order_detail where orderserial='" + orderserial + "' and itemid=0 and cancelyn<>'Y' group by orderserial) d"
    sqlStr = sqlStr + "     on d.orderserial='" + orderserial + "' and m.orderserial=d.orderserial "

    sqlStr = sqlStr + " left join db_order.dbo.tbl_order_PaymentEtc p200 "						'예치금
    sqlStr = sqlStr + " on "
    sqlStr = sqlStr + " 	m.orderserial = p200.orderserial and p200.acctdiv = '200' "
    sqlStr = sqlStr + " left join db_order.dbo.tbl_order_PaymentEtc p900 "						'상품권
    sqlStr = sqlStr + " on "
    sqlStr = sqlStr + " 	m.orderserial = p900.orderserial and p900.acctdiv = '900' "

    sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"
''rw sqlStr

    dbget.Execute sqlStr

end function

function RegWebGiftCardCancelRefundInfo(CsId, orderserial, returnmethod, refundrequire , rebankname, rebankaccount, rebankownername, paygatetid)
    dim sqlStr
    ''전체 취소 환불정보

    sqlStr = " insert into [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " (asid"
    sqlStr = sqlStr + " ,returnmethod"
    sqlStr = sqlStr + " ,refundrequire"
    sqlStr = sqlStr + " ,orgsubtotalprice"
    sqlStr = sqlStr + " ,orgitemcostsum"
    sqlStr = sqlStr + " ,orgbeasongpay"
    sqlStr = sqlStr + " ,orgmileagesum"
    sqlStr = sqlStr + " ,orgcouponsum"
    sqlStr = sqlStr + " ,orgallatdiscountsum"

    ''취소 관련정보
    sqlStr = sqlStr + " ,canceltotal"
    sqlStr = sqlStr + " ,refunditemcostsum"
    sqlStr = sqlStr + " ,refundmileagesum"
    sqlStr = sqlStr + " ,refundcouponsum"
    sqlStr = sqlStr + " ,allatsubtractsum"
    sqlStr = sqlStr + " ,refundbeasongpay"
    sqlStr = sqlStr + " ,refunddeliverypay"
    sqlStr = sqlStr + " ,refundadjustpay"
    sqlStr = sqlStr + " ,rebankname"
    sqlStr = sqlStr + " ,rebankaccount"
    sqlStr = sqlStr + " ,rebankownername"
    sqlStr = sqlStr + " ,paygateTid"

    sqlStr = sqlStr + " ,orggiftcardsum"
    sqlStr = sqlStr + " ,orgdepositsum"
    sqlStr = sqlStr + " ,refundgiftcardsum"
    sqlStr = sqlStr + " ,refunddepositsum"
    sqlStr = sqlStr + " )"

    sqlStr = sqlStr + " values( " + CStr(CsId)
    sqlStr = sqlStr + " ,'" + returnmethod + "'"
    sqlStr = sqlStr + " ," + CStr(refundrequire)
    sqlStr = sqlStr + " ," + CStr(refundrequire)
    sqlStr = sqlStr + " ," + CStr(refundrequire)
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,0"

    sqlStr = sqlStr + " ," + CStr(refundrequire)
    sqlStr = sqlStr + " ," + CStr(refundrequire)
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,'" + rebankname + "'"
    sqlStr = sqlStr + " ,'" + rebankaccount + "'"
    sqlStr = sqlStr + " ,'" + rebankownername + "'"
    sqlStr = sqlStr + " ,'" + CStr(paygatetid) + "'"

    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " )"
''rw sqlStr

    dbget.Execute sqlStr

end function

function CopyWebCancelRefundInfo(FromCsId, ToCsId)
    dim sqlStr
    ''전체 취소 환불정보 복사

    sqlStr = " insert into [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " (asid"
    sqlStr = sqlStr + " 	, returnmethod"
    sqlStr = sqlStr + " 	, refundrequire"
    sqlStr = sqlStr + " 	, orgsubtotalprice"
    sqlStr = sqlStr + " 	, orgitemcostsum"
    sqlStr = sqlStr + " 	, orgbeasongpay"
    sqlStr = sqlStr + " 	, orgmileagesum"
    sqlStr = sqlStr + " 	, orgcouponsum"
    sqlStr = sqlStr + " 	, orgallatdiscountsum"

    ''취소 관련정보
    sqlStr = sqlStr + " 	, canceltotal"
    sqlStr = sqlStr + " 	, refunditemcostsum"
    sqlStr = sqlStr + " 	, refundmileagesum"
    sqlStr = sqlStr + " 	, refundcouponsum"
    sqlStr = sqlStr + " 	, allatsubtractsum"
    sqlStr = sqlStr + " 	, refundbeasongpay"
    sqlStr = sqlStr + " 	, refunddeliverypay"
    sqlStr = sqlStr + " 	, refundadjustpay"
    sqlStr = sqlStr + " 	, rebankname"
    sqlStr = sqlStr + " 	, rebankaccount"
    sqlStr = sqlStr + " 	, rebankownername"
    sqlStr = sqlStr + " 	, encmethod"
    sqlStr = sqlStr + " 	, encaccount"

    sqlStr = sqlStr + " 	, paygateTid"
    sqlStr = sqlStr + " 	, orggiftcardsum"
    sqlStr = sqlStr + " 	, orgdepositsum"
    sqlStr = sqlStr + " 	, refundgiftcardsum"
    sqlStr = sqlStr + " 	, refunddepositsum"
    sqlStr = sqlStr + " )"

    sqlStr = sqlStr + " select " + CStr(ToCsId)
    sqlStr = sqlStr + " 	, returnmethod"
    sqlStr = sqlStr + " 	, refundrequire"
    sqlStr = sqlStr + " 	, orgsubtotalprice"
    sqlStr = sqlStr + " 	, orgitemcostsum"
    sqlStr = sqlStr + " 	, orgbeasongpay"
    sqlStr = sqlStr + " 	, orgmileagesum"
    sqlStr = sqlStr + " 	, orgcouponsum"
    sqlStr = sqlStr + " 	, orgallatdiscountsum"

    ''취소 관련정보
    sqlStr = sqlStr + " 	, canceltotal"
    sqlStr = sqlStr + " 	, refunditemcostsum"
    sqlStr = sqlStr + " 	, refundmileagesum"
    sqlStr = sqlStr + " 	, refundcouponsum"
    sqlStr = sqlStr + " 	, allatsubtractsum"
    sqlStr = sqlStr + " 	, refundbeasongpay"
    sqlStr = sqlStr + " 	, refunddeliverypay"
    sqlStr = sqlStr + " 	, refundadjustpay"
    sqlStr = sqlStr + " 	, rebankname"
    sqlStr = sqlStr + " 	, rebankaccount"
    sqlStr = sqlStr + " 	, rebankownername"
    sqlStr = sqlStr + " 	, encmethod"
    sqlStr = sqlStr + " 	, encaccount"

    sqlStr = sqlStr + " 	, paygateTid"
    sqlStr = sqlStr + " 	, orggiftcardsum"
    sqlStr = sqlStr + " 	, orgdepositsum"
    sqlStr = sqlStr + " 	, refundgiftcardsum"
    sqlStr = sqlStr + " 	, refunddepositsum"
    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_as_refund_info "
    sqlStr = sqlStr + " where asid = " & FromCsId & " "

    dbget.Execute sqlStr


    '관련 CS''''''''''''''''''''*******************************
    sqlStr = " update [db_cs].[dbo].tbl_new_as_list "
    sqlStr = sqlStr + " set refasid = " & CStr(FromCsId) & " "
    sqlStr = sqlStr + " where id = " & CStr(ToCsId) & " "
    dbget.Execute sqlStr

end function

function UpdateWebRefundInfo(id, orderserial, returnmethod, rebankname, rebankaccount, rebankownername)
    dim sqlStr, AssignedRows
    dim opentitle
    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info" + VbCrlf
    sqlStr = sqlStr + " set returnmethod='"&returnmethod&"'" + VbCrlf
    sqlStr = sqlStr + " ,rebankname='"&rebankname&"'" + VbCrlf
    sqlStr = sqlStr + " ,rebankaccount='"&rebankaccount&"'" + VbCrlf
    sqlStr = sqlStr + " ,rebankownername='"&rebankownername&"'" + VbCrlf
    sqlStr = sqlStr + " where asid=" & id

    dbget.Execute sqlStr, AssignedRows

    UpdateWebRefundInfo = (AssignedRows=1)

    ''opentitle 저장 : 변경되었을 수 있음.
    if (returnmethod="R007") then
        opentitle = "주문 취소 후 무통장 환불 요청 접수"
    elseif (returnmethod="R900") then
        opentitle = "주문 취소 후 마일리지 환불 요청 접수"
    elseif (returnmethod="R910") then
        opentitle = "주문 취소 후 예치금전환 요청 접수"
    end if

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"        + VbCrlf
    sqlStr = sqlStr + " set opentitle='" + opentitle + "'"  + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)

    dbget.Execute sqlStr
end function

function RegWebRefundInfo(CsId, orderserial, returnmethod, refundrequire , rebankname, rebankaccount, rebankownername,  canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum, refundbeasongpay, refunddeliverypay)
    dim sqlStr
    ''전체 취소 환불정보

    sqlStr = " insert into [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " (asid"
    sqlStr = sqlStr + " ,returnmethod"
    sqlStr = sqlStr + " ,refundrequire"
    sqlStr = sqlStr + " ,orgsubtotalprice"
    sqlStr = sqlStr + " ,orgitemcostsum"
    sqlStr = sqlStr + " ,orgbeasongpay"
    sqlStr = sqlStr + " ,orgmileagesum"
    sqlStr = sqlStr + " ,orgcouponsum"
    sqlStr = sqlStr + " ,orgallatdiscountsum"

    ''취소 관련정보
    sqlStr = sqlStr + " ,canceltotal"
    sqlStr = sqlStr + " ,refunditemcostsum"
    sqlStr = sqlStr + " ,refundmileagesum"
    sqlStr = sqlStr + " ,refundcouponsum"
    sqlStr = sqlStr + " ,allatsubtractsum"
    sqlStr = sqlStr + " ,refundbeasongpay"
    sqlStr = sqlStr + " ,refunddeliverypay"
    sqlStr = sqlStr + " ,refundadjustpay"
    sqlStr = sqlStr + " ,rebankname"
    sqlStr = sqlStr + " ,rebankaccount"
    sqlStr = sqlStr + " ,rebankownername"
    sqlStr = sqlStr + " ,paygateTid"
    sqlStr = sqlStr + " )"

    sqlStr = sqlStr + " select " + CStr(CsId)
    sqlStr = sqlStr + " ,'" + returnmethod + "'"
    sqlStr = sqlStr + " ," + CStr(refundrequire)
    sqlStr = sqlStr + " ,(m.subtotalprice - IsNull(m.sumPaymentEtc, 0))"
    sqlStr = sqlStr + " ,m.totalsum-IsNULL(d.itemcost,0)"
    sqlStr = sqlStr + " ,IsNULL(d.itemcost,0)"
    sqlStr = sqlStr + " ,m.miletotalprice"
    sqlStr = sqlStr + " ,m.tencardspend"
    sqlStr = sqlStr + " ,m.allatdiscountprice"

    sqlStr = sqlStr + " ," + CStr(canceltotal)
    sqlStr = sqlStr + " ," + CStr(refunditemcostsum)
    sqlStr = sqlStr + " ," + CStr(refundmileagesum*-1)
    sqlStr = sqlStr + " ," + CStr(refundcouponsum*-1)
    sqlStr = sqlStr + " ," + CStr(allatsubtractsum*-1)
    sqlStr = sqlStr + " ," + CStr(refundbeasongpay)
    sqlStr = sqlStr + " ," + CStr(refunddeliverypay*-1)
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,'" + rebankname + "'"
    sqlStr = sqlStr + " ,'" + rebankaccount + "'"
    sqlStr = sqlStr + " ,'" + rebankownername + "'"
    sqlStr = sqlStr + " ,m.paygatetid "
    sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
    sqlStr = sqlStr + "     left join (select orderserial, sum(itemcost) as itemcost from [db_order].[dbo].tbl_order_detail where orderserial='" + orderserial + "' and itemid=0 and cancelyn<>'Y' group by orderserial) d"
    sqlStr = sqlStr + "     on d.orderserial='" + orderserial + "' and m.orderserial=d.orderserial "
    ''sqlStr = sqlStr + "     left join [db_order].[dbo].tbl_order_detail d"
    ''sqlStr = sqlStr + "     on d.orderserial='" + orderserial + "' and m.orderserial=d.orderserial and d.itemid=0 and d.cancelyn<>'Y'"
    sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"

    dbget.Execute sqlStr

end function

function CancelProcess(id, orderserial, isForceAllcancel)
    dim IsAllCancel, IsUpdatedMile, IsUpdatedDeposit, IsUpdatedGiftCard

    dim sqlStr, userid, ipkumdiv, miletotalprice, tencardspend, allatdiscountprice

    dim refundmileagesum, refundcouponsum, allatsubtractsum
    dim refundbeasongpay, refunditemcostsum, refunddeliverypay
    dim refundadjustpay, canceltotal

    dim detailidx, orgbeasongpay, deliveritemoption, deliverbeasongpay
    dim InsureCd
    dim openMessage

    dim regDetailRows, i
    dim remaintencardspend, gubun01, gubun02

    dim orggiftcardsum, refundgiftcardsum, orgdepositsum, refunddepositsum

    if (isForceAllcancel) then
        IsAllCancel = true
    else
        IsAllCancel = IsAllCancelRegValid(id, orderserial)
    end if

    sqlStr = " select userid, ipkumdiv, IsNULL(miletotalprice,0) as miletotalprice "
    sqlStr = sqlStr + " ,IsNULL(tencardspend,0) as tencardspend, IsNULL(allatdiscountprice,0) as allatdiscountprice" + VbCrlf
    sqlStr = sqlStr + " ,IsNULL(InsureCd,'') as InsureCd" + VbCrlf
    sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master" + VbCrlf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        userid              = rsget("userid")
        miletotalprice      = rsget("miletotalprice")
        tencardspend        = rsget("tencardspend")
        allatdiscountprice  = rsget("allatdiscountprice")
        InsureCd            = rsget("InsureCd")
        ipkumdiv            = rsget("ipkumdiv")
    end if
    rsget.close

    sqlStr = " select acctdiv, IsNull(realPayedsum, 0) as realPayedsum " + VbCrlf
    sqlStr = sqlStr + " from " + VbCrlf
    sqlStr = sqlStr + " db_order.dbo.tbl_order_PaymentEtc " + VbCrlf
    sqlStr = sqlStr + " where " + VbCrlf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrlf
    sqlStr = sqlStr + " 	and orderserial = '" + orderserial + "' " + VbCrlf
    sqlStr = sqlStr + " 	and acctdiv in ('200', '900') " + VbCrlf			'200 : 예치금, 900 : 상품권

    rsget.Open sqlStr,dbget,1

	orgdepositsum = 0
	orggiftcardsum = 0
	do until rsget.eof
		if (CStr(rsget("acctdiv")) = "200") then
			orgdepositsum = rsget("realPayedsum")
		elseif (CStr(rsget("acctdiv")) = "900") then
			orggiftcardsum = rsget("realPayedsum")
		end if

		rsget.movenext
	loop
	rsget.close

    sqlStr = " select r.*, a.gubun01, a.gubun02 from [db_cs].[dbo].tbl_new_as_list a"
    sqlStr = sqlStr + " , [db_cs].[dbo].tbl_as_refund_info r"
    sqlStr = sqlStr + " where a.id=" + CStr(id)
    sqlStr = sqlStr + " and a.id=r.asid"
    sqlStr = sqlStr + " and a.deleteyn='N'"
    sqlStr = sqlStr + " and a.currstate<>'B007'"


    rsget.Open sqlStr,dbget,1

    if Not rsget.Eof then
        refundmileagesum    = rsget("refundmileagesum")
        refundcouponsum     = rsget("refundcouponsum")

        refundgiftcardsum   = rsget("refundgiftcardsum")
        refunddepositsum    = rsget("refunddepositsum")

        allatsubtractsum    = rsget("allatsubtractsum")

        refunditemcostsum   = rsget("refunditemcostsum")

        refundbeasongpay    = rsget("refundbeasongpay")
        refunddeliverypay   = rsget("refunddeliverypay")
        refundadjustpay     = rsget("refundadjustpay")
        canceltotal         = rsget("canceltotal")
        gubun01             = rsget("gubun01")
        gubun02             = rsget("gubun02")

    else
        refundmileagesum    = 0
        refundcouponsum     = 0
        allatsubtractsum    = 0

        refundgiftcardsum   = 0
        refunddepositsum    = 0

        refunditemcostsum   = 0

        refundbeasongpay    = 0
        refunddeliverypay   = 0
        refundadjustpay     = 0
        canceltotal         = 0
    end if
    rsget.close

'' 마일리지 환원

    IsUpdatedMile = false
    if (userid<>"") and (IsAllCancel) and (miletotalprice<>0) then
        '' 전체 취소인경우 주문건 취소로 jukyocd : 2 상품구매, 3 : 부분취소시 환원마일리지
        sqlStr = " update [db_user].[dbo].tbl_mileagelog " + VbCrlf
        sqlStr = sqlStr + " set deleteyn='Y' " + VbCrlf
        sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
        sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and jukyocd in ('2','3')" + VbCrlf
        dbget.Execute sqlStr

        IsUpdatedMile = true

        if openMessage="" then
            openMessage = openMessage + "사용 마일리지 환원 : " & miletotalprice
        else
            openMessage = openMessage + VbCrlf + "사용 마일리지 환원 : " & miletotalprice
        end if

    end if

    if (userid<>"") and (Not IsAllCancel) and (refundmileagesum<>0) then
        '' 부분 취소인데 마일리지 환원할 경우.
        sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
        sqlStr = sqlStr + " set miletotalprice=miletotalprice + " + CStr(refundmileagesum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
        dbget.Execute sqlStr


        sqlStr = " insert into [db_user].[dbo].tbl_mileagelog " + VbCrlf
        sqlStr = sqlStr + " (userid, mileage, jukyocd, jukyo, orderserial, deleteyn) " + VbCrlf
        sqlStr = sqlStr + " values ("
        sqlStr = sqlStr + " '" + userid + "'"
        sqlStr = sqlStr + " ," + CStr(refundmileagesum*-1) + ""
        sqlStr = sqlStr + " ,'3'"
        sqlStr = sqlStr + " ,'상품구매 취소 환원'"
        sqlStr = sqlStr + " ,'" + orderserial + "'"
        sqlStr = sqlStr + " ,'N'"
        sqlStr = sqlStr + " )"
        dbget.Execute sqlStr

        IsUpdatedMile = true

        if openMessage="" then
            openMessage = openMessage + "사용 마일리지 환원 : " & refundmileagesum
        else
            openMessage = openMessage + VbCrlf + "사용 마일리지 환원 : " & refundmileagesum
        end if
    end if

'TODO : 상품권환원

'예치금환원
	IsUpdatedDeposit = false
    if (userid<>"") and (IsAllCancel) and (orgdepositsum <> 0) then
        '' 전체 취소인경우 주문건 취소로 jukyocd : 100 상품구매, 10 : 부분취소시 예치금 환원
        sqlStr = " update [db_user].[dbo].tbl_depositlog " + VbCrlf
        sqlStr = sqlStr + " set deleteyn='Y' " + VbCrlf
        sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
        sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and jukyocd in ('100','10')" + VbCrlf					'100 : 상품구매사용 / 10 : 일부환원 (참고 : db_user.dbo.tbl_deposit_gubun)
        dbget.Execute sqlStr

        IsUpdatedDeposit = true

        if openMessage="" then
            openMessage = openMessage + "사용 예치금 환원 : " & orgdepositsum
        else
            openMessage = openMessage + VbCrlf + "사용 예치금 환원 : " & orgdepositsum
        end if

    end if


    if (userid<>"") and (Not IsAllCancel) and (refunddepositsum <> 0) then
        '' 부분 취소인데 예치금 환원할 경우.

        sqlStr = " update [db_order].[dbo].tbl_order_PaymentEtc" + VbCrlf
        sqlStr = sqlStr + " set realPayedsum=realPayedsum + " + CStr(refunddepositsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and acctdiv='200'" + VbCrlf
        dbget.Execute sqlStr

        sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
        sqlStr = sqlStr + " set sumPaymentETC=sumPaymentETC + " + CStr(refunddepositsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
        dbget.Execute sqlStr

        sqlStr = " insert into [db_user].[dbo].tbl_depositlog " + VbCrlf
        sqlStr = sqlStr + " (userid, deposit, jukyocd, jukyo, orderserial, deleteyn) " + VbCrlf
        sqlStr = sqlStr + " values ("
        sqlStr = sqlStr + " '" + userid + "'"
        sqlStr = sqlStr + " ," + CStr(refunddepositsum*-1) + ""
        sqlStr = sqlStr + " ,'10'"
        sqlStr = sqlStr + " ,'상품구매 취소 환원'"
        sqlStr = sqlStr + " ,'" + orderserial + "'"
        sqlStr = sqlStr + " ,'N'"
        sqlStr = sqlStr + " )"
        dbget.Execute sqlStr

        IsUpdatedDeposit = true

        if openMessage="" then
            openMessage = openMessage + "사용 예치금 환원 : " & refunddepositsum
        else
            openMessage = openMessage + VbCrlf + "사용 예치금 환원 : " & refunddepositsum
        end if
    end if

'Gift카드환원
    IsUpdatedGiftCard = false
    if (userid<>"") and (IsAllCancel) and (orggiftcardsum <> 0) then
        '' 전체 취소인경우 주문건 취소로 jukyocd : 200 상품구매, 300 : 부분취소시 Gift카드 환원
        sqlStr = " update [db_user].[dbo].tbl_giftcard_log " + VbCrlf
        sqlStr = sqlStr + " set deleteyn='Y' " + VbCrlf
        sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
        sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and jukyocd in ('200','300')" + VbCrlf					'200 : 상품구매사용 / 300 : 일부환원 (참고 : db_user.dbo.tbl_giftcard_gubun)
        dbget.Execute sqlStr

        IsUpdatedGiftCard = true

        if openMessage="" then
            openMessage = openMessage + "사용 Gift카드 환원 : " & orggiftcardsum
        else
            openMessage = openMessage + VbCrlf + "사용 Gift카드 환원 : " & orggiftcardsum
        end if

    end if

    if (userid<>"") and (Not IsAllCancel) and (refundgiftcardsum <> 0) then
        '' 부분 취소인데 Gift카드 환원할 경우.

        sqlStr = " update [db_order].[dbo].tbl_order_PaymentEtc" + VbCrlf
        sqlStr = sqlStr + " set realPayedsum=realPayedsum + " + CStr(refundgiftcardsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and acctdiv='900'" + VbCrlf
        dbget.Execute sqlStr

        sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
        sqlStr = sqlStr + " set sumPaymentETC=sumPaymentETC + " + CStr(refundgiftcardsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
        dbget.Execute sqlStr

        sqlStr = " insert into [db_user].[dbo].tbl_giftcard_log " + VbCrlf
        sqlStr = sqlStr + " (userid, useCash, jukyocd, jukyo, orderserial, deleteyn, reguserid) " + VbCrlf
        sqlStr = sqlStr + " values ("
        sqlStr = sqlStr + " '" + userid + "'"
        sqlStr = sqlStr + " ," + CStr(refundgiftcardsum*-1) + ""
        sqlStr = sqlStr + " ,'300'"
        sqlStr = sqlStr + " ,'상품구매 취소 환원'"
        sqlStr = sqlStr + " ,'" + orderserial + "'"
        sqlStr = sqlStr + " ,'N'"
        sqlStr = sqlStr + " ,'" + userid + "'"
        sqlStr = sqlStr + " )"
        dbget.Execute sqlStr

        IsUpdatedGiftCard = true

        if openMessage="" then
            openMessage = openMessage + "사용 Gift카드 환원 : " & refundgiftcardsum
        else
            openMessage = openMessage + VbCrlf + "사용 Gift카드 환원 : " & refundgiftcardsum
        end if
    end if


'' 할인권 환급
    if (IsAllCancel) and (tencardspend<>0) then
        sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
	    sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
	    sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "
        sqlStr = sqlStr + " and userid='"&userid&"'"  ''2015/04/13 추가(느리므로)

	    dbget.Execute sqlStr

	    if openMessage="" then
            openMessage = openMessage + "사용 보너스쿠폰 환급"
        else
            openMessage = openMessage + VbCrlf + "사용 보너스쿠폰 환급"
        end if
    end if

    if (Not IsAllCancel) and (refundcouponsum<>0) then
         '' 부분 취소인경우 - 환급한 만큼 깜..
        sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
        sqlStr = sqlStr + " set tencardspend=tencardspend + " + CStr(refundcouponsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        dbget.Execute sqlStr

        ''전체 환급인 경우만 쿠폰을 돌려줌
        sqlStr = "select IsNULL(tencardspend,0) as tencardspend from [db_order].[dbo].tbl_order_master" + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        rsget.Open sqlStr,dbget,1
            remaintencardspend = rsget("tencardspend")
        rsget.close

        ''원래 할인권 사용액이 있고, 남은 쿠폰사용액이 없을경우 전체  환급
        if (tencardspend>0) then
            if (remaintencardspend=0)   then
                sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
            	sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
            	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "

            	dbget.Execute sqlStr

            	if openMessage="" then
                    openMessage = openMessage + "사용 할인권  환급"
                else
                    openMessage = openMessage + VbCrlf + "사용 할인권  환급"
                end if
            else
                ''(또는, %쿠폰인 경우 공통,단순변심인 경우 제외하고 환급해줌./ 부분취소 ) C004 CD01
                if (ipkumdiv>3) and (Not ((gubun01="C004") and (gubun02="CD01"))) then
                    sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
                	sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
                	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "
                	sqlStr = sqlStr + " and coupontype=1"

                	dbget.Execute sqlStr

                	if openMessage="" then
                        openMessage = openMessage + "사용 할인권  환급."
                    else
                        openMessage = openMessage + VbCrlf + "사용 할인권  환급."
                    end if
                end if
            end if
        end if

    end if



    '' 올엣카드 할인 차감
    if (IsAllCancel) and (allatdiscountprice<>0) then

    end if

    if (Not IsAllCancel) and (allatsubtractsum<>0) then
        sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
        sqlStr = sqlStr + " set allatdiscountprice=allatdiscountprice + " + CStr(allatsubtractsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        dbget.Execute sqlStr

        if openMessage="" then
            openMessage = openMessage + "올엣카드 할인 차감 : " & allatsubtractsum
        else
            openMessage = openMessage + VbCrlf + "올엣카드 할인 차감 : " & allatsubtractsum
        end if
    end if


	'배송비도 같이 취소된다. setCancelDetail()

    if (IsAllCancel) then
	    ''전체 취소인경우
	    '' 주문  master 취소 변경
	    call setCancelMaster(id, orderserial)

	    if openMessage="" then
            openMessage = openMessage + "주문취소 완료"
        else
            openMessage = openMessage + VbCrlf + "주문취소 완료"
        end if
	else
	    ''부분 취소인경우
	    '' 주문  detail 취소 변경
	    call setCancelDetail(id, orderserial)

	    call reCalcuOrderMaster(orderserial)

	    if openMessage="" then
            openMessage = openMessage + "주문부분취소 완료"
        else
            openMessage = openMessage + VbCrlf + "주문부분취소 완료"
        end if
	end if

    ''마일리지는 주문건 취소 후 재계산해야함.
    '예치금 재계산
    if (userid<>"") then
        Call updateUserMileage(userid)

        if IsUpdatedDeposit then
        	Call updateUserDeposit(userid)
        end if

        if IsUpdatedGiftCard then
        	Call updateUserGiftCard(userid)
        end if
    end if

    ''최근 주문수량 조정 2015/08/12
    if (userid<>"") and (IsAllCancel) then
        sqlStr = "exec [db_order].[dbo].sp_Ten_Recalcu_His_recent_OrderCNT '" & userid & "'"
        dbget.Execute(sqlStr)
    end if

    ''전자보증서 발급된 경우 취소
    if (InsureCd="0") then
        Call UsafeCancel(orderserial)
    end if

    ''재고 및 한정수량 조절(2007-09-01 서동석 추가)
    ''Call LimitItemRecover(orderserial) : 기존
    if (IsAllCancel) then
	    ''전체 취소인경우 // setCancelMaster 에 통합됨
	    ''sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_cancelOrderAll '" & orderserial & "'"
	    ''dbget.Execute sqlStr
	else
	    ''부분 취소인경우
	    sqlStr = " select itemid,itemoption,regitemno "
        sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_detail "
        sqlStr = sqlStr & " where masterid=" & id
        sqlStr = sqlStr & " and orderserial='" & orderserial & "'"

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            regDetailRows = rsget.getRows()
        end if
        rsget.Close

        if IsArray(regDetailRows) then
            for i=0 to UBound(regDetailRows,2)
    	        sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & regDetailRows(0,i) & ",'" & regDetailRows(1,i) & "'," & regDetailRows(2,i)
                dbget.Execute sqlStr
            Next
        end if
	end if

    ''전자보증서 발급된 경우 취소
    if (InsureCd="0") then
        Call UsafeCancel(orderserial)
    end if

    if (openMessage<>"") then
        call AddCustomerOpenContents(id, openMessage)
    end if
end function


''취소(전체취소 /부분취소 추가 2009 : 기 접수된 내역을 완료처리 할 경우만.)
function CancelProcess1111111111111111(id, orderserial, isForceAllcancel)
    dim IsAllCancel, IsUpdatedMile, IsUpdatedDeposit

    dim sqlStr, userid, ipkumdiv, miletotalprice, tencardspend, allatdiscountprice

    dim refundmileagesum, refundcouponsum, allatsubtractsum
    dim refundbeasongpay, refunditemcostsum, refunddeliverypay
    dim refundadjustpay, canceltotal

    dim detailidx, orgbeasongpay, deliveritemoption, deliverbeasongpay
    dim InsureCd
    dim openMessage

    dim regDetailRows, i
    dim remaintencardspend, gubun01, gubun02

    dim sumPaymentEtc
    dim orggiftcardsum, refundgiftcardsum, orgdepositsum, refunddepositsum

    if (isForceAllcancel) then
        IsAllCancel = true
    else
        IsAllCancel = IsAllCancelRegValid(id, orderserial)
    end if

    sqlStr = " select userid, ipkumdiv, IsNULL(miletotalprice,0) as miletotalprice "
    sqlStr = sqlStr + " ,IsNULL(tencardspend,0) as tencardspend, IsNULL(allatdiscountprice,0) as allatdiscountprice" + VbCrlf
    sqlStr = sqlStr + " ,IsNULL(InsureCd,'') as InsureCd" + VbCrlf
    sqlStr = sqlStr + " ,IsNULL(sumPaymentEtc,0) as sumPaymentEtc" + VbCrlf
    sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master" + VbCrlf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        userid              = rsget("userid")
        miletotalprice      = rsget("miletotalprice")
        tencardspend        = rsget("tencardspend")
        allatdiscountprice  = rsget("allatdiscountprice")
        InsureCd            = rsget("InsureCd")
        ipkumdiv            = rsget("ipkumdiv")
        sumPaymentEtc       = rsget("sumPaymentEtc")
    end if
    rsget.close

IF (ERR) THEN response.write "ERR-step1"

    ''보조결제.
    sqlStr = " select acctdiv, IsNull(realPayedsum, 0) as realPayedsum " + VbCrlf
    sqlStr = sqlStr + " from " + VbCrlf
    sqlStr = sqlStr + " db_order.dbo.tbl_order_PaymentEtc " + VbCrlf
    sqlStr = sqlStr + " where " + VbCrlf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrlf
    sqlStr = sqlStr + " 	and orderserial = '" + orderserial + "' " + VbCrlf
    sqlStr = sqlStr + " 	and acctdiv in ('200', '900') " + VbCrlf			'200 : 예치금, 900 : 상품권

    rsget.Open sqlStr,dbget,1

	orgdepositsum = 0
	orggiftcardsum = 0
	do until rsget.eof
		if (CStr(rsget("acctdiv")) = "200") then
			orgdepositsum = rsget("realPayedsum")
		elseif (CStr(rsget("acctdiv")) = "900") then
			orggiftcardsum = rsget("realPayedsum")
		end if

		rsget.movenext
	loop
	rsget.close

IF (ERR) THEN response.write "ERR-step2"
    ''환불정보 -->
    sqlStr = " select r.*, a.gubun01, a.gubun02 from [db_cs].[dbo].tbl_new_as_list a"
    sqlStr = sqlStr + " , [db_cs].[dbo].tbl_as_refund_info r"
    sqlStr = sqlStr + " where a.id=" + CStr(id)
    sqlStr = sqlStr + " and a.id=r.asid"
    sqlStr = sqlStr + " and a.deleteyn='N'"
    sqlStr = sqlStr + " and a.currstate<>'B007'"

    rsget.Open sqlStr,dbget,1

    if Not rsget.Eof then
        refundmileagesum    = rsget("refundmileagesum")
        refundcouponsum     = rsget("refundcouponsum")
        allatsubtractsum    = rsget("allatsubtractsum")

        refunditemcostsum   = rsget("refunditemcostsum")

        refundbeasongpay    = rsget("refundbeasongpay")
        refunddeliverypay   = rsget("refunddeliverypay")
        refundadjustpay     = rsget("refundadjustpay")
        canceltotal         = rsget("canceltotal")
        gubun01             = rsget("gubun01")
        gubun02             = rsget("gubun02")

        refunddepositsum    = rsget("refunddepositsum")
    else
        refundmileagesum    = 0
        refundcouponsum     = 0
        allatsubtractsum    = 0

        refunditemcostsum   = 0

        refundbeasongpay    = 0
        refunddeliverypay   = 0
        refundadjustpay     = 0
        canceltotal         = 0
        refunddepositsum    = 0
    end if
    rsget.close

'' 마일리지 변경 유무
IF (ERR) THEN response.write "ERR-step3"

    IsUpdatedMile = false

    if (userid<>"") and (IsAllCancel) and (miletotalprice<>0) then
        '' 전체 취소인경우 주문건 취소로 jukyocd : 2 상품구매, 3 : 부분취소시 환급마일리지
        sqlStr = " update [db_user].[dbo].tbl_mileagelog " + VbCrlf
        sqlStr = sqlStr + " set deleteyn='Y' " + VbCrlf
        sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
        sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and jukyocd in ('2','3')" + VbCrlf

        dbget.Execute sqlStr

        IsUpdatedMile = true

        if openMessage="" then
            openMessage = openMessage + "사용 마일리지 환급 : " & miletotalprice
        else
            openMessage = openMessage + VbCrlf + "사용 마일리지 환급 : " & miletotalprice
        end if

    end if

'예치금환원
'TODO : 상품권환원
	IsUpdatedDeposit = false
    if (userid<>"") and (IsAllCancel) and (orgdepositsum <> 0) then
        '' 전체 취소인경우 주문건 취소로 jukyocd : 100 상품구매, 10 : 부분취소시 예치금 환원
        sqlStr = " update [db_user].[dbo].tbl_depositlog " + VbCrlf
        sqlStr = sqlStr + " set deleteyn='Y' " + VbCrlf
        sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
        sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and jukyocd in ('100','10')" + VbCrlf
        dbget.Execute sqlStr

        IsUpdatedDeposit = true

        if openMessage="" then
            openMessage = openMessage + "사용 예치금 환원 : " & orgdepositsum
        else
            openMessage = openMessage + VbCrlf + "사용 예치금 환원 : " & orgdepositsum
        end if

    end if

    if (userid<>"") and (Not IsAllCancel) and (refunddepositsum <> 0) then
        '' 부분 취소인데 예치금 환원할 경우.

        sqlStr = " update [db_order].[dbo].tbl_order_PaymentEtc" + VbCrlf
        sqlStr = sqlStr + " set realPayedsum=realPayedsum + " + CStr(refunddepositsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and acctdiv='200'" + VbCrlf
        dbget.Execute sqlStr

        sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
        sqlStr = sqlStr + " set sumPaymentETC=sumPaymentETC + " + CStr(refunddepositsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
        dbget.Execute sqlStr

        sqlStr = " insert into [db_user].[dbo].tbl_depositlog " + VbCrlf
        sqlStr = sqlStr + " (userid, deposit, jukyocd, jukyo, orderserial, deleteyn) " + VbCrlf
        sqlStr = sqlStr + " values ("
        sqlStr = sqlStr + " '" + userid + "'"
        sqlStr = sqlStr + " ," + CStr(refunddepositsum*-1) + ""
        sqlStr = sqlStr + " ,'10'"
        sqlStr = sqlStr + " ,'상품구매 취소 환원'"
        sqlStr = sqlStr + " ,'" + orderserial + "'"
        sqlStr = sqlStr + " ,'N'"
        sqlStr = sqlStr + " )"
        dbget.Execute sqlStr

        IsUpdatedDeposit = true

        if openMessage="" then
            openMessage = openMessage + "사용 예치금 환원 : " & refunddepositsum
        else
            openMessage = openMessage + VbCrlf + "사용 예치금 환원 : " & refunddepositsum
        end if
    end if

    ''부분취소 추가
    if (userid<>"") and (Not IsAllCancel) and (refundmileagesum<>0) then
        '' 부분 취소인데 마일리지 환급할 경우.
        sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
        sqlStr = sqlStr + " set miletotalprice=miletotalprice + " + CStr(refundmileagesum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        dbget.Execute sqlStr


        sqlStr = " insert into [db_user].[dbo].tbl_mileagelog " + VbCrlf
        sqlStr = sqlStr + " (userid, mileage, jukyocd, jukyo, orderserial, deleteyn) " + VbCrlf
        sqlStr = sqlStr + " values ("
        sqlStr = sqlStr + " '" + userid + "'"
        sqlStr = sqlStr + " ," + CStr(refundmileagesum*-1) + ""
        sqlStr = sqlStr + " ,'3'"
        sqlStr = sqlStr + " ,'상품구매 취소 환급'"
        sqlStr = sqlStr + " ,'" + orderserial + "'"
        sqlStr = sqlStr + " ,'N'"
        sqlStr = sqlStr + " )"

        dbget.Execute sqlStr

        IsUpdatedMile = true

        if openMessage="" then
            openMessage = openMessage + "사용 마일리지 환급 : " & refundmileagesum
        else
            openMessage = openMessage + VbCrlf + "사용 마일리지 환급 : " & refundmileagesum
        end if
    end if

''rw "E1."&Err.Number

'' 할인권 환급
    if (IsAllCancel) and (tencardspend<>0) then
        sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
	    sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
	    sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "
        sqlStr = sqlStr + " and userid='"&userid&"'"  ''2015/04/13 추가(느리므로)

	    dbget.Execute sqlStr

	    if openMessage="" then
            openMessage = openMessage + "사용 보너스쿠폰 환급"
        else
            openMessage = openMessage + VbCrlf + "사용 보너스쿠폰 환급"
        end if
    end if
''rw "E2."&Err.Number
    if (Not IsAllCancel) and (refundcouponsum<>0) then
         '' 부분 취소인경우 - 환급한 만큼 깜..
        sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
        sqlStr = sqlStr + " set tencardspend=tencardspend + " + CStr(refundcouponsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        dbget.Execute sqlStr

        ''전체 환급인 경우만 쿠폰을 돌려줌
        sqlStr = "select IsNULL(tencardspend,0) as tencardspend from [db_order].[dbo].tbl_order_master" + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        rsget.Open sqlStr,dbget,1
            remaintencardspend = rsget("tencardspend")
        rsget.close

        ''원래 할인권 사용액이 있고, 남은 쿠폰사용액이 없을경우 전체  환급
        if (tencardspend>0) then
            if (remaintencardspend=0)   then
                sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
            	sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
            	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "

            	dbget.Execute sqlStr

            	if openMessage="" then
                    openMessage = openMessage + "사용 할인권  환급"
                else
                    openMessage = openMessage + VbCrlf + "사용 할인권  환급"
                end if
            else
                ''(또는, %쿠폰인 경우 공통,단순변심인 경우 제외하고 환급해줌./ 부분취소 ) C004 CD01
                if (ipkumdiv>3) and (Not ((gubun01="C004") and (gubun02="CD01"))) then
                    sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
                	sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
                	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "
                	sqlStr = sqlStr + " and coupontype=1"

                	dbget.Execute sqlStr

                	if openMessage="" then
                        openMessage = openMessage + "사용 할인권  환급."
                    else
                        openMessage = openMessage + VbCrlf + "사용 할인권  환급."
                    end if
                end if
            end if
        end if



    end if

''rw "E3."&Err.Number

    '' 올엣카드 할인 차감
    if (IsAllCancel) and (allatdiscountprice<>0) then
        '' No Action
    end if

    if (Not IsAllCancel) and (allatsubtractsum<>0) then
        sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
        sqlStr = sqlStr + " set allatdiscountprice=allatdiscountprice + " + CStr(allatsubtractsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        dbget.Execute sqlStr

        if openMessage="" then
            openMessage = openMessage + "올엣카드 할인 차감 : " & allatsubtractsum
        else
            openMessage = openMessage + VbCrlf + "올엣카드 할인 차감 : " & allatsubtractsum
        end if
    end if
''rw "E4."&Err.Number

    '' 배송비 재계산. : 현재 배송비와 다를경우만. 부분 취소인 경우만. :: 업체 개별 배송비로 수정
    dim detailRefundBeasongPay
    detailRefundBeasongPay = 0
    sqlStr = " select IsNULL(sum(itemcost),0) as detailRefundBeasongPay from [db_cs].[dbo].tbl_new_as_detail"
    sqlStr = sqlStr + " where masterid=" + CStr(id)
    sqlStr = sqlStr + " and itemid=0"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        detailRefundBeasongPay = rsget("detailRefundBeasongPay")
    end if
    rsget.Close

    if (Not IsAllCancel) and (refundbeasongpay<>0) then
        orgbeasongpay =0

        ''기본배송비.
        sqlStr = " select * from [db_order].[dbo].tbl_order_detail "
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
        sqlStr = sqlStr + " and itemid=0"
        sqlStr = sqlStr + " and IsNULL(makerid,'')=''"
        sqlStr = sqlStr + " and cancelyn<>'Y'"

        rsget.Open sqlStr,dbget,1
            detailidx     = rsget("idx")
            orgbeasongpay = rsget("itemcost")
        rsget.Close

        ''원래 텐배송 비가 >0 이고, 환불배송비가=텐배송비고,
'response.write "orgbeasongpay=" & orgbeasongpay & "<br>"
'response.write "refundbeasongpay=" & refundbeasongpay & "<br>"
'response.write "detailRefundBeasongPay=" & detailRefundBeasongPay & "<br>"

        if (orgbeasongpay>0) and (orgbeasongpay-refundbeasongpay=0) and (refundbeasongpay-detailRefundBeasongPay>0) then
             sqlStr = " update [db_order].[dbo].tbl_order_detail "
             sqlStr = sqlStr + " set itemoption='0000'"
             sqlStr = sqlStr + " ,itemcost=0"
             sqlStr = sqlStr + " where idx=" + CStr(detailidx)

             dbget.Execute sqlStr
             response.write   "원 기본 배송비(" & orgbeasongpay & ") 0 원 처리"
        else

        end if
    end if
''rw "E5."&Err.Number
    if (IsAllCancel) then
	    ''전체 취소인경우
	    '' 주문  master 취소 변경
	    Call setCancelMaster(id, orderserial)

	    if openMessage="" then
            openMessage = openMessage + "주문취소 완료"
        else
            openMessage = openMessage + VbCrlf + "주문취소 완료"
        end if
    else
	    ''부분 취소인경우
	    '' 주문  detail 취소 변경
	    call setCancelDetail(id, orderserial)

	    call reCalcuOrderMaster(orderserial)
''rw "E7."&Err.Number
	    if openMessage="" then
            openMessage = openMessage + "주문부분취소 완료"
        else
            openMessage = openMessage + VbCrlf + "주문부분취소 완료"
        end if
	end if

    ''마일리지는 주문건 취소 후 재계산해야함.
    '예치금 재계산
    if (userid<>"") then
        Call UpdateUserMileage(userid)

        if (IsUpdatedDeposit) then
        	Call updateUserDeposit(userid)
        end if

        if IsUpdatedGiftCard then
        	Call updateUserGiftCard(userid)
        end if
    end if

    ''한정 수량 조정 - setCancelMaster에 포함( 한정 조정 및 재고 업데이트)
    '''Call LimitItemRecover(orderserial)
    if (IsAllCancel) then
        '''setCancelMaster에 포함( 한정 조정 및 재고 업데이트)
    else
	    ''부분 취소인경우
	    sqlStr = " select itemid,itemoption,regitemno "
        sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_detail "
        sqlStr = sqlStr & " where masterid=" & id
        sqlStr = sqlStr & " and orderserial='" & orderserial & "'"

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            regDetailRows = rsget.getRows()
        end if
        rsget.Close

        if IsArray(regDetailRows) then
            for i=0 to UBound(regDetailRows,2)
    	        sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & regDetailRows(0,i) & ",'" & regDetailRows(1,i) & "'," & regDetailRows(2,i)
                dbget.Execute sqlStr
            Next
        end if
	end if
''rw "E10."&Err.Number
    ''전자보증서 발급된 경우 취소
    if (InsureCd="0") then
        Call UsafeCancel(orderserial)
    end if


    if (openMessage<>"") then
        call AddCustomerOpenContents(id, openMessage)
    end if
end function





'function EditCSMaster(divcd, orderserial, modiuserid, title, contents_jupsu, gubun01, gubun02)
'    '' CS Master 수정
'    dim sqlStr
'
'    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
'    sqlStr = sqlStr + " set writeuser='" + modiuserid + "'"
'    sqlStr = sqlStr + " ,title='" + title + "'"
'    sqlStr = sqlStr + " ,contents_jupsu='" + contents_jupsu + "'"
'    sqlStr = sqlStr + " ,gubun01='" + gubun01 + "'"
'    sqlStr = sqlStr + " ,gubun02='" + gubun02 + "'"
'    sqlStr = sqlStr + " where id=" + CStr(id)
'
'    dbget.Execute sqlStr
'
'end function

'function EditCSMasterFinished(divcd, orderserial, modiuserid, title, contents_jupsu, gubun01, gubun02, finishuserid, contents_finish)
'    '' CS Master 완료된 내역 수정
'    dim sqlStr
'
'    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
'    sqlStr = sqlStr + " set finishuser='" + finishuserid + "'"
'    sqlStr = sqlStr + " ,title='" + title + "'"
'    sqlStr = sqlStr + " ,contents_jupsu='" + contents_jupsu + "'"
'    sqlStr = sqlStr + " ,contents_finish='" + contents_finish + "'"
'    sqlStr = sqlStr + " ,gubun01='" + gubun01 + "'"
'    sqlStr = sqlStr + " ,gubun02='" + gubun02 + "'"
'    sqlStr = sqlStr + " where id=" + CStr(id)
'
'    dbget.Execute sqlStr
'end function

function RegCSMasterRefundInfo(asid, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay  , rebankname, rebankaccount, rebankownername, paygateTid)
    dim sqlStr
    if IsNULL(orgmileagesum) then orgmileagesum=0
	if IsNULL(paygateTid) then paygateTid=""

    sqlStr = " insert into [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " (asid"
    sqlStr = sqlStr + " ,returnmethod"
    sqlStr = sqlStr + " ,refundrequire"
    sqlStr = sqlStr + " ,orgsubtotalprice"
    sqlStr = sqlStr + " ,orgitemcostsum"
    sqlStr = sqlStr + " ,orgbeasongpay"
    sqlStr = sqlStr + " ,orgmileagesum"
    sqlStr = sqlStr + " ,orgcouponsum"
    sqlStr = sqlStr + " ,orgallatdiscountsum"
    sqlStr = sqlStr + " ,canceltotal"
    sqlStr = sqlStr + " ,refunditemcostsum"
    sqlStr = sqlStr + " ,refundmileagesum"
    sqlStr = sqlStr + " ,refundcouponsum"
    sqlStr = sqlStr + " ,allatsubtractsum"
    sqlStr = sqlStr + " ,refundbeasongpay"
    sqlStr = sqlStr + " ,refunddeliverypay"
    sqlStr = sqlStr + " ,refundadjustpay"
    sqlStr = sqlStr + " ,rebankname"
    sqlStr = sqlStr + " ,rebankaccount"
    sqlStr = sqlStr + " ,rebankownername"
    sqlStr = sqlStr + " ,paygateTid"
    sqlStr = sqlStr + " )"

    sqlStr = sqlStr + " values("
    sqlStr = sqlStr + " " + CStr(asid)
    sqlStr = sqlStr + " ,'" + returnmethod + "'"
    sqlStr = sqlStr + " ," + CStr(refundrequire)
    sqlStr = sqlStr + " ," + CStr(orgsubtotalprice)
    sqlStr = sqlStr + " ," + CStr(orgitemcostsum)
    sqlStr = sqlStr + " ," + CStr(orgbeasongpay)
    sqlStr = sqlStr + " ," + CStr(orgmileagesum)
    sqlStr = sqlStr + " ," + CStr(orgcouponsum)
    sqlStr = sqlStr + " ," + CStr(orgallatdiscountsum)
    sqlStr = sqlStr + " ," + CStr(canceltotal)
    sqlStr = sqlStr + " ," + CStr(refunditemcostsum)
    sqlStr = sqlStr + " ," + CStr(refundmileagesum)
    sqlStr = sqlStr + " ," + CStr(refundcouponsum)
    sqlStr = sqlStr + " ," + CStr(allatsubtractsum)
    sqlStr = sqlStr + " ," + CStr(refundbeasongpay)
    sqlStr = sqlStr + " ," + CStr(refunddeliverypay)
    sqlStr = sqlStr + " ," + CStr(refundadjustpay)
    sqlStr = sqlStr + " ,'" + rebankname + "'"
    sqlStr = sqlStr + " ,'" + rebankaccount + "'"
    sqlStr = sqlStr + " ,'" + rebankownername + "'"
    sqlStr = sqlStr + " ,'" + paygateTid + "'"
    sqlStr = sqlStr + " )"
    dbget.Execute sqlStr

end function


'function EditCSDetailByArrStr(byval detailitemlist, id, orderserial)
'    dim sqlStr, tmp, buf, i
'    dim dorderdetailidx, dgubun01, dgubun02, dregitemno, dcausecontent
'
'    buf = split(detailitemlist, "|")
'
'    for i = 0 to UBound(buf)
'		if (TRIM(buf(i)) <> "") then
'			tmp = split(buf(i), Chr(9))
'
'			dorderdetailidx = tmp(0)
'			dgubun01        = tmp(1)
'			dgubun02        = tmp(2)
'			dregitemno      = tmp(3)
'			dcausecontent   = tmp(4)
'
'	        call EditOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno, dcausecontent)
'		end if
'	next
'
'end function


'function AddOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno)
'    dim sqlStr
'
'    sqlStr = " insert into [db_cs].[dbo].tbl_new_as_detail"
'    sqlStr = sqlStr + " (masterid, orderdetailidx, gubun01,gubun02"
'    sqlStr = sqlStr + " ,orderserial, itemid, itemoption, makerid, itemname, itemoptionname, regitemno, confirmitemno,orderitemno) "
'    sqlStr = sqlStr + " values(" + CStr(id) + ""
'    sqlStr = sqlStr + " ," + CStr(dorderdetailidx) + ""
'    sqlStr = sqlStr + " ,'" + CStr(dgubun01) + "'"
'    sqlStr = sqlStr + " ,'" + CStr(dgubun02) + "'"
'    sqlStr = sqlStr + " ,'" + CStr(orderserial) + "'"
'    sqlStr = sqlStr + " ,0"
'    sqlStr = sqlStr + " ,''"
'    sqlStr = sqlStr + " ,''"
'    sqlStr = sqlStr + " ,''"
'    sqlStr = sqlStr + " ,''"
'    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
'    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
'    sqlStr = sqlStr + " ,0"
'    sqlStr = sqlStr + " )"
'
'    dbget.Execute sqlStr
'end function


'function EditOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno, dcausecontent)
'    dim sqlStr
'
'    sqlStr = " update [db_cs].[dbo].tbl_new_as_detail"
'    sqlStr = sqlStr + " set gubun01='" + dgubun01 + "'"
'    sqlStr = sqlStr + " , gubun02='" + dgubun02 + "'"
'    sqlStr = sqlStr + " , regitemno=" + dregitemno + ""
'    sqlStr = sqlStr + " , confirmitemno=" + dregitemno + ""
'    sqlStr = sqlStr + " , causecontent='" + dregitemno + "'"
'    sqlStr = sqlStr + " where masterid=" + CStr(id)
'    sqlStr = sqlStr + " and orderdetailidx=" + CStr(dorderdetailidx)
'
'    dbget.Execute sqlStr
'end function


'function AddOneDeliveryInfoCSDetail(id, gubun01, gubun02, orderserial)
'    dim sqlStr
'
'    sqlStr = " insert into [db_cs].[dbo].tbl_new_as_detail"
'    sqlStr = sqlStr + " (masterid, orderdetailidx, gubun01, gubun02,"
'    sqlStr = sqlStr + " orderserial, itemid, itemoption, makerid,itemname, itemoptionname,"
'    sqlStr = sqlStr + " regitemno, confirmitemno, orderitemno, itemcost, buycash, isupchebeasong,regdetailstate) "
'    sqlStr = sqlStr + " select top 1 "
'    sqlStr = sqlStr + " " + CStr(id)
'    sqlStr = sqlStr + " ,d.idx"
'    sqlStr = sqlStr + " ,'" + CStr(gubun01) + "'"
'    sqlStr = sqlStr + " ,'" + CStr(gubun02) + "'"
'    sqlStr = sqlStr + " ,d.orderserial"
'    sqlStr = sqlStr + " ,d.itemid"
'    sqlStr = sqlStr + " ,d.itemoption"
'    sqlStr = sqlStr + " ,IsNULL(d.makerid,'')"
'    sqlStr = sqlStr + " ,IsNULL(d.itemname,'배송료')"
'    sqlStr = sqlStr + " ,IsNULL(d.itemoptionname,(case when d.itemcost=0 then '무료배송' else '일반택배' end))"
'    sqlStr = sqlStr + " ,d.itemno"
'    sqlStr = sqlStr + " ,d.itemno"
'    sqlStr = sqlStr + " ,d.itemno"
'    sqlStr = sqlStr + " ,d.itemcost"
'    sqlStr = sqlStr + " ,d.buycash"
'    sqlStr = sqlStr + " ,d.isupchebeasong"
'    sqlStr = sqlStr + " ,d.currstate"
'    sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d"
'    sqlStr = sqlStr + " where d.orderserial='" + orderserial + "'"
'    sqlStr = sqlStr + " and d.itemid=0"
'    sqlStr = sqlStr + " and d.cancelyn<>'Y'"
'
'    dbget.Execute sqlStr
'
'end function



''바로 완료 처리로 진행 할지 여부.
'function IsDirectProceedFinish(divcd, Asid, orderserial, byRef EtcStr)
'    dim sqlStr
'    dim cancelyn, ipkumdiv
'    IsDirectProceedFinish = false
'
'    '' currstate:2 업체(물류) 통보
'    if (divcd="A008") then
'        ''' 취소 Case
'        '' 등록된 상품중 업체 확인중 상태가 있으면 접수상태로 진행
'        sqlStr = " select count(d.idx) as invalidcount"
'        sqlStr = sqlStr + " from "
'        sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
'        sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
'        sqlStr = sqlStr + " [db_cs].[dbo].tbl_new_as_detail c"
'        sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"
'        sqlStr = sqlStr + " and m.orderserial=d.orderserial"
'        sqlStr = sqlStr + " and d.itemid<>0"
'        sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
'        sqlStr = sqlStr + " and d.idx=c.orderdetailidx"
'        sqlStr = sqlStr + " and d.currstate>=3"
'        sqlStr = sqlStr + " and d.cancelyn<>'Y'"
'
'        rsget.Open sqlStr,dbget,1
'            IsDirectProceedFinish = (rsget("invalidcount")=0)
'        rsget.close
'
'    else
'
'    end if
'
'end function

''검증. 전체 취소 맞는지.
function IsAllCancelRegValid(Asid, orderserial)
    dim sqlStr
    IsAllCancelRegValid = false

    sqlStr = "select count(d.idx) as cnt"
    sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d"
    sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_detail c"
    sqlStr = sqlStr + "     on c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + "     and c.orderdetailidx=d.idx"
    sqlStr = sqlStr + " where d.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and d.itemid<>0"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"
    sqlStr = sqlStr + " and d.itemno<>IsNULL(c.regitemno,0)"

    rsget.Open sqlStr,dbget,1
        IsAllCancelRegValid = (rsget("cnt")=0)
    rsget.close

end function

''검증. 부분 취소 맞는지.
function IsPartialCancelRegValid(Asid, orderserial)
    dim sqlStr
    IsPartialCancelRegValid = false

    sqlStr = "select count(d.idx) as cnt, sum(case when d.itemno=IsNULL(c.regitemno,0) then 1 else 0 end) as Matchcount"
    sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d"
    sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_detail c"
    sqlStr = sqlStr + "     on c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + "     and c.orderdetailidx=d.idx"
    sqlStr = sqlStr + " where d.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and d.itemid<>0"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"

    rsget.Open sqlStr,dbget,1
        IsPartialCancelRegValid = Not (rsget("cnt")=rsget("Matchcount"))
    rsget.close
end function

function SaveCSListHistory(asid)
    dim sqlStr

	'// 이전 처리자 아이디 저장
	sqlStr = " exec [db_log].[dbo].[usp_Ten_SaveCSHistory] " + CStr(asid) + " "
	dbget.Execute(sqlStr)

end function

' 주문 상품쿠폰 적용수 체크		'2023.10.19 한용민 생성
function ItemCouponCount(asid, couponGubun, userid)
    dim sqlStr, returnCount
    returnCount=0

	if asid="" or isnull(asid) then
        ItemCouponCount=returnCount
        exit function
    end if
    asid = trim(asid)
	if userid="" or isnull(userid) then
        ItemCouponCount=returnCount
        exit function
    end if
    userid = trim(userid)

    sqlStr = " select"
    sqlStr = sqlStr & " count(t.itemcouponidx) as itemCouponCount"
    sqlStr = sqlStr & " from ("
    sqlStr = sqlStr & " 	select"
    sqlStr = sqlStr & " 	d.itemcouponidx"
    sqlStr = sqlStr & " 	, isnull((select count(cc.couponidx)"
    sqlStr = sqlStr & "  		from db_item.dbo.tbl_user_item_coupon cc with (nolock)"
    sqlStr = sqlStr & "  		where cc.itemcouponidx = c.itemcouponidx"
    sqlStr = sqlStr & "  		and cc.userid = c.userid"
    sqlStr = sqlStr & "  		and cc.couponidx <> c.couponidx),0) as prevCopiedItemCouponCount"
    sqlStr = sqlStr & "  		, rank() over (partition by c.userid, c.itemcouponidx order by c.couponidx desc) as rk"
    sqlStr = sqlStr & " 	from db_cs.dbo.tbl_new_as_detail ad with (nolock)"
    sqlStr = sqlStr & " 	join db_order.dbo.tbl_order_detail d with (nolock)"
    sqlStr = sqlStr & " 		on ad.orderdetailidx=d.idx"
    sqlStr = sqlStr & " 		and ad.orderserial=d.orderserial"
    sqlStr = sqlStr & " 	join db_item.dbo.tbl_user_item_coupon c with (nolock)"
    sqlStr = sqlStr & " 		on d.itemcouponidx=c.itemcouponidx"
    sqlStr = sqlStr & " 		and ad.orderserial=c.orderserial"
    sqlStr = sqlStr & " 		and c.itemcouponexpiredate>getdate()"	' 유효기간체크

    if couponGubun<>"" then
        sqlStr = sqlStr & " and c.couponGubun='"& couponGubun &"'"
    end if
    if userid<>"" then
        sqlStr = sqlStr & " and c.userid='"& userid &"'"
    end if
    
    sqlStr = sqlStr & " 	where ad.masterid='" & asid & "'"
    sqlStr = sqlStr & " ) as t"
    sqlStr = sqlStr & " where t.rk=1"
    sqlStr = sqlStr & " and t.prevCopiedItemCouponCount=0"    ' 재발급체크

	'response.write sqlStr & "<Br>"
    'response.end
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.Eof) then
        returnCount=rsget("itemCouponCount")
    else
        returnCount=0
    end if
    rsget.Close

    ItemCouponCount=returnCount
end function

' 상품쿠폰 복사발급     ' 2023.10.19 한용민 생성
function CheckAndCopyItemCoupon(asid, reguserid, couponGubun, userid)
	dim orderserial, copyitemcouponinfo, sqlStr, excuteRowCount
    excuteRowCount=0

	if asid="" or isnull(asid) then
        CheckAndCopyItemCoupon = False
        exit function
    end if
    asid = trim(asid)

	sqlStr = " select top 1"
    sqlStr = sqlStr & " a.orderserial, IsNull(r.copyitemcouponinfo, 'N') as copyitemcouponinfo"
	sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_list a with (nolock)"
	sqlStr = sqlStr & " join [db_cs].[dbo].tbl_as_refund_info r with (nolock)"
	sqlStr = sqlStr & " 	on a.id = r.asid "
	sqlStr = sqlStr & " where a.id = "& asid &""
    sqlStr = sqlStr & " and a.divcd in ('A008', 'A004', 'A010')"    ' A008 주문취소 / A004 반품접수(업체배송) / A010 회수신청(텐바이텐배송)

	orderserial = ""
	copyitemcouponinfo = "N"
    'response.write sqlStr & "<br>"
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        orderserial    	= rsget("orderserial")
		copyitemcouponinfo  = rsget("copyitemcouponinfo")
    end if
    rsget.Close

	if (orderserial = "") or (copyitemcouponinfo <> "Y") then
		CheckAndCopyItemCoupon = False
		exit function
	end if

    sqlStr = "insert into db_item.dbo.tbl_user_item_coupon("
    sqlStr = sqlStr & " userid, itemcouponidx, issuedno, itemcoupontype, itemcouponvalue"
    sqlStr = sqlStr & " , itemcouponstartdate, itemcouponexpiredate"
    sqlStr = sqlStr & " , itemcouponname, itemcouponimage, regdate, usedyn, orderserial, couponGubun, csorderserial"
    sqlStr = sqlStr & " )"
    sqlStr = sqlStr & "     select"
    sqlStr = sqlStr & "     t.userid, t.itemcouponidx, t.issuedno, t.itemcoupontype, t.itemcouponvalue"
    sqlStr = sqlStr & "     , t.itemcouponstartdate, t.itemcouponexpiredate"
    sqlStr = sqlStr & "     , t.itemcouponname, t.itemcouponimage, t.regdate, t.usedyn"
    sqlStr = sqlStr & "     , t.orderserial, t.couponGubun, t.csorderserial"
    sqlStr = sqlStr & "     from ("
    sqlStr = sqlStr & " 	    select"
    sqlStr = sqlStr & "         c.userid, c.itemcouponidx, c.issuedno, c.itemcoupontype, c.itemcouponvalue"
    sqlStr = sqlStr & "         , c.itemcouponstartdate, c.itemcouponexpiredate"
    sqlStr = sqlStr & "         , c.itemcouponname, c.itemcouponimage, getdate() as regdate"
    sqlStr = sqlStr & "         , 'N' as usedyn, NULL as orderserial, c.couponGubun, c.orderserial as csorderserial"
    sqlStr = sqlStr & " 	    , isnull((select count(cc.couponidx)"
    sqlStr = sqlStr & "  	    	from db_item.dbo.tbl_user_item_coupon cc with (nolock)"
    sqlStr = sqlStr & "  	    	where cc.itemcouponidx = c.itemcouponidx"
    sqlStr = sqlStr & "  	    	and cc.userid = c.userid"
    sqlStr = sqlStr & "  	    	and cc.couponidx <> c.couponidx),0) as prevCopiedItemCouponCount"
    sqlStr = sqlStr & "  		, rank() over (partition by c.userid, c.itemcouponidx order by c.couponidx desc) as rk"
    sqlStr = sqlStr & " 	    from db_cs.dbo.tbl_new_as_detail ad with (nolock)"
    sqlStr = sqlStr & " 	    join db_order.dbo.tbl_order_detail d with (nolock)"
    sqlStr = sqlStr & " 	    	on ad.orderdetailidx=d.idx"
    sqlStr = sqlStr & " 	    	and ad.orderserial=d.orderserial"
    sqlStr = sqlStr & " 	    join db_item.dbo.tbl_user_item_coupon c with (nolock)"
    sqlStr = sqlStr & " 	    	on d.itemcouponidx=c.itemcouponidx"
    sqlStr = sqlStr & " 	    	and ad.orderserial=c.orderserial"
    sqlStr = sqlStr & " 	    	and c.itemcouponexpiredate>getdate()"	' 유효기간체크

    if couponGubun<>"" then
        sqlStr = sqlStr & "     and c.couponGubun='"& couponGubun &"'"
    end if
    if userid<>"" then
        sqlStr = sqlStr & "     and c.userid='"& userid &"'"
    end if
    
    sqlStr = sqlStr & " 	    where ad.masterid='" & asid & "'"
    sqlStr = sqlStr & "     ) as t"
    sqlStr = sqlStr & "     where t.rk=1"
    sqlStr = sqlStr & "     and t.prevCopiedItemCouponCount=0"    ' 재발급체크

    'response.write sqlStr & "<br>"
	dbget.Execute sqlStr, excuteRowCount

    if excuteRowCount>0 then
	    CheckAndCopyItemCoupon = True
    else
        ' 상품쿠폰 재발급 접수 와 완료처리 사이에 쿠폰 유효기간이 지난경우 재발행여부 N으로 바꾼다.
        Call EditCSCopyItemCouponInfo(asid, "N")

        CheckAndCopyItemCoupon = false
    end if
end function

' 상품쿠폰 복사여부     ' 2023.10.19 한용민 생성
function EditCSCopyItemCouponInfo(asid, copyitemcouponinfo)
	dim sqlStr

	if asid="" or isnull(asid) or copyitemcouponinfo="" or isnull(copyitemcouponinfo) then exit function
    asid = trim(asid)

    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info "
    sqlStr = sqlStr & " set copyitemcouponinfo = '" & copyitemcouponinfo & "' "
    sqlStr = sqlStr & " where asid = " & asid & " "
    
    'response.write sqlStr & "<br>"
    dbget.Execute sqlStr
end function

''주문 상세 내역이 취소 가능한지 체크 - 출고 완료된 내역이 있는지, 주문건이 취소된내역이 있는지
function IsWebCancelValidState(Asid, orderserial)
    dim sqlStr

    IsWebCancelValidState = false

    sqlStr = " select m.cancelyn, m.ipkumdiv, count(d.idx) as invalidcount, sum(case when d.cancelyn='Y' then 1 else 0 end) as detailcancelcount "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
    sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
    sqlStr = sqlStr + " [db_cs].[dbo].tbl_new_as_detail c"
    sqlStr = sqlStr + " where m.orderserial=d.orderserial"
    sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + " and d.idx=c.orderdetailidx"
    sqlStr = sqlStr + " and d.currstate>=7"
    sqlStr = sqlStr + " group by m.cancelyn, m.ipkumdiv"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        IsWebCancelValidState = (rsget("cancelyn")="N") and (rsget("ipkumdiv")<7) and (rsget("invalidcount")<1) and (rsget("detailcancelcount")<1)
    else
        IsWebCancelValidState = true
    end if
    rsget.close

end function

function GetTotalItemNo(orderserial)
    dim sqlStr
    GetTotalItemNo = 0

	sqlStr = " select IsNull(sum(d.itemno),0) as totItemNo "
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d "
	sqlStr = sqlStr + " where d.orderserial = '" & orderserial & "' and d.itemid <> 0 and d.cancelyn <> 'Y' "

    rsget.Open sqlStr,dbget,1
    	GetTotalItemNo = rsget("totItemNo")
    rsget.close

end function

function IsWebReturnValidState(Asid, orderserial, byref iScanErr)
    dim sqlStr
    IsWebReturnValidState = false

    sqlStr = " select ipkumdiv, cancelyn from [db_order].[dbo].tbl_order_master"
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        cancelyn    = rsget("cancelyn")
        ipkumdiv    = rsget("ipkumdiv")
    end if
    esget.Close

    if (cancelyn<>"N") then Exit function

    IsWebReturnValidState = true
end function

function setCancelMaster(Asid, orderserial)
    dim sqlStr
    sqlStr = "update [db_order].[dbo].tbl_order_master" + VbCrlf
    sqlStr = sqlStr + " set cancelyn='Y'" + VbCrlf
    sqlStr = sqlStr + " ,canceldate=getdate()" + VbCrlf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

    dbget.Execute sqlStr

    ''한정수량 조정 및 재고 업데이트
    '''On Error Resume Next
    sqlStr = " exec [db_summary].[dbo].sp_ten_RealtimeStock_cancelOrderAll '" & orderserial & "'"
    dbget.Execute sqlStr
    '''On Error Goto 0
end function



'' 수량이 같으면 취소 Flag 다르면 수량변경
function setCancelDetail(Asid, orderserial)
    dim sqlStr
    ''취소일 추가
    sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
    sqlStr = sqlStr + " set cancelyn='Y'" + VbCrlf
    sqlStr = sqlStr + " ,canceldate=IsNULL(canceldate,getdate())" + VbCrlf
    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_detail c" + VbCrlf
    sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_detail.orderserial='" + orderserial + "'" + VbCrlf
    sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_detail.idx=c.orderdetailidx" + VbCrlf
    sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_detail.itemno=c.regitemno" + VbCrlf
    '''sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_detail.itemid<>0"
    '''배송비도 취소?

    dbget.Execute sqlStr

    '' 수량조정 ::: (몇개 만 취소인경우)
    sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
    sqlStr = sqlStr + " set itemno=itemno-c.regitemno" + VbCrlf
    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_detail c" + VbCrlf
    sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_detail.orderserial='" + orderserial + "'" + VbCrlf
    sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_detail.idx=c.orderdetailidx" + VbCrlf
    sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_detail.itemno>c.regitemno" + VbCrlf
    sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_detail.itemid<>0"

    dbget.Execute sqlStr
end function

function GetPartialCancelRegValidResult(Asid, orderserial)
	'검증 - 일부취소 접수
	'
	' - 부분취소인지
	' - 초과취소인지
	' - 이중취소인지
	' - 마스터 취소 되었는지

    dim sqlStr, result
    GetPartialCancelRegValidResult = ""
    result = ""

	'==========================================================================
	' - 마스터 취소 되었는지
	'==========================================================================
	if (IsMasterCanceled(Asid, orderserial)) then
		GetPartialCancelRegValidResult = "취소된 주문입니다."
		exit function
	end if

	'==========================================================================
	'부분취소인지 - 디테일 취소 접수(CS처리완료제외) 전체의 합이 잔여주문수량보다 작은것이 있는지
	'초과취소인지 - 디테일 취소 접수(CS처리완료제외) 전체의 합이 잔여주문수량보다 큰것이 있는지
	'==========================================================================
	if (IsErrorCancelState(Asid, orderserial)) then
		GetPartialCancelRegValidResult = "주문수량을 초과하여 취소(CS접수 포함)된 상품이 있습니다."
		exit function
	end if

	'==========================================================================
	'이중취소인지 - 취소된 디테일에 대한 취소가 있는지
	'==========================================================================
	if (IsDoubleCancelState(Asid, orderserial)) then
		GetPartialCancelRegValidResult = "취소된 상품에 대한 취소가 있습니다."
		exit function
	end if

end function

function IsMasterCanceled(Asid, orderserial)
    dim sqlStr, result
    IsMasterCanceled = false
    result = ""

	'==========================================================================
	' - 마스터 취소 되었는지
	'==========================================================================
    sqlStr = " select top 1 "
    sqlStr = sqlStr + " 	m.cancelyn as ordercancelyn "
    sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

    if Not rsget.Eof then
    	if (rsget("ordercancelyn") <> "N") then
    		IsMasterCanceled = true
    	end if
    end if
    rsget.close

end function

'초과취소 상태인지
function IsErrorCancelState(Asid, orderserial)
    dim sqlStr, result
    IsErrorCancelState = false

	'==========================================================================
	'초과취소인지 - 디테일 취소 접수(CS처리완료제외) 전체의 합이 잔여주문수량보다 큰지
	'==========================================================================
    sqlStr = " select "
    sqlStr = sqlStr + "     d.itemno "
    sqlStr = sqlStr + "     , sum(IsNULL(csd.regitemno,0)) as totalcancelregno "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
    sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_detail d "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		m.orderserial = d.orderserial "
    sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_new_as_list csm "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and m.orderserial = csm.orderserial "
    sqlStr = sqlStr + " 		and csm.divcd = 'A008' "
    sqlStr = sqlStr + " 		and csm.currstate <> 'B007' "
    sqlStr = sqlStr + " 		and csm.deleteyn <> 'Y' "
    sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_new_as_detail csd "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and csm.id = csd.masterid "
    sqlStr = sqlStr + " 		and csd.orderdetailidx = d.idx "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr + " 	and d.itemid <> 0 "
    sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
    sqlStr = sqlStr + " group by "
    sqlStr = sqlStr + " 	m.idx, d.idx, d.itemno "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
	if  not rsget.EOF  then
		do until rsget.eof
	    	if (rsget("itemno") < rsget("totalcancelregno")) then
	    		IsErrorCancelState = true
				exit do
	    	end if
			rsget.moveNext
		loop
	end if
	rsget.close

end function

'디테일 이중취소 있는지
function IsDoubleCancelState(Asid, orderserial)
    dim sqlStr, result
    IsDoubleCancelState = false

	'==========================================================================
	'이중취소인지 - 취소된 디테일에 대한 취소가 있는지
	'==========================================================================
    sqlStr = " select top 1 "
    sqlStr = sqlStr + "     d.itemid "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
    sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_detail d "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		m.orderserial = d.orderserial "
    sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list csm "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and m.orderserial = csm.orderserial "
    sqlStr = sqlStr + " 		and csm.id = " & Asid & " "
    sqlStr = sqlStr + " 		and csm.divcd = 'A008' "
    sqlStr = sqlStr + " 		and csm.currstate <> 'B007' "
    sqlStr = sqlStr + " 		and csm.deleteyn <> 'Y' "
    sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_detail csd "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and csm.id = csd.masterid "
    sqlStr = sqlStr + " 		and csd.orderdetailidx = d.idx "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr + " 	and d.itemid <> 0 "
    sqlStr = sqlStr + " 	and d.cancelyn = 'Y' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

	if  not rsget.EOF  then
		IsDoubleCancelState = true
	end if
	rsget.close

end function

''주문 상세 내역이 취소 가능한지 체크 - 출고 완료된 내역이 있는지, 주문건이 취소된내역이 있는지
function IsCancelValidState(Asid, orderserial)
    dim sqlStr

    IsCancelValidState = false

    sqlStr = " select m.cancelyn, m.ipkumdiv, sum(case when d.currstate>=7 then 1 else 0 end) as invalidcount, sum(case when d.cancelyn='Y' then 1 when c.confirmitemno > d.itemno then 1 else 0 end) as detailcancelcount "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
    sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
    sqlStr = sqlStr + " [db_cs].[dbo].tbl_new_as_detail c"
    sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and m.orderserial=d.orderserial"
    sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + " and d.idx=c.orderdetailidx"
    ''sqlStr = sqlStr + " and d.currstate>=7"
    sqlStr = sqlStr + " group by m.cancelyn, m.ipkumdiv"
	''response.write sqlStr
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

    if Not rsget.Eof then
        IsCancelValidState = (rsget("cancelyn")="N") and (rsget("ipkumdiv")<=7) and (rsget("invalidcount")<1) and (rsget("detailcancelcount")<1)
    else
        IsCancelValidState = true
    end if
    rsget.close

end function

''주문 마스타 재계산
function recalcuOrderMaster(byVal orderserial)
	dim sqlStr

	sqlStr = "update [db_order].[dbo].tbl_order_master" + VbCrlf
	sqlStr = sqlStr + " set totalsum=IsNULL(T.dtotalsum,0)" + VbCrlf
	''sqlStr = sqlStr + " , totalcost=IsNULL(T.dtotalsum,0)"  + VbCrlf
	sqlStr = sqlStr + " , totalmileage=IsNULL(T.dtotalmileage,0)" + VbCrlf
	sqlStr = sqlStr + " , subtotalpriceCouponNotApplied=IsNULL(T.dtotalitemcostCouponNotApplied,0)" + VbCrlf
	sqlStr = sqlStr + " from (" + VbCrlf
	sqlStr = sqlStr + "     select sum(itemcost*itemno) as dtotalsum, sum(mileage*itemno) as dtotalmileage, sum(IsNull(itemcostCouponNotApplied,0)*itemno) as dtotalitemcostCouponNotApplied" + VbCrlf
	sqlStr = sqlStr + "     from [db_order].[dbo].tbl_order_detail" + VbCrlf
	sqlStr = sqlStr + "     where orderserial='" + orderserial + "'" + VbCrlf
	sqlStr = sqlStr + "     and cancelyn<>'Y'" + VbCrlf
	sqlStr = sqlStr + " ) T" + VbCrlf
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.orderserial='" + orderserial + "'" + VbCrlf

	dbget.Execute sqlStr

	sqlStr = " update m " + VbCrlf
	sqlStr = sqlStr + " set " + VbCrlf
	sqlStr = sqlStr + " 	m.sumPaymentEtc = IsNull(T.realPayedsum, 0) " + VbCrlf
	sqlStr = sqlStr + " from " + VbCrlf
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m " + VbCrlf
	sqlStr = sqlStr + " 	left join ( " + VbCrlf
	sqlStr = sqlStr + " 		select " + VbCrlf
	sqlStr = sqlStr + " 			orderserial " + VbCrlf
	sqlStr = sqlStr + " 			, IsNull(sum(realPayedsum), 0) as realPayedsum " + VbCrlf
	sqlStr = sqlStr + " 		from " + VbCrlf
	sqlStr = sqlStr + " 			[db_order].[dbo].tbl_order_PaymentEtc " + VbCrlf
	sqlStr = sqlStr + " 		where " + VbCrlf
	sqlStr = sqlStr + " 			1 = 1 " + VbCrlf
	sqlStr = sqlStr + " 			and orderserial = '" & orderserial & "' " + VbCrlf
	sqlStr = sqlStr + " 			and acctdiv in ('200', '900') " + VbCrlf
	sqlStr = sqlStr + " 		group by " + VbCrlf
	sqlStr = sqlStr + " 			orderserial " + VbCrlf
	sqlStr = sqlStr + " 	) T " + VbCrlf
	sqlStr = sqlStr + " 	on " + VbCrlf
	sqlStr = sqlStr + " 		m.orderserial = T.orderserial " + VbCrlf
	sqlStr = sqlStr + " where " + VbCrlf
	sqlStr = sqlStr + " 	m.orderserial = '" & orderserial & "' " + VbCrlf

	dbget.Execute sqlStr

	sqlStr = "update [db_order].[dbo].tbl_order_master" + VbCrlf
	sqlStr = sqlStr + " set subtotalprice=totalsum-(IsNULL(tencardspend,0) + IsNULL(miletotalprice,0) + IsNULL(spendmembership,0) + IsNULL(allatdiscountprice,0)) "+ VbCrlf
	'sqlStr = sqlStr + " , subtotalpriceCouponNotApplied=subtotalpriceCouponNotApplied-(IsNULL(tencardspend,0) + IsNULL(miletotalprice,0) + IsNULL(spendmembership,0) + IsNULL(allatdiscountprice,0)) "+ VbCrlf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

    dbget.Execute sqlStr

	sqlStr = " update m "
	sqlStr = sqlStr + " set subtotalpriceCouponNotApplied = (case when T.dtotalitemcostCouponNotApplied = 0 then 0 else subtotalpriceCouponNotApplied end) "
	sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select "
	sqlStr = sqlStr + " 			orderserial, sum(IsNull(itemcostCouponNotApplied,0)*itemno) as dtotalitemcostCouponNotApplied "
	sqlStr = sqlStr + " 		from "
	sqlStr = sqlStr + " 			[db_order].[dbo].tbl_order_detail "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and orderserial = '" & orderserial & "' "
	sqlStr = sqlStr + " 			and cancelyn <> 'Y' "
	sqlStr = sqlStr + " 			and itemid <> 0 "
	sqlStr = sqlStr + " group by "
	sqlStr = sqlStr + " 	orderserial "
	sqlStr = sqlStr + " 	) T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial = T.orderserial "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	m.orderserial = '" & orderserial & "' "

	dbget.Execute sqlStr

end function



function updateUserMileage(byVal userid)
	dim sqlStr

	'==============================================================
	'보너스/사용마일리지 요약 재계산
'	sqlStr = "update [db_user].[dbo].tbl_user_current_mileage" + vbCrlf
'	sqlStr = sqlStr + " set [db_user].[dbo].tbl_user_current_mileage.spendmileage=IsNull(T.totspendmile,0)" + vbCrlf
'	sqlStr = sqlStr + " ,[db_user].[dbo].tbl_user_current_mileage.bonusmileage=IsNull(T.totgainmile,0)" + vbCrlf
'	sqlStr = sqlStr + " from " + vbCrlf
'	sqlStr = sqlStr + " ("
'	sqlStr = sqlStr + "     select sum(case when mileage<0 then mileage*-1 else 0 end) as totspendmile" + vbCrlf
'	sqlStr = sqlStr + "     , sum(case when mileage>0 then mileage else 0 end) as totgainmile" + vbCrlf
'	sqlStr = sqlStr + "     from [db_user].[dbo].tbl_mileagelog" + vbCrlf
'	sqlStr = sqlStr + "     where userid='" + userid + "'" + vbCrlf
'	sqlStr = sqlStr + "     and deleteyn='N'" + vbCrlf
'	sqlStr = sqlStr + " ) as T" + vbCrlf + vbCrlf
'	sqlStr = sqlStr + " where [db_user].[dbo].tbl_user_current_mileage.userid='" + userid + "'"
'	rsget.Open sqlStr,dbget,1

    ''2014/12/23 변경
    sqlStr = " exec [db_user].[dbo].sp_Ten_ReCalcu_His_BonusMileage '"&userid&"'"
    dbget.Execute sqlStr

	'==============================================================
	'주문마일리지 요약 재계산
'    sqlStr = "update [db_user].[dbo].tbl_user_current_mileage" + VbCrlf
'    sqlStr = sqlStr + " set [db_user].[dbo].tbl_user_current_mileage.jumunmileage=IsNull(T.totmile,0)" + VbCrlf
'    sqlStr = sqlStr + " from " + VbCrlf
'    sqlStr = sqlStr + "     (select sum(totalmileage) as totmile" + VbCrlf
'    sqlStr = sqlStr + "     from [db_order].[dbo].tbl_order_master" + VbCrlf
'    sqlStr = sqlStr + "     where userid='" + CStr(userid) + "' " + VbCrlf
'    sqlStr = sqlStr + "     and sitename ='10x10'" + VbCrlf
'    sqlStr = sqlStr + "     and cancelyn='N'" + VbCrlf
'    sqlStr = sqlStr + "     and ipkumdiv>3" + VbCrlf
'    sqlStr = sqlStr + " ) as T" + VbCrlf
'    sqlStr = sqlStr + " where [db_user].[dbo].tbl_user_current_mileage.userid='" + CStr(userid) + "' " + VbCrlf
'    rsget.Open sqlStr,dbget,1

    ''2014/12/23 변경
    sqlStr = " exec [db_order].[dbo].sp_Ten_recalcuHesJumunmileage '"&userid&"'"
    dbget.Execute sqlStr
end function

function updateUserDeposit(byVal userid)
	dim sqlStr
	dim dataexist

	'==============================================================
	'예치금 재계산
	sqlStr = " update c " + vbCrlf
	sqlStr = sqlStr + " set " + vbCrlf
	sqlStr = sqlStr + " 	c.currentdeposit = T.gaindeposit - T.spenddeposit " + vbCrlf
	sqlStr = sqlStr + " 	, c.gaindeposit = T.gaindeposit " + vbCrlf
	sqlStr = sqlStr + " 	, c.spenddeposit = T.spenddeposit " + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " 	db_user.dbo.tbl_user_current_deposit c " + vbCrlf
	sqlStr = sqlStr + " 	join ( " + vbCrlf
	sqlStr = sqlStr + " 		select " + vbCrlf
	sqlStr = sqlStr + " 			'" + userid + "' as userid " + vbCrlf
	sqlStr = sqlStr + " 			, IsNull(sum(case when deposit>0 then deposit else 0 end), 0) as gaindeposit " + vbCrlf
	sqlStr = sqlStr + " 			, IsNull(sum(case when deposit<0 then (deposit * -1) else 0 end), 0) as spenddeposit " + vbCrlf
	sqlStr = sqlStr + " 		from db_user.dbo.tbl_depositlog " + vbCrlf
	sqlStr = sqlStr + "     	where userid='" + userid + "'" + vbCrlf
	sqlStr = sqlStr + "     		and deleteyn='N' " + vbCrlf
	sqlStr = sqlStr + " 	) T " + vbCrlf
	sqlStr = sqlStr + " 	on " + vbCrlf
	sqlStr = sqlStr + " 		c.userid = T.userid " + vbCrlf
	'response.write sqlStr

	rsget.Open sqlStr,dbget

	sqlStr = " select @@rowcount as cnt "
	'response.write sqlStr

    rsget.Open sqlStr,dbget,1
        dataexist = (rsget("cnt") > 0)
    rsget.Close

	'데이타가 없으면 생성한다.
	if (Not dataexist) then

		sqlStr = " if not exists (select * from db_user.dbo.tbl_user_current_deposit where userid = '" + userid + "') begin " + vbCrlf
		sqlStr = sqlStr + " 	insert into db_user.dbo.tbl_user_current_deposit(userid, currentdeposit, gaindeposit, spenddeposit) " + vbCrlf
		sqlStr = sqlStr + " 		select " + vbCrlf
		sqlStr = sqlStr + " 			'" + userid + "' " + vbCrlf
		sqlStr = sqlStr + " 			, IsNull(sum(deposit), 0) as currentdeposit " + vbCrlf
		sqlStr = sqlStr + " 			, IsNull(sum(case when deposit>0 then deposit else 0 end), 0) as gaindeposit " + vbCrlf
		sqlStr = sqlStr + " 			, IsNull(sum(case when deposit<0 then (deposit * -1) else 0 end), 0) as spenddeposit " + vbCrlf
		sqlStr = sqlStr + " 		from db_user.dbo.tbl_depositlog " + vbCrlf
		sqlStr = sqlStr + "     	where userid='" + userid + "'" + vbCrlf
		sqlStr = sqlStr + " end " + vbCrlf

		dbget.Execute sqlStr
	end if

end function

function updateUserGiftCard(byVal userid)
	dim sqlStr
	dim dataexist

	'==============================================================
	'GiftCard 재계산
	sqlStr = " update c " + vbCrlf
	sqlStr = sqlStr + " set " + vbCrlf
	sqlStr = sqlStr + " 	c.currentCash = T.gainCash - T.spendCash - T.refundCash " + vbCrlf
	sqlStr = sqlStr + " 	, c.gainCash = T.gainCash " + vbCrlf
	sqlStr = sqlStr + " 	, c.spendCash = T.spendCash " + vbCrlf
	sqlStr = sqlStr + " 	, c.refundCash = T.refundCash " + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " 	db_user.dbo.tbl_giftcard_current c " + vbCrlf
	sqlStr = sqlStr + " 	join ( " + vbCrlf
	sqlStr = sqlStr + " 		select " + vbCrlf
	sqlStr = sqlStr + " 			'" + userid + "' as userid " + vbCrlf
	sqlStr = sqlStr + " 			, IsNull(sum(case when useCash>0 then useCash else 0 end), 0) as gainCash " + vbCrlf
	sqlStr = sqlStr + " 			, IsNull(sum(case when useCash<0 and (jukyocd not in ('400', '410', '900')) then (useCash * -1) else 0 end), 0) as spendCash " + vbCrlf
	sqlStr = sqlStr + " 			, IsNull(sum(case when useCash<0 and (jukyocd in ('400', '410', '900')) then (useCash * -1) else 0 end), 0) as refundCash " + vbCrlf
	sqlStr = sqlStr + " 		from db_user.dbo.tbl_giftcard_log " + vbCrlf
	sqlStr = sqlStr + "     	where userid='" + userid + "'" + vbCrlf
	sqlStr = sqlStr + "     		and deleteyn='N' " + vbCrlf
	sqlStr = sqlStr + " 	) T " + vbCrlf
	sqlStr = sqlStr + " 	on " + vbCrlf
	sqlStr = sqlStr + " 		c.userid = T.userid " + vbCrlf
	'response.write sqlStr

	rsget.Open sqlStr,dbget

	sqlStr = " select @@rowcount as cnt "
	'response.write sqlStr

    rsget.Open sqlStr,dbget,1
        dataexist = (rsget("cnt") > 0)
    rsget.Close

	'데이타가 없으면 생성한다.
	if (Not dataexist) then

		sqlStr = " if not exists (select * from db_user.dbo.tbl_giftcard_current where userid = '" + userid + "') begin " + vbCrlf
		sqlStr = sqlStr + " 	insert into db_user.dbo.tbl_giftcard_current(userid, currentCash, gainCash, spendCash, refundCash) " + vbCrlf
		sqlStr = sqlStr + " 		select " + vbCrlf
		sqlStr = sqlStr + " 			'" + userid + "' " + vbCrlf
		sqlStr = sqlStr + " 			, IsNull(sum(useCash), 0) as currentCash " + vbCrlf
		sqlStr = sqlStr + " 			, IsNull(sum(case when useCash>0 then useCash else 0 end), 0) as gainCash " + vbCrlf
		sqlStr = sqlStr + " 			, IsNull(sum(case when useCash<0 and (jukyocd not in ('400', '410', '900')) then (useCash * -1) else 0 end), 0) as spendCash " + vbCrlf
		sqlStr = sqlStr + " 			, IsNull(sum(case when useCash<0 and (jukyocd in ('400', '410', '900')) then (useCash * -1) else 0 end), 0) as refundCash " + vbCrlf
		sqlStr = sqlStr + " 		from db_user.dbo.tbl_giftcard_log " + vbCrlf
		sqlStr = sqlStr + "     	where userid='" + userid + "'" + vbCrlf
		sqlStr = sqlStr + " end " + vbCrlf

		dbget.Execute sqlStr
	end if

end function

''사용안함 - setCancelMaster로 통합
function LimitItemRecover(byval orderserial)
    dim sqlStr
    On Error Resume Next
        ''한정수량 조정 -
        sqlStr = "update [db_item].[dbo].tbl_item" + vbCrlf
        sqlStr = sqlStr + " set limitsold=(case when 0>limitsold - T.itemno then 0 else limitsold - T.itemno end)" + vbCrlf
        sqlStr = sqlStr + " from " + vbCrlf
        sqlStr = sqlStr + " ("
        sqlStr = sqlStr + " 	select d.itemid, d.itemno" + vbCrlf
        sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_detail d" + vbCrlf
        sqlStr = sqlStr + " 	where d.orderserial='" + CStr(orderserial) + "'" + vbCrlf
        sqlStr = sqlStr + " 	and d.itemid<>0 "
        sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
        sqlStr = sqlStr + " ) as T" + vbCrlf
        sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=T.Itemid"
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.limityn='Y'"

        dbget.Execute(sqlStr)

        ''옵션있는상품
        sqlStr = "update [db_item].[dbo].tbl_item_option" + vbCrlf
        sqlStr = sqlStr + " set optlimitsold=(case when 0 >optlimitsold - T.itemno then 0 else optlimitsold - T.itemno end)" + vbCrlf
        sqlStr = sqlStr + " from " + vbCrlf
        sqlStr = sqlStr + " ("
        sqlStr = sqlStr + " 	select d.itemid, d.itemoption, d.itemno" + vbCrlf
        sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_detail d " + vbCrlf
        sqlStr = sqlStr + " 	where d.orderserial='" + CStr(orderserial) + "'" + vbCrlf
        sqlStr = sqlStr + " 	and d.itemid<>0" + vbCrlf
        sqlStr = sqlStr + " 	and d.itemoption<>'0000'" + vbCrlf
        sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
        sqlStr = sqlStr + " ) as T" + vbCrlf
        sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid=T.Itemid"
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption"
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.optlimityn='Y'"

        dbget.Execute(sqlStr)
    On Error Goto 0
end function


sub UsafeCancel(byval orderserial)
    '// 전자보증서가 있으면 보증서 취소 요청 (2006.06.15; 운영관리팀 허진원)
    dim objUsafe, result, result_code, result_msg
    On Error Resume Next
    	Set objUsafe = CreateObject( "USafeCom.guarantee.1"  )

    '	Test일 때
    '	objUsafe.Port = 80
    '	objUsafe.Url = "gateway2.usafe.co.kr"
    '	objUsafe.CallForm = "/esafe/guartrn.asp"

        ' Real일 때
        objUsafe.Port = 80
        objUsafe.Url = "gateway.usafe.co.kr"
        objUsafe.CallForm = "/esafe/guartrn.asp"

    	objUsafe.gubun	= "B0"				'// 전문구분 (A0:신규발급, B0:보증서취소, C0:입금확인)
    	objUsafe.EncKey	= ""			'널값인 경우 암호화 안됨
    	objUsafe.mallId	= "ZZcube1010"		'// 쇼핑몰ID
    	objUsafe.oId	= CStr(orderserial)	'// 주문번호

    	'처리 실행!
    	result = objUsafe.cancelInsurance

    	result_code	= Left( result , 1 )
    	result_msg	= Mid( result , 3 )

    	Set objUsafe = Nothing
    On Error Goto 0
end Sub


'function ValidDeleteCS(id)
'    dim sqlStr
'    dim currstate
'
'    ValidDeleteCS = false
'
'    sqlStr = "select * from [db_cs].[dbo].tbl_new_as_list"
'    sqlStr = sqlStr + " where id=" + CStr(id)
'
'    rsget.Open sqlStr,dbget,1
'        currstate = rsget("currstate")
'    rsget.Close
'
'    If (currstate>="B006") then Exit function
'
'    ValidDeleteCS = true
'end function

'function DeleteCSProcess(id, finishuserid)
'    dim sqlStr, resultCount
'
'    sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
'    sqlStr = sqlStr + " set deleteyn='Y'" + VbCrlf
'    sqlStr = sqlStr + " , finishuser = '" + finishuserid+ "'" + VbCrlf
'    sqlStr = sqlStr + " , finishdate = getdate()" + VbCrlf
'    sqlStr = sqlStr + " where id=" + CStr(id)
'    sqlStr = sqlStr + " and currstate<'B006'"
'
'    dbget.Execute sqlStr, resultCount
'
'    DeleteCSProcess = (resultCount>0)
'end function

function GetRefundMethodString(returnmethod)
	dim tmpstr

    'R007 무통장환불
    'R020 실시간이체취소
    'R050 입점몰결제 취소
    'R080 올엣카드취소
    'R100 신용카드취소
    'R550 기프팅취소
    'R560 기프티콘취소
    'R120 신용카드부분취소
    'R400 휴대폰취소
	'R420 휴대폰부분취소
    'R900 마일리지로환불
    'R910 예치금환불
    'R022 실시간이체부분취소(NP)
    'R150 이니렌탈 취소

	tmpstr = ""

    if (returnmethod="R020") or (returnmethod="R022") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R550") or (returnmethod="R560") or (returnmethod="R120") or (returnmethod="R400") or (returnmethod="R420") then
        if (returnmethod="R020") then
            tmpstr = "실시간이체취소"
        elseif (returnmethod="R022") then ''2016/07/21
            tmpstr = "실시간이체부분취소"
        elseif (returnmethod="R080") then
            tmpstr = "올엣카드취소"
        elseif (returnmethod="R100") then
            tmpstr = "신용카드취소"
        elseif (returnmethod="R550") then
            tmpstr = "기프팅취소"
        elseif (returnmethod="R560") then
            tmpstr = "기프티콘취소"
        elseif (returnmethod="R120") then
            tmpstr = "신용카드부분취소"
		elseif (returnmethod="R400") then
            tmpstr = "휴대폰취소"
        elseif (returnmethod="R420") then
            tmpstr = "휴대폰부분취소"
        elseif (returnmethod="R150") then
            tmpstr = "이니렌탈취소"
        end if
    elseif (returnmethod="R050") then
        tmpstr = "입점몰결제 취소"
    elseif (returnmethod="R900") then
        tmpstr = "마일리지 환불"
    elseif (returnmethod="R910") then
        tmpstr = "예치금 환불"
    elseif (returnmethod<>"") then
        tmpstr = "무통장 환불"
    end if

	GetRefundMethodString = tmpstr

end function

function CheckNRegRefund(id, orderserial, reguserid)
    '' A003 환불요청 , A005 외부몰환불요청 , A007 신용카드/실시간이체취소요청
    '' Result -1, or newAsID
    dim divcd
    dim returnmethod, gubun01, gubun02

    dim sqlStr, RegDivCd
    dim title, contents_jupsu
    dim NewRegedID

    CheckNRegRefund = -1

    sqlStr = " select a.divcd, a.gubun01, a.gubun02"
    sqlStr = sqlStr + " , r.returnmethod "
    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list a"
    sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_as_refund_info r"
    sqlStr = sqlStr + "     on a.id=r.asid"
    sqlStr = sqlStr + " where a.id=" + CStr(id)
    sqlStr = sqlStr + " and a.deleteyn='N'"
    sqlStr = sqlStr + " and a.currstate<>'B007'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        divcd                = rsget("divcd")
        returnmethod         = rsget("returnmethod")
        gubun01              = rsget("gubun01")
        gubun02              = rsget("gubun02")

        if IsNULL(returnmethod) then returnmethod="R000"
    end if
    rsget.close


	'R000 환불없음.
    'R007 무통장환불
    'R020 실시간이체취소
    'R050 입점몰결제 취소
    'R080 올엣카드취소
    'R100 신용카드취소
    'R550 기프팅취소
    'R560 기프티콘취소
    'R120 신용카드부분취소
    'R400 휴대폰취소
	'R420 휴대폰부분취소
    'R900 마일리지로환불
    'R910 예치금환불
    'R022 실시간이체부분취소(NP)
    'R150 이니렌탈취소

	title = GetRefundMethodString(returnmethod)

    if (returnmethod="R000") or (returnmethod="") or (trim(returnmethod)="") then
        Exit function
    elseif (returnmethod="R020") or (returnmethod="R022") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R550") or (returnmethod="R560") or (returnmethod="R120") or (returnmethod="R400") or (returnmethod="R420") then
        RegDivCd = "A007"

        ''contents_jupsu = ""
    elseif (returnmethod="R050") then
        RegDivCd = "A005"
    elseif (returnmethod="R900") then
        RegDivCd = "A003"
    elseif (returnmethod="R910") then
        RegDivCd = "A003"
    elseif (returnmethod<>"") then
        RegDivCd = "A003"
        contents_jupsu = ""
    end if

    if (divcd="A008") then
        title = "주문 취소 후 " + title
    elseif (divcd="A004") then
        title = "반품 처리 후 " + title
    elseif (divcd="A010") then
        title = "회수 처리 후 " + title
    elseif (divcd="A100") then
        title = "교환 출고 후 " + title
    end if

    if (RegDivCd<>"") then
        NewRegedID =  RegCSMaster(RegDivCd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

		Call CopyWebCancelRefundInfo(id, NewRegedID)

        CheckNRegRefund = NewRegedID

		''원 CsID에 오픈메세지 저장
        Call AddCustomerOpenContents(id,title)
    end if
end function



function CheckNEditRefundInfo(id,returnmethod,rebankaccount,rebankownername,rebankname,paygateTid,refundrequire)
    dim sqlStr
    dim refundInfoExists, oldrefundrequire
    refundInfoExists     = false
    CheckNEditRefundInfo = false

    if ((returnmethod="") or (returnmethod="R000")) then Exit function
    if ((Not IsNumeric(refundrequire)) or (refundrequire="")) then Exit function


    sqlStr = " select * from [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " where asid=" + CStr(id)

    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
        refundInfoExists = True
        oldrefundrequire = rsget("refundrequire")
    end if
    rsget.Close

    if (Not refundInfoExists) then Exit function


    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info"                             + VbCrlf
    sqlStr = sqlStr + " set returnmethod='" + returnmethod + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankaccount='" + rebankaccount + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankownername='" + rebankownername + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankname='" + rebankname + "'"                          + VbCrlf
    sqlStr = sqlStr + " , paygateTid='" + paygateTid + "'"                          + VbCrlf

    ''무통장이나 마일리지 환불인 경우만 수기 수정 가능
    if ((returnmethod="R007") or (returnmethod="R900") or (returnmethod="R910")) and (refundrequire<>oldrefundrequire) then
        sqlStr = sqlStr + " , refundrequire=" + CStr(refundrequire)                     + VbCrlf
        sqlStr = sqlStr + " , refundadjustpay=" + CStr(refundrequire) + "-canceltotal"  + VbCrlf
    end if
    sqlStr = sqlStr + " where asid=" + CStr(id)

    dbget.Execute sqlStr

    CheckNEditRefundInfo = true
end function

%>
