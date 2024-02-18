<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/outmall/lotteCom/lotteitemcls.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<%
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.9","61.252.133.10","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72", "61.252.133.67", "61.252.133.70","121.78.103.60")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

dim act     : act = requestCheckVar(request("act"),32)
dim param1  : param1 = requestCheckVar(request("param1"),32)
dim param2  : param2 = requestCheckVar(request("param2"),32)
dim param3  : param3 = requestCheckVar(request("param3"),32)
dim param4  : param4 = requestCheckVar(request("param4"),32)
dim sqlStr, i, paramData, retVal
dim retCnt : retCnt = 0

dim ref : ref = Request.ServerVariables("REMOTE_ADDR")
dim redKey  : redKey = requestCheckVar(request("redSsnKey"),32)
If (application("Svr_Info")) <> "Dev" Then
    if (Not CheckVaildIP(ref)) then
        If (param1 = "lotteon" or param1 = "shintvshopping" or param1 = "skstoa" or param1 = "wetoo1300k" or param1 = "lotteimall" or param1 = "coupang" or param1 = "nvstorefarm" or param1 = "nvstoremoonbangu" or param1 = "Mylittlewhoopee" or param1 = "nvstoregift" or param1 = "hmall1010" or param1 = "11st1010" or param1 = "yes24" or param1 = "alphamall" or param1 = "kakaostore" or param1 = "ohou1010" or param1 = "wadsmartstore" or param1 = "casamia_good_com" or param1 = "wconcept1010" or param1 = "withnature1010" or param1 = "goodshop1010" or param1 = "auction1010" or param1 = "lfmall" or param1 = "gmarket1010") and redKey = "system" Then
        Else
            rw ref
            dbget.Close()
            response.end
        End If
    end if
End If

Dim cnt, topCnt
Dim OutMallOrderSerialArr
Dim OrgDetailKeyArr
Dim songjangDivArr
Dim songjangNoArr, sendReqCntArr, beasongdateArr, outmallGoodsIDArr, orderItemOptionArr, orderItemOptionNameArr, outmalloptionnoArr, requireDetail11stYNArr, beasongNum11stArr, reserve01Arr, shoplinkerOrderIDArr
dim oLotteitem, itemidArr, orgsendReqCntArr

If param1 = "nvstorefarm" Then
    topCnt = "50"
Else
    topCnt = "100"
End If


select Case act

    Case "outmallSongJangIp" ''제휴사 송장입력
    'response.end

        sqlStr = "select top "& topCnt &" T.orderserial, T.OutMallOrderSerial"
        sqlStr = sqlStr & " ,T.OrgDetailKey, IsNULL(T.sendState,0) as sendState"
        sqlStr = sqlStr & " ,D.songjangDiv, isNULL(D.songjangNo,'기타') as songjangNo, D.itemNo, D.beasongdate, T.outMallGoodsNo, T.orderItemOption, T.orderItemOptionName " ''isNULL(D.songjangNo,'기타') as songjangNo 2015/10/09
		sqlStr = sqlStr & " ,T.outmalloptionno, T.requireDetail11stYN, T.beasongNum11st, T.orgOrderCNT, T.reserve01, T.shoplinkerOrderID "
        sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T with (nolock) "
        sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master M with (nolock) "
        sqlStr = sqlStr & " 	on T.orderserial=M.orderserial"
        sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_detail D with (nolock) "
        sqlStr = sqlStr & " 	on T.orderserial=D.orderserial"
        ''sqlStr = sqlStr & " 	and T.matchitemid=D.itemid"
        ''sqlStr = sqlStr & " 	and T.matchitemoption=D.itemoption"
		sqlStr = sqlStr & " 	and IsNull(T.changeitemid, T.matchitemid)=D.itemid"					'// 기존 주문에 합쳐진 경우(빨강1개,파랑1개 -> 파랑2개)
		sqlStr = sqlStr & " 	and IsNull(T.changeitemoption, T.matchitemoption)=D.itemoption"
        sqlStr = sqlStr & " 	and D.currstate=7"
        sqlStr = sqlStr & " 	left join db_order.dbo.tbl_songjang_div V with (nolock) "
        sqlStr = sqlStr & " 	on D.songjangDiv=V.divcd"
'        sqlStr = sqlStr & " where datediff(m,T.regdate,getdate())<7"    ''20130304 추가
        sqlStr = sqlStr & " where T.regdate > dateadd(month, -2, getdate()) "    ''7개월 -> 2개월로 변경..2021-11-18 김진영
        sqlStr = sqlStr & " and T.sellsite='"&param1&"'"
        sqlStr = sqlStr & " and T.OrgDetailKey is Not NULL"             ''디테일키 입력 주문건만..
        sqlStr = sqlStr & " and IsNULL(T.sendState,0)=0"
        sqlStr = sqlStr & " and T.sendReqCnt<3"                         ''여러번 시도 안되도록. 추가.
        sqlStr = sqlStr & " and T.matchState not in ('R','D','B')"      ''교환 취소 반품 제외.
        sqlStr = sqlStr & " order by D.beasongdate desc"

        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        cnt = rsget.RecordCount
        ReDim OutMallOrderSerialArr(cnt)
        ReDim OrgDetailKeyArr(cnt)
        ReDim songjangDivArr(cnt)
        ReDim songjangNoArr(cnt)
        Redim sendReqCntArr(cnt)
        Redim orgsendReqCntArr(cnt)
        Redim beasongdateArr(cnt)
        Redim outmallGoodsIDArr(cnt)
		Redim orderItemOptionArr(cnt)
		Redim orderItemOptionNameArr(cnt)
		Redim outmalloptionnoArr(cnt)
		Redim requireDetail11stYNArr(cnt)
		Redim beasongNum11stArr(cnt)
        Redim reserve01Arr(cnt)
        Redim shoplinkerOrderIDArr(cnt)

        i = 0
        if Not rsget.Eof then
            do until rsget.eof
            OutMallOrderSerialArr(i) = rsget("OutMallOrderSerial")
            OrgDetailKeyArr(i) = rsget("OrgDetailKey")
			songjangDivArr(i) = rsget("songjangDiv")
			songjangNoArr(i) = Trim(replace(rsget("songjangNo"),"-",""))
			sendReqCntArr(i) = rsget("itemNo")
            orgsendReqCntArr(i) = rsget("orgOrderCNT")
			beasongdateArr(i) = rsget("beasongdate")
			outmallGoodsIDArr(i) = rsget("outMallGoodsNo")
			orderItemOptionArr(i) = rsget("orderItemOption")
			orderItemOptionNameArr(i) = rsget("orderItemOptionName")
			outmalloptionnoArr(i)	= rsget("outmalloptionno")
			requireDetail11stYNArr(i)= rsget("requireDetail11stYN")
			beasongNum11stArr(i) = rsget("beasongNum11st")
            reserve01Arr(i) = rsget("reserve01")
            shoplinkerOrderIDArr(i) = rsget("shoplinkerOrderID")
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
                    IF (LCASE(param1)="lottecom") then
                        paramData = "redSsnKey=system&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode2LotteDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)
                        ''response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/LotteCom_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="homeplus") then
                        paramData = "redSsnKey=system&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&Server.URLEncode(TenDlvCode2HomeplusDlvCode(songjangDivArr(i)))&"&inv_no="&songjangNoArr(i)
                        ''response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/Homeplus_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="lfmall") then
                        paramData = "redSsnKey=system&OutMallOrderSerial="&OutMallOrderSerialArr(i)&"&OrgDetailKey="&OrgDetailKeyArr(i)&"&hdc_cd="&songjangDivArr(i)&"&songjangNo="&songjangNoArr(i)
                        ''response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/lfmall_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                        
                    ELSEIF (LCASE(param1)="yes24") or (LCASE(param1)="alphamall") or (LCASE(param1)="ohou1010") or (LCASE(param1)="wadsmartstore") or (LCASE(param1)="casamia_good_com") or (LCASE(param1)="wconcept1010") or (LCASE(param1)="withnature1010") or (LCASE(param1)="goodshop1010") then    '사방넷의 lfmall or yes24 or alphamall or ohou1010 or wadsmartstore or casamia_good_com or wconcept1010 or withnature1010 or goodshop1010
                        paramData = "redSsnKey=system&OutMallOrderSerial="&OutMallOrderSerialArr(i)&"&OrgDetailKey="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode2SabangNetDlvCode(songjangDivArr(i))&"&songjangNo="&songjangNoArr(i)&"&shoplinkerorderid="&shoplinkerOrderIDArr(i)
                        response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/sabangnet_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal

                    ELSEIF (LCASE(param1)="auction1010") then
                        paramData = "redSsnKey=system&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode2AuctionDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)&"&songjangDiv="&songjangDivArr(i)
                        ''response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/Auction_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="gmarket1010") then
                        paramData = "redSsnKey=system&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&Server.URLEncode(TenDlvCode2GmarketDlvCode(songjangDivArr(i)))&"&inv_no="&songjangNoArr(i)&"&songjangDiv="&songjangDivArr(i)
                        ''response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/Gmarket_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="halfclub") then
                        paramData = "redSsnKey=system&OutMallOrderSerial="&OutMallOrderSerialArr(i)&"&OrgDetailKey="&OrgDetailKeyArr(i)&"&outmallGoodNo="&outmallGoodsIDArr(i)&"&hdc_cd="&Server.URLEncode(TenDlvCode2HalfClubDlvCode(songjangDivArr(i)))&"&songjangNo="&songjangNoArr(i)&"&songjangDiv="&songjangDivArr(i)&"&itemno="&orgsendReqCntArr(i)&"&outmallOptionCode="&orderItemOptionArr(i)&"&outmallOptionName="&orderItemOptionNameArr(i)
                        ''response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/halfclub_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal


                    ELSEIF (LCASE(param1)="coupang") then
                        paramData = "redSsnKey=system&OutMallOrderSerial="&OutMallOrderSerialArr(i)&"&OrgDetailKey="&OrgDetailKeyArr(i)&"&outmallGoodNo="&outmallGoodsIDArr(i)&"&hdc_cd="&Server.URLEncode(TenDlvCode2CoupangDlvCode(songjangDivArr(i)))&"&songjangNo="&songjangNoArr(i)&"&songjangDiv="&songjangDivArr(i)&"&outmallOptionCode="&outmalloptionnoArr(i)&"&beasongNum="&beasongNum11stArr(i)&"&splitrequire="&requireDetail11stYNArr(i)
                        'response.write paramData&"<br>"
                    	'response.end
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/coupang_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="lotteon") then
                        paramData = "redSsnKey=system&OutMallOrderSerial="&OutMallOrderSerialArr(i)&"&OrgDetailKey="&OrgDetailKeyArr(i)&"&outmallGoodNo="&outmallGoodsIDArr(i)&"&hdc_cd="&Server.URLEncode(TenDlvCode2LotteonDlvCode(songjangDivArr(i)))&"&songjangNo="&songjangNoArr(i)&"&outmallOptionCode="&outmalloptionnoArr(i)&"&beasongNum="&beasongNum11stArr(i)&"&sendQnt="&orgsendReqCntArr(i)
                        'response.write paramData&"<br>"
                    	'response.end
                        if (application("Svr_Info")<>"Dev") then
                            retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/Lotteon_SongjangProc.asp",paramData)
                        else
                            retVal = SendReq("http://localhost:11117/outmall/proc/Lotteon_SongjangProc.asp",paramData)
                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="shintvshopping") then
                        paramData = "redSsnKey=system&OutMallOrderSerial="&OutMallOrderSerialArr(i)&"&OrgDetailKey="&OrgDetailKeyArr(i)&"&outmallGoodNo="&outmallGoodsIDArr(i)&"&hdc_cd="&Server.URLEncode(TenDlvCode2ShintvshoppingDlvCode(songjangDivArr(i)))&"&songjangNo="&songjangNoArr(i)&"&outmallOptionCode="&outmalloptionnoArr(i)&"&beasongNum="&beasongNum11stArr(i)
                        'response.write paramData&"<br>"
                    	'response.end
                        if (application("Svr_Info")<>"Dev") then
                            retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/shintvshopping_SongjangProc.asp",paramData)
                        else
                            retVal = SendReq("http://localhost:11117/outmall/proc/shintvshopping_SongjangProc.asp",paramData)
                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="skstoa") then
                        paramData = "redSsnKey=system&OutMallOrderSerial="&OutMallOrderSerialArr(i)&"&OrgDetailKey="&OrgDetailKeyArr(i)&"&outmallGoodNo="&outmallGoodsIDArr(i)&"&hdc_cd="&Server.URLEncode(TenDlvCode2SkstoaDlvCode(songjangDivArr(i)))&"&songjangNo="&songjangNoArr(i)&"&outmallOptionCode="&outmalloptionnoArr(i)
                        'response.write paramData&"<br>"
                    	'response.end
                        if (application("Svr_Info")<>"Dev") then
                            retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/skstoa_SongjangProc.asp",paramData)
                        else
                            retVal = SendReq("http://localhost:11117/outmall/proc/skstoa_SongjangProc.asp",paramData)
                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="wetoo1300k") then
                        paramData = "redSsnKey=system&OutMallOrderSerial="&OutMallOrderSerialArr(i)&"&OrgDetailKey="&OrgDetailKeyArr(i)&"&outmallGoodNo="&outmallGoodsIDArr(i)&"&hdc_cd="&Server.URLEncode(TenDlvCode2Wetoo1300kDlvCode(songjangDivArr(i)))&"&songjangNo="&songjangNoArr(i)
                        'response.write paramData&"<br>"
                    	'response.end
                        if (application("Svr_Info")<>"Dev") then
                            retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/wetoo1300k_SongjangProc.asp",paramData)
                        else
                            retVal = SendReq("http://localhost:11117/outmall/proc/wetoo1300k_SongjangProc.asp",paramData)
                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="11st1010") then
                        paramData = "redSsnKey=system&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode211stDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)&"&songjangDiv="&beasongNum11stArr(i)
                        'response.write paramData&"<br>"
                    	'response.end
                        if (application("Svr_Info")<>"Dev") then
                            retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/11st_SongjangProc.asp",paramData)
                        else
                            retVal = SendReq("http://localhost:11117/outmall/proc/11st_SongjangProc.asp",paramData)
                        end if
                        response.write retVal

                    ELSEIF (LCASE(param1)="hmall1010") then
                        paramData = "redSsnKey=system&OutMallOrderSerial="&OutMallOrderSerialArr(i)&"&OrgDetailKey="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode2HmallDlvCode(songjangDivArr(i))&"&songjangNo="&songjangNoArr(i)&"&beasongNum="&beasongNum11stArr(i)&"&reserve01="&reserve01Arr(i)
                        'response.write paramData&"<br>"
                    	'response.end
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/hmall_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="nvstorefarm") then
                        paramData = "redSsnKey=system&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode2NvstorefarmDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)&"&songjangDiv="&songjangDivArr(i)
                        ''response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/Nvstorefarm_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="nvstoremoonbangu") then
                        paramData = "redSsnKey=system&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode2NvstorefarmDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)&"&songjangDiv="&songjangDivArr(i)
                        ''response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/Nvstoremoonbangu_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="mylittlewhoopee") then
                        paramData = "redSsnKey=system&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode2NvstorefarmDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)&"&songjangDiv="&songjangDivArr(i)
                        ''response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/Mylittlewhoopee_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="nvstoregift") then
                        paramData = "redSsnKey=system&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode2NvstorefarmDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)&"&songjangDiv="&songjangDivArr(i)
                        ''response.write paramData&"<br>"
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/Nvstoregift_SongjangProc.asp",paramData)
                        else

                        end if
                        response.write retVal
                    ELSEIF (LCASE(param1)="lotteimall") then
                        paramData = "redSsnKey=system&cmdparam=songjangip&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&sendQnt="&sendReqCntArr(i)&"&sendDate="&replace(Left(beasongdateArr(i),10),"-","")&"&outmallGoodsID="&outmallGoodsIDArr(i)&"&hdc_cd="&TenDlvCode2LotteiMallNewDlvCode(songjangDivArr(i))&"&inv_no="&songjangNoArr(i)
                        'rw paramData
                        'response.end
                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/Lotteimall_SongjangProc.asp",paramData)
                             'rw retVal
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
