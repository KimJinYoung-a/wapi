<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/outmall/cjmall/cjmallitemcls.asp"-->
<!-- #include virtual="/outmall/cjmall/incCJmallFunction.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<%
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

' if (Not CheckVaildIP(ref)) then
'     dbget.Close()
'     response.end
' end if

dim act     : act = requestCheckVar(request("act"),32)
dim param1  : param1 = requestCheckVar(request("param1"),32)
dim param2  : param2 = requestCheckVar(request("param2"),32)
dim param3  : param3 = requestCheckVar(request("param3"),32)
dim param4  : param4 = requestCheckVar(request("param4"),32)
dim param5	: param5 = requestCheckVar(request("param5"),32)
dim sqlStr, i, paramData, retVal
dim retCnt : retCnt = 0

Dim cnt
Dim OutMallOrderSerialArr
Dim OrgDetailKeyArr
Dim songjangDivArr
Dim songjangNoArr, sendReqCntArr, beasongdateArr, outmallGoodsIDArr
Dim cjMall, itemidArr, ArrRows
select Case act
    Case "outmallSongJangIp" ''제휴사 송장입력	40=>5*N	2016/04/05			==================================================================
    'response.end
        sqlStr = "select top 30 T.orderserial, T.OutMallOrderSerial"
        sqlStr = sqlStr & " ,T.OrgDetailKey, IsNULL(T.sendState,0) as sendState"
        sqlStr = sqlStr & " ,D.songjangDiv, D.songjangNo, D.itemNo, D.beasongdate, T.outMallGoodsNo"
        sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T"
        sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master M"
        sqlStr = sqlStr & " 	on T.orderserial=M.orderserial"
        sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_detail D"
        sqlStr = sqlStr & " 	on T.orderserial=D.orderserial"
        ''sqlStr = sqlStr & " 	and T.matchitemid=D.itemid"
        ''sqlStr = sqlStr & " 	and T.matchitemoption=D.itemoption"
		sqlStr = sqlStr & " 	and IsNull(T.changeitemid, T.matchitemid)=D.itemid"					'// 기존 주문에 합쳐진 경우(빨강1개,파랑1개 -> 파랑2개)
		sqlStr = sqlStr & " 	and IsNull(T.changeitemoption, T.matchitemoption)=D.itemoption"
        sqlStr = sqlStr & " 	and D.currstate=7"
        sqlStr = sqlStr & " 	left join db_order.dbo.tbl_songjang_div V"
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
        ReDim TenOrderserial(cnt)
        ReDim OutMallOrderSerialArr(cnt)
        ReDim OrgDetailKeyArr(cnt)
        ReDim songjangDivArr(cnt)
        ReDim songjangNoArr(cnt)
        Redim sendReqCntArr(cnt)
        Redim beasongdateArr(cnt)
        Redim outmallGoodsIDArr(cnt)
        i = 0
        if Not rsget.Eof then
            do until rsget.eof
            TenOrderserial(i) = rsget("orderserial")
            OutMallOrderSerialArr(i) = rsget("OutMallOrderSerial")
            OrgDetailKeyArr(i) = rsget("OrgDetailKey")
			songjangDivArr(i) = rsget("songjangDiv")
			songjangNoArr(i) = rsget("songjangNo")
			sendReqCntArr(i) = rsget("itemNo")
			beasongdateArr(i) = rsget("beasongdate")
			outmallGoodsIDArr(i) = rsget("outMallGoodsNo")
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
					IF (LCASE(param1)="cjmall") then
                        ''var params = "ten_ord_no="+tenorderserial+"&ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+OutMallSDiv+"&inv_no="+songjangNo;
                        ''var popwin=window.open('/admin/etc/cjmall/actCJmallSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');

                        paramData = "redSsnKey=system&ten_ord_no="&TenOrderserial(i)&"&ord_no="&OutMallOrderSerialArr(i)&"&ord_dtl_sn="&OrgDetailKeyArr(i)&"&hdc_cd="&TenDlvCode2cjMallDlvCode(songjangDivArr(i))&"&inv_no="&server.URLEncode(songjangNoArr(i))
                        rw paramData

                        if (application("Svr_Info")<>"Dev") then
                             retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCJmallSongjangInputProc.asp",paramData)
                             rw retVal
                        end if
                    end if
                end if
            next
        end if


'    Case "cjmallCheckRDItem"				''cjmall 승인,판매상태 확인 Batch			==================================================================
'        paramData = "redSsnKey=system&cmdparam=confirmItemAuto&subcmd="&param2
'        retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCjMallReq.asp",paramData)
'        response.write retVal&VbCRLF
'	Case "cjmallSoldOutItem"				''품절처리 상품. (10x10 품절, cj판매중)		==================================================================
'		Set cjMall = new CCjmall
'			cjMall.FCurrPage					= 1
'			cjMall.FPageSize					= 30
'			cjMall.FRectExtNotReg				= "D"		'전시
'			cjMall.FRectCjmallYes10x10No	    = "on"
'			cjMall.getCjmallRegedItemList
'
'		If (cjMall.FResultCount < 1) Then
'			response.Write "S_NONE"
'			dbCTget.Close()
'			Set cjMall = Nothing
'			response.end
'		End If
'
'		For i = 0 to cjMall.FResultCount - 1
'		    itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
'		Next
'		Set cjMall= Nothing
'		IF (Right(itemidArr,1) = ",") Then itemidArr = Left(itemidArr, Len(itemidArr) - 1)
'		paramData = "redSsnKey=system&cmdparam=EditSellYn&subcmd=N&cksel="&itemidArr
'		retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCjMallReq.asp",paramData)
'		response.Write "itemidArr="&itemidArr
'		response.Write "<br>"&retVal
'	Case "cjmallexpensive10x10"				'CJ 가격 < 텐바이텐 가격					==================================================================
'		Set cjMall = new CCjmall
'			cjMall.FCurrPage					= 1
'			cjMall.FPageSize					= 20
'			cjMall.FRectExtNotReg				= "D"	'전시
'			cjMall.FRectSellYn					= "Y"
'			cjMall.FRectExtSellYn               = "Y"
'			cjMall.FRectExpensive10x10          = "on"
'			cjMall.FRectOrdType					= "B"	'베스트순
'			cjMall.FRectFailCntOverExcept		= "3"	'3회 이상 실패내역 제낌.
'			cjMall.getCjmallRegedItemList
'
'		If (cjMall.FResultCount < 1) Then
'			response.Write "S_NONE"
'			dbCTget.Close()
'			Set cjMall = Nothing
'			response.end
'		End If
'
'		For i = 0 to cjMall.FResultCount - 1
'			itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
'		Next
'		Set cjMall = Nothing
'		If (Right(itemidArr,1) = ",") Then itemidArr = Left(itemidArr, Len(itemidArr) - 1)
'		paramData = "redSsnKey=system&cmdparam=EditPriceSelect&cksel="&itemidArr
'		retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCjMallReq.asp",paramData)
'		response.Write "itemidArr="&itemidArr
'		response.Write "<br>"&retVal
'	Case "cjmallmarginItem" '' 역마진 상품이면(보통 할인들어간 것) sellcash를 orgprice로==================================================================
'		sqlStr = ""
'		sqlStr = sqlStr & " SELECT TOP 10 "
'		sqlStr = sqlStr & " i.itemid, i.itemname, i.orgPrice, i.sellcash, i.buycash , J.cjmallPrice, J.cjmallSellYn, (i.buycash)/J.cjmallPrice*100 as margin "
'		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_cjmall_regitem as J "
'		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item as i on J.itemid = i.itemid "
'		sqlStr = sqlStr & " WHERE 1 = 1 and i.isusing='Y' "
'		sqlStr = sqlStr & " and i.sellcash >= 1000 "
'		sqlStr = sqlStr & " and J.cjmallStatCd = 7 and J.cjmallPrdNo is Not Null and i.sellYn='Y' "
'		sqlStr = sqlStr & " and J.cjmallSellYn='Y' and i.sellcash<>0 "
'		sqlStr = sqlStr & " and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)<15 "
'		sqlStr = sqlStr & " and J.cjmallSellYn= 'Y' "
'		sqlStr = sqlStr & " and i.orgprice <> J.cjmallPrice "
'		sqlStr = sqlStr & " and J.accFailCNT < 5 "
'		sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
'        rsCTget.Open sqlStr,dbCTget,1
'        cnt = rsCTget.RecordCount
'		If (cnt < 1) Then
'			response.Write "S_NONE"
'			response.end
'		Else
'	        For i = 0 to cnt - 1
'	            itemidArr = itemidArr & rsCTget("itemid") &","
'				rsCTget.MoveNext
'	        Next
'		End If
'		rsCTget.Close
'
'		IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
'		paramData = "redSsnKey=system&cmdparam=EditPriceSelect&cksel="&itemidArr
'		retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCjMallReq.asp",paramData)
'		response.Write "itemidArr="&itemidArr
'		response.Write "<br>"&retVal
'	Case "cjmallmarginNotSaleItem" ''역마진 중에서 세일N인 것들 품절처리				==================================================================
'		Set cjMall = new CCjmall
'			cjMall.FCurrPage					= 1
'			cjMall.FPageSize					= 10
'			cjMall.FRectExtNotReg				= "D"	'전시
'			cjMall.FRectSellYn					= "Y"	'판매Y
'			cjMall.FRectSailYn					= "N"	'할인N
'			cjMall.FRectonlyValidMargin			= "N"	'마진이하
'			cjMall.FRectExtSellYn               = "Y"	'제휴판매Y
'			cjMall.getCjmallRegedItemList
'
'		If (cjMall.FResultCount < 1) Then
'			response.Write "S_NONE"
'			dbCTget.Close()
'			Set cjMall = Nothing
'			response.end
'		End If
'
'		For i = 0 to cjMall.FResultCount - 1
'			itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
'		Next
'		Set cjMall= Nothing
'
'		IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
'		paramData = "redSsnKey=system&cmdparam=EditSellYn&subcmd=N&cksel="&itemidArr
'		retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCjMallReq.asp",paramData)
'		response.Write "itemidArr="&itemidArr
'		response.Write "<br>"&retVal
'	Case "cjmallEditItem"	'cj 상품수정												==================================================================
'		Set cjMall = new CCjmall
'		cjMall.FCurrPage					= param2
'		cjMall.FPageSize					= 5
'		cjMall.FRectExtNotReg				= param5
'		cjMall.FRectMatchCate				= "Y"
'		cjMall.FRectPrdDivMatch				= "Y"
'		cjMall.FRectSellYn					= "Y"
'		cjMall.FRectOrdType					= param3	'베스트 셀러순 "B"
'		If param4 <> "" Then							'한정판매
'			cjMall.FRectLimitYn = "Y"
'		Else
'			cjMall.FRectonlyValidMargin = "Y"			'마진 되는거만
'		End If
'		cjMall.FRectFailCntOverExcept			= "5"
'		cjMall.getCjmallRegedItemList
'
'		If (cjMall.FResultCount < 1) Then
'			response.Write "S_NONE"
'			dbCTget.Close()
'			Set cjMall = Nothing
'			response.end
'		End If
'
'		For i = 0 to cjMall.FResultCount - 1
'			itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
'		Next
'		Set cjMall = Nothing
'
'		IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
'		paramData = "redSsnKey=system&cmdparam=EditSelect2&cksel="&itemidArr
'		retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCjMallReq.asp",paramData)
'		response.Write "itemidArr="&itemidArr
'		response.Write "<br>"&retVal
'    Case "cjmallExpireItem"   '' 품절 처리 요망 (사용안함, 조건배송등)					==================================================================
'		Set cjMall = new CCjmall
'			cjMall.FCurrPage					= 1
'			cjMall.FPageSize					= param2
'			cjMall.FRectExtNotReg				= "D"
'			cjMall.FRectExtSellYn               = "Y"
'			cjMall.FRectFailCntOverExcept		= "3"	'3회 이상 실패내역 제낌.
'			cjMall.getCjmallreqExpireItemList
'
'		If (cjMall.FResultCount < 1) Then
'			response.Write "S_NONE"
'			dbCTget.Close()
'			Set cjMall = Nothing
'			response.end
'		End If
'
'		For i = 0 to cjMall.FResultCount - 1
'			itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
'		Next
'		Set cjMall = Nothing
'
'		IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
'		paramData = "redSsnKey=system&cmdparam=EditSellYn&subcmd=N&cksel="&itemidArr
'		retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCjMallReq.asp",paramData)
'		response.Write "itemidArr="&itemidArr
'		response.Write "<br>"&retVal
'	Case "cjmallEditItemLastupdate"						'상품최종수정일 기준 상품수정	==================================================================
'		Set cjMall = new CCjmall
'			cjMall.FPageSize				= 5
'			cjMall.FCurrPage				= 1
'			cjMall.FRectExtNotReg			= "D"
'			cjMall.FRectMatchCate			= "Y"
'			cjMall.FRectSellYn				= "Y"
'			cjMall.FRectExtSellYn			= "Y"
'			cjMall.FRectOrdType				= "LU"		'LU는 조건 -> isnull(J.lastStatCheckDate,'') = '' and Left(i.lastupdate, 10) <> Left(J.cjmallLastUpdate, 10) | order by i.lastupdate
'			cjMall.FRectFailCntOverExcept	= "3"
'			cjMall.getCjmallRegedItemList
'			If (cjMall.FResultCount < 1) Then
'				response.Write "S_NONE"
'				dbCTget.Close()
'				Set cjMall= Nothing
'				response.end
'			End If
'
'			For i = 0 to cjMall.FResultCount - 1
'				itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
'			Next
'
'			Set cjMall= Nothing
'			IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
'			paramData = "redSsnKey=system&cmdparam=EditSelect2&cksel="&itemidArr
'			retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCjMallReq.asp",paramData)
'			response.Write "itemidArr="&itemidArr
'			response.Write "<br>"&retVal
'	Case "cjmallpriceDiff"	'가격 상이 수정												==================================================================
'		Set cjMall = new CCjmall
'			cjMall.FCurrPage					= 1
'			cjMall.FPageSize					= 10
'			cjMall.FRectExtNotReg				= "D"	'전시
'			cjMall.FRectSellYn					= "Y"	'판매Y
'			cjMall.FRectonlyValidMargin 		= "Y"	'마진이상
'			cjMall.FRectFailCntOverExcept		= "3"	'3회 재낌
'			cjMall.FRectdiffPrc 				= "Y"	'가격 다름 조건Y
'			cjMall.GetCjmallRegedItemList
'
'			If (cjMall.FResultCount < 1) Then
'				response.Write "S_NONE"
'				dbCTget.Close()
'				Set cjMall= Nothing
'				response.end
'			End If
'
'			For i = 0 to cjMall.FResultCount - 1
'				itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
'			Next
'
'			IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
'			paramData = "redSsnKey=system&cmdparam=EditPriceSelect&cksel="&itemidArr
'			retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCjMallReq.asp",paramData)
'			response.Write "itemidArr="&itemidArr
'			response.Write "<br>"&retVal
'    Case "cjmallpriceEdit2" '' 가격 수정(옵션추가금액 관련)								==================================================================
'		sqlStr = ""
'		sqlStr = sqlStr & " select top 10 ro.itemid, r.lastStatCheckDate"
'		sqlStr = sqlStr & " from db_outmall.dbo.tbl_cjmall_regItem r"
'		sqlStr = sqlStr & " Join db_outmall.dbo.tbl_outmall_regedoption ro on ro.itemid=r.itemid"
'		sqlStr = sqlStr & " Join db_AppWish.dbo.tbl_item_option o on ro.itemid=o.itemid and ro.itemoption=o.itemoption"
'		sqlStr = sqlStr & " Join db_AppWish.dbo.tbl_item i on r.itemid=i.itemid"
'		sqlStr = sqlStr & " where ro.mallid='cjmall'"
'		sqlStr = sqlStr & " and r.optaddPrcCnt>0"
'		sqlStr = sqlStr & " and r.cjmallprice+o.optAddprice<>ro.outmallAddPrice"
'		sqlStr = sqlStr & " and r.cjmallsellyn='Y'"
'		sqlStr = sqlStr & " and r.accFailCNT < 5 "
'		sqlStr = sqlStr & " group by ro.itemid,r.lastStatCheckDate "
'		sqlStr = sqlStr & " order by r.lastStatCheckDate"
'		rsCTget.Open sqlStr,dbCTget,1
'		If not rsCTget.Eof Then
'			ArrRows = rsCTget.getRows()
'		End If
'		rsCTget.close
'
'		itemidArr = ""
'		If isArray(ArrRows) then
'			For i =0 To UBound(ArrRows,2)
'				itemidArr = itemidArr + CStr(ArrRows(0,i)) + ","
'			Next
'		Else
'			rw "S_NONE"
'			dbCTget.Close() : response.end
'		End If
'
'		IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
'		paramData = "redSsnKey=system&cmdparam=EditSelect2&cksel="&itemidArr
'		retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCjMallReq.asp",paramData)
'		response.Write "itemidArr="&itemidArr
'		response.Write "<br>"&retVal
'    Case "cjmallpriceEdit3" '' 가격 수정(2014-12-19 진영 추가)							==================================================================
'		sqlStr = ""
'		sqlStr = sqlStr & " select top 10 itemid from "
'		sqlStr = sqlStr & " db_outmall.dbo.tbl_cjmall_regitem "
'		sqlStr = sqlStr & " where optAddPrcCnt > 0  "
'		sqlStr = sqlStr & " and cjmallsellyn = 'Y' "
'		sqlStr = sqlStr & " and cjmallStatCd = '7' "
'		sqlStr = sqlStr & " and isnull(cjmallprdno, '') <> '' "
'		sqlStr = sqlStr & " and accFailCNT < 5 "
'		sqlStr = sqlStr & " order by lastpriceCheckDate asc "
'		rsCTget.Open sqlStr,dbCTget,1
'		if not rsCTget.Eof then
'			ArrRows = rsCTget.getRows()
'		end if
'		rsCTget.close
'
'		itemidArr = ""
'		If isArray(ArrRows) then
'			For i =0 To UBound(ArrRows,2)
'				itemidArr = itemidArr + CStr(ArrRows(0,i)) + ","
'			Next
'		Else
'			rw "S_NONE"
'			dbCTget.Close() : response.end
'		End If
'
'		IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
'		paramData = "redSsnKey=system&cmdparam=EditPriceSelect&cksel="&itemidArr
'		retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCjMallReq.asp",paramData)
'		response.Write "itemidArr="&itemidArr
'		response.Write "<br>"&retVal
'    Case "cjmallEditItemOptChk" '' cj 상품수정(옵션관련)
'		sqlStr = ""
'		sqlStr = sqlStr & " select top 10 ro.itemid"
'		sqlStr = sqlStr & " from db_outmall.dbo.tbl_cjmall_regItem r"
'		sqlStr = sqlStr & " Join db_outmall.dbo.tbl_outmall_regedoption ro on ro.itemid=r.itemid"
'		sqlStr = sqlStr & " Join db_AppWish.dbo.tbl_item_option o on ro.itemid=o.itemid and ro.itemoption=o.itemoption"
'		sqlStr = sqlStr & " Join db_AppWish.dbo.tbl_item i on r.itemid=i.itemid"
'		sqlStr = sqlStr & " where ro.mallid='cjmall'"
'		sqlStr = sqlStr & " and i.optionCnt>0"
'		sqlStr = sqlStr & " and r.cjmallsellyn='Y'"
'		sqlStr = sqlStr & " and r.cjmallStatCd>3"   ''2013/06/20 추가
'		sqlStr = sqlStr & " and (o.optsellyn='N' or (o.optsellyn='Y' and o.optlimityn='Y' and (o.optlimitno-o.optlimitsold<1)))"
'		sqlStr = sqlStr & " and ro.outmallsellyn='Y'"
'		sqlStr = sqlStr & " and r.accFailCnt < 5 "
'		sqlStr = sqlStr & " group by ro.itemid,r.cjmallLastUpdate,lastStatCheckDate,i.lastupdate"
'		sqlStr = sqlStr & " order by r.lastStatCheckDate"
'		rsCTget.Open sqlStr,dbCTget,1
'		if not rsCTget.Eof then
'			ArrRows = rsCTget.getRows()
'		end if
'		rsCTget.close
'
'		itemidArr = ""
'		if isArray(ArrRows) then
'			For i =0 To UBound(ArrRows,2)
'				itemidArr = itemidArr + CStr(ArrRows(0,i)) + ","
'			Next
'		else
'			rw "S_NONE"
'			dbCTget.Close() : response.end
'		end if
'
'		IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
'		paramData = "redSsnKey=system&cmdparam=EditSelect2&cksel="&itemidArr
'		retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCjMallReq.asp",paramData)
'		response.Write "itemidArr="&itemidArr
'		response.Write "<br>"&retVal
'	Case "cjNotOptionMaySellOK"
'		Set cjMall = new CCjmall
'			cjMall.FCurrPage					= 1
'			cjMall.FPageSize					= 20
'			cjMall.FRectSellYn					= "Y"		'판매Y
'			cjMall.FRectExtSellYn               = "N"		'제휴판매Y
'			cjMall.FRectIsReged					= "Q"		'등록상품 판매가능
'			cjMall.FRectIsOption				= "optN"	'옵션=단품
'			cjMall.FRectOPTCntEqual				= "Y"
'			cjMall.FRectFailCntOverExcept		= "3"		'3회 재낌
'			cjMall.GetCjmallRegedItemList
'
'			If (cjMall.FResultCount < 1) Then
'				response.Write "S_NONE"
'				dbCTget.Close()
'				Set cjMall= Nothing
'				response.end
'			End If
'
'			For i = 0 to cjMall.FResultCount - 1
'				itemidArr = itemidArr & cjMall.FItemList(i).FItemID &","
'			Next
'
'			IF (Right(itemidArr,1)=",") then itemidArr=Left(itemidArr,Len(itemidArr)-1)
'			paramData = "redSsnKey=system&cmdparam=EditSelect2&cksel="&itemidArr
'			retVal = SendReq("http://wapi.10x10.co.kr/outmall/cjmall/actCjMallReq.asp",paramData)
'			response.Write "itemidArr="&itemidArr
'			response.Write "<br>"&retVal
'    ''---------------------------------------------------------------------------------------
'    Case ELSE
'        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->