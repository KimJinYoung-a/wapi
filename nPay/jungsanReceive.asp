<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Call Response.AddHeader("Access-Control-Allow-Origin", "*")
Call Response.AddHeader("Access-Control-Allow-Credentials", "true")

Response.CharSet="UTF-8"
%>
<% Server.ScriptTimeOut = 3600 %>
<!-- #include virtual="/lib/util/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/nPay/incNaverpayCommon.asp" -->
<script language="jscript" runat="server">
    function jsURLDecode(v){ return decodeURI(v); }
    function jsURLEncode(v){ return encodeURI(v); }
</script>

<%
''스케줄 :113SVR : /jobs/applistReceive.vbs 오전8시33분

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("192.168.1.67","192.168.1.70","192.168.1.71","192.168.1.72","192.168.1.73","110.93.128.107","121.78.103.60","110.93.128.114","110.93.128.113","112.218.65.244")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

function debugWrite(iTxt)
    if (application("Svr_Info")	= "Dev") then
        if (request("isautoscript")<>"on") then
            response.write iTxt&"<br>"
        end if
    end if
end function

function NpayResultToAppTemp(NPay_Result,PGgubun,PGuserid,sitename)
    Dim listlength,payHistId, merchantPayKey, merchantUserKey, admissionTypeCode, admissionYmdt
    Dim totalPayAmount, primaryPayAmount,npointPayAmount,primaryPayMeans
    Dim cardCorp, cardAuthNo, cardInstCount, bankCorp
    Dim settleExprectAmount, payCommissionAmount

    dim appDivCode, PGkey, PGCSkey, appDate, cancelDate, appMethod
    dim appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, etcPoint
    dim sqlStr
    dim i

    listlength = NPay_Result.body.list.length
    for i=0 to listlength-1
        payHistId = NPay_Result.body.list.get(i).payHistId                          '' 네이버페이 일련번호
        merchantPayKey = NPay_Result.body.list.get(i).merchantPayKey                '' 가맹점 결제 키
        merchantUserKey = NPay_Result.body.list.get(i).merchantUserKey              '' 가맹점 사용자 키(빈값)
        admissionTypeCode = NPay_Result.body.list.get(i).admissionTypeCode          '' 01:원결제승인건, 03:전체취소건, 04:부분취소건 (**)
        admissionYmdt = NPay_Result.body.list.get(i).admissionYmdt                  '' 결제/취소 일시
        totalPayAmount = NPay_Result.body.list.get(i).totalPayAmount                '' 총 결제/취소금액
        primaryPayAmount = NPay_Result.body.list.get(i).primaryPayAmount            '' 주결제수단 결제/취소금액
        npointPayAmount = NPay_Result.body.list.get(i).npointPayAmount              '' 네이버페이 포인트 결제/취소금액
        primaryPayMeans = NPay_Result.body.list.get(i).primaryPayMeans              '' 주결제수단(CARD/BANK)
        cardCorp    = NPay_Result.body.list.get(i).cardCorpCode                     '' 주결제수단 카드사
        ''cardNo    = NPay_Result.body.list.get(i).cardNo                           '' 일부 마스킹된 카드번호
        cardAuthNo  = NPay_Result.body.list.get(i).cardAuthNo                       '' 카드 승인 번호 (취소시에는 승인번호 개념이 없으므로 원결제시만 반환)
        cardInstCount = NPay_Result.body.list.get(i).cardInstCount                  '' 할부개월수 (일시불은 0 )
        bankCorp    = NPay_Result.body.list.get(i).bankCorpCode                     '' 주결제수단은행
        ''accountNo    = NPay_Result.body.list.get(i).accountNo                     '' 일부 마스킹된 계좌번호
        ''productName    = NPay_Result.body.list.get(i).productName                 '' 상품명

        on Error resume next
        settleExprectAmount    = NPay_Result.body.list.get(i).settleInfo.totalSettleAmount   '' 정산예정금액 (결제후 약 1시간 이후 생성, 그전까지는 반환되지 않음)
        if (ERR) then
              settleExprectAmount = 0
        end if
        on Error Goto 0

        on Error resume next
        payCommissionAmount    = NPay_Result.body.list.get(i).settleInfo.totalCommissionAmount   '' 결제수수료금액 (결제후 약 1시간 이후 생성, 그전까지는 반환되지 않음)
        if (ERR) then
              settleExprectAmount = 0
        end if
        on Error Goto 0

        'response.write payHistId&"|"&merchantPayKey&"|"&merchantUserKey&"|"&admissionTypeCode&"|"&admissionYmdt&"|"
        'response.write totalPayAmount&"|"&primaryPayAmount&"|"&npointPayAmount&"|"&primaryPayMeans&"|"
        'response.write cardCorp&"|"&cardAuthNo&"|"&cardInstCount&"|"&bankCorp&"|"
        'response.write settleExprectAmount&"|"&payCommissionAmount
        'response.write "<br>"
        ''&"|"&merchantUserKey&"|"&admissionTypeCode&"|"&admissionYmdt&"|"&totalPayAmount&"|"&primaryPayAmount&

        Select Case admissionTypeCode
    		Case "01"
    			PGkey			= merchantPayKey
    			PGCSkey			= ""
    			appDivCode 		= "A"
    			appDate 		= admissionYmdt
    			appDate 		= "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
    			cancelDate		= "NULL"

    		Case "03"
    			PGkey			= merchantPayKey
    			PGCSkey			= "CANCELALL"
    			appDivCode 		= "C"
    			appDate			= "NULL"
    			cancelDate 		= admissionYmdt
    			cancelDate 		= "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
    		Case "04"
    			PGkey			= merchantPayKey
    			PGCSkey			= payHistId
    			appDivCode 		= "R"
    			appDate			= "NULL"
    			cancelDate 		= admissionYmdt
    			cancelDate 		= "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
    		Case Else
    			PGkey			= merchantPayKey
    			PGCSkey			= "ERROR"
    			appDivCode 		= "E"
    			appDate 		= admissionYmdt
    			appDate 		= "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
    			cancelDate		= "NULL"
    	End Select

    	Select Case primaryPayMeans
    		Case "CARD"
    			appMethod = "100"
    		Case "BANK"
    			appMethod = "20"
    		Case ""
    			'// 빈값은 네이버포인트
    			appMethod = "20"
    		Case Else
    			appMethod = primaryPayMeans
    	End Select

    	appPrice 		= totalPayAmount
    	commPrice		= payCommissionAmount
    	commVatPrice	= Round(1.0 * commPrice * (1.0/11))
    	commPrice 		= commPrice - commVatPrice
    	jungsanPrice 	= settleExprectAmount
    	ipkumdate 		= ""
    	etcPoint		= npointPayAmount

    	if (appDivCode = "A") then
    		commPrice = commPrice * -1
    		commVatPrice = commVatPrice * -1
    	else
    		appPrice = appPrice * -1
    		jungsanPrice = jungsanPrice * -1
    	end if

		''NULL Check
		''PGkey = chkIIF(isNull(PGkey),"",PGkey)
		''PGCSkey = chkIIF(isNull(PGkey),"",PGCSkey)
		''appMethod = chkIIF(isNull(PGkey),"20",appMethod)
		''appDate = chkIIF(isNull(PGkey),"",appDate)
		''cancelDate = chkIIF(isNull(PGkey),"",cancelDate)
		''appPrice = chkIIF(isNull(PGkey),0,appPrice)
		''commPrice = chkIIF(isNull(PGkey),0,commPrice)
		''commVatPrice = chkIIF(isNull(PGkey),0,commVatPrice)
		''jungsanPrice = chkIIF(isNull(PGkey),0,jungsanPrice)
		''etcPoint = chkIIF(isNull(PGkey),0,etcPoint)

        appPrice = CHKIIF(IsNull(appPrice), (jungsanPrice - commPrice - commVatPrice), appPrice)

    	sqlStr = " if NOT Exists( select top 1 * from db_temp.dbo.tbl_onlineApp_log_tmp where PGgubun='" + CStr(PGgubun) + "' and PGkey='" + CStr(PGkey) + "' and PGCSkey='" + CStr(PGCSkey) + "')"&vbCRLF
		sqlStr = sqlStr + " BEGIN"&vbCRLF
		sqlStr = sqlStr + " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, etcPoint) "&vbCRLF
		sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "', " & etcPoint & ") "&vbCRLF
		sqlStr = sqlStr + " END"&vbCRLF

        ''response.write "PGgubun " & PGgubun & "<br />"
        ''response.write "PGkey " & PGkey & "<br />"
        ''response.write "PGCSkey " & PGCSkey & "<br />"
        ''response.write "sitename " & sitename & "<br />"
        ''response.write "appDivCode " & appDivCode & "<br />"
        ''response.write "appMethod " & appMethod & "<br />"
        ''response.write "appDate " & appDate & "<br />"
        ''response.write "cancelDate " & cancelDate & "<br />"
        ''response.write "appPrice " & appPrice & "<br />"
        ''response.write "commPrice " & commPrice & "<br />"
        ''response.write "commVatPrice " & commVatPrice & "<br />"
        ''response.write "jungsanPrice " & jungsanPrice & "<br />"
        ''response.write "ipkumdate " & ipkumdate & "<br />"
        ''response.write "PGuserid " & PGuserid & "<br />"
        ''response.write "etcPoint " & etcPoint & "<br />"

    	''response.write sqlStr + "<br>"
    	dbget.execute sqlStr
    next

    NpayResultToAppTemp = listlength
end function
''''-------------------------------------------------------------
dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    response.write ref
    response.end
end if

Dim i,t
Dim startTime,endTime
Dim npId

Dim NPay_Result, iResultCode, iResultMsg
dim sDate, eDate, page, startPage, endPage, pageSize

dim PGgubun, PGuserid, sitename, retLength
dim tmpTime, sqlStr
dim bufDate ,ArrRows, CSpaygateTid, NPay_Result_CS, retLength_CS
dim pagingData
pagingData = 0

npId  = requestCheckVar(request("npId"), 30)
sDate = requestCheckVar(request("sDate"), 10)
eDate = requestCheckVar(request("eDate"), 10)
page = requestCheckVar(request("page"), 32)

Dim currDT : currDT = dateAdd("d",-1,Now())

if (sDate <> "") and (eDate <> "") then
	startTime = replace(LEFT(sDate,10),"-","")&"000000"
	endTime = replace(LEFT(eDate,10),"-","")&"235959"
else
	startTime = replace(LEFT(currDT,10),"-","")&"000000"
	endTime = replace(LEFT(currDT,10),"-","")&"235959"
end if

tmpTime = Left(endTime, 4) + "-" + Right(Left(endTime, 6), 2) + "-" + Right(Left(endTime, 8), 2) + " " + Right(Left(endTime, 10), 2) + ":" + Right(Left(endTime, 12), 2) + ":" + Right(Left(endTime, 14), 2)

if (DateDiff("h", tmpTime, Now()) < 2) then
	'// 참조 : 정산예정금액 (결제후 약 1시간 이후 생성, 그전까지는 반환되지 않음)
	response.write "S_ERR|결제일자 기준 2시간이 지난 내역까지만 가져올 수 있음."
	dbget.close()	:	response.End
end if

''response.write  startTime
''response.write  endTime

if (npId<>"") then
    SET NPay_Result = fnCallNaverPayCheck(npId)
	fnNpayResultProcess(NPay_Result)
	Set NPay_Result = Nothing
else
	'// 한번에 20개 기준 총 페이지 숫자 가져옴
	pagingData = fnCallNaverPayListPageCheck(startTime,endTime).body.totalPageCount

    if (page = "") then
        '// 페이징 없음
        For t = 1 to pagingData
		    Set NPay_Result = fnCallNaverPayListNew(startTime,endTime,t)
		    fnNpayResultProcess(NPay_Result)
		    Set NPay_Result = Nothing
	    Next
    else
        response.write "TOTAL PAGE : " & pagingData & "<br />"
        if (page = "0") then
            '// 페이지수 표시
            response.write "S_OK"
        else
            '// 페이징
            pageSize = 3
            startPage = 1 + (pageSize * (page - 1))
            endPage = startPage + (pageSize - 1)
            if (startPage > pagingData) then
			    response.write "S_ERR|NO_DATA"
			    dbget.close()	:	response.End
            else
                if (endPage > pagingData) then
                    endPage = pagingData
                end if

                response.write "PROCESSING PAGE : " & startPage & " ~ " & endPage & "<br />"

                For t = startPage to endPage
		            Set NPay_Result = fnCallNaverPayListNew(startTime,endTime,t)
		            fnNpayResultProcess(NPay_Result)
		            Set NPay_Result = Nothing
	            Next
            end if
        end if
    end if
end if


Function fnNpayResultProcess(NPayResult)
	if NPayResult.code="Success" then
		PGgubun = "naverpay"
		PGuserid = "naverpay"
		sitename = "10x10"

		sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
		sqlStr = sqlStr & " where PGgubun = '" & PGgubun & "' " & VbCRLF
		''response.write sqlStr
		dbget.execute sqlStr

		retLength = NpayResultToAppTemp(NPayResult,PGgubun,PGuserid,sitename)

		''CS 내역 단일 조회. 1일치씩만 가능 하기로
		if (npId="") and (LEFT(startTime,8)=LEFT(endTime,8)) then
			bufDate = LEFT(startTime,8)

			bufDate = LEFT(bufDate,4)&"-"&MID(bufDate,5,2)&"-"&MID(bufDate,7,2)

			''response.write  LEFT(dateadd("d",1,bufDate),10)

			sqlStr = "select  c.orderserial,R.paygateTid, R.returnmethod"&VBCRLF
			sqlStr = sqlStr & " from db_cs.dbo.tbl_new_As_list C"&VBCRLF
			sqlStr = sqlStr & " 	JOin db_cs.dbo.tbl_as_refund_info R"&VBCRLF
			sqlStr = sqlStr & " 	on C.id=R.asid"&VBCRLF
			sqlStr = sqlStr & " 	and R.returnmethod in ('R100','R020','R120','R022')"&VBCRLF
			sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master m"&VBCRLF
			sqlStr = sqlStr & " 	on C.orderserial=m.orderserial"&VBCRLF
			sqlStr = sqlStr & " 	and m.pggubun='NP'"&VBCRLF
			sqlStr = sqlStr & " where C.finishdate>='"&bufDate&"'"&VBCRLF
			sqlStr = sqlStr & " and C.finishdate<'"&LEFT(dateadd("d",1,bufDate),10)&"'"&VBCRLF
			sqlStr = sqlStr & " and C.divcd in ('A007')"

			''response.write sqlStr
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			IF not rsget.EOF THEN
				ArrRows = rsget.getRows()
			END IF
			rsget.close

			if IsArray(ArrRows) then
				For i=0 To UBound(ArrRows,2)
					CSpaygateTid = ArrRows(1,i)
					''response.write CSpaygateTid & "<br />"

					retLength_CS = 0
					if (CSpaygateTid<>"") then
						SET NPay_Result_CS = fnCallNaverPayCheck(CSpaygateTid)
						if NPayResult.code="Success" then
							retLength_CS = NpayResultToAppTemp(NPay_Result_CS,PGgubun,PGuserid,sitename)
						end if
						SET NPay_Result_CS = Nothing
						retLength = retLength + retLength_CS
					end if
				Next
			end if

		end if

	''임시. 이곳에서 멈출경우.
	''dbget.close()	:	response.End

		if (retLength<1) then
			response.write "S_ERR|NO_DATA"
			dbget.close()	:	response.End
		else
			''P_TID 가 빈값인 CASE temp_idx=2842006,2842047
			sqlStr = " update at "
			sqlStr = sqlStr + " set at.PGkey = (CASE WHEN isNULL(ot.P_TID,'')='' THEN at.PGkey ELSE ot.P_TID END)"
			sqlStr = sqlStr + " , at.orderserial = ot.orderserial, at.sitename = (case when ot.rdsite = 'mobile' or ot.rdsite = 'app_wish2' then '10x10mobile' else '10x10' end) "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp at "
			sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_temp] ot "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and at.PGgubun = 'naverpay' "
			sqlStr = sqlStr + " 		and at.PGkey = temp_idx "
			''sqlStr = sqlStr + " 		and ot.IsSuccess = 'True' "		'// 우리쪽 실패여도, 승인내역 넘어왔으면 매칭함, skyer9, 2019-12-23
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	at.orderserial is NULL "
			dbget.execute sqlStr

			'// 승인은 성공했으나, 기타 오류로 주문성공 안된 케이스
			sqlStr = " select top 1 1 from db_temp.dbo.tbl_onlineApp_log_tmp a "
			sqlStr = sqlStr + " where PGgubun = 'naverpay' and orderserial = '' "
			''response.write sqlStr
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

			retLength = 0
			IF not rsget.EOF THEN
				retLength = 1
			END IF
			rsget.close

			if (retLength > 0) then
				sqlStr = " update at "
				sqlStr = sqlStr + " set at.orderserial = ot.orderserial "
				sqlStr = sqlStr + "  from  "
				sqlStr = sqlStr + "  	db_temp.dbo.tbl_onlineApp_log_tmp at  "
				sqlStr = sqlStr + "  	join [db_order].[dbo].[tbl_order_master] ot  "
				sqlStr = sqlStr + "  	on  "
				sqlStr = sqlStr + "  		1 = 1  "
				sqlStr = sqlStr + "  		and at.PGgubun = 'naverpay'  "
				sqlStr = sqlStr + "  		and at.PGkey = ot.paygatetid  "
				sqlStr = sqlStr + " 		and at.orderserial = '' "
				dbget.execute sqlStr
			end if

			''2016/08/18 추가
			sqlStr = " update at"
			sqlStr = sqlStr + " SET PGkey=m.paygateTID"
			sqlStr = sqlStr + " from db_temp.dbo.tbl_onlineApp_log_tmp at"
			sqlStr = sqlStr + " 	Join db_order.dbo.tbl_order_master m"
			sqlStr = sqlStr + " 	on at.orderserial=m.orderserial"
			sqlStr = sqlStr + " where at.PGgubun = 'naverpay'"
			sqlStr = sqlStr + " and LEN(at.PGKey)<11"
			dbget.execute sqlStr

			sqlStr = " delete l "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
			sqlStr = sqlStr + " 	join db_order.dbo.tbl_onlineApp_log l "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
			sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
			sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and l.idx is not NULL "
			sqlStr = sqlStr + " 	and l.appDate is NULL "
			sqlStr = sqlStr + " 	and l.cancelDate is NULL "
			sqlStr = sqlStr + " 	and t.PGgubun = '" + CStr(PGgubun) + "' "
			dbget.execute sqlStr

            '// 중복입력 삭제
			sqlStr = " delete T "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_temp] OT "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		T.pgkey = OT.temp_idx "
			sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_onlineApp_log] l "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and OT.orderserial = l.orderserial "
			sqlStr = sqlStr + " 		and T.appDivCode = l.appdivcode "
			sqlStr = sqlStr + " 		and T.pgcskey = l.pgcskey "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.pggubun = 'naverpay' "
			sqlStr = sqlStr + " 	and len(T.pgkey) < 11 "
            dbget.execute sqlStr

			sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate, orderserial, etcPoint) "
			sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121), t.orderserial, t.etcPoint "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
			sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
			sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
			sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
            ''sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_order_temp] ot "
			''sqlStr = sqlStr + " 	on "
			''sqlStr = sqlStr + " 		1 = 1 "
			''sqlStr = sqlStr + " 		and t.PGkey = ot.temp_idx "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and l.idx is NULL "
            ''sqlStr = sqlStr + " 	and ot.temp_idx is NULL "					'// 2022-04-11, skyer9
			sqlStr = sqlStr + " 	and t.PGgubun = '" + CStr(PGgubun) + "' "
			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
			''response.write sqlStr + "<br>"
			dbget.execute sqlStr

			response.write "S_OK"
		end if
	else
		iResultCode = NPayResult.code
		iResultMsg = replace(NPayResult.message,"'","")

		response.write iResultCode&"|"&iResultMsg

		response.write "S_ERR|[" & iResultCode & "] " & iResultMsg
		dbget.close()	:	response.End
	end if
End Function

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
