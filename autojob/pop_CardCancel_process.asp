<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
<!-- #include virtual="/cscenter/action/incKakaopayCommon.asp"-->

<%
''KaKao 신용카드 취소
function CanCelKakaoPay(ipaygatetid,irefundrequire,irdSite,byREF iretval,byREF iResultCode,byREF iResultMsg,byREF iCancelDate,byREF iCancelTime)
    Dim objKMPay

dim otime,orgTim,diffTime
otime = Timer()
orgTim = otime

    '1) 객체 생성
    Set objKMPay = Server.CreateObject("LGCNS.CNSPayService.CnsPayWebConnector")
    objKMPay.RequestUrl = CNSPAY_DEAL_REQUEST_URL

    '2) 로그 정보
    objKMPay.SetCnsPayLogging KMPAY_LOG_DIR, KMPAY_LOG_LEVEL	'-1:로그 사용 안함, 0:Error, 1:Info, 2:Debug

    '3) 요청 페이지 파라메터 셋팅
    objKMPay.AddRequestData "MID", KMPAY_MERCHANT_ID
    objKMPay.AddRequestData "TID", ipaygatetid

    ''objKMPay.AddRequestData "Amt", irefundrequire
	objKMPay.AddRequestData "CancelAmt", irefundrequire

    ''objKMPay.AddRequestData "SupplyAmt",0     ''공급가
    ''objKMPay.AddRequestData "GoodsVat",0      ''부가세
    ''objKMPay.AddRequestData "ServiceAmt",0    ''봉사료
    objKMPay.AddRequestData "CancelMsg","고객요청"
    objKMPay.AddRequestData "PartialCancelCode","0"     '' 0전체취소, 1부분취소.
    objKMPay.AddRequestData "PayMethod","CARD"


    '4) 추가 파라메터 셋팅
    objKMPay.AddRequestData "actionType", "CL0"  															' actionType : CL0 취소, PY0 승인, CI0 조회
    objKMPay.AddRequestData "CancelIP", Request.ServerVariables("LOCAL_ADDR")	' 가맹점 고유 ip
    objKMPay.AddRequestData "CancelPwd", KMPAY_CANCEL_PWD														' 취소 비밀번호 설정

    '5) 가맹점키 셋팅 (MID 별로 틀림)
    objKMPay.AddRequestData "EncodeKey", KMPAY_MERCHANT_KEY

diffTime = FormatNumber(Timer()-otime,4)
rw diffTime
    '6) CNSPAY Lite 서버 접속하여 처리
    objKMPay.RequestAction
rw diffTime
    '7) 결과 처리
    Dim resultCode, resultMsg, cancelAmt, cancelDate, cancelTime, payMethod, resMerchantId, tid, errorCD, errorMsg, authDate, ccPartCl, stateCD

    resultCode = objKMPay.GetResultData("ResultCode") 	' 결과코드 (정상 :2001(취소성공), 2002(취소진행중), 그 외 에러)
    resultMsg = objKMPay.GetResultData("ResultMsg")   	' 결과메시지
    cancelAmt = objKMPay.GetResultData("CancelAmt")   	' 취소금액
    cancelDate = objKMPay.GetResultData("CancelDate") 	' 취소일
    cancelTime = objKMPay.GetResultData("CancelTime")   ' 취소시간
    payMethod = objKMPay.GetResultData("PayMethod")   	' 취소 결제수단
    resMerchantId = objKMPay.GetResultData("MID")     	' 가맹점 ID
    tid = objKMPay.GetResultData("TID")               	' TID
    errorCD = objKMPay.GetResultData("ErrorCD")        	' 상세 에러코드
    errorMsg = objKMPay.GetResultData("ErrorMsg")      	' 상세 에러메시지
    authDate = cancelDate & cancelTime									' 거래시간
    ccPartCl = objKMPay.GetResultData("CcPartCl")       ' 부분취소 가능여부 (0:부분취소불가, 1:부분취소가능)
    stateCD = objKMPay.GetResultData("StateCD")         ' 거래상태코드 (0: 승인, 1:전취소, 2:후취소)

    if (resultCode="2001") then
        iretval = "0"
        iResultCode = resultCode
        iResultMsg = resultMsg
        iCancelDate	= cancelDate
	    iCancelTime	= cancelTime
    else
        iResultCode = resultCode
        iResultMsg = resultMsg
    end if

    Set objKMPay = Nothing

end function

''데이콤 휴대폰 실취소
function CanCelMobileDacom(ipaygatetid,irefundrequire,irdSite,byREF iretval,byREF iResultCode,byREF iResultMsg,byREF iCancelDate,byREF iCancelTime)
    Dim CST_PLATFORM, CST_MID, LGD_MID, LGD_TID,Tradeid, LGD_CANCELREASON, LGD_CANCELREQUESTER, LGD_CANCELREQUESTERIP
    Dim configPath, xpay

    IF (application("Svr_Info") = "Dev") THEN                   ' LG유플러스 결제서비스 선택(test:테스트, service:서비스)
		CST_PLATFORM = "test"
	Else
		CST_PLATFORM = "service"
	End If


    CST_MID              = "tenbyten02"                         ' LG유플러스으로 부터 발급받으신 상점아이디를 입력하세요. //모바일, 서비스 동일.
                                                                ' 테스트 아이디는 't'를 제외하고 입력하세요.
    if CST_PLATFORM = "test" then                               ' 상점아이디(자동생성)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if

    Tradeid     = Split(ipaygatetid,"|")(0)
	LGD_TID     = Split(ipaygatetid,"|")(1)                     ' LG유플러스으로 부터 내려받은 거래번호(LGD_TID) : 24 byte

    LGD_CANCELREASON        = "고객요청"                        ' 취소사유
    LGD_CANCELREQUESTER     = "고객"                            ' 취소요청자
    LGD_CANCELREQUESTERIP   = Request.ServerVariables("REMOTE_ADDR")     ' 취소요청IP


    configPath           = "C:/lgdacom" ''"C:/lgdacom/conf/"&CST_MID         				' LG유플러스에서 제공한 환경파일("/conf/lgdacom.conf") 위치 지정.
    Set xpay = CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "Cancel"
    xpay.Set "LGD_TID", LGD_TID
    xpay.Set "LGD_CANCELREASON", LGD_CANCELREASON
    xpay.Set "LGD_CANCELREQUESTER", LGD_CANCELREQUESTER
    xpay.Set "LGD_CANCELREQUESTERIP", LGD_CANCELREQUESTERIP

    '/*
    ' * 1. 결제취소 요청 결과처리
    ' *
    ' * 취소결과 리턴 파라미터는 연동메뉴얼을 참고하시기 바랍니다.
	' *
	' * [[[중요]]] 고객사에서 정상취소 처리해야할 응답코드
	' * 1. 신용카드 : 0000, AV11
	' * 2. 계좌이체 : 0000, RF00, RF10, RF09, RF15, RF19, RF23, RF25 (환불진행중 응답-> 환불결과코드.xls 참고)
	' * 3. 나머지 결제수단의 경우 0000(성공) 만 취소성공 처리
	' *
    ' */

    if xpay.TX() then
        '1)결제취소결과 화면처리(성공,실패 결과 처리를 하시기 바랍니다.)
'Response.Write("결제취소 요청이 완료되었습니다. <br>")
'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")

        iretval = "0"
        iResultCode = xpay.resCode
		iResultMsg	= xpay.resMsg
    else
        '2)API 요청 실패 화면처리
'Response.Write("결제취소 요청이 실패하였습니다. <br>")
'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
        iResultCode = xpay.resCode
		iResultMsg	= xpay.resMsg
    end if

    iCancelDate	= year(now) & "년 " & month(now) & "월 " & day(now) & "일"
	iCancelTime	= hour(now) & "시 " & minute(now) & "분 " & second(now) & "초"

end function

'''신용카드 부분취소 R120 => 다른 페이지에서 따로 처리.
'''핸드폰 부분취소 R420 => 다른 페이지에서 따로 처리.

dim id, finishuserid, msg, force
dim orgOrderSerial, chgOrderserial
dim jumundiv, accountdiv

id           = RequestCheckVar(request("id"),10)
msg          = RequestCheckVar(request("msg"),50)
finishuserid = session("ssBctID")
force = RequestCheckVar(request("force"),10)

if (msg="") and (IsAutoScript) then msg="배송전취소"

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if


dim orefund
set orefund = New CCSASList
orefund.FRectCsAsID = id

if (id<>"") then
    orefund.GetOneRefundInfo
end if

if (ocsaslist.FResultCount<1) or (orefund.FResultCount<1) then
    if (IsAutoScript) then
        response.write "S_ERR|환불내역이 없거나 유효하지 않은 내역입니다."
    else
        response.write "<script>alert('환불내역이 없거나 유효하지 않은 내역입니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

if (ocsaslist.FOneItem.FCurrstate<>"B001") then
    if (IsAutoScript) then
        response.write "S_ERR|접수 상태가 아닙니다."
    else
        response.write "<script>alert('접수 상태가 아닙니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

'' 신용카드 취소만 가능
'if (orefund.FOneItem.Freturnmethod<>"R100") then
'    response.write "<script>alert('현재 신용카드 거래만 취소 가능합니다.');</script>"
'    response.write "<script>window.close();</script>"
'    dbget.close()	:	response.End
'end if

Dim returnmethod, IsCardPartialCancel
returnmethod = orefund.FOneItem.Freturnmethod

if Not ((returnmethod="R100") or (returnmethod="R020") or (returnmethod="R400")) Then
    if (IsAutoScript) then
        response.write "S_ERR|신용카드 전체취소, 실시간이체 취소, 휴대폰 전체 취소 만 가능."
    else
        response.write "<script>alert('신용카드 전체취소, 실시간이체 취소, 휴대폰 전체 취소 만 가능합니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if


'' IniPay 만 취소만 가능
dim IsInicisTID : IsInicisTID = False
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="IniTechPG_")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIMX_CARD")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIMX_ISP_")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIswtCARD")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIswtISP_")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,6)="Stdpay")
''IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIMX_AUTH")

''if ((Left(orefund.FOneItem.FpaygateTid,3)<>"cns") and (Left(orefund.FOneItem.FpaygateTid,5)<>"KCTEN") and Left(orefund.FOneItem.FpaygateTid,10)<>"IniTechPG_") AND (Left(orefund.FOneItem.FpaygateTid,10)<>"INIMX_CARD") AND (Left(orefund.FOneItem.FpaygateTid,10)<>"INIMX_ISP_") AND (Left(orefund.FOneItem.FpaygateTid,10)<>"INIswtCARD") AND (Left(orefund.FOneItem.FpaygateTid,10)<>"INIswtISP_") AND orefund.FOneItem.Freturnmethod<>"R400" then
if (Left(orefund.FOneItem.FpaygateTid,3)<>"cns") and (Left(orefund.FOneItem.FpaygateTid,5)<>"KCTEN") and Not IsInicisTID AND orefund.FOneItem.Freturnmethod<>"R400" then
    if (IsAutoScript) then
        response.write "S_ERR|이니시스 거래만 취소 가능합니다."
    else
        response.write "<script>alert('이니시스 거래만 취소 가능합니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

''=============전체취소만 가능함.. 부분취소등 취소안됨..=============
dim sqlStr, isSameMoney
dim t_refundrequire, t_MaybeOrgPayPrice
isSameMoney = false

''마이너스 주문일경우 원주문번호// ===> 원주문건으로 변경됨..
sqlStr = " select r.refundrequire, m.orderserial, m.jumundiv, m.linkorderserial"
sqlStr = sqlStr & " from db_cs.dbo.tbl_new_as_list l"
sqlStr = sqlStr & " 	Join db_cs.dbo.tbl_as_refund_info r"
sqlStr = sqlStr & " 	on l.id=r.asid"
sqlStr = sqlStr & " 	and r.returnmethod  in ('R100','R020','R400')"
sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master m"
sqlStr = sqlStr & " 	on l.orderserial=m.orderserial"
sqlStr = sqlStr & " where l.id="&id
sqlStr = sqlStr & " and l.divcd='A007'"

rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
    t_refundrequire=rsget("refundrequire")
    ''if (rsget("jumundiv")="9") then
    ''    orgOrderserial = rsget("linkorderserial")
    ''else
        orgOrderserial = rsget("orderserial")
    ''end if
end if
rsget.Close


'// 교환주문( jumundiv = 6 )이면 원주문에서 결제정보 가져온다.
sqlStr = " select top 1 m.jumundiv, m.accountdiv "
sqlStr = sqlStr + " from "
sqlStr = sqlStr + " 	db_order.dbo.tbl_order_master m "
sqlStr = sqlStr + " where "
sqlStr = sqlStr + " 	1 = 1 "
sqlStr = sqlStr + " 	and m.orderserial = '" & orgOrderserial & "' "
rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
	jumundiv = rsget("jumundiv")
	accountdiv = rsget("accountdiv")
end if
rsget.close

if (jumundiv = "6") then
	sqlStr = " select top 1 c.orgorderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_change_order c "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and c.chgorderserial = '" & orgOrderserial & "' "
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		chgOrderserial = orgOrderserial
		orgOrderserial = rsget("orgorderserial")
	end if
	rsget.close
end if


'''2011-04 이후 tbl_order_paymentEtc 사용.
sqlStr = " select Sum(acctamount) as acctamount"
sqlStr = sqlStr & " from db_order.dbo.tbl_order_paymentEtc"
sqlStr = sqlStr & " where orderserial='"&orgOrderserial&"'"
sqlStr = sqlStr & " and acctdiv in ('100','110')"    ''신용카드 및 OkCashBag은 같이결제됨.
rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
    t_MaybeOrgPayPrice=rsget("acctamount")
    isSameMoney    = (t_refundrequire=(t_MaybeOrgPayPrice))
end if
rsget.Close

IF  (Not isSameMoney) THEN
    IF (force="on") then
        response.write "취소금액과 원금액 상이<br><br>"
    ELSE
        if (IsAutoScript) then
            response.write "S_ERR|취소금액과 원금액 상이"
        else
            response.write "<script>alert('취소금액과 원금액 상이 - 관리자 문의 요망."&t_refundrequire&":"&t_MaybeOrgPayPrice&"');</script>"
            response.write "<script>window.close();</script>"
        end if
        dbget.close()	:	response.End
    End IF
END IF
'''=================================================================


'' Pg_Mid
dim MctID
MctID = Mid(orefund.FOneItem.FpaygateTid,11,10)
'' response.write MctID

dim INIpay, PInst
dim ResultCode, ResultMsg, CancelDate, CancelTime, Rcash_cancel_noappl

''휴대폰 결제 추가 2015/04/21 IsINIMobile
Dim IsINIMobile : IsINIMobile = false
if (orefund.FOneItem.Freturnmethod = "R400") and (Len(orefund.FOneItem.FpaygateTid)=40) then
    IsINIMobile = (LEFT(orefund.FOneItem.Fpaygatetid,LEN("IniTechPG_"))="IniTechPG_") or (LEFT(orefund.FOneItem.Fpaygatetid,LEN("INIMX_HPP_"))="INIMX_HPP_")
end if

Dim IsDacomMobile : IsDacomMobile = false
if (orefund.FOneItem.Freturnmethod = "R400") and (NOT IsINIMobile) then
    if (Len(orefund.FOneItem.FpaygateTid)>=31) then
        IsDacomMobile = True        ''46~49 Tradeid(23) & "|" & vTID(24)  => 263055|tenby2014031117203148569 (31)
    else
        IsDacomMobile = False       ''32~35 Tradeid(23) & "|" & vTID(10)
    end if
end if

''카카오페이
Dim IsKakaoPay : IsKakaoPay = (orefund.FOneItem.Freturnmethod = "R100") and ((Left(orefund.FOneItem.FpaygateTid,3)="cns") or (Left(orefund.FOneItem.FpaygateTid,5)="KCTEN")) ''일단.

'############################################################## 핸드폰 결제 취소 ##############################################################
If (orefund.FOneItem.Freturnmethod = "R400") and (NOT IsINIMobile) Then

    Dim McashCancelObj, Mrchid, Svcid, Tradeid, Prdtprice, Mobilid, retval


    IF (IsDacomMobile) then
        CALL CanCelMobileDacom(orefund.FOneItem.FpaygateTid,orefund.FOneItem.Frefundrequire,Request("rdsite"),retval,ResultCode,ResultMsg,CancelDate,CancelTime)
    ELSE
        '' Not Using MCash
        dim dummi : dummi=1/0
        dbget.close() : response.end


    	Set McashCancelObj = Server.CreateObject("Mcash_Cancel.Cancel.1")

    	Mrchid      = "10030289"
    	If LEFT(Request("rdsite"),6) = "mobile" Then
    		Svcid       = "100302890002"
    	Else
    		Svcid       = "100302890001"
    	End If
    	Tradeid     = Split(orefund.FOneItem.FpaygateTid,"|")(0)
    	Prdtprice   = orefund.FOneItem.Frefundrequire
    	Mobilid     = Split(orefund.FOneItem.FpaygateTid,"|")(1)

    	McashCancelObj.Mrchid			= Mrchid
    	McashCancelObj.Svcid			= Svcid
    	McashCancelObj.Tradeid			= Tradeid
    	McashCancelObj.Prdtprice		= Prdtprice
    	McashCancelObj.Mobilid	        = Mobilid

    	retval = McashCancelObj.CancelData

    	set McashCancelObj = nothing

    	If retval = "0" Then
    		ResultCode 	= "00"
    		ResultMsg	= "정상처리"
    	Else
    		ResultCode = retval
    		Select Case ResultCode
    			Case "14"
    				ResultMsg = "해지"
    			Case "20"
    				ResultMsg = "휴대폰 등록정보 오류(PG사) (LGT의 경우 사용자정보변경에 의한 인증실패)"
    			Case "41"
    				ResultMsg = "거래내역 미존재"
    			Case "42"
    				ResultMsg = "취소기간경과"
    			Case "43"
    				ResultMsg = "승인내역오류 ( 인증정보와의 불일치, 승인번호 유효시간 초과( 3분 ) )"
    			Case "44"
    				ResultMsg = "중복 취소 요청"
    			Case "45"
    				ResultMsg = "취소 요청 시 취소 정보 불일치"
    			Case "97"
    				ResultMsg = "요청자료 오류"
    			Case "98"
    				ResultMsg = "통신사 통신오류"
    			Case "99"
    				ResultMsg = "기타"
    			Case "11"
    				ResultMsg = "고객정보변경건으로 인한 취소불가(11)"
    			Case Else
    				ResultMsg = ""
    		End Select
    	End If

    	CancelDate	= year(now) & "년 " & month(now) & "월 " & day(now) & "일"
    	CancelTime	= hour(now) & "시 " & minute(now) & "분 " & second(now) & "초"
    END IF
ELSEIF (IsKakaoPay) then
    CALL CanCelKakaoPay(orefund.FOneItem.FpaygateTid,orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime)
Else
'############################################################## 카드, 실시간 결제 취소 ##############################################################
		'###############################################################################
		'# 1. 객체 생성 #
		'################
		Set INIpay = Server.CreateObject("INItx41.INItx41.1")


		'###############################################################################
		'# 2. 인스턴스 초기화 #
		'######################
		PInst = INIpay.Initialize("")

		'###############################################################################
		'# 3. 거래 유형 설정 #
		'#####################
		INIpay.SetActionType CLng(PInst), "CANCEL"

		'###############################################################################
		'# 4. 정보 설정 #
		'################
		INIpay.SetField CLng(PInst), "pgid", "IniTechPG_" 'PG ID (고정)
		INIpay.SetField CLng(PInst), "spgip", "203.238.3.10" '예비 PG IP (고정)
		INIpay.SetField CLng(PInst), "mid", MctID '상점아이디
		INIpay.SetField CLng(PInst), "admin", "1111" '키패스워드(상점아이디에 따라 변경)
		INIpay.SetField CLng(PInst), "tid", Request("tid") '취소할 거래번호(TID)
		INIpay.SetField CLng(PInst), "msg", msg '취소 사유
		INIpay.SetField CLng(PInst), "uip", Request.ServerVariables("REMOTE_ADDR") 'IP
		INIpay.SetField CLng(PInst), "debug", "false" '로그모드("true"로 설정하면 상세한 로그를 남김)
		INIpay.SetField CLng(PInst), "merchantreserved", "예비" '예비

		'###############################################################################
		'# 5. 취소 요청 #
		'################
		INIpay.StartAction(CLng(PInst))

		'###############################################################################
		'# 6. 취소 결과 #
		'################
		ResultCode = INIpay.GetResult(CLng(PInst), "resultcode") '결과코드 ("00"이면 취소성공)
		ResultMsg  = INIpay.GetResult(CLng(PInst), "resultmsg") '결과내용
		CancelDate = INIpay.GetResult(CLng(PInst), "pgcanceldate") '이니시스 취소날짜
		CancelTime = INIpay.GetResult(CLng(PInst), "pgcanceltime") '이니시스 취소시각
		Rcash_cancel_noappl = INIpay.GetResult(CLng(PInst), "rcash_cancel_noappl") '현금영수증 취소 승인번호

		'###############################################################################
		'# 7. 인스턴스 해제 #
		'####################
		INIpay.Destroy CLng(PInst)
End If




dim refundrequire,refundresult,userid
dim iorderserial, ibuyhp
dim contents_finish

contents_finish = "결과 " & "[" & ResultCode & "]" & ResultMsg & VbCrlf
contents_finish = contents_finish & "취소일시 : " & CancelDate & " " & CancelTime & VbCrlf
contents_finish = contents_finish & "취소자 ID " & finishuserid

if ((ResultCode="00") or (ResultCode="0000")) or (IsKakaoPay and (resultCode="2001")) then

    sqlStr = "select r.*, a.userid, m.orderserial, m.buyhp from "
    sqlStr = sqlStr + " [db_cs].[dbo].tbl_as_refund_info r,"
    sqlStr = sqlStr + " [db_cs].dbo.tbl_new_as_list a"
    sqlStr = sqlStr + "     left join db_order.dbo.tbl_order_master m "
	sqlStr = sqlStr + "     on a.orderserial=m.orderserial"
    sqlStr = sqlStr + " where r.asid=" + CStr(id)
    sqlStr = sqlStr + " and r.asid=a.id"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        returnmethod    = rsget("returnmethod")
        refundrequire   = rsget("refundrequire")
        refundresult    = rsget("refundresult")
        userid          = rsget("userid")
        iorderserial    = rsget("orderserial")
        ibuyhp          = rsget("buyhp")
    end if
    rsget.Close


    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " set refundresult=" + CStr(refundrequire)
    sqlStr = sqlStr + " where asid=" + CStr(id)
    dbget.Execute sqlStr

	'// OK캐시백 결제인 경우, 반품 및 마이너스 주문 입력 후 카드 전체취소이면 마이너스 주문에 보조결제금액 입력
	if (accountdiv="110") then ''2015/08/05
        sqlStr = " exec [db_order].[dbo].[usp_Ten_AddEtcPaymentWhenCardCancel] '" + CStr(orgOrderserial) + "', '" + CStr(chgOrderserial) + "'"
        dbget.Execute sqlStr
    end if

    Call AddCustomerOpenContents(id, "환불(취소) 완료: " & CStr(refundrequire))


    Call FinishCSMaster(id, finishuserid, contents_finish)

    ''승인 취소 요청 SMS 발송
    if (iorderserial<>"") and (ibuyhp<>"") then
        SendAcctCancelMsg ibuyhp, iorderserial
'		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
'		sqlStr = sqlStr + " values('" + ibuyhp + "',"
'    	sqlStr = sqlStr + " '1644-6030',"
'    	sqlStr = sqlStr + " '1',"
'    	sqlStr = sqlStr + " getdate(),"
'    	sqlStr = sqlStr + " '[텐바이텐]승인 취소 되었습니다. 주문번호 : " + iorderserial + "')"
'		dbget.Execute sqlStr
    end if

    ''메일
    Call SendCsActionMail(id)

    if (IsAutoScript) then
        response.write "S_OK"
    else
        response.write "<script>alert('" & ResultMsg & "');</script>"
        response.write "<script>opener.location.reload();</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End

else
    if (IsAutoScript) then
        response.write "S_ERR|"&ResultMsg
    else
        response.write ResultCode & "<br>"
        response.write ResultMsg & "<br>"
        response.write CancelDate & "<br>"
        response.write CancelTime & "<br>"
        response.write Rcash_cancel_noappl & "<br>"
    end if
end if
%>



<%
set ocsaslist = Nothing
set orefund = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
