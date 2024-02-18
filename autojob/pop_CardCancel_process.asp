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
''KaKao �ſ�ī�� ���
function CanCelKakaoPay(ipaygatetid,irefundrequire,irdSite,byREF iretval,byREF iResultCode,byREF iResultMsg,byREF iCancelDate,byREF iCancelTime)
    Dim objKMPay

dim otime,orgTim,diffTime
otime = Timer()
orgTim = otime

    '1) ��ü ����
    Set objKMPay = Server.CreateObject("LGCNS.CNSPayService.CnsPayWebConnector")
    objKMPay.RequestUrl = CNSPAY_DEAL_REQUEST_URL

    '2) �α� ����
    objKMPay.SetCnsPayLogging KMPAY_LOG_DIR, KMPAY_LOG_LEVEL	'-1:�α� ��� ����, 0:Error, 1:Info, 2:Debug

    '3) ��û ������ �Ķ���� ����
    objKMPay.AddRequestData "MID", KMPAY_MERCHANT_ID
    objKMPay.AddRequestData "TID", ipaygatetid

    ''objKMPay.AddRequestData "Amt", irefundrequire
	objKMPay.AddRequestData "CancelAmt", irefundrequire

    ''objKMPay.AddRequestData "SupplyAmt",0     ''���ް�
    ''objKMPay.AddRequestData "GoodsVat",0      ''�ΰ���
    ''objKMPay.AddRequestData "ServiceAmt",0    ''�����
    objKMPay.AddRequestData "CancelMsg","����û"
    objKMPay.AddRequestData "PartialCancelCode","0"     '' 0��ü���, 1�κ����.
    objKMPay.AddRequestData "PayMethod","CARD"


    '4) �߰� �Ķ���� ����
    objKMPay.AddRequestData "actionType", "CL0"  															' actionType : CL0 ���, PY0 ����, CI0 ��ȸ
    objKMPay.AddRequestData "CancelIP", Request.ServerVariables("LOCAL_ADDR")	' ������ ���� ip
    objKMPay.AddRequestData "CancelPwd", KMPAY_CANCEL_PWD														' ��� ��й�ȣ ����

    '5) ������Ű ���� (MID ���� Ʋ��)
    objKMPay.AddRequestData "EncodeKey", KMPAY_MERCHANT_KEY

diffTime = FormatNumber(Timer()-otime,4)
rw diffTime
    '6) CNSPAY Lite ���� �����Ͽ� ó��
    objKMPay.RequestAction
rw diffTime
    '7) ��� ó��
    Dim resultCode, resultMsg, cancelAmt, cancelDate, cancelTime, payMethod, resMerchantId, tid, errorCD, errorMsg, authDate, ccPartCl, stateCD

    resultCode = objKMPay.GetResultData("ResultCode") 	' ����ڵ� (���� :2001(��Ҽ���), 2002(���������), �� �� ����)
    resultMsg = objKMPay.GetResultData("ResultMsg")   	' ����޽���
    cancelAmt = objKMPay.GetResultData("CancelAmt")   	' ��ұݾ�
    cancelDate = objKMPay.GetResultData("CancelDate") 	' �����
    cancelTime = objKMPay.GetResultData("CancelTime")   ' ��ҽð�
    payMethod = objKMPay.GetResultData("PayMethod")   	' ��� ��������
    resMerchantId = objKMPay.GetResultData("MID")     	' ������ ID
    tid = objKMPay.GetResultData("TID")               	' TID
    errorCD = objKMPay.GetResultData("ErrorCD")        	' �� �����ڵ�
    errorMsg = objKMPay.GetResultData("ErrorMsg")      	' �� �����޽���
    authDate = cancelDate & cancelTime									' �ŷ��ð�
    ccPartCl = objKMPay.GetResultData("CcPartCl")       ' �κ���� ���ɿ��� (0:�κ���ҺҰ�, 1:�κ���Ұ���)
    stateCD = objKMPay.GetResultData("StateCD")         ' �ŷ������ڵ� (0: ����, 1:�����, 2:�����)

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

''������ �޴��� �����
function CanCelMobileDacom(ipaygatetid,irefundrequire,irdSite,byREF iretval,byREF iResultCode,byREF iResultMsg,byREF iCancelDate,byREF iCancelTime)
    Dim CST_PLATFORM, CST_MID, LGD_MID, LGD_TID,Tradeid, LGD_CANCELREASON, LGD_CANCELREQUESTER, LGD_CANCELREQUESTERIP
    Dim configPath, xpay

    IF (application("Svr_Info") = "Dev") THEN                   ' LG���÷��� �������� ����(test:�׽�Ʈ, service:����)
		CST_PLATFORM = "test"
	Else
		CST_PLATFORM = "service"
	End If


    CST_MID              = "tenbyten02"                         ' LG���÷������� ���� �߱޹����� �������̵� �Է��ϼ���. //�����, ���� ����.
                                                                ' �׽�Ʈ ���̵�� 't'�� �����ϰ� �Է��ϼ���.
    if CST_PLATFORM = "test" then                               ' �������̵�(�ڵ�����)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if

    Tradeid     = Split(ipaygatetid,"|")(0)
	LGD_TID     = Split(ipaygatetid,"|")(1)                     ' LG���÷������� ���� �������� �ŷ���ȣ(LGD_TID) : 24 byte

    LGD_CANCELREASON        = "����û"                        ' ��һ���
    LGD_CANCELREQUESTER     = "��"                            ' ��ҿ�û��
    LGD_CANCELREQUESTERIP   = Request.ServerVariables("REMOTE_ADDR")     ' ��ҿ�ûIP


    configPath           = "C:/lgdacom" ''"C:/lgdacom/conf/"&CST_MID         				' LG���÷������� ������ ȯ������("/conf/lgdacom.conf") ��ġ ����.
    Set xpay = CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "Cancel"
    xpay.Set "LGD_TID", LGD_TID
    xpay.Set "LGD_CANCELREASON", LGD_CANCELREASON
    xpay.Set "LGD_CANCELREQUESTER", LGD_CANCELREQUESTER
    xpay.Set "LGD_CANCELREQUESTERIP", LGD_CANCELREQUESTERIP

    '/*
    ' * 1. ������� ��û ���ó��
    ' *
    ' * ��Ұ�� ���� �Ķ���ʹ� �����޴����� �����Ͻñ� �ٶ��ϴ�.
	' *
	' * [[[�߿�]]] ���翡�� ������� ó���ؾ��� �����ڵ�
	' * 1. �ſ�ī�� : 0000, AV11
	' * 2. ������ü : 0000, RF00, RF10, RF09, RF15, RF19, RF23, RF25 (ȯ�������� ����-> ȯ�Ұ���ڵ�.xls ����)
	' * 3. ������ ���������� ��� 0000(����) �� ��Ҽ��� ó��
	' *
    ' */

    if xpay.TX() then
        '1)������Ұ�� ȭ��ó��(����,���� ��� ó���� �Ͻñ� �ٶ��ϴ�.)
'Response.Write("������� ��û�� �Ϸ�Ǿ����ϴ�. <br>")
'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")

        iretval = "0"
        iResultCode = xpay.resCode
		iResultMsg	= xpay.resMsg
    else
        '2)API ��û ���� ȭ��ó��
'Response.Write("������� ��û�� �����Ͽ����ϴ�. <br>")
'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
        iResultCode = xpay.resCode
		iResultMsg	= xpay.resMsg
    end if

    iCancelDate	= year(now) & "�� " & month(now) & "�� " & day(now) & "��"
	iCancelTime	= hour(now) & "�� " & minute(now) & "�� " & second(now) & "��"

end function

'''�ſ�ī�� �κ���� R120 => �ٸ� ���������� ���� ó��.
'''�ڵ��� �κ���� R420 => �ٸ� ���������� ���� ó��.

dim id, finishuserid, msg, force
dim orgOrderSerial, chgOrderserial
dim jumundiv, accountdiv

id           = RequestCheckVar(request("id"),10)
msg          = RequestCheckVar(request("msg"),50)
finishuserid = session("ssBctID")
force = RequestCheckVar(request("force"),10)

if (msg="") and (IsAutoScript) then msg="��������"

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
        response.write "S_ERR|ȯ�ҳ����� ���ų� ��ȿ���� ���� �����Դϴ�."
    else
        response.write "<script>alert('ȯ�ҳ����� ���ų� ��ȿ���� ���� �����Դϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

if (ocsaslist.FOneItem.FCurrstate<>"B001") then
    if (IsAutoScript) then
        response.write "S_ERR|���� ���°� �ƴմϴ�."
    else
        response.write "<script>alert('���� ���°� �ƴմϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

'' �ſ�ī�� ��Ҹ� ����
'if (orefund.FOneItem.Freturnmethod<>"R100") then
'    response.write "<script>alert('���� �ſ�ī�� �ŷ��� ��� �����մϴ�.');</script>"
'    response.write "<script>window.close();</script>"
'    dbget.close()	:	response.End
'end if

Dim returnmethod, IsCardPartialCancel
returnmethod = orefund.FOneItem.Freturnmethod

if Not ((returnmethod="R100") or (returnmethod="R020") or (returnmethod="R400")) Then
    if (IsAutoScript) then
        response.write "S_ERR|�ſ�ī�� ��ü���, �ǽð���ü ���, �޴��� ��ü ��� �� ����."
    else
        response.write "<script>alert('�ſ�ī�� ��ü���, �ǽð���ü ���, �޴��� ��ü ��� �� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if


'' IniPay �� ��Ҹ� ����
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
        response.write "S_ERR|�̴Ͻý� �ŷ��� ��� �����մϴ�."
    else
        response.write "<script>alert('�̴Ͻý� �ŷ��� ��� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

''=============��ü��Ҹ� ������.. �κ���ҵ� ��Ҿȵ�..=============
dim sqlStr, isSameMoney
dim t_refundrequire, t_MaybeOrgPayPrice
isSameMoney = false

''���̳ʽ� �ֹ��ϰ�� ���ֹ���ȣ// ===> ���ֹ������� �����..
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


'// ��ȯ�ֹ�( jumundiv = 6 )�̸� ���ֹ����� �������� �����´�.
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


'''2011-04 ���� tbl_order_paymentEtc ���.
sqlStr = " select Sum(acctamount) as acctamount"
sqlStr = sqlStr & " from db_order.dbo.tbl_order_paymentEtc"
sqlStr = sqlStr & " where orderserial='"&orgOrderserial&"'"
sqlStr = sqlStr & " and acctdiv in ('100','110')"    ''�ſ�ī�� �� OkCashBag�� ���̰�����.
rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
    t_MaybeOrgPayPrice=rsget("acctamount")
    isSameMoney    = (t_refundrequire=(t_MaybeOrgPayPrice))
end if
rsget.Close

IF  (Not isSameMoney) THEN
    IF (force="on") then
        response.write "��ұݾװ� ���ݾ� ����<br><br>"
    ELSE
        if (IsAutoScript) then
            response.write "S_ERR|��ұݾװ� ���ݾ� ����"
        else
            response.write "<script>alert('��ұݾװ� ���ݾ� ���� - ������ ���� ���."&t_refundrequire&":"&t_MaybeOrgPayPrice&"');</script>"
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

''�޴��� ���� �߰� 2015/04/21 IsINIMobile
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

''īī������
Dim IsKakaoPay : IsKakaoPay = (orefund.FOneItem.Freturnmethod = "R100") and ((Left(orefund.FOneItem.FpaygateTid,3)="cns") or (Left(orefund.FOneItem.FpaygateTid,5)="KCTEN")) ''�ϴ�.

'############################################################## �ڵ��� ���� ��� ##############################################################
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
    		ResultMsg	= "����ó��"
    	Else
    		ResultCode = retval
    		Select Case ResultCode
    			Case "14"
    				ResultMsg = "����"
    			Case "20"
    				ResultMsg = "�޴��� ������� ����(PG��) (LGT�� ��� ������������濡 ���� ��������)"
    			Case "41"
    				ResultMsg = "�ŷ����� ������"
    			Case "42"
    				ResultMsg = "��ұⰣ���"
    			Case "43"
    				ResultMsg = "���γ������� ( ������������ ����ġ, ���ι�ȣ ��ȿ�ð� �ʰ�( 3�� ) )"
    			Case "44"
    				ResultMsg = "�ߺ� ��� ��û"
    			Case "45"
    				ResultMsg = "��� ��û �� ��� ���� ����ġ"
    			Case "97"
    				ResultMsg = "��û�ڷ� ����"
    			Case "98"
    				ResultMsg = "��Ż� ��ſ���"
    			Case "99"
    				ResultMsg = "��Ÿ"
    			Case "11"
    				ResultMsg = "��������������� ���� ��ҺҰ�(11)"
    			Case Else
    				ResultMsg = ""
    		End Select
    	End If

    	CancelDate	= year(now) & "�� " & month(now) & "�� " & day(now) & "��"
    	CancelTime	= hour(now) & "�� " & minute(now) & "�� " & second(now) & "��"
    END IF
ELSEIF (IsKakaoPay) then
    CALL CanCelKakaoPay(orefund.FOneItem.FpaygateTid,orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime)
Else
'############################################################## ī��, �ǽð� ���� ��� ##############################################################
		'###############################################################################
		'# 1. ��ü ���� #
		'################
		Set INIpay = Server.CreateObject("INItx41.INItx41.1")


		'###############################################################################
		'# 2. �ν��Ͻ� �ʱ�ȭ #
		'######################
		PInst = INIpay.Initialize("")

		'###############################################################################
		'# 3. �ŷ� ���� ���� #
		'#####################
		INIpay.SetActionType CLng(PInst), "CANCEL"

		'###############################################################################
		'# 4. ���� ���� #
		'################
		INIpay.SetField CLng(PInst), "pgid", "IniTechPG_" 'PG ID (����)
		INIpay.SetField CLng(PInst), "spgip", "203.238.3.10" '���� PG IP (����)
		INIpay.SetField CLng(PInst), "mid", MctID '�������̵�
		INIpay.SetField CLng(PInst), "admin", "1111" 'Ű�н�����(�������̵� ���� ����)
		INIpay.SetField CLng(PInst), "tid", Request("tid") '����� �ŷ���ȣ(TID)
		INIpay.SetField CLng(PInst), "msg", msg '��� ����
		INIpay.SetField CLng(PInst), "uip", Request.ServerVariables("REMOTE_ADDR") 'IP
		INIpay.SetField CLng(PInst), "debug", "false" '�α׸��("true"�� �����ϸ� ���� �α׸� ����)
		INIpay.SetField CLng(PInst), "merchantreserved", "����" '����

		'###############################################################################
		'# 5. ��� ��û #
		'################
		INIpay.StartAction(CLng(PInst))

		'###############################################################################
		'# 6. ��� ��� #
		'################
		ResultCode = INIpay.GetResult(CLng(PInst), "resultcode") '����ڵ� ("00"�̸� ��Ҽ���)
		ResultMsg  = INIpay.GetResult(CLng(PInst), "resultmsg") '�������
		CancelDate = INIpay.GetResult(CLng(PInst), "pgcanceldate") '�̴Ͻý� ��ҳ�¥
		CancelTime = INIpay.GetResult(CLng(PInst), "pgcanceltime") '�̴Ͻý� ��ҽð�
		Rcash_cancel_noappl = INIpay.GetResult(CLng(PInst), "rcash_cancel_noappl") '���ݿ����� ��� ���ι�ȣ

		'###############################################################################
		'# 7. �ν��Ͻ� ���� #
		'####################
		INIpay.Destroy CLng(PInst)
End If




dim refundrequire,refundresult,userid
dim iorderserial, ibuyhp
dim contents_finish

contents_finish = "��� " & "[" & ResultCode & "]" & ResultMsg & VbCrlf
contents_finish = contents_finish & "����Ͻ� : " & CancelDate & " " & CancelTime & VbCrlf
contents_finish = contents_finish & "����� ID " & finishuserid

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

	'// OKĳ�ù� ������ ���, ��ǰ �� ���̳ʽ� �ֹ� �Է� �� ī�� ��ü����̸� ���̳ʽ� �ֹ��� ���������ݾ� �Է�
	if (accountdiv="110") then ''2015/08/05
        sqlStr = " exec [db_order].[dbo].[usp_Ten_AddEtcPaymentWhenCardCancel] '" + CStr(orgOrderserial) + "', '" + CStr(chgOrderserial) + "'"
        dbget.Execute sqlStr
    end if

    Call AddCustomerOpenContents(id, "ȯ��(���) �Ϸ�: " & CStr(refundrequire))


    Call FinishCSMaster(id, finishuserid, contents_finish)

    ''���� ��� ��û SMS �߼�
    if (iorderserial<>"") and (ibuyhp<>"") then
        SendAcctCancelMsg ibuyhp, iorderserial
'		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
'		sqlStr = sqlStr + " values('" + ibuyhp + "',"
'    	sqlStr = sqlStr + " '1644-6030',"
'    	sqlStr = sqlStr + " '1',"
'    	sqlStr = sqlStr + " getdate(),"
'    	sqlStr = sqlStr + " '[�ٹ�����]���� ��� �Ǿ����ϴ�. �ֹ���ȣ : " + iorderserial + "')"
'		dbget.Execute sqlStr
    end if

    ''����
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
