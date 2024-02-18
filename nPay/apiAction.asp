<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Response.CharSet="UTF-8"
%>
<!-- #include virtual="/lib/util/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/nPay/incNaverPayCommon.asp" -->
<script language="jscript" runat="server">
    function jsURLDecode(v){ return decodeURI(v); }
    function jsURLEncode(v){ return encodeURI(v); }
</script>

<%

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("192.168.1.70","192.168.1.71","192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
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

''''-------------------------------------------------------------
dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    response.write ref
    response.end
end if


Dim sqlStr
Dim QueIdx
QueIdx = Trim(requestCheckVar(request("QueIdx"),10))

if (QueIdx="") then
    dbget.close: response.end
end if

dim paymentId, apiAction, orderserial
sqlStr="select * from db_cs.dbo.tbl_Npay_Que where QueIdx="&QueIdx
rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
if Not rsget.Eof then
    paymentId   = rsget("paymentId")
    apiAction   = rsget("apiAction")
    orderserial   = rsget("orderserial")
end if
rsget.Close

'CALL debugWrite(paymentId)
'CALL debugWrite(apiAction)
'CALL debugWrite(NPay_API_URL)


Dim iErrMsg,inpointCashAmount,iprimaryCashAmount,iprimaryMeans,itotalCashAmount,isupplyCashAmount,ivatCashAmount

dim NPay_Result, iResultCode, iResultMsg
dim confirmTime, iCancelDate, iCancelTime
dim retVal

if (apiAction="FIN") then
    Set NPay_Result = fnCallNaverPayDlvFinish(paymentId)
    
    if NPay_Result.code="Success" then
        iResultCode = "00"                      ''00 사용할것.
        iResultMsg = replace(NPay_Result.message,"'","")
        if (iResultMsg="") then iResultMsg = NPay_Result.code
            
        confirmTime = NPay_Result.body.confirmTime
        
        'CALL debugWrite("confirmTime:"&confirmTime)
        
        iCancelDate = LEFT(confirmTime,4)&"-"&MID(confirmTime,5,2)&"-"&MID(confirmTime,7,2)
        iCancelTime = MID(confirmTime,9,2)&":"&MID(confirmTime,11,2)&":"&MID(confirmTime,13,2)
        
    else
        iResultCode = NPay_Result.code
        iResultMsg = replace(NPay_Result.message,"'","")
        
        retVal = "S_FAIL||"&apiAction&"||"&paymentId ''iResultMsg
    end if
    
    sqlStr =" update Q" & vbCRLF
    sqlStr = sqlStr&" SET findate=getdate()" & vbCRLF
    sqlStr = sqlStr&" ,resultCode='"&iResultCode&"'" & vbCRLF
    sqlStr = sqlStr&" ,resultEtc=convert(varchar(100),'"&iResultMsg&"')" & vbCRLF
    sqlStr = sqlStr&" from db_cs.dbo.tbl_Npay_Que Q where QueIdx="&QueIdx
    dbget.Execute sqlStr
    
    Set NPay_Result = Nothing
    
    if (retVal="") then
        response.write "S_OK||"&apiAction&"||"&paymentId
    else
        response.write retVal
    end if
    
    ''현금영수증 조회 큐 넣음.
    sqlStr =" insert into db_cs.dbo.tbl_Npay_Que " & vbCRLF
    sqlStr = sqlStr&" (paymentid,orderserial,apiaction)" & vbCRLF
    sqlStr = sqlStr&"  select paymentid,orderserial,'CASHAMT'" & vbCRLF
    sqlStr = sqlStr&"  from db_cs.dbo.tbl_Npay_Que Q " & vbCRLF
    sqlStr = sqlStr&"  where QueIdx="& QueIdx & vbCRLF
    dbget.Execute sqlStr
    
elseif (apiAction="CASHAMT") then
    Set NPay_Result = fnCallNaverPayCashAmt(paymentId)
    
    if NPay_Result.code="Success" then
        iResultCode = "00"                      ''00 사용할것.
        iResultMsg = replace(NPay_Result.message,"'","")
        if (iResultMsg="") then iResultMsg = NPay_Result.code
        
        itotalCashAmount  =NPay_Result.body.totalCashAmount
        isupplyCashAmount =NPay_Result.body.supplyCashAmount
        
        
        sqlStr =" update Q" & vbCRLF
        sqlStr = sqlStr&" SET findate=getdate()" & vbCRLF
        sqlStr = sqlStr&" ,resultCode='"&iResultCode&"'" & vbCRLF
        sqlStr = sqlStr&" ,resultEtc=convert(varchar(100),'"&iResultMsg&"')" & vbCRLF
        sqlStr = sqlStr&" ,resultAmt="&itotalCashAmount& vbCRLF
        sqlStr = sqlStr&" ,resultSupp="&isupplyCashAmount& vbCRLF
        sqlStr = sqlStr&" from db_cs.dbo.tbl_Npay_Que Q where QueIdx="&QueIdx
        dbget.Execute sqlStr
        
        Set NPay_Result = Nothing
        
        response.write "S_OK||"&apiAction&"||"&paymentId&"_"&itotalCashAmount&"_"&isupplyCashAmount
        'response.write "totalCashAmount:"&itotalCashAmount
        'response.write "supplyCashAmount:"&isupplyCashAmount
        
        if (orderserial<>"") then
            sqlStr =" update C" & vbCRLF
            sqlStr = sqlStr&" SET ConfirmCashAmt=("&itotalCashAmount&"+isNULL(m.sumpaymentEtc,0))"& vbCRLF
            sqlStr = sqlStr&" ,ConfirmCashSupp=("&isupplyCashAmount&"+convert(int,isNULL(m.sumpaymentEtc,0)*10/11))"& vbCRLF
            sqlStr = sqlStr&" from db_log.dbo.tbl_cash_receipt C"& vbCRLF
            sqlStr = sqlStr&"   Join db_order.dbo.tbl_order_master m"& vbCRLF
	        sqlStr = sqlStr&"   on C.orderserial=m.orderserial"& vbCRLF
	        sqlStr = sqlStr&"   and m.cancelyn='N' "& vbCRLF
	        sqlStr = sqlStr&"   and m.orderserial='"&orderserial&"'"& vbCRLF
            sqlStr = sqlStr&" where C.orderserial='"&orderserial&"'"& vbCRLF
            sqlStr = sqlStr&" and C.cancelyn='N'"& vbCRLF
            sqlStr = sqlStr&" and C.resultCode='R'"& vbCRLF
            dbget.Execute sqlStr
            
            
            sqlStr =" update C" & vbCRLF
            sqlStr = sqlStr&" SET cr_price=C.ConfirmCashAmt" & vbCRLF
            sqlStr = sqlStr&" ,sup_price=c.ConfirmCashSupp" & vbCRLF
	        sqlStr = sqlStr&" ,tax=c.ConfirmCashAmt-c.ConfirmCashSupp" & vbCRLF
            sqlStr = sqlStr&" from db_log.dbo.tbl_cash_receipt C"& vbCRLF
            sqlStr = sqlStr&" where C.orderserial='"&orderserial&"'"& vbCRLF
            sqlStr = sqlStr&" and C.cancelyn='N'"& vbCRLF
            sqlStr = sqlStr&" and C.resultCode='R'"& vbCRLF
            sqlStr = sqlStr&" and ("& vbCRLF
            sqlStr = sqlStr&"   (cr_price<>isNULL(C.ConfirmCashAmt,0))"& vbCRLF
            sqlStr = sqlStr&"    or (sup_price<>isNULL(C.ConfirmCashSupp,0))"& vbCRLF
            sqlStr = sqlStr&" )"& vbCRLF

            dbget.Execute sqlStr

            ''절삭관련. 2016/08/22
            ' sqlStr =" update C" & vbCRLF
            ' sqlStr = sqlStr&" SET sup_price=c.ConfirmCashSupp" & vbCRLF
	        ' sqlStr = sqlStr&" ,tax=c.ConfirmCashAmt-c.ConfirmCashSupp" & vbCRLF
            ' sqlStr = sqlStr&" from db_log.dbo.tbl_cash_receipt C"& vbCRLF
            ' sqlStr = sqlStr&" where C.orderserial='"&orderserial&"'"& vbCRLF
            ' sqlStr = sqlStr&" and C.cancelyn='N'"& vbCRLF
            ' sqlStr = sqlStr&" and C.resultCode='R'"& vbCRLF
            ' sqlStr = sqlStr&" and C.cr_price=C.ConfirmCashAmt"& vbCRLF
            ' sqlStr = sqlStr&" and ABS(C.sup_price-isNULL(c.ConfirmCashSupp,0))=1"& vbCRLF
            ' dbget.Execute sqlStr
        end if
    else
        iResultCode = NPay_Result.code
        iResultMsg = replace(NPay_Result.message,"'","")
        
        retVal = "S_FAIL||"&iResultCode&"||"&iResultMsg
        response.write retVal
    end if
elseif (apiAction="XXX") then
    ''NPay_Result = fnCallNaverListOfPayment("","20160629","20160629",iErrMsg)
    CALL debugWrite("apiAction|"&apiAction)
    dbget.Close() : response.end
else
    CALL debugWrite("apiAction|"&apiAction)
    dbget.Close() : response.end
end if


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->