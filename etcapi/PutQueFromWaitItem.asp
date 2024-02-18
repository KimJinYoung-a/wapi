<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/function.asp" -->
<%
'접근ip 확인
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("172.16.0.225","192.168.1.71","::1")  '젠킨스,테스트(kobula),로컬
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
    response.write "S_ERR"
    dbget.Close() :     response.end
end if


dim apiAction : apiAction=requestCheckVar(request("apiAction"),32)
dim dispCate  : dispCate = requestCheckVar(request("dispCate"),16)
dim makerid   : makerid = requestCheckVar(request("makerid"),32)
dim currstate : currstate = requestCheckVar(request("currstate"),1)
dim itemid    : itemid = requestCheckVar(request("itemid"),12)
dim startNo   : startNo = requestCheckVar(request("startNo"),8)
dim endNo     : endNo = requestCheckVar(request("endNo"),8)
dim ctrState  : ctrState = requestCheckVar(request("ctrState"),1)

startNo = 1
if endNo="" then endNo = 100
currstate = "W"     '최근 30분 이전 승인대기 상품
ctrState = "Y"      '계약완료 업체만

dim sqlStr

if (apiAction="readToSetWaitItem") then

    sqlStr = "db_temp.dbo.sp_Ten_wait_item_getItemList4Que '"&dispCate&"','"&makerid&"','"&currstate&"','"&itemid&"','"&startNo&"','"&endNo&"','"&ctrState&"'"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly
    if Not rsget.Eof then
        do until rsget.eof
            'GET으로 전송
            Call SendReqGet("http://110.93.128.100:8090/scmapi/nsqmessage/singlecontaincollect","waititemid=" & rsget("itemid") & "&scmid=" & rsget("makerid"))
            rsget.MoveNext
        loop
    end if
    rsget.close

    response.write "S_OK"
else
    response.write "S_ERR|Not Valid - "&apiAction
end if


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->