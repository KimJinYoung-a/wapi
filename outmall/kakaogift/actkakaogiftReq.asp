<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim arrItemid : arrItemid = request("cksel")
Dim i, strParam, iErrStr, ret1
Dim sqlStr, strSql, AssignedRow, SubNodes
Dim chgSellYn, actCnt, retErrStr
Dim buf, buf2, CNT10, CNT20, CNT30, iitemid
Dim ArrRows
Dim retFlag
Dim iMessage
dim iItemName, pregitemname
Dim ccd
ccd		  = request("CommCD")
retFlag   = request("retFlag")
chgSellYn = request("chgSellYn")
arrItemid = Trim(arrItemid)

dim imallid : imallid = "kakaogift"
dim iapiaction : iapiaction = ""
dim ilastUpdateid : ilastUpdateid = session("ssBctID")

'' CHKSTAT : 상태체크
'' EDPRSL : 수정 (현상태로맞게)

if (cmdparam="EditSellYn") and (chgSellYn="N") then
    iapiaction = "SELLN"
elseif (cmdparam="EditSellYn") and (chgSellYn="N") then
    iapiaction = "SELLY"
elseif (cmdparam="EDIT") then
    iapiaction = "EDPRSL"
elseif (cmdparam="CHK") then
    iapiaction = "CHKSTAT"
else
    response.write "<script>alert('미지정:"&cmdparam&"');</script>"
    dbget.close() : response.end
end if
    
dim iarrItemids : iarrItemids = split(arrItemid,",")
if isArray(iarrItemids) then    
    for i=LBound(iarrItemids) to Ubound(iarrItemids)
        iitemid = iarrItemids(i)
        if (iitemid<>"") then
            strSQL = ""
            strSQL = strSQL & " INSERT INTO [db_etcmall].[dbo].[tbl_outmall_API_Que] (mallid, apiAction, itemid, priority, regdate,lastUserid) VAlUES " & VBCRLF
            strSQL = strSQL & " ('"& imallid &"', '"& iapiaction &"', '"& iitemid &"', '999999', getdate(), '"& ilastUpdateid &"') " & VBCRLF
            dbget.Execute strSQL
        end if
    next
end if

response.write "cmdparam:"&cmdparam&"<br>"
response.write "ccd:"&ccd&"<br>"
response.write "retFlag:"&retFlag&"<br>"
response.write "chgSellYn:"&chgSellYn&"<br>"
response.write "arrItemid:"&arrItemid

response.write"<script>alert('수정요청되었습니다.');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
