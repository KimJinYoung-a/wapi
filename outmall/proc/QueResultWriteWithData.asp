<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
''UTF-8 로 해야 한글이 안깨지게 받음.

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("222.109.123.95","211.206.236.117","115.94.163.42","61.252.133.88","192.168.1.70","192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next

    dim validToken : validToken = Array("bd2acd564c264459908cd0d744986ea0","554740c796aa47b1aae8ee9bacd2643c")
    dim authtkn : authtkn = LCASE(request("authtkn"))
    for i=0 to UBound(validToken)
        if (validToken(i)=authtkn) then
            CheckVaildIP = true
            exit function
        end if
    next
end function


dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    response.write ref
    response.end
end if


dim idx     : Idx=requestCheckVar(request("idx"),10)
dim itemid  : itemid=requestCheckVar(request("itemid"),10)
dim ErrCode : ErrCode=requestCheckVar(request("ErrCode"),32)
dim ErrMsg  : ErrMsg=requestCheckVar(request("ErrMsg"),500)

dim sellyn  : sellyn=requestCheckVar(request("sellyn"),10)                   '' mall 전송 판매 상태 Y/N/S
dim salePrice : salePrice=requestCheckVar(request("salePrice"),10)           '' mall 전송 판매 가격  13000
dim regitemname : regitemname=requestCheckVar(request("regitemname"),100)      '' mall 전송 상품명 문주란다이어리.

dim mode : mode=requestCheckVar(request("mode"),16) 
dim sellsite : sellsite=requestCheckVar(request("sellsite"),32) 
dim mallitemid : mallitemid=requestCheckVar(request("mallitemid"),32) 
dim apiAction : apiAction=requestCheckVar(request("apiAction"),32) 
dim orgprice : orgprice=requestCheckVar(request("orgprice"),10)
dim mallregdt : mallregdt=requestCheckVar(request("mallregdt"),19)
dim malloptioncnt : malloptioncnt=requestCheckVar(request("malloptioncnt"),10)
dim optdataarr : optdataarr=request("optdataarr")
dim optoneLine, optname,optttlstock, optsaledstock, optremainstock,malloptcode,optcode,optdisp, optstockexists
dim sellerProductItemId, vendorItemId, mallsellprice, mallsellyn

dim sqlStr, i, j

if (apiAction="CHKSTAT") then
    if (mode="OPTSTAT") then
        optdataArr = split(optdataarr,vbCRLF)
        if IsArray(optdataArr) then
            for i=Lbound(optdataArr) to Ubound(optdataArr)
                optoneLine = Trim(optdataArr(i))
                if (optoneLine<>"") then
                    optoneLine = split(optoneLine,"|")
                    if IsArray(optoneLine) then
                        malloptcode = ""
                        for j=Lbound(optoneLine) to Ubound(optoneLine)
                            if (j=0) then malloptcode     = Trim(optoneLine(j))
                            if (j=1) then optname         = Trim(optoneLine(j))
                            if (j=2) then optttlstock     = Trim(optoneLine(j))
                            if (j=3) then optsaledstock   = Trim(optoneLine(j))
                            if (j=4) then optremainstock  = Trim(optoneLine(j))
                            if (j=5) then optcode         = Trim(optoneLine(j))
                            if (j=6) then optdisp         = Trim(optoneLine(j))
                            if (j=7) then optstockexists  = Trim(optoneLine(j))
                        next
                        ''response.write optremainstock&"|"&optcode
                        if (optdisp="전시") then 
                            optdisp="Y"
                        else
                            optdisp="N"
                        end if
                        
                        if (optstockexists="재고있음") then 
                            optstockexists="Y"
                        else
                            optstockexists="N"
                        end if
                        
                        if (optremainstock="무제한") then optremainstock="9999"
                        
                        if (malloptcode<>"") then
                            if Not IsNumeric(itemid) then 
                                response.write "S_ERR"
                                dbget.close()
                                response.end
                            else
                                sqlStr = "db_etcmall.[dbo].[sp_Ten_OutMall_API_SaleStatWriteOption] "&itemid&",'"&mallitemid&"','"&sellsite&"','"&malloptcode&"','"&optcode&"',"&optremainstock&",'"&optdisp&"','"&optstockexists&"','"&html2db(optname)&"'"
                                dbget.Execute sqlStr
                            end if
                        end if
                    end if
                end if
            next

            if (sellsite="kakaogift") then
                sqlStr = "db_etcmall.[dbo].[sp_Ten_OutMall_API_SaleStatUpdateByOptionSummary] "&itemid&",'"&mallitemid&"','"&sellsite&"'"
                dbget.Execute sqlStr
            end if
        end if
    elseif (mode="OPTSTAT2") then
        optdataArr = split(optdataarr,"@!@!")  ''vbCRLF
        if IsArray(optdataArr) then
            for i=Lbound(optdataArr) to Ubound(optdataArr)
                optoneLine = Trim(optdataArr(i))
                if (optoneLine<>"") then
                    optoneLine = split(optoneLine,"|")
                    if IsArray(optoneLine) then
                        vendorItemId = ""
                        optremainstock = 0
                        mallsellprice = 0
                        mallsellyn = "Y"
                        sellerProductItemId = ""
                        optcode = ""
                        optname = ""

                        for j=Lbound(optoneLine) to Ubound(optoneLine)
                            ''7135|3747143985|9999|33000.0|True|730157346|[ART]Hello RainCats 캣앤독 자동우산 1 아이보리(ivory)\r\n
                            ''7149|3747144028|9999|33000.0|True|730157347|[ART]Hello RainCats 캣앤독 자동우산 1 화이트(white)'}
                            if (j=0) then optcode    = Trim(optoneLine(j))
                            if (j=1) then vendorItemId    = Trim(optoneLine(j))
                            if (j=2) then optremainstock  = Trim(optoneLine(j))
                            if (j=3) then mallsellprice   = Trim(optoneLine(j))
                            if (j=4) then mallsellyn      = Trim(optoneLine(j))
                            if (j=5) then sellerProductItemId      = Trim(optoneLine(j))
                            if (j=6) then optname         = Trim(optoneLine(j))
                        next

                        ''response.write optremainstock&"|"&optcode
                        if (LCASE(mallsellyn)="true") then 
                            mallsellyn="Y"
                        else
                            mallsellyn="N"
                        end if
                        
                        mallsellprice = CLNG(mallsellprice)

                        if (optcode<>"") and (vendorItemId<>"") and (vendorItemId<>"") then
                            if Not IsNumeric(itemid) then 
                                response.write "S_ERR"
                                dbget.close()
                                response.end
                            else
                                ''@itemid	,@itemoption ,@vendorItemId ,@sellerProductItemId ,@mallid 
	                            '',@outmallOptName ,@outmallSellyn,@outmalllimityn,@outmalllimitno,@outmallAddPrice,@outmallSellPrice 
                                sqlStr = "db_etcmall.[dbo].[sp_Ten_OutMall_API_SaleStatWriteOption_coupang] "&itemid&",'"&optcode&"','"&vendorItemId&"','"&sellerProductItemId&"','"&sellsite&"','"&replace(optname,"'","''")&"','"&mallsellyn&"','Y','"&optremainstock&"',0,'"&mallsellprice&"'"
                                dbget.Execute sqlStr
                            end if
                        end if
                    end if
                end if
            next

            if (sellsite="coupang") then
                sqlStr = "db_etcmall.[dbo].[sp_Ten_OutMall_API_SaleStatUpdateByOptionSummary] "&itemid&",'"&mallitemid&"','"&sellsite&"'"
                dbget.Execute sqlStr
            end if
        end if

    else
        if Not IsNumeric(itemid) then 
            response.write "S_ERR"
            dbget.close()
            response.end
        else
            sqlStr = "db_etcmall.[dbo].[sp_Ten_OutMall_API_SaleStatWrite] "&itemid&",'"&mallitemid&"','"&sellsite&"','"&apiAction&"','"&sellyn&"',"&orgprice&","&salePrice&",'"&regitemname&"','"&mallregdt&"',"&malloptioncnt&","&idx&",'"&ErrCode&"','"&ErrMsg&"'"
            'response.write sqlStr
            dbget.Execute sqlStr
        end if
    end if

else
    sqlStr = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWriteWithData] "&idx&","&itemid&",'"&sellyn&"',"&salePrice&",'"&html2DB(regitemname)&"','"&ErrCode&"','"&html2DB(ErrMsg)&"'"
    'response.write sqlStr
    dbget.Execute sqlStr
end if



response.write "S_OK"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->