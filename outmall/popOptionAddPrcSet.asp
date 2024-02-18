<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/popheader.asp"-->
<%
dim itemid : itemid = requestCheckvar(request("itemid"),10)
dim mallid  : mallid = requestCheckvar(request("mallid"),32)
dim mode    : mode = requestCheckvar(request("mode"),32)
dim mngOptAdd : mngOptAdd = requestCheckvar(request("mngOptAdd"),10)
dim optAddPrcRegType: optAddPrcRegType = requestCheckvar(request("optAddPrcRegType"),10)

dim sqlStr, arrRows
dim regitemname,outmallGoodNo,optaddPrcCnt, lastStatCheckDate
dim TEN_URI, OUT_URI
dim i

'rw itemid
'rw mallid

if (mode="EDTRegType") then
    if (mallid="lotteCom") then
        sqlStr = ""
        sqlStr = sqlStr & " update R"&VbCRLF
        sqlStr = sqlStr & " set optAddPrcRegType="&optAddPrcRegType&VbCRLF
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_lotte_regItem R"&VbCRLF
        sqlStr = sqlStr & " where R.itemid="&itemid&VbCRLF

        dbget.Execute sqlStr
    elseif (mallid="lotteimall") then
        sqlStr = ""
        sqlStr = sqlStr & " update R"&VbCRLF
        sqlStr = sqlStr & " set optAddPrcRegType="&optAddPrcRegType&VbCRLF
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_Ltimall_regItem R"&VbCRLF
        sqlStr = sqlStr & " where R.itemid="&itemid&VbCRLF

        dbget.Execute sqlStr
    else
        rw "미지정["&mallid&"]"
        dbget.Close() : response.end
    end if
end if



if (mallid="cjmall") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, '' as outmallregName, cjmallprdno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_outmall].[dbo].tbl_cjmall_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
else
    rw "미지정["&mallid&"]"
    dbCTget.Close() : response.end
end if
rsCTget.Open sqlStr, dbCTget
if Not(rsCTget.EOF or rsCTget.BOF) then
    regitemname     = rsCTget("outmallregName")
    outmallGoodNo   = rsCTget("outmallGoodNo")
    optaddPrcCnt    = rsCTget("optaddPrcCnt")
    optAddPrcRegType = rsCTget("optAddPrcRegType")
    lastStatCheckDate= rsCTget("lastStatCheckDate")
end if
rsCTget.close


sqlStr = ""
sqlStr = sqlStr & " select top 300 '1' as Sitename,o.itemid,o.itemoption,o.isusing,o.optsellyn"&VbCRLF
sqlStr = sqlStr & " ,o.optlimityn,o.optlimitno,o.optlimitsold"&VbCRLF
sqlStr = sqlStr & " ,o.optionname,o.optaddprice,o.optaddbuyprice"&VbCRLF
sqlStr = sqlStr & " ,o.optionTypeName"&VbCRLF
sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
sqlStr = sqlStr & " from [db_Appwish].[dbo].tbl_item_option o"&VbCRLF
sqlStr = sqlStr & " 	left join [db_outmall].[dbo].tbl_OutMall_regedoption r"&VbCRLF
sqlStr = sqlStr & " 	on r.mallid='"&mallid&"'"&VbCRLF
sqlStr = sqlStr & " 	and o.itemid=r.itemid"&VbCRLF
sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
sqlStr = sqlStr & " where o.itemid="&itemid&VbCRLF
sqlStr = sqlStr & " Union"&VbCRLF
sqlStr = sqlStr & " select top 100 'Z' as Sitename,r.itemid,r.itemoption,NULL as isusing,NULL as optsellyn"&VbCRLF
sqlStr = sqlStr & " ,NULL as optlimityn,NULL as optlimitno,NULL as optlimitsold"&VbCRLF
sqlStr = sqlStr & " ,NULL as optionname,NULL as optaddprice,NULL as optaddbuyprice"&VbCRLF
sqlStr = sqlStr & " ,NULL as optionTypeName"
sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
sqlStr = sqlStr & " from [db_outmall].[dbo].tbl_OutMall_regedoption r"&VbCRLF
sqlStr = sqlStr & " 	left join [db_Appwish].[dbo].tbl_item_option o"&VbCRLF
sqlStr = sqlStr & " 	on o.itemid=r.itemid"&VbCRLF
sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
sqlStr = sqlStr & " where r.mallid='"&mallid&"'"&VbCRLF
sqlStr = sqlStr & " and r.itemid="&itemid&VbCRLF
sqlStr = sqlStr & " and o.itemid is NULL"&VbCRLF
sqlStr = sqlStr & " order by Sitename"&VbCRLF
''rw sqlStr
rsCTget.Open sqlStr, dbCTget
if Not(rsCTget.EOF or rsCTget.BOF) then
    arrRows = rsCTget.getRows
end if
rsCTget.close

TEN_URI = "http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid
if (mallid="lotteCom") then
    OUT_URI = "http://www.lotte.com/goods/viewGoodsDetail.lotte?goods_no="&outmallGoodNo
elseif (mallid="lotteimall") then
    OUT_URI = "http://www.lotteimall.com/product/Product.jsp?i_code="&outmallGoodNo
elseif (mallid="cjmall") then
    OUT_URI = "http://www.cjmall.com/prd/detail_cate.jsp?item_cd="&outmallGoodNo
elseif (mallid="gsshop") then
    OUT_URI = "http://www.gsshop.com/prd/prd.gs?prdid="&outmallGoodNo
elseif (mallid="homeplus") then
    OUT_URI = "http://direct.homeplus.co.kr/app.product.Product.ghs?comm=usr.product.detail&i_style="&outmallGoodNo
else
    OUT_URI = ""
end if

dim isOptSoldOut, isOutOptSoldOut
dim isLimit, isOutLimit
dim limitno, outLimitno

Dim isOptAddPriceExistsItem : isOptAddPriceExistsItem = false

if isArray(arrRows) then
For i =0 To UBound(ArrRows,2)
    isOptAddPriceExistsItem = (isOptAddPriceExistsItem or ArrRows(9,i)>0)
Next
end if
%>
<script language='javascript'>
function saveThis(comp){
    if (confirm('수정 하시겠습니까?')){
        comp.form.submit();
    }
}

function refreshSellStat(itemid,mallid){
    if (mallid=="lotteCom"){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "ChkStockSelect";
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "/admin/etc/lotte/actRegLotteItem.asp"
        document.frmSvArr.retFlag.value="retFunc(2)";
        document.frmSvArr.submit();
    }else if (mallid=="cjmall"){
        alert('cjMall 단품 재수신의 경우 5~10분정도 지연됨');
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "confirmItem";                       //조회
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "/outmall/cjmall/actCjmallReq.asp"
        document.frmSvArr.retFlag.value="retFunc(2)";
        document.frmSvArr.submit();
   }else if (mallid=="lotteimall"){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "ChkStockSelect";
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "/admin/etc/LtiMall/actLotteiMallReq.asp"
        document.frmSvArr.retFlag.value="retFunc(2)";
        document.frmSvArr.submit();
   }
}

function outItemDtlProc(itemid,mallid){
    if (mallid=="lotteCom"){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditSelect";
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "/admin/etc/lotte/actRegLotteItem.asp"
        document.frmSvArr.retFlag.value="retFunc(1)";
        document.frmSvArr.submit();
    }else if (mallid=="cjmall"){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSelect2";                       //가격/한정수량/단품수정
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "/outmall/cjmall/actCjmallReq.asp"
        document.frmSvArr.retFlag.value="retFunc(1)";
        document.frmSvArr.submit();
    }else if (mallid=="lotteimall"){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSelect";
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "/admin/etc/LtiMall/actLotteiMallReq.asp"
        document.frmSvArr.retFlag.value="retFunc(1)";
        document.frmSvArr.submit();
    }else if (mallid=="gsshop"){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditOPT";
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "/admin/etc/gsshop/actgsshopReq.asp"
        document.frmSvArr.submit();
    }else if (mallid=="homeplus"){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditItemSelect";
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}

function retFunc(retval){
    if (retval==1){
        refreshSellStat('<%=itemid%>','<%=mallid%>');
    }else if (retval==2){
        alert('ok');
        document.location.reload();
    }
}
</script>
<table width="100%" cellpadding="5" cellspacing="1" bgcolor="#BBBBBB" class="a">
<form name="frmSv" method="post" action="/admin/etc/popOptionAddPrcSet.asp">
<input type="hidden" name="mode" value="EDTRegType">
<input type="hidden" name="itemid" value="<%=itemid %>">
<input type="hidden" name="mallid" value="<%=mallid %>">
<input type="hidden" name="mngOptAdd" value="<%=mngOptAdd %>">
<tr  bgcolor="#FFFFFF" align="center">
    <td bgcolor="#FFFFFF" align="center" colspan="4"><%= mallid %></td>
    <td align="right"><a href="javascript:document.location.reload();"><img src="http://scm.10x10.co.kr/images/icon_reload.gif" border="0"></a></td>
</tr>
<tr  bgcolor="#FFFFFF" align="center">
    <td width="20%" colspan="2"><a href="<%= TEN_URI %>" target=_blank><%= itemid %></a></td>
    <td>
        <%= regitemname %> / <%= lastStatCheckDate %>
    </td>
    <td width="20%" colspan="2"><a href="<%= OUT_URI %>" target=_blank><%= outmallGoodNo %></a></td>
</tr>
<% if (mngOptAdd="1") then %>
<tr>
    <td bgcolor="#FFFFFF" align="center" colspan="5">
        <input type="radio" name="optAddPrcRegType" value="0" <%=CHKIIF(optAddPrcRegType="0","checked","")%> > 미지정(자동품절)
        <input type="radio" name="optAddPrcRegType" value="1" <%=CHKIIF(optAddPrcRegType="1","checked","")%> > 옵션추가금액 없는상품만 판매
        <input type="radio" name="optAddPrcRegType" value="9" disabled > 옵션추가금액 별도 등록
    </td>
</tr>
<tr>
    <td bgcolor="#FFFFFF" align="center" colspan="5">
    <input type="button" value=" 저 장 " onClick="saveThis(this);">
    </td>
</tr>
<% end if %>
</form>
</table>
<p>
<table width="100%" cellpadding="5" cellspacing="1" bgcolor="#BBBBBB" class="a">
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td colspan="7">10x10</td>
    <td width="1" bgcolor="#FFFFFF">
    <input type="button" value=">>" onClick="outItemDtlProc('<%=itemid %>','<%=mallid %>');">
    </td>
    <td colspan="4"><%= mallid %> <% If (mallid <> "gsshop") and (mallid <> "homeplus") Then %> <input type="button" value="단품재수신" class="button" onClick="refreshSellStat('<%=itemid %>','<%=mallid %>')"> <% End If %>
    <br><br>(복합옵션은 정확하지 않음)
    </td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td>상품코드</td>
    <td>옵션코드</td>
    <td>옵션타입</td>
    <td>옵션명</td>
    <td>한정</td>
    <td>판매</td>
    <td>옵션추가액</td>
    <td width="1" bgcolor="#FFFFFF">

    </td>
    <td>옵션명</td>
    <td>옵션코드</td>
    <td>한정</td>
    <td>판매</td>
</tr>
<% if isArray(arrRows) then %>
<% For i =0 To UBound(ArrRows,2) %>
<%
    if isNULL(ArrRows(3,i)) then
        isOptSoldOut = false
    else
        isOptSoldOut = ((ArrRows(3,i)="N") or (ArrRows(4,i)="N") or (((ArrRows(5,i)="Y") and (ArrRows(6,i)-ArrRows(7,i)<1))))
    end if

    if isNULL(ArrRows(5,i)) then
        isLimit = false
        limitno = 0
    else
        isLimit = (ArrRows(5,i)="Y")
        limitno = (ArrRows(6,i)-ArrRows(7,i))

    end if

    if (limitno<1) then limitno=0

%>
<%
    if isNULL(ArrRows(14,i)) or isNULL(ArrRows(15,i)) or isNULL(ArrRows(16,i)) then
        isOutOptSoldOut = false
    else
        isOutOptSoldOut = (ArrRows(14,i)="N") or ((ArrRows(15,i)="Y") and (ArrRows(16,i)<1))
    end if

    if isNULL(ArrRows(15,i)) then
        isOutLimit = false
        outLimitno = 0
    else
        isOutLimit = (ArrRows(15,i)="Y")
        outLimitno = ArrRows(16,i)

    end if

    if (outLimitno<1) then outLimitno=0
%>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(1,i) %></td>
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(2,i) %></td>
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(11,i) %></td>
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(8,i) %></td>
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" >
        <% if isLimit then %>
        <font color="blue"><%= limitno %></font>
        <% end if %>
    </td>
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" >
        <%= CHKIIF(isOptSoldOut,"<font color=red>품절</font>","") %>
    </td>
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(9,i) %></td>
    <td width="1" bgcolor="#FFFFFF"></td>
    <td bgcolor="<%= CHKIIF(isOutOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(13,i) %></td>
    <td bgcolor="<%= CHKIIF(isOutOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(12,i) %></td>
    <td bgcolor="<%= CHKIIF(isOutOptSoldOut,"#EEEEEE","#FFFFFF") %>" >
        <% if isOutLimit then %>
        <font color="blue"><%= outLimitno %></font>
        <% end if %>
    </td>
    <td bgcolor="<%= CHKIIF(isOutOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= CHKIIF(isOutOptSoldOut,"<font color=red>품절</font>","") %></td>
</tr>
<% next %>
<% end if %>

</table>

<% if (mngOptAdd<>"1") then %>
<p>
<table width="100%" cellpadding="5" cellspacing="1" bgcolor="#BBBBBB" class="a">
<tr>
    <td align="center" bgcolor="#FFFFFF" height="20">
    <input type="button" value=" 닫기 " onClick="self.close();">
    </td>
</tr>
</table>
<% end if %>

<form name="frmSvArr" method="post" >
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="cksel" value="">
<input type="hidden" name="retFlag" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="1" width="100%" height="100"></iframe>
<!-- #include virtual="/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->