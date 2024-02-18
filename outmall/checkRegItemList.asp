<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/popheader.asp"-->
<%
function getLastCheckTimByreg(iSellSite)
    dim sqlStr
    sqlStr = "select top 1 convert(varchar(10),GoodsRegDtime,21) as yyyymmdd "&vbCRLF
    sqlStr = sqlStr & " from db_temp.dbo.tbl_XSite_regItemCheck C"&vbCRLF
    sqlStr = sqlStr & " where sellsite='"&iSellSite&"'"&vbCRLF
    sqlStr = sqlStr & " order by GoodsRegDtime desc"&vbCRLF

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        getLastCheckTimByreg = rsget("yyyymmdd")
    end if
    rsget.close
end function

function getMayErrItemList(iSellSite)
    dim sqlStr
    sqlStr = "select top 100 SellSite,goodsno,SaleStatCd,SalePrc,goodsNm,GoodsRegDtime,DispYn,regDtKey,mayTenItemID,maymidx from db_temp.dbo.tbl_XSite_regItemCheck C"&vbCRLF
    sqlStr = sqlStr & " where  sellsite='"&iSellSite&"'"&vbCRLF
    sqlStr = sqlStr & " and mayTenItemid is NULL"&vbCRLF
    sqlStr = sqlStr & " and maymidx is NULL"&vbCRLF
    sqlStr = sqlStr & " and SaleStatCd not in ('N','X')"&vbCRLF
    sqlStr = sqlStr & " order by GoodsRegDtime desc"&vbCRLF

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        getMayErrItemList = rsget.getRows()
    end if
    rsget.close
end function

Dim i
Dim sellsite : sellsite = requestCheckvar(request("sellsite"),32)
Dim arrItemid : arrItemid = requestCheckvar(request("arrItemid"),2000)
Dim cmdparam : cmdparam = requestCheckvar(request("cmdparam"),32)
Dim yyyymmdd1 : yyyymmdd1 = requestCheckvar(request("yyyymmdd1"),10)
Dim yyyymmdd2 : yyyymmdd2 = requestCheckvar(request("yyyymmdd2"),10)
Dim research : research = requestCheckvar(request("research"),10)
Dim chgSellYn

arrItemid = Trim(arrItemid)


if (cmdparam="") then cmdparam="chkitembydate"
if (yyyymmdd2="") then yyyymmdd2=LEFT(dateadd("d",-1,now()),10)
if (yyyymmdd1="") then yyyymmdd1=LEFT(dateadd("d",-2,now()),10)

dim maypredate 
if (research="") then
    maypredate = getLastCheckTimByreg(sellsite)
    yyyymmdd1 = maypredate
end if

dim arrRows : arrRows = getMayErrItemList(sellsite)
Dim bufOutitemArr
%>
<script type="text/javascript" src="/lib/js/jquery-1.7.1.min.js"></script>

<script type="text/javascript">
<% if (FALSE) then %>
	var items = "<%=arrItemid%>";
	var itemArr = items.split(", ");
	var rotation;
	var rno = 0;

	function x_loadRotation() {
		if(itemArr[rno] == undefined){
			alert('완료하였습니다')
			return;
		}
		rotation = arrSubmit(itemArr[rno]);
		rno++;
		if(rno > itemArr.length-1){
			clearTimeout(rotation);
			//setTimeout("alert('완료하였습니다')", 500);
		}else{
			//setTimeout('loadRotation()', 2000);
		}
	}

	function arrSubmit(ino){
		document.frmSvArr.target = "xLink2";
        document.frmSvArr.act.value = "<%=cmdparam%>";
        document.frmSvArr.itemid.value = ino;
        document.frmSvArr.chgSellYn.value = "<%=chgSellYn%>";
        document.frmSvArr.action = '/outmall/lotteCom/lotteComActProc.asp';
        //document.frmSvArr.submit();		
	}
	//window.onload = new Function('setTimeout("loadRotation()", 500)');
<% end if %>

var actyyyymmdd = "";

function loadRotation(){
    var stdt = $("#yyyymmdd1").val();
    var eddt = $("#yyyymmdd2").val();

    if (actyyyymmdd==""){
        actyyyymmdd = stdt;
    }else{
        //actyyyymmdd = dateToYMD((new Date(actyyyymmdd)).setDate((new Date(actyyyymmdd)).getDate()+1));
        actyyyymmdd = dateToYMD(new Date(actyyyymmdd.substring(0,4),actyyyymmdd.substring(5,7)*1-1,actyyyymmdd.substring(8,10)*1+1)); 
    }
    
    if ((new Date(actyyyymmdd))>(new Date(eddt))) {
        alert("FIN");
        return;
    }

    <% if (LCASE(sellsite)="lotteimall") then %>
    document.frmSearchArr.action ="/outmall/ltimall/ltimallActProc.asp";
    <% else %> 
    document.frmSearchArr.action ="/outmall/LotteCom/lotteComActProc.asp";
    <% end if %>
    document.frmSearchArr.target = "xLink2";
    document.frmSearchArr.act.value = "CHKITEMLIST"
    document.frmSearchArr.itemid .value = "-1"
    document.frmSearchArr.yyyymmdd.value = actyyyymmdd;

    document.frmSearchArr.submit();

}

function dateToYMD(date) {
    var d = date.getDate();
    var m = date.getMonth() + 1; //Month from 0 to 11
    var y = date.getFullYear();
    return '' + y + '-' + (m<=9 ? '0' + m : m) + '-' + (d <= 9 ? '0' + d : d);
}

function chkItemByRegDateLotteCom(){
    actyyyymmdd = "";
    $("#actStr").html("");
    loadRotation();
}

function reAct(yyyymmdd){
    $("#yyyymmdd1").val(yyyymmdd);
    $("#yyyymmdd2").val(yyyymmdd);
    chkItemByRegDateLotteCom();
}
</script>
<form name="frmSvArr">
	<input type="hidden" name="act">
	<input type="hidden" name="itemid">
	<input type="hidden" name="chgSellYn">
</form>

<form name="frmSearchArr">
	<input type="hidden" name="act">
    <input type="hidden" name="itemid">
	<input type="hidden" name="yyyymmdd">
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BBBBBB">
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
    site : <%=sellsite%>
    &nbsp;&nbsp;
    <input type="text" name="yyyymmdd1" id="yyyymmdd1" value="<%=yyyymmdd1%>" size="10">
    <input type="text" name="yyyymmdd2" id="yyyymmdd2" value="<%=yyyymmdd2%>" size="10">
    <input type="button" value="조회" onClick="chkItemByRegDateLotteCom()">
    </td>
	<td align="center" width="100">
    <input type="button" value="검색" onClick="location.href='?sellsite=<%=sellsite%>&research=on&yyyymmdd1=<%=yyyymmdd1%>&yyyymmdd2=<%=yyyymmdd2%>'">
    </td>
</tr>
</table>
<p>
<div style="overflow:scroll; width:100%; height:250px; ">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BBBBBB">
<% if IsArray(ArrRows) then %>
<% ''SellSite,goodsno,SaleStatCd,SalePrc,goodsNm,GoodsRegDtime,DispYn,regDtKey,mayTenItemID,maymidx %>
<tr bgcolor="#FFFFFF" >
    <td>SellSite</td>
    <td>제휴상품번호</td>
    <td>판매상태</td>
    <td>판매가</td>
    <td>goodsNm</td>
    <td>상품등록일</td>
    <td>전시YN</td>
    <td>조회일</td>
    <td>상품코드</td>
    <td>옵션Key</td>
    <td>다시조회</td>
</tr>
<% for i=LBound(ArrRows,2) to UBound(ArrRows,2) %>
<%
bufOutitemArr = bufOutitemArr&ArrRows(1,i)&vbCRLF
%>
<tr bgcolor="#FFFFFF" >
    <td><%=ArrRows(0,i)%></td>
    <td><%=ArrRows(1,i)%></td>
    <td><%=ArrRows(2,i)%></td>
    <td align="right"><%=ArrRows(3,i)%></td>
    <td><%=ArrRows(4,i)%></td>
    <td><%=ArrRows(5,i)%></td>
    <td align="center"><%=ArrRows(6,i)%></td>
    <td><%=ArrRows(7,i)%></td>
    <td><%=ArrRows(8,i)%></td>
    <td><%=ArrRows(9,i)%></td>
    <td align="center"><img src="http://scm.10x10.co.kr/images/icon_arrow_link.gif" style="cursor:pointer" onClick="reAct('<%=LEFT(ArrRows(5,i),10)%>')"></td>
</tr>
<% next %>
<% end if %>
</table>
</div>
<% if bufOutitemArr<>"" then %>
<textarea cols="30" rows="5"><%=bufOutitemArr%></textarea>
<% end if %>
<div id="actStr"></div>
<iframe name="xLink2" id="xLink2" frameborder="1" width="100%" height="300"></iframe>
<!-- #include virtual="/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
