<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim arrIdx : arrIdx = Trim(request("cksel"))
Dim oGSShop, i, strParam, iErrStr, ret1
Dim sqlStr, strSql, AssignedRow, SubNodes
Dim chgSellYn, actCnt, retErrStr
Dim buf, buf2, CNT10, CNT20, CNT30, iitemid
Dim ArrRows
Dim retFlag
Dim iMessage
dim iItemName, pregitemname

retFlag   = request("retFlag")
chgSellYn = request("chgSellYn")

If (cmdparam = "EditStatCd") Then				''승인대기 -> 승인완료 프로세스
	Dim chgStatItemCode
	chgStatItemCode = request("chgStatItemCode")
	strSql = ""
	strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem SET "
	strSql = strSql & " GSShopStatCd = '7' "
	strSql = strSql & " WHERE itemid = '"& chgStatItemCode &"' "
	dbget.Execute(strSql)
	rw chgStatItemCode & " : 승인완료로 수정"
	response.end
End If

%>
<script type="text/javascript">
	var items = "<%=arrIdx%>";
	var itemArr = items.split(", ");
	var rotation;
	var rno = 0;

	function loadRotation() {
		if(itemArr[rno] == undefined){
			alert('완료하였습니다')
			return;
		}
		rotation = arrSubmit(itemArr[rno]);
		rno++;
		if(rno > itemArr.length-1){
			clearTimeout(rotation);
		}
	}

	function arrSubmit(ino){
		document.frmSvArr.target = "xLink2";
        document.frmSvArr.act.value = "<%=cmdparam%>";
        document.frmSvArr.idx.value = ino;
        document.frmSvArr.chgSellYn.value = "<%=chgSellYn%>";
        document.frmSvArr.action = '/outmall/gsshopAddOpt/gsshopActProc.asp';
        document.frmSvArr.submit();
	}
	window.onload = new Function('setTimeout("loadRotation()", 200)');
</script>
<form name="frmSvArr">
	<input type="hidden" name="act">
	<input type="hidden" name="idx">
	<input type="hidden" name="chgSellYn">
</form>

<div id="actStr"></div>
<iframe name="xLink2" id="xLink2" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
