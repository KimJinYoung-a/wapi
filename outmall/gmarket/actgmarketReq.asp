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
Dim gMakername, gBrandname
retFlag   = request("retFlag")
chgSellYn = request("chgSellYn")
arrItemid = Trim(arrItemid)
gMakername = Trim(request("gMakername"))
gBrandname = Trim(request("gBrandname"))

If cmdparam = "DELETE" Then
	If Right(arrItemid,1) = "," Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
	sqlStr = ""
	sqlStr = sqlStr &" INSERT INTO [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] " & VBCRLF
	sqlStr = sqlStr &" SELECT 'gmarket1010', i.itemid, r.gmarketGoodNo, r.gmarketRegdate, getdate(), r.lastErrStr " & VBCRLF
	sqlStr = sqlStr &" FROM db_item.dbo.tbl_item as i " & VBCRLF
	sqlStr = sqlStr &" JOIN db_etcmall.dbo.tbl_gmarket_regitem as r on i.itemid = r.itemid " & VBCRLF
	sqlStr = sqlStr &" WHERE i.itemid in (" & arrItemid & ") "
	dbget.Execute(sqlStr)

	sqlStr = ""
	sqlStr = sqlStr & " delete from db_etcmall.dbo.tbl_gmarket_regitem where itemid in (" & arrItemid & ")" & VbCrlf
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 상품 삭제"

	sqlStr = ""
	sqlStr = sqlStr & " delete from db_item.dbo.tbl_OutMall_regedoption where itemid in (" & arrItemid & ") and mallid = 'gmarket1010' "
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 옵션 삭제"

	sqlStr = ""
	sqlStr = sqlStr & " delete from db_etcmall.dbo.tbl_outmall_API_Que where itemid in (" & arrItemid & ") and mallid = 'gmarket1010' "
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 Que 삭제"
	response.end
End If
%>
<script type="text/javascript">
	var items = "<%=arrItemid%>";
	var itemArr = items.split(", ");
	var rotation;
	var rno = 0;

	function loadRotation() {
		if(itemArr[rno] == undefined){
			//alert('완료하였습니다')
			window.parent.postMessage({
				action: "systemAlert"
				, message: "완료하였습니다"
			}, "*");
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
        document.frmSvArr.gMakername.value = "<%=gMakername%>";
        document.frmSvArr.gBrandname.value = "<%=gBrandname%>";
        document.frmSvArr.action = '/outmall/gmarket/gmarketActProc.asp';
        document.frmSvArr.submit();
	}
	window.onload = new Function('setTimeout("loadRotation()", 200)');
</script>
<form name="frmSvArr">
	<input type="hidden" name="act">
	<input type="hidden" name="itemid">
	<input type="hidden" name="chgSellYn">
	<input type="hidden" name="gMakername">
	<input type="hidden" name="gBrandname">
</form>

<div id="actStr"></div>
<iframe name="xLink2" id="xLink2" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/lib/db/dbclose.asp" -->
