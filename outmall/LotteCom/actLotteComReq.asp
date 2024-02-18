<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim arrItemid : arrItemid = request("cksel")
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
arrItemid = Trim(arrItemid)

If cmdparam = "RegSelectWait" Then
	If Right(arrItemid,1) = "," Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
	sqlStr = ""
	sqlStr = sqlStr & " INSERT into db_item.dbo.tbl_lotte_regItem "
	sqlStr = sqlStr & " (itemid, regdate, reguserid, LotteStatCd)"
	sqlStr = sqlStr & " SELECT i.itemid, getdate(), '"&session("SSBctID")&"', '00' "
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i"
	sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_lotte_regItem R on i.itemid = R.itemid "
	sqlStr = sqlStr & " WHERE i.itemid in ("&arrItemid&") "
	sqlStr = sqlStr & " and R.itemid is NULL"
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 예정등록됨"

	sqlStr = ""
	sqlStr = sqlStr & " update R "
	sqlStr = sqlStr & " set optAddPrcCnt= T.optAddPrcCnt "
	sqlStr = sqlStr & " from db_item.dbo.tbl_lotte_regitem R "
	sqlStr = sqlStr & " Join ( "
	sqlStr = sqlStr & " 	select ii.itemid,count(*) as optAddPrcCnt "
	sqlStr = sqlStr & " 	from db_item.dbo.tbl_item ii 	 "
	sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item_option o 	 "
	sqlStr = sqlStr & "		on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'	 "
	sqlStr = sqlStr & " 	group by ii.itemid "
	sqlStr = sqlStr & " ) T on R.itemid =T.itemid "
	sqlStr = sqlStr & " WHERE R.itemid in ("&arrItemid&") "
	dbget.Execute sqlStr,AssignedRow
	response.end
ElseIf (cmdparam = "DelSelectWait") Then
	If Right(arrItemid,1) = "," Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
	sqlStr = ""
	sqlStr = sqlStr & " DELETE FROM db_item.dbo.tbl_lotte_regItem "
	sqlStr = sqlStr & " WHERE LotteStatCd in ('00')"
	sqlStr = sqlStr & " and itemid in ("&arrItemid&")"
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 예정삭제됨"
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
        document.frmSvArr.submit();
	}
	window.onload = new Function('setTimeout("loadRotation()", 200)');
</script>
<form name="frmSvArr">
	<input type="hidden" name="act">
	<input type="hidden" name="itemid">
	<input type="hidden" name="chgSellYn">
</form>

<div id="actStr"></div>
<iframe name="xLink2" id="xLink2" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
