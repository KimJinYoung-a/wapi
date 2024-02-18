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
	sqlStr = sqlStr & " INSERT into db_item.dbo.tbl_LTiMall_regItem "
	sqlStr = sqlStr & " (itemid, regdate, reguserid, LtiMallStatCD)"
	sqlStr = sqlStr & " SELECT i.itemid, getdate(), '"&session("SSBctID")&"', '0' "
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i"
	sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_LTiMall_regItem R on i.itemid = R.itemid "
	sqlStr = sqlStr & " WHERE i.itemid in ("&arrItemid&") "
	sqlStr = sqlStr & " and R.itemid is NULL"
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 예정등록됨"

	sqlStr = ""
	sqlStr = sqlStr & " update R "
	sqlStr = sqlStr & " set optAddPrcCnt= T.optAddPrcCnt "
	sqlStr = sqlStr & " from db_item.dbo.tbl_LTiMall_regItem R "
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
	sqlStr = sqlStr & " DELETE FROM db_item.dbo.tbl_LTiMall_regItem "
	sqlStr = sqlStr & " WHERE LtimallStatCD in ('0')"
	sqlStr = sqlStr & " and itemid in ("&arrItemid&")"
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 예정삭제됨"
	response.end
ElseIf cmdparam = "DELETE" Then
	If Right(arrItemid,1) = "," Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
	sqlStr = ""
	sqlStr = sqlStr & " DELETE FROM [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] WHERE sellsite='interpark' and itemid in (" & arrItemid & ") "
	dbget.Execute(sqlStr)

	sqlStr = sqlStr &" INSERT INTO [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] " & VBCRLF
	sqlStr = sqlStr &" SELECT 'lotteimall', i.itemid, r.LTiMallGoodNo, isnull(r.LTiMallRegdate, r.LTiMallLastUpdate), getdate(), r.lastErrStr " & VBCRLF
	sqlStr = sqlStr &" FROM db_item.dbo.tbl_item as i " & VBCRLF
	sqlStr = sqlStr &" JOIN db_item.dbo.tbl_LTiMall_regItem as r on i.itemid = r.itemid " & VBCRLF
	sqlStr = sqlStr &" WHERE i.itemid in (" & arrItemid & ") "
	dbget.Execute(sqlStr)

	sqlStr = ""
	sqlStr = sqlStr & " delete from db_item.dbo.tbl_LTiMall_regItem where itemid in (" & arrItemid & ")" & VbCrlf
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 상품 삭제"

	sqlStr = ""
	sqlStr = sqlStr & " delete from db_item.dbo.tbl_OutMall_regedoption where itemid in (" & arrItemid & ") and mallid = 'lotteimall' "
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 옵션 삭제"

	sqlStr = ""
	sqlStr = sqlStr & " delete from db_etcmall.dbo.tbl_outmall_API_Que where itemid in (" & arrItemid & ") and mallid = 'lotteimall' "
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
        document.frmSvArr.action = '/outmall/ltimall/ltimallActProc.asp';
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
