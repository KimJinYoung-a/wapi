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

retFlag		= request("retFlag")
chgSellYn	= request("chgSellYn")
arrIdx 		= Trim(arrIdx)

If cmdparam = "RegSelectWait" Then
	If Right(arrIdx,1) = "," Then arrIdx = Left(arrIdx, Len(arrIdx) - 1)
	sqlStr = ""
	sqlStr = sqlStr & " INSERT into db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] "
	sqlStr = sqlStr & " (midx, regdate, reguserid, LtiMallStatCD)"
	sqlStr = sqlStr & " SELECT idx, getdate(), '"&session("SSBctID")&"', '0' "
	sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M "
	sqlStr = sqlStr & " LEFT JOIN db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] R on M.idx = R.midx "
	sqlStr = sqlStr & " WHERE M.idx in ("&arrIdx&") "
	sqlStr = sqlStr & " and M.mallid = 'lotteimall' "
	sqlStr = sqlStr & " and R.midx is NULL"
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 예정등록됨"
	response.end
ElseIf (cmdparam = "DelSelectWait") Then
	If Right(arrIdx,1) = "," Then arrIdx = Left(arrIdx, Len(arrIdx) - 1)
	sqlStr = ""
	sqlStr = sqlStr & " DELETE FROM db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] "
	sqlStr = sqlStr & " WHERE LtiMallStatCD in ('0')"
	sqlStr = sqlStr & " and midx in ("&arrIdx&")"
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 예정삭제됨"
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
        document.frmSvArr.action = '/outmall/ltimallAddOpt/ltimallActProc.asp';
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
