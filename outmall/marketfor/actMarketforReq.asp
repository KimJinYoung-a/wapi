<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim arrItemid : arrItemid = request("cksel")
Dim chkXML : chkXML = request("chkXML")
Dim auto : auto = request("auto")
Dim oGSShop, i, strParam, iErrStr, ret1
Dim sqlStr, strSql, AssignedRow, SubNodes, getRegdate
Dim chgSellYn, actCnt, retErrStr
Dim buf, buf2, CNT10, CNT20, CNT30, iitemid
Dim ArrRows
Dim retFlag
Dim iMessage
dim iItemName, pregitemname

retFlag   = request("retFlag")
chgSellYn = request("chgSellYn")
arrItemid = Trim(arrItemid)
getRegdate = request("getRegdate")

If cmdparam = "STAT" Then
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_etcmall.dbo.tbl_ezwel_regitem "
	sqlStr = sqlStr & " SET ezwelStatcd = 7 "
	sqlStr = sqlStr & " WHERE ezwelStatcd = 3 "
	If getRegdate <> "" Then
		sqlStr = sqlStr & " and Ezwelregdate between '"&getRegdate&" 00:00:00' and '"&getRegdate&" 23:59:59' "
	End If
	dbget.Execute sqlStr, AssignedRow
	rw AssignedRow&"건 승인처리"
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
			<% if (auto <> "Y") then %>
			window.parent.postMessage({
				action: "systemAlert"
				, message: "완료하였습니다"
			}, "*");
			<% end if %>
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
		document.frmSvArr.chkXML.value = "<%=chkXML%>";
        document.frmSvArr.action = '/outmall/marketfor/marketforActProc.asp';
        document.frmSvArr.submit();
	}
	window.onload = new Function('setTimeout("loadRotation()", 200)');
</script>
<form name="frmSvArr">
	<input type="hidden" name="act">
	<input type="hidden" name="itemid">
	<input type="hidden" name="chgSellYn">
	<input type="hidden" name="chkXML">
</form>

<div id="actStr"></div>
<iframe name="xLink2" id="xLink2" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
