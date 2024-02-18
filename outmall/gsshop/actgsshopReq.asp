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

If cmdparam = "sugiRegedoption" Then
	Dim ckLimit, arrGSShopInfo
	ckLimit = request("ckLimit")
	If ckLimit = "" Then
		Response.Write "<script language=javascript>alert('한정 여부 선택 후 진행하세요');</script>"
		dbget.Close: Response.End
	End If

	strSql = ""
	strSql = strSql & " SELECT itemid, gsshopgoodno FROM db_item.dbo.tbl_gsshop_regitem WHERE itemid in ("&arrItemid&") "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		arrGSShopInfo = rsget.getrows()
	rsget.Close

	rw "--실제 실행되는 쿼리가 아닙니다~!"
	For i = 0 To Ubound(arrGSShopInfo,2)
		If ckLimit = "N" Then
			rw "insert into db_item.dbo.tbl_outmall_regedoption (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate) values "
			rw "('"&arrGSShopInfo(0,i)&"', '0000', 'gsshop', '"&arrGSShopInfo(1,i)&"001', '단일상품', 'Y', 'N', '999', '0', getdate())"&"<br>"
		ElseIf ckLimit = "Y" Then
			rw "insert into db_item.dbo.tbl_outmall_regedoption (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate) values "
			rw "('"&arrGSShopInfo(0,i)&"', '0000', 'gsshop', '"&arrGSShopInfo(1,i)&"001', '단일상품', 'Y', 'Y', '220', '0', getdate())"&"<br>"'
		End If
	Next
	response.end
ElseIf (cmdparam = "EditStatCd") Then				''승인대기 -> 승인완료 프로세스
	Dim chgStatItemCode
	chgStatItemCode = request("chgStatItemCode")
	strSql = ""
	strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem SET "
	strSql = strSql & " GSShopStatCd = '7' "
	strSql = strSql & " WHERE itemid = '"& chgStatItemCode &"' "
	dbget.Execute(strSql)
	rw chgStatItemCode & " : 승인완료로 수정"
	response.end
ElseIf cmdparam = "DELETE" Then
	If Right(arrItemid,1) = "," Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
	sqlStr = ""
	sqlStr = sqlStr &" INSERT INTO [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] " & VBCRLF
	sqlStr = sqlStr &" SELECT 'gsshop', i.itemid, r.GSShopGoodNo, r.GSShopRegdate, getdate(), r.lastErrStr " & VBCRLF
	sqlStr = sqlStr &" FROM db_item.dbo.tbl_item as i " & VBCRLF
	sqlStr = sqlStr &" JOIN db_item.dbo.tbl_gsshop_regitem as r on i.itemid = r.itemid " & VBCRLF
	sqlStr = sqlStr &" WHERE i.itemid in (" & arrItemid & ") "
	dbget.Execute(sqlStr)

	sqlStr = ""
	sqlStr = sqlStr & " delete from db_item.dbo.tbl_gsshop_regItem where itemid in (" & arrItemid & ")" & VbCrlf
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 상품 삭제"

	sqlStr = ""
	sqlStr = sqlStr & " delete from db_item.dbo.tbl_OutMall_regedoption where itemid in (" & arrItemid & ") and mallid = 'gsshop' "
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 옵션 삭제"

	sqlStr = ""
	sqlStr = sqlStr & " delete from db_etcmall.dbo.tbl_outmall_API_Que where itemid in (" & arrItemid & ") and mallid = 'gsshop' "
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
        document.frmSvArr.action = '/outmall/gsshop/gsshopActProc.asp';
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
