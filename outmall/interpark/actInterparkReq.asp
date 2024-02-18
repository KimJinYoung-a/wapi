<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/outmall/interpark/interparkItemcls.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim arrItemid : arrItemid = request("cksel")
Dim auto : auto = request("auto")
Dim i, strParam, iErrStr, ret1
Dim sqlStr, strSql, AssignedRow, SubNodes
Dim chgSellYn, actCnt, retErrStr
Dim buf, buf2, CNT10, CNT20, CNT30, iitemid
Dim ArrRows
Dim retFlag
Dim iMessage
dim iItemName, pregitemname
Dim ccd
Dim param1, param2
ccd		  = request("CommCD")
retFlag   = request("retFlag")
chgSellYn = request("chgSellYn")
arrItemid = Trim(arrItemid)
param1 = request("param1")
param2 = request("param2")

If cmdparam = "RegSelectWait" Then
	If Right(arrItemid,1) = "," Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
	sqlStr = ""
	sqlStr = sqlStr & " INSERT INTO [db_item].[dbo].tbl_interpark_reg_item " & VbCrlf
	sqlStr = sqlStr & " (itemid,reguserid) " & VbCrlf
	sqlStr = sqlStr & " SELECT top 1000 i.itemid,'" & session("ssBctID") & "'" & VbCrlf
	sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item i" & VbCrlf
	sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_interpark_reg_item t on i.itemid = t.itemid" & VbCrlf
	sqlStr = sqlStr & " WHERE (" & VbCrlf
	sqlStr = sqlStr & " 	(i.cate_large='110' and  i.cate_mid in ('010','020','030','040','050','070','090'))" & VbCrlf
	sqlStr = sqlStr & " 	or" & VbCrlf
	sqlStr = sqlStr & " 	(i.cate_large in ('010','020','025','030','035','040','045','050','055','060','070','075','080','090','100'))" & VbCrlf
	sqlStr = sqlStr & " )"
	sqlStr = sqlStr & " and Not (i.cate_large='110' and i.cate_mid='030' and i.cate_small='040')" & VbCrlf  ''음반
	sqlStr = sqlStr & " and t.itemid is null" & VbCrlf
	sqlStr = sqlStr & " and i.itemid in (" & arrItemid & ")" & VbCrlf
	sqlStr = sqlStr & " and sellcash<>0" & VbCrlf
'	sqlStr = sqlStr & " and ((sellcash-buycash)/sellcash)*100>=" & CMAXMARGIN & VbCrlf			'2016-04-20 김진영 마진미만이라도 예정등록되게 수정
	sqlStr = sqlStr & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
	sqlStr = sqlStr & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
	sqlStr = sqlStr & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
	sqlStr = sqlStr & " 				) THEN 'Y' "
	sqlStr = sqlStr & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
	''특정상품제외
	sqlStr = sqlStr & " and i.itemid <> 114039" & VbCrlf
	sqlStr = sqlStr & " and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')" & VbCrlf
	''등록시 오류.
	sqlStr = sqlStr & " and i.makerid<>'haba'" & VbCrlf
	sqlStr = sqlStr & " and ((i.deliverytype<6) or " & VbCrlf
	sqlStr = sqlStr & "     ((i.deliverytype=9) " & VbCrlf
	sqlStr = sqlStr & "     and " & VbCrlf
	sqlStr = sqlStr & "     i.sellcash>=10000 " & VbCrlf ''' 조건배송은 1만원 이상짜리만..
	sqlStr = sqlStr & " ))" & VbCrlf
	''제휴 사용안함인거 걸러냄. isExtusing = 'N'
	sqlStr = sqlStr & " and i.isExtusing = 'Y'"
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 예정등록됨"
	response.end
ElseIf (cmdparam = "DelSelectWait") Then
	If Right(arrItemid,1) = "," Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
	sqlStr = ""
	sqlStr = sqlStr & " DELETE R " & VbCrlf
    sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_interpark_reg_item R " & VbCrlf
    sqlStr = sqlStr & " WHERE R.itemid in (" & arrItemid & ")" & VbCrlf
    sqlStr = sqlStr & " and interparkregdate is NULL" & VbCrlf
    sqlStr = sqlStr & " and interparkPrdNo is NULL" & VbCrlf
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 예정삭제됨"
	response.end
ElseIf cmdparam = "DELETE" Then
	If Right(arrItemid,1) = "," Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
	sqlStr = ""
	sqlStr = sqlStr & " DELETE FROM [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] WHERE sellsite='interpark' and itemid in (" & arrItemid & ") "
	dbget.Execute(sqlStr)

	sqlStr = sqlStr &" INSERT INTO [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] " & VBCRLF
	sqlStr = sqlStr &" SELECT 'interpark', i.itemid, r.interparkPrdNo, isnull(r.interparkRegdate, r.interparklastupdate), getdate(), r.lastErrStr " & VBCRLF
	sqlStr = sqlStr &" FROM db_item.dbo.tbl_item as i " & VBCRLF
	sqlStr = sqlStr &" JOIN  [db_item].[dbo].tbl_interpark_reg_item as r on i.itemid = r.itemid " & VBCRLF
	sqlStr = sqlStr &" WHERE i.itemid in (" & arrItemid & ") "
	dbget.Execute(sqlStr)

	sqlStr = ""
	sqlStr = sqlStr & " delete from  [db_item].[dbo].tbl_interpark_reg_item where itemid in (" & arrItemid & ")" & VbCrlf
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 상품 삭제"

	sqlStr = ""
	sqlStr = sqlStr & " delete from db_item.dbo.tbl_OutMall_regedoption where itemid in (" & arrItemid & ") and mallid = 'interpark' "
	dbget.Execute sqlStr,AssignedRow
	rw AssignedRow&"건 옵션 삭제"

	sqlStr = ""
	sqlStr = sqlStr & " delete from db_etcmall.dbo.tbl_outmall_API_Que where itemid in (" & arrItemid & ") and mallid = 'interpark' "
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
			<% if (auto <> "Y") then %>
			//alert('완료하였습니다');
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
		document.frmSvArr.ccd.value = "<%=ccd%>";
		document.frmSvArr.param1.value = "<%=param1%>";
		document.frmSvArr.param2.value = "<%=param2%>";
        document.frmSvArr.action = '/outmall/interpark/interparkActProc.asp';
        document.frmSvArr.submit();
	}
	window.onload = new Function('setTimeout("loadRotation()", 200)');
</script>
<form name="frmSvArr">
	<input type="hidden" name="act">
	<input type="hidden" name="itemid">
	<input type="hidden" name="chgSellYn">
	<input type="hidden" name="ccd">
	<input type="hidden" name="param1">
	<input type="hidden" name="param2">
</form>

<div id="actStr"></div>
<iframe name="xLink2" id="xLink2" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/lib/db/dbclose.asp" -->
