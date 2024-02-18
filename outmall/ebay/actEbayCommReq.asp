<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim arrItemid : arrItemid = request("cksel")
Dim vGubun : vGubun = request("vGubun")
Dim sIdx : sIdx = request("sIdx")
Dim eIdx : eIdx = request("eIdx")
Dim i, strParam, iErrStr, ret1
Dim sqlStr, strSql, AssignedRow, SubNodes
Dim chgSellYn, actCnt, retErrStr
Dim buf, buf2, CNT10, CNT20, CNT30, iitemid
Dim ArrRows
Dim retFlag
Dim iMessage
dim iItemName, pregitemname
Dim depth
depth	  = request("depth")
retFlag   = request("retFlag")
chgSellYn = request("chgSellYn")
arrItemid = Trim(arrItemid)

If cmdparam = "GETSITECATE" OR cmdparam = "GETCATE" OR cmdparam = "GETMATCHCATE"  Then
	If cmdparam = "GETSITECATE" Then
		If (depth <> "o") AND (depth = "2") OR (depth = "3") OR (depth = "4") OR (depth = "5") Then
			ArrRows = fnGetSiteCateCodes(vGubun, depth)
			If isArray(ArrRows) Then
				For i = 0 To Ubound(ArrRows, 2)
					arrItemid = arrItemid & arrRows(0, i) & ", "
				Next
				arrItemid = trim(arrItemid)
				If Right(arrItemid, 1) = "," Then
					arrItemid = Trim(Left(arrItemid, Len(arrItemid) - 1))
				End If
			Else
				rw "해당 Depth 없음"
				response.end
			End If
		End If
	ElseIf cmdparam = "GETMATCHCATE" Then
		ArrRows = fnGetESMCateCodes()
		If isArray(ArrRows) Then
			For i = 0 To Ubound(ArrRows, 2)
				arrItemid = arrItemid & arrRows(0, i) & ", "
			Next
			arrItemid = trim(arrItemid)
			If Right(arrItemid, 1) = "," Then
				arrItemid = Trim(Left(arrItemid, Len(arrItemid) - 1))
			End If
		Else
			rw "해당 Depth 없음"
			response.end
		End If
	End If
Else
	rw "이 req 파일에서 사용할 수 없는 액션입니다."
	response.end
End If

Function fnGetSiteCateCodes(vGubun, vDepth)
	Dim strSql
	strSql = ""
	strSql = "EXEC [db_temp].[dbo].[usp_TEN_OutMall_Ebay_Cate_Get] '"&vGubun&"', '"&vDepth-1&"' "
	rsget.CursorLocation = adUseClient
	rsget.CursorType = adOpenStatic
	rsget.LockType = adLockOptimistic
	rsget.Open strSql, dbget
	If Not(rsget.EOF or rsget.BOF) then
		fnGetSiteCateCodes = rsget.getRows
	End If
	rsget.close
End Function

Function fnGetESMCateCodes()
	Dim strSql
	strSql = ""
	strSql = "EXEC [db_temp].[dbo].[usp_TEN_OutMall_Ebay_EsmCate_Get] "
	rsget.CursorLocation = adUseClient
	rsget.CursorType = adOpenStatic
	rsget.LockType = adLockOptimistic
	rsget.Open strSql, dbget
	If Not(rsget.EOF or rsget.BOF) then
		fnGetESMCateCodes = rsget.getRows
	End If
	rsget.close
End Function
%>
<script type="text/javascript">
	var items = "<%=arrItemid%>";
	var itemArr = items.split(", ");
	var rotation;
	var rno = 0;

	function loadRotation() {
		if(itemArr[rno] == undefined){
			alert('완료하였습니다');
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
        document.frmSvArr.itemid.value = ino;
        document.frmSvArr.chgSellYn.value = "<%=chgSellYn%>";
        document.frmSvArr.depth.value = "<%=depth%>";
		document.frmSvArr.vGubun.value = "<%=vGubun%>";
		document.frmSvArr.sIdx.value = "<%=sIdx%>";
		document.frmSvArr.eIdx.value = "<%=eIdx%>";
        document.frmSvArr.action = '/outmall/test/ebayCommActProc.asp';
        document.frmSvArr.submit();
	}
	window.onload = new Function('setTimeout("loadRotation()", 10)');
</script>
<form name="frmSvArr">
	<input type="hidden" name="act">
	<input type="hidden" name="itemid">
	<input type="hidden" name="chgSellYn">
	<input type="hidden" name="depth">
	<input type="hidden" name="vGubun">
	<input type="hidden" name="sIdx">
	<input type="hidden" name="eIdx">
</form>

<div id="actStr"></div>
<iframe name="xLink2" id="xLink2" frameborder="0" width="100%" height="100%"></iframe>
<!-- #include virtual="/lib/db/dbclose.asp" -->
