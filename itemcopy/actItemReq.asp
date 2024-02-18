<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim arrItemid : arrItemid = request("itemarr")
Dim arrBrandId : arrBrandId = request("brandarr")
Dim arrItemdiv : arrItemdiv = request("itemdivarr")
arrItemid = Trim(arrItemid)
arrBrandId = Trim(arrBrandId)
arrItemdiv = Trim(arrItemdiv)

arrItemid = Replace(arrItemid, "||", ",")
arrBrandId = Replace(arrBrandId, "||", ",")
arrItemdiv = Replace(arrItemdiv, "||", ",")

If Right(arrItemid,1) = "," Then arrItemid = Left(arrItemid, Len(arrItemid) - 1)
If Right(arrBrandId,1) = "," Then arrBrandId = Left(arrBrandId, Len(arrBrandId) - 1)
If Right(arrItemdiv,1) = "," Then arrItemdiv = Left(arrItemdiv, Len(arrItemdiv) - 1)
%>
<script type="text/javascript">
	var items = "<%=arrItemid%>";
	var itemArr = items.split(",");

	var makerids = "<%=arrBrandId%>";
	var makeridArr = makerids.split(",");

	var itemdivs = "<%=arrItemdiv%>";
	var itemdivArr = itemdivs.split(",");

	var rotation;
	var rno = 0;

	function loadRotation() {
		if(itemArr[rno] == undefined){
			alert('완료하였습니다');
			return;
		}
		rotation = arrSubmit(itemArr[rno], makeridArr[rno], itemdivArr[rno]);
		rno++;
		if(rno > itemArr.length-1){
			clearTimeout(rotation);
		}else{
			//setTimeout('loadRotation()', 200);
		}
	}

	function arrSubmit(ino, mid, idv){
		document.frmSvArr.target = "xLink2";
		document.frmSvArr.act.value = "<%=cmdparam%>";
		document.frmSvArr.itemid.value = ino;
		document.frmSvArr.makerid.value = mid;
		document.frmSvArr.itemdiv.value = idv;
		document.frmSvArr.action = '/itemcopy/itemActProc.asp';
		document.frmSvArr.submit();
	}
	window.onload = new Function('setTimeout("loadRotation()", 200)');
</script>
<form name="frmSvArr">
	<input type="hidden" name="act">
	<input type="hidden" name="itemid">
	<input type="hidden" name="makerid">
	<input type="hidden" name="itemdiv">
</form>

<div id="actStr"></div>
<iframe name="xLink2" id="xLink2" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/lib/db/dbclose.asp" -->
