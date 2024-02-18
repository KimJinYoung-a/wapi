<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/outmall/cjmall/cjmallitemcls.asp"-->
<%
Dim ocjmall, i
Dim cdl, cdm, cds, dispNo '', dispNm, dispFull
cdl		= request("cdl")
cdm		= request("cdm")
cds		= request("cds")
dispNo	= request("dspNo")

If cdl = "" Then
	Call Alert_Close("카테고리 코드가 없습니다.")
	dbget.Close: Response.End
End IF

'// 카테고리 내용 접수
Set ocjmall = new CCjmall
	ocjmall.FPageSize = 20
	ocjmall.FCurrPage = 1
	ocjmall.FRectCDL = cdl
	ocjmall.FRectCDM = cdm
	ocjmall.FRectCDS = cds
	ocjmall.getTencjmallCateList

If ocjmall.FResultCount <= 0 Then
	Call Alert_Close("해당 카테고리 정보가 없습니다.")
	dbget.Close: Response.End
End If
%>
<script language="javascript">
<!--
	// 매칭 저장하기
	function fnSaveForm() {
		var frm = document.frmAct;

		if(frm.dspNo.value=="") {
			alert("매칭할 cjmall 카테고리를 선택해주세요.");
			return;
		}

		if(confirm("선택하신 카테고리로 매칭하시겠습니까?")) {
			frm.mode.value="saveCate";
			frm.action="proccjmall.asp";
			frm.submit();
		}
	}

    function fnDelForm(iDspNo) {
		var frm = document.frmAct;
		if (iDspNo=="") {
		    alert("삭제할 cjmall 카테고리를 선택해주세요.");
			return;
		}

		if(confirm("현재 매칭된 카테고리를 연결해제 하시겠습니까?\n\n※ 상품 또는 카테고리가 삭제되는 것은 아니며, 연결된 정보만 삭제됩니다.")) {
			frm.mode.value="delCate";
			frm.dspNo.value=iDspNo;
			frm.action="proccjmall.asp";
			frm.submit();
		}
	}

	// 창닫기
	function fnCancel() {
		if(confirm("작업을 취소하고 창을 닫으시겠습니까?")) {
			self.close();
		}
	}

	// cjmall 카테고리 검색
	function fnSearchCJCate(disptpcd) {
		var pFCL = window.open("","popcjmallCate","width=900,height=700,scrollbars=yes,resizable=yes");
		pFCL.focus();
		srcFrm.target="popcjmallCate";
		srcFrm.action="popFindcjmallCate.asp";
		srcFrm.submit();
	}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/lib/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/lib/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/lib/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/lib/images/tbl_blue_round_04.gif"></td>
	<td><img src="/lib/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>cjmall 카테고리 매칭</strong></font></td>
	<td background="/lib/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/lib/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/lib/images/tbl_blue_round_08.gif"></td>
	<td><img src="/lib/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<p>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/lib/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/lib/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/lib/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/lib/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 상단바 끝-->
<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/lib/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/lib/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 텐바이텐 카테고리 정보</td>
	<td width="10" align="left" background="/lib/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">대분류</td>
	<td bgcolor="#FFFFFF">[<%=cdl%>] <%=ocjmall.FItemList(0).FtenCDLName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">중분류</td>
	<td bgcolor="#FFFFFF">[<%=cdm%>] <%=ocjmall.FItemList(0).FtenCDMName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">소분류</td>
	<td bgcolor="#FFFFFF">[<%=cds%>] <%=ocjmall.FItemList(0).FtenCDSName%></td>
</tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/lib/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/lib/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> cjmall 전시 카테고리 매칭 정보 <!--(변경하려면 삭제후 저장) 1:1 카테고리로.--></td>
	<td width="10" align="left" background="/lib/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<form name="srcFrm" method="GET" onsubmit="return false" style="margin:0px;">
<input type="hidden" name="srcDiv" value="CNM">
<input type="hidden" name="disptpcd" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="2" >검색</td>
	<td bgcolor="#FFFFFF">
		카테고리명 <input type="text" name="srcKwd" class="text">
		<input type="button" value="검색" class="button" onClick="fnSearchCJCate()">
	</td>
</tr>
<tr id="BrRow" style="display:">
	<td bgcolor="#F2F2F2">추가 : <b><span id="selBr"></span></b></td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="<%= ocjmall.FResultCount + 1 %>" >등록된<br>카테고리</td>
	<td bgcolor="#FFFFFF" height="1"></td>
</tr>
<% For i = 0 to ocjmall.FResultCount - 1 %>
<% If Not IsNULL(ocjmall.FItemList(i).FDispNo) Then %>
<tr>
    <td bgcolor="#F2F2F2"><b><span id="selBr"><%=ocjmall.FItemList(i).getDispGubunNm%> [<%=ocjmall.FItemList(i).FDispNo%>] <%=ocjmall.FItemList(i).FDispNm%></span></b>
    &nbsp;&nbsp;&nbsp;&nbsp;<img src="/lib/images/icon_delete.gif" width="45" height="20" border="0" onclick="fnDelForm('<%=ocjmall.FItemList(i).FDispNo%>')" style="cursor:pointer" align="absmiddle">
    </td>
</tr>
<% End If %>
<% Next %>
</table>
</form>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="28">
    <td width="10" align="right" background="/lib/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left"></td>
    <td valign="bottom" align="right">
		<img src="/lib/images/icon_cancel.gif" width="45" height="20" border="0" onclick="fnCancel()" style="cursor:pointer" align="absmiddle"> &nbsp;&nbsp;&nbsp;
		<img src="/lib/images/icon_save.gif" width="45" height="20" border="0" onclick="fnSaveForm()" style="cursor:pointer" align="absmiddle">
    </td>
    <td width="10" align="left" background="/lib/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/lib/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td colspan="2" background="/lib/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/lib/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 하단바 끝-->
<form name="frmAct" method="POST" target="xLink" style="margin:0px;">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="cdm" value="<%=cdm%>">
<input type="hidden" name="cds" value="<%=cds%>">
<input type="hidden" name="dspNo" value="">
<input type="hidden" name="mode" value="saveCate">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="110" height="110"></iframe>
</p>
<% Set ocjmall = Nothing %>
<!-- #include virtual="/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
