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
Dim infodiv, infodivnm, mode
mode	= request("mode")
cdl		= request("cdl")
cdm		= request("cdm")
cds		= request("cds")
infodiv	= request("infodiv")
dispNo	= request("dspNo")

Select Case infodiv
	Case "01"	infodivnm = "의류"
	Case "02"	infodivnm = "구두/신발"
	Case "03"	infodivnm = "가방"
	Case "04"	infodivnm = "패션잡화(모자/벨트/액세서리)"
	Case "05"	infodivnm = "침구류/커튼"
	Case "06"	infodivnm = "가구(침대/소파/싱크대/DIY제품)"
	Case "07"	infodivnm = "영상가전(TV류)"
	Case "08"	infodivnm = "가정용 전기제품(냉장고/세탁기/식기세척기/전자레인지)"
	Case "09"	infodivnm = "계절가전(에어컨/온풍기)"
	Case "10"	infodivnm = "사무용기기(컴퓨터/노트북/프린터)"
	Case "11"	infodivnm = "광학기기(디지털카메라/캠코더)"
	Case "12"	infodivnm = "소형전자(MP3/전자사전 등)"
	Case "13"	infodivnm = "휴대폰"
	Case "14"	infodivnm = "내비게이션"
	Case "15"	infodivnm = "자동차용품(자동차부품/기타 자동차용품)"
	Case "16"	infodivnm = "의료기기"
	Case "17"	infodivnm = "주방용품"
	Case "18"	infodivnm = "화장품"
	Case "19"	infodivnm = "귀금속/보석/시계류"
	Case "20"	infodivnm = "식품(농수산물)"
	Case "21"	infodivnm = "가공식품"
	Case "22"	infodivnm = "건강기능식품"
	Case "23"	infodivnm = "영유아용품"
	Case "24"	infodivnm = "악기"
	Case "25"	infodivnm = "스포츠용품"
	Case "26"	infodivnm = "서적"
	Case "27"	infodivnm = "호텔/펜션 예약"
	Case "28"	infodivnm = "여행패키지"
	Case "29"	infodivnm = "항공권"
	Case "30"	infodivnm = "자동차 대여 서비스(렌터카)"
	Case "31"	infodivnm = "물품대여 서비스(정수기, 비데, 공기청정기 등)"
	Case "32"	infodivnm = "물품대여 서비스(서적, 유아용품, 행사용품 등)"
	Case "33"	infodivnm = "디지털 콘텐츠(음원, 게임, 인터넷강의 등)"
	Case "34"	infodivnm = "상품권/쿠폰"
	Case "35"	infodivnm = "기타"
End Select


If cdl = "" Then
	Call Alert_Close("카테고리 코드가 없습니다.")
	dbget.Close: Response.End
End IF

'// 카테고리 내용 접수
Set ocjmall = new CCjmall
	ocjmall.FRectCDL = cdl
	ocjmall.FRectCDM = cdm
	ocjmall.FRectCDS = cds
	ocjmall.Finfodiv = infodiv
	ocjmall.getTencjmallOneprdDiv

'If ocjmall.FResultCount <= 0 Then
'	Call Alert_Close("해당 카테고리 정보가 없습니다.")
'	dbget.Close: Response.End
'End If
%>
<script language="javascript">
<!--
	// 매칭 저장하기
	function fnSaveForm() {
		var frm = document.frmAct;

		if(frm.dspNo.value=="") {
			alert("매칭할 cjmall 상품분류를 선택해주세요.");
			return;
		}

		if(confirm("선택하신 상품분류로 매칭하시겠습니까?")) {
			frm.mode.value="saveCate";
			frm.action="proccjmall2.asp";
			frm.submit();
		}
	}

    function fnDelForm(iDspNo) {
		var frm = document.frmAct;
		if (iDspNo=="") {
		    alert("삭제할 cjmall 상품분류를 선택해주세요.");
			return;
		}

		if(confirm("현재 매칭된 상품분류를 연결해제 하시겠습니까?\n\n※ 상품분류가 삭제되는 것은 아니며, 연결된 정보만 삭제됩니다.")) {
			frm.mode.value="delPrddiv";
			frm.dspNo.value=iDspNo;
			frm.action="proccjmall2.asp";
			frm.submit();
		}
	}

	// 창닫기
	function fnCancel() {
		if(confirm("작업을 취소하고 창을 닫으시겠습니까?")) {
			self.close();
		}
	}

	// cjmall 상품분류 검색
	function fnSearchCJPrddiv(disptpcd, cddnm) {
		var pFCL = window.open("popFindcjmallPrddiv.asp?infodiv="+disptpcd+"&cdd_NAME="+cddnm,"popcjmallPrddiv","width=900,height=700,scrollbars=yes,resizable=yes");
		pFCL.focus();
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
	<font color="red"><strong>cjmall 상품분류 매칭</strong></font></td>
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
	<td width="80" align="center" bgcolor="#DDDDFF">품목정보</td>
	<td bgcolor="#FFFFFF">[<%=infodiv%>] <%=infodivnm%></td>
</tr>
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
	<td align="left"><img src="/lib/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> cjmall 상품분류 매칭 정보</td>
	<td width="10" align="left" background="/lib/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<form name="srcFrm" method="GET" onsubmit="return false" style="margin:0px;">
<input type="hidden" name="srcDiv" value="CNM">
<input type="hidden" name="disptpcd" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<% If mode <> "U" Then %>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="2" >검색</td>
	<td bgcolor="#FFFFFF">
		세분류명 <input type="text" name="srcKwd" class="text">
		<input type="button" value="검색" class="button" onClick="fnSearchCJPrddiv('<%=infodiv%>',document.srcFrm.srcKwd.value)">
	</td>
</tr>
<tr id="BrRow" style="display:">
	<td bgcolor="#F2F2F2">추가 : <b><span id="selBr"></span></b></td>
</tr>
<% End If %>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="<%= ocjmall.FResultCount + 1 %>" >등록된<br>상품분류</td>
	<td bgcolor="#FFFFFF" height="1"></td>
</tr>
<% If Not IsNULL(ocjmall.FItemList(0).FCddKey) Then %>
<tr>
    <td bgcolor="#F2F2F2"><b><span id="selBr">[<%=ocjmall.FItemList(0).FCddKey%>] <%=ocjmall.FItemList(0).Fcdd_Name%></span></b>
    &nbsp;&nbsp;&nbsp;&nbsp;<img src="/lib/images/icon_delete.gif" width="45" height="20" border="0" onclick="fnDelForm('<%=ocjmall.FItemList(0).FCddKey%>')" style="cursor:pointer" align="absmiddle">
    </td>
</tr>
<% End If %>
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
<input type="hidden" name="infodiv" value="<%=infodiv%>">
<input type="hidden" name="CdmKey" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="110" height="110"></iframe>
</p>
<% Set ocjmall = Nothing %>
<!-- #include virtual="/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
