<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/outmall/cjmall/cjmallitemcls.asp"-->
<%
Dim ocjmall, i, page, isMapping, srcDiv, srcKwd
Dim disptpcd
Dim cateAllNm

page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")
disptpcd    = request("disptpcd")

If page = ""	Then page = 1
If srcDiv = ""	Then srcDiv = "CCD"

'// 목록 접수
Set ocjmall = new CCjmall
	ocjmall.FPageSize 		= 20
	ocjmall.FCurrPage		= page
	ocjmall.FRectIsMapping	= isMapping
	ocjmall.FRectSDiv		= srcDiv
	ocjmall.FRectKeyword	= srcKwd
	ocjmall.FRectCDL		= request("cdl")
	ocjmall.FRectCDM		= request("cdm")
	ocjmall.FRectCDS		= request("cds")
	ocjmall.FRectdisptpcd	= disptpcd
	ocjmall.getTencjmallCateList
%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg) {
		frm.page.value = pg;
		frm.submit();
	}

	// 검색
	function serchItem() {
		frm.page.value = 1;
		frm.submit();
	}

	// cjmall 카테고리 매칭 팝업
	function popCjCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("popcjmallCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&dspNo="+dno,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
		pCM.focus();
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
	<font color="red"><strong>cjmall 카테고리 관리</strong></font></td>
	<td background="/lib/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/lib/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/lib/images/tbl_blue_round_08.gif"></td>
	<td><img src="/lib/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 액션 -->
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;">
		텐바이텐 <!-- #include virtual="/lib/module/categoryselectbox.asp"--><br>
		전시구분 : 
	    <select name="disptpcd" class="select">
	        <option value="">선택</option>
			<option value="B" <%=chkIIF(disptpcd="B","selected","")%>>전문</option>
			<option value="D" <%=chkIIF(disptpcd="D","selected","")%>>일반</option>
		</select> /
		매칭여부 :
		<select name="ismap" class="select">
			<option value="">전체</option>
			<option value="Y" <%=chkIIF(isMapping="Y","selected","")%>>매칭완료</option>
			<option value="N" <%=chkIIF(isMapping="N","selected","")%>>미매칭</option>
		</select> /
		검색구분 :
		<select name="srcDiv" class="select">
			<option value="CCD" <%=chkIIF(srcDiv="CCD","selected","")%>>cjmall 코드</option>
			<option value="CNM" <%=chkIIF(srcDiv="CNM","selected","")%>>카테고리명</option>
		</select> /
		검색어 :
		<input type="text" name="srcKwd" size="15" value="<%=srcKwd%>" class="text">
	</td>
	<td width="55" align="right" style="padding-top:5px;">
		<input id="btnRefresh" type="button" class="button" value="검색" onclick="serchItem()" style="width:50px;height:40px;">
	</td>
</tr>
</table>
</form>
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
	<td align="left"><img src="/lib/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=ocjmall.FtotalCount%></strong></td>
	<td width="10" align="left" background="/lib/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="4">텐바이텐 카테고리</td>
	<td colspan="4">cjmall 카테고리</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>코드</td>
	<td>대분류</td>
	<td>중분류</td>
	<td>소분류</td>
	<td>구분</td>
	<td>코드</td>
	<td>카테고리명</td>
	<td>cjmall 전시(한글)</td>
</tr>
<% If ocjmall.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to ocjmall.FresultCount - 1
			cateAllNm 	= ocjmall.FItemList(i).FDispLrgNm & ">" & ocjmall.FItemList(i).FDispMidNm & ">" & ocjmall.FItemList(i).FDispSmlNm & ">" & ocjmall.FItemList(i).FDispNm
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(ocjmall.FItemList(i).FCateIsUsing="Y","#FFFFFF","#CCCCCC") %>">
	<td><%= ocjmall.FItemList(i).FtenCateLarge & ocjmall.FItemList(i).FtenCateMid & ocjmall.FItemList(i).FtenCateSmall %></td>
	<td><%= ocjmall.FItemList(i).FtenCDLName %></td>
	<td><%= ocjmall.FItemList(i).FtenCDMName %></td>
	<td><%= ocjmall.FItemList(i).FtenCDSName %></td>
	<% If ocjmall.FItemList(i).FDispNo="" OR isNull(ocjmall.FItemList(i).FDispNo) Then %>
	<td colspan="4"><input type="button" class="button" value="cjmall 카테 매칭" onClick="popCjCateMap('<%= ocjmall.FItemList(i).FtenCateLarge %>','<%= ocjmall.FItemList(i).FtenCateMid %>','<%= ocjmall.FItemList(i).FtenCateSmall %>','')"></td>
	<% Else %>
	<td title="<%=cateAllNm%>" onClick="popCjCateMap('<%= ocjmall.FItemList(i).FtenCateLarge %>','<%= ocjmall.FItemList(i).FtenCateMid %>','<%= ocjmall.FItemList(i).FtenCateSmall %>','<%=ocjmall.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= ocjmall.FItemList(i).getDispGubunNm %></td>
	<td title="<%=cateAllNm%>" onClick="popCjCateMap('<%= ocjmall.FItemList(i).FtenCateLarge %>','<%= ocjmall.FItemList(i).FtenCateMid %>','<%= ocjmall.FItemList(i).FtenCateSmall %>','<%=ocjmall.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= ocjmall.FItemList(i).FDispNo %></td>
	<td title="<%=cateAllNm%>" onClick="popCjCateMap('<%= ocjmall.FItemList(i).FtenCateLarge %>','<%= ocjmall.FItemList(i).FtenCateMid %>','<%= ocjmall.FItemList(i).FtenCateSmall %>','<%=ocjmall.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= ocjmall.FItemList(i).FDispNm %></td>
	<td><%=cateAllNm%></td>
	<% End If %>
</tr>
<%
		Next
	End If
%>
</table>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/lib/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% If ocjmall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= ocjmall.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + ocjmall.StartScrollPage to ocjmall.FScrollCount + ocjmall.StartScrollPage - 1 %>
			<% If i > ocjmall.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If ocjmall.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% Else %>
			[next]
		<% End If %>
    </td>
    <td width="10" align="left" background="/lib/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/lib/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/lib/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/lib/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<% Set ocjmall = Nothing %>
<!-- #include virtual="/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->