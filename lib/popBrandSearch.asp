<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partnerusercls.asp"-->

<%
dim i,page
dim rectDesigner
dim usingonly, research, userdiv, rect, crect, mrect, mduserid
dim catecode, vIsJSDomain

dim frmName, compName, socname_kr

frmName= request("frmName")
compName= request("compName")

rectDesigner= request("rectDesigner")
usingonly   = request("usingonly")

research    = request("research")
userdiv     = request("userdiv")
rect        = requestCheckVar(request("rect"),60)
mduserid    = request("mduserid")
catecode    = request("catecode")
crect       = requestCheckVar(request("crect"),60)
mrect       = requestCheckVar(request("mrect"),60)
socname_kr  = requestCheckVar(request("socname_kr"),60)

page        = request("page")

vIsJSDomain = request("isjsdomain")


if ((research="") and (usingonly="")) then usingonly="all"
if ((research="") and (userdiv="")) then userdiv="02"

if page="" then page=1

dim opartner
set opartner = new CPartnerUser
opartner.FCurrpage = page
opartner.FPageSize = 100
opartner.FRectDesignerID = rectDesigner
opartner.FrectIsUsing = usingonly
opartner.FRectDesignerDiv = userdiv
opartner.FRectMdUserID = mduserid
opartner.FRectInitial = rect
opartner.FRectCompanyname = crect
opartner.FRectManagerName = mrect
opartner.FRectCatecode = catecode
opartner.FRectSOCName  = socname_kr
opartner.GetPartnerNUserCList
%>

<script language='javascript'>
<% If vIsJSDomain = "o" Then Response.Write "document.domain = '10x10.co.kr';" End If %>

function NextPage(page){
	frm.page.value = page;
	frm.submit();
}

function research(frm,order){
	frm.rectorder.value = order;
	frm.submit();
}

function selectThis(selval){
    opener.<%= frmName %>.<%= compName %>.value = selval;
    window.close();
}
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="rectorder" value="">
	
	<input type="hidden" name="frmName" value="<%= frmName %>">
	<input type="hidden" name="compName" value="<%= compName %>">
	<input type="hidden" name="isjsdomain" value="<%= vIsJSDomain %>">
	
	
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/lib/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/lib/images/tbl_blue_round_02.gif"></td>
	        <td background="/lib/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/lib/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/lib/images/tbl_blue_round_04.gif"></td>
	        <td>
	            <% if (Not C_IS_FRN_SHOP) THen %>
	        	<input type=radio name=usingonly  value="all" <% if usingonly="all" then response.write "checked" %> >전체
				<input type=radio name=usingonly  value="on" <% if usingonly="on" then response.write "checked" %> >사용함
				
				<input type=radio name=usingonly  value="off_new" <% if usingonly="off_new" then response.write "checked" %> >사용안함(신규)
				<input type=radio name=usingonly  value="off_old" <% if usingonly="off_old" then response.write "checked" %> >사용안함(한달이전)
				<br>
				담당자 : <% drawSelectBoxCoWorker "mduserid", mduserid %>
				&nbsp;
				카테고리 : <% SelectBoxBrandCategory "catecode", catecode %>
				&nbsp;
				업체구분 : <% DrawBrandGubunCombo "userdiv", userdiv %>
				&nbsp;
				<br>
				<% else %>
				<input type="hidden" name="usingonly" value="<%= usingonly %>">
				<input type="hidden" name="userdiv" value="<%= userdiv %>">
				<% end if %>
				브랜드ID : <input type="text" class="text" name="rect" value="<%= rect %>" Maxlength="32" size="16">
				&nbsp;&nbsp;
				스트리트명(한글) : <input type="text" class="text" name="socname_kr" value="<%= socname_kr %>" Maxlength="32" size="16">
				&nbsp;&nbsp;
				회사명 : <input type="text" class="text" name="crect" value="<%= crect %>" Maxlength="32" size="16">
				<% if (Not C_IS_FRN_SHOP) THen %>
				&nbsp;&nbsp;
				담당자명 : <input type="text" class="text" name="mrect" value="<%= mrect %>" Maxlength="32" size="16">
				<% end if %>
	        </td>
	        <td align="right">
	        	<input type="image" src="/lib/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/lib/images/tbl_blue_round_05.gif"></td>
	</tr>

</table>

<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="#BABABA"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/lib/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
        		<tr>
        			<TD width="8"><A href="?rect=a&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">A</A></TD>
					<TD width="8"><A href="?rect=b&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">B</A></TD>
					<TD width="8"><A href="?rect=c&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">C</A></TD>
					<TD width="8"><A href="?rect=d&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">D</A></TD>
					<TD width="8"><A href="?rect=e&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">E</A></TD>
					<TD width="8"><A href="?rect=f&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">F</A></TD>
					<TD width="8"><A href="?rect=g&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">G</A></TD>
					<TD width="8"><A href="?rect=h&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">H</A></TD>
					<TD width="8"><A href="?rect=i&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">I</A></TD>
					<TD width="8"><A href="?rect=j&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">J</A></TD>
					<TD width="8"><A href="?rect=k&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">K</A></TD>
					<TD width="8"><A href="?rect=l&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">L</A></TD>
					<TD width="8"><A href="?rect=m&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">M</A></TD>
					<TD width="8"><A href="?rect=n&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">N</A></TD>
					<TD width="8"><A href="?rect=o&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">O</A></TD>
					<TD width="8"><A href="?rect=p&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">P</A></TD>
					<TD width="8"><A href="?rect=q&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">Q</A></TD>
					<TD width="8"><A href="?rect=r&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">R</A></TD>
					<TD width="8"><A href="?rect=s&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">S</A></TD>
					<TD width="8"><A href="?rect=t&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">T</A></TD>
					<TD width="8"><A href="?rect=u&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">U</A></TD>
					<TD width="8"><A href="?rect=v&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">V</A></TD>
					<TD width="8"><A href="?rect=w&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">W</A></TD>
					<TD width="8"><A href="?rect=x&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">X</A></TD>
					<TD width="8"><A href="?rect=y&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">Y</A></TD>
					<TD width="8"><A href="?rect=z&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">Z</A></TD>
					<TD width="8"><A href="?rect=etc&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>&isjsdomain=<%=vIsJSDomain%>">etc</A></TD>
					<TD></TD>
        		</tr>
        	</table>
        </td>
       	<td align="right">
       		총 <%= opartner.FtotalCount %>건
       	</td>
        <td width="10" align="left" background="/lib/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- 표 중간바 끝-->



<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td >구분</td>
		<td >브랜드ID</td>
		<td >스트리트명(한글)<br>스트리트명(영문)</td>
		<td >회사명</td>
		<td >담당자</td>
	</tr>
	<% for i=0 to opartner.FresultCount-1 %>
	<% if opartner.FPartnerList(i).Fisusing="Y"	then %>
	<tr bgcolor="#FFFFFF">
	<% else %>
	<tr bgcolor="#EEEEEE">
	<% end if %>
		<td align="center"><%= opartner.FPartnerList(i).GetUserDivName %></a></td>
		<td><a href="javascript:selectThis('<%= opartner.FPartnerList(i).FID %>')"><%= opartner.FPartnerList(i).FID %></a></td>
		<td>
			<%= opartner.FPartnerList(i).FSocName_Kor %><br>
			<%= opartner.FPartnerList(i).FSocName %>
		</td>
		<td><a href="javascript:selectThis('<%= opartner.FPartnerList(i).FID %>')"><%= opartner.FPartnerList(i).Fcompany_name %></a></td>
		<td align="center"><%= opartner.FPartnerList(i).Fmanager_name %></td>
		
	</tr>
	<% next %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/lib/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if opartner.HasPreScroll then %>
			<a href="javascript:NextPage('<%= opartner.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + opartner.StartScrollPage to opartner.FScrollCount + opartner.StartScrollPage - 1 %>
    			<% if i>opartner.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if opartner.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>

        </td>
        <td width="10" align="left" background="/lib/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/lib/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/lib/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/lib/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set opartner = Nothing
%>


<!-- #include virtual="/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->