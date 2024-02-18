<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Sub drawSelectBoxDesignerwithName(selectBoxName,selectedId)
	Dim strRst
	strRst = "<input type=""text"" class=""text"" name=""" & selectBoxName & """ id=""[off,off,off,off][브랜드ID]"" value=""" & selectedId & """ size=""20"" >" & vbCrLf
	strRst = strRst & "<input type=""button"" class=""button"" value=""IDSearch"" onclick=""jsSearchBrandID(this.form.name,'" & selectBoxName & "');"" >"
	Response.Write strRst
End Sub

Function CurrURL()
	CurrURL = Request.ServerVariables("PATH_INFO")
End Function

Sub drawSelectBoxCoWorker(byval selectBoxName, selectedId)
   dim tmp_str,query1
   %>
   <select class='select' name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
   <%
   query1 = " select userid, username from"
   query1 = query1 + " [db_partner].[dbo].tbl_user_tenbyten "
   query1 = query1 + " where  isusing= 1 and statediv = 'Y' and part_sn in('11','13','14','15','16') and userid <> '' "
   query1 = query1 + " order by username asc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='" + rsget("userid") + "' "&tmp_str&">" + db2html(rsget("username")) + " (" + rsget("userid") + ")</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end Sub

Sub SelectBoxBrandCategory(byval selectname, byval selectedId)
   dim tmp_str,query1

   if IsNULL(selectedId) then selectedId=""

   %><select class='select' name="<%= selectname %>">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_large, code_nm from [db_item].[dbo].tbl_Cate_large "
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " order by code_large Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("code_large")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code_large")&"' "&tmp_str&">"&rsget("code_nm")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

Sub DrawBrandGubunCombo(selectedname, selectedId)
   dim tmp_str,query1
   %>
   <select class='select' name="<%= selectedname %>" >
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
  	 <option value='02' <% if (selectedId="02") or (selectedId="03") or (selectedId="04") or (selectedId="05") or (selectedId="06") or (selectedId="07") or (selectedId="08") or (selectedId="13") then response.write " selected"%>>매입처(일반)</option>
  	 <option value='14' <%if selectedId="14" then response.write " selected"%>>아카데미</option>
  	 <option value='21' <%if selectedId="21" then response.write " selected"%>>출고처</option>
  	 <option value='20' <%if selectedId="20" then response.write " selected"%>>가맹점매입처</option>
  	 <option value='50' <%if selectedId="50" then response.write " selected"%>>제휴사(온라인)</option>
  	 <option value='95' <%if selectedId="95" then response.write " selected"%>>사용안함</option>
   </select>
  <%
End Sub



Sub OutmallAdminInfo(mall)
	Select Case mall
		Case "cjmall"
			response.write "<a href='http://partner.cjmall.com/login.jsp' target='_blank'>CJ몰Admin바로가기</a>"
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="iredfish") Then
				response.write "<font color='GREEN'>[ 411378 | store10x10 | cube1010$ ]</font>"
			End If
		Case "gsshop"
			response.write "<a href='https://withgs.gsshop.com/cmm/login' target='_blank'>GSShop Admin바로가기</a>"
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="iredfish") Then
				response.write "<font color='GREEN'>[  1003890 | cube1010* ]</font>"
			End If
	End Select
End Sub

Public Sub drawSelectBoxEtcLinkGbn(selectBoxName,selectedId,isDpAll)
	Dim tmp_str,query1
%>
	<select class="select" name="<%=selectBoxName%>" onchange="lgbn(this.value);">
	<% If (isDpAll) Then %>
		<option value='' <% If selectedId="" Then response.write " selected"%> >ALL</option>
	<% End If %>
<%
	query1 = " select linkgbn,valtype,linkDesc from db_outmall.dbo.tbl_OutMall_etcLinkGubun where 1=1 " & VBCRLF
	If poomok <> "05" Then
		query1 = query1 & " AND linkgbn <> 'infoDiv21Lotte' " & VBCRLF
	End If
	rsCTget.Open query1,dbCTget,1
	If  not rsCTget.EOF  Then
		Do until rsCTget.EOF
			If Lcase(selectedId) = Lcase(rsCTget("linkgbn")) Then
				tmp_str = " selected"
			End If
			response.write("<option value='"&rsCTget("linkgbn")&"' "&tmp_str&">" + rsCTget("linkDesc") + "</option>")
			tmp_str = ""
		rsCTget.MoveNext
		loop
	End If
	rsCTget.close
	response.write("</select>")
End Sub

Public Sub drawSelectBoxXSiteAPIPartner(selBoxName, selVal)
	Dim retStr
	retStr = "<select name='"&selBoxName&"' class='select'>"
	retStr = retStr & " <option value=''>전체"
	retStr = retStr & " <option value='interpark' "& CHKIIF(selVal="interpark","selected","") &" >인터파크"
	retStr = retStr & " <option value='lotteCom' "& CHKIIF(selVal="lotteCom","selected","") &" >롯데닷컴"
	retStr = retStr & " <option value='lotteimall' "& CHKIIF(selVal="lotteimall","selected","") &" >롯데iMall"
	retStr = retStr & " <option value='GSShop' "& CHKIIF(selVal="GSShop","selected","") &" >GSShop"
	retStr = retStr & " <option value='Homeplus' "& CHKIIF(selVal="Homeplus","selected","") &" >Homeplus"
	retStr = retStr & " </select> "
	response.write retStr
End Sub

'//날짜형식 2013-01-01 오후 03:00:00 형식을 2013-01-01 15:00:00로 변환		'/2013.04.22 한용민 생성
function dateconvert(dateval)
	dim tmpval
	if dateval = "" then exit function

	tmpval = year(dateval)
	tmpval = tmpval & "-" & Format00(2,month(dateval))
	tmpval = tmpval & "-" & Format00(2,day(dateval))
	tmpval = tmpval & " " & Format00(2,hour(dateval))
	tmpval = tmpval & ":" & Format00(2,minute(dateval))
	tmpval = tmpval & ":" & Format00(2,second(dateval))

	dateconvert = left(tmpval,19)
end function
%>