<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim code, shop_id
code    = request("code")
shop_id = request("shop_id")
rw "code : " & code
rw "shopId : " & shop_id
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->