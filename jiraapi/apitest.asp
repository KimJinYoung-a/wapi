<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.charset = "utf-8" %>
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/aspJSON1.17.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<script language="jscript" runat="server">
    function jsURLDecode(v){ return decodeURI(v); }
    function jsURLEncode(v){ return encodeURI(v); }
</script>
<%
dim param1 : param1=request("param1")
dim param2 : param2=request("param2")
dim param3 : param3=request("param3")
dim itemid : itemid=request("itemid")

dim oJson
'// json객체 선언
Set oJson = jsObject()
oJson("param1") = param1
oJson("param2") = param2
oJson("param3") = param3
'oJson.flush
Set oJson = Nothing

response.write "itemid="&itemid
%>