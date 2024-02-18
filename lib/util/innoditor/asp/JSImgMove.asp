<%@ language=vbscript %>
<% option explicit %>
<%
	dim strUrl
	strUrl = request("url")
%>
<script type="text/javascript">
<!--
	parent.fnUploadResult("<%=strUrl%>");
//-->
</script>