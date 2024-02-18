<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbAnalget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsAnalget></OBJECT>

<%
dbAnalget.Open Application("db_analyze") 
%>
