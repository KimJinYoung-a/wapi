<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbDatamart_dbget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=dbDatamart_rsget></OBJECT>
<%
dbDatamart_dbget.Open Application("db_Datamart")
%>
