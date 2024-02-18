<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbget_Logistics></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsget_Logistics></OBJECT>
<%
dbget_Logistics.Open Application("db_alogistics")
%>
