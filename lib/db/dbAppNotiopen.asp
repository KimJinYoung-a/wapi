<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbAppNotiget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsAppNotiget></OBJECT>

<%
'/���� �ֱ��� ������Ʈ ���� ������ ó�� '2011.11.11 �ѿ�� ����
'/������� ������ �ֽð� ������ ���� �ּ���
Call serverupdate_underconstruction()

dbAppNotiget.Open Application("db_appNoti") 
%>
