<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbSTSget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsSTSget></OBJECT>

<%
'/���� �ֱ��� ������Ʈ ���� ������ ó�� '2011.11.11 �ѿ�� ����
'/������� ������ �ֽð� ������ ���� �ּ���
Call serverupdate_underconstruction()

dbSTSget.Open Application("db_statistics")
%>
