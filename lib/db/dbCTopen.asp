<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbCTget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsCTget></OBJECT>

<%
'/���� �ֱ��� ������Ʈ ���� ������ ó�� '2011.11.11 �ѿ�� ����
'/������� ������ �ֽð� ������ ���� �ּ���
Call serverupdate_underconstruction()

dbCTget.Open Application("db_appWish")
%>
