<OBJECT RUNAT=server PROGID=ADODB.Connection id=dbSTSget></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset  id=rsSTSget></OBJECT>

<%
'/서버 주기적 업데이트 위한 공사중 처리 '2011.11.11 한용민 생성
'/리뉴얼시 이전해 주시고 지우지 말아 주세요
Call serverupdate_underconstruction()

dbSTSget.Open Application("db_statistics")
%>
