<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ǰ ǰ����ǰ/������� �ȳ�
' History : �̻� ����
'           2020.10.27 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/lib/email/smslib.asp"-->

<%
dim kakaomsgstr, btnJson, smstitlestr, smsmsgstr, orderserial, itemName, buyhp
orderserial="01234567891"
itemName="test��ǰ��"
buyhp="010-9177-8708"
						    ' smstitlestr = "[�ٹ�����]�ֹ��Ͻ� ��ǰ �ù��ľ� ��ۺҰ� �ȳ�"
							' smsmsgstr = "[10x10] �ù��ľ� ��ۺҰ��ȳ�" & vbCrLf & vbCrLf
							' smsmsgstr = smsmsgstr & "�˼��մϴ�. ����" & vbCrLf
							' smsmsgstr = smsmsgstr & "�ù��ľ����� ���� ������ ������� �ù�߼��� ��ưԵǾ� �ȳ��帳�ϴ�." & vbCrLf
							' smsmsgstr = smsmsgstr & "���� ����簳 ���� ������ �˼����� ��Ȳ����" & vbCrLf
							' smsmsgstr = smsmsgstr & "��Ÿ������ �ֹ���� �ȳ��帮�� �� ���غ�Ź�帳�ϴ�." & vbCrLf
							' smsmsgstr = smsmsgstr & "�ֹ���ǰ�� ���� �ڵ� ��� �� ȯ�ҿ����Դϴ�." & vbCrLf & vbCrLf & vbCrLf
							' smsmsgstr = smsmsgstr & "�� �ֹ���ȣ : "& orderserial &"" & vbCrLf
							' smsmsgstr = smsmsgstr & "�� ��ǰ�� : "& itemName &"" & vbCrLf & vbCrLf & vbCrLf
							' smsmsgstr = smsmsgstr & "�����մϴ�." & vbCrLf
							' smsmsgstr = smsmsgstr & "����ϱ� : http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx="& orderserial &""
							' btnJson = "{""button"":[{""name"":""����ϱ�"",""type"":""WL"", ""url_mobile"":""http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx="& orderserial &"""}]}"
							' kakaomsgstr = "[10x10] �ù��ľ� ��ۺҰ��ȳ�" & vbCrLf & vbCrLf
							' kakaomsgstr = kakaomsgstr & "�˼��մϴ�. ����" & vbCrLf
							' kakaomsgstr = kakaomsgstr & "�ù��ľ����� ���� ������ ������� �ù�߼��� ��ưԵǾ� �ȳ��帳�ϴ�." & vbCrLf
							' kakaomsgstr = kakaomsgstr & "���� ����簳 ���� ������ �˼����� ��Ȳ����" & vbCrLf
							' kakaomsgstr = kakaomsgstr & "��Ÿ������ �ֹ���� �ȳ��帮�� �� ���غ�Ź�帳�ϴ�." & vbCrLf
							' kakaomsgstr = kakaomsgstr & "�ֹ���ǰ�� ���� �ڵ� ��� �� ȯ�ҿ����Դϴ�." & vbCrLf & vbCrLf & vbCrLf
							' kakaomsgstr = kakaomsgstr & "�� �ֹ���ȣ : "& orderserial &"" & vbCrLf
							' kakaomsgstr = kakaomsgstr & "�� ��ǰ�� : "& itemName &"" & vbCrLf & vbCrLf & vbCrLf
							' kakaomsgstr = kakaomsgstr & "�����մϴ�."
							' Call SendKakaoCSMsg_LINK("",buyhp,"1644-6030","KC-0025",kakaomsgstr,"LMS", smstitlestr, smsmsgstr,btnJson,"","")
' call SendNormalLMS("01091778708", "����1", "1644-6030", "����1")
' call SendNormalLMSTimeFix("01091778708", "����2", "1644-6030", "����2")
' call SendNormalSMS_LINK("01091778708", "1644-6030", "����3")
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
