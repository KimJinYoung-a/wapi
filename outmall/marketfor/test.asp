<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/base64.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%
Dim com, encodeStr
set Com = server.createobject("kjy8517AES.AES")
    encodeStr = Com.AESEncrypt256("������� : ����Ư���� ������ ������ 508(���ﵿ GS Ÿ��)\n ���� : GS ������", "8bc7f784046609702a21e4a47b6bd8cf")
    rw "Encript : " & encodeStr
    rw "Decript : " & Com.AESDecrypt256(encodeStr, "8bc7f784046609702a21e4a47b6bd8cf")
    response.end
set Com = nothing
%>