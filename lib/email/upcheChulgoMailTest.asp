<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- include virtual="/lib/email/smsLib.asp"-->
<!-- include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/email/mailLib2.asp"-->
<!-- #include virtual="/lib/email/mailFunc_Designer.asp"-->
<!-- include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<!-- include virtual="/cscenter/lib/csAsfunction.asp" -->
<%
    call fcSendMailFinish_Dlv_Designer("11070858898","therain001")
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->