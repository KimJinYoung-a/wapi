<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<%
' http://www.10x10.co.kr/lib/email/bank_ok_mail.asp?buyemail=lcseung@yahoo.com&buyname=이철승
        dim sql,discountrate,subtotalprice
        dim mailfrom, mailto, mailtitle, mailcontent,buyemail,buyname
 
        buyemail = request.form("buyemail")
        buyname = request.form("buyname")
'        buyemail = request("buyemail")
'        buyname = request("buyname")

        if(buyemail="") then
            response.write("주문자 이메일이 넘어오지 않았습니다.")
            dbget.close()	:	response.End
        end if
        if(buyname="") then
            response.write("주문자 이름이 넘어오지 않았습니다.")
            dbget.close()	:	response.End
        end if

        mailcontent = sendmailbankok(buyemail,buyname)

'        response.write mailcontent

        response.write "S_OK"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->