<%
''check referer
dim refIP : refIP = request.ServerVariables("REMOTE_ADDR")

'if (LEFT(refIP,11)<>"61.252.133.") and (LEFT(refIP,12)<>"115.94.163.4") then 
'    response.end
'end if




dim disp : disp=request("disp")
''response.write "���"&disp
dim dispname
dim ifontsize : ifontsize = 8

if (disp="101") then
    dispname="�����ι���"
elseif (disp="102") then
    dispname="������/�ڵ���"
elseif (disp="103") then
    dispname="ķ��/Ʈ����"
elseif (disp="104") then
    dispname="����"
elseif (disp="121") then
    dispname="����/����"
elseif (disp="122") then
    dispname="����/�ö��"
elseif (disp="120") then
    dispname="�к긯/����"
elseif (disp="112") then
    dispname="Űģ"
elseif (disp="119") then
    dispname="Ǫ��"
elseif (disp="117") then
    dispname="�м��Ƿ�"
elseif (disp="116") then
    dispname="����/����/�־�"
elseif (disp="118") then
    dispname="��Ƽ"
elseif (disp="115") then
    dispname="���̺�/Ű��"
elseif (disp="110") then
    dispname="Cat&Dog"
else
    dispname=disp
end if


dim photoCmd : photoCmd = "cmd=text&fontsize="&ifontsize&"&x=0&y=0&w=200&h=30&color=000000&fontstyle=1&text="&dispname

response.redirect "http://thumbnail.10x10.co.kr/webimage/common/blank200.jpg?"&photoCmd



%>