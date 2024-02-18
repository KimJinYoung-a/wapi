<%
''check referer
dim refIP : refIP = request.ServerVariables("REMOTE_ADDR")

'if (LEFT(refIP,11)<>"61.252.133.") and (LEFT(refIP,12)<>"115.94.163.4") then 
'    response.end
'end if




dim disp : disp=request("disp")
''response.write "상그"&disp
dim dispname
dim ifontsize : ifontsize = 8

if (disp="101") then
    dispname="디자인문구"
elseif (disp="102") then
    dispname="디지털/핸드폰"
elseif (disp="103") then
    dispname="캠핑/트래블"
elseif (disp="104") then
    dispname="토이"
elseif (disp="121") then
    dispname="가구/조명"
elseif (disp="122") then
    dispname="데코/플라워"
elseif (disp="120") then
    dispname="패브릭/수납"
elseif (disp="112") then
    dispname="키친"
elseif (disp="119") then
    dispname="푸드"
elseif (disp="117") then
    dispname="패션의류"
elseif (disp="116") then
    dispname="가방/슈즈/주얼리"
elseif (disp="118") then
    dispname="뷰티"
elseif (disp="115") then
    dispname="베이비/키즈"
elseif (disp="110") then
    dispname="Cat&Dog"
else
    dispname=disp
end if


dim photoCmd : photoCmd = "cmd=text&fontsize="&ifontsize&"&x=0&y=0&w=200&h=30&color=000000&fontstyle=1&text="&dispname

response.redirect "http://thumbnail.10x10.co.kr/webimage/common/blank200.jpg?"&photoCmd



%>