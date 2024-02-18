<html>
<head>
<title>T</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">

<script type="text/javascript" src="/lib/jstest/t1.js"></script>
<script type="text/javascript" src="/lib/jstest/t2.js"></script>
<script type="text/javascript" src="/lib/jstest/t3.js"></script>


<script type="text/javascript" src="/lib/jstest2/t5.js"></script>
<script type="text/javascript" src="/lib/jstest2/t6.js"></script>
<script type="text/javascript" src="/lib/jstest2/t7.js"></script>
<script type="text/javascript" src="/lib/jstest2/t8.js"></script>

<script type="text/javascript" src="/lib/jstest/t4.js"></script>
<!--
<script type="text/javascript" src="/lib/jstest/t10.js"></script>
<script type="text/javascript" src="/lib/jstest/t11.js"></script>
<script type="text/javascript" src="/lib/jstest/t12.js"></script>
<script type="text/javascript" src="/lib/jstest/t13.js"></script>
<script type="text/javascript" src="/lib/jstest/t14.js"></script>
<script type="text/javascript" src="/lib/jstest/t15.js"></script>
-->
<!--
<script type="text/javascript" src="/lib/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/lib/jstest/t1.js"></script>
-->

</head>
<body>

<%
'response.write "<br>"
'response.write session.sessionID
response.write "<br>"
response.write request.servervariables("remote_addr")
response.write "<br>"

dim tm : tm = split(Timer,".")
dim buf
if ubound(tm)>0 then 
    buf=tm(1) 
else 
    buf=0
end if

response.write now() & "." &  buf
%>
<!--
<img src="http://webimage.10x10.co.kr/eventIMG/2015/66069/group20150911175943.jpg" width="120"><br>
<img src="http://image.ithinkso.co.kr/files/itemimage/9000/9661/150x150/20111215154243.jpg" width="120"><br>
<img src="http://thumbnail.10x10.co.kr/webimage/image/basic600/135/B001356780.jpg?cmd=thumb&w=100&h=120&fit=true&ws=false"><br>
<img src="http://thumbnail.10x10.co.kr/webimage/image/basic600/135/B001356780.jpg?cmd=thumb&w=101&h=120&fit=true&ws=false"><br>
<img src="http://thumbnail.10x10.co.kr/webimage/image/basic600/135/B001356780.jpg?cmd=thumb&w=102&h=120&fit=true&ws=false"><br>
<img src="http://thumbnail.10x10.co.kr/webimage/image/basic600/135/B001356780.jpg?cmd=thumb&w=103&h=120&fit=true&ws=false"><br>
<img src="http://thumbnail.10x10.co.kr/webimage/image/basic600/135/B001356780.jpg?cmd=thumb&w=104&h=120&fit=true&ws=false"><br>
---------------------------------------->

<br>
<a href="javascript:document.location.reload()">reload </a>

<br><br>


<br><br>
<a href="http://52.79.73.177:5000/">aws </a>

<br><br>
<a href="http://m.11st.co.kr">11st </a>

<br><br><br><br><br><br><br><br><br><br><br>
...............<br>
<br><br><br><br>
...............<br>
</body>
</html>