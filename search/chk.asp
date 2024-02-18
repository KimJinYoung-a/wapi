<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.charset = "utf-8" %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%

dim kwyArray : kwyArray= Array("다이어리","노호혼","문주란")
%>
<script type="text/javascript" src="/lib/js/jquery-1.7.1.min.js"></script>

<script>
function GetOneKey(ikeyword) {
	var str = "";
	str = $.ajax({
					type: "POST",
					url: "getdiffcheck.asp?q=" + encodeURIComponent(ikeyword),
					dataType: "text",
					async: false,
					cache: false
				}).responseText;

	return str;
}

function st(){
    var ikeywords = $("#keywords" ).val();
    
    var ikeywordsArr = ikeywords.split("\n")
    var kkkk='';
    for (var i=0;i<ikeywordsArr.length;i++){
        ikeyword = ikeywordsArr[i];
     
        var oneret = GetOneKey(ikeyword);
        kkkk+=(ikeyword+"\t"+oneret+"\n");
        
        sleep(20)
        
    }
    $("#retdata" ).val(kkkk)
    alert('fin')
}

function sleep (delay) {
   var start = new Date().getTime();
   while (new Date().getTime() < start + delay);
}


</script>
<textarea id="keywords" cols="80" rows="10"></textarea>
<textarea id="retdata" cols="80" rows="10"></textarea>
<input type="button" onClick="st()" value="START">