<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<script type="text/javascript" src="/lib/js/jquery-1.7.1.min.js"></script>
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'#######################################################
'	History	: 2015.07.20 원승현 생성
'	Description : 텐바이텐 sns 가져오기
'	/http://wapi.10x10.co.kr/sns/snsdataselect.asp
'#######################################################

dim eCode
IF application("Svr_Info") = "Dev" THEN
	eCode   =  66157		''PLAY #31. 물 좋은 곳
Else
	eCode   =  71532		''PLAY #31. 물 좋은 곳
End If

On Error Resume Next

Dim instaSns, userid, tagsTxt, vMaxId, vJsonSnsUrl, vChkSnsId, refer
Dim  i, j
	j = 0

dim iTxtVal

vChkSnsId = ""
userid=getloginuserid()
vMaxId = request("MaxId")
tagsTxt = Server.URLencode("텐바이텐향기")
'vJsonSnsUrl = "https://api.instagram.com/v1/users/711689678/media/recent/?access_token=711689678.19795ba.a63ebd633d9c4e93a66b25f1e196850f&max_id="&vMaxId

''태그(tagsTxt)로 긁어오기
'vJsonSnsUrl = "https://api.instagram.com/v1/tags/"&tagsTxt&"/media/recent/?access_token=711689678.19795ba.a63ebd633d9c4e93a66b25f1e196850f&max_id="&vMaxId
vJsonSnsUrl = "https://api.instagram.com/v1/tags/"&tagsTxt&"/media/recent?access_token=711689678.19795ba.a63ebd633d9c4e93a66b25f1e196850f&max_id="&vMaxId

''텐바이텐 계정꺼 긁어오기
'vJsonSnsUrl = "https://api.instagram.com/v1/users/self/media/recent/?count=10&access_token=711689678.19795ba.a63ebd633d9c4e93a66b25f1e196850f&max_id="&vMaxId
																										


'If userid="thensi7" Or userid="kobula" Or userid="bborami" Or userid="tozzinet" Or userid="baboytw" Or userid="ppono2" Then
'Else
'	response.write "<script>alert('관계자만 볼 수 있는 페이지 입니다.');window.close();</script>"
'	response.End
'End If

Function chkSnsins(ckid)
	Dim sqlstr

	sqlstr = "select count(*) "
	sqlstr = sqlstr & " from [db_Appwish].[dbo].[tbl_snsSelectData]"
	sqlstr = sqlstr & " Where snsid='"&ckid&"' and evt_code="& eCode &""
	
	'response.write sqlstr & "<br>"
	rsCTget.Open sqlstr, dbCTget, adOpenForwardOnly, adLockReadOnly
		chkSnsins = rsCTget(0)
	rsCTget.close
End Function
%>
<html>
<head>
<style type="text/css">
.table {width:900px; margin:0 auto; font-family:'malgun gothic'; border-collapse:collapse;}
.table th {padding:12px 0; font-size:13px; font-weight:bold;  color:#fff; background:#444;}
.table td {padding:12px 3px; font-size:12px; border:1px solid #ddd; border-bottom:2px solid #ddd;}
.table td.lt {text-align:left; padding:12px 10px;}
.tBtn {display:inline-block; border:1px solid #2b90b6; background:#03a0db; padding:0 10px 2px; line-height:26px; height:26px; font-weight:bold; border-radius:5px; color:#fff !important;}
.tBtn:hover {text-decoration:none;}
.table td input {border:1px solid #ddd; height:30px; padding:0 3px; font-size:14px; color:#ec0d02; text-align:right;}
</style>

<script>

function goActionSns(snsid)
{
	$.ajax({
		type:"GET",
		url:"/sns/incInsDelAjax.asp?snsid="+snsid+"&JsonSnsUrl=<%=Server.URLencode(vJsonSnsUrl)%>",
		dataType: "text",
		async:false,
		cache:true,
		success : function(Data, textStatus, jqXHR){
			if (jqXHR.readyState == 4) {
				if (jqXHR.status == 200) {
					if(Data!="") {
						var str;
						for(var i in Data)
						{
							 if(Data.hasOwnProperty(i))
							{
								str += Data[i];
							}
						}
						str = str.replace("undefined","");
						res = str.split("|");
						if (res[0]=="OK")
						{
							if (res[1]=="1")
							{
								$("#tr"+res[2]).css("background-color","#8DE2DA");
								$("#actionTxt"+res[2]).empty().html("삭제");
							}
							else
							{
								$("#tr"+res[2]).css("background-color","");
								$("#actionTxt"+res[2]).empty().html("추가");
							}
						}
						else
						{
							errorMsg = res[1].replace(">?n", "\n");
							alert(errorMsg );
							return false;
						}
					} else {
						alert("오류가 발생하였습니다.");
						return false;
					}
				}
			}
		},
		error:function(jqXHR, textStatus, errorThrown){
			alert("오류가 발생하였습니다.");
			var str;
			for(var i in jqXHR)
			{
				 if(jqXHR.hasOwnProperty(i))
				{
					str += jqXHR[i];
				}
			}
			alert(str);
			document.location.reload();
			return false;
		}
	});
}
</script>
</head>
<body>
<%
'	Set instaSns = JSON.parse(getJsonAsp("https://api.instagram.com/v1/tags/"&tagsTxt&"/media/recent/?access_token=711689678.19795ba.a63ebd633d9c4e93a66b25f1e196850f",""))
	Set instaSns = JSON.parse(getJsonAsp(""&vJsonSnsUrl&"",""))
%>

<table class="table" style="width:70%;">
	<colgroup>
		<col width="5%" />
		<col width="10%" />
		<col width="15%" />
		<col width="30%" />
		<col width="5%" />
		<col width="5%" />
	</colgroup>
	<tr align="center" bgcolor="#E6E6E6">
		<th><strong>관리</strong></th>
		<th><strong>id</strong></th>
		<th><strong>이미지</strong></th>
		<th><strong>내용</strong></th>
		<th><strong>sns유저아이디</strong></th>
		<th><strong>sns유저명</strong></th>
	</tr>

	<div id="setCont">
	<%
		If IsNull(instaSns) Or instaSns<>"" Then
			For i=0 To 30
			If instaSns.data.Get(i).caption.id ="" Then 
				Exit For
			End If
	%>
	<tr align="center" id="tr<%=instaSns.data.Get(i).caption.id%>" <% If chkSnsins(instaSns.data.Get(i).caption.id)>=1 Then %>style="background-color:#8DE2DA"<% End If %>>
		<td>
		
			<a href="" onclick="goActionSns('<%=instaSns.data.Get(i).caption.id%>');return false;"><span id="actionTxt<%=instaSns.data.Get(i).caption.id%>"><% If chkSnsins(instaSns.data.Get(i).caption.id)>=1 Then %>삭제<% Else %>추가<% End If %></span></a>
		</td>
		<td><%=instaSns.data.Get(i).caption.id%></td>
		<td><img src='<%=instaSns.data.Get(i).images.low_resolution.url%>' alt='' width="120" height="120"/></td>
		<td><%=chrbyte(instaSns.data.Get(i).caption.text, 185, "N")%></td>
		<td><%=instaSns.data.Get(i).user.id%></td>
		<td><%=instaSns.data.Get(i).user.username%></td>
	</tr>
	<%
			Next
		End If
	%>
	</div>
</table>
<br>
<table class="table" style="width:30%;height:10px;">

	<tr align="center">
		<td><% If vMaxId <>"" Then %><a href="" onclick="history.back();return false;"><- 이전</a>&nbsp;&nbsp;&nbsp;<% End If %><a href="/sns/snsDataSelect.asp?maxid=<%=instaSns.pagination.next_max_id%>">다음 -></a></td>
	</tr>

</table>

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<%
'// asp json 거시기
Function getJsonAsp(url, param)
	Dim objHttp
	Dim strJsonText
	Set objHttp = server.CreateObject("Microsoft.XMLHTTP")
	If IsNull(objHttp) Then
		response.write "서버 연결 오류"
		response.End
	End If
	objHttp.Open "Get", url, False
	objHttp.SetRequestHeader "Content-Type","text/plain"
	objHttp.Send param
	strJsonText = objHttp.responseText
	Set objHttp = Nothing

	getJsonAsp = strJsonText

End Function

Set instaSns = Nothing
%>