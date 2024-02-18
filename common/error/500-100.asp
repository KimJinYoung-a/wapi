<%@ language="VBScript" %>
<%
  Option Explicit


  Const lngMaxFormBytes = 200

  Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
  Dim strMethod, lngPos, datNow, strQueryString, strURL

  If Response.Buffer Then
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/html"
    Response.Expires = 0
  End If

  Set objASPError = Server.GetLastError
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<HTML><HEAD><TITLE>10x10 페이지에 오류가 발생했습니다.</TITLE>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=euc-kr">
<STYLE type="text/css">
  BODY { font: 9pt/12pt 굴림 }
  H1 { font: 13pt/15pt 굴림; font-weight:bold; }
  H2 { font: 9pt/12pt 굴림 }
  A:link { color: red }
  A:visited { color: maroon }
</STYLE>
</HEAD>
<BODY>
<TABLE width="620" border="0" cellspacing="5">
<TR>
	<td width="117" valign="top"><img src="/fiximage/web2012/common/footer_logo.png" /></td>
	<TD>
		<h1>이 페이지를 표시할 수 없습니다.</h1>
		연결하려는 페이지에 오류가 발생되어 표시할 수 없습니다.
	</td>
</tr>
<tr>
	<td colspan="2">
		<hr>
		<p>다음과 같은 내용을 텐바이텐 <strong>시스템팀</strong>에게 전달해주세요.</p>
		<ul>
			<li>오류가 발생한 현재 페이지의 <strong>URL 주소</strong>를 알려주세요.</li>
			<li>아래 기술 정보를 복사해서 전달해 주세요.</li>
			<li>유입 과정 또는 발생 상황을 알려주시면 해결에 더욱 도움이 됩니다.</li>
		</ul>
		<h2>HTTP 500.100 - 내부 서버 오류: ASP 오류입니다.<br>IIS(인터넷 정보 서비스)</h2>
		<hr>
		<p><strong>기술 정보</strong> (시스템팀에게 전달해주세요.)</p>
		<ul>
			<li>오류 유형:<br> <%
			  Dim bakCodepage
			  on error resume next
				bakCodepage = Session.Codepage
				Session.Codepage = 1252
			  on error goto 0
			  Response.Write Server.HTMLEncode(objASPError.Category)
			  If objASPError.ASPCode > "" Then Response.Write Server.HTMLEncode(", " & objASPError.ASPCode)
			    Response.Write Server.HTMLEncode(" (0x" & Hex(objASPError.Number) & ")" ) & "<br>"
			  If objASPError.ASPDescription > "" Then
				Response.Write Server.HTMLEncode(objASPError.ASPDescription) & "<br>"
			  elseIf (objASPError.Description > "") Then
				Response.Write Server.HTMLEncode(objASPError.Description) & "<br>"
			  end if
			  blnErrorWritten = False
			  ' Only show the Source if it is available and the request is from the same machine as IIS
			  If objASPError.Source > "" Then
			    strServername = LCase(Request.ServerVariables("SERVER_NAME"))
			    strServerIP = Request.ServerVariables("LOCAL_ADDR")
			    strRemoteIP =  Request.ServerVariables("REMOTE_ADDR")
			    If (strServerIP = strRemoteIP) And objASPError.File <> "?" Then
			      Response.Write Server.HTMLEncode(objASPError.File)
			      If objASPError.Line > 0 Then Response.Write ", line " & objASPError.Line
			      If objASPError.Column > 0 Then Response.Write ", column " & objASPError.Column
			      Response.Write "<br>"
			      Response.Write "<font style=""COLOR:000000; FONT: 8pt/11pt courier new""><b>"
			      Response.Write Server.HTMLEncode(objASPError.Source) & "<br>"
			      If objASPError.Column > 0 Then Response.Write String((objASPError.Column - 1), "-") & "^<br>"
			      Response.Write "</b></font>"
			      blnErrorWritten = True
			    End If
			  End If
			  If Not blnErrorWritten And objASPError.File <> "?" Then
			    Response.Write "<b>" & Server.HTMLEncode(  objASPError.File)
			    If objASPError.Line > 0 Then Response.Write Server.HTMLEncode(", line " & objASPError.Line)
			    If objASPError.Column > 0 Then Response.Write ", column " & objASPError.Column
			    Response.Write "</b><br>"
			  End If
			%>
			</li>
			<li>브라우저 종류:<br> <%= Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT")) %> <br><br></li>
			<li>페이지:<br> <%
			  strMethod = Request.ServerVariables("REQUEST_METHOD")
			  Response.Write strMethod & " "
			  If strMethod = "POST" Then
			    Response.Write Request.TotalBytes & " bytes to "
			  End If
			  Response.Write Request.ServerVariables("SCRIPT_NAME")
			  Response.Write "</li>"
			  If strMethod = "POST" Then
			    Response.Write "<p><li>POST Data:<br>"
			    ' On Error in case Request.BinaryRead was executed in the page that triggered the error.
			    On Error Resume Next
			    If Request.TotalBytes > lngMaxFormBytes Then
			      Response.Write Server.HTMLEncode(Left(Request.Form, lngMaxFormBytes)) & " . . ."
			    Else
			      Response.Write Server.HTMLEncode(Request.Form)
			    End If
			    On Error Goto 0
			    Response.Write "</li>"
			  End If
			%> <br><br></li>
			<li>시간:<br> <%
			  datNow = Now()
			  Response.Write Server.HTMLEncode(FormatDateTime(datNow, 1) & ", " & FormatDateTime(datNow, 3))
			  on error resume next
				Session.Codepage = bakCodepage
			  on error goto 0
			%> </li>
		</ul>
	</TD>
</TR>
</TABLE>
</BODY>
</HTML>
