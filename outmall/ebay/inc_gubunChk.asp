<%
Dim CMALLNAME
If request("vGubun") = "A" Then
	CMALLNAME = "auction1010"
ElseIf request("vGubun") = "G" Then
	CMALLNAME = "gmarket1010"
End If
%>