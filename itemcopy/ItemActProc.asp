<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/itemcopy/incItemFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, makerid, itemdiv, action, failCnt, arrRows
Dim iErrStr, strSql, SumErrStr, SumOKStr, i, strparam
Dim isExistMakerid, isExistItemdiv, isExistItemid, retCopyitemid, vidx
itemid			= requestCheckVar(request("itemid"),9)
makerid			= requestCheckVar(request("makerid"),32)
itemdiv			= requestCheckVar(request("itemdiv"),10)
action			= request("act")
failCnt			= 0

If action = "itemcopy" Then			'��ǰ����
	strSql = "EXEC [db_item].[dbo].[usp_API_itemcopy_Upd] 'I', '', '"& itemid &"', '"& session("ssBctId") &"', '', '', '' "
	rsget.CursorLocation = adUseClient
	rsget.CursorType = adOpenStatic
	rsget.LockType = adLockOptimistic
	rsget.Open strSql, dbget
	If Not(rsget.EOF or rsget.BOF) Then
		vidx = rsget("idx")
	End If
	rsget.Close

	isExistMakerid	= fnFindMakerid(makerid)
	isExistItemdiv	= fnFindItemdiv(itemdiv)
	isExistItemid	= fnFindItemid(itemid, retCopyitemid)

	If isExistMakerid = "N" Then
		iErrStr = "ERR||"&itemid&"||���� ������ �귣�尡 �ƴմϴ�.("& makerid &")"
		failCnt = failCnt + 1
	End If

	If isExistItemdiv = "N" Then
		iErrStr = "ERR||"&itemid&"||���� ������ ��ǰ������ �ƴմϴ�.("& itemdiv &")"
		failCnt = failCnt + 1
	End If

	If isExistItemid = "Y" Then
		iErrStr = "ERR||"&itemid&"||�̹� ������ ��ǰ�Դϴ�. ������ǰ�ڵ� : ("& retCopyitemid &")"
		failCnt = failCnt + 1
	End If

	If failCnt = 0 Then
		Call fnItemCopy(itemid, makerid, itemdiv, vidx, iErrStr)
	Else
		strSql = "EXEC [db_item].[dbo].[usp_API_itemcopy_Upd] 'U', '"& vidx &"', '', '', 'ERR', '"& Split(iErrStr,"||")(2) &"', ''"
		dbget.execute strSql
	End If
End If

response.write  "<script>" & vbCrLf &_
				"	var str, t; " & vbCrLf &_
				"	t = parent.document.getElementById('actStr') " & vbCrLf &_
				"	str = t.innerHTML; " & vbCrLf &_
				"	str = '"&iErrStr&"<br>' + str " & vbCrLf &_
				"	t.innerHTML = str; " & vbCrLf &_
				"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
				"</script>"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->