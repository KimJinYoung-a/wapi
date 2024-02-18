<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/ebay/utils.asp"-->
<!-- #include virtual="/outmall/ebay/ebayItemcls.asp"-->
<!-- #include virtual="/outmall/ebay/incEbayFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, o11st, oAuctionOpt, failCnt, chgSellYn, arrRows, skipItem, t11stGoodno, isAllRegYn, getMustprice
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, depth, isItemIdChk, vOptCnt
Dim isoptionyn, isText, i, vGubun, v, cateCode, sIdx, eIdx
itemid			= requestCheckVar(request("itemid"),100)
action			= request("act")
chgSellYn		= request("chgSellYn")
depth			= request("depth")
vGubun			= request("vGubun")
sIdx			= request("sIdx")
eIdx			= request("eIdx")
failCnt			= 0

' rw getToken(vGubun)
' response.end

'######################################################## ebay API ########################################################
If action = "GETSITECATE" Then			'1.1 Siteī�װ���ȸ API
''0.	GUBUN : A or G�� �ٲٱ�, DELETE FROM db_temp.dbo.tbl_ebay_siteCategory �ؼ� ����
''1.	http://localhost:11117/outmall/ebay/actEbayCommReq.asp?cmdparam=GETSITECATE&depth=1&vGubun=G ȣ�� 1depthó��
''2.	http://localhost:11117/outmall/ebay/actEbayCommReq.asp?cmdparam=GETSITECATE&vGubun=G&depth=2 ȣ�� 2depthó��
''3.	http://localhost:11117/outmall/ebay/actEbayCommReq.asp?cmdparam=GETSITECATE&vGubun=G&depth=3 ȣ�� 3depthó��
''4.	http://localhost:11117/outmall/ebay/actEbayCommReq.asp?cmdparam=GETSITECATE&vGubun=G&depth=4 ȣ�� 4depthó��
''5.	http://localhost:11117/outmall/ebay/actEbayReq.asp?cmdparam=GETSITECATE&depth=o&vGubun=G ȣ�� ����Լ��� ���� ���̺� �����
	If depth = "1" Then
		Call fnEbaytGetSiteCate(getToken(vGubun), depth, cateCode, vGubun, iErrStr)
	ElseIf depth = "o" Then
		Call fnEbaytMakeSiteCate(vGubun)
	Else
		Call fnEbaytGetSiteCate(getToken(vGubun), depth, itemid, vGubun, iErrStr)
	End If
ElseIf action = "GETCATE" Then			'1.2 ESMī�װ���ȸ API
'	http://localhost:11117/outmall/ebay/actEbayCommReq.asp?cmdparam=GETCATE&vGubun=G ȣ��..������ A or G�� A�� ���� ȣ���� �ʿ� ����
	Call fnEbaytGetCate(getToken(vGubun), "0", iErrStr)
ElseIf action = "GETMATCHCATE" Then		'1.3 Site-ESMī�װ���Ī��ȸ API
'	http://localhost:11117/outmall/ebay/actEbayCommReq.asp?cmdparam=GETMATCHCATE&vGubun=A ȣ��..������ A or G�� A�� ���� ȣ���� �ʿ� ����
	Call fnEbaytGetMatchCate(getToken(vGubun), itemid, iErrStr)
End If

If iErrStr <> "" Then
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str = '"&iErrStr&"<br>' + str " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 10);" & vbCrLf &_
					"</script>"
End If
'###################################################### LotteiMall API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
