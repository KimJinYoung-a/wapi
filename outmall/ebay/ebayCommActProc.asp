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
If action = "GETSITECATE" Then			'1.1 Site카테고리조회 API
''0.	GUBUN : A or G로 바꾸기, DELETE FROM db_temp.dbo.tbl_ebay_siteCategory 해서 비우기
''1.	http://localhost:11117/outmall/ebay/actEbayCommReq.asp?cmdparam=GETSITECATE&depth=1&vGubun=G 호출 1depth처리
''2.	http://localhost:11117/outmall/ebay/actEbayCommReq.asp?cmdparam=GETSITECATE&vGubun=G&depth=2 호출 2depth처리
''3.	http://localhost:11117/outmall/ebay/actEbayCommReq.asp?cmdparam=GETSITECATE&vGubun=G&depth=3 호출 3depth처리
''4.	http://localhost:11117/outmall/ebay/actEbayCommReq.asp?cmdparam=GETSITECATE&vGubun=G&depth=4 호출 4depth처리
''5.	http://localhost:11117/outmall/ebay/actEbayReq.asp?cmdparam=GETSITECATE&depth=o&vGubun=G 호출 재귀함수로 실제 테이블 만들기
	If depth = "1" Then
		Call fnEbaytGetSiteCate(getToken(vGubun), depth, cateCode, vGubun, iErrStr)
	ElseIf depth = "o" Then
		Call fnEbaytMakeSiteCate(vGubun)
	Else
		Call fnEbaytGetSiteCate(getToken(vGubun), depth, itemid, vGubun, iErrStr)
	End If
ElseIf action = "GETCATE" Then			'1.2 ESM카테고리조회 API
'	http://localhost:11117/outmall/ebay/actEbayCommReq.asp?cmdparam=GETCATE&vGubun=G 호출..구분을 A or G라서 A를 따로 호출할 필요 없음
	Call fnEbaytGetCate(getToken(vGubun), "0", iErrStr)
ElseIf action = "GETMATCHCATE" Then		'1.3 Site-ESM카테고리매칭조회 API
'	http://localhost:11117/outmall/ebay/actEbayCommReq.asp?cmdparam=GETMATCHCATE&vGubun=A 호출..구분을 A or G라서 A를 따로 호출할 필요 없음
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
