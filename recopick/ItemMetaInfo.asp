<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% session.CodePage = "65001" %>
<% response.charset = "utf-8" %>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->

<%


	'// 레코픽 서비스 종료에 따른 메타 페이지 서비스 종료(150630 원승현)
	response.End

	Dim query1, vItemid, vItemName, vItemImgUrl, vItemDescription, vBrand, vPrice, vorgprice, vSalechk, vlimitno, vlimityn, vlimitsold, isSoldout, vsellyn


	vItemid = request("item_id")


	If vItemid="" Or IsNull(vItemid) Then
		response.write "<error>잘못된 접근 입니다.</error>"
		response.End
	End If

    ''2015/06/15 추가(eastone)
    if Not(isNumeric(vItemid)) then
    	response.write "<error>잘못된 접근 입니다(2).</error>"
    	dbCTget.close()	:	response.end
    end if

	query1 = " Select i.itemid, i.itemname, icon1image, c.designercomment, m.socname, i.sellcash, i.orgprice, i.sailyn, limitno, limityn, limitsold, sellyn "
	query1 = query1 + " From db_AppWish.dbo.tbl_item i "
	query1 = query1 + " inner join db_AppWish.dbo.tbl_item_contents c on i.itemid = c.itemid "
	query1 = query1 + " inner join db_AppWish.dbo.tbl_user_c m on i.makerid = m.userid "
	query1 = query1 + " Where i.itemid='"&vItemid&"' "

	rsCTget.Open query1,dbCTget,1

	If Not(rsCTget.bof Or rsCTget.eof) Then
		vItemName = rsCTget("itemname")
		vItemImgUrl = "http://webimage.10x10.co.kr/image/icon1/" + Num2Str(CStr(Clng(vItemid) \ 10000),2,"0","R")  + "/" + rsCTget("icon1image")
		vItemDescription = "생활감성채널 텐바이텐 - " & Replace(rsCTget("itemname") & " " & Trim(rsCTget("designercomment")),"""","")
		vBrand = rsCTget("socname")
		vPrice = rsCTget("sellcash")
		vorgprice = rsCTget("orgprice")
		vSalechk = rsCTget("sailyn")
		vlimitno = rsCTget("limitno")
		vlimityn = rsCTget("limityn")
		vlimitsold = rsCTget("limitsold")
		vsellyn =  rsCTget("sellyn")
	Else
		vItemName = ""
		vItemImgUrl = ""
		vItemDescription = ""
		vBrand = ""
		vPrice = ""
		vorgprice = ""
		vSalechk = ""
		vlimitno = ""
		vlimityn = ""
		vlimitsold = ""
		vsellyn =  ""
   End If


	'// 솔드아웃 체크
	IF vlimitno<>"" and vlimitsold<>"" Then
		isSoldOut = (vsellyn<>"Y") or ((vlimityn = "Y") and (clng(vlimitno)-clng(vlimitsold)<1))
	Else
		isSoldOut = (vsellyn<>"Y")
	End If

	Function Num2Str(inum,olen,cChr,oalign)
		dim i, ilen, strChr

		ilen = len(Cstr(inum))
		strChr = ""

		if ilen < olen then
			for i=1 to olen-ilen
				strChr = strChr & cChr
			next
		end if

		'결합방법에따른 결과 분기
		if oalign="L" then
			'왼쪽기준
			Num2Str = inum & strChr
		else
			'오른쪽 기준 (기본값)
			Num2Str = strChr & inum
		end if

	End Function

%>

<html>
<head>
	<meta http-equiv="content-type" content="text/html; charset=UTF-8">
	<meta property="og:title" content="<%=Replace(vItemName,"""","")%>">
	<meta property="og:image" content="<%=vItemImgUrl%>">
	<meta property="og:description" content="<%=vItemDescription%>">
	 
	<%' 상품의 저자, 메이커가 있는 경우, 아래 태그를 추가 해 주세요 %>
	<meta name="author" content="<%=vBrand%>">
	 
	<%' 상품의 가격 정보가 있는 경우, 상품의 가격 정보에 맞게 아래 태그를 추가 해 주세요. %>
	<meta property="product:price:amount" content="<%=vorgprice%>">
	<meta property="product:price:currency" content="KRW">

	<%' 상품의 할인 가격이 존재한다면 %>
	<% If vSalechk  = "Y" Then %>
	<meta property="product:sale_price:amount" content="<%=vPrice%>">
	<meta property="product:sale_price:currency" content="KRW">
	<% End If %>
	 
	<%' 상품이 품절 상태인 경우에만 아래 태그를 추가 해 주세요. oos(out of service)의 약자입니다. %>
	<% If isSoldOut Then %>
	<meta property="product:availability" content="oos">
	<% End If %>
</head>
<body>
</body>
</html>


<!-- #include virtual="/lib/db/dbCTclose.asp" -->