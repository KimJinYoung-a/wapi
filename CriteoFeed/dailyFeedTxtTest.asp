<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% session.CodePage = "65001" %>
<% Response.AddHeader "Content-type","text/xml" %>
<% Server.ScriptTimeOut = 1200 %>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
	'// 크리테오 피드 관련 xml 페이지
	Dim query1, vItemid, vItemName, vItemImgUrl, vItemDescription, vBrand, vPrice, vorgprice, vSalechk, vlimitno, vlimityn, vlimitsold, isSoldout, vsellyn, vExpirationDate, xmlPars, rss, Channel
	Dim title, channel_link, description, xItem, vCateFullName, xAppLink
	Dim savePath, rstMsg, strIdx, EndIdx, FileName, idx, arrRst, maxLoopCount, i


	savePath = server.mappath("/CriteoFeed/") + "\"


	vItemid = request("item_id")
	idx		= request("idx")

	If (idx="") Then
		response.write "<error>잘못된 접근 입니다.</error>"
       	dbCTget.close()	:	response.end
	End If

	FileName = "CriteoFeed"&idx&".xml"
	EndIdx = idx*50000
	If idx = 1 Then
		strIdx = (EndIdx-50000)+1
	Else
		strIdx = (EndIdx-50000)
	End If

	query1 = " Select num, itemid, itemname, icon1image, designercomment, socname, sellcash, orgprice, sailyn, limitno, limityn, "
	query1 = query1 + " limitsold, sellyn, catecode, basicimage600, basicimage, basicimage1000, mainimage, smallimage, sellEndDate, "
    query1 = query1 + " cate1, cate2, cate3, cate4, cate5, ISNULL(itempoint,0) AS itempoint "
	query1 = query1 + "From "
	query1 = query1 + " ( "
	query1 = query1 + " 	Select ROW_NUMBER() OVER( ORDER BY i.itemid ASC) AS NUM, i.itemid, i.itemname, icon1image, c.designercomment, m.socname, i.sellcash,  "
	query1 = query1 + " 	i.orgprice, i.sailyn, limitno, limityn, limitsold, sellyn, "
	query1 = query1 + " 	ci.catecode, i.basicimage600, i.basicimage, i.basicimage1000, i.mainimage, i.smallimage, convert(varchar(10), i.sellEndDate, 120) as sellEndDate, "
	query1 = query1 + " 	(Select top 1 catename From db_AppWish.dbo.tbl_display_cate Where catecode = substring(cast(ci.catecode as nvarchar(max)), 1, 3)) as cate1, "
	query1 = query1 + " 	case when len(ci.catecode)>6 then "
	query1 = query1 + " 		(Select top 1 catename From db_AppWish.dbo.tbl_display_cate WITH(NOLOCK) Where catecode = substring(cast(ci.catecode as nvarchar(max)), 1, 6)) "
	query1 = query1 + " 	else "
	query1 = query1 + " 		null "
	query1 = query1 + " 	end as cate2, "
	query1 = query1 + " 	case when len(ci.catecode)>9 then "
	query1 = query1 + " 		(Select top 1 catename From db_AppWish.dbo.tbl_display_cate WITH(NOLOCK) Where catecode = substring(cast(ci.catecode as nvarchar(max)), 1, 9)) "
	query1 = query1 + " 	else "
	query1 = query1 + " 		null "
	query1 = query1 + " 	end as cate3, "
	query1 = query1 + " 	case when len(ci.catecode)>12 then "
	query1 = query1 + " 		(Select top 1 catename From db_AppWish.dbo.tbl_display_cate WITH(NOLOCK) Where catecode = substring(cast(ci.catecode as nvarchar(max)), 1, 12)) "
	query1 = query1 + " 	else "
	query1 = query1 + " 		null "
	query1 = query1 + " 	end as cate4, "
	query1 = query1 + " 	case when len(ci.catecode)>15 then "
	query1 = query1 + " 		(Select top 1 catename From db_AppWish.dbo.tbl_display_cate WITH(NOLOCK) Where catecode = substring(cast(ci.catecode as nvarchar(max)), 1, 15)) "
	query1 = query1 + " 	else "
	query1 = query1 + " 		null "
	query1 = query1 + " 	end as cate5, "
    query1 = query1 + "     (SELECT CEILING(CONVERT(FLOAT,SUM(TotalPoint))/CONVERT(FLOAT,COUNT(*))*2)/2 "
    query1 = query1 + "     FROM db_AppWish.dbo.tbl_item_evaluate WITH(NOLOCK) WHERE itemid = i.itemid) AS itempoint "
	query1 = query1 + " 	From db_AppWish.dbo.tbl_item i  WITH(NOLOCK) "
	query1 = query1 + " 	inner join db_AppWish.dbo.tbl_item_contents c WITH(NOLOCK) on i.itemid = c.itemid  "
	query1 = query1 + " 	inner join db_AppWish.dbo.tbl_user_c m WITH(NOLOCK) on i.makerid = m.userid "
	query1 = query1 + " 	inner join db_AppWish.[dbo].[tbl_display_cate_item] ci WITH(NOLOCK) on i.itemid = ci.itemid And ci.isdefault = 'y' "
	query1 = query1 + " 	Where i.isusing='Y' And i.itemid <> 0 And i.sellyn = 'Y' And ci.depth>=2  And (c.sellcount>0 or datediff(day, i.regdate, getdate())<=20) And i.itemid not in ('1513445', '1603471', '1611936') "
	query1 = query1 + " 	And i.makerid not in ('imir10X10','clivia','secret01','drmtest1','pdccompany','cookie07','piooda07','sistalkkorea1','brandpick','cvkorea') "		'2017-02-01 15:15 김진영 수정
	query1 = query1 + " 	And i.adulttype = 0 "		'2019-12-16 성인용품 보내지 않음
	query1 = query1 + " )Tot "
	If strIdx<>"" Then
		query1 = query1 + " Where num >= "&strIdx&" And num < "&EndIdx&" "
	End If    
	dbCTget.CommandTimeOut = 480
	rsCTget.CursorLocation = adUseClient
	rsCTget.Open query1,dbCTget, adOpenForwardOnly, adLockReadOnly

	dim retCount : retCount = rsCTget.recordcount
	If Not(rsCTget.bof Or rsCTget.eof) Then

		maxLoopCount = Fix(retCount/50000)+1
        arrRst = rsCTget.getRows()

	End If
	rsCTget.close


    if isArray(arrRst) then

		Set xmlPars = Server.CreateObject("Msxml2.DOMDocument")
		xmlPars.preserveWhiteSpace = True
		xmlPars.appendChild(xmlPars.createProcessingInstruction("xml","version=""1.0"" encoding=""utf-8"""))

		Set rss = xmlPars.CreateElement("rss")
		rss.setAttribute "xmlns:g","http://base.google.com/ns/1.0"
		rss.setAttribute "version","2.0"
		xmlPars.AppendChild(rss)

		Set Channel = xmlPars.CreateElement("channel") 
		rss.AppendChild(Channel)

		'<title>정보 
		Set title = xmlPars.CreateElement("title") 
		Channel.AppendChild(title)
		Channel.childnodes(0).text = "10x10"  '제목

		'<link>정보 
		Set channel_link = xmlPars.CreateElement("link") 
		Channel.AppendChild(channel_link)
		Channel.childNodes(1).appendChild(xmlPars.createCDATASection("name_Cdata"))
		Channel.childnodes(1).childnodes(0).text = "http://www.10x10.co.kr"  '주소

		'<description>정보 
		Set description = xmlPars.CreateElement("description") 
		Channel.AppendChild(description) 
		Channel.childNodes(2).appendChild(xmlPars.createCDATASection("name_Cdata"))
		Channel.childnodes(2).childnodes(0).text = "10x10 Criteo Dynamic Ads Feed"  '설명

		for i=0 to ubound(arrRst,2)
            '0  - num
            '1  - itemid
            '2  - itemname
            '3  - icon1image
            '4  - designercomment
            '5  - socname
            '6  - sellcash
            '7  - orgprice
            '8  - sailyn
            '9  - limitno
            '10 - limityn
            '11 - limitsold
            '12 - sellyn
            '13 - catecode
            '14 - basicimage600
            '15 - basicimage
            '16 - basicimage1000
            '17 - mainimage
            '18 - smallimage
            '19 - sellEndDate
            '20 - cate1
            '21 - cate2
            '22 - cate3
            '23 - cate4
            '24 - cate5
            '25 - itempoint

			Set xItem = xmlPars.CreateElement("item")
			Channel.AppendChild(xItem) 

			xItem.AppendChild(xmlPars.CreateElement("g:id") ) 
			xItem.childnodes(0).text = arrRst(1,i)  '상품번호

			xItem.AppendChild(xmlPars.CreateElement("g:title"))
			xItem.childnodes(1).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(1).childnodes(0).text = Trim(Replace(Replace(Replace(arrRst(2,i), """", ""), Chr(32), ""), "", ""))  '상품명

			xItem.AppendChild(xmlPars.CreateElement("g:description") ) 
			xItem.childnodes(2).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(2).childnodes(0).text = "생활감성채널 텐바이텐 - " & stripHTML(Trim(Replace(Replace(Replace(arrRst(2,i), """", ""), Chr(32), ""), "", "")))  '상품설명


			vCateFullName = arrRst(20,i)
			If Not(arrRst(21,i)="" Or IsNull(arrRst(21,i))) Then
				vCateFullName = vCateFullName & " > "&arrRst(21,i)
			End If
			If Not(arrRst(22,i)="" Or IsNull(arrRst(22,i))) Then
				vCateFullName = vCateFullName & " > "&arrRst(22,i)
			End If
			If Not(arrRst(23,i)="" Or IsNull(arrRst(23,i))) Then
				vCateFullName = vCateFullName & " > "&arrRst(23,i)
			End If
			If Not(arrRst(24,i)="" Or IsNull(arrRst(24,i))) Then
				vCateFullName = vCateFullName & " > "&arrRst(24,i)
			End If
			If vCateFullName="" Or IsNull(vCateFullName) Then
				vCateFullName = ""
			End If
			xItem.AppendChild(xmlPars.CreateElement("g:product_type") ) 
			xItem.childnodes(3).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(3).childnodes(0).text = vCateFullName  '카테고리

			xItem.AppendChild(xmlPars.CreateElement("g:link") ) 
			xItem.childnodes(4).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(4).childnodes(0).text = "http://www.10x10.co.kr/shopping/category_prd.asp?utm_source=criteo&utm_medium=ad&utm_campaign=catalog&utm_term=criteo&rdsite=criteo&itemid="&arrRst(1,i)  '링크값(일단은 pc로 보내면 자동으로 모바일로 갈테니 pc주소로 보냄)

			'// url encode 제거
			'xItem.childnodes(4).childnodes(0).text = "https://tenten.app.link/3p?$3p=a_facebook&branch_ad_format=Product&$ios_url=http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&arrRst(1,i)&"&utm_source=instagram&utm_medium=referral&utm_campaign=posting&utm_term=producttag&rdsite=producttag&$android_url=http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&arrRst(1,i)&"&utm_source=instagram&utm_medium=referral&utm_campaign=posting&utm_term=producttag&rdsite=producttag&$desktop_url=http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&arrRst(1,i)&"&utm_source=instagram&utm_medium=referral&utm_campaign=posting&utm_term=producttag&rdsite=producttag&$deeplink_path=m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid="&arrRst(1,i)&"&utm_source=instagram&utm_medium=referral&utm_campaign=posting&utm_term=producttag&rdsite=producttag&isbranch=true&~feature=paid+advertising&$uri_redirect_mode=1&$ios_redirect_timeout=1000&$ios_passive_deepview=false&$android_passive_deepview=false&~ad_id={{ad.id}}&~ad_name={{ad.name}}&~ad_set_id={{adset.id}}&~ad_set_name={{adset.name}}&~campaign=posting&~campaign_id={{campaign.id}}&~keyword="&arrRst(1,i)
						

			If Trim(arrRst(14,i))="" Or IsNull(arrRst(14,i)) Then
				vItemImgUrl = "http://webimage.10x10.co.kr/image/basic/" + Num2Str(CStr(Clng(arrRst(1,i)) \ 10000),2,"0","R")  + "/" + arrRst(15,i)
			Else
				vItemImgUrl = "http://webimage.10x10.co.kr/image/basic600/" + Num2Str(CStr(Clng(arrRst(1,i)) \ 10000),2,"0","R")  + "/" + arrRst(14,i)
			End If
			If IsNull(vItemImgUrl) Or vItemImgUrl="" Then
				vItemImgUrl = "http://webimage.10x10.co.kr/image/main/" + Num2Str(CStr(Clng(arrRst(1,i)) \ 10000),2,"0","R")  + "/" + arrRst(17,i)
			End If
			If IsNull(vItemImgUrl) Or vItemImgUrl="" Then
				vItemImgUrl = "http://webimage.10x10.co.kr/image/small/" + Num2Str(CStr(Clng(arrRst(1,i)) \ 10000),2,"0","R")  + "/" + arrRst(18,i)
			End If

			xItem.AppendChild(xmlPars.CreateElement("g:image_link")) 
			xItem.childnodes(5).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(5).childnodes(0).text = vItemImgUrl  '이미지 링크

			xItem.AppendChild(xmlPars.CreateElement("g:condition")) 
			xItem.childnodes(6).text = "new"  '컨디션(new로 픽스)

			xItem.AppendChild(xmlPars.CreateElement("g:availability")) 
			xItem.childnodes(7).text = "in stock"  '재고현황(in stock로 픽스)

			xItem.AppendChild(xmlPars.CreateElement("g:price")) 
			xItem.childnodes(8).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(8).childnodes(0).text = arrRst(6,i)  '가격(sellcash값 보내줌)

			If arrRst(19,i)="" Or IsNull(arrRst(19,i)) Then
				vExpirationDate = "2039-01-18"
			Else
				vExpirationDate = arrRst(19,i)
			End If

			xItem.AppendChild(xmlPars.CreateElement("g:expiration_date")) 
			xItem.childnodes(9).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(9).childnodes(0).text = vExpirationDate  '만료일(보통은 없지않나??)

			xItem.AppendChild(xmlPars.CreateElement("g:brand")) 
			xItem.childnodes(10).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(10).childnodes(0).text = arrRst(5,i)  '브랜드명

			xItem.AppendChild(xmlPars.CreateElement("g:star")) 
			xItem.childnodes(11).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(11).childnodes(0).text = arrRst(25,i)  '상품평점            
			xItem.AppendChild(xmlPars.createTextNode(vbNewLine))

			Set xAppLink = xmlPars.CreateElement("applink")
			xItem.AppendChild(xAppLink)
			xAppLink.setAttribute "property","ios_url"
			xAppLink.setAttribute "content","tenwishapp://http://m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?utm_source=criteo&utm_medium=ad&utm_campaign=catalog&utm_term=criteo&rdsite=criteo&itemid="&arrRst(1,i)

			Set xAppLink = xmlPars.CreateElement("applink")
			xItem.AppendChild(xAppLink)
			xAppLink.setAttribute "property","ios_app_store_id"
			xAppLink.setAttribute "content","864817011"

			Set xAppLink = xmlPars.CreateElement("applink")
			xItem.AppendChild(xAppLink)
			xAppLink.setAttribute "property","ios_app_name"
			xAppLink.setAttribute "content","10x10"

			Set xAppLink = xmlPars.CreateElement("applink")
			xItem.AppendChild(xAppLink)
			xAppLink.setAttribute "property","android_url"
			xAppLink.setAttribute "content","tenwishapp://http://m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?utm_source=criteo&utm_medium=ad&utm_campaign=catalog&utm_term=criteo&rdsite=criteo&itemid="&arrRst(1,i)

			Set xAppLink = xmlPars.CreateElement("applink")
			xItem.AppendChild(xAppLink)
			xAppLink.setAttribute "property","android_package"
			xAppLink.setAttribute "content","kr.tenbyten.shopping"

			Set xAppLink = xmlPars.CreateElement("applink")
			xItem.AppendChild(xAppLink)
			xAppLink.setAttribute "property","android_app_name"
			xAppLink.setAttribute "content","10x10"            
		next
    End If

	'// XML파일 저장
	if (retCount>0) then
		xmlPars.save(savePath & FileName)
	end if 

'		rstMsg = "데이터 파일 [CriteoFeed.xml] 생성 완료"
	Set xmlPars = Nothing
	Set xAppLink = Nothing
	Set xItem = Nothing

	'response.write "("&retCount&")"

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

    '// HTML태그 제거 //
    function stripHTML(strng)
    Dim regEx
    Set regEx = New RegExp
    regEx.Pattern = "[<][^>]*[>]"
    regEx.IgnoreCase = True
    regEx.Global = True
    stripHTML = regEx.Replace(strng, " ")
    Set regEx = nothing
    End Function    
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->