<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% session.CodePage = "65001" %>
<% Response.AddHeader "Content-type","text/xml" %>
<% Server.ScriptTimeOut = 1200 %>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
	'// 페이스북 피드 관련 xml 페이지
	Dim query1, vItemid, vItemName, vItemImgUrl, vItemDescription, vBrand, vPrice, vorgprice, vSalechk, vlimitno, vlimityn, vlimitsold, isSoldout, vsellyn, vExpirationDate, xmlPars, rss, Channel
	Dim title, channel_link, description, xItem, vCateFullName, xAppLink
	Dim savePath, rstMsg, strIdx, EndIdx, FileName, newLineNode


	savePath = server.mappath("/Files/fbfeed/") + "\"


	vItemid = request("item_id")
	strIdx = request("strIdx")
	EndIdx = request("EndIdx")

	If (strIdx="" Or EndIdx="") Then
		response.write "<error>잘못된 접근 입니다.</error>"
    	dbCTget.close()	:	response.end
	End If

	If strIdx = 1 Then
		FileName = "FaceBookFeed1.xml"
	End If

	If strIdx = 50000 Then
		FileName = "FaceBookFeed2.xml"
	End If

	If strIdx = 100000 Then
		FileName = "FaceBookFeed3.xml"
	End If

	If strIdx = 150000 Then
		FileName = "FaceBookFeed4.xml"
	End If

	If strIdx = 200000 Then
		FileName = "FaceBookFeed5.xml"
	End If

	If strIdx = 250000 Then
		FileName = "FaceBookFeed6.xml"
	End If

	If strIdx = 300000 Then
		FileName = "FaceBookFeed7.xml"
	End If

	If strIdx = 350000 Then
		FileName = "FaceBookFeed8.xml"
	End If

	If strIdx = 400000 Then
		FileName = "FaceBookFeed9.xml"
	End If

	If strIdx = 450000 Then
		FileName = "FaceBookFeed10.xml"
	End If	

	If strIdx = 500000 Then
		FileName = "FaceBookFeed11.xml"
	End If

	If strIdx = 550000 Then
		FileName = "FaceBookFeed12.xml"
	End If

	If strIdx = 600000 Then
		FileName = "FaceBookFeed13.xml"
	End If

	If strIdx = 650000 Then
		FileName = "FaceBookFeed14.xml"
	End If

	If strIdx = 700000 Then
		FileName = "FaceBookFeed15.xml"
	End If

	If strIdx = 750000 Then
		FileName = "FaceBookFeed16.xml"
	End If

	If strIdx = 800000 Then
		FileName = "FaceBookFeed17.xml"
	End If

	If strIdx = 850000 Then
		FileName = "FaceBookFeed18.xml"
	End If

	If strIdx = 900000 Then
		FileName = "FaceBookFeed19.xml"
	End If

	If strIdx = 950000 Then
		FileName = "FaceBookFeed20.xml"
	End If

	If strIdx >= 1000000 Then
		FileName = "FaceBookFeed21.xml"
	End If								

	If strIdx = 120000 Then
		FileName = "FaceBookFeedEtc.xml"
	End If

	query1 = " Select num, itemid, itemname, icon1image, designercomment, socname, sellcash, orgprice, sailyn, limitno, limityn, "
	query1 = query1 + " limitsold, sellyn, catecode, basicimage600, basicimage, basicimage1000, mainimage, smallimage, sellEndDate, cate1, cate2, cate3, cate4, cate5, "
	query1 = query1 + " itemcouponyn, curritemcouponidx, itemcoupontype, itemcouponvalue "	
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
	query1 = query1 + " 	i.itemcouponyn, i.curritemcouponidx, i.itemcoupontype, i.itemcouponvalue "	
	query1 = query1 + " 	From db_AppWish.dbo.tbl_item i  WITH(NOLOCK) "
	query1 = query1 + " 	inner join db_AppWish.dbo.tbl_item_contents c WITH(NOLOCK) on i.itemid = c.itemid  "
	query1 = query1 + " 	inner join db_AppWish.dbo.tbl_user_c m WITH(NOLOCK) on i.makerid = m.userid "
	query1 = query1 + " 	inner join db_AppWish.[dbo].[tbl_display_cate_item] ci WITH(NOLOCK) on i.itemid = ci.itemid And ci.isdefault = 'y' "
	query1 = query1 + " 	Where i.isusing='Y' And i.itemid <> 0 And i.sellyn = 'Y' And ci.depth>=2  And (c.sellcount>0 or datediff(day, i.regdate, getdate())<=20) "
	query1 = query1 + " 	And i.itemid not in ( Select itemid From db_outmall.dbo.tbl_EpShop_not_in_itemid Where mallgubun='ggshop' AND isusing = 'Y' )"'// 마케팅 요청으로 인한 제외 상품 코드 추가(구글 피드 제외 상품과 동기회 2022.2.8)
	query1 = query1 + " 	And i.makerid not in ( Select makerid From db_outmall.dbo.tbl_EpShop_not_in_makerid Where mallgubun='ggshop' AND isusing = 'N' ) "'// 2021-04-27 14:56 원승현 수정(구글피드 제외 브랜드와 동기화)
	query1 = query1 + " 	And i.adulttype = 0 "		'2019-12-16 성인용품 보내지 않음
	query1 = query1 + " 	And i.dispcate1 <> 119 "	'2021-04-27 마케팅 요청으로 푸드 카테고리 보내지 않음
	query1 = query1 + " 	And ci.catecode not in ('120114101','117103101104','117109105') " '2021-04-27 마케팅 요청으로 해당 카테고리 코드 상품 보내지 않음	
	query1 = query1 + " )Tot "
	If strIdx<>"" Then
		query1 = query1 + " Where num >= "&strIdx&" And num < "&EndIdx&" "
	End If
	dbCTget.CommandTimeOut = 480
	rsCTget.CursorLocation = adUseClient
	rsCTget.Open query1,dbCTget, adOpenForwardOnly, adLockReadOnly

	dim retCount : retCount = rsCTget.recordcount
	If Not(rsCTget.bof Or rsCTget.eof) Then

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
		Channel.childnodes(2).childnodes(0).text = "10x10 Facebook Dynamic Ads Feed"  '설명

		Do Until rsCTget.eof


			Set xItem = xmlPars.CreateElement("item")
			Channel.AppendChild(xItem) 

			xItem.AppendChild(xmlPars.CreateElement("g:id") ) 
			xItem.childnodes(0).text = rsCTget("itemid")  '상품번호

			xItem.AppendChild(xmlPars.CreateElement("g:title"))
			xItem.childnodes(1).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(1).childnodes(0).text = Trim(Replace(Replace(Replace(Replace(rsCTget("itemname"), """", ""), Chr(32), ""), "", ""), Chr(8), ""))  '상품명

			xItem.AppendChild(xmlPars.CreateElement("g:description") ) 
			xItem.childnodes(2).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(2).childnodes(0).text = "생활감성채널 텐바이텐 - " & Trim(Replace(Replace(Replace(Replace(rsCTget("itemname") & " " & Trim(rsCTget("designercomment")),"""",""), Chr(32), ""), "", ""), Chr(8), ""))  '상품설명


			vCateFullName = rsCTget("cate1")
			If Not(rsCTget("cate2")="" Or IsNull(rsCTget("cate2"))) Then
				vCateFullName = vCateFullName & " > "&rsCTget("cate2")
			End If
			If Not(rsCTget("cate3")="" Or IsNull(rsCTget("cate3"))) Then
				vCateFullName = vCateFullName & " > "&rsCTget("cate3")
			End If
			If Not(rsCTget("cate4")="" Or IsNull(rsCTget("cate4"))) Then
				vCateFullName = vCateFullName & " > "&rsCTget("cate4")
			End If
			If Not(rsCTget("cate5")="" Or IsNull(rsCTget("cate5"))) Then
				vCateFullName = vCateFullName & " > "&rsCTget("cate5")
			End If
			If vCateFullName="" Or IsNull(vCateFullName) Then
				vCateFullName = ""
			End If
			xItem.AppendChild(xmlPars.CreateElement("g:product_type") ) 
			xItem.childnodes(3).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(3).childnodes(0).text = vCateFullName  '카테고리

			xItem.AppendChild(xmlPars.CreateElement("g:link") ) 
			xItem.childnodes(4).appendChild(xmlPars.createCDATASection("name_Cdata"))
			'xItem.childnodes(4).childnodes(0).text = "http://www.10x10.co.kr/shopping/category_prd.asp?rdsite=fbec5&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&term=fbdpa_echo&itemid="&rsCTget("itemid")  '링크값(일단은 pc로 보내면 자동으로 모바일로 갈테니 pc주소로 보냄)
			
			'// 웹 url은 브랜치로 연결
			'xItem.childnodes(4).childnodes(0).text = "https://m.10x10.co.kr/common/tenlanding.asp?urltype=item&itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&term=fbdpa_echo&rdsite=fbec5"
			
			'// 웹 url은 브랜치로 연결하되 Bridge Page 사용하지 않고 직링크로 배포
			'xItem.childnodes(4).childnodes(0).text = "https://tenten.app.link/3p?%243p=a_facebook&%24deeplink_no_attribution=true&branch_ad_format=Product&%24ios_url="&Server.URLEncode("http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5")&"&%24android_url="&Server.URLEncode("http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5")&"&%24desktop_url="&Server.URLEncode("http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5")&"&%24deeplink_path="&Server.URLEncode("http://m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5&isbranch=true")&"&~feature=paid+advertising"
			
			'xItem.childnodes(4).childnodes(0).text = "https://tenten.app.link/3p?%243p=a_facebook&branch_ad_format=Product&%24ios_url="&Server.URLEncode("http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5")&"&%24android_url="&Server.URLEncode("http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5")&"&%24desktop_url="&Server.URLEncode("http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5")&"&%24deeplink_path="&Server.URLEncode("m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5&isbranch=true")&"&~feature=paid+advertising&%24uri_redirect_mode=1&%24ios_redirect_timeout=1000&%24ios_passive_deepview=false&%24android_passive_deepview=false&~ad_id={{ad.id}}&~ad_name={{ad.name}}&~ad_set_id={{adset.id}}&~ad_set_name={{adset.name}}&~campaign=dpa&~campaign_id={{campaign.id}}&~keyword="&rsCTget("itemid")
			
			
			'// 브랜치 오류로 인해 일단 해당 feed는 $web_only 속성 추가(브랜치 오류 수정되면 바로 위 주소로 다시 바꿔야됨)
			'xItem.childnodes(4).childnodes(0).text = "https://tenten.app.link/3p?%243p=a_facebook&branch_ad_format=Product&%24ios_url="&Server.URLEncode("http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5")&"&%24android_url="&Server.URLEncode("http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5")&"&%24desktop_url="&Server.URLEncode("http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5")&"&%24deeplink_path="&Server.URLEncode("m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5&isbranch=true")&"&~feature=paid+advertising&%24uri_redirect_mode=1&%24ios_redirect_timeout=1000&%24ios_passive_deepview=false&%24android_passive_deepview=false&%24web_only=true&~ad_id={{ad.id}}&~ad_name={{ad.name}}&~ad_set_id={{adset.id}}&~ad_set_name={{adset.name}}&~campaign=dpa&~campaign_id={{campaign.id}}"

			'// url encode 제거
			'xItem.childnodes(4).childnodes(0).text = "https://tenten.app.link/3p?$3p=a_facebook&branch_ad_format=Product&$ios_url=http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5&$android_url=http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5&$desktop_url=http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5&$deeplink_path=m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=fbec5&isbranch=true&~feature=paid+advertising&$uri_redirect_mode=1&$ios_redirect_timeout=1000&$ios_passive_deepview=false&$android_passive_deepview=false&~ad_id={{ad.id}}&~ad_name={{ad.name}}&~ad_set_id={{adset.id}}&~ad_set_name={{adset.name}}&~campaign=dpa&~campaign_id={{campaign.id}}&~keyword="&rsCTget("itemid")

			'// 마케팅팀에서 요청한 변경되는 url
			'xItem.childnodes(4).childnodes(0).text = "https://tenten.app.link/yxxdjfxbBfb?%243p=a_facebook&~branch_ad_format=Product&$ios_url=http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=adef&$android_url=http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=adef&$desktop_url=http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=adef&$deeplink_path=m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=adef&isbranch=true&~feature=paid%20advertising&$uri_redirect_mode=1&$ios_redirect_timeout=1000&$ios_passive_deepview=false&$android_passive_deepview=false&~ad_id={{ad.id}}&~ad_name={{ad.name}}&~ad_set_id={{adset.id}}&~ad_set_name={{adset.name}}&~campaign=dpa&~campaign_id=6199033136881&~keyword="&rsCTget("itemid")

			'// 마케팅팀에서 요청한 변경되는 url(새로운 버전...)
			xItem.childnodes(4).childnodes(0).text = "https://tenten.app.link/yxxdjfxbBfb?%243p=a_facebook&~branch_ad_format=Product&%24fallback_url="&Server.URLEncode("https://m.10x10.co.kr/category/category_itemprd.asp?itemid="&rsCTget("itemid")&"&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&utm_term=fbdpa_echo&rdsite=adef")&"&%24deeplink_path="&Server.URLEncode("m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid="&rsCTget("itemid"))&"&~campaign=dpa&~campaign_id=6199033136881&~keyword="&rsCTget("itemid")&"&~ad_id={{ad.id}}&~ad_name={{ad.name}}&~ad_set_id={{adset.id}}&~ad_set_name={{adset.name}}"


			If Trim(rsCTget("basicimage600"))="" Or IsNull(rsCTget("basicimage600")) Then
				vItemImgUrl = "http://webimage.10x10.co.kr/image/basic/" + Num2Str(CStr(Clng(rsCTget("itemid")) \ 10000),2,"0","R")  + "/" + rsCTget("basicimage")
			Else
				vItemImgUrl = "http://webimage.10x10.co.kr/image/basic600/" + Num2Str(CStr(Clng(rsCTget("itemid")) \ 10000),2,"0","R")  + "/" + rsCTget("basicimage600")
			End If
			If IsNull(vItemImgUrl) Or vItemImgUrl="" Then
				vItemImgUrl = "http://webimage.10x10.co.kr/image/main/" + Num2Str(CStr(Clng(rsCTget("itemid")) \ 10000),2,"0","R")  + "/" + rsCTget("mainimage")
			End If
			If IsNull(vItemImgUrl) Or vItemImgUrl="" Then
				vItemImgUrl = "http://webimage.10x10.co.kr/image/small/" + Num2Str(CStr(Clng(rsCTget("itemid")) \ 10000),2,"0","R")  + "/" + rsCTget("smallimage")
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
			'xItem.childnodes(8).childnodes(0).text = FormatNumber(rsCTget("sellcash"), 0)&" KRW"  '가격(sellcash값 보내줌)
			xItem.childnodes(8).childnodes(0).text = FormatNumber(rsCTget("orgprice"), 0)&" KRW"  '가격(orgprice값 보내줌) 2021년 4월 7일 마케팅 요청으로 변경

			If rsCTget("sellEndDate")="" Or IsNull(rsCTget("sellEndDate")) Then
				vExpirationDate = "2039-01-18"
			Else
				vExpirationDate = rsCTget("sellEndDate")
			End If

			xItem.AppendChild(xmlPars.CreateElement("g:expiration_date")) 
			xItem.childnodes(9).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(9).childnodes(0).text = vExpirationDate  '만료일(보통은 없지않나??)

			xItem.AppendChild(xmlPars.CreateElement("g:brand")) 
			xItem.childnodes(10).appendChild(xmlPars.createCDATASection("name_Cdata"))
			xItem.childnodes(10).childnodes(0).text = Replace(rsCTget("socname"), Chr(8), "")  '브랜드명

			'// 할인가, 쿠폰할인가 전송(할인기간은 일단 보내지 않음)
			If ((rsCTget("sailyn")="Y") and (rsCTget("orgprice")-rsCTget("sellcash")>0)) or (rsCTget("itemcouponyn") = "Y") And isCouponPriceDisplay(rsCTget("itemid")) Then
				xItem.AppendChild(xmlPars.CreateElement("g:sale_price")) 
				xItem.childnodes(11).appendChild(xmlPars.createCDATASection("name_Cdata"))
				xItem.childnodes(11).childnodes(0).text = FormatNumber(GetCouponAssignPrice(rsCTget("sellcash"), GetCouponDiscountPrice(rsCTget("itemcoupontype"), rsCTget("itemcouponvalue"), rsCTget("sellcash")), rsCTget("itemcouponyn")),0)&" KRW"  '가격(할인값 보내줌)

				'// 할인기간 셋팅
				'xItem.AppendChild(xmlPars.CreateElement("g:sale_price_effective_date")) 
				'xItem.childnodes(12).appendChild(xmlPars.createCDATASection("name_Cdata"))
				'xItem.childnodes(12).childnodes(0).text = "2021-03-29T00:00+09:00/2021-04-25T23:59+09:00"
			End If

			xItem.AppendChild(xmlPars.createTextNode(vbNewLine))			

			'// 딥링크 부분은 사용안함
			'Set xAppLink = xmlPars.CreateElement("applink")
			'xItem.AppendChild(xAppLink)
			'xAppLink.setAttribute "property","ios_url"
			'xAppLink.setAttribute "content","tenwishapp://http://m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?rdsite=fbec5&utm_source=facebook&utm_medium=ad&'utm_campaign=dpa&term=fbdpa_echo&itemid="&rsCTget("itemid")

			'Set xAppLink = xmlPars.CreateElement("applink")
			'xItem.AppendChild(xAppLink)
			'xAppLink.setAttribute "property","ios_app_store_id"
			'xAppLink.setAttribute "content","864817011"

			'Set xAppLink = xmlPars.CreateElement("applink")
			'xItem.AppendChild(xAppLink)
			'xAppLink.setAttribute "property","ios_app_name"
			'xAppLink.setAttribute "content","10x10"

			'Set xAppLink = xmlPars.CreateElement("applink")
			'xItem.AppendChild(xAppLink)
			'xAppLink.setAttribute "property","android_url"
			'xAppLink.setAttribute "content","tenwishapp://http://m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?rdsite=fbec5&utm_source=facebook&utm_medium=ad&utm_campaign=dpa&term=fbdpa_echo&itemid="&rsCTget("itemid")

			'Set xAppLink = xmlPars.CreateElement("applink")
			'xItem.AppendChild(xAppLink)
			'xAppLink.setAttribute "property","android_package"
			'xAppLink.setAttribute "content","kr.tenbyten.shopping"

			'Set xAppLink = xmlPars.CreateElement("applink")
			'xItem.AppendChild(xAppLink)
			'xAppLink.setAttribute "property","android_app_name"
			'xAppLink.setAttribute "content","10x10"

		rsCTget.movenext
		Loop
	End If
	rsCTget.close


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

	'// 쿠폰할인 적용시 쿠폰가격 노출 하지 않는 상품
	Function isCouponPriceDisplay(iid)
		isCouponPriceDisplay = true
		Select Case cStr(iid)
			Case "2624996", "2624995"
				isCouponPriceDisplay = false
		End Select		
	End Function

	'// 쿠폰 할인가 '?
	Function GetCouponDiscountPrice(cptype, cpvalue, sellcash) 
		Select case cptype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(cpvalue*sellcash/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = cpvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

	End function

	'// 쿠폰 적용가
	Function GetCouponAssignPrice(sellcash, coupondiscountprice, itemcouponyn)
		if Trim(itemcouponyn)="Y" then
			GetCouponAssignPrice = sellcash - coupondiscountprice
		else
			GetCouponAssignPrice = sellcash
		end if
	end Function	

	'// XML파일 저장
	if (retCount>0) then
		xmlPars.save(savePath & FileName)
	end if 

'		rstMsg = "데이터 파일 [FaceBookFeed.xml] 생성 완료"
	Set xmlPars = Nothing
	Set xAppLink = Nothing
	Set xItem = Nothing

	response.write "("&retCount&")"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->