<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% session.CodePage = "65001" %>
<% Server.ScriptTimeOut = 1200 %>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Const MaxPage   = 90		'140 -> 30(15만개) -> 90 (45만개)
Const PageSize = 5000

Dim appPath : appPath = server.mappath("/outmall/googleFeed/") + "\"
Dim FileName: FileName = "googleFeed_temp.xml"
Dim newFileName: newFileName = "googleFeed.xml"
Dim fso, tFile

Function WriteMakeGooleFeedFile(tFile, arrList, byref iLastItemid)
    Dim intLoop, iRow, strSql
    Dim bufstr, isMake
    Dim itemid, deliv, lp, barcode, ArrCateNM
    Dim itemname, designerComment, description, deliveryFixday, adultType, vAddImageArr, vSplitAddImage, vSplitGubun
	Dim item, shipping
	Dim q, tp
    iRow = UBound(arrList,2)

	For intLoop=0 to iRow
		q = ""
		tp = 1
		vAddImageArr	= ""
		vSplitAddImage	= ""
		itemid			= arrList(1,intLoop)
		itemname		= arrList(2,intLoop)
		itemname		= Replace(itemname,"무료배송","")
		itemname		= Replace(itemname,"무료 배송","")
		itemname		= Replace(itemname,"&nbsp;","")
		itemname		= Replace(itemname,"&nbsp","")
		itemname		= Replace(itemname,"""","")

		designerComment	= arrList(3,intLoop)
		deliv 			= arrList(13,intLoop)  ''배송비 /2000, 2500, 0

		If designerComment <> "" Then
			description = "생활감성채널 텐바이텐- "&Replace(Trim(designerComment),"""","")
		Else
			description = "생활감성채널 10x10(텐바이텐)은 디자인소품, 아이디어상품, 독특한 인테리어 및 패션 상품 등으로 고객에게 즐거운 경험을 주는 디자인전문 쇼핑몰 입니다."
		End If

		If isNULL(arrList(9,intLoop)) Then
		    ArrCateNM		= ""
		Else
    		ArrCateNM		= Split(arrList(9,intLoop),"||")(0)
			ArrCateNM		= Replace(ArrCateNM, ",", " &gt; " )
        End If

		adultType = arrList(10,intLoop)
		If (adultType="1" or adultType="2") Then
			adultType = "yes"
		Else
			adultType = "no"
		End If

		barcode = "10" & CHKIIF(itemid >= 1000000, Format00(8, itemid), Format00(6, itemid)) & "0000"
		Set item = xmlPars.CreateElement("item")
			Channel.AppendChild(item)
		'** 기본 제품 데이터 **
			item.AppendChild(xmlPars.CreateElement("g:id"))				'#[ID] 제품의 고유 식별자
			item.childnodes(0).text = itemid

			item.AppendChild(xmlPars.CreateElement("g:title"))			'#[제목] 제품 이름
			item.childnodes(1).appendChild(xmlPars.createCDATASection("g:title_Cdata"))
			item.childnodes(1).childnodes(0).text = itemname

			item.AppendChild(xmlPars.CreateElement("g:description"))	'#[설명] 제품 설명 | 제품을 정확하게 설명하고 방문 페이지의 설명과 일치하게 합니다. '무료 배송'과 같은 프로모션 텍스트, 모두 대문자로 구성된 문구, 변칙적인 외국어 문자를 포함해서는 안 됩니다.
			item.childnodes(2).appendChild(xmlPars.createCDATASection("g:description_Cdata"))
			item.childnodes(2).childnodes(0).text = description

			item.AppendChild(xmlPars.CreateElement("g:link"))			'#[링크] 제품의 방문 페이지
			item.childnodes(3).appendChild(xmlPars.createCDATASection("g:link_Cdata"))
			item.childnodes(3).childnodes(0).text = "http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&utm_source=google&utm_medium=ad&utm_campaign=shopping_w&utm_term=ggshop&rdsite=ggshop"

			item.AppendChild(xmlPars.CreateElement("g:image_link"))		'#[이미지_링크] 제품 기본 이미지의 URL
			item.childnodes(4).appendChild(xmlPars.createCDATASection("g:image_Cdata"))
			item.childnodes(4).childnodes(0).text = "http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) + "/" & arrList(8,intLoop)

			vAddImageArr = arrList(18,intLoop)

			If vAddImageArr = "" Then
				tp = 1
			Else
				vSplitAddImage = Split(vAddImageArr, "|")
				vSplitGubun = ""
				For lp=1 to Ubound(vSplitAddImage) + 1
					vSplitGubun	= Split(vSplitAddImage(lp-1), "^*^*")

					item.AppendChild(xmlPars.CreateElement("g:additional_image_link"))	'[추가_이미지_링크] 제품에 대한 추가 이미지의 URL | 최대10개까지
					item.childnodes(4+lp).appendChild(xmlPars.createCDATASection("g:additional_image_link_cdata"))
					item.childnodes(4+lp).childnodes(0).text = "http://webimage.10x10.co.kr/image/add" & vSplitGubun(0) & "/" & GetImageSubFolderByItemid(itemid) & "/" & vSplitGubun(1)
					tp = tp + 1
				Next
			End IF

			q = 4 + tp

			item.AppendChild(xmlPars.CreateElement("g:mobile_link"))	'[모바일_링크] 모바일과 데스크톱 트래픽에 대한 URL이 다른 경우 모바일에 최적화된 제품 방문 페이지
			item.childnodes(q).appendChild(xmlPars.createCDATASection("g:mobile_link_Cdata"))
			'하단 링크text는 브랜치로 변경될 수 있음..현재는 테스트로 m에 직접 호출..2019-08-26 진영
			item.childnodes(q).childnodes(0).text = "http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&itemid&"&utm_source=google&utm_medium=ad&utm_campaign=shopping_m&utm_term=ggshop&rdsite=ggshop"
		'** 가격 및 재고 **
			item.AppendChild(xmlPars.CreateElement("g:availability"))	'#[재고] 제품 재고 | in stock[재고 있음], out of stock[재고 없음], preorder[선주문]
			item.childnodes(q+1).text = "in stock"

			If arrList(7,intLoop) <> "Y" Then		'할인이 아니면
				item.AppendChild(xmlPars.CreateElement("g:price"))		'#[가격] 제품 가격
				item.childnodes(q+2).text = arrList(5,intLoop)&" KRW"

				item.AppendChild(xmlPars.CreateElement("g:sale_price"))	'[할인가] 제품 할인가
				item.childnodes(q+3).text = ""
			Else
				item.AppendChild(xmlPars.CreateElement("g:price"))		'#[가격] 제품 가격
				item.childnodes(q+2).text = arrList(4,intLoop)&" KRW"

				item.AppendChild(xmlPars.CreateElement("g:sale_price"))	'[할인가] 제품 할인가
				item.childnodes(q+3).text = arrList(6,intLoop)&" KRW"
			End If
		'** 제품 카테고리 **
			item.AppendChild(xmlPars.CreateElement("g:google_product_category"))	'[Google_제품_카테고리] 제품에 대해 Google에서 정의한 제품 카테고리 (일단 다이어리만 할꺼라 사무용품>일반 사무용품>종이 제품에 매칭)
			item.childnodes(q+4).text = "956"

			item.AppendChild(xmlPars.CreateElement("g:product_type"))	'[제품_유형] 제품에 대해 정의한 제품 카테고리
			item.childnodes(q+5).appendChild(xmlPars.createCDATASection("g:product_type_cdata"))
			item.childnodes(q+5).childnodes(0).text  = ArrCateNM
		'** 제품 식별자 **
			item.AppendChild(xmlPars.CreateElement("g:brand"))			'#[브랜드] 모든 새 제품의 경우 필수사항이며 영화, 도서, 음반 브랜드는 제외
			item.childnodes(q+6).appendChild(xmlPars.createCDATASection("g:brand"))
			item.childnodes(q+6).childnodes(0).text  = arrList(12,intLoop)

			item.AppendChild(xmlPars.CreateElement("g:mpn"))			'[MPN] 새 제품에 제조업체에서 할당한 GTIN이 없는 경우만 해당
			item.childnodes(q+7).text = barcode

			item.AppendChild(xmlPars.CreateElement("g:identifier_exists"))	'[식별자_존재] 제품에 상품 고유 식별자(UPI) GTIN, MPN, 브랜드가 있는지 여부를 명시하려면 사용합니다.
			item.childnodes(q+8).text = "yes"
		'** 상세 제품 설명 **
			item.AppendChild(xmlPars.CreateElement("g:condition"))		'#[상태] | new[새 상품] 새로운 상품이나 오리지널 상품, 포장 개봉 전, refurbished[리퍼 상품] ,전문적으로 정상 상태로 복원한 상품, 보증 제공, 원래의 포장일 수도 있고 아닐 수도 있음, used[중고품] 이미 사용된 상품, 원래의 포장이 개봉되었거나 없어진 상태
			item.childnodes(q+9).text = "new"

			item.AppendChild(xmlPars.CreateElement("g:adult"))			'#[성인] 제품에 성인용 콘텐츠가 포함된 경우 | yes[예] no[아니오]
			item.childnodes(q+10).text = adultType
		'** 배송 **
			Set shipping = xmlPars.CreateElement("shipping")			'[배송] 제품의 배송비
				item.AppendChild(shipping)

				shipping.AppendChild(xmlPars.CreateElement("g:country"))	'[국가] ISO 3166 국가 코드
				shipping.childnodes(0).text = "KR"

				shipping.AppendChild(xmlPars.CreateElement("g:price"))		'#[가격] 고정 배송비(필요한 경우 VAT 포함)
				shipping.childnodes(1).text = deliv & " KRW"
			Set shipping = nothing
		Set item = nothing
		iLastItemid = itemid
    Next
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

''작성시간 체크
sqlStr = ""
sqlStr = sqlStr & " INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog (ref) "
sqlStr = sqlStr & " VALUES ('googleFeed_ST')"
dbCTget.execute sqlStr

''데이터 카운트
sqlStr ="[db_outmall].[dbo].[usp_Ten_Google_FeedDataCount]"
dbCTget.CommandTimeout = 120
rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
	FTotCnt = rsCTget(0)
END IF
rsCTget.close
'response.write FTotCnt&"<br>"

Dim i, ArrRows, bufstr1
Dim iLastItemid : iLastItemid=9999999
Dim xmlPars, rss, Channel, title, link, description
If FTotCnt > 0 Then
    FTotPage = CLNG(FTotCnt / PageSize)
    If FTotPage <> (FTotCnt / PageSize) Then FTotPage = FTotPage + 1
    If (FTotPage > MaxPage) Then FTotPage = MaxPage

	Set xmlPars = Server.CreateObject("Msxml2.DOMDocument")
		xmlPars.appendChild(xmlPars.createProcessingInstruction("xml","version=""1.0"" encoding=""UTF-8"""))
		Set rss = xmlPars.CreateElement("rss")
			rss.setAttribute "xmlns:g", "http://base.google.com/ns/1.0"
			rss.setAttribute "version", "2.0"
			xmlPars.AppendChild(rss)
			Set Channel = xmlPars.CreateElement("channel")
				rss.AppendChild(Channel)
				Set title = xmlPars.CreateElement("title")
					Channel.AppendChild(title)
					Channel.childnodes(0).text = "10x10"  '제목
				Set title = nothing
				Set link = xmlPars.CreateElement("link")
					Channel.AppendChild(link)
					Channel.childNodes(1).appendChild(xmlPars.createCDATASection("name_Cdata"))
					Channel.childnodes(1).childnodes(0).text = "http://www.10x10.co.kr"  '주소
				Set title = nothing
				Set description = xmlPars.CreateElement("description")
					Channel.AppendChild(description)
					Channel.childNodes(2).appendChild(xmlPars.createCDATASection("name_Cdata"))
					Channel.childnodes(2).childnodes(0).text = "10x10 Google Feed"  '설명
				Set description = nothing

				For i = 0 to FTotPage - 1
					ArrRows = ""
					sqlStr = "[db_outmall].[dbo].[usp_Ten_Google_FeedData] ("&i+1&", "&PageSize&", "&iLastItemid&")"
					dbCTget.CommandTimeout = 120
					rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
					If Not (rsCTget.EOF OR rsCTget.BOF) Then
						ArrRows = rsCTget.getRows()
					End If
					rsCTget.close

					If isArray(ArrRows) Then
						CALL WriteMakeGooleFeedFile(tFile, ArrRows, iLastItemid)
					End If

					''작성시간 체크
					sqlStr = ""
					sqlStr = sqlStr & " INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog (ref) "
					sqlStr = sqlStr & " VALUES ('googleFeed_"&(i+1)*PageSize&"_"&iLastItemid&"')"
					dbCTget.execute sqlStr
				Next
			Set Channel = nothing
		Set rss = nothing
		xmlPars.save(appPath & FileName)
	Set xmlPars = Nothing
End If

''작성시간 체크
sqlStr = ""
sqlStr = sqlStr & " INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog (ref) "
sqlStr = sqlStr & " VALUES ('googleFeed_ED')"
dbCTget.execute sqlStr

Dim Newfso
Set Newfso = Server.CreateObject("Scripting.FileSystemObject")
	Newfso.CopyFile appPath & FileName ,appPath & newFileName
Set Newfso = nothing

Dim makeCnt
If FTotCnt > (MaxPage * PageSize) Then
	makeCnt = MaxPage * PageSize
Else
	makeCnt = FTotCnt
End If
response.write "Count : " & makeCnt & " make ["&newFileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->