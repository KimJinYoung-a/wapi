<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% session.CodePage = "65001" %>
<% Server.ScriptTimeOut = 1200 %>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'' 구글 쇼핑 파일 Make / 일별
Const MaxPage   = 300		'140 -> 20(10만개)
Const PageSize = 5000

Dim appPath : appPath = server.mappath("/outmall/googleFeed/") + "\"
Dim FileName: FileName = "googleFeed_temp.txt"
Dim newFileName: newFileName = "googleFeed.txt"
Dim fso, tFile

Function WriteMakeGooleFeedFile(tfso, arrList, byref iLastItemid)
    Dim intLoop, iRow, strSql
    Dim bufstr, isMake
    Dim itemid, deliv, lp, barcode, ArrCateNM, vLink, vMobLink, vImageLink, vAddImageArr, vSplitAddImage, vSplitGubun, additional_image_link, vLastSellcash
    Dim itemname, designerComment, description, deliveryFixday, adultType
	Dim item, shipping
	Dim q, tp
    iRow = UBound(arrList,2)

	For intLoop=0 to iRow
		q = ""
		tp = 1
		vAddImageArr	= ""
		vSplitAddImage	= ""
		additional_image_link = ""
		vLastSellcash	= ""
		itemid			= arrList(1,intLoop)
		itemname		= arrList(2,intLoop)
		itemname		= Replace(itemname,"무료배송","")
		itemname		= Replace(itemname,"무료 배송","")
		itemname		= Replace(itemname,"&nbsp;","")
		itemname		= Replace(itemname,"&nbsp","")
		itemname		= Replace(itemname,"""","")
		itemname 		= Replace(itemname, vbTab, "")
		itemname 		= Replace(itemname, chr(13), "")
		itemname 		= Replace(itemname, chr(10), "")

		vLink			= "http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&utm_source=google&utm_medium=ad&utm_campaign=shopping_w&utm_term=ggshop&rdsite=ggshop"
		vMobLink		= "http://m.10x10.co.kr/category/category_itemprd.asp?itemid="&itemid&"&utm_source=google&utm_medium=ad&utm_campaign=shopping_m&utm_term=ggshop&rdsite=ggshop"
		vImageLink		= "http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) + "/" & arrList(8,intLoop)

		designerComment	= arrList(3,intLoop)
		deliv 			= arrList(13,intLoop)  ''배송비 /2000, 2500, 0

		If designerComment <> "" Then
			description = "생활감성채널 텐바이텐- "&Replace(Trim(designerComment),"""","")
			description = Replace(description, vbTab, "")
		Else
			description = "생활감성채널 10x10(텐바이텐)은 디자인소품, 아이디어상품, 독특한 인테리어 및 패션 상품 등으로 고객에게 즐거운 경험을 주는 디자인전문 쇼핑몰 입니다."
		End If

		If isNULL(arrList(9,intLoop)) Then
		    ArrCateNM		= ""
		Else
    		ArrCateNM		= Split(arrList(9,intLoop),"||")(0)
			ArrCateNM		= Replace(ArrCateNM, ",", " > " )
        End If

		adultType = arrList(10,intLoop)
		If (adultType="1" or adultType="2") Then
			adultType = "yes"
		Else
			adultType = "no"
		End If

		barcode = "10" & CHKIIF(itemid >= 1000000, Format00(8, itemid), Format00(6, itemid)) & "0000"
		bufstr = itemid & vbTab & Replace(itemname, vbTab, "") & vbTab & Replace(description, vbTab, "") & vbTab & vLink & vbTab  		'#[ID] 제품의 고유 식별자 | '#[제목] 제품 이름 | '#[설명] 제품 설명 | 제품을 정확하게 설명하고 방문 페이지의 설명과 일치하게 합니다. '무료 배송'과 같은 프로모션 텍스트, 모두 대문자로 구성된 문구, 변칙적인 외국어 문자를 포함해서는 안 됩니다. | '#[링크] 제품의 방문 페이지
		bufstr = bufstr & vImageLink & vbTab		'#[이미지_링크] 제품 기본 이미지의 URL

		vAddImageArr = arrList(18,intLoop)
		vSplitAddImage = Split(vAddImageArr, "|")
		vSplitGubun = ""
		For lp=1 to Ubound(vSplitAddImage) + 1
			vSplitGubun	= Split(vSplitAddImage(lp-1), "^*^*")
			additional_image_link = additional_image_link & "http://webimage.10x10.co.kr/image/add" & vSplitGubun(0) & "/" & GetImageSubFolderByItemid(itemid) & "/" & vSplitGubun(1) & ","
		Next

		If Right(additional_image_link,1) = "," Then
			additional_image_link = Left(additional_image_link, Len(additional_image_link) - 1)
		End If
		bufstr = bufstr & additional_image_link & vbTab		'[추가_이미지_링크] 제품에 대한 추가 이미지의 URL | 최대10개까지
		bufstr = bufstr & vMobLink & vbTab & "in stock" & vbTab 	'[모바일_링크] 모바일과 데스크톱 트래픽에 대한 URL이 다른 경우 모바일에 최적화된 제품 방문 페이지 | '#[재고] 제품 재고 | in stock[재고 있음], out of stock[재고 없음], preorder[선주문]

		' If arrList(7,intLoop) <> "Y" Then		'할인이 아니면
		' 	bufstr = bufstr & arrList(5,intLoop)&" KRW" & vbTab & "" & vbTab 	'[모바일_링크] 모바일과 데스크톱 트래픽에 대한 URL이 다른 경우 모바일에 최적화된 제품 방문 페이지 | '#[재고] 제품 재고 | in stock[재고 있음], out of stock[재고 없음], preorder[선주문]
		' Else
		' 	bufstr = bufstr & arrList(4,intLoop)&" KRW" & vbTab & arrList(6,intLoop)&" KRW" & vbTab 	'#[가격] 제품 가격 | '[할인가] 제품 할인가
		' End If

		If arrList(19,intLoop) <> arrList(4,intLoop) Then
			vLastSellcash = arrList(19,intLoop)
		End If
		bufstr = bufstr & arrList(4,intLoop)&" KRW" & vbTab & CHKIIF(vLastSellcash <> "", vLastSellcash&" KRW", "") & vbTab 	'#[가격] 제품 가격 | '[할인가] 제품 할인가
		bufstr = bufstr & "956" & vbTab & ArrCateNM & vbTab & Replace(Replace(arrList(12,intLoop),"&nbsp;",""), vbTab, "") & vbTab & barcode & vbTab & "yes" & vbTab & "new" & vbTab & adultType & vbTab  		'[Google_제품_카테고리] 제품에 대해 Google에서 정의한 제품 카테고리 (일단 다이어리만 할꺼라 사무용품>일반 사무용품>종이 제품에 매칭) | '[제품_유형] 제품에 대해 정의한 제품 카테고리 | '#[브랜드] 모든 새 제품의 경우 필수사항이며 영화, 도서, 음반 브랜드는 제외 | '[MPN] 새 제품에 제조업체에서 할당한 GTIN이 없는 경우만 해당 | '[식별자_존재] 제품에 상품 고유 식별자(UPI) GTIN, MPN, 브랜드가 있는지 여부를 명시하려면 사용합니다. | '#[상태] | new[새 상품] 새로운 상품이나 오리지널 상품, 포장 개봉 전, refurbished[리퍼 상품] ,전문적으로 정상 상태로 복원한 상품, 보증 제공, 원래의 포장일 수도 있고 아닐 수도 있음, used[중고품] 이미 사용된 상품, 원래의 포장이 개봉되었거나 없어진 상태 | '#[성인] 제품에 성인용 콘텐츠가 포함된 경우 | yes[예] no[아니오]
		'bufstr = bufstr & "KR::"& deliv &" KRW"	'shipping(country:postal_code:price) -> shipping으로 변경 전 데이터
		bufstr = bufstr & "KR:::"& deliv &" KRW"
		'bufstr = "id222"& vbTab &"title"& vbTab &"price_pc"& vbTab &"price_mobile"& vbTab &"normal_price"& vbTab &"link"& vbTab &"mobile_link"& vbTab &"image_link"& vbTab &"add_image_link"& vbTab &"category_name1"& vbTab &"category_name2"& vbTab &"category_name3"& vbTab &"category_name4"& vbTab &"naver_category"& vbTab &"naver_product_id"& vbTab &"condition"& vbTab &"import_flag"& vbTab &"parallel_import"& vbTab &"order_made"& vbTab &"product_flag"& vbTab &"adult"& vbTab &"goods_type"& vbTab &"barcode"& vbTab &"manufacture_define_number"& vbTab &"model_number"& vbTab &"brand"& vbTab &"maker"& vbTab &"origin"& vbTab &"card_event"& vbTab &"event_words"& vbTab &"coupon"& vbTab &"partner_coupon_download"& vbTab &"interest_free_event"& vbTab &"point"& vbTab &"installation_costs"& vbTab &"pre_match_code"& vbTab &"search_tag"& vbTab &"group_id"& vbTab &"vendor_id"& vbTab &"coordi_id"& vbTab &"minimum_purchase_quantity"& vbTab &"review_count"& vbTab &"shipping"& vbTab &"delivery_grade"& vbTab &"delivery_detail"& vbTab &"attribute"& vbTab &"option_detail"& vbTab &"seller_id"& vbTab &"age_group"& vbTab &"gender"
		tfso.WriteText bufstr, 1
 		iLastItemid = itemid
     Next
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage
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

IF FTotCnt > 0 THEN
    FTotPage = CLNG(FTotCnt/PageSize)
    IF FTotPage<>(FTotCnt/PageSize) THEn FTotPage=FTotPage+1
    IF (FTotPage>MaxPage) THEn FTotPage=MaxPage

    Set fso = CreateObject("ADODB.Stream")
		fso.Mode = 3
		fso.Type = 2
		fso.CharSet = "UTF-8"
		fso.Open
'		bufstr1 = "id"& vbTab &"title"& vbTab &"description"& vbTab &"link"& vbTab &"image_link"& vbTab &"additional_image_link"& vbTab &"mobile_link"& vbTab &"availability"& vbTab &"price"& vbTab &"sale_price"& vbTab &"google_product_category"& vbTab &"product_type"& vbTab &"brand"& vbTab &"mpn"& vbTab &"identifier_exists"& vbTab &"condition"& vbTab &"adult"& vbTab &"shipping(country:postal_code:price)"
		bufstr1 = "id"& vbTab &"title"& vbTab &"description"& vbTab &"link"& vbTab &"image_link"& vbTab &"additional_image_link"& vbTab &"mobile_link"& vbTab &"availability"& vbTab &"price"& vbTab &"sale_price"& vbTab &"google_product_category"& vbTab &"product_type"& vbTab &"brand"& vbTab &"mpn"& vbTab &"identifier_exists"& vbTab &"condition"& vbTab &"adult"& vbTab &"shipping"
		fso.WriteText bufstr1, 1

		For i=0 to FTotPage-1
			ArrRows = ""
			sqlStr = "[db_outmall].[dbo].[usp_Ten_Google_FeedData] ("&i+1&", "&PageSize&", "&iLastItemid&")"
			dbCTget.CommandTimeout = 120 ''2019/01/16 추가
			rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
				ArrRows = rsCTget.getRows()
			END IF
			rsCTget.close

			if isArray(ArrRows) then
				CALL WriteMakeGooleFeedFile(fso, ArrRows, iLastItemid)
			end if

			sqlStr = ""
			sqlStr = sqlStr & " INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog (ref) "
			sqlStr = sqlStr & " VALUES ('googleFeed_"&(i+1)*PageSize&"_"&iLastItemid&"')"
			dbCTget.execute sqlStr
		NExt
		Call fso.SaveToFile(appPath & FileName, 2)
	Set fso = Nothing
END IF

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