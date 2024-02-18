<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 1200  ''초단위
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Const MaxPage   = 140
Const PageSize = 5000

Dim appPath : appPath = server.mappath("/outmall/googleFeed/") + "\"
Dim FileName: FileName = "googleFeed_temp.xml"
Dim newFileName: newFileName = "googleFeed.xml"
Dim fso, tFile

Function WriteMakeGooleFeedFile(tFile, arrList, byref iLastItemid)
    Dim intLoop, iRow, strSql
    Dim bufstr, isMake
    Dim itemid, deliv, lp, barcode, ArrCateNM
    Dim itemname, designerComment, description, deliveryFixday, adultType
    iRow = UBound(arrList,2)

    For intLoop=0 to iRow
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

		bufstr = "		<item>"
		'** 기본 제품 데이터 **
		bufstr = bufstr & "		<g:id>"&itemid&"</g:id>"						'#[ID] 제품의 고유 식별자
		bufstr = bufstr & "		<g:title><![CDATA["&itemname&"]]></g:title>"	'#[제목] 제품 이름
		bufstr = bufstr & "		<g:description><![CDATA["&description&"]]></g:description>"	'#[설명] 제품 설명 | 제품을 정확하게 설명하고 방문 페이지의 설명과 일치하게 합니다. '무료 배송'과 같은 프로모션 텍스트, 모두 대문자로 구성된 문구, 변칙적인 외국어 문자를 포함해서는 안 됩니다.
		bufstr = bufstr & "		<g:link>http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"</g:link>"	'#[링크] 제품의 방문 페이지
		bufstr = bufstr & "		<g:image_link>http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) + "/" & arrList(5,intLoop)&"</g:image_link>"	'#[이미지_링크] 제품 기본 이미지의 URL

		strSql = ""
		strSql = strSql & " SELECT TOP 30 gubun, ImgType, addimage_400, addimage_600, addimage_1000 "
		strSql = strSql & " FROM [db_AppWish].[dbo].[tbl_item_addimage] "
		strSql = strSql & " WHERE itemid = '"&itemid&"' "
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open strSql, dbCTget, adOpenForwardOnly, adLockReadOnly
		IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
			For lp=1 to rsCTget.RecordCount
				If rsCTget("imgType")="0" Then
					bufstr = bufstr & "		<g:additional_image_link>http://webimage.10x10.co.kr/image/add" & rsCTget("gubun") & "/" & GetImageSubFolderByItemid(itemid) & "/" & rsCTget("addimage_400") &"</g:additional_image_link>"	'[추가_이미지_링크] 제품에 대한 추가 이미지의 URL | 최대10개까지
				End If
				rsCTget.MoveNext
				If lp >= 10 Then Exit For
			Next
		END IF
		rsCTget.close

		bufstr = bufstr & "		<g:mobile_link>http://m.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"</g:mobile_link>"	'[모바일_링크] 모바일과 데스크톱 트래픽에 대한 URL이 다른 경우 모바일에 최적화된 제품 방문 페이지
		'** 가격 및 재고 **
		bufstr = bufstr & "		<g:availability>in stock</g:availability>"										'#[재고] 제품 재고 | in stock[재고 있음], out of stock[재고 없음], preorder[선주문]
'		bufstr = bufstr & "		<g:availability_date>YYYY-MM-DD</g:availability_date>"							'[재고_여부_날짜] 선주문 제품의 배송 가능 날짜 | preorder[선주문]로 availability[재고]를 제출하는 경우에 이 속성을 사용합니다.
'		bufstr = bufstr & "		<g:cost_of_goods_sold>23000.00 KRW</g:cost_of_goods_sold>"						'[매출원가] 특정 상품 판매와 연관된 비용으로, 설정한 회계 기준으로 정의됩니다. 이러한 비용에는 재료, 임금, 화물이나 기타 관리비용이 포함될 수 있습니다. 제품에 매출원가를 제출하면 쇼핑 광고로 발생한 총이익과 매출 규모와 같은 다른 측정항목을 파악할 수 있습니다.
'		bufstr = bufstr & "		<g:expiration_date>YYYY-MM-DD</g:expiration_date>"								'[만료일] 제품 표시가 중단되어야 하는 날짜 | 앞으로 30일 이내의 날짜를 사용합니다.

		If arrList(7,intLoop) <> "Y" Then		'할인이 아니면
			bufstr = bufstr & "		<g:price>"&arrList(5,intLoop)&" KRW</g:price>"								'#[가격] 제품 가격
		Else
			bufstr = bufstr & "		<g:price>"&arrList(4,intLoop)&" KRW</g:price>"								'#[가격] 제품 가격
			bufstr = bufstr & "		<g:sale_price>"&arrList(6,intLoop)&" KRW</g:sale_price>"					'[할인가] 제품 할인가
'			bufstr = bufstr & "		<g:sale_price_effective_date>YYYY-MM-DD</g:sale_price_effective_date>"		'[할인가_적용_일] 제품의 sale_price가 적용되는 기간
		End If
'		bufstr = bufstr & "		<g:unit_pricing_measure>1.5kg</g:unit_pricing_measure>"						'[단가_책정_단위] 제품 판매 시점의 측정치 및 크기 | 무게: oz, lb, mg, g, kg 용량(미국 인치법): floz, pt, qt, gal 용량 단위: ml, cl, l, cbm  길이: in, ft, yd, cm, m 면적: sqft, sqm 단위당: ct
'		bufstr = bufstr & "		<g:unit_pricing_base_measure>100g</g:unit_pricing_base_measure>"			'[단가_책정_기준_단위] 제품의 가격 책정 기준 단위(예: 100ml는 100ml 단위로 가격이 계산됨) | unit_pricing_measure[단가_책정_단위]를 제출하는 경우에 선택사항입니다.
'		bufstr = bufstr & "		<g:installment>"														'[할부] 할부 결제 방식의 세부정보
'		bufstr = bufstr & "			<g:months>6</g:months>"													'[개월] 정수이며 구매자가 결제해야 하는 할부 횟수입니다.
'		bufstr = bufstr & "			<g:amount>50BRL</g:amount>"												'[월납부액] ISO 4217 표준을 따라야 하며 구매자가 매월 결제해야 하는 금액입니다.
'		bufstr = bufstr & "		</g:installment>"
'		bufstr = bufstr & "		<g:subscription_cost>"													'[구독_요금] 무선 제품과 통신 서비스 계약을 번들로 함께 제공하는 월간 또는 연간 요금제의 세부정보
'		bufstr = bufstr & "			<g:period>개월</g:period>"												'#[기간] 단일 구독 단위의 기간으로, 'month[월]' 또는 'year[연]' 단위입니다.
'		bufstr = bufstr & "			<g:period_length>12</g:period_length>"									'#[기간_길이] 구매자가 결제해야 하는 월 또는 연 단위 구독 기간의 길이(정수)입니다.
'		bufstr = bufstr & "			<g:amount>50000 KRW</g:amount>"											'#[월납부액]   ISO 4217 표준을 따라야 하며 구매자가 매월 결제해야 하는 금액입니다. 이 금액을 표시할 때 공간을 덜 차지하도록 가장 가까운 현지 통화 단위로 금액이 반올림될 수 있습니다. 하지만 제공한 값은 사이트에 표시되는 금액과 정확히 일치해야 합니다.
'		bufstr = bufstr & "		</g:subscription_cost>"
'		bufstr = bufstr & "		<g:loyalty_points>" 													'[적립_포인트] (일본만 해당)제품을 구매할 때 고객이 받는 적립 포인트와 유형
'		bufstr = bufstr & "			<g:name>Program A</g:name>"												'#[포인트_값] 제품으로 획득한 포인트
'		bufstr = bufstr & "			<g:points_value>100</g:points_value>"									'[이름] 일본어 12자 또는 로마자 24자로 구성된 적립 포인트 제도의 이름
'		bufstr = bufstr & "			<g:ratio>1.0</g:ratio>"													'[비율] 통화로 전환 시 포인트 비율(숫자)
'		bufstr = bufstr & "		</g:loyalty_points>"
		'** 제품 카테고리 **
		bufstr = bufstr & "		<g:google_product_category>956</g:google_product_category>"					'[Google_제품_카테고리] 제품에 대해 Google에서 정의한 제품 카테고리 (일단 다이어리만 할꺼라 사무용품>일반 사무용품>종이 제품에 매칭)
		bufstr = bufstr & "		<g:product_type><![CDATA["&ArrCateNM&"]]></g:product_type>"					'[제품_유형] 제품에 대해 정의한 제품 카테고리
		'** 제품 식별자 **
		bufstr = bufstr & "		<g:brand><![CDATA["&arrList(12,intLoop)&"]]></g:brand>"						'#[브랜드] 모든 새 제품의 경우 필수사항이며 영화, 도서, 음반 브랜드는 제외
'		bufstr = bufstr & "		<g:gtin>71919219405200</g:gtin>"	'[GTIN] 제조업체가 할당한 GTIN이 있는 모든 새 제품의 경우 | 왠지 itemstock 테이블의 barcode를 써야 될 것 같은데..옵션별로 되있음
		bufstr = bufstr & "		<g:mpn>"&barcode&"</g:mpn>"	'[MPN] 새 제품에 제조업체에서 할당한 GTIN이 없는 경우만 해당
		bufstr = bufstr & "		<g:identifier_exists>yes</g:identifier_exists>"								'[식별자_존재] 제품에 상품 고유 식별자(UPI) GTIN, MPN, 브랜드가 있는지 여부를 명시하려면 사용합니다.
		'** 상세 제품 설명 **
		bufstr = bufstr & "		<g:condition>new</g:condition>"												'#[상태] | new[새 상품] 새로운 상품이나 오리지널 상품, 포장 개봉 전, refurbished[리퍼 상품] ,전문적으로 정상 상태로 복원한 상품, 보증 제공, 원래의 포장일 수도 있고 아닐 수도 있음, used[중고품] 이미 사용된 상품, 원래의 포장이 개봉되었거나 없어진 상태
		bufstr = bufstr & "		<g:adult>"&adultType&"</g:adult>"											'#[성인] 제품에 성인용 콘텐츠가 포함된 경우 | yes[예] no[아니오]
'		bufstr = bufstr & "		<g:multipack>6</g:multipack>"												'[패키지 상품] 판매자가 정의한 패키지 상품에 포함되어 판매되는 동일 제품의 수량
'		bufstr = bufstr & "		<g:is_bundle>no</g:is_bundle>"												'[번들_여부] 1개의 주요 제품과 이를 보조하는 여러 제품으로 판매자가 정의한 맞춤 그룹 제품임을 명시
'		bufstr = bufstr & "		<g:energy_efficiency_class>A+</g:energy_efficiency_class>"					'[에너지_효율_등급] 제품의 에너지 라벨
'		bufstr = bufstr & "		<g:min_energy_efficiency_class>A+++</g:min_energy_efficiency_class>"		'[최소_에너지_효율_등급] 제품의 에너지 라벨
'		bufstr = bufstr & "		<g:max_energy_efficiency_class>D</g:max_energy_efficiency_class>"			'[최대_에너지_효율_등급] 제품의 에너지 라벨
'		bufstr = bufstr & "		<g:age_group>infant</g:age_group>"											'[연령대] 제품의 대상 인구통계 | newborn[신생아] 3개월 이하, infant[영아] 3개월~12개월, toddler[유아] 1세~5세, kids[어린이] 5세~13세, adult[성인] 일반적으로 10대 이상
'		bufstr = bufstr & "		<g:color>Black</g:color>"													'[색상] 제품의 색상
'		bufstr = bufstr & "		<g:gender>unisex</g:gender>"												'[성별] 제품의 대상 성별 | male[남성], female[여성], unisex[남녀공용]
'		bufstr = bufstr & "		<g:material>leather</g:material>"											'[소재] 제품의 원단 또는 소재
'		bufstr = bufstr & "		<g:pattern>striped</g:pattern>"												'[패턴] 제품의 패턴 또는 그래픽 프린트
'		bufstr = bufstr & "		<g:size>XL</g:size>"														'[크기] 제품의 사이즈
'		bufstr = bufstr & "		<g:size_type>regular</g:size_type>"											'[크기_유형] 의류 제품의 컷 | regular[일반], petite[쁘띠], plus[플러스], big and tall[빅 사이즈], maternity[임산부]
'		bufstr = bufstr & "		<g:size_system>US</g:size_system>"											'[사이즈_체계] 제품에 사용되는 사이즈 체계의 국가 | US, UK, EU, DE, FR, JP, CN(중국), IT, BR, MEX, AU
'		bufstr = bufstr & "		<g:item_group_id>AB12345</g:item_group_id>"									'[상품_그룹_ID] 여러 버전(변형)으로 제공되는 제품 그룹의 ID
		'** 쇼핑 캠페인 및 기타 구성 **
'		bufstr = bufstr & "		<g:ads_redirect>http://www.example.com/product.html</g:ads_redirect>"		'[ads_리디렉션] 제품 페이지의 추가 매개변수를 지정하는 데 사용되는 URL입니다. 사용자는 link[링크] 또는 mobile_link[모바일_링크]에 제출된 값이 아니라 이 URL로 이동합니다.
'		bufstr = bufstr & "		<g:custom_label_0>시즌 상품</g:custom_label_0>"								'[맞춤_라벨_0] 쇼핑 캠페인의 입찰 및 보고를 구성하기 위해 제품에 할당하는 라벨입니다. | 이 속성을 여러 번 포함하여 제품당 최대 5개까지 맞춤 라벨을 제출합니다. custom_label_0[맞춤_라벨_0], custom_label_1[맞춤_라벨_1], custom_label_2[맞춤_라벨_2], custom_label_3[맞춤_라벨_3], custom_label_4[맞춤_라벨_4]
'		bufstr = bufstr & "		<g:promotion_id>ABC123</g:promotion_id>"									'[프로모션_ID] 제품을 판매자 프로모션에 연결할 수 있는 식별자입니다.
		'** 목적지 **
'		bufstr = bufstr & "		<g:excluded_destination>Shopping Ads</g:excluded_destination>"				'[제외되는_유형] 특정 유형의 광고 캠페인에 제품이 참여하지 않도록 제외하는 데 사용할 수 있는 설정 | Shopping Ads[쇼핑 광고], Shopping Actions[쇼핑 작업], Display Ads[디스플레이 광고], Surfaces across Google[여러 Google 제품에 게재]
'		bufstr = bufstr & "		<g:included_destination>Shopping Ads</g:included_destination>"				'[포함되는_유형] 특정 유형의 광고 캠페인에 제품을 포함하는 데 사용할 수 있는 설정 | Shopping Ads[쇼핑 광고], Shopping Actions[쇼핑 작업], Display Ads[디스플레이 광고], Surfaces across Google[여러 Google 제품에 게재]
		'** 배송 **
		bufstr = bufstr & "		<g:shipping>"															'[배송] 제품의 배송비
		bufstr = bufstr & "			<g:country>KR</g:country>"												'[국가] ISO 3166 국가 코드
'		bufstr = bufstr & "			<g:region>MA</g:region>"												'[지역] 주, 준주, 현을 제출합니다. 미국, 오스트레일리아, 일본에 지원됩니다. 국가 접두어 없이 ISO 3166-1 국가 코드를 제출합니다(예: CA, NSW, 03).
		bufstr = bufstr & "			<g:service>일반 배송</g:service>"										'[서비스] 서비스 등급 또는 배송 속도
		bufstr = bufstr & "			<g:price>"&deliv&" KRW</g:price>"							'#[가격] 고정 배송비(필요한 경우 VAT 포함)
		bufstr = bufstr & "		</g:shipping>"
'		bufstr = bufstr & "		<g:shipping_label>신선 제품</g:shipping_label>"								'[배송물_라벨] 판매자 센터 계정 설정에서 올바른 배송비를 할당하기 위해 제품에 할당하는 라벨
'		bufstr = bufstr & "		<g:shipping_weight>3kg</g:shipping_weight>"									'[배송물_중량] 배송비를 계산하는 데 사용되는 제품의 중량
'		bufstr = bufstr & "		<g:shipping_length>20cm</g:shipping_length>"								'[배송물_길이] 용적 중량별 배송비를 계산하는 데 사용되는 제품의 길이
'		bufstr = bufstr & "		<g:shipping_width>20cm</g:shipping_width>"									'[배송물_폭] 용적 중량별 배송비를 계산하는 데 사용되는 제품의 폭
'		bufstr = bufstr & "		<g:shipping_height>20cm</g:shipping_height>"								'[배송물_높이] 용적 중량별 배송비를 계산하는 데 사용되는 제품의 높이
'		bufstr = bufstr & "		<g:transit_time_label>시애틀 출고</g:transit_time_label>"					'[운송_시간_라벨] 판매자 센터 계정 설정에 다른 운송 시간을 할당하는 데 도움이 되도록 제품에 할당하는 라벨.
'		bufstr = bufstr & "		<g:max_handling_time>3</g:max_handling_time>"								'[최대_상품_준비_기간] 제품을 주문한 후 배송되기까지 걸리는 최대 시간입니다.
'		bufstr = bufstr & "		<g:min_handling_time>3</g:min_handling_time>"								'[최소_상품_준비_기간] 제품을 주문한 후 배송되기까지 걸리는 최단 시간입니다.
		'** 세금 **
'		bufstr = bufstr & "		<g:tax>"																'#[세금] 미국만해당 | 퍼센트 단위의 제품 판매세율
'		bufstr = bufstr & "			<g:rate>5.00</g:rate>"													'#[세율] 퍼센트 단위의 세율
'		bufstr = bufstr & "			<g:country>US</g:country>"												'[국가] ISO 3166 국가 코드
'		bufstr = bufstr & "			<g:region>MA</g:region>"												'[지역]
'		bufstr = bufstr & "			<g:tax_ship>예</g:tax_ship>"											'[배송세_여부] 배송비에 세금을 부과할지 여부를 지정합니다. 허용되는 값은 yes[예] 또는 no[아니요]입니다.
'		bufstr = bufstr & "		</g:tax>"
'		bufstr = bufstr & "		<g:tax_category>apparel</g:tax_category>"									'[세금_카테고리] 특정 세금 규칙으로 제품을 분류하는 카테고리
		bufstr = bufstr & "</item>"
		tFile.WriteLine bufstr
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

If FTotCnt > 0 Then
    FTotPage = CLNG(FTotCnt / PageSize)
    If FTotPage <> (FTotCnt / PageSize) Then FTotPage = FTotPage + 1
    If (FTotPage > MaxPage) Then FTotPage = MaxPage

    Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(appPath & FileName )

			bufstr1 = ""
			bufstr1 = bufstr1 & "<?xml version=""1.0"" encoding=""UTF-8""?>"
			bufstr1 = bufstr1 & "<rss xmlns:g=""http://base.google.com/ns/1.0"" version=""2.0"">"
			bufstr1 = bufstr1 & "	<channel>"
			bufstr1 = bufstr1 & "		<title>10x10</title>"
			bufstr1 = bufstr1 & "		<link>http://www.10x10.co.kr</link>"
			bufstr1 = bufstr1 & "		<description>10x10 Google Feed</description>"
			tFile.WriteLine bufstr1

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
			bufstr1 = ""
			bufstr1 = bufstr1 & "	</channel>"
			bufstr1 = bufstr1 & "</rss>"
			tFile.WriteLine bufstr1
    		tFile.Close
		Set tFile = Nothing
	Set fso = Nothing
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
response.write FTotCnt&"건 생성 ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->