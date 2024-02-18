<%
Dim shintvshoppingAPIURL, linkCode, entpCode, entpId, entpPass, shipCostCode, mdCode, entpManSeq, returnManSeq, shipManSeq, makecoCode, originCode, brandCode, mdManId

IF application("Svr_Info") = "Dev" THEN
	shintvshoppingAPIURL = "http://open-api-dev.shinsegaetvshopping.com"
	linkCode		= "TENBY"			'연결코드
	entpCode		= "410000"			'업체코드
	entpId			= "E410000"			'업체사용자ID
	entpPass		= "E410000"			'업체PASSWORD
	entpManSeq		= "002"				'업체담당자
	returnManSeq	= "004"				'회수담당자
	shipManSeq		= "003"				'출고담당자
	shipCostCode	= "B001"			'배송비정책코드 | 5만원이상 3천원
	mdCode 			= "061"				'MD
	mdManId			= "KYMMD"			'담당MD ID
	makecoCode		= "AES2"			'제조업체 | 개발서버에서는 "상세설명참조"
	originCode		= "9999"			'원산지 | 개발서버에서는 "상세설명참조"
	brandCode		= "022652"			'브랜드 | test
Else
	Dim shintvshoppingStrSql
	shintvshoppingStrSql = ""
	shintvshoppingStrSql = shintvshoppingStrSql & " SELECT TOP 1 isnull(iniVal, '') as iniVal "
	shintvshoppingStrSql = shintvshoppingStrSql & " FROM db_etcmall.dbo.tbl_outmall_ini " & VbCRLF
	shintvshoppingStrSql = shintvshoppingStrSql & " where mallid='shintvshopping' " & VbCRLF
	shintvshoppingStrSql = shintvshoppingStrSql & " and inikey='pass'"
	rsget.CursorLocation = adUseClient
	rsget.Open shintvshoppingStrSql, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.Eof then
		entpPass	= rsget("iniVal")
	end if
	rsget.close

	shintvshoppingAPIURL = "http://open-api.shinsegaetvshopping.com"
	linkCode		= "TENBY"			'연결코드
	entpCode		= "419803"			'업체코드
	entpId			= "E419803"			'업체사용자ID
'	entpPass		= "ten101010*"		'업체PASSWORD
	entpManSeq		= "001"				'업체담당자 | 변장혁
	returnManSeq	= "005"				'회수담당자 | 최유미
	shipManSeq		= "004"				'출고담당자 | 최유미
	shipCostCode	= "B01"				'배송비정책코드 | 5만원이상 무료배송
	mdCode 			= "061"				'MD | 061 : 온라인
	mdManId			= "011074"			'담당MD ID
	makecoCode		= "AES2"			'제조업체 | 상세설명참조
	originCode		= "9999"			'원산지 | 상세설명참조
	brandCode		= "031506"			'브랜드
End if
%>