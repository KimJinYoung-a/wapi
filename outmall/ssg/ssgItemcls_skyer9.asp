<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "ssg"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST ssgAPIURL = "http://eapi.ssgadm.com"
CONST ssgSSLAPIURL = "https://eapi.ssgadm.com"
CONST ssgApiKey = "18a8d870-12a7-4b36-afaf-1e9d38e2b988"
CONST CDEFALUT_STOCK = 999
CONST SSGMARGIN = 12									'17%는 계약상 최대치..12로 쏘자

Class CSsgItem
	Public Fitemid
	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public Fitemname
	Public FitemDiv
	Public FsmallImage
	Public Fmakerid
	Public Fregdate
	Public FlastUpdate
	Public ForgPrice
	Public ForgSuplyCash
	Public FSellCash
	Public FBuyCash
	Public FsellYn
	Public FsaleYn
	Public FisUsing
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public Fkeywords
	Public Fvatinclude
	Public ForderComment
	Public FoptionCnt
	Public FregedOptCnt
	Public FaccFailCNT
	Public FlastErrStr
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Ficon2Image
	Public FListimage
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public FSafetyNum
	Public Fitemcontent
	Public FSsgStatCD
	Public Fdeliverfixday
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FmaySoldOut
	Public FSsgGoodno
	Public FDisplayDate
	Public Fregitemname
	Public FregImageName
	Public FbasicImageNm
	Public Fsocname_kor
	Public FDepthCode
	Public FDepth4Code
	Public Fcdmkey
	Public Fcddkey
	Public FG9GoodNo
	Public FMapCnt
	Public FMwdiv
	Public FItemsize
	Public FItemsource

	Public FNotinCate
	Public FSafeAuthType
	Public FAuthItemTypeCode
	Public FIsChildrenCate
	Public FOverlap

	Public Function getLimitEa()
		Dim ret
		If FLimitYn = "Y" Then
			ret = FLimitNo - FLimitSold - 5
			If ret > 1000 Then
				ret = CDEFALUT_STOCK
			End If
		Else
			ret = CDEFALUT_STOCK
		End If

		If (ret < 1) Then ret = 0
		getLimitEa = ret
	End Function

	Function RightCommaDel(ostr)
		Dim restr
		restr = ""
		If IsNULL(ostr) Then Exit Function
		restr = Trim(ostr)
		If (Right(restr,1)=",") Then restr = Left(restr,Len(restr)-1)
		RightCommaDel = restr
	End Function

	'// 품절여부
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold <= CMAXLIMITSELL))
	end function

	Public Function MustPrice()
		Dim GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
	End Function

	Public Function getFiftyUpDown()
		Dim strSql, zoptaddprice, tmpPrice
		If FOptionCnt = 0 Then
			getFiftyUpDown = "N"
		Else
			strSql = ""
			strSql = strSql &" SELECT Max(optaddprice) optaddprice "
			strSql = strSql &" FROM db_item.dbo.tbl_item_option "
			strSql = strSql &" WHERE itemid = '"&FItemid&"' "
			rsget.Open strSql,dbget,1
			If Not(rsget.EOF or rsget.BOF) Then
				zoptaddprice = rsget("optaddprice")
			End If
			rsget.Close

			If zoptaddprice = 0 Then
				getFiftyUpDown = "N"
			Else
				tmpPrice = Clng(MustPrice / 2)
				If tmpPrice > zoptaddprice Then
					getFiftyUpDown = "N"
				Else
					getFiftyUpDown = "Y"
				End If
			End If
		End If
	End Function

	'// SSG 판매여부 반환
	Public Function getSsgSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getSsgSellYn = "Y"
			Else
				getSsgSellYn = "N"
			End If
		Else
			getSsgSellYn = "N"
		End If
	End Function

    Public Function getItemNameFormat()
		Dim buf
		If application("Svr_Info") = "Dev" Then
			FItemName = "TEST상품 "&FItemName
		End If

		buf = replace(FItemName,"'","")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","프로")
		buf = replace(buf,"[무료배송]","")
		buf = replace(buf,"[무료 배송]","")

		'2017-07-03 김진영 상품명에 특문 제거
		buf = replace(buf,"ː","")
		buf = replace(buf,"?","")
		buf = replace(buf,"★","")
		buf = replace(buf,"™","")
		buf = replace(buf,"π","")
		buf = replace(buf,"№","")
		buf = replace(buf,"♥"," ")
		buf = replace(buf,"×","x")
		buf = replace(buf,"：",":")
		buf = replace(buf,"º","")
		buf = replace(buf,"’","'")
		buf = replace(buf,"`","")
		buf = replace(buf,"，",",")
		buf = replace(buf,"［","[")
		buf = replace(buf,"］","]")
		'2017-07-03 김진영 상품명에 특문 제거끝
		getItemNameFormat = buf
    End Function

	Public Function checkTenItemOptionValid()
		Dim strSql, chkRst, chkMultiOpt
		Dim cntType, cntOpt
		chkRst = true
		chkMultiOpt = false

		If FoptionCnt > 0 Then
			'// 이중옵션확인
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				cntType = rsget.RecordCount
			End If
			rsget.Close

			If chkMultiOpt Then
				'// 이중옵션 일때
				strSql = "Select optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
						cntOpt = ubound(split(db2Html(rsget("optionname")), ",")) + 1
						If cntType <> cntOpt then
							chkRst = false
						End If
						rsget.MoveNext
					Loop
				Else
					chkRst = false
				End If
				rsget.Close
			Else
				'// 단일옵션일 때
				strSql = "Select optionTypeName, optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1
				If (rsget.EOF or rsget.BOF) Then
					chkRst = false
				End If
				rsget.Close
			End If
		End If
		'//결과 반환
		checkTenItemOptionValid = chkRst
	End Function

	Public Function getSourcearea()
		Dim arrAreaName, arrAreaCode, i
		arrAreaName = Array("세르비아", "상품상세참조", "원양칠레", "북서부대서양", "중북부대서양", "중서부대서양", "중동부대서양", "서남부대서양", "동남부대서양", "남극부대서양", "북대서양", "원양노르웨이", "원양브라질", "원양아르헨티나", "포클랜드", "서인도양", "국내산", "중국", "일본", "미국산", "베트남", "마카오", "북태평양산", "북서부태평양", "중북부태평양", "중서부태평양", "중동부태평양", "서남부태평양", "이탈리아", "벨기에", "북한", "프랑스", "독일", "대만", "가나", "가봉", "가이아나", "감비아", "과테말라", "그레나다", "그루지야", "동남부태평양", "남극부태평양", "베네수엘라", "벨로루시", "벨리즈", "보츠와나", "볼리비아", "부룬디", "부르키나파소", "부탄", "불가리아", "브라질", "브루나이", "사우디아라비아", "요르단", "우간다", "우루과이", "우즈베키스탄", "우크라이나", "산마리노", "상투메프린시페", "서사모아", "세네갈", "세이셸", "세인트루시아", "세인트빈센트그레나딘", "수입산", "국내산", "상세설명참조", "그리스", "몰타", "멕시코", "모나코", "기니-비사우", "나미비아", "나우루", "나이지리아", "남아프리카공화국", "네덜란드", "네팔", "노르웨이", "뉴질랜드", "니제르", "니카라과", "대한민국", "덴마크", "도미니카", "동티모르", "라스팔마스", "라오스", "라이베리아", "라트비아", "러시아", "레바논", "레소토", "루마니아", "룩셈부르크", "르완다", "리히텐슈타인", "마다가스카르", "마샬", "마이크로네시아 연방", "마케도니아", "말라위", "기니아", "리투아니아", "리비아", "말레이시아", "말리", "모로코", "모리셔스", "모리타니", "모잠비크", "몰도바", "몰디브", "몽골", "미얀마", "바누아투", "바베이도스", "바티칸", "바하마", "방글라데시", "베냉", "소말리아", "솔로몬", "수단", "스리랑카", "스와질란드", "스웨덴", "스위스", "스페인", "슬로바키아", "슬로베니아", "시리아", "시에라리온", "아랍에미리트", "아르메니아", "아르헨티나", "아이슬랜드", "아이티", "아일랜드", "아제르바이잔", "아프가니스탄", "안도라", "알제리", "앙골라", "앤티가바부다", "에리트레아", "에스토니아", "에콰도르", "엘살바도르", "예멘", "오만", "오스트리아", "온두라스", "필리핀", "헝가리", "카자흐스탄", "카타르", "캐나다", "케냐", "코모로", "코스타리카", "코트디브와르", "콜롬비아", "콩고", "콩고민주공화국", "쿠바", "쿠웨이트", "크로아티아", "키르기스탄", "키리바시", "키프로스", "타지키스탄", "탄자니아", "태국", "토고", "통가", "투르크메니스탄", "튀니지", "트리니드다드", "파나마", "파라과이", "파키스탄", "파푸아뉴기니아", "팔라우", "팔레스타인", "페루", "포르투갈", "폴란드", "푸에르토리코", "호주", "홍콩", "유고슬라비아", "이디오피아", "이라크", "이란", "이스라엘", "이집트", "인도", "자메이카", "잠비아", "적도기니", "중앙아프리카공화국", "지부티", "짐바브웨", "차드", "체코", "칠레", "카메룬", "카보베르데", "원양뉴질랜드", "원양러시아", "원양미국", "원양멕시코", "북해도", "알라스카", "원양인도네시아", "수리남", "영국", "터키", "인도네시아", "바레인", "알바니아", "캄보디아", "핀란드", "싱가폴", "투발루", "피지", "원양중국", "원양산", "동인도양", "남극부인도양", "원양파키스탄", "국내", "대서양", "해외기타", "아마존닷컴", "기타", "매핑없음", "브랜드 네트웍스", "기타", "동아", "오스트리아", "아시아", "싱가포르", "슬로바키아", "스웨덴", "중국", "북한", "북태평양", "미국", "라트비아", "덴마크", "터어키산", "캄보디아", "칠레", "우크라이나", "아르헨티나", "이탈리아", "스페인", "북아메리카", "독일", "뉴질랜드", "노르웨이", "핀란드", "포르투갈", "터키", "콜롬비아", "엘살바도르", "아일랜드", "영국", "불가리아", "벨기에", "미얀마", "모로코", "멕시코", "남아프리카공화국", "튀니지", "태국", "이집트", "온두라스", "슬로베니아", "마카오", "도미니카", "네덜란드", "피지", "크로아티아", "인도", "모나코", "마다가스카르", "루마니아", "라오스", "그리스", "프랑스", "페루", "일본", "이란", "시리아", "호주", "헝가리", "한국", "브라질", "베트남", "러시아", "폴란드", "파키스탄", "이스라엘", "스리랑카", "스위스", "동남아", "과테말라", "필리핀", "캐나다", "체코", "유럽", "조지아(Georgia)", "GERMANY", "INDONESIA (R&D GERMANY)", "CHINA", "JAPAN", "Germany(독일)", "Italy(이탈리아)", "France(프랑스)", "미국/영국", "영국/프랑스", "영국/미국", "CHINA OEM", "중국 OEM", "이탈리아시아크", "대한민국성과향상센터", "ENGLAND", "ITALY", "IYALY", "호주", "MALAYSIA(R&D GERMANY)", "대한민국 (R&D GERMANY)", "KOREA", "CHINA(R&D FRANCE)", "CHINA (R&D FRANCE)", "보스니아", "알바니아", "르샤트라", "마케도니아", "fusha", "네팔", "바레인", "니카라과", "에스토니아", "리투아니아", "이태리", "몰타", "타이완", "몽골", "리샤", "원양산", "베키아에누보", "황태촌", "홍콩", "국내산", "모리셔스공화국", "코스타리카", "아르메니아", "몰도바", "파푸아뉴기니", "케냐", "몰디브", "세르비아", "에티오피아", "스코트랜드", "인도네시아", "요르단", "방글라데시", "말레이시아", "대만")
		arrAreaCode = Array("1000000235", "2000000033", "1000000217", "1000000218", "1000000219", "1000000220", "1000000221", "1000000222", "1000000223", "1000000224", "1000000225", "1000000226", "1000000227", "1000000228", "1000000229", "1000000230", "1000000001", "1000000002", "1000000003", "1000000004", "1000000005", "1000000199", "1000000201", "1000000202", "1000000203", "1000000204", "1000000205", "1000000206", "1000000006", "1000000007", "1000000008", "1000000009", "1000000010", "1000000011", "1000000012", "1000000013", "1000000014", "1000000015", "1000000016", "1000000017", "1000000018", "1000000207", "1000000208", "1000000075", "1000000076", "1000000077", "1000000079", "1000000080", "1000000081", "1000000082", "1000000083", "1000000084", "1000000085", "1000000086", "1000000087", "1000000132", "1000000133", "1000000134", "1000000135", "1000000136", "1000000088", "1000000089", "1000000090", "1000000091", "1000000092", "1000000093", "1000000094", "1000000999", "1000000672", "1000000000", "1000000019", "1000000056", "1000000057", "1000000058", "1000000021", "1000000022", "1000000023", "1000000024", "1000000025", "1000000026", "1000000027", "1000000028", "1000000029", "1000000030", "1000000031", "1000000032", "1000000033", "1000000034", "1000000035", "1000000036", "1000000037", "1000000038", "1000000039", "1000000040", "1000000041", "1000000042", "1000000043", "1000000044", "1000000045", "1000000048", "1000000049", "1000000050", "1000000051", "1000000052", "1000000053", "1000000020", "1000000047", "1000000046", "1000000054", "1000000055", "1000000059", "1000000060", "1000000061", "1000000062", "1000000063", "1000000064", "1000000066", "1000000067", "1000000068", "1000000070", "1000000071", "1000000072", "1000000073", "1000000074", "1000000096", "1000000097", "1000000098", "1000000100", "1000000101", "1000000102", "1000000103", "1000000104", "1000000105", "1000000106", "1000000107", "1000000108", "1000000110", "1000000111", "1000000112", "1000000113", "1000000114", "1000000115", "1000000116", "1000000117", "1000000118", "1000000120", "1000000121", "1000000122", "1000000123", "1000000124", "1000000125", "1000000126", "1000000128", "1000000129", "1000000130", "1000000131", "1000000195", "1000000196", "1000000156", "1000000157", "1000000159", "1000000160", "1000000161", "1000000162", "1000000163", "1000000164", "1000000165", "1000000166", "1000000167", "1000000168", "1000000169", "1000000170", "1000000171", "1000000172", "1000000173", "1000000174", "1000000175", "1000000177", "1000000178", "1000000179", "1000000181", "1000000182", "1000000183", "1000000184", "1000000185", "1000000186", "1000000187", "1000000188", "1000000189", "1000000190", "1000000191", "1000000192", "1000000197", "1000000198", "1000000137", "1000000138", "1000000139", "1000000140", "1000000141", "1000000142", "1000000143", "1000000145", "1000000146", "1000000147", "1000000148", "1000000149", "1000000150", "1000000151", "1000000152", "1000000153", "1000000154", "1000000155", "1000000209", "1000000210", "1000000211", "1000000212", "1000000213", "1000000214", "1000000215", "1000000099", "1000000127", "1000000176", "1000000144", "1000000069", "1000000119", "1000000158", "1000000194", "1000000109", "1000000180", "1000000193", "1000000216", "1000000234", "1000000231", "1000000232", "1000000233", "2001000013", "2001000021", "2000000079", "2000000043", "2000000003", "2000999999", "2001000024", "1000000990", "1000000259", "2000000048", "2000000044", "2000000041", "2000000038", "2000000035", "2000000060", "2000000030", "2000000029", "2000000023", "2000000014", "2000000009", "1000000200", "2000000063", "2000000062", "2000000051", "2000000042", "2000000056", "2000000037", "2000000028", "2000000011", "2000000007", "2000000006", "2000000076", "2000000072", "2000000068", "2000000065", "2000000046", "2000000045", "2000000047", "2000000031", "2000000027", "2000000024", "2000000022", "2000000020", "2000000004", "2000000069", "2000000067", "2000000055", "2000000049", "2000000039", "2000000018", "2000000010", "2000000005", "2000000075", "2000000066", "2000000057", "2000000021", "2000000017", "2000000016", "2000000013", "2000000002", "2000000074", "2000000071", "2000000059", "2000000053", "2000000040", "2000000081", "2000000080", "2000000078", "2000000032", "2000000026", "2000000015", "2000000073", "2000000070", "2000000054", "2000000034", "2000000036", "2000000012", "2000000001", "2000000077", "2000000064", "2000000061", "2000000052", "3000000001", "1000000236", "1000000237", "1000000238", "1000000239", "1000000240", "1000000241", "1000000242", "1000000243", "1000000244", "1000000245", "1000000246", "1000000247", "1000000248", "1000000249", "1000000250", "1000000251", "1000000252", "1000000253", "1000000254", "1000000255", "1000000256", "1000000257", "1000000258", "2001000028", "2001000029", "2001000023", "2001000018", "2001000004", "2001000001", "2001000009", "2001000010", "2001000020", "2001000032", "2001000033", "2001000031", "2001000030", "2001000003", "2001000027", "2001000026", "2001000025", "2001000022", "2000000082", "2001000014", "2001000002", "2001000007", "2001000016", "2001000015", "2001000008", "2001000006", "2001000012", "2001000011", "2001000005", "2001000019", "2000000058", "2000000050", "2000000025", "2000000019", "2000000008")

		If FSourcearea = "한국" Then
			getSourcearea = "대한민국"
		End If

		For i =0 To Ubound(arrAreaName)
			If Trim(arrAreaName(i)) = Trim(FSourcearea) Then
				getSourcearea = Trim(arrAreaCode(i))
				Exit For
			End If
		Next

		If getSourcearea = "" Then
			getSourcearea = "1000000000"		''상세설명참조
		End If
	End Function

	Public Function getShopLeadTime()
		Dim CateLargeMid, leadTime
		CateLargeMid = CStr(FtenCateLarge) & CStr(FtenCateMid)
		Select Case CateLargeMid
			Case "030331", "055070", "055080"
				leadTime = 15
			Case "040010", "040020", "040030", "040040", "040050", "040070", "040080", "040090", "040100", "055100", "055110", "055120"
				leadTime = 10
			Case "050045"
				leadTime = 7
			Case "040011", "040121", "045002", "045003", "050010", "050020", "050030", "050040", "055090", "055222"
				leadTime = 5
			Case Else
				leadTime = 3
		End Select
		getShopLeadTime = leadTime
	End Function

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	Public Function getSsgContParamToReg()
		Dim strRst, strSQL
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '돋움','돋움' }</style><br>"
		strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_ssg.jpg'></p><br>"

		If Fitemsource <> "" Then
			strRst = strRst & "- 재료 : " &  Fitemsource & "<br>"
		End If
		strRst = strRst & Replace(Replace(FItemContent,"",""),"","")

		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
				If rsget("imgType") = "1" Then
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#기본 상품 설명이미지
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=http://webimage.10x10.co.kr/image/main/" & GetImageSubFolderByItemid(FItemID) & "/" & Fmainimage & "><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=http://webimage.10x10.co.kr/image/main2/" & GetImageSubFolderByItemid(FItemID) & "/" & Fmainimage2 & "><br>")

		'#배송 주의사항
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ssg.jpg"">")
		getSsgContParamToReg = strRst
	End Function

    public function isImageChanged()
        Dim ibuf : ibuf = getBasicImage
        if InStr(ibuf,"-")<1 then
            isImageChanged = FALSE
            Exit function
        end if
        isImageChanged = ibuf <> FregImageName
    end function

    public function getBasicImage()
        if IsNULL(FbasicImageNm) or (FbasicImageNm="") then Exit function
        getBasicImage = FbasicImageNm
    end function

	Public Function getCertInfoParam(iCode)
		Dim strRst, strSql, isChild, isSafe, isElec, isHarm
		Dim chldCertYn, chldCertDivCd, chldCertNo
		Dim certKind, certYn, certDivCd, certNo
		strSql = ""
		strSql = strSql & " SELECT TOP 1 chldCertTgtYn, safeCertTgtYn, elecCertTgtYn, harmCertTgtYn "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_ssg_mmg_category] "
		strSql = strSql & " WHERE stdCtgDclsId = '"&iCode&"' "
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			isChild	= rsget("chldCertTgtYn")
			isSafe	= rsget("safeCertTgtYn")
			isElec	= rsget("elecCertTgtYn")
			isHarm	= rsget("harmCertTgtYn")
		End If
		rsget.Close

		If (FSafetyyn = "Y") And (FsafetyDiv = "50") Then
			chldCertYn		= "Y"
			chldCertDivCd	= "10"
			chldCertNo		= FSafetyNum
		Else
			chldCertYn		= "N"
		End If

		If isChild = "Y" Then certKind = "6000000001"
		If isSafe = "Y" Then certKind = "6000000002"
		If isElec = "Y" Then certKind = "6000000003"
		If isHarm = "Y" Then certKind = "6000000004"

		If (FSafetyyn = "Y") AND (FSafetyNum <> "")  Then
			certYn = "Y"
			certDivCd = "10"
			certNo = FSafetyNum
		Else
			certYn = "N"
		End If

		strRst = ""
		strRst = strRst & "	<chldCert>"
		strRst = strRst & "		<chldCertYn>"&chldCertYn&"</chldCertYn>" 									'#어린이인증 여부
		strRst = strRst & "		<chldCertDivCd>"&chldCertDivCd&"</chldCertDivCd>"  							'어린이인증 구분 (commCd:I368) | (어린이인증 여부가 Y일 경우에만 필수) 10 : 안전인증대상, 20 : 안전확인대상, 30 : 공급자적합성확인
		strRst = strRst & "		<chldCertNo>"&chldCertNo&"</chldCertNo>" 									'인증번호 | 어린이인증 구분이 10, 20 일경우에만 필수-
		strRst = strRst & "	</chldCert>"
		If certKind <> "" Then
			strRst = strRst & "	<certInfos>"
			strRst = strRst & "		<certInfo>"
	'뭐지..아래 값에 certKin -> 6000000004 이걸로 강제로 넣고 등록하니 등록되네? 상품코드 : 366690 .. 2017-12-20 19:52 김진영
			strRst = strRst & "			<certKind>"&certKind&"</certKind>"										'#인증종류 (commCd:I387) | 인증대상 카테고리 일 경우 필수..6000000001 : 어린이인증 대상여부, 6000000002 : 안전인증 대상여부, 6000000003 : 전파인증 적합성평가 대상여부, 6000000004 : 위해우려제품 표시대상여부
			strRst = strRst & "			<certYn>"&certYn&"</certYn>"											'#인증 여부
			strRst = strRst & "			<certDivCd>"&certDivCd&"</certDivCd>"									'인증 구분 (commCd:I368) | 인증여부가 Y이고 인증종류가 (certKind=6000000001 | 6000000002) 일 경우 필수..10 : 안전인증대상, 20 : 안전확인대상, 30 : 공급자적합성확인
			strRst = strRst & "			<certNo>"&certNo&"</certNo>"											'인증번호 | 인증 구분이 10, 20 일경우에만 필수-
			strRst = strRst & "		</certInfo>"
			strRst = strRst & "	</certInfos>"
		End If
		getCertInfoParam = strRst
'response.write strRst
'response.end
	End Function

	Public Function getSsgAddImageParam()
		Dim strRst, strSQL, i
		strRst = ""
		strRst = strRst & "	<itemImgs>"
		strRst = strRst & "		<imgInfo>"
		strRst = strRst & "			<dataSeq>1</dataSeq>"													'#자료순번
		strRst = strRst & "			<dataFileNm><![CDATA["&FbasicImage&"]]></dataFileNm>"					'#자료파일명
		strRst = strRst & "			<rplcTextNm>대표이미지</rplcTextNm>"												'#대체 텍스트 명
		strRst = strRst & "		</imgInfo>"

		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=2 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					strRst = strRst & "		<imgInfo>"
					strRst = strRst & "			<dataSeq>"&i&"</dataSeq>"													'#자료순번
					strRst = strRst & "			<dataFileNm><![CDATA[http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & "]]></dataFileNm>"					'#자료파일명
					strRst = strRst & "			<rplcTextNm>상품 이미지"&i&"</rplcTextNm>"												'#대체 텍스트 명
					strRst = strRst & "		</imgInfo>"
				End If
				rsget.MoveNext
				If i>=9 Then Exit For
			Next
		End If
		rsget.Close
		strRst = strRst & "	</itemImgs>"
'		strRst = strRst & "	<qualityViewImgs>"
'		strRst = strRst & "		<imgInfo>"
'		strRst = strRst & "			<dataSeq></dataSeq>"													'#자료순번
'		strRst = strRst & "			<dataFileNm></dataFileNm>"												'#자료파일명
'		strRst = strRst & "			<rplcTextNm></rplcTextNm>"												'#대체 텍스트 명
'		strRst = strRst & "		</imgInfo>"
'		strRst = strRst & "	</qualityViewImgs>"
		getSsgAddImageParam = strRst
	End Function

	Public Function getRegedOptionCnt
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as Cnt  "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption "
		sqlStr = sqlStr & " WHERE mallid= 'ssg' "
		sqlStr = sqlStr & " and itemoption <> '0000' "
		sqlStr = sqlStr & " and itemid=" & FItemid
		rsget.Open sqlStr,dbget,1
			getRegedOptionCnt = rsget("Cnt")
		rsget.Close
	End Function

	Public Function IsMayLimitSoldout
		If FOptionCnt = 0 Then
			Exit Function
		End If
		Dim sqlStr, optLimit, limitYCnt
		sqlStr = ""
		sqlStr = sqlStr & " SELECT itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item_option "
		sqlStr = sqlStr & " WHERE isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
		rsget.Open sqlStr,dbget,1
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
				optLimit = rsget("optLimit")
				optLimit = optLimit-5
				If (optLimit < 1) Then optLimit = 0
				If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

				If (optLimit <> 0) Then
					limitYCnt =  limitYCnt + 1
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		If limitYCnt = 0 Then
			IsMayLimitSoldout = "Y"
		Else
			IsMayLimitSoldout = "N"
		End If
	End Function

	Function getiszeroWonSoldOut(iitemid)
		Dim sqlStr, i, goptlimitno, goptlimitsold, cnt
		i = 0
		If Flimityn = "Y" Then
			sqlStr = ""
			sqlStr = sqlStr & "SELECT Count(*) as cnt FROM db_item.dbo.tbl_item_option where itemid = '"&iitemid&"' and optaddprice > 0 "
			rsget.Open sqlStr,dbget,1
				cnt = rsget("cnt")
			rsget.Close

			If cnt = 0 Then
				getiszeroWonSoldOut = "N"
			Else
				sqlStr = ""
				sqlStr = sqlStr & " SELECT itemid, itemoption, optlimitno, optlimitsold "
				sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option  "
				sqlStr = sqlStr & " where itemid = '"&iitemid&"'  "
				sqlStr = sqlStr & " and optaddprice = 0 "
				sqlStr = sqlStr & " and isusing = 'Y' "
				sqlStr = sqlStr & " and optsellyn = 'Y' "
				rsget.Open sqlStr,dbget,1
				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
						goptlimitno		= rsget("optlimitno")
						goptlimitsold	= rsget("optlimitsold")
						If goptlimitno - goptlimitsold > CMAXLIMITSELL Then
							i = i + 1
						End If
						rsget.MoveNext
					Loop

					If i = 0 Then		'0원 옵션의 재고가 5개 이하면 품절
						getiszeroWonSoldOut = "Y"
					Else
						getiszeroWonSoldOut = "N"
					End If
				Else
					getiszeroWonSoldOut = "Y"
				End If
				rsget.Close
			End If
		Else
			getiszeroWonSoldOut = "N"
		End If
	End Function

	Public Function getSsgCategoryParam()
		Dim sqlStr, i, standardCode, arrDepthCode, arrSiteNo
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 2 stdCtgDClsCd, depthCode, siteNo "
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_ssg_cate_mapping "
		sqlStr = sqlStr & " WHERE tenCateLarge = '"& FtenCateLarge &"' "
		sqlStr = sqlStr & " and tenCateMid = '"& FtenCateMid &"' "
		sqlStr = sqlStr & " and tenCateSmall = '"& FtenCateSmall &"' "
		sqlStr = sqlStr & " ORDER BY siteNo DESC "
		rsget.Open sqlStr, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				standardCode		= rsget("stdCtgDClsCd")
				arrDepthCode		= arrDepthCode & rsget("depthCode") & ","
				arrSiteNo			= arrSiteNo & rsget("siteNo") & ","
				rsget.MoveNext
			Next
			arrDepthCode = RightCommaDel(arrDepthCode)
			arrSiteNo = RightCommaDel(arrSiteNo)
		End If
		rsget.Close
		getSsgCategoryParam = standardCode & "|_|" & arrDepthCode & "|_|" & arrSiteNo
	End Function

	Public Function getSsgOptParamtoEDIT()
		Dim strRst, strRst2, strRst3, strSql, chkMultiOpt, requireDetailStr, i
		Dim itemoption, outmalloptcode, outmalloptName, optlimityn, isusing, optsellyn, opt1name, opt2name, opt3name, preged, optNameDiff, oopt, optaddprice
		Dim itemSellTypeCd, OptTypeNm1, OptTypeNm2, OptTypeNm3, optLimit, arrOptionname
		Dim arrRows, isOptionExists, sellStatCd
		Dim arrOptTypeNm

		If FOptionCnt = 0 Then			'단품
			itemSellTypeCd = "10"
		Else
			itemSellTypeCd = "20"
		End If
		strRst = ""
		strRst2 = ""
		strRst3 = ""
		strRst = strRst & "	<itemSellTypeCd>"&itemSellTypeCd&"</itemSellTypeCd>"							'#상품판매유형코드 (commCd:I006) | 10 : 일반, 20 : 옵션
		strRst = strRst & "	<itemSellTypeDtlCd>10</itemSellTypeDtlCd>"

		If (FOptionCnt > 0) Then
			strRst = strRst & "	<uitems>"

			'#옵션명 생성
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				Do until rsget.EOF
					arrOptTypeNm = arrOptTypeNm & Replace(db2Html(rsget("optionTypeName")),",","")
					rsget.MoveNext
					If Not(rsget.EOF) Then arrOptTypeNm = arrOptTypeNm & ","
				Loop
			End If
			rsget.Close
			arrOptTypeNm = Split(arrOptTypeNm, ",")

			strSql = "EXEC db_item.dbo.usp_Ten_OutMall_Ssg_optEditParamList '"&CMallName&"'," & FItemid
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows
			End If
			rsget.close

			If chkMultiOpt Then					'###################### 이중옵션일 때 '######################
				Select Case Ubound(arrOptTypeNm)
					Case "1"
						OptTypeNm1 = Trim(arrOptTypeNm(0))
						OptTypeNm2 = Trim(arrOptTypeNm(1))
						OptTypeNm3 = ""
					Case "2"
						OptTypeNm1 = Trim(arrOptTypeNm(0))
						OptTypeNm2 = Trim(arrOptTypeNm(1))
						OptTypeNm3 = Trim(arrOptTypeNm(2))
				End Select

				For i = 0 To UBound(ArrRows,2)
					itemoption		= ArrRows(1,i)
					outmalloptcode	= ArrRows(2,i)
					outmalloptName	= Replace(Replace(db2Html(ArrRows(3,i)),":",""),",","")
					optlimit		= ArrRows(4,i)
					optlimityn		= ArrRows(5,i)
					isusing			= ArrRows(6,i)
					optsellyn		= ArrRows(7,i)
					opt1name		= ArrRows(8,i)
					opt2name		= ArrRows(9,i)
					opt3name		= ArrRows(10,i)
					preged			= (ArrRows(11,i)=1)
					optNameDiff		= (ArrRows(12,i)=1)
					oopt			= ArrRows(13,i)
					optaddprice		= ArrRows(14,i)

				    optLimit = optLimit - 5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

					If preged = 0 Then
						If (isUsing="N") or (optsellyn="N") or (FLimityn = "Y" AND optLimit <= 0) Then
							sellStatCd = "80"
					    Else
							sellStatCd = "20"
						End If
					Else
						If (optNameDiff) or (isUsing="N") or (optsellyn="N") or (FLimityn = "Y" AND optLimit <= 0) Then
							sellStatCd = "80"
					    Else
							sellStatCd = "20"
						End If
					End If

					strRst = strRst & "		<uitem>"
				If preged = 0 Then
					strRst = strRst & "			<tempUitemId>"&itemoption&"</tempUitemId>"						'#단품ID (임시번호)
					strRst = strRst & "			<uitemOptnTypeNm1>"&OptTypeNm1&"</uitemOptnTypeNm1>"			'#단품 옵션 유형명1
					strRst = strRst & "			<uitemOptnNm1>"&opt1name&"</uitemOptnNm1>"					'#단품 옵션 명1
					strRst = strRst & "			<uitemOptnTypeNm2>"&OptTypeNm2&"</uitemOptnTypeNm2>"			'단품 옵션 유형명2
					strRst = strRst & "			<uitemOptnNm2>"&opt2name&"</uitemOptnNm2>"					'단품 옵션 명2
					strRst = strRst & "			<uitemOptnTypeNm3>"&OptTypeNm3&"</uitemOptnTypeNm3>"			'단품 옵션 유형명3
					strRst = strRst & "			<uitemOptnNm3>"&opt3name&"</uitemOptnNm3>"					'단품 옵션 명3
					strRst = strRst & "			<uitemOptnTypeNm4></uitemOptnTypeNm4>"							'단품 옵션 유형명4
					strRst = strRst & "			<uitemOptnNm4></uitemOptnNm4>"									'단품 옵션 명4
					strRst = strRst & "			<uitemOptnTypeNm5></uitemOptnTypeNm5>"							'단품 옵션 유형명5
					strRst = strRst & "			<uitemOptnNm5></uitemOptnNm5>"									'단품 옵션 명5
				Else
					strRst = strRst & "			<uitemId>"&outmalloptcode&"</uitemId>"								'#단품ID
				End If
					strRst = strRst & "			<sellStatCd>"&sellStatCd&"</sellStatCd>"						'판매상태코드 | 20:판매중, 80:일시판매중지, 90:영구판매중지
					strRst = strRst & "			<baseInvQty>"&optLimit&"</baseInvQty>"							'재고 수량
					strRst = strRst & "			<useYn>Y</useYn>"												'사용 여부...Y로 그냥 보내도 되나??
					strRst = strRst & "		</uitem>"

					strRst3 = strRst3 & "		<uitemPrc>"
				If preged = 0 Then
					strRst3 = strRst3 & "			<tempUitemId>"&itemoption&"</tempUitemId>"					'#단품ID (임시번호)
				Else
					strRst3 = strRst3 & "			<uitemId>"&outmalloptcode&"</uitemId>"							'#단품ID
				End If
					strRst3 = strRst3 & "			<siteNo>6004</siteNo>"										'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
					strRst3 = strRst3 & "			<splprc>"&(MustPrice + optaddprice) * 0.85&"</splprc>"		'#공급가
					strRst3 = strRst3 & "			<sellprc>"&MustPrice + optaddprice&"</sellprc>"				'#판매가
					strRst3 = strRst3 & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"								'#마진율
					strRst3 = strRst3 & "		</uitemPrc>"
				Next
			Else
				For i = 0 To UBound(ArrRows,2)
					itemoption		= ArrRows(1,i)
					outmalloptcode	= ArrRows(2,i)
					outmalloptName	= Replace(Replace(db2Html(ArrRows(3,i)),":",""),",","")
					optlimit		= ArrRows(4,i)
					optlimityn		= ArrRows(5,i)
					isusing			= ArrRows(6,i)
					optsellyn		= ArrRows(7,i)
					opt1name		= ArrRows(13,i)
					opt2name		= ""
					opt3name		= ""
					preged			= (ArrRows(11,i)=1)
					optNameDiff		= (ArrRows(12,i)=1)
					oopt			= ArrRows(13,i)
					optaddprice		= ArrRows(14,i)
					OptTypeNm1		= ArrRows(15,i)

					If OptTypeNm1 = "" Then
						OptTypeNm1 = "선택"
					End If

				    optLimit = optLimit - 5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

					If preged = 0 Then
						If (isUsing="N") or (optsellyn="N") or (FLimityn = "Y" AND optLimit <= 0) Then
							sellStatCd = "80"
					    Else
							sellStatCd = "20"
						End If
					Else
						If (optNameDiff) or (isUsing="N") or (optsellyn="N") or (FLimityn = "Y" AND optLimit <= 0) Then
							sellStatCd = "80"
					    Else
							sellStatCd = "20"
						End If
					End If
					strRst = strRst & "		<uitem>"
				If preged = 0 Then
					strRst = strRst & "			<tempUitemId>"&itemoption&"</tempUitemId>"						'#단품ID (임시번호)
					strRst = strRst & "			<uitemOptnTypeNm1>"&OptTypeNm1&"</uitemOptnTypeNm1>"			'#단품 옵션 유형명1
					strRst = strRst & "			<uitemOptnNm1>"&opt1name&"</uitemOptnNm1>"					'#단품 옵션 명1
					strRst = strRst & "			<uitemOptnTypeNm2></uitemOptnTypeNm2>"							'단품 옵션 유형명2
					strRst = strRst & "			<uitemOptnNm2></uitemOptnNm2>"									'단품 옵션 명2
					strRst = strRst & "			<uitemOptnTypeNm3></uitemOptnTypeNm3>"							'단품 옵션 유형명3
					strRst = strRst & "			<uitemOptnNm3></uitemOptnNm3>"									'단품 옵션 명3
					strRst = strRst & "			<uitemOptnTypeNm4></uitemOptnTypeNm4>"							'단품 옵션 유형명4
					strRst = strRst & "			<uitemOptnNm4></uitemOptnNm4>"									'단품 옵션 명4
					strRst = strRst & "			<uitemOptnTypeNm5></uitemOptnTypeNm5>"							'단품 옵션 유형명5
					strRst = strRst & "			<uitemOptnNm5></uitemOptnNm5>"									'단품 옵션 명5
				Else
					strRst = strRst & "			<uitemId>"&outmalloptcode&"</uitemId>"								'#단품ID
				End If
					strRst = strRst & "			<sellStatCd>"&sellStatCd&"</sellStatCd>"						'판매상태코드 | 20:판매중, 80:일시판매중지, 90:영구판매중지
					strRst = strRst & "			<baseInvQty>"&optLimit&"</baseInvQty>"							'재고 수량
					strRst = strRst & "			<useYn>Y</useYn>"												'사용 여부...Y로 그냥 보내도 되나??
					strRst = strRst & "		</uitem>"

					strRst3 = strRst3 & "		<uitemPrc>"
				If preged = 0 Then
					strRst3 = strRst3 & "			<tempUitemId>"&itemoption&"</tempUitemId>"					'#단품ID (임시번호)
				Else
					strRst3 = strRst3 & "			<uitemId>"&outmalloptcode&"</uitemId>"							'#단품ID
				End If
					strRst3 = strRst3 & "			<siteNo>6004</siteNo>"										'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
					strRst3 = strRst3 & "			<splprc>"&(MustPrice + optaddprice) * 0.85&"</splprc>"		'#공급가
					strRst3 = strRst3 & "			<sellprc>"&MustPrice + optaddprice&"</sellprc>"				'#판매가
					strRst3 = strRst3 & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"								'#마진율
					strRst3 = strRst3 & "		</uitemPrc>"
				Next
			End If
		strRst = strRst & "	</uitems>"
		End If

		If FitemDiv = "06" Then
			requireDetailStr = ""
			requireDetailStr = requireDetailStr & "	<itemOrdOptns>"
			requireDetailStr = requireDetailStr & "		<itemOrdOptn>"
			requireDetailStr = requireDetailStr & "			<addOrdOptnSeq>1</addOrdOptnSeq>"						'#추가 주문 옵션 순번
			requireDetailStr = requireDetailStr & "			<addOrdOptnNm>주문제작문구</addOrdOptnNm>"				'#추가 주문 옵션명
			requireDetailStr = requireDetailStr & "		</itemOrdOptn>"
			requireDetailStr = requireDetailStr & "	</itemOrdOptns>"
		End If

		If FOptionCnt > 0 Then
			strRst2 = strRst2 & "	<uitemPluralPrcs>"
			strRst2 = strRst2 & strRst3
			strRst2 = strRst2 & Replace(strRst3, "<siteNo>6004</siteNo>", "<siteNo>6001</siteNo>")					'// 이마트몰 추가
			strRst2 = strRst2 & "	</uitemPluralPrcs>"
		End If
'response.write strRst & requireDetailStr & strRst2
'response.end
		getSsgOptParamtoEDIT = strRst & requireDetailStr & strRst2
	End Function

	Public Function getSsgOptParamtoREG()
		Dim strRst, strRst2, strRst3, strSql, chkMultiOpt, arrOptTypeNm, requireDetailStr
		Dim itemSellTypeCd, OptTypeNm1, OptTypeNm2, OptTypeNm3, optLimit, itemoption, arrOptionname, optionname1, optionname2, optionname3, optaddprice

		If FOptionCnt = 0 Then			'단품
			itemSellTypeCd = "10"
		Else
			itemSellTypeCd = "20"
		End If
		strRst = ""
		strRst2 = ""
		strRst3 = ""
		strRst = strRst & "	<itemSellTypeCd>"&itemSellTypeCd&"</itemSellTypeCd>"							'#상품판매유형코드 (commCd:I006) | 10 : 일반, 20 : 옵션
		strRst = strRst & "	<itemSellTypeDtlCd>10</itemSellTypeDtlCd>"										'#상품판매유형상세코드 (commCd:I007) | 10 : 일반, 30 : 30 기획 (신세계몰은 기획상품 불가능)

		If FOptionCnt > 0 Then
			strRst = strRst & "	<uitems>"
			'#옵션명 생성
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				Do until rsget.EOF
					arrOptTypeNm = arrOptTypeNm & Replace(db2Html(rsget("optionTypeName")),",","")
					rsget.MoveNext
					If Not(rsget.EOF) Then arrOptTypeNm = arrOptTypeNm & ","
				Loop
			End If
			rsget.Close
			arrOptTypeNm = Split(arrOptTypeNm, ",")

			If chkMultiOpt Then					'###################### 이중옵션일 때 '######################
				Select Case Ubound(arrOptTypeNm)
					Case "1"
						OptTypeNm1 = Trim(arrOptTypeNm(0))
						OptTypeNm2 = Trim(arrOptTypeNm(1))
						OptTypeNm3 = ""
					Case "2"
						OptTypeNm1 = Trim(arrOptTypeNm(0))
						OptTypeNm2 = Trim(arrOptTypeNm(1))
						OptTypeNm3 = Trim(arrOptTypeNm(2))
				End Select

				strSql = ""
				strSql = strSql & " SELECT itemid, itemoption, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optionTypeName, optionname, optaddprice, (optlimitno-optlimitsold) as optLimit "
				strSql = strSql & " FROM db_item.dbo.tbl_item_option "
				strSql = strSql & " WHERE isusing = 'Y' and itemid=" & FItemid &"  "
				strSql = strSql & " ORDER BY itemoption ASC "
				rsget.Open strSql,dbget,1
				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit - 5
					    If (optLimit < 1) Then optLimit = 0
					    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

						itemoption = rsget("itemoption")
						arrOptionname = rsget("optionname")
						arrOptionname = Split(arrOptionname, ",")
						optaddprice = rsget("optaddprice")

						Select Case Ubound(arrOptTypeNm)
							Case "1"
								optionname1 = Trim(arrOptionname(0))
								optionname2 = Trim(arrOptionname(1))
								optionname3 = ""
							Case "2"
								optionname1 = Trim(arrOptionname(0))
								optionname2 = Trim(arrOptionname(1))
								optionname3 = Trim(arrOptionname(2))
						End Select
						strRst = strRst & "		<uitem>"
						strRst = strRst & "			<tempUitemId>"&itemoption&"</tempUitemId>"						'#단품ID (임시번호)
						strRst = strRst & "			<uitemOptnTypeNm1>"&OptTypeNm1&"</uitemOptnTypeNm1>"			'#단품 옵션 유형명1
						strRst = strRst & "			<uitemOptnNm1>"&optionname1&"</uitemOptnNm1>"					'#단품 옵션 명1
						strRst = strRst & "			<uitemOptnTypeNm2>"&OptTypeNm2&"</uitemOptnTypeNm2>"			'단품 옵션 유형명2
						strRst = strRst & "			<uitemOptnNm2>"&optionname2&"</uitemOptnNm2>"					'단품 옵션 명2
						strRst = strRst & "			<uitemOptnTypeNm3>"&OptTypeNm3&"</uitemOptnTypeNm3>"			'단품 옵션 유형명3
						strRst = strRst & "			<uitemOptnNm3>"&optionname3&"</uitemOptnNm3>"					'단품 옵션 명3
						strRst = strRst & "			<uitemOptnTypeNm4></uitemOptnTypeNm4>"							'단품 옵션 유형명4
						strRst = strRst & "			<uitemOptnNm4></uitemOptnNm4>"									'단품 옵션 명4
						strRst = strRst & "			<uitemOptnTypeNm5></uitemOptnTypeNm5>"							'단품 옵션 유형명5
						strRst = strRst & "			<uitemOptnNm5></uitemOptnNm5>"									'단품 옵션 명5
						strRst = strRst & "			<baseInvQty>"&optLimit&"</baseInvQty>"							'재고 수량
						strRst = strRst & "			<useYn>Y</useYn>"												'사용 여부...Y로 그냥 보내도 되나??
						strRst = strRst & "		</uitem>"

						strRst3 = strRst3 & "		<uitemPrc>"
						strRst3 = strRst3 & "			<tempUitemId>"&itemoption&"</tempUitemId>"					'#단품ID (임시번호)
						strRst3 = strRst3 & "			<siteNo>6004</siteNo>"										'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
						strRst3 = strRst3 & "			<splprc>"&(MustPrice + optaddprice) * 0.85&"</splprc>"											'#공급가
						strRst3 = strRst3 & "			<sellprc>"&MustPrice + optaddprice&"</sellprc>"				'#판매가
						strRst3 = strRst3 & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"								'#마진율
						strRst3 = strRst3 & "		</uitemPrc>"
						rsget.MoveNext
					Loop
				End If
				rsget.Close
			Else								'###################### 단일옵션일 때 '######################
				strSql = ""
				strSql = strSql & " SELECT itemid, itemoption, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, isnull(optionTypeName, '') as optionTypeName, optionname, optaddprice, (optlimitno-optlimitsold) as optLimit "
				strSql = strSql & " FROM db_item.dbo.tbl_item_option "
				strSql = strSql & " WHERE isusing = 'Y' and itemid=" & FItemid &"  "
				strSql = strSql & " ORDER BY itemoption ASC "
				rsget.Open strSql,dbget,1
				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit - 5
					    If (optLimit < 1) Then optLimit = 0
					    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

						itemoption = rsget("itemoption")
						optionname1 = rsget("optionname")
						OptTypeNm1 = rsget("optionTypeName")
						optaddprice = rsget("optaddprice")
						If OptTypeNm1 = "" Then
							OptTypeNm1 = "선택"
						End If
						strRst = strRst & "		<uitem>"
						strRst = strRst & "			<tempUitemId>"&itemoption&"</tempUitemId>"						'#단품ID (임시번호)
						strRst = strRst & "			<uitemOptnTypeNm1>"&OptTypeNm1&"</uitemOptnTypeNm1>"			'#단품 옵션 유형명1
						strRst = strRst & "			<uitemOptnNm1>"&optionname1&"</uitemOptnNm1>"					'#단품 옵션 명1
						strRst = strRst & "			<uitemOptnTypeNm2></uitemOptnTypeNm2>"							'단품 옵션 유형명2
						strRst = strRst & "			<uitemOptnNm2></uitemOptnNm2>"									'단품 옵션 명2
						strRst = strRst & "			<uitemOptnTypeNm3></uitemOptnTypeNm3>"							'단품 옵션 유형명3
						strRst = strRst & "			<uitemOptnNm3></uitemOptnNm3>"									'단품 옵션 명3
						strRst = strRst & "			<uitemOptnTypeNm4></uitemOptnTypeNm4>"							'단품 옵션 유형명4
						strRst = strRst & "			<uitemOptnNm4></uitemOptnNm4>"									'단품 옵션 명4
						strRst = strRst & "			<uitemOptnTypeNm5></uitemOptnTypeNm5>"							'단품 옵션 유형명5
						strRst = strRst & "			<uitemOptnNm5></uitemOptnNm5>"									'단품 옵션 명5
						strRst = strRst & "			<baseInvQty>"&optLimit&"</baseInvQty>"							'재고 수량
						strRst = strRst & "			<useYn>Y</useYn>"												'사용 여부...Y로 그냥 보내도 되나??
						strRst = strRst & "		</uitem>"

						strRst3 = strRst3 & "		<uitemPrc>"
						strRst3 = strRst3 & "			<tempUitemId>"&itemoption&"</tempUitemId>"					'#단품ID (임시번호)
						strRst3 = strRst3 & "			<siteNo>6004</siteNo>"										'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
						strRst3 = strRst3 & "			<splprc>"&(MustPrice + optaddprice) * 0.85&"</splprc>"											'#공급가
						strRst3 = strRst3 & "			<sellprc>"&MustPrice + optaddprice&"</sellprc>"				'#판매가
						strRst3 = strRst3 & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"								'#마진율
						strRst3 = strRst3 & "		</uitemPrc>"
						rsget.MoveNext
					Loop
				End If
				rsget.Close
			End If
			strRst = strRst & "	</uitems>"
		End If

		If FitemDiv = "06" Then
			requireDetailStr = ""
			requireDetailStr = requireDetailStr & "	<itemOrdOptns>"
			requireDetailStr = requireDetailStr & "		<itemOrdOptn>"
			requireDetailStr = requireDetailStr & "			<addOrdOptnSeq>1</addOrdOptnSeq>"						'#추가 주문 옵션 순번
			requireDetailStr = requireDetailStr & "			<addOrdOptnNm>주문제작문구</addOrdOptnNm>"				'#추가 주문 옵션명
			requireDetailStr = requireDetailStr & "		</itemOrdOptn>"
			requireDetailStr = requireDetailStr & "	</itemOrdOptns>"
		End If
		If FOptionCnt > 0 Then
			strRst2 = strRst2 & "	<uitemPluralPrcs>"
			strRst2 = strRst2 & strRst3
			strRst2 = strRst2 & Replace(strRst3, "<siteNo>6004</siteNo>", "<siteNo>6001</siteNo>")					'// 이마트몰 추가
			strRst2 = strRst2 & "	</uitemPluralPrcs>"
		End If
		getSsgOptParamtoREG = strRst & requireDetailStr & strRst2
	End Function

	Public Function getSsgItemInfoCdToReg(iareaCode)
		Dim strSql, buf, lp
		Dim mallinfoCd, infoContent
		strSql = ""
		strSql = strSql & " SELECT top 100 M.* , "
		strSql = strSql & " CASE WHEN (M.infoCdAdd='00000') AND (F.chkDiv='Y') THEN 'Y' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00000') AND (F.chkDiv='N') THEN 'N' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00001') AND (F.chkDiv='Y') THEN '10' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00001') AND (F.chkDiv='N') THEN '20' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00002') AND (F.chkDiv='Y') THEN 'O' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00002') AND (F.chkDiv='N') THEN 'N' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00002') AND (F.chkDiv='N') THEN 'N' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00003') THEN 'N' "
		strSql = strSql & " 	 WHEN (M.mallinfoCd='0000000011') THEN '"&iareaCode&"' "
		strSql = strSql & " 	 WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035' "
		strSql = strSql & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent "
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"' "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='"&FItemid&"' "
		strSql = strSql & " WHERE M.mallid = 'ssg' and IC.itemid='"&FItemid&"' "
		rsget.Open strSql,dbget,1
		buf = ""
		If not rsget.EOF Then
			buf = buf & "	<itemMngPropClsId>"& rsget("infoETC") &"</itemMngPropClsId>"
			buf = buf & "	<itemMngAttrs>"
			Do until rsget.EOF
				infoContent = rsget("infocontent")
				mallinfocd = rsget("mallinfocd")
				buf = buf & "	<itemMngAttr>"
				buf = buf & "		<itemMngPropId>"&mallinfocd&"</itemMngPropId>"
				buf = buf & "		<itemMngCntt><![CDATA["&infoContent&"]]></itemMngCntt>"
				buf = buf & "	</itemMngAttr>"
				rsget.MoveNext
			Loop
			buf = buf & "	</itemMngAttrs>"
		End If
		rsget.Close
		getSsgItemInfoCdToReg = buf
	End Function

	'SSG 등록 XML
	Public Function getSsgItemRegParameter()
		Dim strRst, i, sellStatCd, areaCode, shppItemDivCd, shppRqrmDcnt, shppRqrmDcntChngRsnCntt

		sellStatCd = 20
		'################################ 카테고리 항목 호출 ########################################
		Dim callCategory , standardCateCode, arrDisplayCateCode, arrSiteNum
		callCategory = getSsgCategoryParam()
		standardCateCode = Split(callCategory, "|_|")(0)
		arrDisplayCateCode = Split(Split(callCategory, "|_|")(1), ",")
		arrSiteNum = Split(Split(callCategory, "|_|")(2), ",")
		'##########################################################################################
		'################################### 원산지  호출 ##########################################
		areaCode = getSourcearea()
		'##########################################################################################
		'################################### 배송기일  호출 #########################################
		shppRqrmDcnt = getShopLeadTime()
		'##########################################################################################
'		If FItemdiv = "06" OR FItemdiv = "16" Then
'			shppItemDivCd = "05"
'			If FRequireMakeDay < 1 Then
'				shppRqrmDcnt = 7
'			Else
'				shppRqrmDcnt = FRequireMakeDay
'			End If
'			shppRqrmDcntChngRsnCntt = "주문제작상품"
'		Else
'			shppItemDivCd = "01"
'			shppRqrmDcnt = 3
'		End If

		shppItemDivCd = "01"
		If getShopLeadTime > 3 Then
			shppRqrmDcntChngRsnCntt = "주문제작상품"
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst & "<insertItem>"
		strRst = strRst & "	<itemNm><![CDATA["&getItemNameFormat&"]]></itemNm>"								'#상품명
		strRst = strRst & "	<mdlNm></mdlNm>"																'모델명
		strRst = strRst & "	<brandId>2000047517</brandId>"													'#브랜드ID | 텐바이텐(2000047517)
		strRst = strRst & "	<stdCtgId>"&standardCateCode&"</stdCtgId>"										'#표준카테고리ID
		strRst = strRst & "	<sites>"
		strRst = strRst & "		<site>"
		strRst = strRst & "			<siteNo>6004</siteNo>"													'#사이트번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
		strRst = strRst & "			<sellStatCd>"&sellStatCd&"</sellStatCd>"								'#판매 상태 코드 | 20 : 판매중, 80 : 일시판매중지
		strRst = strRst & "		</site>"
		strRst = strRst & "		<site>"
		strRst = strRst & "			<siteNo>6001</siteNo>"													'#사이트번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
		strRst = strRst & "			<sellStatCd>"&sellStatCd&"</sellStatCd>"								'#판매 상태 코드 | 20 : 판매중, 80 : 일시판매중지
		strRst = strRst & "		</site>"
		strRst = strRst & "	</sites>"
		strRst = strRst & "	<itemAplRngTypeCd></itemAplRngTypeCd>"											'상품 적용 범위 | 00 : 전체적용, 10 : B2C적용, 20 : B2E적용
		strRst = strRst & "	<b2eAplRngCd>10</b2eAplRngCd>"													'B2E 적용 범위 | 10 : 전체 적용, 20 : 적용 않음, 30 : 회원사 지정
		strRst = strRst & "	<b2cAplRngCd>20</b2cAplRngCd>"													'B2C 적용 범위 | 10 : 적용, 20 : 적용 (대행 제휴사 제외), 70 : 적용 않음
		strRst = strRst & "	<itemChrctDivCd>10</itemChrctDivCd>"											'#상품 특성 구분 코드 | 10 : 일반, 40 : 미가공 귀금속, 50 : 모바일 기프트, 60 : 상품권, 70 : 쇼핑 충전금
		strRst = strRst & "	<itemChrctDtlCd></itemChrctDtlCd>"												'#상품 특성 상세 코드 | 상품 특성 구분 코드(itemChrctDivCd = 50 | 60) 일 경우 상품 특성 구분 코드(itemChrctDivCd = 50) => 10 : 일반, 50 : 상품권, 상품 특성 구분 코드(itemChrctDivCd = 60) => 60 : 신세계 지류 상품권, 70 : 외부 지류 상품권, 80 : 기프트 카드, 90 : 맞춤형 기프트 카드
		strRst = strRst & "	<exusItemDivCd>10</exusItemDivCd>"												'#전용 상품 구분 코드 | 10 : 일반, 20 : GIFT(일반)
		strRst = strRst & "	<exusItemDtlCd>10</exusItemDtlCd>"												'#전용 상품 상세 코드 | 10 : 일반, 20 : 특정점
		strRst = strRst & "	<dispAplRngTypeCd>10</dispAplRngTypeCd>"										'#전시 적용 범위 유형 코드 | 10 : 전체 (모바일 + PC), 30 : 모바일 (모바일 선택시 전체로 설정 불가)
		strRst = strRst & "	<speSalestrNo></speSalestrNo>"													'특정 영업점 번호 (특정점 (exusItemDtlCd=20)일 경우 입력) | ※ 특정점코드 API 참조
		strRst = strRst & getSsgItemInfoCdToReg(areaCode)
		strRst = strRst & "	<manufcoNm><![CDATA["&Trim(FMakername)&"]]></manufcoNm>"						'#제조사명
		strRst = strRst & "	<prodManufCntryId>"&areaCode&"</prodManufCntryId>"								'#생산 제조 국가 ID | (참고 : 원산지조회API(listOrplc API))
		strRst = strRst & "	<dispCtgs>"
	For i = 0 to Ubound(arrDisplayCateCode)
		strRst = strRst & "		<dispCtg>"
		strRst = strRst & "			<siteNo>"&arrSiteNum(i)&"</siteNo>"										'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰, 6005 : SSG.COM
		strRst = strRst & "			<dispCtgId>"&arrDisplayCateCode(i)&"</dispCtgId>"						'#전시 카테고리 ID
		strRst = strRst & "			<repDispOrdr>"&i+1&"</repDispOrdr>"										'#대표 전시 순서 | 순서대로, 중복 허용하지 않음. 사이트별 최대 3개까지 선택 가능
		strRst = strRst & "		</dispCtg>"
	Next
		strRst = strRst & "	</dispCtgs>"
		strRst = strRst & "	<dispStrtDts>"&Replace(Date(), "-", "")&"</dispStrtDts>"						'#전시시작일시(YYYYMMDD OR YYYYMMDDHH24MISS)
		strRst = strRst & "	<dispEndDts>29991231</dispEndDts>"												'#전시종료일시(YYYYMMDD OR YYYYMMDDHH24MISS)
'		strRst = strRst & "	<spDispCtgs>"																	'-------- MayBe 전문카테고리 인 듯.. --------
'		strRst = strRst & "		<dispCtg>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
'		strRst = strRst & "			<dispCtgId></dispCtgId>"												'#전시 카테고리 ID
'		strRst = strRst & "			<repDispOrdr></repDispOrdr>"											'#대표 전시 순서 | 순서대로, 중복 허용하지 않음. 사이트별 최대 3개까지 선택 가능
'		strRst = strRst & "		</dispCtg>"
'		strRst = strRst & "	</spDispCtgs>"
		strRst = strRst & "	<srchPsblYn>Y</srchPsblYn>"														'검색 가능 여부
		strRst = strRst & "	<itemSrchwdNm><![CDATA["&RightCommaDel(Trim(FKeywords))&"]]></itemSrchwdNm>"	'상품검색어명
		strRst = strRst & "	<aplMbrGrdCd></aplMbrGrdCd>"													'노출 회원 등급 (값이 존재하지 않을 경우 ALL) | 10 : 패밀리, 20 : 브론즈, 30 : 실버, 40 : 골드, 50 : VIP, 90 : VVIP
		strRst = strRst & "	<minOnetOrdPsblQty>1</minOnetOrdPsblQty>"										'#최소 1회 주문 가능 수량
		strRst = strRst & "	<maxOnetOrdPsblQty>9999</maxOnetOrdPsblQty>"									'#최대 1회 주문 가능 수량
		strRst = strRst & "	<max1dyOrdPsblQty>9999</max1dyOrdPsblQty>"										'#최대 1일 주문 가능 수량
		strRst = strRst & "	<adultItemTypeCd>90</adultItemTypeCd>"											'#성인 상품 타입 코드 (commCd:I408) | 10 : 성인 상품, 20 : 주류 상품, 90 : 일반 상품
		strRst = strRst & "	<hriskItemYn>N</hriskItemYn>"													'#고 위험 상품 여부
		strRst = strRst & "	<nitmAplYn>N</nitmAplYn>" 														'#신 상품 적용 여부
'		strRst = strRst & "	<sellPnts>"																		'-------- MayBe 셀링포인트 인 듯.. --------
'		strRst = strRst & "		<sellPnt>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
'		strRst = strRst & "			<sellpntNm></sellpntNm>"												'#
'		strRst = strRst & "			<dispStrtDts></dispStrtDts>"											'#전시 시작 일시 (YYYYMMDD)
'		strRst = strRst & "			<dispEndDts></dispEndDts>"												'#전시 종료 일시 (YYYYMMDD)
'		strRst = strRst & "			<useYn></useYn>"														'#사용 여부
'		strRst = strRst & "		</sellPnt>"
'		strRst = strRst & "	</sellPnts>"
		strRst = strRst & "	<sellCapaUnitCd></sellCapaUnitCd>"												'판매 용량 단위 코드 (commCd:I159) | 01 ea, 02 cc, 03 g, 04 kg, 05 m, 06 ml, 07 mm, 08 개, 09 매, 10 포
		strRst = strRst & "	<sellTotCapa></sellTotCapa>"													'판매 총 용량
		strRst = strRst & "	<sellUnitCapa></sellUnitCapa>"													'판매 단위 용량
		strRst = strRst & "	<sellUnitQty>0</sellUnitQty>"													'판매 단위 수량
		strRst = strRst & "	<buyFrmCd>60</buyFrmCd>"														'#매입 형태 코드 (commCd:I002) | 10 : 직매입, 20 : 직매입2(판매분), 40 : 특정매입, 60 : 위수탁
		strRst = strRst & "	<txnDivCd>"&CHKIIF(FVatInclude="N","20","10")&"</txnDivCd>"						'#과세 구분 코드 (commCd:I005) | 10 : 과세, 20 : 면세, 30 : 영세
		strRst = strRst & "	<prcMngMthd>1</prcMngMthd>"														'가격설정방식 | 1 : 공급가 자동계산 (Default), 2 : 판매가 자동계산, 3 : 마진 자동계산..이 값 설정시 SALE_PRC_INFO, B2E_PRC 둘다 적용 받는다. 값은 모두 입력 받아도 상관 없으나 해당 값 설정에 따라 해당 값이 자동으로 계산됨.
		strRst = strRst & "	<salesPrcInfos>"
		strRst = strRst & "		<uitemPrc>"
		strRst = strRst & "			<siteNo>6004</siteNo>"													'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
		strRst = strRst & "			<splprc>"&MustPrice()*0.85&"</splprc>"														'#공급가
		strRst = strRst & "			<sellprc>"&MustPrice()&"</sellprc>"										'#판매가
		strRst = strRst & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"											'#마진율
		strRst = strRst & "		</uitemPrc>"
		strRst = strRst & "		<uitemPrc>"
		strRst = strRst & "			<siteNo>6001</siteNo>"													'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
		strRst = strRst & "			<splprc>"&MustPrice()*0.85&"</splprc>"														'#공급가
		strRst = strRst & "			<sellprc>"&MustPrice()&"</sellprc>"										'#판매가
		strRst = strRst & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"											'#마진율
		strRst = strRst & "		</uitemPrc>"
		strRst = strRst & "	</salesPrcInfos>"
'		strRst = strRst & "	<b2ePrcAplTgts>"
'		strRst = strRst & "		<b2ePrcAplTgt>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
'		strRst = strRst & "			<b2eMbrcoId></b2eMbrcoId>"												'#B2E회원사ID
'		strRst = strRst & "			<b2eSplprc></b2eSplprc>"												'#B2E 공급가
'		strRst = strRst & "			<b2eSellprc></b2eSellprc>"												'#B2E 판매가
'		strRst = strRst & "			<b2eMrgrt></b2eMrgrt>"													'#B2E 마진율
'		strRst = strRst & "		</b2ePrcAplTgt>"
'		strRst = strRst & "	</b2ePrcAplTgts>"
		strRst = strRst & "	<invMngYn>Y</invMngYn>"															'#재고 관리 여부
		strRst = strRst & "	<baseInvQty>"&getLimitEa()&"</baseInvQty>"										'#재고 수량
		strRst = strRst & "	<invQtyMarkgYn>Y</invQtyMarkgYn>"												'#재고 수량 표기 여부
'		strRst = strRst & "	<rsvSaleInfo>"
'		strRst = strRst & "		<aplStrtDt></aplStrtDt>"													'#예약판매 시작일 (YYYYMMDD)
'		strRst = strRst & "		<aplEndDt></aplEndDt>"														'#예약판매 종료일 (YYYYMMDD)
'		strRst = strRst & "		<whoutStrtDt></whoutStrtDt>"												'#출고 시작 일자 (YYYYMMDD)
'		strRst = strRst & "		<rstctInvQty></rstctInvQty>"												'#예약 판매 수량
'		strRst = strRst & "	</rsvSaleInfo>"
		strRst = strRst & getSsgOptParamtoREG()
		strRst = strRst & "	<shppItemDivCd>"&shppItemDivCd&"</shppItemDivCd>"								'#배송상품구분코드 (commCd:I070) | 01 : 일반, 02 : 해외구매대행, 03 : 설치(유료), 04 : 설치(무료), 05 : 주문제작, 06 : 해외직배송
		strRst = strRst & "	<exprtCntryId></exprtCntryId>"													'수출국가(해외 직배송 적출국)shppItemDivCd=06(해외직배송) 인 경우 필수 | 원산지 조회 API 참고(listOrplc API)
		strRst = strRst & "	<pcusMngCd></pcusMngCd>"														'개인 통관 고유 부호 | shppItemDivCd=06(해외직배송) 인 경우 필수 10 : 선택 입력, 20 : 필수 입력, 30 : 입력 안함
		strRst = strRst & "	<retExchPsblYn>Y</retExchPsblYn>"												'#반품 교환 가능 여부
		strRst = strRst & "	<shppMainCd>41</shppMainCd>"													'#배송 주체 코드 (commCd:P017) | 31 : 자사창고, 32 : 업체창고, 41 : 협력업체
		strRst = strRst & "	<shppMthdCd>20</shppMthdCd>"													'#배송 방법 코드 (commCd:P021) | 10 : 자사배송, 20 : 택배배송, 30 : 매장방문, 40 : 등기, 50 : 미배송, 60 : 미발송
		strRst = strRst & "	<mareaShppYn></mareaShppYn>"													'#수도권 배송여부
		strRst = strRst & "	<shppRqrmDcnt>"&shppRqrmDcnt&"</shppRqrmDcnt>"									'#배송 소요 일수
		strRst = strRst & "	<shppRqrmDcntChngRsnCntt>"&shppRqrmDcntChngRsnCntt&"</shppRqrmDcntChngRsnCntt>"	'#배송 소요 일수 변경 사유 | 상품배송구분이 일반(01) 이고 배송소요일수가 4일 이상일 경우 필수
		strRst = strRst & "	<splVenItemId>"&FItemid&"</splVenItemId>"										'업체 상품 번호
		strRst = strRst & "	<whoutShppcstId>0000517199</whoutShppcstId>"									'#출고 배송비 ID
		strRst = strRst & "	<retShppcstId>0000011336</retShppcstId>"										'#반품 배송비 ID
		strRst = strRst & "	<whoutAddrId>0000006297</whoutAddrId>"											'#출고 주소 ID
		strRst = strRst & "	<snbkAddrId>0000006297</snbkAddrId>"											'#반품 주소 ID
		strRst = strRst & "	<frgShppPsblYn>N</frgShppPsblYn>"												'#해외 배송 가능 여부
		strRst = strRst & "	<itemTotWgt></itemTotWgt>"												  		'상품 총 무게
		strRst = strRst & "	<hopeShppDdDivCd></hopeShppDdDivCd>"											'희망 발송일 구분 코드 (commCd:I015) | 10 : 15일이내, 20 : 15일이후 30일이내, 30 : 30일이후, 90 : 발송일 최대 날짜 지정
		strRst = strRst & "	<hopeShppDdEndDts></hopeShppDdEndDts>"											'희망 발송일 종료 일시 (YYYYMMDD) | 희망발송일 구분코드가 (hopeShppDdEndDts=90) 일경우 필수
		strRst = strRst & getSsgAddImageParam()
		strRst = strRst & "	<itemDesc><![CDATA["&getSsgContParamToReg()&"]]></itemDesc>"					'#상품 상세 설명
		strRst = strRst & "	<sizeDesc><![CDATA["&FItemsize&"]]></sizeDesc>"									'사이즈 조견표
		strRst = strRst & "	<purchGuideCntt></purchGuideCntt>"												'구매 안내 내용
		strRst = strRst & "	<asMemoCntt></asMemoCntt>"														'AS 메모 내용
'		strRst = strRst & "	<qualityFiles>"
'		strRst = strRst & "		<qualityFile>"
'		strRst = strRst & "			<itemDescDivCd></itemDescDivCd>"										'#품질 검증 파일 구분 코드 (commCd:I045) | 61 원산지증명서, 65 수입신고필증, 63 KC성적서, 64 유기농인증, 65 수입신고필증, 66 광고심의필, 6B 기타
'		strRst = strRst & "			<imgFileNm></imgFileNm>" 												'#이미지 파일 위치
'		strRst = strRst & "		</qualityFile>"
'		strRst = strRst & "	</qualityFiles>"
		strRst = strRst & getCertInfoParam(standardCateCode)
		strRst = strRst & "	<giftPsblYn>Y</giftPsblYn>"														'#선물 가능 여부
		strRst = strRst & "	<shppMsgId></shppMsgId>"														'배송 메시지 ID
		strRst = strRst & "	<ssgstrSellYn></ssgstrSellYn>"													'#SSG 스토어(하남) 판매 여부
		strRst = strRst & "	<vodExtnlPathUrl></vodExtnlPathUrl>"											'동영상 외부 경로 URL (허용 업체에 한하여)
		strRst = strRst & "	<palimpItemYn>N</palimpItemYn>"													'#병행 수입 상품 여부
		strRst = strRst & "	<itemSellWayCd>10</itemSellWayCd>"												'#상품 판매 방식 코드 (commCd:I392) | 10 일반, 20 렌탈, 30 사전 예약, 40 할부,
		strRst = strRst & "	<itemStatTypeCd>10</itemStatTypeCd>"											'#상품 상태 유형 코드 (commCd:I393) | 10 새상품, 20 중고, 30 리퍼, 40 전시, 50 반품, 60 스크래치
		strRst = strRst & "	<whinNotiYn>N</whinNotiYn>"														'#입고 알림 여부
'    <book>		'책관련 필드는 생략..
'    </book>
		strRst = strRst & "	<giftPackPsblYn>N</giftPackPsblYn>"												'선물 포장 가능 여부
		strRst = strRst & "</insertItem>"
		getSsgItemRegParameter = strRst
	End Function

	'SSG 수정 XML
	Public Function getssgItemEditParameter(ichgSellYn)
		Dim strRst, i, sellStatCd, areaCode, shppItemDivCd, shppRqrmDcnt, shppRqrmDcntChngRsnCntt
		If ichgSellYn = "Y" Then
			sellStatCd = 20
		Else
			sellStatCd = 80
		End If
		'################################ 카테고리 항목 호출 ########################################
		Dim callCategory , standardCateCode, arrDisplayCateCode, arrSiteNum
		callCategory = getSsgCategoryParam()
		standardCateCode = Split(callCategory, "|_|")(0)
		arrDisplayCateCode = Split(Split(callCategory, "|_|")(1), ",")
		arrSiteNum = Split(Split(callCategory, "|_|")(2), ",")
		'##########################################################################################
		'################################### 원산지  호출 ##########################################
		areaCode = getSourcearea()
		'##########################################################################################
		'################################### 배송기일  호출 #########################################
		shppRqrmDcnt = getShopLeadTime()
		'##########################################################################################
'		If FItemdiv = "06" OR FItemdiv = "16" Then
'			shppItemDivCd = "05"
'			If FRequireMakeDay < 1 Then
'				shppRqrmDcnt = 7
'			Else
'				shppRqrmDcnt = FRequireMakeDay
'			End If
'			shppRqrmDcntChngRsnCntt = "주문제작상품"
'		Else
'			shppItemDivCd = "01"
'			shppRqrmDcnt = 3
'		End If

		shppItemDivCd = "01"
		If getShopLeadTime > 3 Then
			shppRqrmDcntChngRsnCntt = "주문제작상품"
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst & "<updateItem>"
		strRst = strRst & "	<itemId>"&FSsgGoodno&"</itemId>"												'상품ID
		strRst = strRst & "	<itemNm><![CDATA["&getItemNameFormat&"]]></itemNm>"								'상품명
		strRst = strRst & "	<mdlNm></mdlNm>"																'모델명
		strRst = strRst & "	<deleteMdlNmYn></deleteMdlNmYn>"												'모델명 삭제여부(온라인 상품의 경우만 가능)
		strRst = strRst & "	<brandId>2000047517</brandId>"													'#브랜드ID | 텐바이텐(2000047517)
		strRst = strRst & "	<sites>"
		strRst = strRst & "		<site>"
		strRst = strRst & "			<siteNo>6004</siteNo>"													'#사이트번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
		strRst = strRst & "			<sellStatCd>20</sellStatCd>"											'#판매 상태 코드 | 20 : 판매중, 80 : 일시판매중지
		strRst = strRst & "		</site>"
		strRst = strRst & "		<site>"
		strRst = strRst & "			<siteNo>6001</siteNo>"													'#사이트번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
		strRst = strRst & "			<sellStatCd>20</sellStatCd>"											'#판매 상태 코드 | 20 : 판매중, 80 : 일시판매중지
		strRst = strRst & "		</site>"
		strRst = strRst & "	</sites>"
		strRst = strRst & "	<itemAplRngTypeCd></itemAplRngTypeCd>"											'상품 적용 범위 | 00 : 전체적용, 10 : B2C적용, 20 : B2E적용
		strRst = strRst & "	<b2eAplRngCd>10</b2eAplRngCd>"													'B2E 적용 범위 | 10 : 전체 적용, 20 : 적용 않음, 30 : 회원사 지정
		strRst = strRst & "	<b2cAplRngCd>20</b2cAplRngCd>"													'B2C 적용 범위 | 10 : 적용, 20 : 적용 (대행 제휴사 제외), 70 : 적용 않음
		strRst = strRst & "	<itemChrctDivCd>10</itemChrctDivCd>"											'상품 특성 구분 코드 | 10 : 일반, 40 : 미가공 귀금속, 50 : 모바일 기프트, 60 : 상품권, 70 : 쇼핑 충전금
		strRst = strRst & "	<itemChrctDtlCd></itemChrctDtlCd>"												'상품 특성 상세 코드 | 상품 특성 구분 코드(itemChrctDivCd = 50 | 60) 일 경우 상품 특성 구분 코드(itemChrctDivCd = 50) => 10 : 일반, 50 : 상품권, 상품 특성 구분 코드(itemChrctDivCd = 60) => 60 : 신세계 지류 상품권, 70 : 외부 지류 상품권, 80 : 기프트 카드, 90 : 맞춤형 기프트 카드
		strRst = strRst & "	<exusItemDivCd>10</exusItemDivCd>"												'전용 상품 구분 코드 | 10 : 일반, 20 : GIFT(일반)
		strRst = strRst & "	<exusItemDtlCd>10</exusItemDtlCd>"												'전용 상품 상세 코드 | 10 : 일반, 20 : 특정점
		strRst = strRst & "	<dispAplRngTypeCd>10</dispAplRngTypeCd>"										'전시 적용 범위 유형 코드 | 10 : 전체 (모바일 + PC), 30 : 모바일 (모바일 선택시 전체로 설정 불가)
		strRst = strRst & "	<speSalestrNo></speSalestrNo>"													'특정 영업점 번호 (특정점 (exusItemDtlCd=20)일 경우 입력) | ※ 특정점코드 API 참조
		strRst = strRst & "	<sellStatCd>"&sellStatCd&"</sellStatCd>"										'판매 상태 코드 | 20 : 판매중, 80 : 일시판매중지, 90 : 영구판매중지
		strRst = strRst & getSsgItemInfoCdToReg(areaCode)
		strRst = strRst & "	<manufcoNm><![CDATA["&Trim(FMakername)&"]]></manufcoNm>"						'제조사명
		strRst = strRst & "	<prodManufCntryId>"&areaCode&"</prodManufCntryId>"								'생산 제조 국가 ID | (참고 : 원산지조회API(listOrplc API))
		strRst = strRst & "	<dispCtgs>"
	For i = 0 to Ubound(arrDisplayCateCode)
		strRst = strRst & "		<dispCtg>"
		strRst = strRst & "			<siteNo>"&arrSiteNum(i)&"</siteNo>"										'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰, 6005 : SSG.COM
		strRst = strRst & "			<delYn></delYn>"														'삭제 여부
		strRst = strRst & "			<dispCtgId>"&arrDisplayCateCode(i)&"</dispCtgId>"						'#전시 카테고리 ID
		strRst = strRst & "			<repDispOrdr>"&i+1&"</repDispOrdr>"										'대표 전시 순서 | 순서대로, 중복 허용하지 않음. 사이트별 최대 3개까지 선택 가능
		strRst = strRst & "		</dispCtg>"
	Next
		strRst = strRst & "	</dispCtgs>"
		strRst = strRst & "	<dispStrtDts>"&Replace(Date(), "-", "")&"</dispStrtDts>"						'전시시작일시(YYYYMMDD OR YYYYMMDDHH24MISS)
		strRst = strRst & "	<dispEndDts>29991231</dispEndDts>"												'전시종료일시(YYYYMMDD OR YYYYMMDDHH24MISS)
'		strRst = strRst & "	<spDispCtgs>"																	'-------- MayBe 전문카테고리 인 듯.. --------
'		strRst = strRst & "		<dispCtg>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
'		strRst = strRst & "			<delYn></delYn>"														'삭제 여부
'		strRst = strRst & "			<dispCtgId></dispCtgId>"												'#전시 카테고리 ID
'		strRst = strRst & "			<repDispOrdr></repDispOrdr>"											'대표 전시 순서 | 순서대로, 중복 허용하지 않음. 사이트별 최대 3개까지 선택 가능
'		strRst = strRst & "		</dispCtg>"
'		strRst = strRst & "	</spDispCtgs>"
		strRst = strRst & "	<srchPsblYn>Y</srchPsblYn>"														'검색 가능 여부
		strRst = strRst & "	<itemSrchwdNm><![CDATA["&RightCommaDel(Trim(FKeywords))&"]]></itemSrchwdNm>"	'상품검색어명
		strRst = strRst & "	<aplMbrGrdCd></aplMbrGrdCd>"													'노출 회원 등급 (값이 존재하지 않을 경우 ALL) | 10 : 패밀리, 20 : 브론즈, 30 : 실버, 40 : 골드, 50 : VIP, 90 : VVIP
		strRst = strRst & "	<minOnetOrdPsblQty>1</minOnetOrdPsblQty>"										'최소 1회 주문 가능 수량
		strRst = strRst & "	<maxOnetOrdPsblQty>9999</maxOnetOrdPsblQty>"									'최대 1회 주문 가능 수량
		strRst = strRst & "	<max1dyOrdPsblQty>9999</max1dyOrdPsblQty>"										'최대 1일 주문 가능 수량
		strRst = strRst & "	<adultItemTypeCd>90</adultItemTypeCd>"											'#성인 상품 타입 코드 (commCd:I408) | 10 : 성인 상품, 20 : 주류 상품, 90 : 일반 상품
		strRst = strRst & "	<hriskItemYn>N</hriskItemYn>"													'고 위험 상품 여부
		strRst = strRst & "	<nitmAplYn>N</nitmAplYn>" 														'신 상품 적용 여부
'		strRst = strRst & "	<sellPnts>"																		'-------- MayBe 셀링포인트 인 듯.. --------
'		strRst = strRst & "		<sellPnt>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
'		strRst = strRst & "			<sellpntId></sellpntId>"												'#셀링 포인트 ID
'		strRst = strRst & "			<sellpntNm></sellpntNm>"												'#셀링 포인트 명
'		strRst = strRst & "			<dispStrtDts></dispStrtDts>"											'#전시 시작 일시 (YYYYMMDD)
'		strRst = strRst & "			<dispEndDts></dispEndDts>"												'#전시 종료 일시 (YYYYMMDD)
'		strRst = strRst & "			<useYn></useYn>"														'#사용 여부
'		strRst = strRst & "		</sellPnt>"
'		strRst = strRst & "	</sellPnts>"
		strRst = strRst & "	<sellCapaUnitCd></sellCapaUnitCd>"												'판매 용량 단위 코드 (commCd:I159) | 01 ea, 02 cc, 03 g, 04 kg, 05 m, 06 ml, 07 mm, 08 개, 09 매, 10 포
		strRst = strRst & "	<sellTotCapa></sellTotCapa>"													'판매 총 용량
		strRst = strRst & "	<sellUnitCapa></sellUnitCapa>"													'판매 단위 용량
		strRst = strRst & "	<sellUnitQty>0</sellUnitQty>"													'판매 단위 수량
		strRst = strRst & "	<prcMngMthd>1</prcMngMthd>"														'가격설정방식 | 1 : 공급가 자동계산 (Default), 2 : 판매가 자동계산, 3 : 마진 자동계산..이 값 설정시 SALE_PRC_INFO, B2E_PRC 둘다 적용 받는다. 값은 모두 입력 받아도 상관 없으나 해당 값 설정에 따라 해당 값이 자동으로 계산됨.
		strRst = strRst & "	<salesPrcInfos>"
		strRst = strRst & "		<uitemPrc>"
		strRst = strRst & "			<siteNo>6004</siteNo>"													'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
		strRst = strRst & "			<splprc>"&MustPrice()*0.85&"</splprc>"									'#공급가
		strRst = strRst & "			<sellprc>"&MustPrice()&"</sellprc>"										'#판매가
		strRst = strRst & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"											'#마진율
		strRst = strRst & "		</uitemPrc>"
		strRst = strRst & "		<uitemPrc>"
		strRst = strRst & "			<siteNo>6001</siteNo>"													'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
		strRst = strRst & "			<splprc>"&MustPrice()*0.85&"</splprc>"									'#공급가
		strRst = strRst & "			<sellprc>"&MustPrice()&"</sellprc>"										'#판매가
		strRst = strRst & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"											'#마진율
		strRst = strRst & "		</uitemPrc>"
		strRst = strRst & "	</salesPrcInfos>"
'		strRst = strRst & "	<chgSalesPrcInfos>"
'		strRst = strRst & "		<uitemPrc>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#사이트 번호, 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
'		strRst = strRst & "			<splprc></splprc>"														'#공급가
'		strRst = strRst & "			<sellprc></sellprc>"													'#판매가
'		strRst = strRst & "			<mrgrt></mrgrt>"														'#마진율
'		strRst = strRst & "			<aplStrtDts></aplStrtDts>"												'#적용 시작 일시(YYYYMMDDHH24MISS)
'		strRst = strRst & "		</uitemPrc>"
'		strRst = strRst & "	</chgSalesPrcInfos>"
'		strRst = strRst & "	<returnSalesPrcInfos>"
'		strRst = strRst & "		<uitemPrc>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#사이트 번호, 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
'		strRst = strRst & "			<splprc></splprc>"														'#공급가
'		strRst = strRst & "			<sellprc></sellprc>"													'#판매가
'		strRst = strRst & "			<mrgrt></mrgrt>"														'#마진율
'		strRst = strRst & "			<aplStrtDts></aplStrtDts>"												'#적용 시작 일시(YYYYMMDDHH24MISS)
'		strRst = strRst & "		</uitemPrc>"
'		strRst = strRst & "	</returnSalesPrcInfos>"
'		strRst = strRst & "	<b2ePrcAplTgts>"
'		strRst = strRst & "		<b2ePrcAplTgt>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
'		strRst = strRst & "			<b2eMbrcoId></b2eMbrcoId>"												'#B2E회원사ID
'		strRst = strRst & "			<b2eSplprc></b2eSplprc>"												'#B2E 공급가
'		strRst = strRst & "			<b2eSellprc></b2eSellprc>"												'#B2E 판매가
'		strRst = strRst & "			<b2eMrgrt></b2eMrgrt>"													'#B2E 마진율
'		strRst = strRst & "		</b2ePrcAplTgt>"
'		strRst = strRst & "	</b2ePrcAplTgts>"
		strRst = strRst & "	<invMngYn>Y</invMngYn>"															'재고 관리 여부
		strRst = strRst & "	<baseInvQty>"&getLimitEa()&"</baseInvQty>"										'재고 수량
		strRst = strRst & "	<invQtyMarkgYn>Y</invQtyMarkgYn>"												'재고 수량 표기 여부
'		strRst = strRst & "	<rsvSaleInfo>"
'		strRst = strRst & "		<aplStrtDt></aplStrtDt>"													'#예약판매 시작일 (YYYYMMDD)
'		strRst = strRst & "		<aplEndDt></aplEndDt>"														'#예약판매 종료일 (YYYYMMDD)
'		strRst = strRst & "		<whoutStrtDt></whoutStrtDt>"												'#출고 시작 일자 (YYYYMMDD)
'		strRst = strRst & "		<rstctInvQty></rstctInvQty>"												'#예약 판매 수량
'		strRst = strRst & "		<rsvSaleEndTp></rsvSaleEndTp>"												'#예약 판매 종료(Y로 입력시 예약판매 강제 종료)
'		strRst = strRst & "	</rsvSaleInfo>"
'		If ichgSellYn = "Y" Then	'품절에 해당하지 않을 때만 옵션 수정하기..
			strRst = strRst & getSsgOptParamtoEDIT()
'		End If
		strRst = strRst & "	<shppItemDivCd>"&shppItemDivCd&"</shppItemDivCd>"								'배송상품구분코드 (commCd:I070) | 01 : 일반, 02 : 해외구매대행, 03 : 설치(유료), 04 : 설치(무료), 05 : 주문제작, 06 : 해외직배송
		strRst = strRst & "	<exprtCntryId></exprtCntryId>"													'수출국가(해외 직배송 적출국)shppItemDivCd=06(해외직배송) 인 경우 필수 | 원산지 조회 API 참고(listOrplc API)
		strRst = strRst & "	<pcusMngCd></pcusMngCd>"														'개인 통관 고유 부호 | shppItemDivCd=06(해외직배송) 인 경우 필수 10 : 선택 입력, 20 : 필수 입력, 30 : 입력 안함
		strRst = strRst & "	<retExchPsblYn>Y</retExchPsblYn>"												'반품 교환 가능 여부
		strRst = strRst & "	<shppMainCd>41</shppMainCd>"													'배송 주체 코드 (commCd:P017) | 31 : 자사창고, 32 : 업체창고, 41 : 협력업체
		strRst = strRst & "	<shppMthdCd>20</shppMthdCd>"													'배송 방법 코드 (commCd:P021) | 10 : 자사배송, 20 : 택배배송, 30 : 매장방문, 40 : 등기, 50 : 미배송, 60 : 미발송
		strRst = strRst & "	<mareaShppYn></mareaShppYn>"													'수도권 배송여부
		strRst = strRst & "	<shppRqrmDcnt>"&shppRqrmDcnt&"</shppRqrmDcnt>"									'배송 소요 일수
		strRst = strRst & "	<shppRqrmDcntChngRsnCntt>"&shppRqrmDcntChngRsnCntt&"</shppRqrmDcntChngRsnCntt>"	'배송 소요 일수 변경 사유 | 상품배송구분이 일반(01) 이고 배송소요일수가 4일 이상일 경우 필수
		strRst = strRst & "	<splVenItemId>"&FItemid&"</splVenItemId>"										'업체 상품 번호
		strRst = strRst & "	<whoutShppcstId>0000517199</whoutShppcstId>"									'출고 배송비 ID
		strRst = strRst & "	<retShppcstId>0000011336</retShppcstId>"										'반품 배송비 ID
		strRst = strRst & "	<whoutAddrId>0000006297</whoutAddrId>"											'출고 주소 ID
		strRst = strRst & "	<snbkAddrId>0000006297</snbkAddrId>"											'반품 주소 ID
		strRst = strRst & "	<frgShppPsblYn>N</frgShppPsblYn>"												'해외 배송 가능 여부
		strRst = strRst & "	<itemTotWgt></itemTotWgt>"												  		'상품 총 무게
		strRst = strRst & "	<hopeShppDdDivCd></hopeShppDdDivCd>"											'희망 발송일 구분 코드 (commCd:I015) | 10 : 15일이내, 20 : 15일이후 30일이내, 30 : 30일이후, 90 : 발송일 최대 날짜 지정
		strRst = strRst & "	<hopeShppDdEndDts></hopeShppDdEndDts>"											'희망 발송일 종료 일시 (YYYYMMDD) | 희망발송일 구분코드가 (hopeShppDdEndDts=90) 일경우 필수
		If isImageChanged Then
			strRst = strRst & getSsgAddImageParam()
		End If
		strRst = strRst & "	<itemDesc><![CDATA["&getSsgContParamToReg()&"]]></itemDesc>"					'상품 상세 설명
		strRst = strRst & "	<sizeDesc><![CDATA["&FItemsize&"]]></sizeDesc>"									'사이즈 조견표
		strRst = strRst & "	<purchGuideCntt></purchGuideCntt>"												'구매 안내 내용
		strRst = strRst & "	<asMemoCntt></asMemoCntt>"														'AS 메모 내용
'		strRst = strRst & "	<qualityFiles>"
'		strRst = strRst & "		<qualityFile>"
'		strRst = strRst & "			<itemDescDivCd></itemDescDivCd>"										'#품질 검증 파일 구분 코드 (commCd:I045) | 61 원산지증명서, 65 수입신고필증, 63 KC성적서, 64 유기농인증, 65 수입신고필증, 66 광고심의필, 6B 기타
'		strRst = strRst & "			<imgFileNm></imgFileNm>" 												'#이미지 파일 위치
'		strRst = strRst & "		</qualityFile>"
'		strRst = strRst & "	</qualityFiles>"
		strRst = strRst & getCertInfoParam(standardCateCode)
		strRst = strRst & "	<giftPsblYn>Y</giftPsblYn>"														'선물 가능 여부
		strRst = strRst & "	<shppMsgId></shppMsgId>"														'배송 메시지 ID
		strRst = strRst & "	<ssgstrSellYn></ssgstrSellYn>"													'SSG 스토어(하남) 판매 여부
		strRst = strRst & "	<vodExtnlPathUrl></vodExtnlPathUrl>"											'동영상 외부 경로 URL (허용 업체에 한하여)
		strRst = strRst & "	<palimpItemYn>N</palimpItemYn>"													'병행 수입 상품 여부
		strRst = strRst & "	<itemSellWayCd>10</itemSellWayCd>"												'상품 판매 방식 코드 (commCd:I392) | 10 일반, 20 렌탈, 30 사전 예약, 40 할부,
		strRst = strRst & "	<itemStatTypeCd>10</itemStatTypeCd>"											'상품 상태 유형 코드 (commCd:I393) | 10 새상품, 20 중고, 30 리퍼, 40 전시, 50 반품, 60 스크래치
		strRst = strRst & "	<whinNotiYn>N</whinNotiYn>"														'입고 알림 여부
'    <book>		'책관련 필드는 생략..
'    </book>
		strRst = strRst & "	<giftPackPsblYn>N</giftPackPsblYn>"												'선물 포장 가능 여부
		strRst = strRst & "</updateItem>"
		getssgItemEditParameter = strRst
	End Function

End Class

Class CSsg
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectItemID
	Public FRectMustSellyn

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	'// 미등록 상품 목록(등록용)
	Public Sub getSsgNotRegOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & "	SELECT itemid FROM ("
            addSql = addSql & "     SELECT itemid"
            addSql = addSql & " 	,count(*) as optCNT"
			addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	FROM db_item.dbo.tbl_item_option"
            addSql = addSql & " 	WHERE itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and isusing='Y'"
            addSql = addSql & " 	GROUP BY itemid"
            addSql = addSql & " ) T"
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, isNULL(C.safetyNum, '') as safetyNum "
		strSql = strSql & "	, isNULL(R.ssgStatCD,-9) as ssgStatCD, cm.mapCnt, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	, UC.socname_kor, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_ssg_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ssg_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " WHERE i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		ELSE
		    strSql = strSql & " and (i.deliveryType<>9)"
	    END IF
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.sellcash > i.buycash "
		strSql = strSql & " and i.itemdiv not in ('08', '09', '21') "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "						'플라워/화물배송
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)>=" & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and isnull(R.ssgGoodNo, '') = '' "
		strSql = strSql & " and cm.mapCnt is Not Null "
'		strSql = strSql & " and (i.mwdiv='M' or i.mwdiv='W') "		'매입 or 위탁
'		strSql = strSql & " and i.deliveryType = 1 "				'탠배
'2018-01-29 15:00 김진영 하단 주석처리..
'		strSql = strSql & " and ( ((i.mwdiv='M' or i.mwdiv='W') and (i.deliveryType = 1)) OR (i.makerid in ('meaningless01', 'mandarinebrothers', 'fromamour', 'woolly02' ,'dalbampicnic')) ) "
		strSql = strSql & addSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CSsgItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FtenCateLarge		= rsget("cate_large")
				FOneItem.FtenCateMid		= rsget("cate_mid")
				FOneItem.FtenCateSmall		= rsget("cate_small")
				FOneItem.Fitemname			= db2html(rsget("itemname"))
				FOneItem.FitemDiv			= rsget("itemdiv")
				FOneItem.FsmallImage		= rsget("smallImage")
				FOneItem.Fmakerid			= rsget("makerid")
				FOneItem.Fregdate			= rsget("regdate")
				FOneItem.FlastUpdate		= rsget("lastUpdate")
				FOneItem.ForgPrice			= rsget("orgPrice")
				FOneItem.ForgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FsellYn			= rsget("sellYn")
				FOneItem.FsaleYn			= rsget("sailyn")
				FOneItem.FisUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.Fkeywords			= rsget("keywords")
				FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.ForderComment		= db2html(rsget("ordercomment"))
				FOneItem.FoptionCnt			= rsget("optionCnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.FListimage			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetyDiv			= rsget("safetyDiv")
				FOneItem.FSsgStatCD			= rsget("ssgStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.Fsocname_kor		= rsget("socname_kor")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FMapCnt 			= rsget("mapCnt")
				FOneItem.FMwdiv 			= rsget("mwdiv")
				FOneItem.FItemsize 			= rsget("itemsize")
				FOneItem.FItemsource 		= rsget("itemsource")
				FOneItem.FRequireMakeDay 	= rsget("requireMakeDay")
		End If
		rsget.Close
	End Sub

	Public Sub getSsgEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		If FRectMustSellyn <> "Y" Then
	        ''//연동 제외상품
	        addSql = addSql & " and i.itemid not in ("
	        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
	        addSql = addSql & "     where stDt < getdate()"
	        addSql = addSql & "     and edDt > getdate()"
	        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
	        addSql = addSql & "     and linkgbn='donotEdit'"
	        addSql = addSql & " )"
		End If
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, isNULL(C.safetyNum, '') as safetyNum "
		strSql = strSql & "	, isNULL(m.ssgStatCD,-9) as ssgStatCD, cm.mapCnt, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	, UC.socname_kor, isNULL(c.requireMakeDay,0) as requireMakeDay, m.ssgGoodNo "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv = '21' "
'		strSql = strSql & " 	or i.mwdiv not in ('M', 'W') "
'		strSql = strSql & " 	or i.deliveryType <> 1 "
'2018-01-29 15:00 김진영 하단 주석처리..
'		strSql = strSql & "		or ( ((i.mwdiv not in ('M', 'W')) or (i.deliveryType <> 1)) and i.makerid not in ('meaningless01', 'mandarinebrothers', 'fromamour', 'woolly02' ,'dalbampicnic') )"
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.itemdiv = '09' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " LEFT JOIN ( "
		strSql = strSql & " 	SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & " 	FROM db_etcmall.dbo.tbl_ssg_cate_mapping "
		strSql = strSql & " 	GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_ssg_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.ssgGoodNo is Not Null "		'등록 상품만
		'strSql = strSql & " and m.ssgStatCD = 7' "				'승인완료된 애들만 수정이 된다함..TEST 해봐야 됨
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CSsgItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FtenCateLarge		= rsget("cate_large")
				FOneItem.FtenCateMid		= rsget("cate_mid")
				FOneItem.FtenCateSmall		= rsget("cate_small")
				FOneItem.Fitemname			= db2html(rsget("itemname"))
				FOneItem.FitemDiv			= rsget("itemdiv")
				FOneItem.FsmallImage		= rsget("smallImage")
				FOneItem.Fmakerid			= rsget("makerid")
				FOneItem.Fregdate			= rsget("regdate")
				FOneItem.FlastUpdate		= rsget("lastUpdate")
				FOneItem.ForgPrice			= rsget("orgPrice")
				FOneItem.ForgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FsellYn			= rsget("sellYn")
				FOneItem.FsaleYn			= rsget("sailyn")
				FOneItem.FisUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.Fkeywords			= rsget("keywords")
				FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.ForderComment		= db2html(rsget("ordercomment"))
				FOneItem.FoptionCnt			= rsget("optionCnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.FListimage			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetyDiv			= rsget("safetyDiv")
				FOneItem.FSsgStatCD			= rsget("ssgStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.Fsocname_kor		= rsget("socname_kor")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FMapCnt 			= rsget("mapCnt")
				FOneItem.FMwdiv 			= rsget("mwdiv")
				FOneItem.FItemsize 			= rsget("itemsize")
				FOneItem.FItemsource 		= rsget("itemsource")
				FOneItem.FRequireMakeDay 	= rsget("requireMakeDay")
				FOneItem.FmaySoldOut		= rsget("maySoldOut")
				FOneItem.FSsgGoodno			= rsget("ssgGoodno")
		End If
		rsget.Close
	End Sub
End Class

'SSG 상품코드 얻기
Function getSsgGoodNo(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 ssgGoodNo FROM db_etcmall.dbo.tbl_ssg_regitem WHERE itemid = '"&iitemid&"' "
	rsget.Open strSql, dbget, 1
		getSsgGoodNo = rsget("ssgGoodNo")
	rsget.Close
End Function

'// 상품이미지 존재여부 검사
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function

Public Function GetRaiseValue(value)
    If Fix(value) < value Then
    	GetRaiseValue = Fix(value) + 1
    Else
    	GetRaiseValue = Fix(value)
    End If
End Function

function replaceRst(checkvalue)
	dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, "&", "&amp;")
    v = replace(v, """", "&quot;")
	'v = Replace(v,"<br>","&#xA;")
	'v = Replace(v,"</br>","&#xA;")
	'v = Replace(v,"<br />","&#xA;")
	v = Replace(v,"<","&lt;")
	v = Replace(v,">","&gt;")
    replaceRst = v
end function

function replaceMsg(v)
	if IsNull(v) then
		replaceMsg = ""
		Exit function
	end if
	v = Replace(v, vbcrlf,"")
	v = Replace(v, vbCr,"")
	v = Replace(v, vbLf,"")
    replaceMsg = v
end function
%>