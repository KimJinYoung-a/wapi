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
	Public FAdultType
	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FmaySoldOut
	Public FSsgGoodno
	Public FSsgPrice
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
	Public FOrderMaxNum

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "999999" Then
			getOrderMaxNum = 999999
		End If
	End Function

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

	public Function getKeywords()
		Dim strRst
		strRst = FKeywords
		strRst = replace(strRst, "인기", "")
		strRst = replace(strRst, "인치", "")
		strRst = replace(strRst, "모기퇴치", "")
		If strRst = "" Then
			strRst = "텐바이텐"
		End If
		getKeywords = strRst
	End Function

	'// 품절여부
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold <= CMAXLIMITSELL))
	end function

	Public Function isUpchePriceSoldout
		If Fdeliverytype = "9" and MustPrice < 10000 Then
			isUpchePriceSoldout = "Y"
		Else
			isUpchePriceSoldout = "N"
		End If
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, sqlStr, specialPrice, tmpPrice, vBigPrice, vSmallPrice
		Dim ownItemCnt
		sqlStr = ""
		sqlStr = sqlStr & " SELECT mustPrice "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] "
		sqlStr = sqlStr & " WHERE mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " and itemid = '"& Fitemid &"' "
		sqlStr = sqlStr & " and getdate() >= startDate and getdate() <= endDate "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			specialPrice = rsget("mustPrice")
		End If
		rsget.Close

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as CNT "
		sqlStr = sqlStr & " FROM db_partner.dbo.tbl_partner "
		sqlStr = sqlStr & " WHERE purchaseType in ('3','5','6') "		'3 : PB, 5 : ODM, 6 : 수입
		sqlStr = sqlStr & " and id = '"& FMakerId &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			ownItemCnt = rsget("CNT")
		End If
		rsget.Close

		If specialPrice <> "" Then
			MustPrice = specialPrice
		ElseIf ownItemCnt > 0 Then
			MustPrice = Forgprice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If FSsgprice = 0 Then
				If GetTenTenMargin < CMAXMARGIN Then
					MustPrice = Forgprice
				Else
					MustPrice = FSellCash
				End If
			Else
				If GetTenTenMargin < CMAXMARGIN Then
					MustPrice = Forgprice
				Else
					If (FSellCash < Round(FSsgprice * 0.25, 0)) Then
						MustPrice = CStr(GetRaiseValue(Round(FSsgprice * 0.25, 0)/10)*10)
					Else
						MustPrice = CStr(GetRaiseValue(FSellCash/10)*10)
					End If
				End If
			End If
		End If
	End Function

    Public Function getSSGMargin()
    	Dim standardCode, strSql, isCategoryMarginExist, isItemMarginExists
		Dim cateMargin, itemMargin
		strSql = ""
		strSql = strSql & " SELECT TOP 1 m.margin "
		strSql = strSql & " from db_etcmall.[dbo].[tbl_ssg_marginCate_master] as m "
		strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_ssg_marginCate_detail] as d on m.idx = d.midx "
		strSql = strSql & " WHERE m.isusing = 'Y' "
		strSql = strSql & " and convert(char(10), getdate(), 120) between m.startDate and m.enddate "
		strSql = strSql & " and d.cdl = '"& FtenCateLarge &"' "
		strSql = strSql & " and d.cdm = '"& FtenCateMid &"' "
		strSql = strSql & " and m.mallid = '"&CMALLNAME&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			cateMargin = rsget("margin")
			isCategoryMarginExist = "Y"
		End If
		rsget.Close

		strSql = ""
		strSql = strSql & " SELECT TOP 1 m.margin "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_ssg_marginItem_master as m "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_ssg_marginItem_detail as d on m.idx = d.midx "
		strSql = strSql & " WHERE m.isusing = 'Y' "
		strSql = strSql & " and CONVERT(char(10), getdate(), 12) between m.startDate and m.endDate "
		strSql = strSql & " and d.itemid = '"& FItemid &"' "
		strSql = strSql & " and m.mallid = '"&CMALLNAME&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			itemMargin = rsget("margin")
			isItemMarginExists = "Y"
		End If
		rsget.Close

		'매입마진 변경된 카테고리와 상품이 중복된다면 상품이 우선 순위가 높음..
		If isItemMarginExists = "Y" Then
			getSSGMargin = itemMargin
		ElseIf isCategoryMarginExist = "Y" Then
			getSSGMargin = cateMargin
		Else
			getSSGMargin = SSGMARGIN
		End If


    	' standardCode = Split(getSsgCategoryParam(), "|_|")(0)
		' If FtenCateLarge = "055" AND (Now() > #10/05/2018 00:00:00# AND Now() < #10/15/2018 23:59:59#) Then
		' 	getSSGMargin = 15
		' ElseIf FtenCateLarge <> "055" AND (Now() > #09/28/2018 00:00:00# AND Now() < #10/08/2018 23:59:59#) Then
		' 	If getMarginChgCategory = "Y" Then
		' 		getSSGMargin = 17
		' 	Else
		' 		getSSGMargin = SSGMARGIN
		' 	End If
		' ElseIf (Now() > #08/31/2018 00:00:00# AND Now() < #09/10/2018 23:59:59#) Then
		' 	getSSGMargin = getMarginChgCategory2(standardCode)
		' Else
		' 	getSSGMargin = SSGMARGIN
		' End If
    End Function

	Public Function getMarginChgCategory
		Select Case FtenCateLarge
			'Case "010","020","030","035","040","045","050","055","060"
			'Case "040", "055"			'가구, 패브릭
			Case "010", "020", "030", "035"		'디자인문구, 오피스/개인소품, 키덜트, 여행/취미
				getMarginChgCategory = "Y"
			Case Else
				getMarginChgCategory = "N"
		End Select
	End Function

	Public Function getMarginChgCategory2(standardCode)
		Select Case standardCode
			Case "1000020917", "1000020919", "1000020920", "1000020921", "1000020922", "1000020923", "1000021518", "1000021519", "1000021520", "1000021521", "1000021539", "1000021937", "1000022354", "1000022607", "4000002193", "4000002194", "4000002195", "4000002196", "4000002197", "4000002198", "4000002200", "4000002201", "4000002202", "4000002203", "4000002204", "4000002205", "4000002206", "4000002207", "4000002209", "4000002210", "4000002216", "4000002218", "4000002221", "4000002223", "4000002230", "4000002232", "4000002235", "4000002236", "4000002239", "4000002244", "4000002249", "4000002251", "4000002257", "4000002259", "4000002264", "4000002265", "4000002266", "4000002267", "4000002268", "4000002271", "4000002273", "4000002274", "4000002276", "4000002277", "4000002278", "4000002279", "4000002280", "4000002282", "4000002285", "4000002288", "4000002289", "4000002290", "4000002291", "4000002292", "4000002294", "4000002295", "4000002320", "4000002321", "4000002323", "4000002340", "4000002352", "4000002360", "4000002372", "4000002377", "4000002935", "4000002937", "4000002938", "4000002942", "4000002944", "4000002947", "4000002949", "4000002950"	getMarginChgCategory2 = "17"
			'Case "4000002194", "4000002193", "4000002195", "4000002210", "4000002249", "4000002277", "4000002276", "4000002285", "4000002278", "4000002280", "4000002279", "4000002274", "4000002202", "4000002230", "4000002266", "4000002206", "4000002268", "4000002271", "4000002196", "4000002288", "4000002290", "4000002294", "4000002292", "4000002291", "4000002282", "4000002295", "4000002289", "4000002251", "4000002265", "4000002207", "4000002244", "4000002197", "4000002232", "4000002205", "4000002239", "4000002201", "4000002221", "4000002218", "4000002203", "4000002223", "4000002264", "4000002259", "4000002198", "4000002235", "4000002216", "4000002257", "4000002236", "4000002204", "4000002209", "4000002200", "1000020917", "1000021539", "1000020919", "1000020920", "1000020921", "1000020923", "1000020922", "1000021937", "4000002321", "4000002323", "4000002352", "4000002320", "4000002340", "4000002372", "4000002360", "4000002377", "1000021521", "4000002949", "1000021520", "4000002947", "4000002942", "4000002938", "4000002937", "1000021518", "4000002950", "4000002944", "1000021519", "4000002935", "1000022607", "4000002273", "4000002267", "1000022354", "4000001787", "4000001786", "4000001815", "4000001804", "4000001807", "4000001806", "4000001805", "4000001801", "4000001812", "4000001802", "4000001800", "4000001793", "4000001756", "4000001811", "4000001778", "4000001785", "4000001779", "4000001783", "4000001780", "4000001781", "4000001784", "4000001797", "4000001798", "4000001799"	getMarginChgCategory2 = "15"
			Case Else	getMarginChgCategory2 = SSGMARGIN
		End Select
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
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
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
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

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
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
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
		Dim i, strSql, retId
		strSql = ""
		strSql = strSql & " SELECT TOP 1 id "
		'strSql = strSql & " FROM db_etcmall.dbo.tbl_ssg_sourceAreaCode "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_ssg_sourceAreaCodeMapping "
		strSql = strSql & " WHERE sourcearea = '"&Trim(FSourcearea)&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			retId = rsget("id")
		End If
		rsget.Close

		If retId = "" Then
			getSourcearea = "1000000000"		''상세설명참조
'			getSourcearea = "2000000078"		'2022-11-28 김진영 / 상세설명참조 코드 사라짐..없으면 2000000078(한국)으로 바꿈..대신 상세설명에 원산지명 표기
		Else
			getSourcearea = retId
		End If
	End Function

	Public Function IsAdultItem()
		Select Case FAdultType
			Case "1", "2"
				IsAdultItem = "Y"
			Case Else
				IsAdultItem = "N"
		End Select
	End Function

	Public Function getShopLeadTime()
		Dim CateLargeMid, leadTime
		If isNull(FtenCateLarge) AND isNull(FtenCateMid) Then
			FtenCateLarge = "999"
			FtenCateMid = "999"
		End If

		CateLargeMid = CStr(FtenCateLarge) & CStr(FtenCateMid)
		Select Case CateLargeMid
			Case "030331", "040010", "040011", "040020", "040030", "040040", "040050", "040070", "040080", "040090", "040100", "040121", "055070", "055080"
				leadTime = 15
			Case "050777", "055090", "055100", "055110", "055120", "060070"
				leadTime = 10
			Case "050045", "080007", "080010", "080020", "080030", "080031", "080040", "080050", "080051", "080060", "080070", "080071", "080080", "080090", "090005", "090010", "090011", "090020", "090030", "090040", "090050"
				leadTime = 7
			Case "010130", "010140", "010150", "010160", "020001", "020010", "020020", "020030", "020060", "020070", "020090", "020100", "020110", "020111", "020120", "020130", "020222", "020333", "020334", "025014", "025015", "025020", "025022", "025030", "025040", "025050", "025060", "025070", "025080", "025100", "025101", "025102", "025103", "025104", "025105", "025106", "025107", "025108", "025109", "025110", "025111", "025112", "025113", "025114", "025115", "025116", "025117", "025118", "025120", "025456", "030020", "030300", "030320", "030330", "030340", "030345", "030350", "030360", "030370", "030380", "030420", "030421", "030450", "035009", "035010", "035011", "035012", "035013", "035014", "035015", "035016", "035017", "035018", "035019", "035020", "035021", "035022", "045001", "045002", "045003", "045004", "045005", "045006", "045007", "045008", "045009", "045010", "045011", "050010", "050020", "050030", "050040", "050050", "050070", "050110", "050120", "050666", "055222", "060010", "060020", "060040", "060050", "060060", "060080", "060090", "060120", "060130", "060140", "060150", "060160", "070010", "070020", "070030", "070040", "070050", "070070", "070110", "070120", "070140", "070150", "070160", "070200", "070201", "070202", "070203", "090060", "090061", "090070", "090071", "090080"
				leadTime = 5
			Case Else
				leadTime = 3
		End Select
		getShopLeadTime = leadTime
	End Function

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	Public Function getSsgContParamToReg()
		Dim strRst, strSQL, textVal
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '돋움','돋움' }</style><br>"
		strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_ssg.jpg'></p><br>"

		If ForderComment <> "" Then
			Fordercomment = replace(Fordercomment, "할인", "할 인")
			strRst = strRst & "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
		End If

		If FSourcearea <> "" Then
			strRst = strRst & "- 원산지 : " &  FSourcearea & "<br>"
		End If

		If Fitemsource <> "" Then
			strRst = strRst & "- 재료 : " &  Fitemsource & "<br>"
		End If
		strRst = strRst & Replace(Replace(Replace(Replace(Replace(FItemContent,"",""),"",""), "=""/", ""),"조기품절",""), "주문폭주", "")

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
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ ><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ ><br>")

		'#배송 주의사항
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ssg.jpg"">")
		getSsgContParamToReg = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		if Not(rsget.EOF or rsget.BOF) then
			textVal = rsget("textVal")
			strRst = ""
			strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '돋움','돋움' }</style><br>"
			strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_ssg.jpg'></p><br>"

			If ForderComment <> "" Then
				strRst = strRst & "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
			End If

			If Fitemsource <> "" Then
				strRst = strRst & "- 재료 : " &  Fitemsource & "<br>"
			End If
			strRst = strRst & textVal & "<br/><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ssg.jpg"">"
			getSsgContParamToReg = strRst
		End If
		rsget.Close
	End Function

    public function isImageChanged()
        Dim ibuf : ibuf = getBasicImage
        ' if InStr(ibuf,"-")<1 then
        '     isImageChanged = FALSE
        '     Exit function
        ' end if
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
		Dim arrRows, notarrRows
		Dim childStrRst, childIntoStrRst, harmStrRst, safeStrRst, elecStrRst
		strSql = ""
		strSql = strSql & " SELECT TOP 1 chldCertTgtYn, safeCertTgtYn, elecCertTgtYn, harmCertTgtYn "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_ssg_mmg_category] "
		strSql = strSql & " WHERE stdCtgDclsId = '"&iCode&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			isChild	= rsget("chldCertTgtYn")
			isSafe	= rsget("safeCertTgtYn")
			isElec	= rsget("elecCertTgtYn")
			isHarm	= rsget("harmCertTgtYn")
		End If
		rsget.Close

		If isChild = "Y" Then
			strSql = ""
			strSql = strSql & " SELECT TOP 5 certNum, safetyDiv " & vbcrlf
			strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg " & vbcrlf
			strSql = strSql & " WHERE itemid='"&FItemID&"' " & vbcrlf
			strSql = strSql & " AND safetyDiv in ('70', '80', '90') "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				Select Case rsget("safetyDiv")
					Case "70"	chldCertDivCd = "10"
					Case "80"	chldCertDivCd = "20"
					Case "90"	chldCertDivCd = "30"
				End Select

				childStrRst = ""
				childStrRst = childStrRst & "	<chldCert>"
				childStrRst = childStrRst & "		<chldCertYn>Y</chldCertYn>" 			'#어린이인증 여부
				childStrRst = childStrRst & "		<chldCertDivCd>"&chldCertDivCd&"</chldCertDivCd>"  	'어린이인증 구분 (commCd:I368) | (어린이인증 여부가 Y일 경우에만 필수) 10 : 안전인증대상, 20 : 안전확인대상, 30 : 공급자적합성확인
				childStrRst = childStrRst & "		<chldCertNo>"&rsget("certNum")&"</chldCertNo>" 			'인증번호 | 어린이인증 구분이 10, 20 일경우에만 필수-
				childStrRst = childStrRst & "	</chldCert>"

				childIntoStrRst = ""
				childIntoStrRst = childIntoStrRst & "		<certInfo>"
				childIntoStrRst = childIntoStrRst & "			<certKind>6000000001</certKind>"			'#인증종류 (commCd:I387) | 인증대상 카테고리 일 경우 필수..6000000001 : 어린이인증 대상여부, 6000000002 : 안전인증 대상여부, 6000000003 : 전파인증 적합성평가 대상여부, 6000000004 : 위해우려제품 표시대상여부
				childIntoStrRst = childIntoStrRst & "			<certYn>Y</certYn>"							'#인증 여부
				childIntoStrRst = childIntoStrRst & "			<certDivCd>"&chldCertDivCd&"</certDivCd>"	'인증 구분 (commCd:I368) | 인증여부가 Y이고 인증종류가 (certKind=6000000001 | 6000000002) 일 경우 필수..10 : 안전인증대상, 20 : 안전확인대상, 30 : 공급자적합성확인
				childIntoStrRst = childIntoStrRst & "			<certNo>"&rsget("certNum")&"</certNo>"		'인증번호 | 인증 구분이 10, 20 일경우에만 필수-
				childIntoStrRst = childIntoStrRst & "		</certInfo>"
			Else
				' If (FSafetyyn = "Y") And (FsafetyDiv = "50") Then
				' 	chldCertYn		= "Y"
				' 	chldCertDivCd	= "10"
				' 	chldCertNo		= FSafetyNum
				' Else
				' 	chldCertYn		= "N"
				' End If
				' 2019-01-16 김진영 수정..위 조건 삭제 후 전부 인증없음 처리
				chldCertYn = "N"
				chldCertDivCd = ""
				chldCertNo = ""

				childStrRst = ""
				childStrRst = childStrRst & "	<chldCert>"
				childStrRst = childStrRst & "		<chldCertYn>"&chldCertYn&"</chldCertYn>" 			'#어린이인증 여부
				childStrRst = childStrRst & "		<chldCertDivCd>"&chldCertDivCd&"</chldCertDivCd>"  	'어린이인증 구분 (commCd:I368) | (어린이인증 여부가 Y일 경우에만 필수) 10 : 안전인증대상, 20 : 안전확인대상, 30 : 공급자적합성확인
				childStrRst = childStrRst & "		<chldCertNo>"&chldCertNo&"</chldCertNo>" 			'인증번호 | 어린이인증 구분이 10, 20 일경우에만 필수-
				childStrRst = childStrRst & "	</chldCert>"

				childIntoStrRst = ""
				childIntoStrRst = childIntoStrRst & "		<certInfo>"
				childIntoStrRst = childIntoStrRst & "			<certKind>6000000001</certKind>"			'#인증종류 (commCd:I387) | 인증대상 카테고리 일 경우 필수..6000000001 : 어린이인증 대상여부, 6000000002 : 안전인증 대상여부, 6000000003 : 전파인증 적합성평가 대상여부, 6000000004 : 위해우려제품 표시대상여부
				childIntoStrRst = childIntoStrRst & "			<certYn>"&chldCertYn&"</certYn>"			'#인증 여부
				childIntoStrRst = childIntoStrRst & "			<certDivCd>"&chldCertDivCd&"</certDivCd>"	'인증 구분 (commCd:I368) | 인증여부가 Y이고 인증종류가 (certKind=6000000001 | 6000000002) 일 경우 필수..10 : 안전인증대상, 20 : 안전확인대상, 30 : 공급자적합성확인
				childIntoStrRst = childIntoStrRst & "			<certNo>"&chldCertNo&"</certNo>"			'인증번호 | 인증 구분이 10, 20 일경우에만 필수-
				childIntoStrRst = childIntoStrRst & "		</certInfo>"
			End If
			rsget.Close
		End If

		If isHarm = "Y" Then
			harmStrRst = ""
			harmStrRst = harmStrRst & "		<certInfo>"
			harmStrRst = harmStrRst & "			<certKind>6000000004</certKind>"			'#인증종류 (commCd:I387) | 인증대상 카테고리 일 경우 필수..6000000001 : 어린이인증 대상여부, 6000000002 : 안전인증 대상여부, 6000000003 : 전파인증 적합성평가 대상여부, 6000000004 : 위해우려제품 표시대상여부
			harmStrRst = harmStrRst & "			<certYn>N</certYn>"							'#인증 여부
			harmStrRst = harmStrRst & "			<certDivCd></certDivCd>"					'인증 구분 (commCd:I368) | 인증여부가 Y이고 인증종류가 (certKind=6000000001 | 6000000002) 일 경우 필수..10 : 안전인증대상, 20 : 안전확인대상, 30 : 공급자적합성확인
			harmStrRst = harmStrRst & "			<certNo></certNo>"							'인증번호 | 인증 구분이 10, 20 일경우에만 필수-
			harmStrRst = harmStrRst & "		</certInfo>"
		End If

		If isSafe = "Y" Then
			strSql = ""
			strSql = strSql & " SELECT TOP 1 certNum, safetyDiv " & vbcrlf
			strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg " & vbcrlf
			strSql = strSql & " WHERE itemid='"&FItemID&"' " & vbcrlf
			strSql = strSql & " AND safetyDiv not in ('70', '80', '90') "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				Select Case rsget("safetyDiv")
					Case "10", "40"		certKind = "10"
					Case "20", "50"		certKind = "20"
					Case "30", "60"		certKind = "30"
				End Select

				safeStrRst = ""
				safeStrRst = safeStrRst & "		<certInfo>"
				safeStrRst = safeStrRst & "			<certKind>6000000002</certKind>"				'#인증종류 (commCd:I387) | 인증대상 카테고리 일 경우 필수..6000000001 : 어린이인증 대상여부, 6000000002 : 안전인증 대상여부, 6000000003 : 전파인증 적합성평가 대상여부, 6000000004 : 위해우려제품 표시대상여부
				safeStrRst = safeStrRst & "			<certYn>Y</certYn>"								'#인증 여부
				safeStrRst = safeStrRst & "			<certDivCd>"&certKind&"</certDivCd>"			'인증 구분 (commCd:I368) | 인증여부가 Y이고 인증종류가 (certKind=6000000001 | 6000000002) 일 경우 필수..10 : 안전인증대상, 20 : 안전확인대상, 30 : 공급자적합성확인
				safeStrRst = safeStrRst & "			<certNo>"&rsget("certNum")&"</certNo>"			'인증번호 | 인증 구분이 10, 20 일경우에만 필수-
				safeStrRst = safeStrRst & "		</certInfo>"
			Else
				safeStrRst = ""
				safeStrRst = safeStrRst & "		<certInfo>"
				safeStrRst = safeStrRst & "			<certKind>6000000002</certKind>"			'#인증종류 (commCd:I387) | 인증대상 카테고리 일 경우 필수..6000000001 : 어린이인증 대상여부, 6000000002 : 안전인증 대상여부, 6000000003 : 전파인증 적합성평가 대상여부, 6000000004 : 위해우려제품 표시대상여부
				safeStrRst = safeStrRst & "			<certYn>N</certYn>"							'#인증 여부
				safeStrRst = safeStrRst & "			<certDivCd></certDivCd>"					'인증 구분 (commCd:I368) | 인증여부가 Y이고 인증종류가 (certKind=6000000001 | 6000000002) 일 경우 필수..10 : 안전인증대상, 20 : 안전확인대상, 30 : 공급자적합성확인
				safeStrRst = safeStrRst & "			<certNo></certNo>"							'인증번호 | 인증 구분이 10, 20 일경우에만 필수-
				safeStrRst = safeStrRst & "		</certInfo>"
			End If
			rsget.Close
		End If

		If isElec = "Y" Then
			elecStrRst = ""
			elecStrRst = elecStrRst & "		<certInfo>"
			elecStrRst = elecStrRst & "			<certKind>6000000003</certKind>"			'#인증종류 (commCd:I387) | 인증대상 카테고리 일 경우 필수..6000000001 : 어린이인증 대상여부, 6000000002 : 안전인증 대상여부, 6000000003 : 전파인증 적합성평가 대상여부, 6000000004 : 위해우려제품 표시대상여부
			elecStrRst = elecStrRst & "			<certYn>N</certYn>"							'#인증 여부
			elecStrRst = elecStrRst & "			<certDivCd></certDivCd>"					'인증 구분 (commCd:I368) | 인증여부가 Y이고 인증종류가 (certKind=6000000001 | 6000000002) 일 경우 필수..10 : 안전인증대상, 20 : 안전확인대상, 30 : 공급자적합성확인
			elecStrRst = elecStrRst & "			<certNo></certNo>"							'인증번호 | 인증 구분이 10, 20 일경우에만 필수-
			elecStrRst = elecStrRst & "		</certInfo>"
		End If

		If isHarm = "Y" OR isElec = "Y" OR isSafe = "Y" OR isChild = "Y" Then
			strRst = strRst & "	<certInfos>"
			strRst = strRst & harmStrRst & elecStrRst & safeStrRst & childIntoStrRst
			strRst = strRst & "	</certInfos>"
		End If
		getCertInfoParam = childStrRst & strRst
	End Function

	Public Function getCertInfoNewParam(iCode)
		Dim strRst, strSql
		Dim safetyDiv, certNum
		Dim mndtyYnCnt, itemAppePropClsId, itemAppePropId, itemAppePropTypeCd

		strSql = ""
		strSql = strSql & " SELECT TOP 5 t.safetyDiv, isnull(f.certNum, '') as certNum " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg as t " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_safetycert_info as f on t.itemid = f.itemid " & vbcrlf
		strSql = strSql & " WHERE t.itemid='"&FItemID&"' " & vbcrlf
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			safetyDiv	= rsget("safetyDiv")
			certNum		= rsget("certNum")
		End If
		rsget.Close

		'텐바이텐에 인증번호가 입력이 안 된 경우
		If certNum = "" Then
			strSql = ""
			strSql = strSql & " SELECT COUNT(*) as cnt FROM db_etcmall.[dbo].[tbl_ssg_mmg_cate_SafeInfo] WHERE stdCtgDclsId = '"&iCode&"' and mndtyYn = 'Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				mndtyYnCnt	= rsget("cnt")
			End If
			rsget.Close

			If mndtyYnCnt = 0 Then
				getCertInfoNewParam = ""
			Else
				strSql = ""
				strSql = strSql & " SELECT TOP 100 f.itemAppePropClsId, f.itemAppePropId, f.itemAppePropTypeCd "
				strSql = strSql & " FROM db_etcmall.[dbo].[tbl_ssg_mmg_cate_SafeInfo] as f "
				strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_ssg_stdCate_mapping] as m on f.stdCtgDclsId = m.stdCtgDClsCd "
				strSql = strSql & " WHERE f.mndtyYn = 'Y' "
				strSql = strSql & " and f.stdCtgDclsId = '"&iCode&"' "
				strSql = strSql & " and m.tenCateLarge = '"& FtenCateLarge &"' "
				strSql = strSql & " and m.tenCateMid = '"& FtenCateMid &"' "
				strSql = strSql & " and m.tenCateSmall = '"& FtenCateSmall &"' "
				strSql = strSql & " ORDER BY f.itemAppePropId ASC "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) then
					strRst = ""
					strRst = strRst & "<certificationProps>"
					Do until rsget.EOF
						itemAppePropClsId	= rsget("itemAppePropClsId")
						itemAppePropId		= rsget("itemAppePropId")
						itemAppePropTypeCd	= rsget("itemAppePropTypeCd")

						strRst = strRst & "<prop>"
						strRst = strRst & "	<itemAppePropClsId>"&itemAppePropClsId&"</itemAppePropClsId>"
						strRst = strRst & "	<itemAppePropId>"&itemAppePropId&"</itemAppePropId>"
						If itemAppePropTypeCd = "30" Then
							strRst = strRst & "	<itemAppePropCntt>10</itemAppePropCntt>"
						Else
							strRst = strRst & "	<itemAppePropCntt>refer-ItemView</itemAppePropCntt>"
						End If
						strRst = strRst & "</prop>"
						rsget.MoveNext
					Loop
					strRst = strRst & "</certificationProps>"
				End If
				rsget.Close
				getCertInfoNewParam = strRst
			End If
		Else		'텐바이텐에 입력이 된 경우
			Dim chkMappCode1, chkMappCode2
			Select Case safetyDiv
				Case "10"
					chkMappCode1 = "6100000100"
					chkMappCode2 = "6100000103"
				Case "20"
					chkMappCode1 = "6100000110"
					chkMappCode2 = "6100000113"
				Case "30"
					chkMappCode1 = "6100000120"
					chkMappCode2 = ""
				Case "40"
					chkMappCode1 = "6100000070"
					chkMappCode2 = "6100000073"
				Case "50"
					chkMappCode1 = "6100000080"
					chkMappCode2 = "6100000083"
				Case "60"
					chkMappCode1 = "6100000090"
					chkMappCode2 = ""
				Case "70"
					chkMappCode1 = "6100000010"
					chkMappCode2 = "6100000013"
				Case "80"
					chkMappCode1 = "6100000020"
					chkMappCode2 = "6100000023"
				Case "90"
					chkMappCode1 = "6100000030"
					chkMappCode2 = ""
			End Select

			strSql = ""
			strSql = strSql & " SELECT COUNT(*) as cnt FROM db_etcmall.[dbo].[tbl_ssg_mmg_cate_SafeInfo] WHERE stdCtgDclsId = '"&iCode&"' and mndtyYn = 'Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				mndtyYnCnt	= rsget("cnt")
			End If
			rsget.Close
			'ssg에서 요구하는 값과 우리값과 다른 경우
			If mndtyYnCnt = 0 Then

				strSql = ""
				strSql = strSql & " SELECT TOP 100 f.itemAppePropClsId, f.itemAppePropId, f.itemAppePropTypeCd "
				strSql = strSql & " FROM db_etcmall.[dbo].[tbl_ssg_mmg_cate_SafeInfo] as f "
				strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_ssg_stdCate_mapping] as m on f.stdCtgDclsId = m.stdCtgDClsCd "
				strSql = strSql & " WHERE f.stdCtgDclsId = '"&iCode&"' "
				strSql = strSql & " and f.itemAppePropClsId = '"& chkMappCode1 &"' "
				strSql = strSql & " and m.tenCateLarge = '"& FtenCateLarge &"' "
				strSql = strSql & " and m.tenCateMid = '"& FtenCateMid &"' "
				strSql = strSql & " and m.tenCateSmall = '"& FtenCateSmall &"' "
				strSql = strSql & " GROUP BY f.itemAppePropClsId, f.itemAppePropId, f.itemAppePropTypeCd "
				strSql = strSql & " ORDER BY f.itemAppePropId ASC "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) then
					strRst = ""
					strRst = strRst & "<certificationProps>"
					Do until rsget.EOF
						itemAppePropClsId	= rsget("itemAppePropClsId")
						itemAppePropId		= rsget("itemAppePropId")
						itemAppePropTypeCd	= rsget("itemAppePropTypeCd")

						strRst = strRst & "<prop>"
						strRst = strRst & "	<itemAppePropClsId>"&itemAppePropClsId&"</itemAppePropClsId>"
						strRst = strRst & "	<itemAppePropId>"&itemAppePropId&"</itemAppePropId>"
						If itemAppePropTypeCd = "30" Then
							strRst = strRst & "	<itemAppePropCntt>10</itemAppePropCntt>"
						ElseIf itemAppePropId = chkMappCode2 Then
							strRst = strRst & "	<itemAppePropCntt>"&certNum&"</itemAppePropCntt>"
						Else
							strRst = strRst & "	<itemAppePropCntt></itemAppePropCntt>"
						End If
						strRst = strRst & "</prop>"
						rsget.MoveNext
					Loop
					strRst = strRst & "</certificationProps>"
				End If
				rsget.Close
				getCertInfoNewParam = strRst
			Else
			'ssg에서 요구하는 값과 우리값과 같은 경우
				strSql = ""
				strSql = strSql & " SELECT TOP 100 f.itemAppePropClsId, f.itemAppePropId, f.itemAppePropTypeCd "
				strSql = strSql & " FROM db_etcmall.[dbo].[tbl_ssg_mmg_cate_SafeInfo] as f "
				strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_ssg_stdCate_mapping] as m on f.stdCtgDclsId = m.stdCtgDClsCd "
				strSql = strSql & " WHERE f.mndtyYn = 'Y' "
				strSql = strSql & " and f.stdCtgDclsId = '"&iCode&"' "
				strSql = strSql & " and m.tenCateLarge = '"& FtenCateLarge &"' "
				strSql = strSql & " and m.tenCateMid = '"& FtenCateMid &"' "
				strSql = strSql & " and m.tenCateSmall = '"& FtenCateSmall &"' "
				strSql = strSql & " ORDER BY f.itemAppePropId ASC "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) then
					strRst = ""
					strRst = strRst & "<certificationProps>"
					Do until rsget.EOF
						itemAppePropClsId	= rsget("itemAppePropClsId")
						itemAppePropId		= rsget("itemAppePropId")
						itemAppePropTypeCd	= rsget("itemAppePropTypeCd")

						strRst = strRst & "<prop>"
						strRst = strRst & "	<itemAppePropClsId>"&itemAppePropClsId&"</itemAppePropClsId>"
						strRst = strRst & "	<itemAppePropId>"&itemAppePropId&"</itemAppePropId>"
						If itemAppePropTypeCd = "30" Then
							strRst = strRst & "	<itemAppePropCntt>10</itemAppePropCntt>"
						ElseIf itemAppePropId = chkMappCode2 Then
							strRst = strRst & "	<itemAppePropCntt>"&certNum&"</itemAppePropCntt>"
						Else
							strRst = strRst & "	<itemAppePropCntt>refer-ItemView</itemAppePropCntt>"
						End If
						strRst = strRst & "</prop>"
						rsget.MoveNext
					Loop
					strRst = strRst & "</certificationProps>"
				End If
				rsget.Close
				getCertInfoNewParam = strRst
			End If
		End If

		If  (session("ssBctID")="kjy8517") Then
			'rw getCertInfoNewParam
		End If
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
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
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
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
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
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
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
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
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
		sqlStr = sqlStr & " SELECT TOP 1 stdCtgDClsCd "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_stdCate_mapping] "
		sqlStr = sqlStr & " WHERE tenCateLarge = '"& FtenCateLarge &"' "
		sqlStr = sqlStr & " and tenCateMid = '"& FtenCateMid &"' "
		sqlStr = sqlStr & " and tenCateSmall = '"& FtenCateSmall &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			standardCode		= rsget("stdCtgDClsCd")
		End If
		rsget.Close

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 3 dispCtgId, tenCateLarge, tenCateMid, tenCateSmall, siteNo, lastupdate "
		sqlStr = sqlStr & " from db_etcmall.[dbo].[tbl_ssg_DispCate_mapping]  "
		sqlStr = sqlStr & " WHERE tenCateLarge = '"& FtenCateLarge &"' "
		sqlStr = sqlStr & " and tenCateMid = '"& FtenCateMid &"' "
		sqlStr = sqlStr & " and tenCateSmall = '"& FtenCateSmall &"' "
		sqlStr = sqlStr & " ORDER BY siteNo DESC "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				arrDepthCode		= arrDepthCode & rsget("dispCtgId") & ","
				arrSiteNo			= arrSiteNo & rsget("siteNo") & ","
				rsget.MoveNext
			Next
			arrDepthCode = RightCommaDel(arrDepthCode)
			arrSiteNo = RightCommaDel(arrSiteNo)
		End If
		rsget.Close

		getSsgCategoryParam = standardCode & "|_|" & arrDepthCode & "|_|" & arrSiteNo
	End Function

	Public Function getSsgOptParamtoEDIT(vArrSiteNum)
		Dim strRst, strRst2, strRst3, strSql, chkMultiOpt, requireDetailStr, i, j
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
					strRst = strRst & "			<uitemOptnNm1><![CDATA["&opt1name&"]]></uitemOptnNm1>"			'#단품 옵션 명1
					strRst = strRst & "			<uitemOptnTypeNm2>"&OptTypeNm2&"</uitemOptnTypeNm2>"			'단품 옵션 유형명2
					strRst = strRst & "			<uitemOptnNm2><![CDATA["&opt2name&"]]></uitemOptnNm2>"			'단품 옵션 명2
					strRst = strRst & "			<uitemOptnTypeNm3>"&OptTypeNm3&"</uitemOptnTypeNm3>"			'단품 옵션 유형명3
					strRst = strRst & "			<uitemOptnNm3><![CDATA["&opt3name&"]]></uitemOptnNm3>"			'단품 옵션 명3
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
'					For j = 0 to Ubound(vArrSiteNum)
'						If vArrSiteNum(j) <> "6005" Then
							strRst3 = strRst3 & "		<uitemPrc>"
							If preged = 0 Then
								strRst3 = strRst3 & "		<tempUitemId>"&itemoption&"</tempUitemId>"					'#단품ID (임시번호)
							Else
								strRst3 = strRst3 & "		<uitemId>"&outmalloptcode&"</uitemId>"						'#단품ID
							End If
'							strRst3 = strRst3 & "			<siteNo>"&vArrSiteNum(j)&"</siteNo>"						'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
							strRst3 = strRst3 & "			<splprc>"&Clng((MustPrice + optaddprice) * 0.85)&"</splprc>"		'#공급가
							strRst3 = strRst3 & "			<sellprc>"&MustPrice + optaddprice&"</sellprc>"				'#판매가
							strRst3 = strRst3 & "			<mrgrt>"&getSSGMargin&"</mrgrt>"								'#마진율
							strRst3 = strRst3 & "		</uitemPrc>"
'						End if
'					Next
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
					strRst = strRst & "			<uitemOptnNm1><![CDATA["&opt1name&"]]></uitemOptnNm1>"			'#단품 옵션 명1
					strRst = strRst & "			<uitemOptnTypeNm2></uitemOptnTypeNm2>"							'단품 옵션 유형명2
					strRst = strRst & "			<uitemOptnNm2></uitemOptnNm2>"									'단품 옵션 명2
					strRst = strRst & "			<uitemOptnTypeNm3></uitemOptnTypeNm3>"							'단품 옵션 유형명3
					strRst = strRst & "			<uitemOptnNm3></uitemOptnNm3>"									'단품 옵션 명3
					strRst = strRst & "			<uitemOptnTypeNm4></uitemOptnTypeNm4>"							'단품 옵션 유형명4
					strRst = strRst & "			<uitemOptnNm4></uitemOptnNm4>"									'단품 옵션 명4
					strRst = strRst & "			<uitemOptnTypeNm5></uitemOptnTypeNm5>"							'단품 옵션 유형명5
					strRst = strRst & "			<uitemOptnNm5></uitemOptnNm5>"									'단품 옵션 명5
				Else
					strRst = strRst & "			<uitemId>"&outmalloptcode&"</uitemId>"							'#단품ID
				End If
					strRst = strRst & "			<sellStatCd>"&sellStatCd&"</sellStatCd>"						'판매상태코드 | 20:판매중, 80:일시판매중지, 90:영구판매중지
					strRst = strRst & "			<baseInvQty>"&optLimit&"</baseInvQty>"							'재고 수량
					strRst = strRst & "			<useYn>Y</useYn>"												'사용 여부...Y로 그냥 보내도 되나??
					strRst = strRst & "		</uitem>"


'					For j = 0 to Ubound(vArrSiteNum)
'						If vArrSiteNum(j) <> "6005" Then
							strRst3 = strRst3 & "		<uitemPrc>"
							If preged = 0 Then
								strRst3 = strRst3 & "			<tempUitemId>"&itemoption&"</tempUitemId>"				'#단품ID (임시번호)
							Else
								strRst3 = strRst3 & "			<uitemId>"&outmalloptcode&"</uitemId>"					'#단품ID
							End If
'							strRst3 = strRst3 & "			<siteNo>"&vArrSiteNum(j)&"</siteNo>"						'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
							strRst3 = strRst3 & "			<splprc>"&Clng((MustPrice + optaddprice) * 0.85)&"</splprc>"		'#공급가
							strRst3 = strRst3 & "			<sellprc>"&MustPrice + optaddprice&"</sellprc>"				'#판매가
							strRst3 = strRst3 & "			<mrgrt>"&getSSGMargin&"</mrgrt>"								'#마진율
							strRst3 = strRst3 & "		</uitemPrc>"
'						End If
'					Next
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
'			strRst2 = strRst2 & Replace(strRst3, "<siteNo>6004</siteNo>", "<siteNo>6001</siteNo>")					'// 이마트몰 추가
			strRst2 = strRst2 & "	</uitemPluralPrcs>"
		End If
'response.write strRst & requireDetailStr & strRst2
'response.end
		getSsgOptParamtoEDIT = strRst & requireDetailStr & strRst2
	End Function

	Public Function getSsgOptParamtoREG(vArrSiteNum)
		Dim strRst, strRst2, strRst3, strSql, chkMultiOpt, arrOptTypeNm, requireDetailStr, i
		Dim itemSellTypeCd, OptTypeNm1, OptTypeNm2, OptTypeNm3, optLimit, itemoption, arrOptionname, optionname1, optionname2, optionname3, optaddprice
		Dim vssgMargin, vSpecialPrice
		vssgMargin = getSSGMargin
		vSpecialPrice = MustPrice()

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
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
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
						strRst = strRst & "			<uitemOptnNm1><![CDATA["&optionname1&"]]></uitemOptnNm1>"		'#단품 옵션 명1
						strRst = strRst & "			<uitemOptnTypeNm2>"&OptTypeNm2&"</uitemOptnTypeNm2>"			'단품 옵션 유형명2
						strRst = strRst & "			<uitemOptnNm2><![CDATA["&optionname2&"]]></uitemOptnNm2>"		'단품 옵션 명2
						strRst = strRst & "			<uitemOptnTypeNm3>"&OptTypeNm3&"</uitemOptnTypeNm3>"			'단품 옵션 유형명3
						strRst = strRst & "			<uitemOptnNm3><![CDATA["&optionname3&"]]></uitemOptnNm3>"		'단품 옵션 명3
						strRst = strRst & "			<uitemOptnTypeNm4></uitemOptnTypeNm4>"							'단품 옵션 유형명4
						strRst = strRst & "			<uitemOptnNm4></uitemOptnNm4>"									'단품 옵션 명4
						strRst = strRst & "			<uitemOptnTypeNm5></uitemOptnTypeNm5>"							'단품 옵션 유형명5
						strRst = strRst & "			<uitemOptnNm5></uitemOptnNm5>"									'단품 옵션 명5
						strRst = strRst & "			<baseInvQty>"&optLimit&"</baseInvQty>"							'재고 수량
						strRst = strRst & "			<useYn>Y</useYn>"												'사용 여부...Y로 그냥 보내도 되나??
						strRst = strRst & "		</uitem>"
'					For i = 0 to Ubound(vArrSiteNum)
'						If vArrSiteNum(i) <> "6005" Then
						strRst3 = strRst3 & "		<uitemPrc>"
						strRst3 = strRst3 & "			<tempUitemId>"&itemoption&"</tempUitemId>"					'#단품ID (임시번호)
'						strRst3 = strRst3 & "			<siteNo>"&vArrSiteNum(i)&"</siteNo>"						'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
						strRst3 = strRst3 & "			<splprc>"&Clng((vSpecialPrice + optaddprice) * 0.85)&"</splprc>"		'#공급가
						strRst3 = strRst3 & "			<sellprc>"&vSpecialPrice + optaddprice&"</sellprc>"				'#판매가
						strRst3 = strRst3 & "			<mrgrt>"&vssgMargin&"</mrgrt>"								'#마진율
						strRst3 = strRst3 & "		</uitemPrc>"
'						End If
'					Next
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
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
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
						strRst = strRst & "			<uitemOptnNm1><![CDATA["&optionname1&"]]></uitemOptnNm1>"		'#단품 옵션 명1
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

'					For i = 0 to Ubound(vArrSiteNum)
'						If vArrSiteNum(i) <> "6005" Then
						strRst3 = strRst3 & "		<uitemPrc>"
						strRst3 = strRst3 & "			<tempUitemId>"&itemoption&"</tempUitemId>"					'#단품ID (임시번호)
'						strRst3 = strRst3 & "			<siteNo>"&vArrSiteNum(i)&"</siteNo>"						'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
						strRst3 = strRst3 & "			<splprc>"&Clng((vSpecialPrice + optaddprice) * 0.85)&"</splprc>"		'#공급가
						strRst3 = strRst3 & "			<sellprc>"&vSpecialPrice + optaddprice&"</sellprc>"				'#판매가
						strRst3 = strRst3 & "			<mrgrt>"&vssgMargin&"</mrgrt>"								'#마진율
						strRst3 = strRst3 & "		</uitemPrc>"
'						End If
'					Next
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
'			strRst2 = strRst2 & Replace(strRst3, "<siteNo>6004</siteNo>", "<siteNo>6001</siteNo>")					'// 이마트몰 추가
			strRst2 = strRst2 & "	</uitemPluralPrcs>"
		End If
		getSsgOptParamtoREG = strRst & requireDetailStr & strRst2
	End Function

	Public Function getSsgItemInfoCdToReg(iareaCode)
		Dim strSql, buf, lp
		Dim mallinfoCd, infoContent, importYn

'		If FSourcearea = "한국" OR FSourcearea = "대한민국" OR FSourcearea = "국내산" OR FSourcearea = "양산시" OR UCASE(FSourcearea) = "KOREA" OR FSourcearea = "국내" Then
		If iareaCode = "2000000078" Then
			importYn = "N"
		Else
			importYn = "Y"
		End If

		strSql = ""
		strSql = strSql & " SELECT top 100 M.* , "
		strSql = strSql & " CASE WHEN (M.infoCdAdd='00000') AND ('"&importYn&"'='Y') THEN 'Y' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00000') AND ('"&importYn&"'='N') THEN 'N' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00001') AND (F.chkDiv='Y') THEN '10' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00001') AND (F.chkDiv='N') THEN '20' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00002') AND (F.chkDiv='Y') THEN 'O' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00002') AND (F.chkDiv='N') THEN 'N' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00003') THEN '"&importYn&"' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00004') THEN 'N' "	'SSG에서 가전제품 "으뜸효율가전 환급여부" 라는 것을 고시에 심어놈..전부 비대상처리
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00005') AND (LEN(isNULL(F.infocontent, '')) < 2) THEN 'N' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00005') AND (LEN(isNULL(F.infocontent, '')) >= 2) THEN 'O' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00006') THEN '0000000086' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00007') THEN '상세페이지 참고' "
		strSql = strSql & " 	 WHEN (M.mallinfoCd='0000000022') AND (LEN(isNULL(F.infocontent, '')) < 2) THEN '"& replace(getItemNameFormat, "'", "") &"' "
		strSql = strSql & " 	 WHEN (M.mallinfoCd='0000000022') AND (LEN(isNULL(F.infocontent, '')) >= 2) THEN F.infocontent "
		strSql = strSql & " 	 WHEN (M.mallinfoCd='0000000103' OR M.mallinfoCd='0000000058' OR M.mallinfoCd='0000000106' OR M.mallinfoCd='0000000408') AND (F.chkDiv='N') THEN 'N' "
		strSql = strSql & " 	 WHEN (M.mallinfoCd='0000000103' OR M.mallinfoCd='0000000058' OR M.mallinfoCd='0000000106' OR M.mallinfoCd='0000000408') AND (F.chkDiv='Y') THEN 'Y' "
		strSql = strSql & " 	 WHEN (M.mallinfoCd='0000000009') AND ('"&importYn&"'='N') THEN 'XXXX' "
		strSql = strSql & " 	 WHEN (M.mallinfoCd='0000000011' OR M.mallinfoCd='0000000196') THEN '"&iareaCode&"' "
		strSql = strSql & " 	 WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035' "
		strSql = strSql & " 	 WHEN LEN(isnull(F.infocontent, '') + isNULL(F2.infocontent,'')) < 2 THEN '상세페이지 참고' "
		strSql = strSql & " ELSE isNull(F.infocontent, '') + isNULL(F2.infocontent,'') END AS infocontent "
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"' "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='"&FItemid&"' "
		strSql = strSql & " WHERE M.mallid = 'ssg' and IC.itemid='"&FItemid&"' "
		strSql = strSql & " ORDER BY M.mallinfocd ASC "
'  If  (session("ssBctID")="kjy8517") Then
'  	rw strSql
'  End If
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		buf = ""
		If not rsget.EOF Then
			buf = buf & "	<itemMngPropClsId>"& rsget("infoETC") &"</itemMngPropClsId>"
			buf = buf & "	<itemMngAttrs>"
			Do until rsget.EOF
				infoContent = rsget("infocontent")
				infoContent = replace(infoContent, "_", "")
				mallinfocd = rsget("mallinfocd")

				If infoContent <> "XXXX" Then
					buf = buf & "	<itemMngAttr>"
					buf = buf & "		<itemMngPropId>"&mallinfocd&"</itemMngPropId>"
					buf = buf & "		<itemMngCntt><![CDATA["&infoContent&"]]></itemMngCntt>"
					buf = buf & "	</itemMngAttr>"
				End If
				rsget.MoveNext
			Loop
			buf = buf & "	</itemMngAttrs>"
		End If
		rsget.Close
		getSsgItemInfoCdToReg = buf
'response.write buf
	End Function

	'SSG 등록 XML
	Public Function getSsgItemRegParameter(imustPrice)
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

		If getShopLeadTime > 3 Then
			If getShopLeadTime = 5 Then	'5일 소요 카테고리는 일반으로 요청옴..2021-07-28 김진영 수정
				If FItemdiv = "06" OR FItemdiv = "16" Then
					shppRqrmDcnt = "7"
					shppItemDivCd = "05"
					shppRqrmDcntChngRsnCntt = "주문제작상품"
				Else
					shppItemDivCd = "01"
				End If
			Else
				shppItemDivCd = "05"
				shppRqrmDcntChngRsnCntt = "주문제작상품"
			End If
		Else
			If FItemdiv = "06" OR FItemdiv = "16" Then
				shppRqrmDcnt = "7"
				shppItemDivCd = "05"
				shppRqrmDcntChngRsnCntt = "주문제작상품"
			Else
				shppItemDivCd = "01"
			End If
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst & "<insertItem>"
		strRst = strRst & "	<itemNm><![CDATA["&getItemNameFormat&"]]></itemNm>"								'#상품명
		strRst = strRst & "	<mdlNm></mdlNm>"																'모델명
		strRst = strRst & "	<brandId>2000047517</brandId>"													'#브랜드ID | 텐바이텐(2000047517)
		strRst = strRst & "	<stdCtgId>"&standardCateCode&"</stdCtgId>"										'#표준카테고리ID
		strRst = strRst & "	<sites>"
	For i = 0 to Ubound(arrSiteNum)
		If arrSiteNum(i) <> "6005" Then
			strRst = strRst & "		<site>"
			strRst = strRst & "			<siteNo>"&arrSiteNum(i)&"</siteNo>"									'#사이트번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
			strRst = strRst & "			<sellStatCd>"&sellStatCd&"</sellStatCd>"							'#판매 상태 코드 | 20 : 판매중, 80 : 일시판매중지
			strRst = strRst & "		</site>"
		End If
	Next
		strRst = strRst & "	</sites>"
		strRst = strRst & "	<itemAplRngTypeCd></itemAplRngTypeCd>"											'상품 적용 범위 | 00 : 전체적용, 10 : B2C적용, 20 : B2E적용
		strRst = strRst & "	<b2eAplRngCd>10</b2eAplRngCd>"													'B2E 적용 범위 | 10 : 전체 적용, 20 : 적용 않음, 30 : 회원사 지정
		strRst = strRst & "	<b2cAplRngCd>10</b2cAplRngCd>"													'B2C 적용 범위 | 10 : 적용, 20 : 적용 (대행 제휴사 제외), 70 : 적용 않음
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
	For i = 0 to Ubound(arrSiteNum)
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
		strRst = strRst & "	<itemSrchwdNm><![CDATA["&RightCommaDel(Trim(getKeywords()))&"]]></itemSrchwdNm>"	'상품검색어명
		strRst = strRst & "	<aplMbrGrdCd></aplMbrGrdCd>"													'노출 회원 등급 (값이 존재하지 않을 경우 ALL) | 10 : 패밀리, 20 : 브론즈, 30 : 실버, 40 : 골드, 50 : VIP, 90 : VVIP
		strRst = strRst & "	<minOnetOrdPsblQty>1</minOnetOrdPsblQty>"										'#최소 1회 주문 가능 수량
		strRst = strRst & "	<maxOnetOrdPsblQty>"& getOrderMaxNum &"</maxOnetOrdPsblQty>"					'#최대 1회 주문 가능 수량
		strRst = strRst & "	<max1dyOrdPsblQty>9999</max1dyOrdPsblQty>"										'#최대 1일 주문 가능 수량
		strRst = strRst & "	<adultItemTypeCd>"&Chkiif(IsAdultItem() = "Y", "10", "90")&"</adultItemTypeCd>"	'#성인 상품 타입 코드 (commCd:I408) | 10 : 성인 상품, 20 : 주류 상품, 90 : 일반 상품
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
'		strRst = strRst & "	<sellTotCapa></sellTotCapa>"													'판매 총 용량
'		strRst = strRst & "	<sellUnitCapa></sellUnitCapa>"													'판매 단위 용량
		strRst = strRst & "	<sellUnitQty>0</sellUnitQty>"													'판매 단위 수량
		strRst = strRst & "	<buyFrmCd>60</buyFrmCd>"														'#매입 형태 코드 (commCd:I002) | 10 : 직매입, 20 : 직매입2(판매분), 40 : 특정매입, 60 : 위수탁
		strRst = strRst & "	<txnDivCd>"&CHKIIF(FVatInclude="N","20","10")&"</txnDivCd>"						'#과세 구분 코드 (commCd:I005) | 10 : 과세, 20 : 면세, 30 : 영세
		strRst = strRst & "	<prcMngMthd>1</prcMngMthd>"														'가격설정방식 | 1 : 공급가 자동계산 (Default), 2 : 판매가 자동계산, 3 : 마진 자동계산..이 값 설정시 SALE_PRC_INFO, B2E_PRC 둘다 적용 받는다. 값은 모두 입력 받아도 상관 없으나 해당 값 설정에 따라 해당 값이 자동으로 계산됨.
		strRst = strRst & "	<salesPrcInfos>"
'	For i = 0 to Ubound(arrSiteNum)
'		If arrSiteNum(i) <> "6005" Then
			strRst = strRst & "		<uitemPrc>"
'			strRst = strRst & "			<siteNo>"&arrSiteNum(i)&"</siteNo>"									'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
			strRst = strRst & "			<splprc>"&Clng(imustPrice*0.85)&"</splprc>"							'#공급가
			strRst = strRst & "			<sellprc>"&imustPrice&"</sellprc>"									'#판매가
			strRst = strRst & "			<mrgrt>"&getSSGMargin&"</mrgrt>"									'#마진율
			strRst = strRst & "		</uitemPrc>"
'		End If
'	Next
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
		strRst = strRst & getSsgOptParamtoREG(arrSiteNum)
		strRst = strRst & "	<shppItemDivCd>"&shppItemDivCd&"</shppItemDivCd>"								'#배송상품구분코드 (commCd:I070) | 01 : 일반, 02 : 해외구매대행, 03 : 설치(유료), 04 : 설치(무료), 05 : 주문제작, 06 : 해외직배송
		strRst = strRst & "	<exprtCntryId></exprtCntryId>"													'수출국가(해외 직배송 적출국)shppItemDivCd=06(해외직배송) 인 경우 필수 | 원산지 조회 API 참고(listOrplc API)
		strRst = strRst & "	<pcusMngCd></pcusMngCd>"														'개인 통관 고유 부호 | shppItemDivCd=06(해외직배송) 인 경우 필수 10 : 선택 입력, 20 : 필수 입력, 30 : 입력 안함
		strRst = strRst & "	<retExchPsblYn>"&chkiif(shppItemDivCd <> "05", "Y", "N")&"</retExchPsblYn>"												'#반품 교환 가능 여부
		strRst = strRst & "	<shppMainCd>41</shppMainCd>"													'#배송 주체 코드 (commCd:P017) | 31 : 자사창고, 32 : 업체창고, 41 : 협력업체
		strRst = strRst & "	<shppMthdCd>20</shppMthdCd>"													'#배송 방법 코드 (commCd:P021) | 10 : 자사배송, 20 : 택배배송, 30 : 매장방문, 40 : 등기, 50 : 미배송, 60 : 미발송
		strRst = strRst & "	<mareaShppYn></mareaShppYn>"													'#수도권 배송여부
		strRst = strRst & "	<shppRqrmDcnt>"&shppRqrmDcnt&"</shppRqrmDcnt>"									'#배송 소요 일수
		strRst = strRst & "	<shppRqrmDcntChngRsnCntt>"&shppRqrmDcntChngRsnCntt&"</shppRqrmDcntChngRsnCntt>"	'#배송 소요 일수 변경 사유 | 상품배송구분이 일반(01) 이고 배송소요일수가 4일 이상일 경우 필수
		strRst = strRst & "	<tdShppPsblYn>N</tdShppPsblYn>"													'#오늘출발 가능여부(Y/N) | 배송방법이 자사배송(10) 또는 택배배송(20)이고 배송상품구분이 일반(01) 인 경우 입력 가능
		strRst = strRst & "	<splVenItemId>"&FItemid&"</splVenItemId>"										'업체 상품 번호
		strRst = strRst & "	<whoutShppcstId>0000517199</whoutShppcstId>"									'#출고 배송비 ID
		strRst = strRst & "	<retShppcstId>0000520999</retShppcstId>"										'#반품 배송비 ID
'		strRst = strRst & "	<retShppcstId>"& Chkiif(MustPrice() >= 30000, "0000011336", "0000520999") &"</retShppcstId>"		'반품 배송비 ID
	If Fmakerid <> "gagufeeling" Then
		strRst = strRst & "	<ismtarAddShppcstId>0000543853</ismtarAddShppcstId>"							'도서산간 추가배송비 ID
		strRst = strRst & "	<jejuAddShppcstId>0000543854</jejuAddShppcstId>"								'제주도 추가배송비 ID
	Else
		strRst = strRst & "	<ismtarAddShppcstId></ismtarAddShppcstId>"										'도서산간 추가배송비 ID
		strRst = strRst & "	<jejuAddShppcstId></jejuAddShppcstId>"											'제주도 추가배송비 ID
	End If
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
'		strRst = strRst & getCertInfoParam(standardCateCode)
		strRst = strRst & getCertInfoNewParam(standardCateCode)
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
'		strRst = strRst & "	<giftPackPsblYn>N</giftPackPsblYn>"												'선물 포장 가능 여부
		strRst = strRst & "</insertItem>"
		getSsgItemRegParameter = strRst
'response.write Replace(strRst, "<?xml","ASDASDASD")
'response.end
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

'		shppItemDivCd = "01"
'		If getShopLeadTime > 3 Then
'			shppRqrmDcntChngRsnCntt = "주문제작상품"
'		End If

		' If getShopLeadTime > 3 Then
		' 	shppItemDivCd = "05"
		' 	shppRqrmDcntChngRsnCntt = "주문제작상품"
		' Else
		' 	shppItemDivCd = "01"
		' End If

		If getShopLeadTime > 3 Then
			If getShopLeadTime = 5 Then	'5일 소요 카테고리는 일반으로 요청옴..2021-07-28 김진영 수정
				If FItemdiv = "06" OR FItemdiv = "16" Then
					shppRqrmDcnt = "7"
					shppItemDivCd = "05"
					shppRqrmDcntChngRsnCntt = "주문제작상품"
				Else
					shppItemDivCd = "01"
				End If
			Else
				shppItemDivCd = "05"
				shppRqrmDcntChngRsnCntt = "주문제작상품"
			End If
		Else
			If FItemdiv = "06" OR FItemdiv = "16" Then
				shppRqrmDcnt = "7"
				shppItemDivCd = "05"
				shppRqrmDcntChngRsnCntt = "주문제작상품"
			Else
				shppItemDivCd = "01"
			End If
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
	For i = 0 to Ubound(arrSiteNum)
		If arrSiteNum(i) <> "6005" Then
			strRst = strRst & "		<site>"
			strRst = strRst & "			<siteNo>"&arrSiteNum(i)&"</siteNo>"									'#사이트번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
			strRst = strRst & "			<sellStatCd>20</sellStatCd>"										'#판매 상태 코드 | 20 : 판매중, 80 : 일시판매중지
			strRst = strRst & "		</site>"
		End If
	Next
		strRst = strRst & "	</sites>"
		strRst = strRst & "	<itemAplRngTypeCd></itemAplRngTypeCd>"											'상품 적용 범위 | 00 : 전체적용, 10 : B2C적용, 20 : B2E적용
		strRst = strRst & "	<b2eAplRngCd>10</b2eAplRngCd>"													'B2E 적용 범위 | 10 : 전체 적용, 20 : 적용 않음, 30 : 회원사 지정
		strRst = strRst & "	<b2cAplRngCd>10</b2cAplRngCd>"													'B2C 적용 범위 | 10 : 적용, 20 : 적용 (대행 제휴사 제외), 70 : 적용 않음
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
	For i = 0 to Ubound(arrSiteNum)
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
		strRst = strRst & "	<itemSrchwdNm><![CDATA["&RightCommaDel(Trim(getKeywords()))&"]]></itemSrchwdNm>"	'상품검색어명
		strRst = strRst & "	<aplMbrGrdCd></aplMbrGrdCd>"													'노출 회원 등급 (값이 존재하지 않을 경우 ALL) | 10 : 패밀리, 20 : 브론즈, 30 : 실버, 40 : 골드, 50 : VIP, 90 : VVIP
		strRst = strRst & "	<minOnetOrdPsblQty>1</minOnetOrdPsblQty>"										'최소 1회 주문 가능 수량
		strRst = strRst & "	<maxOnetOrdPsblQty>"& getOrderMaxNum &"</maxOnetOrdPsblQty>"					'최대 1회 주문 가능 수량
		strRst = strRst & "	<max1dyOrdPsblQty>9999</max1dyOrdPsblQty>"										'최대 1일 주문 가능 수량
		strRst = strRst & "	<adultItemTypeCd>"&Chkiif(IsAdultItem() = "Y", "10", "90")&"</adultItemTypeCd>"	'#성인 상품 타입 코드 (commCd:I408) | 10 : 성인 상품, 20 : 주류 상품, 90 : 일반 상품
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
'		strRst = strRst & "	<sellTotCapa></sellTotCapa>"													'판매 총 용량
'		strRst = strRst & "	<sellUnitCapa></sellUnitCapa>"													'판매 단위 용량
		strRst = strRst & "	<sellUnitQty>0</sellUnitQty>"													'판매 단위 수량
		strRst = strRst & "	<prcMngMthd>1</prcMngMthd>"														'가격설정방식 | 1 : 공급가 자동계산 (Default), 2 : 판매가 자동계산, 3 : 마진 자동계산..이 값 설정시 SALE_PRC_INFO, B2E_PRC 둘다 적용 받는다. 값은 모두 입력 받아도 상관 없으나 해당 값 설정에 따라 해당 값이 자동으로 계산됨.
		strRst = strRst & "	<salesPrcInfos>"
'	For i = 0 to Ubound(arrSiteNum)
'		If arrSiteNum(i) <> "6005" Then
			strRst = strRst & "		<uitemPrc>"
'			strRst = strRst & "			<siteNo>"&arrSiteNum(i)&"</siteNo>"									'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
			strRst = strRst & "			<splprc>"&Clng(MustPrice()*0.85)&"</splprc>"						'#공급가
			strRst = strRst & "			<sellprc>"&MustPrice()&"</sellprc>"									'#판매가
			strRst = strRst & "			<mrgrt>"&getSSGMargin&"</mrgrt>"									'#마진율
			strRst = strRst & "		</uitemPrc>"
'		End If
'	Next
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
			strRst = strRst & getSsgOptParamtoEDIT(arrSiteNum)
'		End If
		strRst = strRst & "	<shppItemDivCd>"&shppItemDivCd&"</shppItemDivCd>"								'배송상품구분코드 (commCd:I070) | 01 : 일반, 02 : 해외구매대행, 03 : 설치(유료), 04 : 설치(무료), 05 : 주문제작, 06 : 해외직배송
		strRst = strRst & "	<exprtCntryId></exprtCntryId>"													'수출국가(해외 직배송 적출국)shppItemDivCd=06(해외직배송) 인 경우 필수 | 원산지 조회 API 참고(listOrplc API)
		strRst = strRst & "	<pcusMngCd></pcusMngCd>"														'개인 통관 고유 부호 | shppItemDivCd=06(해외직배송) 인 경우 필수 10 : 선택 입력, 20 : 필수 입력, 30 : 입력 안함
		strRst = strRst & "	<retExchPsblYn>"&chkiif(shppItemDivCd <> "05", "Y", "N")&"</retExchPsblYn>"												'반품 교환 가능 여부
		strRst = strRst & "	<shppMainCd>41</shppMainCd>"													'배송 주체 코드 (commCd:P017) | 31 : 자사창고, 32 : 업체창고, 41 : 협력업체
		strRst = strRst & "	<shppMthdCd>20</shppMthdCd>"													'배송 방법 코드 (commCd:P021) | 10 : 자사배송, 20 : 택배배송, 30 : 매장방문, 40 : 등기, 50 : 미배송, 60 : 미발송
		strRst = strRst & "	<mareaShppYn></mareaShppYn>"													'수도권 배송여부
		strRst = strRst & "	<shppRqrmDcnt>"&shppRqrmDcnt&"</shppRqrmDcnt>"									'배송 소요 일수
		strRst = strRst & "	<shppRqrmDcntChngRsnCntt>"&shppRqrmDcntChngRsnCntt&"</shppRqrmDcntChngRsnCntt>"	'배송 소요 일수 변경 사유 | 상품배송구분이 일반(01) 이고 배송소요일수가 4일 이상일 경우 필수
		strRst = strRst & "	<tdShppPsblYn>N</tdShppPsblYn>"													'오늘출발 가능여부(Y/N) | 배송방법이 자사배송(10) 또는 택배배송(20)이고 배송상품구분이 일반(01) 인 경우 입력 가능
		strRst = strRst & "	<splVenItemId>"&FItemid&"</splVenItemId>"										'업체 상품 번호
		strRst = strRst & "	<whoutShppcstId>0000517199</whoutShppcstId>"									'출고 배송비 ID
		strRst = strRst & "	<retShppcstId>0000520999</retShppcstId>"										'반품 배송비 ID
'		strRst = strRst & "	<retShppcstId>"& Chkiif(MustPrice() >= 30000, "0000011336", "0000520999") &"</retShppcstId>"		'반품 배송비 ID
	If Fmakerid <> "gagufeeling" Then
		strRst = strRst & "	<ismtarAddShppcstId>0000543853</ismtarAddShppcstId>"							'도서산간 추가배송비 ID
		strRst = strRst & "	<jejuAddShppcstId>0000543854</jejuAddShppcstId>"								'제주도 추가배송비 ID
	Else
		strRst = strRst & "	<ismtarAddShppcstId></ismtarAddShppcstId>"										'도서산간 추가배송비 ID
		strRst = strRst & "	<jejuAddShppcstId></jejuAddShppcstId>"											'제주도 추가배송비 ID
	End If
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
'		strRst = strRst & getCertInfoParam(standardCateCode)
		strRst = strRst & getCertInfoNewParam(standardCateCode)
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
'		strRst = strRst & "	<giftPackPsblYn>N</giftPackPsblYn>"												'선물 포장 가능 여부
		strRst = strRst & "</updateItem>"
'response.write strRst
'response.end
		getssgItemEditParameter = strRst
	End Function

	'SSG 품절 수정 XML // 필수값만 넣고 나머지 다 땜..옵션수정부분도
	Public Function getssgItemEditSellynParameter(ichgSellYn)
		Dim strRst, i, sellStatCd, areaCode, shppItemDivCd, shppRqrmDcnt, shppRqrmDcntChngRsnCntt
		If ichgSellYn = "Y" Then
			sellStatCd = 20
		ElseIf ichgSellYn = "X" Then
			sellStatCd = 90
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
		'################################### 배송기일  호출 #########################################
		shppRqrmDcnt = getShopLeadTime()
		'##########################################################################################
		' If getShopLeadTime > 3 Then
		' 	shppItemDivCd = "05"
		' 	shppRqrmDcntChngRsnCntt = "주문제작상품"
		' Else
		' 	shppItemDivCd = "01"
		' End If

		If getShopLeadTime > 3 Then
			If getShopLeadTime = 5 Then	'5일 소요 카테고리는 일반으로 요청옴..2021-07-28 김진영 수정
				If FItemdiv = "06" OR FItemdiv = "16" Then
					shppRqrmDcnt = "7"
					shppItemDivCd = "05"
					shppRqrmDcntChngRsnCntt = "주문제작상품"
				Else
					shppItemDivCd = "01"
				End If
			Else
				shppItemDivCd = "05"
				shppRqrmDcntChngRsnCntt = "주문제작상품"
			End If
		Else
			If FItemdiv = "06" OR FItemdiv = "16" Then
				shppRqrmDcnt = "7"
				shppItemDivCd = "05"
				shppRqrmDcntChngRsnCntt = "주문제작상품"
			Else
				shppItemDivCd = "01"
			End If
		End If


		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst & "<updateItem>"
		strRst = strRst & "	<itemId>"&FSsgGoodno&"</itemId>"												'상품ID
		strRst = strRst & "	<brandId>2000047517</brandId>"													'#브랜드ID | 텐바이텐(2000047517)
		strRst = strRst & "	<sites>"
	For i = 0 to Ubound(arrSiteNum)
		If arrSiteNum(i) <> "6005" Then
			strRst = strRst & "		<site>"
			strRst = strRst & "			<siteNo>"&arrSiteNum(i)&"</siteNo>"									'#사이트번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
			strRst = strRst & "			<sellStatCd>20</sellStatCd>"										'#판매 상태 코드 | 20 : 판매중, 80 : 일시판매중지
			strRst = strRst & "		</site>"
		End If
	Next
		strRst = strRst & "	</sites>"
		strRst = strRst & "	<sellStatCd>"&sellStatCd&"</sellStatCd>"										'판매 상태 코드 | 20 : 판매중, 80 : 일시판매중지, 90 : 영구판매중지
'		strRst = strRst & getSsgItemInfoCdToReg(areaCode)
		strRst = strRst & "	<dispCtgs>"
	For i = 0 to Ubound(arrSiteNum)
		strRst = strRst & "		<dispCtg>"
		strRst = strRst & "			<siteNo>"&arrSiteNum(i)&"</siteNo>"										'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰, 6005 : SSG.COM
		strRst = strRst & "			<delYn></delYn>"														'삭제 여부
		strRst = strRst & "			<dispCtgId>"&arrDisplayCateCode(i)&"</dispCtgId>"						'#전시 카테고리 ID
		strRst = strRst & "			<repDispOrdr>"&i+1&"</repDispOrdr>"										'대표 전시 순서 | 순서대로, 중복 허용하지 않음. 사이트별 최대 3개까지 선택 가능
		strRst = strRst & "		</dispCtg>"
	Next
		strRst = strRst & "	</dispCtgs>"
		strRst = strRst & "	<shppItemDivCd>"&shppItemDivCd&"</shppItemDivCd>"								'#배송상품구분코드 (commCd:I070) | 01 : 일반, 02 : 해외구매대행, 03 : 설치(유료), 04 : 설치(무료), 05 : 주문제작, 06 : 해외직배송
		strRst = strRst & "	<retExchPsblYn>"&chkiif(shppItemDivCd <> "05", "Y", "N")&"</retExchPsblYn>"		'#반품 교환 가능 여부
		strRst = strRst & "	<shppRqrmDcnt>"&shppRqrmDcnt&"</shppRqrmDcnt>"									'#배송 소요 일수
		strRst = strRst & "	<shppRqrmDcntChngRsnCntt>"&shppRqrmDcntChngRsnCntt&"</shppRqrmDcntChngRsnCntt>"	'#배송 소요 일수 변경 사유 | 상품배송구분이 일반(01) 이고 배송소요일수가 4일 이상일 경우 필수
		strRst = strRst & "	<tdShppPsblYn>N</tdShppPsblYn>"													'#오늘출발 가능여부(Y/N) | 배송방법이 자사배송(10) 또는 택배배송(20)이고 배송상품구분이 일반(01) 인 경우 입력 가능
'		strRst = strRst & getCertInfoParam(standardCateCode)
		strRst = strRst & getCertInfoNewParam(standardCateCode)
'		strRst = strRst & "	<salesPrcInfos>"
'	For i = 0 to Ubound(arrSiteNum)
'		If arrSiteNum(i) <> "6005" Then
'			strRst = strRst & "		<uitemPrc>"
'			strRst = strRst & "			<siteNo>"&arrSiteNum(i)&"</siteNo>"									'#사이트 번호 | 6001 : 이마트몰, 6002 : 트레이더스몰, 6003 : 분스몰, 6004 : 신세계몰, 6009 : 신세계백화점몰, 6200 : 신세계TV쇼핑몰
'			strRst = strRst & "			<splprc>"&MustPrice()*0.85&"</splprc>"								'#공급가
'			strRst = strRst & "			<sellprc>"&MustPrice()&"</sellprc>"									'#판매가
'			strRst = strRst & "			<mrgrt>"&getSSGMargin&"</mrgrt>"										'#마진율
'			strRst = strRst & "		</uitemPrc>"
'		End If
'	Next
'		strRst = strRst & "	</salesPrcInfos>"
		strRst = strRst & "</updateItem>"
'response.write strRst
'response.end
		getssgItemEditSellynParameter = strRst
		If  (session("ssBctID")="kjy8517") Then
			rw getssgItemEditSellynParameter
		End If
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
'		strSql = strSql & ", isNULL(k.keywords, c.keywords) as keywords "
		strSql = strSql & "	, (SELECT [db_etcmall].[dbo].[getOutmallKeywords] ('"& CMALLNAME &"', '" & FRectItemID & "')) as keywords "
		strSql = strSql & "	, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, isNULL(C.safetyNum, '') as safetyNum "
		strSql = strSql & "	, isNULL(R.ssgStatCD,-9) as ssgStatCD, cm.mapCnt, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	, UC.socname_kor, isNULL(c.requireMakeDay,0) as requireMakeDay, IsNull(R.ssgPrice, 0) as ssgPrice"
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_ssg_DispCate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ssg_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
'		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_outmall_keywords] as k on i.itemid = k.itemid and k.mallid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		IF (CUPJODLVVALID) then
'		    strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		ELSE
		    strSql = strSql & " and (i.deliveryType<>9)"
	    END IF
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.sellcash > i.buycash "
		strSql = strSql & " and i.itemdiv not in ('08', '09', '21', '30', '23') "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "						'플라워/화물배송/해외직구
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
'		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)>=" & CMAXMARGIN & ")" 2019-06-17 | 2397378 윤현주 예외처리 요청
'		strSql = strSql & " and ((i.itemid = 2397378) OR ( (i.itemid <> 2397378) and i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)>= " & CMAXMARGIN & "))"
		strSql = strSql & "	and ( "
		strSql = strSql & "		(i.itemid = 2397378) OR "
		strSql = strSql & "		( "
		strSql = strSql & "			(i.itemid <> 2397378) "
		strSql = strSql & "			and (i.sellcash <> 0) "
		strSql = strSql & "			and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & "		 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >=  "&CMAXLIMITSELL&" ) "
		strSql = strSql & "		 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >=  "&CMAXLIMITSELL&" ) "
		strSql = strSql & "		 				) THEN 'Y' "
		strSql = strSql & "		 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >=  "&CMAXLIMITSELL&" ) THEN 'Y' ELSE 'N' END "
		strSql = strSql & "		) "
		strSql = strSql & "	) "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'등록제외 카테고리
		strSql = strSql & "	and isnull(R.ssgGoodNo, '') = '' "
		strSql = strSql & " and cm.mapCnt is Not Null "
'		strSql = strSql & " and (i.mwdiv='M' or i.mwdiv='W') "		'매입 or 위탁
'		strSql = strSql & " and i.deliveryType = 1 "				'탠배
'2018-01-29 15:00 김진영 하단 주석처리..
'		strSql = strSql & " and ( ((i.mwdiv='M' or i.mwdiv='W') and (i.deliveryType = 1)) OR (i.makerid in ('meaningless01', 'mandarinebrothers', 'fromamour', 'woolly02' ,'dalbampicnic')) ) "
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
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
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FSsgPrice			= rsget("ssgPrice")
				FOneItem.FOrderMaxNum		= rsget("orderMaxNum")
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
'		strSql = strSql & "	, isNULL(k.keywords, c.keywords) as keywords "
		strSql = strSql & "	, (SELECT [db_etcmall].[dbo].[getOutmallKeywords] ('"& CMALLNAME &"', '" & FRectItemID & "')) as keywords "
		strSql = strSql & "	, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, isNULL(C.safetyNum, '') as safetyNum "
		strSql = strSql & "	, isNULL(m.ssgStatCD,-9) as ssgStatCD, cm.mapCnt, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	, UC.socname_kor, isNULL(c.requireMakeDay,0) as requireMakeDay, m.ssgGoodNo, m.ssgPrice "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"	'2019-05-30 김진영 i.sailyn='N' 조건 추가
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '30', '23') "
'		strSql = strSql & " 	or i.mwdiv not in ('M', 'W') "
'		strSql = strSql & " 	or i.deliveryType <> 1 "
'2018-01-29 15:00 김진영 하단 주석처리..
'		strSql = strSql & "		or ( ((i.mwdiv not in ('M', 'W')) or (i.deliveryType <> 1)) and i.makerid not in ('meaningless01', 'mandarinebrothers', 'fromamour', 'woolly02' ,'dalbampicnic') )"
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.itemdiv = '09' or i.cate_large = '999' or i.cate_large=''"
'		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "	'2019-06-17 | 2397378 윤현주 예외처리 요청
'		strSql = strSql & "		or ((i.sailyn = 'N') and ( (i.itemid <> 2397378) AND ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "

		strSql = strSql & "		or ( "
		strSql = strSql & "				(i.itemid <> 2397378) "
		strSql = strSql & "				AND ( "
		strSql = strSql & "					((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < "&CMAXMARGIN&" ) AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) <  "&CMAXMARGIN&" ))) "
		strSql = strSql & "					or ((i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) <  "&CMAXMARGIN&" )) "
		strSql = strSql & "				) "
		strSql = strSql & "		) "

		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " LEFT JOIN ( "
		strSql = strSql & " 	SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & " 	FROM db_etcmall.dbo.tbl_ssg_DispCate_mapping "
		strSql = strSql & " 	GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_ssg_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
'		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_outmall_keywords] as k on i.itemid = k.itemid and k.mallid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.ssgGoodNo is Not Null "		'등록 상품만
		'strSql = strSql & " and m.ssgStatCD = 7' "				'승인완료된 애들만 수정이 된다함..TEST 해봐야 됨
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
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
				FOneItem.FSsgPrice			= rsget("ssgPrice")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum		= rsget("orderMaxNum")
		End If
		rsget.Close
	End Sub
End Class

'SSG 상품코드 얻기
Function getSsgGoodNo(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 ssgGoodNo FROM db_etcmall.dbo.tbl_ssg_regitem WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
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
