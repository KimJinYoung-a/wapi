<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 1200  ''초단위
'상품EP는 78번 DB를 바라보고, 판매EP는 77번DB를 바라본다
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'' 네이버 지식쇼핑 파일 Make / 일별
Const MaxPage   = 300   ''maxpage 변경 40->50으로 2013-12-13수정, 50->60으로 2014-09-23 김진영 변경, 60->70으로 2014-10-08 변경 ,70->100 으로 2016-06-29
Const PageSize = 5000  ''3000->5000

Dim appPath
If application("Svr_Info")="Dev" Then
	appPath = server.mappath("/outmall/naverEP/") + "\"
Else
	appPath = server.mappath("/Files/naverEP/") + "\"
End If

Dim FileName: FileName = "naverNewVerDailyEP_temp.txt"
Dim newFileName: newFileName = "naverNewVerDailyEP.txt"
Dim fso, tFile

Dim IsChangedEP : IsChangedEP = (request("epType")="chg")
If (IsChangedEP) Then
	FileName = "naverNewVerChangedEP_temp.txt"
	newFileName = "naverNewVerChangedEP.txt"
End If

Function WriteMakeNaverFile(tFile, arrList, isIsChangedEP, byref iLastItemid)
    Dim intLoop,iRow
    Dim bufstr, isMake, basicImage, basic600Image, displayImageUrl, cardStr
    Dim itemid,deliverytype, deliv, dispCash
    Dim ArrCateNM, ArrCateCD, jaehu3depNM, CntNM, CntCD, lp, lp2
    Dim tmpLastDeptNM, itemname, evtText, isCouponDown, nvcpnVal, iNvCouponPro, iNvCouponValue, deliveryFixday, importFlagYN, adultType
    iRow = UBound(arrList,2)

    For intLoop=0 to iRow
'이하는 전시카테고리
		displayImageUrl = ""
		itemid			= arrList(1,intLoop)
		deliverytype	= arrList(8,intLoop)
		deliv 			= arrList(19,intLoop)  ''배송비 /2000, 2500, 0

		IF isNULL(arrList(20,intLoop)) then  ''2013/12/07 추가
		    ArrCateNM		= ""
    		CntNM			= Split(ArrCateNM,",")
    		ArrCateCD		= ""
    		CntCD			= Split(ArrCateCD,",")
    		jaehu3depNM		= ""
		else
    		ArrCateNM		= Split(arrList(20,intLoop),"||")(0)
    		CntNM			= Split(ArrCateNM,",")
    		ArrCateCD		= Split(arrList(20,intLoop),"||")(1)
    		CntCD			= Split(ArrCateCD,",")
    		jaehu3depNM		= Split(arrList(20,intLoop),"||")(2)

    		'2뎁쓰면 2뎁쓰명이 나오게 수정..2017-10-17 김진영
    		If Ubound(CntNM) = 1 then
				jaehu3depNM = Split(ArrCateNM, ",")(1)
	    	End If
        end if

		itemname		= arrList(2,intLoop)
		itemname		= Replace(itemname,"&nbsp;","")
		itemname		= Replace(itemname,"&nbsp","")
		itemname		= Replace(itemname,"""","")

		basicImage		= arrList(4,intLoop)
		basic600Image	= arrList(34,intLoop)
		cardStr			= arrList(35,intLoop)

		If basic600Image <> "" Then
			displayImageUrl = "http://webimage.10x10.co.kr/image/basic600/" & GetImageSubFolderByItemid(itemid) & "/" & arrList(4,intLoop)
		Else
			displayImageUrl = "http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) & "/" & arrList(4,intLoop)
		End If

		If itemid = "1831400" then	''2017-12-01 권태돈 차장님 요청..상품명 변경 실험
			itemname = "1.1M 프리미엄 북유럽 크리스마스트리 전구풀세트 (레드)_(540048)_트리"
		End If

		If (deliverytype = "7") Then deliv=-1
		If arrList(27,intLoop) = "06" OR arrList(27,intLoop) = "16" Then
			isMake = "Y"
		Else
			isMake = "N"
		End If

		If arrList(28,intLoop) > 0 Then					'네이버 쿠폰값이 있으면...쿠폰확인하여 쿠폰문구와 nvcpnVal의 값을 수정해야함
			dispCash	= CDBL(arrList(28,intLoop))
		Else
			dispCash	= CDBL(arrList(3,intLoop))
		End If

		'' 이벤트 문구 변경 2019/05/20
		'' 이벤트 문구 DB화 2019-09-25 김진영 추가
		evtText		= arrList(33,intLoop)
		isCouponDown= ""
		nvcpnVal	= ""

		'우선 순위 Depth3ItemName > Depth3MakerName > 전시카테고리명
		If (arrList(24,intLoop) <> "") OR (arrList(25,intLoop) <> "") Then
			IF (isIsChangedEP) then			'요약EP
				If arrList(21,intLoop) = "U" Then	'수정상태(U)
					If (arrList(25,intLoop) <> "") Then
						jaehu3depNM = db2html(arrList(25,intLoop))
					ElseIf (arrList(24,intLoop) <> "") Then
						jaehu3depNM = db2html(arrList(24,intLoop))
					End If
				End If
			Else
				If (arrList(24,intLoop) <> "") OR (arrList(25,intLoop) <> "") Then
					If (arrList(25,intLoop) <> "") Then
						jaehu3depNM = db2html(arrList(25,intLoop))
					ElseIf (arrList(24,intLoop) <> "") Then
						jaehu3depNM = db2html(arrList(24,intLoop))
					End If
				End If
			End If
		End If

		deliveryFixday = arrList(31,intLoop)
		If deliveryFixday = "G" Then
			importFlagYN = "Y"
		Else
			importFlagYN = ""
		End If

		''2019/04/25
		adultType = arrList(32,intLoop)
		if (adultType="1" or adultType="2") then
			adultType="Y"
		else
			adultType=""
		end if

		bufstr = itemid & vbTab & Replace(itemname, vbTab, "") & CHKIIF(jaehu3depNM="",""," ") & jaehu3depNM & vbTab & dispCash & vbTab & dispCash & vbTab  		'상품코드 | 상품명 | pc판매가격 | 모바일 판매가격
'2019-04-11 하단 주석처리
'		bufstr = bufstr & CLNG(arrList(26,intLoop)) & vbTab & "http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=nvshop_sp&utm_source=organic&utm_medium=shopping_w&utm_campaign=nvshop_w&term=nvshop" & vbTab	'정가 | 상품URL
'		bufstr = bufstr & "http://m.10x10.co.kr/category/category_itemPrd.asp?itemid="&itemid&"&rdsite=nvshop_sp&utm_source=organic&utm_medium=shopping_m&utm_campaign=nvshop_m&term=nvshop" & vbTab									'상품모바일URL
'2019-04-11 남궁병준님 요청 하단으로 utmParam 변경
		bufstr = bufstr & CLNG(arrList(26,intLoop)) & vbTab & "http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&utm_source=naver&utm_medium=organic&utm_campaign=shopping_w&term=nvshop_w&rdsite=nvshop_sp" & vbTab	'정가 | 상품URL
		'// 모바일은 브랜치로 연결
		bufstr = bufstr & "http://m.10x10.co.kr/common/tenlanding.asp?urltype=item&itemid="&itemid&"&utm_source=naver&utm_medium=organic&utm_campaign=shopping_m&term=nvshop_m&rdsite=nvshop_sp" & vbTab									'상품모바일URL
		bufstr = bufstr & displayImageUrl & vbTab & "" & vbTab	'이미지URL | 추가 이미지URL

		For lp = 1 to Ubound(CntNM) + 1
			If lp>4 Then Exit For
			bufstr = bufstr & Replace(CntNM(lp-1),"&nbsp;","") & vbTab																						'제휴사 카테고리명(대/중/소/세)
		Next
		If lp < 5 Then
			For lp=lp to 4
				bufstr = bufstr & "" & vbTab
			Next
		End If

		if (itemid="2142647") then  ''원부매핑테스트
		bufstr = bufstr & "" & vbTab & "15883309361" & vbTab & "신상품" & vbTab & importFlagYN & vbTab & "" & vbTab & isMake & vbTab		'네이버카테고리 | 가격비교 페이지ID | 상품상태 | 해외구매대행여부 | 병행수입여부 | 주문제작상품여부
		elseif (itemid="2091984") then  ''원부매핑테스트
		bufstr = bufstr & "" & vbTab & "15558147004" & vbTab & "신상품" & vbTab & importFlagYN & vbTab & "" & vbTab & isMake & vbTab		'네이버카테고리 | 가격비교 페이지ID | 상품상태 | 해외구매대행여부 | 병행수입여부 | 주문제작상품여부
		elseif (itemid="1864887") then  ''원부매핑테스트
		bufstr = bufstr & "" & vbTab & "13874181171" & vbTab & "신상품" & vbTab & importFlagYN & vbTab & "" & vbTab & isMake & vbTab		'네이버카테고리 | 가격비교 페이지ID | 상품상태 | 해외구매대행여부 | 병행수입여부 | 주문제작상품여부
		elseif (itemid="2117554") then  ''원부매핑테스트 20190425->0000000000 으로변경해봄 // 카테고리 가구1depth
		bufstr = bufstr & "50000004" & vbTab & "0000000000" & vbTab & "신상품" & vbTab & importFlagYN & vbTab & "" & vbTab & isMake & vbTab		'네이버카테고리 | 가격비교 페이지ID | 상품상태 | 해외구매대행여부 | 병행수입여부 | 주문제작상품여부
		else
		bufstr = bufstr & "" & vbTab & "" & vbTab & "신상품" & vbTab & importFlagYN & vbTab & "" & vbTab & isMake & vbTab		'네이버카테고리 | 가격비교 페이지ID | 상품상태 | 해외구매대행여부 | 병행수입여부 | 주문제작상품여부
		end if
		bufstr = bufstr & "" & vbTab & adultType & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab			 	'판매방식구분 | 미성년자구매불가상품여부 | 상품구분 | 바코드 | 제품코드 | 모델명
		bufstr = bufstr & Replace(Replace(arrList(14,intLoop),"&nbsp;",""), vbTab, "") & vbTab & Replace(Replace(arrList(6,intLoop),"&nbsp;",""), vbTab, "") & vbTab & "" & vbTab	'브랜드 | 제조사 | 원산지
		bufstr = bufstr & cardStr & vbTab		 '카드명/카드할인가격
		bufstr = bufstr & evtText & vbTab																			'이벤트

		If (arrList(28,intLoop) > 0) THEN
			bufstr = bufstr & nvcpnVal & vbTab																		'일반/제휴쿠폰
		ElseIf (arrList(22,intLoop) <> "") THEN
			bufstr = bufstr & Replace(arrList(22,intLoop),"&nbsp;","") & vbTab
		Else
			bufstr = bufstr & "" & vbTab
		End if

		bufstr = bufstr & isCouponDown & vbTab																		'쿠폰다운로드필요여부
		bufstr = bufstr & "" & vbTab & arrList(11,intLoop) & vbTab & "" & vbTab & "" & vbTab						'카드무이자할부정보 | 포인트 | 별도설치비유무 | 사전매칭코드
		bufstr = bufstr & "" & vbTab	'검색태그..확인필요
		bufstr = bufstr & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & arrList(15,intLoop) & vbTab			'그룹ID | 제휴사상품ID | 코디상품ID | 최소구매수량 | 상품평 개수

		if (itemid="3672138") then  ''원부매핑테스트
			bufstr = bufstr & deliv & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "블루^"&dispCash&"|화이트^"&dispCash&"|그레이^"&dispCash&"|레몬^"&dispCash&"" & vbTab							'배송료 | 차등배송비여부 | 차등배송비내용 | 상품속성 | 구매옵션
		else
			bufstr = bufstr & deliv & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab							'배송료 | 차등배송비여부 | 차등배송비내용 | 상품속성 | 구매옵션
		end if

		bufstr = bufstr & "" & vbTab & "" & vbTab																	'셀러ID | 주이용고객층
		IF (isIsChangedEP) then
			bufstr = bufstr & "" & vbTab & arrList(36,intLoop) & vbTab & arrList(21,intLoop) & vbTab & arrList(10,intLoop)						'성별 | ISBN | I,U,D | 상품정보생성시각
		Else
			bufstr = bufstr & "" & vbTab & arrList(36,intLoop)	'성별 | ISBN
		End If
		tFile.WriteLine bufstr
		iLastItemid = itemid
    Next
End function

Dim sqlStr, logIdx
Dim FTotCnt, FTotPage, FCurrPage

''작성시간 체크
If (IsChangedEP) Then
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr & " (ref) values('nvshop_NewCH_ST')"
    dbCTget.execute sqlStr
Else
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr & " (ref) values('nvshop_NewDY_ST')"
    dbCTget.execute sqlStr
End If

If (IsChangedEP) Then
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	sqlStr = sqlStr & " FROM db_outmall.[dbo].[tbl_naverchgep_log] with (nolock) "
	rsCTget.CursorLocation = adUseClient
	rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly
	IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
		FTotCnt = rsCTget(0)
	END IF
	rsCTget.close

    sqlStr ="[db_outmall].[dbo].[usp_Ten_NaverEP_JobLog_Set] ('150', '네이버EP 요약 생성', '"& FTotCnt &"', 'Jenkins Batch') "
	dbCTget.CommandTimeout = 120 ''2019/01/16 추가
	rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
	IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
		logIdx = rsCTget(0)
	END IF
	rsCTget.close
Else
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	sqlStr = sqlStr & " FROM db_outmall.[dbo].[tbl_naverep_log] with (nolock) "
	rsCTget.CursorLocation = adUseClient
	rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly
	IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
		FTotCnt = rsCTget(0)
	END IF
	rsCTget.close

	sqlStr = ""
	sqlStr = sqlStr & " IF NOT EXISTS ( "
	sqlStr = sqlStr & " 	SELECT TOP 1 madedt "
	sqlStr = sqlStr & " 	FROM db_outmall.[dbo].[tbl_EpShop_Report] "
	sqlStr = sqlStr & " 	WHERE madedt = CONVERT(Date, GETDATE()) "
	sqlStr = sqlStr & " 	and mallid = 'nvshop' "
	sqlStr = sqlStr & " ) "
	sqlStr = sqlStr & " BEGIN "
	sqlStr = sqlStr & " 	INSERT INTO db_outmall.[dbo].[tbl_EpShop_Report] (mallid, madedt, epCnt) VALUES ('nvshop', CONVERT(Date, GETDATE()), '"& FTotCnt &"') "
	sqlStr = sqlStr & " END "
	dbCTget.execute sqlStr

    sqlStr ="[db_outmall].[dbo].[usp_Ten_NaverEP_JobLog_Set] ('100', '네이버EP 전체 생성', '"& FTotCnt &"', 'Jenkins Batch') "
	dbCTget.CommandTimeout = 120 ''2019/01/16 추가
	rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
	IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
		logIdx = rsCTget(0)
	END IF
	rsCTget.close
End If
'response.write FTotCnt&"<br>"

Dim i, ArrRows, bufstr1
Dim iLastItemid : iLastItemid=9999999

IF FTotCnt > 0 THEN
	FTotPage = CLNG(FTotCnt/PageSize)
	IF FTotPage<>(FTotCnt/PageSize) THEn FTotPage=FTotPage+1
	IF (FTotPage>MaxPage) THEn
		FTotPage=MaxPage
		FTotCnt=MaxPage*PageSize
	ENd IF
    Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )

	If (IsChangedEP) Then
		bufstr1 = "id"& vbTab &"title"& vbTab &"price_pc"& vbTab &"price_mobile"& vbTab &"normal_price"& vbTab &"link"& vbTab &"mobile_link"& vbTab &"image_link"& vbTab &"add_image_link"& vbTab &"category_name1"& vbTab &"category_name2"& vbTab &"category_name3"& vbTab &"category_name4"& vbTab &"naver_category"& vbTab &"naver_product_id"& vbTab &"condition"& vbTab &"import_flag"& vbTab &"parallel_import"& vbTab &"order_made"& vbTab &"product_flag"& vbTab &"adult"& vbTab &"goods_type"& vbTab &"barcode"& vbTab &"manufacture_define_number"& vbTab &"model_number"& vbTab &"brand"& vbTab &"maker"& vbTab &"origin"& vbTab &"card_event"& vbTab &"event_words"& vbTab &"coupon"& vbTab &"partner_coupon_download"& vbTab &"interest_free_event"& vbTab &"point"& vbTab &"installation_costs"& vbTab &"pre_match_code"& vbTab &"search_tag"& vbTab &"group_id"& vbTab &"vendor_id"& vbTab &"coordi_id"& vbTab &"minimum_purchase_quantity"& vbTab &"review_count"& vbTab &"shipping"& vbTab &"delivery_grade"& vbTab &"delivery_detail"& vbTab &"attribute"& vbTab &"option_detail"& vbTab &"seller_id"& vbTab &"age_group"& vbTab &"gender"& vbTab &"isbn"& vbTab &"class"& vbTab &"update_time"
	Else
		bufstr1 = "id"& vbTab &"title"& vbTab &"price_pc"& vbTab &"price_mobile"& vbTab &"normal_price"& vbTab &"link"& vbTab &"mobile_link"& vbTab &"image_link"& vbTab &"add_image_link"& vbTab &"category_name1"& vbTab &"category_name2"& vbTab &"category_name3"& vbTab &"category_name4"& vbTab &"naver_category"& vbTab &"naver_product_id"& vbTab &"condition"& vbTab &"import_flag"& vbTab &"parallel_import"& vbTab &"order_made"& vbTab &"product_flag"& vbTab &"adult"& vbTab &"goods_type"& vbTab &"barcode"& vbTab &"manufacture_define_number"& vbTab &"model_number"& vbTab &"brand"& vbTab &"maker"& vbTab &"origin"& vbTab &"card_event"& vbTab &"event_words"& vbTab &"coupon"& vbTab &"partner_coupon_download"& vbTab &"interest_free_event"& vbTab &"point"& vbTab &"installation_costs"& vbTab &"pre_match_code"& vbTab &"search_tag"& vbTab &"group_id"& vbTab &"vendor_id"& vbTab &"coordi_id"& vbTab &"minimum_purchase_quantity"& vbTab &"review_count"& vbTab &"shipping"& vbTab &"delivery_grade"& vbTab &"delivery_detail"& vbTab &"attribute"& vbTab &"option_detail"& vbTab &"seller_id"& vbTab &"age_group"& vbTab &"gender"& vbTab &"isbn"
	End If
	tFile.WriteLine bufstr1

    For i=0 to FTotPage-1
        ArrRows = ""
        If (IsChangedEP) Then
			sqlStr ="[db_outmall].[dbo].[sp_Ten_Naver_EPDataChgInsert_MakeText] ("&i+1&","&PageSize&")"
        Else
            sqlStr ="[db_outmall].[dbo].[sp_Ten_Naver_EPDataInsert_MakeText] ("&i+1&","&PageSize&")"
        End If
		dbCTget.CommandTimeout = 120 ''2019/01/16 추가
        rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
        IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
        	ArrRows = rsCTget.getRows()
        END IF
        rsCTget.close

        if isArray(ArrRows) then
            CALL WriteMakeNaverFile(tFile, ArrRows, IsChangedEP, iLastItemid)
        end if

        ''작성시간 체크
        IF(IsChangedEP) then
            sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
			sqlStr = sqlStr & " (ref) values('nvshop_NewCH_"&(i+1) * PageSize & "')"
            dbCTget.execute sqlStr
        Else
			sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
			sqlStr = sqlStr & " (ref) values('nvshop_NewDY_"&(i+1) * PageSize & "')"
			dbCTget.execute sqlStr
        End If
    Next

    tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
END IF

''작성시간 체크
IF(IsChangedEP) then
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr & " (ref) values('nvshop_NewCH_ED')"
    dbCTget.execute sqlStr
else
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr & " (ref) values('nvshop_NewDY_ED')"
    dbCTget.execute sqlStr

    sqlStr = ""
    sqlStr = sqlStr & " UPDATE db_outmall.[dbo].[tbl_EpShop_Report] "
	sqlStr = sqlStr & " SET completedt = GETDATE() "
	sqlStr = sqlStr & " WHERE mallid = 'nvshop' "
	sqlStr = sqlStr & " and madedt = CONVERT(Date, GETDATE()) "
    dbCTget.execute sqlStr
end if

sqlStr ="[db_outmall].[dbo].[usp_Ten_NaverEP_JobLog_Upd] ('"& logIdx &"', 'C', '') "
dbCTget.execute sqlStr

'2013-12-10 15:40 김진영 추가 TEMP파일을 원본 파일로 복사
Dim Newfso
Set Newfso = Server.CreateObject("Scripting.FileSystemObject")
	Newfso.CopyFile appPath & FileName ,appPath & newFileName
Set Newfso = nothing
response.write FTotCnt&"건 생성 ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->