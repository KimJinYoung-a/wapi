<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 1000  ''초단위
'상품EP는 78번 DB를 바라보고, 판매EP는 77번DB를 바라본다
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'' 위메프 파일 Make / 일별
Const MaxPage   = 200   ''maxpage 변경 40->50으로 2013-12-13수정, 50->60으로 2014-09-23 김진영 변경, 60->70으로 2014-10-08 변경 ,70->100 으로 2016-06-29
Const PageSize = 5000  ''3000->5000

Dim appPath : appPath = server.mappath("/outmall/wemakepriceEP/") + "\"
Dim FileName: FileName = "wemakePriceDailyEP_temp.txt"
Dim newFileName: newFileName = "wemakePriceDailyEP.txt"
Dim fso, tFile

Dim IsChangedEP : IsChangedEP = (request("epType")="chg")
If (IsChangedEP) Then
	FileName = "wemakePriceDailyChangedEP_temp.txt"
	newFileName = "wemakePriceDailyChangedEP.txt"
End If

Function WriteMakeWeMakePriceFile(tFile, arrList, isIsChangedEP,byref iLastItemid )
    Dim intLoop,iRow
    Dim bufstr, isMake
    Dim itemid,deliverytype, deliv, dispCash
    Dim ArrCateNM, ArrCateCD, CntNM, CntCD, lp, lp2
    Dim tmpLastDeptNM, itemname, evtText, isCouponDown, nvcpnVal
    iRow = UBound(arrList,2)

    For intLoop=0 to iRow
'이하는 전시카테고리
		itemid			= arrList(1,intLoop)
		deliverytype	= arrList(8,intLoop)
		deliv 			= arrList(19,intLoop)  ''배송비 /2000, 2500, 0

		IF isNULL(arrList(20,intLoop)) then  ''2013/12/07 추가
		    ArrCateNM		= ""
    		CntNM			= Split(ArrCateNM,",")
    		ArrCateCD		= ""
    		CntCD			= Split(ArrCateCD,",")
		Else
    		ArrCateNM		= Split(arrList(20,intLoop),"||")(0)
    		CntNM			= Split(ArrCateNM,",")
    		ArrCateCD		= Split(arrList(20,intLoop),"||")(1)
    		CntCD			= Split(ArrCateCD,",")
        End If
        itemname		= "[텐바이텐]"&arrList(2,intLoop)	'2017-12-22 11:40 김진영..상품명 앞에 텐바이텐 추가
		itemname		= Replace(itemname,"&nbsp;","")
		itemname		= Replace(itemname,"&nbsp","")
		itemname		= Replace(itemname,"""","")

		If (deliverytype = "7") Then deliv=-1
		If arrList(25,intLoop) = "06" OR arrList(25,intLoop) = "16" Then
			isMake = "Y"
		Else
			isMake = "N"
		End If

'		If arrList(26,intLoop) > 0 Then					'위메프 쿠폰값이 있으면...
'			dispCash	= CLNG(arrList(26,intLoop))
'			evtText		= "★위메프쇼핑 5% 추가할인★"
'			isCouponDown= "Y"
'			nvcpnVal	= "^5"
'		Else
			dispCash	= CLNG(arrList(3,intLoop))
'			If (Now() > #13/10/2017 00:00:00# AND Now() < #25/10/2017 20:59:59#) Then
'				evtText		= "텐바이텐 16주년 쿠폰쇼! 최대 30% 할인찬스"
'			Else
				evtText		= "▶ 구매 시 마일리지 적립 & 신규회원 가입 시 보너스쿠폰 증정!"
'			End If
			isCouponDown= ""
			nvcpnVal	= ""
'		End If

		bufstr = itemid & vbTab & Replace(itemname, vbTab, "") & vbTab & dispCash & vbTab & dispCash & vbTab  		'상품코드 | 상품명 | pc판매가격 | 모바일 판매가격
		bufstr = bufstr & CLNG(arrList(24,intLoop)) & vbTab & "http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=wmprc" & vbTab	'정가 | 상품URL
		bufstr = bufstr & "http://m.10x10.co.kr/category/category_itemPrd.asp?itemid="&itemid&"&rdsite=wmprc" & vbTab									'상품모바일URL
		bufstr = bufstr & "http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) & "/" & arrList(4,intLoop) & vbTab & "" & vbTab	'이미지URL | 추가 이미지URL

		For lp = 1 to Ubound(CntNM) + 1
			If lp>4 Then Exit For
			bufstr = bufstr & Replace(CntNM(lp-1),"&nbsp;","") & vbTab																						'제휴사 카테고리명(대/중/소/세)
		Next
		If lp < 5 Then
			For lp=lp to 4
				bufstr = bufstr & "" & vbTab
			Next
		End If

		bufstr = bufstr & "" & vbTab & "" & vbTab & "신상품" & vbTab & "" & vbTab & "" & vbTab & isMake & vbTab		'위메프카테고리 | 가격비교 페이지ID | 상품상태 | 해외구매대행여부 | 병행수입여부 | 주문제작상품여부
		bufstr = bufstr & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab			 	'판매방식구분 | 미성년자구매불가상품여부 | 상품구분 | 바코드 | 제품코드 | 모델명
		bufstr = bufstr & Replace(Replace(arrList(14,intLoop),"&nbsp;",""), vbTab, "") & vbTab & Replace(Replace(arrList(6,intLoop),"&nbsp;",""), vbTab, "") & vbTab & "" & vbTab & "" & vbTab		'브랜드 | 제조사 | 원산지 | 카드명/카드할인가격
		bufstr = bufstr & evtText & vbTab																			'이벤트

'		If (arrList(26,intLoop) > 0) THEN
'			bufstr = bufstr & nvcpnVal & vbTab																		'일반/제휴쿠폰
'		ElseIf (arrList(22,intLoop) <> "") THEN
		If (arrList(22,intLoop) <> "") THEN
			bufstr = bufstr & Replace(arrList(22,intLoop),"&nbsp;","") & vbTab
		Else
			bufstr = bufstr & "" & vbTab
		End if

		bufstr = bufstr & isCouponDown & vbTab																		'쿠폰다운로드필요여부
		bufstr = bufstr & "" & vbTab & arrList(11,intLoop) & vbTab & "" & vbTab & "" & vbTab						'카드무이자할부정보 | 포인트 | 별도설치비유무 | 사전매칭코드
		bufstr = bufstr & "" & vbTab	'검색태그..확인필요
		bufstr = bufstr & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & arrList(15,intLoop) & vbTab			'그룹ID | 제휴사상품ID | 코디상품ID | 최소구매수량 | 상품평 개수
		bufstr = bufstr & deliv & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab							'배송료 | 차등배송비여부 | 차등배송비내용 | 상품속성 | 구매옵션
		bufstr = bufstr & "" & vbTab & "" & vbTab																	'셀러ID | 주이용고객층
		IF (isIsChangedEP) then
			bufstr = bufstr & "" & vbTab & arrList(21,intLoop) & vbTab & arrList(10,intLoop)						'성별 | I,U,D | 상품정보생성시각
		Else
			bufstr = bufstr & ""	'성별
		End If
		tFile.WriteLine bufstr
		iLastItemid = itemid
    Next
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

''작성시간 체크
IF(IsChangedEP) then
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('wemakePrice_CH_ST')"
    dbCTget.execute sqlStr
else
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('wemakePrice_DY_ST')"
    dbCTget.execute sqlStr
end if


if (IsChangedEP) then
    sqlStr ="[db_outmall].[dbo].[usp_Ten_Outmall_Wemakeprice_EPDataCount](1)"
else
    sqlStr ="[db_outmall].[dbo].[usp_Ten_Outmall_Wemakeprice_EPDataCount]"
end if
dbCTget.CommandTimeout = 120 ''2019/01/16 추가
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
    IF (FTotPage>MaxPage) THEn
		FTotPage=MaxPage
		FTotCnt=MaxPage*PageSize
	end if

    Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )

	If (IsChangedEP) Then
		bufstr1 = "id"& vbTab &"title"& vbTab &"price_pc"& vbTab &"price_mobile"& vbTab &"normal_price"& vbTab &"link"& vbTab &"mobile_link"& vbTab &"image_link"& vbTab &"add_image_link"& vbTab &"category_name1"& vbTab &"category_name2"& vbTab &"category_name3"& vbTab &"category_name4"& vbTab &"naver_category"& vbTab &"naver_product_id"& vbTab &"condition"& vbTab &"import_flag"& vbTab &"parallel_import"& vbTab &"order_made"& vbTab &"product_flag"& vbTab &"adult"& vbTab &"goods_type"& vbTab &"barcode"& vbTab &"manufacture_define_number"& vbTab &"model_number"& vbTab &"brand"& vbTab &"maker"& vbTab &"origin"& vbTab &"card_event"& vbTab &"event_words"& vbTab &"coupon"& vbTab &"partner_coupon_download"& vbTab &"interest_free_event"& vbTab &"point"& vbTab &"installation_costs"& vbTab &"pre_match_code"& vbTab &"search_tag"& vbTab &"group_id"& vbTab &"vendor_id"& vbTab &"coordi_id"& vbTab &"minimum_purchase_quantity"& vbTab &"review_count"& vbTab &"shipping"& vbTab &"delivery_grade"& vbTab &"delivery_detail"& vbTab &"attribute"& vbTab &"option_detail"& vbTab &"seller_id"& vbTab &"age_group"& vbTab &"gender"& vbTab &"class"& vbTab &"update_time"
	Else
		bufstr1 = "id"& vbTab &"title"& vbTab &"price_pc"& vbTab &"price_mobile"& vbTab &"normal_price"& vbTab &"link"& vbTab &"mobile_link"& vbTab &"image_link"& vbTab &"add_image_link"& vbTab &"category_name1"& vbTab &"category_name2"& vbTab &"category_name3"& vbTab &"category_name4"& vbTab &"naver_category"& vbTab &"naver_product_id"& vbTab &"condition"& vbTab &"import_flag"& vbTab &"parallel_import"& vbTab &"order_made"& vbTab &"product_flag"& vbTab &"adult"& vbTab &"goods_type"& vbTab &"barcode"& vbTab &"manufacture_define_number"& vbTab &"model_number"& vbTab &"brand"& vbTab &"maker"& vbTab &"origin"& vbTab &"card_event"& vbTab &"event_words"& vbTab &"coupon"& vbTab &"partner_coupon_download"& vbTab &"interest_free_event"& vbTab &"point"& vbTab &"installation_costs"& vbTab &"pre_match_code"& vbTab &"search_tag"& vbTab &"group_id"& vbTab &"vendor_id"& vbTab &"coordi_id"& vbTab &"minimum_purchase_quantity"& vbTab &"review_count"& vbTab &"shipping"& vbTab &"delivery_grade"& vbTab &"delivery_detail"& vbTab &"attribute"& vbTab &"option_detail"& vbTab &"seller_id"& vbTab &"age_group"& vbTab &"gender"
	End If
	tFile.WriteLine bufstr1

    For i=0 to FTotPage-1
        ArrRows = ""
        if (IsChangedEP) then
            sqlStr ="[db_outmall].[dbo].[usp_Ten_Outmall_Wemakeprice_EPData]("&i+1&","&PageSize&",1,"&iLastItemid&")"
        else
            sqlStr ="[db_outmall].[dbo].[usp_Ten_Outmall_Wemakeprice_EPData]("&i+1&","&PageSize&",0,"&iLastItemid&")"
        end if
		dbCTget.CommandTimeout = 120 ''2019/01/16 추가
        rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
        IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
        	ArrRows = rsCTget.getRows()
        END IF
        rsCTget.close

        if isArray(ArrRows) then
            CALL WriteMakeWeMakePriceFile(tFile,ArrRows, IsChangedEP, iLastItemid)
        end if

        ''작성시간 체크
        IF(IsChangedEP) then
            sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
            sqlStr = sqlStr + " (ref) values('wemakePrice_CH_"&(i+1)*PageSize&"_"&iLastItemid&"')"
            dbCTget.execute sqlStr
        else
            sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
            sqlStr = sqlStr + " (ref) values('wemakePrice_DY_"&(i+1)*PageSize&"_"&iLastItemid&"')"
            dbCTget.execute sqlStr
        end if
    NExt

    tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
END IF

''작성시간 체크
IF(IsChangedEP) then
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('wemakePrice_CH_ED')"
    dbCTget.execute sqlStr
else
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('wemakePrice_DY_ED')"
    dbCTget.execute sqlStr
end if

'2013-12-10 15:40 김진영 추가 TEMP파일을 원본 파일로 복사
Dim Newfso
Set Newfso = Server.CreateObject("Scripting.FileSystemObject")
	Newfso.CopyFile appPath & FileName ,appPath & newFileName
Set Newfso = nothing
response.write FTotCnt&"건 생성 ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->