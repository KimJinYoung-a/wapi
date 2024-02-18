<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 1000 ''초단위
'상품EP는 78번 DB를 바라보고, 판매EP는 77번DB를 바라본다
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Const MaxPage   = 1   ''maxpage 100..2015-05-07 김진영 변경
Const PageSize = 30000  ''3000->5000

Dim appPath : appPath = server.mappath("/outmall/coochaEP/") + "\"
Dim FileName: FileName = "HotDeal_temp.xml"
Dim newFileName: newFileName = "HotDeal.xml"
Dim fso, tFile

Function WriteMakeCoochaFile(tFile, arrList, byref iLastItemid )
    Dim intLoop,iRow
    Dim bufstr
    Dim itemid, deliv
    Dim lp, lp2, limitSu, itemcontent, ussinghtml, mainimage, mainimage2
    Dim tmpLastDeptNM, itemname, couponCash, couponPer
    iRow = UBound(arrList,2)

		bufstr = "<?xml version=""1.0"" encoding=""euc-kr"" ?>"&VbCRLF
		bufstr = bufstr&"<products>"&VbCRLF													'*## 전체 상품 정보의 시작/끝 | 전체 상품 정보의 시작부터 끝을 <products> ~ </products>로 선언하여 묶음
		tFile.WriteLine bufstr
	For intLoop=0 to iRow
		itemid			= arrList(1,intLoop)
		itemname		= "[텐바이텐]"&arrList(2,intLoop)
		itemname		= Replace(itemname,"&nbsp;","")
		itemname		= Replace(itemname,"&nbsp","")

		ussinghtml		= arrList(17,intLoop)
		itemcontent		= arrList(18,intLoop)
		mainimage		= "http://webimage.10x10.co.kr/image/main/" & GetImageSubFolderByItemid(itemid) & "/" & arrList(19,intLoop)
		mainimage2		= "http://webimage.10x10.co.kr/image/main2/" & GetImageSubFolderByItemid(itemid) & "/" & arrList(20,intLoop)

		If arrList(14,intLoop) = "Y" Then
			limitSu = arrList(15,intLoop) - arrList(16,intLoop)
		Else
			limitSu = "999999"
		End If

		If arrList(8,intLoop) = "Y" Then
			If isNull(arrList(21,intLoop)) OR isNull(arrList(22,intLoop)) Then
				couponCash	= ""
				couponPer	= ""
			Else
				couponCash	= Clng(arrList(21,intLoop))
				couponPer	= CInt(arrList(22,intLoop))
			End If
		Else
			couponCash	= ""
			couponPer	= ""
		End If

		bufstr = ""
		bufstr = bufstr&"<product>"&VbCRLF														'*## 개별 상품 정보의 시작/끝 | 개별상품의 정보를 입력 (만약, 상품이 1개 이상이라면 <product> ~ </product> 의 정보를 연속해서 입력)
		bufstr = bufstr&"<product_id>"&itemid&"</product_id>"&VbCRLF							'*## 상품 고유번호
		bufstr = bufstr&"<product_title><![CDATA["&itemname&"]]></product_title>"&VbCRLF		'*## 상품 타이틀 | 상품명
'		bufstr = bufstr&"<product_desc><![CDATA["&getContent(itemid, ussinghtml, itemcontent, mainimage, mainimage2)&"]]></product_desc>"&VbCRLF			'*##상품 상세 문구 | 상세설명
		bufstr = bufstr&"<product_desc></product_desc>"&VbCRLF									'*##상품 상세 문구 | 상세설명(쿠차에서 Null처리 하라함)
		bufstr = bufstr&"<product_url><![CDATA[http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=coocha]]></product_url>"&VbCRLF	'*## 쿠차-상품 연결 웹 URL
		bufstr = bufstr&"<mobile_url><![CDATA[http://m.10x10.co.kr/category/category_itemPrd.asp?itemid="&itemid&"&rdsite=coocha]]></mobile_url>"&VbCRLF	'*## 쿠차-상품 연결 모바일 URL
		bufstr = bufstr&"<product_url2><![CDATA[http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=coomoa]]></product_url2>"&VbCRLF	'*## 쿠폰모아-상품 연결 웹 URL
		bufstr = bufstr&"<mobile_url2><![CDATA[http://m.10x10.co.kr/category/category_itemPrd.asp?itemid="&itemid&"&rdsite=coomoa]]></mobile_url2>"&VbCRLF	'*## 쿠폰모아-상품 연결 모바일 URL
		bufstr = bufstr&"<sale_start><![CDATA["&arrList(23,intLoop)&"]]></sale_start>"&VbCRLF	'*## 상품 판매 일시
		bufstr = bufstr&"<sale_end><![CDATA["&arrList(24,intLoop)&"]]></sale_end>"&VbCRLF		'*## 상품 판매 종료일시
		bufstr = bufstr&"<price_normal>"&CLng(arrList(3,intLoop))&"</price_normal>"&VbCRLF		'*## 정상가격
		bufstr = bufstr&"<price_discount>"&Clng(arrList(5,intLoop))&"</price_discount>"&VbCRLF	'*## 할인가격
		bufstr = bufstr&"<discount_rate>"&CInt(arrList(6,intLoop))&"</discount_rate>"&VbCRLF	'*## 할인율
		bufstr = bufstr&"<coupon_use_start></coupon_use_start>"&VbCRLF							'*## 쿠폰 유효기간 시작일 | null / 협의내용 : 상품쿠폰이 아닌 기프티콘 같은 쿠폰이라함
		bufstr = bufstr&"<coupon_use_end></coupon_use_end>"&VbCRLF								'*## 쿠폰 유효기간 종료일 | null / 협의내용 : 상품쿠폰이 아닌 기프티콘 같은 쿠폰이라함
		bufstr = bufstr&"<now_use></now_use>"&VbCRLF											'*## 바로사용 가능여부 표시 | null
		bufstr = bufstr&"<category1><![CDATA["&arrList(11,intLoop)&"]]></category1>"&VbCRLF		'*## 1차 카테고리
		bufstr = bufstr&"<category2><![CDATA["&arrList(12,intLoop)&"]]></category2>"&VbCRLF		'*## 2차 카테고리
		bufstr = bufstr&"<category3><![CDATA["&arrList(13,intLoop)&"]]></category3>"&VbCRLF		'*## 3차 카테고리
		'bufstr = bufstr&"<category4></category4>"&VbCRLF										'4차 카테고리
		'bufstr = bufstr&"<buy_limit>0</buy_limit>"&VbCRLF										'할인 적용 최소 인원 수
		bufstr = bufstr&"<buy_max>"&limitSu&"</buy_max>"&VbCRLF									'*## 최대 구매자 수 | 한정이면 한정갯수, 아니면 999999
		bufstr = bufstr&"<buy_count>0</buy_count>"&VbCRLF										'*## 구매자 수
		bufstr = bufstr&"<free_shipping>"&Chkiif(arrList(10,intLoop) > 0, "C", "F") &"</free_shipping>"&VbCRLF	'## *배송비조건 입력 | 무료배송 : F or 조건부 무료배송 : A or 유료배송 : C, 배송상품이 아닌 경우 null값 처리
		bufstr = bufstr&"<shipping_fee>"&arrList(10,intLoop)&"</shipping_fee>"&VbCRLF							'## *배송비입력
		bufstr = bufstr&"<image_url1><![CDATA[http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) + "/" & arrList(7,intLoop)&"]]></image_url1>"&VbCRLF	'*## 대표이미지 URL 1
		bufstr = bufstr&"<coupon_price>"&couponCash&"</coupon_price>"&VbCRLF					'*## 쿠폰가격
		bufstr = bufstr&"<coupon_rate>"&couponPer&"</coupon_rate>"&VbCRLF						'*## 쿠폰할인율
		bufstr = bufstr&"<shop_name></shop_name>"&VbCRLF										'*## 상품 판매업소명
		bufstr = bufstr&"<shop_tel></shop_tel>"&VbCRLF											'*## 상품 판매업소 연락처
		bufstr = bufstr&"<shop_address></shop_address>"&VbCRLF									'*## 상품 판매점 주소
		bufstr = bufstr&"<shop_latitude></shop_latitude>"&VbCRLF								'*## 상품 판매점 위도 (x값)
		bufstr = bufstr&"<shop_longitude></shop_longitude>"&VbCRLF								'*## 상품 판매점 경도 (y값)
		bufstr = bufstr&"</product>"&VbCRLF
		tFile.WriteLine bufstr
		bufstr = ""
		iLastItemid = itemid
	Next
		'이 부분이 반복되어야 함 / 끝
		bufstr = bufstr&"</products>"
		tFile.WriteLine bufstr
		bufstr = ""
End function

Function getContent(iitemid, iussinghtml, iitemcontent, imainimage, imainimage2)
	Dim strRst, strSQL
	strRst = ("<div align=""center"">")

	Select Case iussinghtml
		Case "Y"
			strRst = strRst & (iitemcontent & "<br>")
		Case "H"
			strRst = strRst & (nl2br(iitemcontent) & "<br>")
		Case Else
			strRst = strRst & (nl2br(ReplaceBracket(iitemcontent)) & "<br>")
	End Select
	'# 추가 상품 설명이미지 접수
	strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & iitemid
	rsget.CursorLocation = adUseClient
	rsget.CursorType=adOpenStatic
	rsget.Locktype=adLockReadOnly
	rsget.Open strSQL, dbget
	If Not(rsget.EOF or rsget.BOF) Then
		Do Until rsget.EOF
			If rsget("imgType") = "1" Then
				strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(iitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%""><br>")
			End If
			rsget.MoveNext
		Loop
	End If
	rsget.Close

	'#기본 상품 설명이미지
	If ImageExists(imainimage) Then strRst = strRst & ("<img src=""" & imainimage & """ border=""0"" style=""width:100%""><br>")
	If ImageExists(imainimage2) Then strRst = strRst & ("<img src=""" & imainimage2 & """ border=""0"" style=""width:100%""><br>")
	strRst = strRst & ("</div>")
	getContent = strRst

End function

'// 상품이미지 존재여부 검사
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('coocha_HOT_ST')"
dbCTget.execute sqlStr

sqlStr ="[db_outmall].[dbo].[sp_Ten_Coocha_EPDataCount](2)"
rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
	FTotCnt = rsCTget(0)
END IF
rsCTget.close

'response.write FTotCnt&"<br>"

Dim i, ArrRows
Dim iLastItemid : iLastItemid=9999999

IF FTotCnt > 0 THEN
    FTotPage = CLNG(FTotCnt/PageSize)
    IF FTotPage<>(FTotCnt/PageSize) THEn FTotPage=FTotPage+1
    IF (FTotPage>MaxPage) THEn FTotPage=MaxPage
    Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )

'    For i=0 to FTotPage-1
		ArrRows = ""
		sqlStr ="[db_outmall].[dbo].[sp_Ten_Coocha_EPData]("&i+1&","&PageSize&",2,"&iLastItemid&")"
        rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
        IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
        	ArrRows = rsCTget.getRows()
        END IF
        rsCTget.close

        if isArray(ArrRows) then
            CALL WriteMakeCoochaFile(tFile, ArrRows, iLastItemid)
        end if

		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
		sqlStr = sqlStr + " (ref) values('coocha_HOT"&(i+1)*PageSize&"_"&iLastItemid&"')"
		dbCTget.execute sqlStr
 '   NExt

    tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
END IF

sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('coocha_HOT_ED')"
dbCTget.execute sqlStr

'2013-12-10 15:40 김진영 추가 TEMP파일을 원본 파일로 복사
Dim Newfso
Set Newfso = Server.CreateObject("Scripting.FileSystemObject")
	Newfso.CopyFile appPath & FileName ,appPath & newFileName
Set Newfso = nothing
response.write FTotCnt&"건 생성 ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->