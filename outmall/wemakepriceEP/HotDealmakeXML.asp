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
Const PageSize = 100  ''3000->5000

Dim appPath : appPath = server.mappath("/outmall/wemakepriceEP/") + "\"
Dim FileName: FileName = "HotDeal_temp.xml"
Dim newFileName: newFileName = "HotDeal.xml"
Dim fso, tFile

Function WriteMakeWeMakePriceFile(tFile, arrList, byref iLastItemid )
    Dim intLoop,iRow
    Dim bufstr
    Dim itemid, deliv, ArrCateNM, CntNM
    Dim lp, lp2, limitSu, itemdivCost, freeDlvYn
    Dim tmpLastDeptNM, itemname, couponCash, couponPer
	Dim isDealItem
    iRow = UBound(arrList,2)

		bufstr = "<?xml version=""1.0"" encoding=""euc-kr"" ?>"&VbCRLF
		bufstr = bufstr&"<products>"&VbCRLF													'*## 전체 상품 정보의 시작/끝 | 전체 상품 정보의 시작부터 끝을 <products> ~ </products>로 선언하여 묶음
		tFile.WriteLine bufstr
	For intLoop=0 to iRow
		itemid			= arrList(1,intLoop)
		itemname		= "[텐바이텐]"&arrList(2,intLoop)
		itemname		= Replace(itemname,"&nbsp;","")
		itemname		= Replace(itemname,"&nbsp","")
		ArrCateNM		= Split(arrList(12,intLoop),"||")(0)
		CntNM			= Split(ArrCateNM,",")

		itemdivCost		= arrList(8,intLoop)
		If itemdivCost > 0 Then
			freeDlvYn = "Y"
		Else
			freeDlvYn = "N"
		End If

		If arrList(9,intLoop) = "Y" Then
			limitSu = arrList(10,intLoop) - arrList(11,intLoop)
		Else
			limitSu = "999999"
		End If

		isDealItem = (arrList(16,intLoop)="21")

		bufstr = ""
		bufstr = bufstr&"<product>"&VbCRLF
		bufstr = bufstr&"<mallPid>"&itemid&"</mallPid>"&VbCRLF									'*몰 상품 아이디(상품번호)
		bufstr = bufstr&"<poplrDgr>"&arrList(0,intLoop)&"</poplrDgr>"&VbCRLF					'*인기도
		bufstr = bufstr&"<prodName><![CDATA["&itemname&"]]></prodName>"&VbCRLF					'*상품명
		if (isDealItem) then
			bufstr = bufstr&"<prodUrl><![CDATA[http://www.10x10.co.kr/deal/deal.asp?itemid="&itemid&"&rdsite=wmprchot]]></prodUrl>"&VbCRLF				'*pc 랜딩 페이지
			bufstr = bufstr&"<mblProdUrl><![CDATA[http://m.10x10.co.kr/deal/deal.asp?itemid="&itemid&"&rdsite=wmprchot]]></mblProdUrl>"&VbCRLF		'*모바일 랜딩페이지
			bufstr = bufstr&"<prodImgUrl1><![CDATA[http://webimage.10x10.co.kr/image/basic/" & arrList(7,intLoop)&"]]></prodImgUrl1>"&VbCRLF	'*대표 이미지URL(정사각형)
		else
			bufstr = bufstr&"<prodUrl><![CDATA[http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=wmprchot]]></prodUrl>"&VbCRLF				'*pc 랜딩 페이지
			bufstr = bufstr&"<mblProdUrl><![CDATA[http://m.10x10.co.kr/category/category_itemPrd.asp?itemid="&itemid&"&rdsite=wmprchot]]></mblProdUrl>"&VbCRLF		'*모바일 랜딩페이지
			bufstr = bufstr&"<prodImgUrl1><![CDATA[http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) + "/" & arrList(7,intLoop)&"]]></prodImgUrl1>"&VbCRLF	'*대표 이미지URL(정사각형)
		end if
		bufstr = bufstr&"<nomlPrc>"&CLng(arrList(3,intLoop))&"</nomlPrc>"&VbCRLF				'*정상 가격
		bufstr = bufstr&"<mblDcPrc>"&Clng(arrList(5,intLoop))&"</mblDcPrc>"&VbCRLF				'*모바일 할인가격
		bufstr = bufstr&"<mblDcRt>"&CInt(arrList(6,intLoop))&"</mblDcRt>"&VbCRLF				'모바일 할인율
		bufstr = bufstr&"<freeDlvYn>"&freeDlvYn&"</freeDlvYn>"&VbCRLF							'*무료배송 여부
		bufstr = bufstr&"<saleCnt>"&arrList(15,intLoop)&"</saleCnt>"&VbCRLF						'*구매 수량(딜 시작 후 누적)
		For lp=1 to Ubound(CntNM)+1
			If lp>5 Then Exit For
			bufstr = bufstr&"<catNm"&lp&"><![CDATA["&Replace(CntNM(lp-1),"&nbsp;","")&"]]></catNm"&lp&">"&VbCRLF	'*제휴사 카테고리명(대/중/소/세분류 카테고리)
		Next
		If lp < 5 Then
			For lp=lp to 4
				bufstr = bufstr&"<catNm"&lp&"></catNm"&lp&">"&VbCRLF
			Next
		End If
		bufstr = bufstr&"<modelNm></modelNm>"&VbCRLF											'모델명
		bufstr = bufstr&"<brandNm><![CDATA["&arrList(13,intLoop)&"]]></brandNm>"&VbCRLF			'브랜드명
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

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('wemakePrice_HOT_ST')"
dbCTget.execute sqlStr

Dim i, ArrRows
Dim iLastItemid : iLastItemid=9999999

Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )

		ArrRows = ""
		sqlStr ="[db_outmall].[dbo].[usp_Ten_Outmall_Wemakeprice_EPData]("&i+1&","&PageSize&",2,"&iLastItemid&")"
		dbCTget.CommandTimeout = 120 ''2019/01/16 추가
	    rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
	    IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
	    	ArrRows = rsCTget.getRows()
	    END IF
	    rsCTget.close

	    if isArray(ArrRows) then
	        CALL WriteMakeWeMakePriceFile(tFile, ArrRows, iLastItemid)
	    end if

		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
		sqlStr = sqlStr + " (ref) values('wemakePrice_HOT"&(i+1)*PageSize&"_"&iLastItemid&"')"
		dbCTget.execute sqlStr

	    tFile.Close
	Set tFile = Nothing
Set fso = Nothing

sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('wemakePrice_HOT_ED')"
dbCTget.execute sqlStr

'2013-12-10 15:40 김진영 추가 TEMP파일을 원본 파일로 복사
Dim Newfso
Set Newfso = Server.CreateObject("Scripting.FileSystemObject")
	Newfso.CopyFile appPath & FileName ,appPath & newFileName
Set Newfso = nothing
response.write "100건 생성 ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->