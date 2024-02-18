<%@ Language="VBScript" CodePage=65001 %>
<% option explicit %>
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Server.ScriptTimeOut = 1000  ''초단위
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/JSON_noenc2.2.0.4.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim appPath
If application("Svr_Info")="Dev" Then
	appPath = server.mappath("/outmall/naverEP/") + "\"
Else
	appPath = server.mappath("/Files/naverEP/") + "\"
End If

Dim FileName: FileName = "naverDailyBookEP_temp.json"
Dim newFileName: newFileName = "naverDailyBookEP.json"
Dim IsChangedEP : IsChangedEP = (request("epType")="chg")
If (IsChangedEP) Then
	FileName = "naverChangedBookEP_temp.json"
	newFileName = "naverChangedBookEP.json"
End If
Dim sqlStr, FTotCnt

''작성시간 체크
If (IsChangedEP) Then
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr & " (ref) values('nvshop_NewBookCH_ST')"
    dbCTget.execute sqlStr
Else
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr & " (ref) values('nvshop_NewBookDY_ST')"
    dbCTget.execute sqlStr
End If

If (IsChangedEP) Then
	sqlStr = ""
	sqlStr = sqlStr & " EXEC db_outmall.[dbo].[sp_Ten_Naver_EPDataBook] 'COUNT', 1 "
	rsCTget.CursorLocation = adUseClient
	rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly
	IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
		FTotCnt = rsCTget(0)
	END IF
	rsCTget.close
Else
	sqlStr = ""
	sqlStr = sqlStr & " EXEC db_outmall.[dbo].[sp_Ten_Naver_EPDataBook] 'COUNT', 0 "
	rsCTget.CursorLocation = adUseClient
	rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly
	IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
		FTotCnt = rsCTget(0)
	END IF
	rsCTget.close
End If
'response.write FTotCnt&"<br>"

Dim i, ArrRows
If FTotCnt > 0 Then
	ArrRows = ""
	sqlStr = ""
	If (IsChangedEP) Then
		sqlStr = " EXEC db_outmall.[dbo].[sp_Ten_Naver_EPDataBook] 'LIST', 1 "
	Else
		sqlStr = " EXEC db_outmall.[dbo].[sp_Ten_Naver_EPDataBook] 'LIST', 0 "
	End If
	dbCTget.CommandTimeout = 120 ''2019/01/16 추가
	rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly
	If Not (rsCTget.EOF OR rsCTget.BOF) Then
		ArrRows = rsCTget.getRows()
	End If
	rsCTget.close

	If isArray(ArrRows) Then
		Dim intLoop, iRow, q, outputStream, BinaryStream
		Dim bufstr, obj, isbn13, isbn10, isbn
		iRow = UBound(ArrRows,2)

		Set outputStream = Server.CreateObject("ADODB.Stream")
			outputStream.Charset = "utf-8"
			outputStream.Open
			For intLoop=0 to iRow
				isbn = ""
				isbn13 = CSTR(ArrRows(2, intLoop))
				isbn10 = CSTR(ArrRows(3, intLoop))
				If Len(ArrRows(2, intLoop)) = 13 Then
					isbn = ArrRows(2, intLoop)
				Else
					isbn = ArrRows(3, intLoop)
				End If

				Set obj = jsObject()
					obj("id") = CSTR(ArrRows(0, intLoop))				'상품ID
					obj("goods_type") = CSTR(ArrRows(1, intLoop))		'상품 타입
					obj("isbn") = isbn									'ISBN코드
					obj("title") = ArrRows(4, intLoop)					'상품명
					obj("normal_price") = ArrRows(5, intLoop)			'도서 원가
					obj("price_pc") = ArrRows(6, intLoop)				'판매가
					obj("link") = "http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&ArrRows(0, intLoop)&"&utm_source=naver&utm_medium=organic&utm_campaign=shopping_w&term=nvshop_w&rdsite=nvshop_sp"		'상품 URL
					obj("mobile_link") = "http://m.10x10.co.kr/common/tenlanding.asp?urltype=item&itemid="&ArrRows(0, intLoop)&"&utm_source=naver&utm_medium=organic&utm_campaign=shopping_m&term=nvshop_m&rdsite=nvshop_sp"		'모바일 상품URL
					obj("category_name1") = CSTR(ArrRows(7, intLoop))	'제휴사 카테고리명(대분류)
					obj("image_link") = "http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(ArrRows(0, intLoop)) & "/" & ArrRows(8,intLoop)
					obj("publisher") = ArrRows(12, intLoop)				'출판사
					obj("shipping") = ArrRows(9, intLoop)				'배송비
					If IsChangedEP Then
						obj("update_time") = ArrRows(11, intLoop)		'요약 EP 등록 시간 정보
						obj("class") = ArrRows(10, intLoop)				'업데이트 구분
					End If
					bufstr = obj.jsString & VBCRLF
				Set obj = nothing
				outputStream.WriteText bufstr
			Next

			Set BinaryStream = CreateObject("adodb.stream")'오브젝트를 생성
				BinaryStream.Type = 1 
				BinaryStream.Mode = 3
				BinaryStream.Open
			outputStream.Position = 3 '쓰기 위치
			outputStream.CopyTo BinaryStream ' 오브젝트스트림 내용을 바이너리스트림으로 복사
			outputStream.Flush 
			outputStream.Close
				BinaryStream.SaveToFile appPath & FileName, 2
				BinaryStream.Close
			Set BinaryStream = nothing
		Set outputStream = Nothing
	End If
End If

''작성시간 체크
IF (IsChangedEP) Then
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr & " (ref) values('nvshop_BookCH_ED')"
    dbCTget.execute sqlStr
Else
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr & " (ref) values('nvshop_BookDY_ED')"
    dbCTget.execute sqlStr
End If

'2013-12-10 15:40 김진영 추가 TEMP파일을 원본 파일로 복사
If FTotCnt > 0 Then
Dim Newfso
Set Newfso = Server.CreateObject("Scripting.FileSystemObject")
	Newfso.CopyFile appPath & FileName, appPath & newFileName
Set Newfso = nothing
End If
response.write FTotCnt&"건 생성 ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->