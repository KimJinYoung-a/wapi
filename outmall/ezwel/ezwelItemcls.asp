<%
CONST CMAXMARGIN = 14.9
CONST CMALLNAME = "ezwel"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CEzwelMARGIN = 10									'이지웰페어 마진 10%
CONST cspCd		= "10040413"							'CP업체코드(이지웰 발급)
CONST crtCd		= "8e5a6dbdd27efb49fc600c293884ef47"	'보안코드(이지웰 발급)
CONST cspDlvrId	= "10040413"							'배송처코드

Class CEzwelItem
	Public FItemid
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
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FezwelStatCD
	Public FinfoDiv
	Public FDeliveryType
	Public FdepthCode
	Public FbasicimageNm
	Public FezwelGoodNo
	Public Fezwelprice
	Public FezwelSellYn
	Public FregedOptCnt
	Public FaccFailCNT
	Public FlastErrStr
	Public FrequireMakeDay
	Public Fsafetyyn
	Public FsafetyDiv
    Public FsafetyNum
    Public FmaySoldOut
	Public FAdultType

    Public Fregitemname
    Public FregImageName
	Public FOrderMaxNum

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "999" Then
			getOrderMaxNum = 999
		End If
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
		getKeywords = Server.URLEncode(strRst)
	End Function

	public Function getNewKeywords()
		Dim strRst
		strRst = FKeywords
		strRst = replace(strRst, "인기", "")
		strRst = replace(strRst, "인치", "")
		strRst = replace(strRst, "모기퇴치", "")
		If strRst = "" Then
			strRst = "텐바이텐"
		End If
		getNewKeywords = strRst
	End Function

	'// 품절여부
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
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
				If (FLimitYN <> "Y") Then optLimit = 999

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
' o11st.FOneItem.FLimityn
		i = 0
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
	End Function

	Public Function fngetMustPrice
		Dim strRst, GetTenTenMargin, sqlStr, specialPrice
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

		If specialPrice <> "" Then
			fngetMustPrice = specialPrice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If GetTenTenMargin < CMAXMARGIN Then
				fngetMustPrice = Forgprice
			Else
				fngetMustPrice = FSellCash
			End If
		End If
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, sqlStr, specialPrice
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

		If specialPrice <> "" Then
			MustPrice = specialPrice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If GetTenTenMargin < CMAXMARGIN Then
				MustPrice = Forgprice
			Else
				MustPrice = FSellCash
			End If
		End If
	End Function

	'// Ezwel 판매여부 반환
	Public Function getEzwelSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getEzwelSellYn = "Y"
			Else
				getEzwelSellYn = "N"
			End If
		Else
			getEzwelSellYn = "N"
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

	Public Function getLimitEzwelEa()
		Dim ret
		If FLimitYn = "Y" Then
			ret = FLimitNo - FLimitSold - 5
			If ret > 10000 Then
				ret = 10000
			End If
		Else
			ret = 10000
		End If

		If (ret < 1) Then ret = 0
		getLimitEzwelEa = ret
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

    Function getEzwelAddSuplyPrice(addprice)
		getEzwelAddSuplyPrice = CLNG((addprice)  * (100-CEzwelMARGIN) / 100)
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
'				strSql = strSql & " 	and optaddprice=0 "
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

	Function getItemNameFormat()
		Dim buf
		buf = "[텐바이텐]"&replace(FItemName,"'","")		'최초 상품명 앞에 [텐바이텐] 이라고 붙임
		buf = replace(buf,"&#8211;","-")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","프로")
		buf = replace(buf,"[무료배송]","")
		buf = replace(buf,"[무료 배송]","")
		getItemNameFormat = buf
	End Function

	Public Function IsFreeBeasong()
		IsFreeBeasong = False
		If (FdeliveryType=2) or (FdeliveryType=4) or (FdeliveryType=5) then				'2(텐무), 4,5(업무)
			IsFreeBeasong = True
		End If
'		If (FSellcash>=30000) then IsFreeBeasong=True
		If (FdeliveryType=9) Then														'업체조건
'			If (Clng(FSellcash) >= Clng(FdefaultfreeBeasongLimit)) then
'				IsFreeBeasong=True
'			End If
			IsFreeBeasong = False
		End If
    End Function

	Public Function getBrandCode(v)
		Dim strSql
		strSql = strSql & " SELECT TOP 1 brandCd "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_ezwel_brandList] "
		'strSql = strSql & " WHERE brandNm like '%"& html2db(v) &"%' "
		strSql = strSql & " WHERE brandNm = '"& html2db(v) &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If (Not rsget.EOF) Then
			getBrandCode = rsget("brandCd")
		Else
			getBrandCode = "143289"
		End If
		rsget.Close
	End Function

	Public Function getMafcCode(v)
		Dim strSql
		strSql = strSql & " SELECT TOP 1 mafcCd "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_ezwel_mafcList] "
		'strSql = strSql & " WHERE mafcNm like '%"& html2db(v) &"%' "
		strSql = strSql & " WHERE mafcNm = '"& html2db(v) &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If (Not rsget.EOF) Then
			getMafcCode = rsget("mafcCd")
		Else
			getMafcCode = "184231"
		End If
		rsget.Close
	End Function

	'상품설명 파라메터 생성
	Public Function getEzwelItemContParam()
		Dim strRst, strSQL,strRst2
		strRst = ("<div align=""center"">")
		strRst = strRst & ("<p><center><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_ezwel.jpg""></center></p><br>")
		Fitemcontent = rpTxt(Fitemcontent)

		If ForderComment <> "" Then
			strRst = strRst & "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
		End If

		'#기본 상품설명
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "<br>")
			Case Else
				strRst = strRst & (nl2br(ReplaceBracket(Fitemcontent)) & "<br>")
		End Select
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
				If rsget("imgType") = "1" Then
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%""><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#기본 상품 설명이미지
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%""><br>")

		'#배송 주의사항
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ezwel.jpg"">")
		strRst = strRst & ("</div>")

		strRst = replace(replace(strRst, "<script", ""), "</script>", "")
		getEzwelItemContParam = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','ezwel') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strRst2 = rpTxt(rsget("textVal"))
		'response.end
			strRst = ("<div align=""center"">")
			strRst = strRst & ("<p><center><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_ezwel.jpg""></center></p><br>")
			strRst = strRst & strRst2
			strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ezwel.jpg"">")
			strRst = strRst & ("</div>")
			getEzwelItemContParam = strRst
		End If
		rsget.Close

	End Function

	'상품설명 파라메터 생성
	Public Function getEzwelNewItemContParam()
		Dim strRst, strSQL,strRst2
		strRst = ("<div align=""center"">")
		strRst = strRst & ("<p><center><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_ezwel.jpg""></center></p><br>")

		If ForderComment <> "" Then
			strRst = strRst & "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
		End If

		'#기본 상품설명
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "<br>")
			Case Else
				strRst = strRst & (nl2br(ReplaceBracket(Fitemcontent)) & "<br>")
		End Select
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
				If rsget("imgType") = "1" Then
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%""><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#기본 상품 설명이미지
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%""><br>")

		'#배송 주의사항
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ezwel.jpg"">")
		strRst = strRst & ("</div>")

		strRst = replace(replace(strRst, "<script", ""), "</script>", "")
		getEzwelNewItemContParam = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','ezwel') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strRst2 = rsget("textVal")
		'response.end
			strRst = ("<div align=""center"">")
			strRst = strRst & ("<p><center><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_ezwel.jpg""></center></p><br>")
			strRst = strRst & strRst2
			strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ezwel.jpg"">")
			strRst = strRst & ("</div>")
			getEzwelNewItemContParam = strRst
		End If
		rsget.Close
	End Function

	'// 상품등록: 상품추가이미지 파라메터 생성
	Public Function getEzwelAddImageParam()
		Dim strRst, strSQL, i
		strRst = ""
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If

		strRst = strRst &"	<imgPath><![CDATA["&FbasicImage&"]]></imgPath>"		'메인이미지경로 | ex)http://www.ezwel.com/img/goods1.gif
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		'추가이미지경로1~3
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					strRst = strRst &"	<imgPath"&i&"><![CDATA[http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") &"]]></imgPath"&i&">"
				End If
				rsget.MoveNext
				If i >= 3 Then Exit For
			Next

		End If
		rsget.Close
		getEzwelAddImageParam = strRst
	End Function

	Public Function getEzwelNewAddImageParam(obj)
		Dim strSQL, i
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If
		obj("imgPath") = FbasicImage						'메인이미지경로

		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		'추가이미지경로1~3
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					obj("imgPath"&i&"") = "http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")		'추가이미지경로1~3
				End If
				rsget.MoveNext
				If i >= 3 Then Exit For
			Next
		End If
		rsget.Close
	End Function

	'상품품목정보
    public function getEzwelItemInfoCd()
		Dim buf1, buf2, buf3, strSQL, mallinfoCd, infoContent, mallinfodiv
		strSQL = ""
		strSQL = strSQL & " SELECT top 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN IC.safetyNum " & vbcrlf
		strSql = strSql & "		 WHEN (M.infoCd='00000') AND (IC.safetyyn <> 'Y' ) THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00001') THEN '상세페이지참고' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='10000') THEN '공정거래위원회 고시(소비자분쟁해결기준)에 의거하여 보상해 드립니다.' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='J' and F.chkDiv='N' THEN '해당없음' " & vbcrlf
		strSql = strSql & "		 WHEN c.infotype='K' and F.chkDiv='N' THEN '해당없음' " & vbcrlf
		'strSQL = strSQL & " 	 WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035'  " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN '텐바이텐 상품문의 / Q&A 작성'  " & vbcrlf
		strSql = strSql & "		 WHEN LEN( isNull(F.infocontent, '')) < 2 THEN '상품 상세 참고' " & vbcrlf
		strSQL = strSQL & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd  " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"'  " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='" & FItemid &"' " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'ezwel' and IC.itemid='"&FItemid&"'  " & vbcrlf
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		''mallinfodiv = "10" & rsget("mallinfodiv")  '' 이동 eastone 2016/08/17
		If Not(rsget.EOF or rsget.BOF) then
		    mallinfodiv = "10" & rsget("mallinfodiv")
			If mallinfodiv = "1047" Then
				mallinfodiv = "1039"
			ElseIf mallinfodiv = "1048" Then
				mallinfodiv = "1040"
			End If

			buf1 = "<goodsGrpCd>"&mallinfodiv&"</goodsGrpCd>"		'##*상품고시 코드 | 별도첨부
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")

				If FMakerid = "indigoshop" and (rsget("infocd") = "35002") Then
					infoContent = ".."
				end if

				If rsget("infocontent") = "" or isnull(infocontent) Then
					infoContent = "상세페이지 참고"
				End If

				buf2 = buf2 & " 		<arrLayoutDesc><![CDATA["& Server.URLEncode(infoContent) &"]]></arrLayoutDesc>"
				buf2 = buf2 & " 		<arrLayoutSeq>"&mallinfoCd&"</arrLayoutSeq>"
				rsget.MoveNext
			Loop
			buf3 = buf1 & buf2
		End If
		rsget.Close
		getEzwelItemInfoCd = buf3
	End Function

	Public Function getEzwelItemNewInfoCd(obj)
		Dim strSQL, mallinfoCd, infoContent, mallinfodiv, i
		strSQL = "EXEC [db_etcmall].[dbo].[usp_API_Ezwel_InfoCodeMap_Get] " & FItemID
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
		    mallinfodiv = "10" & rsget("mallinfodiv")
			If mallinfodiv = "1047" Then
				mallinfodiv = "1039"
			ElseIf mallinfodiv = "1048" Then
				mallinfodiv = "1040"
			End If
			obj("goodsGrpCd") = mallinfodiv							'#상품고시 코드
			Set obj("arrLayoutDesc") = jsArray()					'#상품고시 내용
			Set obj("arrLayoutSeq") = jsArray()						'#상품고시 항목 순번
			i = 0
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")

				If FMakerid = "indigoshop" and (rsget("infocd") = "35002") Then
					infoContent = ".."
				end if

				If rsget("infocontent") = "" or isnull(infocontent) Then
					infoContent = "상세페이지 참고"
				End If
				obj("arrLayoutDesc")(i) = infoContent
				obj("arrLayoutSeq")(i) = mallinfoCd
				i = i + 1
				rsget.MoveNext
			Loop
		End If
		rsget.Close
	End Function

   Public Function getEzwelOptionParam()
		Dim strSql, strRst, i, optLimit, sellOptcnt
    	Dim buf, optDc, itemsu, addprice, addbuyprice, optTaxCk, optTax, optUsingCk, optUsing

    	buf = ""
		If FoptionCnt>0 then
			strSql = ""
			strSql = strSql &  "SELECT COUNT(*) as cnt "
			strSql = strSql & " FROM [db_item].[dbo].tbl_item_option with (nolock) "
			strSql = strSql & " where itemid=" & FItemid
			strSql = strSql & " and optsellyn='Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				sellOptcnt = rsget("cnt")
			rsget.Close

			If sellOptcnt > 0 Then
				strSql = ""
				strSql = strSql &  "SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

				optDc = ""
				optLimit = ""
				If FVatInclude = "N" Then
					optTaxCk = "N"
				Else
					optTaxCk = "Y"
				End If

				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
						optLimit = rsget("optLimit")
						optLimit = optLimit-5
						If (optLimit < 1) Then optLimit = 0
						If (FLimitYN <> "Y") Then optLimit = 999   ''2013/06/12 재고관리여부 모두 Y로 변경 되므로
						optUsingCk = "Y"
						optDc = optDc & Server.URLEncode(rpTxt(db2Html(replace(rsget("optionname"), ":", ""))))

						itemsu = itemsu & optLimit
						addprice = addprice & rsget("optaddprice")
						addbuyprice = addbuyprice & getEzwelAddSuplyPrice(rsget("optaddprice"))
						optTax = optTax & optTaxCk
						optUsing = optUsing & optUsingCk

						rsget.MoveNext
						If Not(rsget.EOF) Then
							optDc	= optDc & "|"
							itemsu = itemsu & "|"
							addprice = addprice & "|"
							addbuyprice = addbuyprice & "|"
							optTax	= optTax & "|"
							optUsing = optUsing & "|"
						End If
					Loop
				End If
				rsget.Close
				buf = buf & "		<useYn>Y</useYn>"												'상품옵션사용여부 | 옵션이 있을경우(Y) 없을경우(N)
				buf = buf & "		<arrOptionCdNm>"&Server.URLEncode("선택")&"</arrOptionCdNm>"	'상품옵션명
				buf = buf & "		<arrOptionContent>"&optDc&"</arrOptionContent>"					'상품옵션 내용
				buf = buf & "		<arrOptionUseYn>Y</arrOptionUseYn>"								'옵션별에 따른 사용여부 | Y:N
				buf = buf & "		<arrOptionAddAmt>"&itemsu&"</arrOptionAddAmt>"					'*(옵션이 존재하는 경우만) | 상품옵션 수량 | Default: 10000
				buf = buf & "		<arrOptionAddPrice>"&addprice&"</arrOptionAddPrice>"			'상품옵션추가가격
				buf = buf & "		<arrOptionAddBuyPrice>"&addbuyprice&"</arrOptionAddBuyPrice>"	'공급가
				buf = buf & "		<arrOptionAddTaxYn>"&optTax&"</arrOptionAddTaxYn>"				'과세여부 | 과세(Y), 면세(N), 영세(숫자 0)
				buf = buf & "		<arrOptionFullUseYn>"&optUsing&"</arrOptionFullUseYn>"			'옵션 상세별에 따른 사용여부 |||    Y|Y|Y:N|N:N
			Else
				strSql = ""
				strSql = strSql &  "SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

				optDc = ""
				optLimit = ""
				If FVatInclude = "N" Then
					optTaxCk = "N"
				Else
					optTaxCk = "Y"
				End If

				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
						optLimit = rsget("optLimit")
						optLimit = optLimit-5
						If (optLimit < 1) Then optLimit = 0
						If (FLimitYN <> "Y") Then optLimit = 999   ''2013/06/12 재고관리여부 모두 Y로 변경 되므로
						optUsingCk = "N"
						optDc = optDc & Server.URLEncode(rpTxt(db2Html(replace(rsget("optionname"), ":", ""))))

						itemsu = itemsu & optLimit
						addprice = addprice & rsget("optaddprice")
						addbuyprice = addbuyprice & getEzwelAddSuplyPrice(rsget("optaddprice"))
						optTax = optTax & optTaxCk
						optUsing = optUsing & optUsingCk

						rsget.MoveNext
						If Not(rsget.EOF) Then
							optDc	= optDc & "|"
							itemsu = itemsu & "|"
							addprice = addprice & "|"
							addbuyprice = addbuyprice & "|"
							optTax	= optTax & "|"
							optUsing = optUsing & "|"
						End If
					Loop
				End If
				rsget.Close
				buf = buf & "		<useYn>Y</useYn>"												'상품옵션사용여부 | 옵션이 있을경우(Y) 없을경우(N)
				buf = buf & "		<arrOptionCdNm>"&Server.URLEncode("선택")&"</arrOptionCdNm>"	'상품옵션명
				buf = buf & "		<arrOptionContent>"&optDc&"</arrOptionContent>"					'상품옵션 내용
				buf = buf & "		<arrOptionUseYn>Y</arrOptionUseYn>"								'옵션별에 따른 사용여부 | Y:N
				buf = buf & "		<arrOptionAddAmt>"&itemsu&"</arrOptionAddAmt>"					'*(옵션이 존재하는 경우만) | 상품옵션 수량 | Default: 10000
				buf = buf & "		<arrOptionAddPrice>"&addprice&"</arrOptionAddPrice>"			'상품옵션추가가격
				buf = buf & "		<arrOptionAddBuyPrice>"&addbuyprice&"</arrOptionAddBuyPrice>"	'공급가
				buf = buf & "		<arrOptionAddTaxYn>"&optTax&"</arrOptionAddTaxYn>"				'과세여부 | 과세(Y), 면세(N), 영세(숫자 0)
				buf = buf & "		<arrOptionFullUseYn>"&optUsing&"</arrOptionFullUseYn>"			'옵션 상세별에 따른 사용여부 |||    Y|Y|Y:N|N:N
			End If
		Else
			buf = buf & "		<useYn>N</useYn>"												'상품옵션사용여부 | 옵션이 있을경우(Y) 없을경우(N)
		End If
		getEzwelOptionParam = buf
    End Function

	Public Function getEzwelNewOptionParam(obj)
		Dim strSql, strRst, i, optLimit
    	Dim buf, optDc, itemsu, addprice, addbuyprice, optTaxCk, optTax, optUsingCk, optUsing
'FoptionCnt = 0
    	buf = ""
		If FoptionCnt>0 then
			strSql = ""
			strSql = strSql &  "SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice, itemoption "
			strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

			optDc = ""
			optLimit = ""
			If FVatInclude = "N" Then
				optTaxCk = "N"
			Else
			 	optTaxCk = "Y"
			End If

			If Not(rsget.EOF or rsget.BOF) Then
				obj("useYn") = "Y"										'상품옵션사용여부
				obj("optType") = "1001"									'상품옵션유형 | 단독형(1001), 조합형(1002)
				Set obj("optionContentList") = jsArray()				'상품옵션목록
					Set obj("optionContentList")(0) = jsObject()
						obj("optionContentList")(0)("optionCdNm") = "선택"
				Set obj("optionFullContentList") = jsArray()

				i = 0
				Do until rsget.EOF
				    optLimit = rsget("optLimit")
				    optLimit = optLimit-5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = 999

					Set obj("optionFullContentList")(i) = jsObject()
						obj("optionFullContentList")(i)("optionCdNm") = "선택"						'상품옵션명
						obj("optionFullContentList")(i)("optionContent1") = db2Html(rsget("optionname"))		'옵션내용1
						obj("optionFullContentList")(i)("optionAddAmt") = optLimit 					'옵션수량
						obj("optionFullContentList")(i)("optionAddBuyPrice") = getEzwelAddSuplyPrice(rsget("optaddprice"))	'옵션매입가
						obj("optionFullContentList")(i)("optionAddPrice") =  rsget("optaddprice")	'옵션추가가격
						obj("optionFullContentList")(i)("useYn") = "Y"								'옵션상세사용여부
						obj("optionFullContentList")(i)("imgPath") = ""								'옵션썸네일이미지
						obj("optionFullContentList")(i)("imgDispYn") = "N"							'옵션이미지노출여부
						obj("optionFullContentList")(i)("sortNo") = i + 1							'옵션정렬순번
						obj("optionFullContentList")(i)("imgDtlPath") = ""							'옵션상세이미지
						obj("optionFullContentList")(i)("cspOptionFullNum") = rsget("itemoption")	'업체옵션상세코드
					rsget.MoveNext
					i = i + 1
				Loop
			End If
			rsget.Close
		Else
			obj("useYn") = "N"					'상품옵션사용여부
		End If
	End Function

	Public Function getEzwelCertParameter(obj)
		Dim strSql, safetyDiv, certNum, certOrganName, modelName, certDate, isRegCert
		Dim authType, authNum, certDiv

		strSql = ""
		strSql = strSql & " SELECT TOP 1 i.itemid, t.safetyDiv, isNull(t.certNum, '') as certNum, isNull(f.modelName, '') as modelName, isNull(f.certDate, '') as certDate, isNull(f.certOrganName, '') as certOrganName, isNull(f.certDiv, '') as certDiv "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
		strSql = strSql & " WHERE i.itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			safetyDiv	= rsget("safetyDiv")
			certNum		= rsget("certNum")
			certOrganName = rsget("certOrganName")
			modelName	= rsget("modelName")
			certDate	= rsget("certDate")
			certDiv		= rsget("certDiv")
			isRegCert	= "Y"
		Else
			isRegCert	= "N"
		End If
		rsget.Close

		If isRegCert = "Y" Then
			Select Case safetyDiv
				Case "10", "40", "70"
					authType		= "1001"
					authNum			= certNum
				Case "20", "50", "80"
					authType		= "1002"
					authNum			= certNum
				Case "30", "60", "90"
					authType		= "1003"
					authNum			= ""
			End Select

			If len(certDate) = 8 Then
				certDate = Left(certDate,4)&"-"&Mid(certDate,5,2)&"-"&Mid(certDate,7,2)
			Else
				certDate = ""
			End If

			obj("safeAuthYn") = "Y"											'인증대상 유무
			obj("authType") = authType										'인증대상 품목 | 1001:안전인증/1002:안전확인/1003:공급자적함성확인
			obj("authNum") = authNum										'인증번호
			obj("authDt") = certDate										'안전인증 일자 | ex)20220404
			obj("authDiv") = certDiv										'안전인증 항목
			obj("authOrganNm") = certOrganName								'안전인증 기관
		Else
			obj("safeAuthYn") = "N"											'인증대상 유무
			obj("authType") = ""											'인증대상 품목 | 1001:안전인증/1002:안전확인/1003:공급자적함성확인
			obj("authNum") = ""												'인증번호
			obj("authDt") = ""												'안전인증 일자 | ex)20220404
			obj("authDiv") = ""												'안전인증 항목
			obj("authOrganNm") = ""											'안전인증 기관
		End If	
	End Function
	
	Public Function getEzwelDlvrCode(iDepthCode)
		Select Case iDepthCode
			Case "45020518", "45020519", "45110106", "45110105", "45110101", "45110214", "45110212", "45110213", "45110210", "45110211", "45110207", "45110201", "45110205", "45110203", "45110202", "45110215", "70040114"	getEzwelDlvrCode = "1003"
			Case Else
				If FItemdiv = "06" OR FItemdiv = "16" Then
					getEzwelDlvrCode = "1003"
				Else
					getEzwelDlvrCode = "1001"
				End If
		End Select
	End Function

	'상품등록/수정 XML 생성
	Public Function getEzwelItemRegXML(ezwelMethod, ichkXML)
		Dim strRst
		Dim EzwelStatus
		Select Case ezwelMethod
			Case "Reg"			EzwelStatus = "1001"
			Case "SellY"		EzwelStatus = "1002"
			Case "SellN"		EzwelStatus = "1005"
			Case "MustNotOpt"	EzwelStatus = "1005"
		End Select
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "	<dataSet>"
		strRst = strRst & "		<cspCd>"&cspCd&"</cspCd>"					'##*CP 업체코드 | 이지웰 발급(고정값)
		If ezwelMethod <> "Reg" Then
		strRst = strRst & "		<goodsCd>"&FEzwelGoodno&"</goodsCd>"		'##*값이 존재하면 수정 존재하지 않으면 입력 | 상품코드 | 이지웰 상품코드
		End If
		strRst = strRst & "		<cspGoodsCd>"&FItemid&"</cspGoodsCd>"		'##업체상품코드
		strRst = strRst & "		<goodsNm><![CDATA["&Server.URLEncode(Trim(getItemNameFormat))&"]]></goodsNm>"			'##*상품명
		strRst = strRst & "		<taxYn>"&CHKIIF(FVatInclude="N","N","Y")&"</taxYn>"										'##*과세여부 | 과세(Y), 면세(N), 영세(숫자 0)
'		If EzwelStatus <> "1002" Then
			strRst = strRst & "		<goodsStatus>"&EzwelStatus&"</goodsStatus>"												'##상품상태 | 등록(1001), 판매중(1002), 판매중지(1005), 삭제(1006), 일시품절(1004) 2017-11-13 김진영..1005로 할결우 MD 승인받아야 판매중으로 변경됨
'		End If
		strRst = strRst & "		<dlvrPrice>"&CHKIIF(IsFreeBeasong=False,"3000","0")&"</dlvrPrice>"						'##배송가격
		strRst = strRst & "		<dlvrPriceApplYn>"&CHKIIF(IsFreeBeasong=True,"Y","P")&"</dlvrPriceApplYn>"				'##*착불/선결제/무료 | 무료: Y/ 소비자부담:N /착불만: A /선결제만:P
		strRst = strRst & "		<realSalePrice>"&Clng(GetEzwel10wonDown(MustPrice/10)*10)&"</realSalePrice>"			'##*판매가
		strRst = strRst & "		<normalSalePrice>"&Clng(GetRaiseValue(ForgPrice/10)*10)&"</normalSalePrice>"			'##*정상(시중)가
		strRst = strRst & "		<brandNm><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("상품설명 참조"),Server.URLEncode(rpTxt(Fmakername)))&"]]></brandNm>"	'##브랜드명
		strRst = strRst & "		<buyPrice>"&GetEzwelBuyPrice(Clng(GetEzwel10wonDown(MustPrice/10)*10))&"</buyPrice>"	'##*공급가(매입가)
		strRst = strRst & "		<modelNum>"&FItemid&"</modelNum>"														'상품모델
		strRst = strRst & "		<orginNm><![CDATA["&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),Server.URLEncode("상품설명 참조"),Server.URLEncode(Fsourcearea))&"]]></orginNm>"	'##원산지
		strRst = strRst & "		<mafcNm><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("상품설명 참조"),Server.URLEncode(rpTxt(Fmakername)))&"]]></mafcNm>"		'##제조사
		strRst = strRst & "		<enterAmt>"&getLimitEzwelEa()&"</enterAmt>"						'##*입고수량 | Default: 10000
		strRst = strRst & "		<cspDlvrId>"&cspDlvrId&"</cspDlvrId>"			'##출고지ID | 이지웰 발급(고정값)
		strRst = strRst & "		<goodsDesc><![CDATA["&Server.URLEncode(getEzwelItemContParam())&"]]></goodsDesc>"		'##상품설명
		If (ezwelMethod <> "Reg") Then		'2014-12-02 김진영 추가 | 이미지 전송 시간 오래걸림
			If isImageChanged Then
				strRst = strRst & getEzwelAddImageParam()
			End If
		Else
			strRst = strRst & getEzwelAddImageParam()
		End If
		strRst = strRst & "		<ctgCd>"&FDepthCode&"</ctgCd>"					'##*관리카테고리 | 별도첨부
		strRst = strRst & "		<dispCtgCd>"&FDepthCode&"</dispCtgCd>"			'##*전시 카테고리 | 별도첨부
		strRst = strRst & getEzwelItemInfoCd()									'##상품정보제공고시 필드정보 | 상품정보제공 고시를 위한 필드정보
		If ezwelMethod = "MustNotOpt" Then
			strRst = strRst & "	<useYn>N</useYn>"
		Else
			strRst = strRst & getEzwelOptionParam()
		End If

		strRst = strRst & "		<arrIconCd>1008</arrIconCd>"					'아이콘 | 제휴 = 1008 / 복지샵 = 1010 / 레인보우 = 1007	'2018-08-23 윤현주 1008요청
		strRst = strRst & "		<marginRate>"&CEzwelMARGIN&"</marginRate>"		'##현아대리님 10%라고 답변 | *마진률 | 9.0
		strRst = strRst & "		<dlvrForm>"&getEzwelDlvrCode(FDepthCode)&"</dlvrForm>"			'배송형태 | 1001 : 일반택배, 1002 : 자체배송, 1003 : 주문제작, 1004 : 설치제품
		strRst = strRst & "		<keyword><![CDATA["&RightCommaDel(Trim(getKeywords()))&"]]></keyword>"			'검색키워드 | 다중 키워드 입력가능 (,)로 구분 ex)긴팔,해외직구,유명브랜드
		strRst = strRst & "		<unitOrderQty>"& FOrderMaxNum &"</unitOrderQty>"	'인당구매수량 | 1회에 구매할 수 있는 수량 제어 * 값을 보내지 않거나 0인경우 제어하지 않음
		strRst = strRst & "</dataSet>"
		getEzwelItemRegXML = strRst
If (session("ssBctID")="kjy8517") Then
		response.write replace(strRst, "?xml", "?AAAAAl")
'		response.end
End If
	End Function

	'상품등록/수정 Json 생성
	Public Function getEzwelItemRegJson(v)
		Dim obj
		Set obj = jsObject()
			If v = "EDIT" Then
				obj("goodsCd") = FEzwelGoodNo
			End If

			If application("Svr_Info")="Dev" Then
				FDepthCode = "70040114"
			End If
			obj("cspGoodsCd") = FItemid										'업체상품코드
			obj("goodsNm") = getItemNameFormat()							'#상품명
			obj("taxYn") = CHKIIF(FVatInclude="N","N","Y")					'#과세여부 | 과세(Y), 면세(N), 영세(숫자 0)
			obj("goodsStatus") = "1001"										'상품상태 | 등록(1001), 판매중(1002), 판매중지(1005), 삭제(1006)
			obj("dlvrPrice") = CHKIIF(IsFreeBeasong=False,"3000","0")		'배송가격
			obj("addJejuDlvrPrice") = "3000"								'추가배송비(제주)
			obj("addSanganDlvrPrice") = "3000"								'추가배송비(도서산간)
			obj("dlvrPriceApplYn") = CHKIIF(IsFreeBeasong=True,"Y","P")		'#착불./선결제/무료 | 무료: Y / 소비자부담:N /착불만: A /선결제만:P /무료(지역별상이):C
			obj("realSalePrice") = Clng(GetEzwel10wonDown(MustPrice/10)*10)	'#판매가
			obj("normalSalePrice") = Clng(GetRaiseValue(ForgPrice/10)*10)	'#정상(시중)가
			obj("brandNm") = chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조", Fmakername)	'브랜드명
			obj("brandCd") = getBrandCode(Fmakername)						'#브랜드코드
			obj("buyPrice") = GetEzwelBuyPrice(Clng(GetEzwel10wonDown(MustPrice/10)*10))							'#공급가(매입가)
			obj("modelNum") = FItemid										'상품모델
			obj("orginNm") = chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea), "상품설명 참조", Fsourcearea)		'원산지
			obj("mafcNm") = chkIIF(trim(Fmakername)="" or isNull(Fmakername), "상품설명 참조", Fmakername)	'제조사
			obj("mafcCd") = getMafcCode(Fmakername)							'#제조사코드
			obj("enterAmt") = getLimitEzwelEa()								'#입고수량
			obj("cspDlvrId") = cspDlvrId									'출고지ID
			obj("saleStartDt") = Replace(Date(), "-", "")					'판매시작일자
			obj("saleEndDt") = "20991231"									'판매종료일자
			obj("goodsDesc") = getEzwelNewItemContParam()					'상품설명
			Call getEzwelNewAddImageParam(obj)
			obj("ctgCd") = FDepthCode										'#관리 카테고리
			obj("dispCtgCd") = FDepthCode									'#전시 카테고리
			Call getEzwelItemNewInfoCd(obj)
			Call getEzwelNewOptionParam(obj)
'			obj("iconCd") = ""												'식품아이콘 | 카테고리가 식품일 경우 1011(상온)/1012(냉동)/1013(냉장)/1014(해당없음)
			Set obj("arrIconCd") = jsArray()								'아이콘 | 제휴 = 1008 / 복지샵(레인보우) = 1007
				obj("arrIconCd")(0) = "1008"
			obj("marginRate") = CEzwelMARGIN								'#마진률
			obj("dlvrForm") = getEzwelDlvrCode(FDepthCode)					'#상품유형 | 1001:일반택배,1002:자체배송,1003:주문제작,1004:설치제품,1005:해외직배송,1006:판매종료후발주,1007:냉장/냉동식품,1008:신선식품
			obj("exchgPrice") = 3000										'교환 배송비
			obj("returnPrice") = 3000										'반품 배송비
'			obj("bndlNonChgReturnYn") = ""									'묶음교환/반품불가 | Y/N
			obj("keyword") = RightCommaDel(Trim(getNewKeywords()))			'검색키워드 | 다중 키워드입력가능 (,)로 구분 ex)긴팔, 해외직구, 유명브랜드
'			obj("shortDesc") = ""											'상품홍보문구
			obj("policyNo") = "10744781"									'#발송정책 시퀀스 | CP발송정책 시퀀스로 전송
'			obj("imgPath640") = ""											'모바일이미지경로(640*320)
			obj("dlvrFreeYn") = "Y"											'조건부 무료배송 여부 | Y:사용, N:미사용 *출고지ID(cspDlvrId)에 설정된 조건부 무료배송 여부가 N인경우 무조건 N으로 등록(출고/반품지 등록/수정 API 참고)
			obj("unitOrderQty") = FOrderMaxNum								'#1회당  구매수량 | 1회에 구매할 수 있는 수량 제어 * 값을 보내지 않거나 0인경우 제어하지 않음
			obj("idUnitOrderQty") = 0										'인당구매수량(년도) | 1년에 한 아이디 당 구매할 수 있는 수량 제어 * 값을 보내지 않거나 0인 경우 제어하지 않음
'			obj("minPriceYn") = ""											'최저가확인 | 최저가적용 사용:Y/사용안함:N/단품사용:D/미매칭:M
'			obj("minPriceUrl") = ""											'최저가 링크 | 예시)https://search.shopping.naver.com/detail/detail.nhn?nv_mid=xxxx
			obj("exceptBndlDlvrYn") = "N"									'묶음배송 제외여부 | Y:묶음배송 불가, N:묶음배송 가능
			obj("goodsType") = "1001"										'상품유형 | 1001: 일반상품, 1002: 휴대폰상품
'			obj("arrQuotaAmt") = ""											'#할부원금
'			obj("arrSaleTypeSp") = ""										'판매구분 | 1001:신규가입, 1002:번호이동, 1003:기기변경
'			obj("arrGuide") = ""											'#신규가입안내메시지
'			obj("arrSaleStopDesc") = ""										'#판매종료안내메시지
'			obj("arrJoinAmt1") = ""											'가입비 -선납 분할청구 메시지
'			obj("arrJoinAmt2") = ""											'가입비 - 무료지원 대납 메시지
'			obj("arrJoinAmt3") = ""											'가입비 - 면제, 장애인/국가유공자 메시지
'			obj("arrJoinAmt4") = ""											'가입비 - 면제, 재가입 메시지
'			obj("arrJoinYn1") = ""											'가입비 -선납 분할청구 메시지 사용여부
'			obj("arrJoinYn2") = ""											'가입비 - 무료지원 대납 메시지 사용여부
'			obj("arrJoinYn3") = ""											'가입비 - 면제, 장애인/국가유공자 메시지 사용여부
'			obj("arrJoinYn4") = ""											'가입비 - 면제, 재가입 메시지 사용여부
'			obj("arrUsimAmt1") = ""											'유심비 - 선납 메시지
'			obj("arrUsimAmt2") = ""											'유심비 - 면제 메시지
'			obj("arrUsimYn1") = ""											'유심비 - 선납 메시지 사용여부
'			obj("arrUsimYn2") = ""											'유심비 - 면제 메시지 사용여부
'			obj("arrTerm1") = ""											'조건 - 판매조건 메시지
'			obj("arrTemr2") = ""											'조건 - 부가서비스 메시지
'			obj("arrSaleStopYn") = ""										'판매종료 여부
'			obj("arrQuotaMonth") = ""										'할부개월 수 | 24: 24개월, 30:30개월
'			obj("arrMsg") = ""												'#할부개월 메시지
'			obj("arrPrepayAmt") = ""										'#선결제금액
'			obj("saleTypeUrl1") = ""										'신규가입 개통 URL
'			obj("saleTypeUrl2") = ""										'번호이동 개통 URL
'			obj("saleTypeUrl3") = ""										'기기변경 개통 URL
'			obj("noticeNm") = ""											'유의사항명
'			obj("noticeDesc") = ""											'유의사항내용
'			obj("noticeOrderNo") = ""										'유의사항정렬순서
'			obj("arrPriceCd") = ""											'요금제코드
'			obj("arrMobileUseYn") = ""										'사용여부
'			obj("arrMbDcCd") = ""											'단말기할인코드
'			obj("arrFixDcCd") = ""											'약정할인코드
'			obj("arrQuotaMonthPrice") = ""									'할부개월수 | 24: 24개월, 30:30개월
			obj("adultAuthYn") = IsAdultItem()								'성인 본인인증 필요상품대상유무 | Y일시 사용자 화면에서 19세 이상 본인인증을 거친 후, 상품구매가능
			Call getEzwelCertParameter(obj)
			getEzwelItemRegJson = obj.jsString
		Set obj = nothing
	End Function

	'상품가격변경 Json 생성
	Public Function getEzwelItemPriceJson()
		Dim obj
		Set obj = jsObject()
			obj("goodsCd") = FEzwelGoodno
			obj("realSalePrice") = Clng(GetEzwel10wonDown(MustPrice/10)*10)	'#판매가
			obj("buyPrice") = GetEzwelBuyPrice(Clng(GetEzwel10wonDown(MustPrice/10)*10))	'#공급가(매입가)
			getEzwelItemPriceJson = obj.jsString
		Set obj = nothing
	End Function

	'상품옵션변경 Json 생성
	Public Function getEzwelItemOptionJson()
		Dim obj
		Dim strSql, strRst, i, optLimit
    	Dim buf, optDc, itemsu, addprice, addbuyprice, optTaxCk, optTax, optUsingCk, optUsing
'FoptionCnt = 0
		Set obj = jsObject()
			obj("goodsCd") = FEzwelGoodno

    	buf = ""
		If FoptionCnt>0 then
			strSql = ""
			strSql = strSql &  "SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice, itemoption "
			strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

			optDc = ""
			optLimit = ""
			If FVatInclude = "N" Then
				optTaxCk = "N"
			Else
			 	optTaxCk = "Y"
			End If

			If Not(rsget.EOF or rsget.BOF) Then
				obj("useYn") = "Y"										'상품옵션사용여부
				obj("optType") = "1001"									'상품옵션유형 | 단독형(1001), 조합형(1002)
				Set obj("optionContentList") = jsArray()				'상품옵션목록
					Set obj("optionContentList")(0) = jsObject()
						obj("optionContentList")(0)("optionCdNm") = "선택"
				Set obj("optionFullContentList") = jsArray()

				i = 0
				Do until rsget.EOF
				    optLimit = rsget("optLimit")
				    optLimit = optLimit-5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = 999

					Set obj("optionFullContentList")(i) = jsObject()
						obj("optionFullContentList")(i)("optionCdNm") = "선택"						'상품옵션명
						obj("optionFullContentList")(i)("optionContent1") = db2Html(rsget("optionname"))		'옵션내용1
						obj("optionFullContentList")(i)("optionAddAmt") = optLimit 					'옵션수량
						obj("optionFullContentList")(i)("optionAddBuyPrice") = getEzwelAddSuplyPrice(rsget("optaddprice"))	'옵션매입가
						obj("optionFullContentList")(i)("optionAddPrice") =  rsget("optaddprice")	'옵션추가가격
						obj("optionFullContentList")(i)("useYn") = "Y"								'옵션상세사용여부
						obj("optionFullContentList")(i)("imgPath") = ""								'옵션썸네일이미지
						obj("optionFullContentList")(i)("imgDispYn") = "N"							'옵션이미지노출여부
						obj("optionFullContentList")(i)("sortNo") = i + 1							'옵션정렬순번
						obj("optionFullContentList")(i)("imgDtlPath") = ""							'옵션상세이미지
						obj("optionFullContentList")(i)("cspOptionFullNum") = rsget("itemoption")	'업체옵션상세코드
					rsget.MoveNext
					i = i + 1
				Loop
			End If
			rsget.Close
		Else
			obj("useYn") = "N"					'상품옵션사용여부
		End If
			getEzwelItemOptionJson = obj.jsString
		Set obj = nothing
	End Function
End Class

Class CEzwel
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

	Private Sub Class_Initialize()
		Redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function


	Public Sub getEzwelNotRegOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			'옵션 전체 품절인 경우 등록 불가.
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & " select itemid from ("
            addSql = addSql & "     select itemid"
            addSql = addSql & " 	,count(*) as optCNT"
            addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	from db_item.dbo.tbl_item_option"
            addSql = addSql & " 	where itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and isusing='Y'"
            addSql = addSql & " 	group by itemid"
            addSql = addSql & " ) T"
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"
            addSql = addSql & " and isNULL(c.infodiv,'') not in ('','18','20','21','22')"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, isNULL(R.ezwelStatCD,-9) as ezwelStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(bm.depthCode, '') as depthCode "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_Newcate_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_regItem R on i.itemid=R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.itemdiv not in ('21', '23', '30') "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "									'플라워/화물배송/해외직구 상품 제외
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' and i.itemdiv not in ('06', '16') "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
'		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (Select makerid From db_etcmall.dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (Select itemid From db_etcmall.dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and i.itemid not in (Select itemid From db_etcmall.dbo.tbl_ezwel_regItem where ezwelStatCD>3) "
		strSql = strSql & "	and uc.isExtUsing='Y'"  ''20130304 브랜드 제휴사용여부 Y만.
		strSql = strSql & addSql																				'카테고리 매칭 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CEzwelItem
				FOneItem.FItemid			= rsget("itemid")
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
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
                FOneItem.FezwelStatCD		= rsget("ezwelStatCD")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FdepthCode			= rsget("depthCode")
                FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub

	Public Sub getEzwelEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

        ''//연동 제외상품
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt < getdate()"
        addSql = addSql & "     and edDt > getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, m.ezwelGoodNo, m.ezwelprice, m.ezwelSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	,isnull(bm.depthCode, '') as depthCode "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
'		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
'		strSql = strSql & "		or i.itemdiv = '06' "
		strSql = strSql & "		or i.itemdiv in ('06', '16') "

		'홈/데코 > 조화/플라워 > 식물/플라워 카테고리면서 꽃다발, 전국택배 속하면 품절
		strSql = strSql & "		or "
		strSql = strSql & "		( "
		strSql = strSql & "			(i.cate_large = '050' and i.cate_mid = '110' and i.cate_small = '030') "
		strSql = strSql & "			AND ((i.itemname like '%꽃다발%') or (i.itemname like '%전국택배%')) "
		strSql = strSql & "		) "

		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "

		strSql = strSql & "		or i.makerid  in (Select makerid From [db_etcmall].dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_etcmall].dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_ezwel_regitem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_Newcate_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_ezwel_Newcate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.ezwelGoodNo is Not Null "									'#등록 상품만
''rw strSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CezwelItem
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
				FOneItem.ForderComment		= db2html(rsget("ordercomment"))
				FOneItem.FoptionCnt			= rsget("optionCnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FezwelGoodNo		= rsget("ezwelGoodNo")
				FOneItem.Fezwelprice		= rsget("ezwelprice")
				FOneItem.FezwelSellYn		= rsget("ezwelSellYn")

                FOneItem.FoptionCnt         = rsget("optionCnt")
                FOneItem.FregedOptCnt       = rsget("regedOptCnt")
                FOneItem.FaccFailCNT        = rsget("accFailCNT")
                FOneItem.FlastErrStr        = rsget("lastErrStr")
                FOneItem.Fdeliverytype      = rsget("deliverytype")
                FOneItem.FrequireMakeDay    = rsget("requireMakeDay")

                FOneItem.FinfoDiv       	= rsget("infoDiv")
                FOneItem.Fsafetyyn      	= rsget("safetyyn")
                FOneItem.FsafetyDiv     	= rsget("safetyDiv")
                FOneItem.FsafetyNum     	= rsget("safetyNum")
                FOneItem.FmaySoldOut    	= rsget("maySoldOut")

                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FdepthCode			= rsget("depthCode")
                FOneItem.Fregitemname		= rsget("regitemname")
                FOneItem.FregImageName		= rsget("regImageName")
                FOneItem.FbasicImageNm		= rsget("basicimage")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub
End Class

'Ezwel 상품코드 얻기
Function getEzwelGoodno(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 ezwelgoodno FROM db_etcmall.dbo.tbl_ezwel_regitem WHERE itemid = '"&iitemid&"' and ezwelStatcd <> '4' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		getEzwelGoodno = rsget("ezwelgoodno")
	End If
	rsget.Close
End Function

Function GetRaiseValue(value)
    If Fix(value) < value Then
    	GetRaiseValue = Fix(value) + 1
    Else
    	GetRaiseValue = Fix(value)
    End If
End Function

Function GetEzwel10wonDown(value)
   	GetEzwel10wonDown = Fix(value/10)*10
End Function

Function rpTxt(checkvalue)
	Dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, "&", "&amp;")
    v = Replace(v, """", "&quot;")
    v = Replace(v, "'", "&apos;")
    v = replace(v, "<", "&lt;")
    v = replace(v, ">", "&gt;")
	v = replace(v, "", "&gt;")
	'v = replace(v, ":", "")			'http:// 의 :가 치환되므로 패스
    rpTxt = v
End Function

Function rpContent(checkvalue)
	Dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, "<script>", "")
    v = replace(v, "</script>", "")
    v = Replace(v, "<embed>", "")
    v = Replace(v, "</embed>", "")
    v = Replace(v, "<body>", "")
    v = Replace(v, "</body>", "")
    v = replace(v, "<iframe>", "")
    v = replace(v, "</iframe>", "")
    v = replace(v, "<meta>", "")
    v = replace(v, "</meta>", "")
	v = replace(v, "<object>", "")
	v = replace(v, "</object>", "")
	v = replace(v, "<style>", "")
	v = replace(v, "</style>", "")
	v = replace(v, "<link>", "")
	v = replace(v, "</link>", "")
	v = replace(v, "<base>", "")
	v = replace(v, "</base>", "")
	v = replace(v, "<applet>", "")
	v = replace(v, "</applet>", "")
    rpContent = v
End Function

Function GetEzwelBuyPrice(value)
   	GetEzwelBuyPrice = Clng(value - (value / CEzwelMARGIN))
End Function

'// 상품이미지 존재여부 검사
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function

Function getAccessToken()
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 isnull(accessToken, '') as accessToken, lastupdate "&VbCRLF
	strSql = strSql & " FROM db_etcmall.dbo.tbl_outmall_ini"&VbCRLF
	strSql = strSql & " WHERE mallid='"& CMALLNAME &"'"&VbCRLF
	strSql = strSql & " and inikey = 'auth'"
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.Eof Then
		getAccessToken	= rsget("accessToken")
	End If
	rsget.close
End Function
%>
