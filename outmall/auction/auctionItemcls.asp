<%
CONST CMAXMARGIN = 18
CONST CMALLNAME = "auction1010"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST auctionAPIURL = "https://api.auction.co.kr"
CONST auctionTicket = "d3XubWMyHSXucjs2uJ0Fz5C+xyg9FcHga9EzBIM0tnWQbtoXF80ywv34kCmUo0SWnQpl8+H+T3b5IV8/TT/OLSsYCP+TKLkPrVW7EBCTz6xkSTmYMZ/Lqnvif78jMZCBgoDYVsOQwSiPM1IJXZ6zJfe0j1DOu4fWlwKNSeqmcswq5BLj0NaQJmHqPZLx6feNdAZ3NYzh3PfEGa1XGGkXEt4="
CONST CDEFALUT_STOCK = 100

Class CAuctionItem
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
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public FSafetyNum
	Public Fitemcontent
	Public FAuctionStatCD
	Public Fdeliverfixday
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FmaySoldOut
	Public Fregitemname
	Public FregImageName
	Public FbasicImageNm
	Public Fsocname_kor
	Public FDepthCode
	Public Fcdmkey
	Public Fcddkey
	Public FAuctionGoodNo
	Public FAuctionprice
	Public FAuctionSellYn
	Public FAPIadditem
	Public FAPIaddopt

	Public FNotinCate
	Public FSafeAuthType
	Public FAuthItemTypeCode
	Public FIsChildrenCate
	Public FIsLifeCate
	Public FIsElecCate
	Public FOverlap
	Public FRawMaterialsType
	Public FIsbn13
	Public FSellSTDate
	Public FAdultType
	Public FOrderMaxNum
	Public FOutmallstandardMargin

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "999" Then
			getOrderMaxNum = 999
		End If
	End Function

	'// 품절여부
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	end function

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

	Public Function MustPrice()
		Dim GetTenTenMargin, sqlStr, specialPrice, ownItemCnt, outmallstandardMargin
		sqlStr = ""
		sqlStr = sqlStr & " SELECT mustPrice, isnull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] as m "
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE m.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " and m.itemid = '"& Fitemid &"' "
		sqlStr = sqlStr & " and getdate() >= startDate and getdate() <= endDate "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			specialPrice = rsget("mustPrice")
			outmallstandardMargin	= rsget("outmallstandardMargin")
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
			If outmallstandardMargin = "" Then
				outmallstandardMargin	= FOutmallstandardMargin
			End If

			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If GetTenTenMargin < outmallstandardMargin Then
				MustPrice = Forgprice
			Else
				MustPrice = FSellCash
			End If
		End If
	End Function

	'최대 구매 수량
	Public Function getLimitAuctionEa()
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
		getLimitAuctionEa = ret
	End Function

	'// 옥션 판매여부 반환
	Public Function getAuctionSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getAuctionSellYn = "Y"
			Else
				getAuctionSellYn = "N"
			End If
		Else
			getAuctionSellYn = "N"
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

    public function getItemNameFormat()
        dim buf
		If application("Svr_Info") = "Dev" Then
			FItemName = "[TEST상품] "&FItemName
		End If
        buf = replace(FItemName,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","프로")
        buf = replace(buf,"&","＆")
        buf = replace(buf,"[무료배송]","")
        buf = replace(buf,"[무료 배송]","")
        buf = LeftB(buf, 52)
        getItemNameFormat = buf
    end function

    public function getItemNameFormat2()
        dim buf
		If application("Svr_Info") = "Dev" Then
			FItemName = "[TEST상품] "&FItemName
		End If
        buf = replace(FItemName,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","프로")
        buf = replace(buf,"&","＆")
        buf = replace(buf,"[무료배송]","")
        buf = replace(buf,"[무료 배송]","")
        getItemNameFormat2 = buf
    end function

	Public Function checkItemContent()
		Dim strSql, chkRst, etcLinkStr, isVal
		isVal = "N"
		strSql = ""
		strSql = strSql & " SELECT itemid, mallid, linkgbn, textVal, 'Y' as isVal " & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSql = strSql & " where mallid in ('','auction1010') and linkgbn = 'contents' and itemid = '"&FItemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			etcLinkStr	= rsget("textVal")
			isVal		= rsget("isVal")
		End If
		rsget.Close

		If Instr(LCase(etcLinkStr), "<iframe") > 0 Then
			checkItemContent = "Y"
		ElseIf isVal <> "Y" AND Instr(LCase(FItemcontent), "<iframe") > 0 Then
			checkItemContent = "Y"
		Else
			checkItemContent = "N"
		End If
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

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	Public Function getAuctionItemContParamToReg()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
		'2014-01-17 10:00 김진영 탑 이미지 추가
		strRst = strRst & ("<p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_auction.jpg></p>&#xA;")
		strRst = strRst & ("<div style=""width:100%; max-width:700px; margin:0; padding:0; margin-bottom:14px; padding-bottom:6px; background:url(http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_namebg.png) left bottom no-repeat;"">")
		strRst = strRst & ("<table cellpadding=""0"" cellspacing=""0"" width=""100%"">")
		strRst = strRst & ("<tr>")
		strRst = strRst & ("<th style=""vertical-align:middle; width:73px; height:42px; text-align:center; margin:0; padding:3px 0 0 0;""><img src=""http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_nametit.png"" alt=""상품명"" style=""vertical-align:top; display:inline;""/></th>")
		strRst = strRst & ("<td style=""width:627px; vertical-align:middle; text-align:left; font-size:14px; line-height:1.2; color:#000; font-weight:bold; font-family:dotum, dotumche, '돋움', sans-serif; margin:0; padding:4px 0 0 0;"">")
		strRst = strRst & ("<p style=""letter-spacing:-0.03em; margin:0; padding:12px 10px;"">")
		strRst = strRst & getItemNameFormat2
		strRst = strRst & ("</p>")
		strRst = strRst & ("</td>")
		strRst = strRst & ("</tr>")
		strRst = strRst & ("</table>")
		strRst = strRst & ("</div>")

		If ForderComment <> "" Then
			strRst = strRst & "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
		End If

		'#기본 상품설명
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "&#xA;")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "&#xA;")
			Case Else
				strRst = strRst & (nl2br(Fitemcontent) & "&#xA;")
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
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%"">&#xA;")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		'#기본 상품 설명이미지
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%"">&#xA;")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%"">&#xA;")

		'#배송 주의사항
		strRst = strRst & ("&#xA;<img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_auction.jpg>")

		strRst = strRst & ("</div>")
		getAuctionItemContParamToReg = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		if Not(rsget.EOF or rsget.BOF) then
			strRst = rsget("textVal")
			strRst = "<div align=""center""><p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_auction.jpg></p>&#xA;" & strRst & "&#xA;<img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_auction.jpg></div>"
			getAuctionItemContParamToReg = strRst
		End If
		rsget.Close
	End Function

	'// 검색어
	Public Function getItemKeyword()
		Dim arrRst, arrRst2, q, Keyword1, strRst
		If trim(Fkeywords) = "" Then Exit Function
		Fkeywords  = replace(Fkeywords,"%", "")
		Fkeywords  = replace(Fkeywords,"/", ",")
		Fkeywords  = replace(Fkeywords,".", "")
		Fkeywords  = replace(Fkeywords,"+", "")
		Fkeywords  = replace(Fkeywords,"_", "")
		Fkeywords  = replace(Fkeywords,"(", "")
		Fkeywords  = replace(Fkeywords,")", "")
		Fkeywords  = replace(Fkeywords,"&", "")
		Fkeywords  = replace(Fkeywords,";", "")
		Fkeywords  = replace(Fkeywords,"#", "")
		Fkeywords  = replace(Fkeywords,"'", "")
		Fkeywords  = replace(Fkeywords,"[", "")
		Fkeywords  = replace(Fkeywords,"]", "")
		Fkeywords  = replace(Fkeywords,":", "")
		Fkeywords  = replace(Fkeywords,"\", "")

		Fkeywords  = replace(Fkeywords,chr(13), "")
		Fkeywords  = replace(Fkeywords,chr(10), "")
		Fkeywords  = replace(Fkeywords,chr(9), "")
		Fkeywords  = replace(Fkeywords,chr(32), "")

		arrRst = Split(Fkeywords,",")
		If Ubound(arrRst) = 0 then
			arrRst2 = split(arrRst(0),";")
			If Ubound(arrRst2) > 0 then
				arrRst = split(Fkeywords,";")
			End If
		End If

		If Ubound(arrRst)+1 >= 5 then
			getItemKeyword = LeftB(arrRst(0), 20) &","&LeftB(arrRst(1), 20) &","& LeftB(arrRst(2), 20) &","& LeftB(arrRst(3), 20) &","& LeftB(arrRst(4), 20)
		Else
			For q = 0 to Ubound(arrRst)
				Keyword1 = Keyword1&LeftB(arrRst(q), 20) &","
			Next
			If Right(keyword1,1) = "," Then
				keyword1 = Left(keyword1,Len(keyword1)-1)
			End If
			getItemKeyword = keyword1
		End If
'rw getItemKeyword
'response.end
	End Function

	Public Function getAuctionDate()
		Dim strSQL, strRst, vmadeDate, vuseDate, isVal
		strRst = ""
		strSQL = ""
		strSQL = strSQL & " SELECT TOP 1 madeDate, useDate " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid = 'auction1010' and linkgbn = 'auctionDate' and itemid = '"&Fitemid&"' and valtype = '4' " & VBCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			isVal = "o"
			vmadeDate	= rsget("madeDate")
			vuseDate	= rsget("useDate")
		Else
			isVal = "x"
		End If
		rsget.Close

		If isVal = "o" Then
			If vmadeDate <> "" Then
				strRst = strRst & " ProductionDate="""&vmadeDate&""""
			End If

			If vuseDate <> "" Then
				strRst = strRst & " Expiry="""&vuseDate&""""
			End If
			getAuctionDate = strRst
		Else
			getAuctionDate = ""
		End If
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

	Public Function getAuctionAddImageParam()
		Dim strRst, strSQL, i, addImgUrl
		Dim basicImageStr, addImageStr
		basicImageStr = FbasicImage & "/10x10/resize/600/"
'2023-07-19 김진영..하단 주석처리 전부 그냥 포토서버 url로 변경
'		If Instr(FbasicImage, "/image/basic600") > 0 Then
'			basicImageStr = FbasicImage
'		Else
'			basicImageStr = FbasicImage & "/10x10/resize/600/"
'		End If

		strRst = ""
		strRst = strRst & "					<ItemPicture xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"">"'고정, 기본, 목록사진 모두 저희쪽 기준으로 기본이미지 (썸네일 이미지) 넣으면 될 것 같아요
		strRst = strRst & "						<FixImage Uri="""& basicImageStr &""" Description=""FixImage"" />"				'고정 이미지
		strRst = strRst & "						<Picture1 Uri="""& basicImageStr &""" Description=""Picture1"" />"				'기본 사진 권장(300x300)
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If (IsNULL(rsget("addimage_600")) or (rsget("addimage_600")="")) Then
					addImgUrl = "add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")
				Else
					addImgUrl = "add" & rsget("gubun") & "_600/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_600")
				End If

				If rsget("imgType") = "0" Then
'2023-07-19 김진영..하단 주석처리 전부 그냥 포토서버 url로 변경
'					If Instr(addImgUrl, "add" & rsget("gubun") & "_600/") > 0 Then
'						strRst = strRst & "						<Picture"&i+1&" Uri="""&"http://webimage.10x10.co.kr/image/"&addImgUrl&""" Description=""Picture"&i+1&""" />"	'추가 사진1 권장(300x300)
'					Else
'						strRst = strRst & "						<Picture"&i+1&" Uri="""&"http://webimage.10x10.co.kr/image/"&addImgUrl&"/10x10/resize/600/"&""" Description=""Picture"&i+1&""" />"	'추가 사진1 권장(300x300)
'					End If
					strRst = strRst & "						<Picture"&i+1&" Uri="""&"http://webimage.10x10.co.kr/image/"&addImgUrl&"/10x10/resize/600/"&""" Description=""Picture"&i+1&""" />"	'추가 사진1 권장(300x300)
				End If
				rsget.MoveNext
				If i>=2 Then Exit For		'3이상은 프리미엄(유료)시 노출됨 / 우리는 안 쓸거니 등록할 이유없음. 이미지등록시 api속도 저하가 되므로..
			Next
		End If
		rsget.Close

		strRst = strRst & "					</ItemPicture>"
		getAuctionAddImageParam = strRst
	End Function

	'기본정보 등록 soap XML
	Public Function getAuctionItemRegParameter
		Dim strRst, tt, isMadeInKorea, ImportedCode, ImportedAgency
		If Fsourcearea = "한국" OR Fsourcearea = "대한민국" Then
			isMadeInKorea = "Domestic"		'국내
		Else
			isMadeInKorea = "Imported"		'수입
			ImportedCode = getNationName2Code(Fsourcearea,ImportedAgency)
			'CoastalWaters  연근해
			'Domestic  국내
			'Imported  수입
			'Ocean  원양산
			'Unknown  모름
		End If

 		If len(FDepthCode) = 7 Then FDepthCode = "0"&CStr(FDepthCode)

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "	<soap:Header>"
		strRst = strRst & "		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
		strRst = strRst & "			<Value>"&auctionTicket&"</Value>"
		strRst = strRst & "		</EncryptedTicket>"
		strRst = strRst & "	</soap:Header>"
		strRst = strRst & "	<soap:Body>"
		strRst = strRst & "		<AddItem xmlns=""http://www.auction.co.kr/APIv1/ShoppingService"">"
		strRst = strRst & "			 <req Version=""1"">"
		strRst = strRst & "			 	<Item "
		strRst = strRst & "					BrandCode=""3104"""										'브랜드 코드(텐바이텐 : 3104)
		strRst = strRst & "				 	BuyableQuantity="""&getOrderMaxNum&"""" 				'최대 구매허용 수량
		strRst = strRst & "				 	BuyLimitTypeCode=""OnceLimited"""						'최대 구매 허용 수량코드 | OnceLimited	1회제한, OneManLimited	1인제한, PeriodLimited	기간제한, Unlimited	제한없음
		strRst = strRst & "				 	CategoryCode="""&FDepthCode&"""" 						'카테고리 코드
		strRst = strRst & "				 	DescriptionVerType=""New"""								'물품 상세 정보 (HTML 직접 입력 형태)
		strRst = strRst & "				 	Name="""&getItemNameFormat&""""							'물품명
		strRst = strRst & "				 	ItemStatusType=""New"""									'물품 상태 구분 New  신상품
		strRst = strRst & "				 	PlaceOfOrigin="""&isMadeInKorea&""""					'원산지
		If FIsbn13 <> "" Then
		strRst = strRst & "				 	ISBN="""&FIsbn13&""""									'ISBN 코드
		strRst = strRst & "				 	ProductionDate="""&FSellSTDate&""""						'제조일자/발행일자
		End If
		If FRawMaterialsType = "Y" Then		'카테고리에서 식품유형과 매칭된 카테고리라면..
		strRst = strRst & "				 	RawMaterialsType =""Inside"""							'식품유형
		End If
		strRst = strRst & "					IsAdult="""&Chkiif(IsAdultItem() = "Y", "true", "false")&""""	'성인물품 여부
		strRst = strRst & "					IsPCS=""true"""											'가격비교 사이트 등록여부
		strRst = strRst & getAuctionDate()
		strRst = strRst & "				 	Price="""&Clng(GetRaiseValue(MustPrice/10)*10)&""""		'판매가
		strRst = strRst & "				 	SellingArea=""Nationwide"""								'판매 지역
		strRst = strRst & "				 	WishKeyword="""&getItemKeyword&""""						'희망검색어
		strRst = strRst & "				 	ItemCode="""&FItemid&""" xmlns=""http://schema.auction.co.kr/Arche.Sell3.Service.xsd"">"	'판매자 관리코드 | 판매자가 상품에 임의의 코드를 지정하는 데 사용
		strRst = strRst & getAuctionAddImageParam()
		strRst = strRst & "					<ShippingFee xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"""
		strRst = strRst & "						ShippingType=""Door2Door"""							'배송방법 | 옥션코드 : Door2Door(택배)
		strRst = strRst & "						ShippingFeeChargeType=""Amount"""
		strRst = strRst & "						IsPrepayable=""false"""								'선결제 가능 여부
		strRst = strRst & "						IsArrival=""false"""
		strRst = strRst & "						IsDefault=""false"">"
		strRst = strRst & "						<ShipingFeeType>SellerShipping</ShipingFeeType>"	'배송비 부담 방식 | 옥션코드 : SellerShipping(판매자 조건부)
		strRst = strRst & "						<ShippingPlaceSeq>1557709</ShippingPlaceSeq>"		'배송출하지 SEQ번호 | 1557709
'		strRst = strRst & "						<ShippingPolicyNo>3555055</ShippingPolicyNo>"		'판매자 묶음배송비 선택시 배송출하지에 묶여있는 묶음배송정책 번호 | '134383728 5만원 3천원 코드로 바뀌어야 함		strRst = strRst & "					</ShippingFee>"
		strRst = strRst & "						<ShippingPolicyNo>134383728</ShippingPolicyNo>"		'판매자 묶음배송비 선택시 배송출하지에 묶여있는 묶음배송정책 번호 | '134383728 5만원 3천원 코드로 바뀌어야 함		strRst = strRst & "					</ShippingFee>"
		strRst = strRst & "					</ShippingFee>"
		strRst = strRst & "					<ItemReturn DeliveryAgency=""cjgls"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"">"
		strRst = strRst & "						<Address SellerAddrNo=""102197944"" />"				'판매자주소록 No
'		strRst = strRst & "						<ExtraInfo ReturnFee=""2500"" />"
		strRst = strRst & "						<ExtraInfo ReturnFee=""3000"" />"
		strRst = strRst & "					</ItemReturn>"
		strRst = strRst & "					<ItemContentsHtml ItemHtml="""&replaceRst(getAuctionItemContParamToReg)&""" ItemPromotionHtml="""" ItemAddHtml="""" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />" '상품상세 HTML
		If (isMadeInKorea = "Imported") OR (FIsChildrenCate = "Y") OR (FIsLifeCate = "Y") OR (FIsElecCate = "Y")  Then	'국산이 아닐 때 OR 상품인증번호 필요한 카테라면 꼭 아래것을 호출해야 한다함
		strRst = strRst & "					<ItemExtra xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"">"
			If isMadeInKorea = "Imported" Then
		strRst = strRst & "						<ImportedItem ImportAgency="""&ImportedAgency&""" IsMultiple=""false"" Nation="""&ImportedCode&""" />"
			End If
		strRst = strRst & getAuctionCertInfo()
		strRst = strRst & "					</ItemExtra>"
		End If
		strRst = strRst & "				</Item>"
		strRst = strRst & "			</req>"
		strRst = strRst & "		</AddItem>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
		getAuctionItemRegParameter = strRst
'response.write strRst
'response.end
	End Function

	'옵션등록 Soap XML
	Public Function getAuctionOPTRegParameter()
		Dim strSQL, strRst, strRst1, strRst2, strRst3
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "	<soap:Header>"
		strRst = strRst & "		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
		strRst = strRst & "			<Value>"&auctionTicket&"</Value>"
		strRst = strRst & "		</EncryptedTicket>"
		strRst = strRst & "	</soap:Header>"
		strRst = strRst & "	<soap:Body>"
		strRst = strRst & "		<ReviseItemStock xmlns=""http://www.auction.co.kr/APIv1/ShoppingService"">"
		strRst = strRst & "			<req Version=""1"">"
		strRst = strRst & getAuctionOptParamtoReg()
		strRst = strRst & "			</req>"
		strRst = strRst & "		</ReviseItemStock>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
		getAuctionOPTRegParameter = strRst
'response.write getAuctionOPTRegParameter
'response.end
	End Function

	Public Function getAuctionOptParamtoReg()
		Dim strRst, strSql, chkMultiOpt, optIsusing, optSellYn, optaddprice, MultiTypeCnt, arrMultiTypeNm, type1, type2, type3, optDc1, optDc2, optDc3
		Dim optNm, optDc, optLimit, itemoption, AuctionoptionSoldout, IsDisplayable, AuctionMultiType, MultiYN
		chkMultiOpt = false
		MultiTypeCnt = 0

		If FOptionCnt = 0 Then			'단품
			If FItemdiv = "06" Then		'단품이면서 주문제작문구 있는 상품
				strRst = ""
				strRst = strRst & "				<ItemStock ItemID="""&FAuctionGoodNo&""" Type=""BuyerDescriptive"" OptionStockType=""NotAvailable"" IsStockQtyMng=""false"" UseOptionBuyQty=""false"" xmlns=""http://schema.auction.co.kr/Arche.Sell3.Service.xsd"">"
				strRst = strRst & "					<OrderStock Quantity="""&getLimitAuctionEa&""" Price=""0"" IsDisplayable=""true"" ChangeType=""Add"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
				strRst = strRst & "					<StockText DescriptiveText=""텍스트를 입력하세요"" IsDisplayable=""true"" ChangeType=""Add"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
				strRst = strRst & "				</ItemStock>"
			Else						'단품이면서 주문제작문구가 없는 상품
				strRst = ""
				strRst = strRst & "				<ItemStock ItemID="""&FAuctionGoodNo&""" Type=""NotAvailable"" OptionStockType=""NotAvailable"" OptVerType=""New"" ImageMatchingFinishYN=""false"" OptRepImageLevel=""0"" OptDetailImageLevel=""0"" IsStockQtyMng=""false"" UseOptionBuyQty=""false"" xmlns=""http://schema.auction.co.kr/Arche.Sell3.Service.xsd"">"
				strRst = strRst & "					<Seller MemberID=""10x10store"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
				strRst = strRst & "					<OrderStock Section=""_"" Text=""_"" Quantity="""&getLimitAuctionEa&""" Price=""0"" IsDisplayable=""true"" StockMasterSeqNo=""0"" SkuMatchingVerNo=""0"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
				strRst = strRst & "				</ItemStock>"
			End If
		Else							'옵션있는 상품
			strSql = ""
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				MultiTypeCnt = rsget.recordcount
				Do until rsget.EOF
					arrMultiTypeNm = arrMultiTypeNm & replaceRst(db2Html(rsget("optionTypeName")))&","
					rsget.MoveNext
				Loop
			End If
			rsget.Close

			If FItemdiv = "06" Then		'옵션이 있으면서 주문제작문구가 있는 상품
				If chkMultiOpt = false Then		'일반 옵션 일 경우
					strRst = ""
					strRst = ""
					strRst = strRst & "				<ItemStock ItemID="""&FAuctionGoodNo&""" Type=""StandAloneMixed"" OptionStockType=""NotAvailable"" OptVerType=""New"" IsStockQtyMng=""true"" UseOptionBuyQty=""false"" xmlns=""http://schema.auction.co.kr/Arche.Sell3.Service.xsd"">"
					strSql = "Select itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
					strSql = strSql & " From [db_item].[dbo].tbl_item_option "
					strSql = strSql & " where isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If Not(rsget.EOF or rsget.BOF) then
						If db2Html(rsget("optionTypeName")) <> "" Then
							optNm = db2Html(rsget("optionTypeName"))
						Else
							optNm = "옵션"
						End If
						Do until rsget.EOF
							optLimit = rsget("optLimit")
							optLimit = optLimit-5
							If (optLimit < 1) Then optLimit = 0
							If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
							itemoption	= rsget("itemoption")
							optDc		= db2Html(rsget("optionname"))
							optDc		= replaceRst(optDc)
							optIsusing	= rsget("isusing")
							optSellYn	= rsget("optsellyn")
							optaddprice	= rsget("optaddprice")

							If (optIsusing <> "Y") OR (optSellYn <> "Y") OR (optLimit = 0) Then
								AuctionoptionSoldout	= "true"
								IsDisplayable			= "false"
							Else
								AuctionoptionSoldout	= "false"
								IsDisplayable			= "true"
							End If
							strRst = strRst & "					<StockStandAlone xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" Section="""&optNm&""" Text="""&optDc&""" SellerStockCode="""&itemoption&""" StockQty="""&optLimit&""" IsSoldOut="""&AuctionoptionSoldout&""" UseYN=""true"" Price="""&optaddprice&""" ChangeType=""Add"" />"
							rsget.MoveNext
						Loop
					end if
					rsget.Close
					strRst = strRst & "					<OrderStock Quantity="""&getLimitAuctionEa&""" Price=""0"" IsDisplayable="""&IsDisplayable&""" ChangeType=""Add"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
					strRst = strRst & "					<StockText DescriptiveText=""텍스트를 입력하세요"" IsDisplayable=""true"" ChangeType=""Add"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
					strRst = strRst & "				</ItemStock>"
				Else							'복합 옵션 일 경우
					If Right(arrMultiTypeNm,1) = "," Then
						arrMultiTypeNm = Left(arrMultiTypeNm, Len(arrMultiTypeNm) - 1)
					End If

					If MultiTypeCnt = 2 Then	'2중 옵션일 경우
						AuctionMultiType = "Mixed"
						type1 				= Split(arrMultiTypeNm, ",")(0)
						type2 				= Split(arrMultiTypeNm, ",")(1)
					Else						'3중 옵션일 경우
						AuctionMultiType	= "ThreeCombinationMixed"
						type1 				= Split(arrMultiTypeNm, ",")(0)
						type2 				= Split(arrMultiTypeNm, ",")(1)
						type3 				= Split(arrMultiTypeNm, ",")(2)
					End If

					strRst = ""
					strRst = strRst & "				<ItemStock ItemID="""&FAuctionGoodNo&""" Type="""&AuctionMultiType&""" OptionStockType=""NotAvailable"" OptVerType=""New"" IsStockQtyMng=""true"" UseOptionBuyQty=""false"" xmlns=""http://schema.auction.co.kr/Arche.Sell3.Service.xsd"">"
					strRst = strRst & "					<Seller MemberID=""10x10store"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
					If AuctionMultiType = "BuyerSelective" Then
					strRst = strRst & "					<OptionObjectName ClaseName1="""&type1&""" ObjOptNo1=""0"" ClaseName2="""&type2&""" ObjOptNo2=""0"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
					Else
					strRst = strRst & "					<OptionObjectName ClaseName1="""&type1&""" ObjOptNo1=""0"" ClaseName2="""&type2&""" ObjOptNo2=""0"" ClaseName3="""&type3&""" ObjOptNo3=""0"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
					End If

					strSql = ""
					strSql = strSql & "Select itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
					strSql = strSql & ",(case when CHARINDEX(',',optionname)=0 then 'N' else 'Y' end) as MultiYN "	'상품코드 : 1116421 옵션이 일반,복합 섞임; 2015-09-11 진영//발견 후 추가
					strSql = strSql & " From [db_item].[dbo].tbl_item_option "
					strSql = strSql & " where isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If Not(rsget.EOF or rsget.BOF) then
						Do until rsget.EOF
							optLimit = rsget("optLimit")
							optLimit = optLimit-5
							If (optLimit < 1) Then optLimit = 0
							If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
							itemoption	= rsget("itemoption")
							optDc		= db2Html(rsget("optionname"))
							optDc		= replaceRst(optDc)
							optIsusing	= rsget("isusing")
							optSellYn	= rsget("optsellyn")
							optaddprice	= rsget("optaddprice")
							MultiYN		= rsget("MultiYN")

							If (optIsusing <> "Y") OR (optSellYn <> "Y") OR (optLimit = 0) Then
								AuctionoptionSoldout	= "true"
								IsDisplayable			= "false"
							Else
								AuctionoptionSoldout	= "false"
								IsDisplayable			= "true"
							End If

							If MultiTypeCnt = 2 Then
								If MultiYN = "Y" Then
									optDc1 = split(optDc,",")(0)
									optDc2 = split(optDc,",")(1)
									strRst = strRst & "					<OrderStock Code="""&itemoption&""" Section="""&optDc1&""" Text="""&optDc2&""" Quantity="""&optLimit&""" Price="""&optaddprice&""" IsDisplayable="""&IsDisplayable&""" ChangeType=""Add""  xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
								End If
							Else
								optDc1 = split(optDc,",")(0)
								optDc2 = split(optDc,",")(1)
								optDc3 = split(optDc,",")(2)
								strRst = strRst & "					<OrderStock Code="""&itemoption&""" Section="""&optDc1&""" Text="""&optDc2&""" Text2="""&optDc3&""" Quantity="""&optLimit&""" Price="""&optaddprice&""" IsDisplayable="""&IsDisplayable&""" ChangeType=""Add""  xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
							End If
							rsget.MoveNext
						Loop
					end if
					rsget.Close
					strRst = strRst & "					<StockText DescriptiveText=""텍스트를 입력하세요"" IsDisplayable=""true"" ChangeType=""Add"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
					strRst = strRst & "				</ItemStock>"
				End If
			Else						'옵션이 있으면서 주문제작문구가 없는 상품
				If chkMultiOpt = false Then		'일반 옵션 일 경우
					strRst = ""
					strRst = strRst & "				<ItemStock ItemID="""&FAuctionGoodNo&""" Type=""StandAlone"" OptionStockType=""NotAvailable"" OptVerType=""New"" IsStockQtyMng=""true"" UseOptionBuyQty=""false"" xmlns=""http://schema.auction.co.kr/Arche.Sell3.Service.xsd"">"

					strSql = "Select itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
					strSql = strSql & " From [db_item].[dbo].tbl_item_option "
					strSql = strSql & " where isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If Not(rsget.EOF or rsget.BOF) then
						If db2Html(rsget("optionTypeName")) <> "" Then
							optNm = db2Html(rsget("optionTypeName"))
							optNm = replaceRst(optNm)
						Else
							optNm = "옵션"
						End If
						Do until rsget.EOF
							optLimit = rsget("optLimit")
							optLimit = optLimit-5
							If (optLimit < 1) Then optLimit = 0
							If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
							itemoption	= rsget("itemoption")
							optDc		= db2Html(rsget("optionname"))
							optDc		= replaceRst(optDc)

							optIsusing	= rsget("isusing")
							optSellYn	= rsget("optsellyn")
							optaddprice	= rsget("optaddprice")

							If (optIsusing <> "Y") OR (optSellYn <> "Y") OR (optLimit = 0) Then
								AuctionoptionSoldout	= "true"
							Else
								AuctionoptionSoldout	= "false"
							End If

							strRst = strRst & "					<StockStandAlone xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" Section="""&optNm&""" Text="""&optDc&""" SellerStockCode="""&itemoption&""" StockQty="""&optLimit&""" IsSoldOut="""&AuctionoptionSoldout&""" UseYN=""true"" Price="""&optaddprice&""" ChangeType=""Add"" />"
							rsget.MoveNext
						Loop
					end if
					rsget.Close
					strRst = strRst & "				</ItemStock>"
				Else							'복합 옵션 일 경우
					If Right(arrMultiTypeNm,1) = "," Then
						arrMultiTypeNm = Left(arrMultiTypeNm, Len(arrMultiTypeNm) - 1)
					End If

					If MultiTypeCnt = 2 Then	'2중 옵션일 경우
						AuctionMultiType = "BuyerSelective"
						type1 				= Split(arrMultiTypeNm, ",")(0)
						type2 				= Split(arrMultiTypeNm, ",")(1)
					Else						'3중 옵션일 경우
						AuctionMultiType	= "ThreeCombination"
						type1 				= Split(arrMultiTypeNm, ",")(0)
						type2 				= Split(arrMultiTypeNm, ",")(1)
						type3 				= Split(arrMultiTypeNm, ",")(2)
					End If

					strRst = ""
					strRst = strRst & "				<ItemStock ItemID="""&FAuctionGoodNo&""" Type="""&AuctionMultiType&""" OptionStockType=""NotAvailable"" OptVerType=""New"" IsStockQtyMng=""true"" UseOptionBuyQty=""false"" xmlns=""http://schema.auction.co.kr/Arche.Sell3.Service.xsd"">"
					strRst = strRst & "					<Seller MemberID=""10x10store"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
					If AuctionMultiType = "BuyerSelective" Then
					strRst = strRst & "					<OptionObjectName ClaseName1="""&type1&""" ObjOptNo1=""0"" ClaseName2="""&type2&""" ObjOptNo2=""0"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
					Else
					strRst = strRst & "					<OptionObjectName ClaseName1="""&type1&""" ObjOptNo1=""0"" ClaseName2="""&type2&""" ObjOptNo2=""0"" ClaseName3="""&type3&""" ObjOptNo3=""0"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
					End If

					strSql = ""
					strSql = strSql & "Select itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
					strSql = strSql & ",(case when CHARINDEX(',',optionname)=0 then 'N' else 'Y' end) as MultiYN "
					strSql = strSql & " From [db_item].[dbo].tbl_item_option "
					strSql = strSql & " where isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
					rsget.CursorLocation = adUseClient
					rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
					If Not(rsget.EOF or rsget.BOF) then
						Do until rsget.EOF
							optLimit = rsget("optLimit")
							optLimit = optLimit-5
							If (optLimit < 1) Then optLimit = 0
							If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
							itemoption	= rsget("itemoption")
							optDc		= db2Html(rsget("optionname"))
							optDc		= replaceRst(optDc)
							optIsusing	= rsget("isusing")
							optSellYn	= rsget("optsellyn")
							optaddprice	= rsget("optaddprice")
							MultiYN		= rsget("MultiYN")

							If (optIsusing <> "Y") OR (optSellYn <> "Y") OR (optLimit = 0) Then
								AuctionoptionSoldout	= "true"
								IsDisplayable			= "false"
							Else
								AuctionoptionSoldout	= "false"
								IsDisplayable			= "true"
							End If

							If MultiTypeCnt = 2 Then
								If MultiYN = "Y" Then
									optDc1 = split(optDc,",")(0)
									optDc2 = split(optDc,",")(1)
									strRst = strRst & "					<OrderStock Code="""&itemoption&""" Section="""&optDc1&""" Text="""&optDc2&""" Quantity="""&optLimit&""" Price="""&optaddprice&""" IsDisplayable="""&IsDisplayable&""" ChangeType=""Add""  xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
								End If
							Else
								If MultiYN = "Y" Then
									optDc1 = split(optDc,",")(0)
									optDc2 = split(optDc,",")(1)
									optDc3 = split(optDc,",")(2)
									strRst = strRst & "					<OrderStock Code="""&itemoption&""" Section="""&optDc1&""" Text="""&optDc2&""" Text2="""&optDc3&""" Quantity="""&optLimit&""" Price="""&optaddprice&""" IsDisplayable="""&IsDisplayable&""" ChangeType=""Add""  xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
								End If
							End If
							rsget.MoveNext
						Loop
					end if
					rsget.Close
					strRst = strRst & "				</ItemStock>"
				End If
			End If
		End If
		getAuctionOptParamtoReg = strRst
	End Function

	'기본정보 수정 Soap XML
	Public Function getAuctionItemInfoEditParameter()
		Dim strRst, tt, isMadeInKorea, ImportedCode, ImportedAgency
		If Fsourcearea = "한국" OR Fsourcearea = "대한민국" Then
			isMadeInKorea = "Domestic"		'국내
		Else
			isMadeInKorea = "Imported"		'수입
			ImportedCode = getNationName2Code(Fsourcearea,ImportedAgency)
			'CoastalWaters  연근해
			'Domestic  국내
			'Imported  수입
			'Ocean  원양산
			'Unknown  모름
		End If

		'유효한 지역코드가 아닙니다. 에러 때문에 하단을 Unknown으로..2015-10-15 13:32 김진영 수정
		If (FaccFailCNT > 0 AND InStr(FlastErrStr, "유효한 지역코드가 아닙니다") > 0) Then
			isMadeInKorea = "Unknown"
		End If

 		If len(FDepthCode) = 7 Then FDepthCode = "0"&CStr(FDepthCode)

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "	<soap:Header>"
		strRst = strRst & "		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
		strRst = strRst & "			<Value>"&auctionTicket&"</Value>"
		strRst = strRst & "		</EncryptedTicket>"
		strRst = strRst & "	</soap:Header>"
		strRst = strRst & "	<soap:Body>"
		strRst = strRst & "		<ReviseItem xmlns=""http://www.auction.co.kr/APIv1/ShoppingService"">"
		strRst = strRst & "			<req Version=""1"">"
		strRst = strRst & "				<Item xmlns=""http://schema.auction.co.kr/Arche.Sell3.Service.xsd"""
		strRst = strRst & "					ItemID="""&FAuctionGoodNo&""""							'옥션 상품 코드
		strRst = strRst & "					BrandCode=""3104"""										'브랜드 코드(텐바이텐 : 3104)
		strRst = strRst & "					ItemStatusType=""New"""									'New 신상품
		strRst = strRst & "					Name="""&getItemNameFormat&""""							'상품명
		If Fitemid = "1295914" Then
			strRst = strRst & "				 	CategoryCode=""28130303""" 						'카테고리 코드
		Else
			strRst = strRst & "				 	CategoryCode="""&FDepthCode&"""" 						'카테고리 코드
		End If
		strRst = strRst & "					Price="""&Clng(GetRaiseValue(MustPrice/10)*10)&""""		'가격
		strRst = strRst & "				 	BuyableQuantity="""&getOrderMaxNum&"""" 				'최대 구매허용 수량
		strRst = strRst & "				 	BuyLimitTypeCode=""OnceLimited"""						'최대 구매 허용 수량코드 | OnceLimited	1회제한, OneManLimited	1인제한, PeriodLimited	기간제한, Unlimited	제한없음
		strRst = strRst & "					PlaceOfOrigin="""&isMadeInKorea&""""					'원산지
		If FIsbn13 <> "" Then
		strRst = strRst & "				 	ISBN="""&FIsbn13&""""									'ISBN 코드
		strRst = strRst & "				 	ProductionDate="""&FSellSTDate&""""						'제조일자/발행일자
		End If
		strRst = strRst & "				 	WishKeyword="""&getItemKeyword&""""						'희망검색어
		strRst = strRst & "					IsAdult="""&Chkiif(IsAdultItem() = "Y", "true", "false")&""""	'성인물품 여부
		strRst = strRst & "					IsPCS=""true"""											'가격비교 사이트 등록여부
		strRst = strRst & getAuctionDate()
		strRst = strRst & "					SellingArea=""Nationwide"">"							'판매 지역
		If isImageChanged Then																		'MayBe 이미지 수정됐을 시 수정(느리기땜에)
			strRst = strRst & getAuctionAddImageParam()
		End If

		strRst = strRst & "					<ShippingFee xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"""
		strRst = strRst & "						ShippingType=""Door2Door"""							'배송방법 | 옥션코드 : Door2Door(택배)
		strRst = strRst & "						ShippingFeeChargeType=""Amount"""
		strRst = strRst & "						IsPrepayable=""false"""								'선결제 가능 여부
		strRst = strRst & "						IsArrival=""false"""
		strRst = strRst & "						IsDefault=""false"">"
		strRst = strRst & "						<ShipingFeeType>SellerShipping</ShipingFeeType>"	'배송비 부담 방식 | 옥션코드 : SellerShipping(판매자 조건부)
		strRst = strRst & "						<ShippingPlaceSeq>1557709</ShippingPlaceSeq>"		'배송출하지 SEQ번호 | 1557709
		strRst = strRst & "						<ShippingPolicyNo>134383728</ShippingPolicyNo>"		'판매자 묶음배송비 선택시 배송출하지에 묶여있는 묶음배송정책 번호 | '134383728 5만원 3천원 코드로 바뀌어야 함
		strRst = strRst & "					</ShippingFee>"
		strRst = strRst & "					<ItemReturn DeliveryAgency=""cjgls"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"">"
		strRst = strRst & "						<Address SellerAddrNo=""102197944"" />"				'판매자주소록 No
		strRst = strRst & "						<ExtraInfo ReturnFee=""3000"" />"
		strRst = strRst & "					</ItemReturn>"

		strRst = strRst & "					<ItemContentsHtml ItemHtml="""&replaceRst(getAuctionItemContParamToReg)&""" ItemPromotionHtml="""" ItemAddHtml="""" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />" '상품상세 HTML
		If (isMadeInKorea = "Imported") OR (FIsChildrenCate = "Y") OR (FIsLifeCate = "Y") OR (FIsElecCate = "Y")  Then	'국산이 아닐 때 OR 상품인증번호 필요한 카테라면 꼭 아래것을 호출해야 한다함
		strRst = strRst & "					<ItemExtra xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"">"
			If isMadeInKorea = "Imported" Then
		strRst = strRst & "						<ImportedItem ImportAgency="""&ImportedAgency&""" IsMultiple=""false"" Nation="""&ImportedCode&""" />"
			End If
		strRst = strRst & getAuctionCertInfo()
		strRst = strRst & "					</ItemExtra>"
		End If
		strRst = strRst & "				</Item>"
		strRst = strRst & "			</req>"
		strRst = strRst & "		</ReviseItem>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
		getAuctionItemInfoEditParameter = strRst
'response.write strRst
'response.end
	End Function

	Public Function fnCertCodes(iitemid, iGubun, icertNo, icertDiv, icertDate, imodelName)
		Dim strSql, addSql
		If iGubun = "ELEC" Then
			addSql = addSql & " and r.safetyDiv in ('10', '20', '30') "
		ElseIf iGubun = "LIFE" Then
			addSql = addSql & " and r.safetyDiv in ('40', '50', '60') "
		Else
			addSql = addSql & " and r.safetyDiv in ('70', '80', '90') "
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP 1 r.certNum "
		strSql = strSql & "	,Case When r.safetyDiv in ('10', '40', '70') THEN 'SafeCert' "
		strSql = strSql & "		  When r.safetyDiv in ('20', '50', '80') THEN 'SafeCheck' "
		strSql = strSql & " 	  When r.safetyDiv in ('30', '60', '90') THEN 'SupplierCheck' end as safetyStr "
		strSql = strSql & " ,convert(date, f.certDate) as certDate, f.modelName " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg as r " & vbcrlf
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on r.itemid = f.itemid " & vbcrlf
		strSql = strSql & " WHERE r.itemid='"&iitemid&"' "
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			icertNo		= rsget("certNum")
			icertDiv	= rsget("safetyStr")
			icertDate	= rsget("certDate")
			imodelName	= rsget("modelName")
		End If
		rsget.Close
	End Function

	Public Function getAuctionCertInfo()
		Dim buf, certNo, certDiv, certDate, modelName, strRst
		If FIsChildrenCate = "Y" Then
			Call fnCertCodes(FItemid, "CHILD", certNo, certDiv, certDate, modelName)
			buf = buf & "		<IntegrateSafeCert>"
			If certNo <> "" Then
				buf = buf & "			<IntegrateSafeCertGroupList CertificationGroupNo=""Child"" CertificationType=""RequireCert"" >"
				buf = buf & "				<IntegrateSafeCertDetailList CertificationNo="""&certNo&""" CertificationTargetCode="""&certDiv&""" CertificationStatus=""적합"" CertificationDate="""&certDate&""" CertificationType="""" FirstCertificationNo="""" ProductName="""" ModelName="""&modelName&""" CertificationImgUrl="""" InputType=""SystemInput"" />"
				buf = buf & "			</IntegrateSafeCertGroupList>"
			Else
				buf = buf & "			<IntegrateSafeCertGroupList CertificationGroupNo=""Child"" CertificationType=""AddDescription"" />"
			End If
			buf = buf & "		</IntegrateSafeCert>"
		End If

		If FIsLifeCate = "Y" Then
			Call fnCertCodes(FItemid, "LIFE", certNo, certDiv, certDate, modelName)
			buf = buf & "		<IntegrateSafeCert>"
			If certNo <> "" Then
				buf = buf & "			<IntegrateSafeCertGroupList CertificationGroupNo=""Life"" CertificationType=""RequireCert"" >"
				buf = buf & "				<IntegrateSafeCertDetailList CertificationNo="""&certNo&""" CertificationTargetCode="""&certDiv&""" CertificationStatus=""적합"" CertificationDate="""&certDate&""" CertificationType="""" FirstCertificationNo="""" ProductName="""" ModelName="""&modelName&""" CertificationImgUrl="""" InputType=""SystemInput"" />"
				buf = buf & "			</IntegrateSafeCertGroupList>"
			Else
				buf = buf & "			<IntegrateSafeCertGroupList CertificationGroupNo=""Life"" CertificationType=""AddDescription"" />"
			End If
			buf = buf & "		</IntegrateSafeCert>"
		End If

		If FIsElecCate = "Y" then
			Call fnCertCodes(FItemid, "ELEC", certNo, certDiv, certDate, modelName)
			buf = buf & "		<IntegrateSafeCert>"
			If certNo <> "" Then
				buf = buf & "			<IntegrateSafeCertGroupList CertificationGroupNo=""Electric"" CertificationType=""RequireCert"" >"
				buf = buf & "				<IntegrateSafeCertDetailList CertificationNo="""&certNo&""" CertificationTargetCode="""&certDiv&""" CertificationStatus=""적합"" CertificationDate="""&certDate&""" CertificationType="""" FirstCertificationNo="""" ProductName="""" ModelName="""&modelName&""" CertificationImgUrl="""" InputType=""SystemInput"" />"
				buf = buf & "			</IntegrateSafeCertGroupList>"
			Else
				buf = buf & "			<IntegrateSafeCertGroupList CertificationGroupNo=""Electric"" CertificationType=""AddDescription"" />"
			End If
			buf = buf & "		</IntegrateSafeCert>"
		End If
		getAuctionCertInfo = buf
	End Function

	'옵션 삭제 Soap XML
	Public Function getAuctionOPTDeleteParameter()
		Dim strSQL, strRst, strRst1, strRst2, strRst3
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "	<soap:Header>"
		strRst = strRst & "		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
		strRst = strRst & "			<Value>"&auctionTicket&"</Value>"
		strRst = strRst & "		</EncryptedTicket>"
		strRst = strRst & "	</soap:Header>"
		strRst = strRst & "	<soap:Body>"
		strRst = strRst & "		<ReviseItemStock xmlns=""http://www.auction.co.kr/APIv1/ShoppingService"">"
		strRst = strRst & "			<req Version=""1"">"
		strRst = strRst & "				<ItemStock ItemID="""&FAuctionGoodNo&""" Type=""NotAvailable"" OptionStockType=""NotAvailable"" OptVerType=""New"" ImageMatchingFinishYN=""false"" OptRepImageLevel=""0"" OptDetailImageLevel=""0"" IsStockQtyMng=""false"" UseOptionBuyQty=""false"" xmlns=""http://schema.auction.co.kr/Arche.Sell3.Service.xsd"">"
		strRst = strRst & "					<Seller MemberID=""10x10store"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
		strRst = strRst & "					<OrderStock Section=""_"" Text=""_"" Quantity=""1"" Price=""0"" IsDisplayable=""true"" StockMasterSeqNo=""0"" SkuMatchingVerNo=""0"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
		strRst = strRst & "				</ItemStock>"
		strRst = strRst & "			</req>"
		strRst = strRst & "		</ReviseItemStock>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
		getAuctionOPTDeleteParameter = strRst
	End Function

	'단품 재고 수정 Soap XML
	Public Function getAuctionDanPoomModParameter()
		Dim strSQL, strRst, danPoomCode
		strSQL = ""
		strSQL = " SELECT TOP 1 outmallOptCode FROM db_item.dbo.tbl_outmall_regedoption WHERE itemid = '"&FItemid&"' and mallid = '"&CMALLNAME&"'  "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			danPoomCode = rsget("outmallOptCode")
		End If
		rsget.Close

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "	<soap:Header>"
		strRst = strRst & "		<EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
		strRst = strRst & "			<Value>"&auctionTicket&"</Value>"
		strRst = strRst & "		</EncryptedTicket>"
		strRst = strRst & "	</soap:Header>"
		strRst = strRst & "	<soap:Body>"
		strRst = strRst & "		<ReviseItemStock xmlns=""http://www.auction.co.kr/APIv1/ShoppingService"">"
		strRst = strRst & "			<req Version=""1"">"
		If FItemdiv = "06" Then		'단품이면서 주문제작문구 있는 상품
			strRst = strRst & "				<ItemStock ItemID="""&FAuctionGoodNo&""" Type=""BuyerDescriptive"" OptionStockType=""NotAvailable"" IsStockQtyMng=""false"" UseOptionBuyQty=""false"" xmlns=""http://schema.auction.co.kr/Arche.Sell3.Service.xsd"">"
			strRst = strRst & "					<OrderStock StockNo="""&danPoomCode&""" Quantity="""&getLimitAuctionEa&""" Price=""0"" IsDisplayable=""true"" ChangeType=""Update"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
			strRst = strRst & "				</ItemStock>"
		Else						'단품이면서 주문제작문구가 없는 상품
			strRst = strRst & "				<ItemStock ItemID="""&FAuctionGoodNo&""" Type=""NotAvailable"" OptionStockType=""NotAvailable"" OptVerType=""New"" ImageMatchingFinishYN=""false"" OptRepImageLevel=""0"" OptDetailImageLevel=""0"" IsStockQtyMng=""false"" UseOptionBuyQty=""false"" xmlns=""http://schema.auction.co.kr/Arche.Sell3.Service.xsd"">"
			strRst = strRst & "					<Seller MemberID=""10x10store"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
			strRst = strRst & "					<OrderStock StockNo="""&danPoomCode&""" Section=""_"" Text=""_"" Quantity="""&getLimitAuctionEa&""" Price=""0"" ChangeType=""Update"" IsDisplayable=""true"" StockMasterSeqNo=""0"" SkuMatchingVerNo=""0"" xmlns=""http://schema.auction.co.kr/Arche.Service.xsd"" />"
			strRst = strRst & "				</ItemStock>"
		End If
		strRst = strRst & "			</req>"
		strRst = strRst & "		</ReviseItemStock>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
		getAuctionDanPoomModParameter = strRst
	End Function

End Class

Class CAuction
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
	Public Sub getAuctionNotRegOneItem
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
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum "
		strSql = strSql & "	, isNULL(R.auctionStatCD,-9) as auctionStatCD "
		strSql = strSql & "	, UC.socname_kor, am.depthCode "
'		strSql = strSql & "	, isnull(tm.notinCate, '') as notinCate, tm.SafeAuthType, isnull(tm.AuthItemTypeCode, '') as AuthItemTypeCode, tm.isChildrenCate, tm.overlap, tm.RawMaterialsType "
		strSql = strSql & "	, tm.isChildrenCate, tm.isLifeCate, tm.isElecCate, tm.RawMaterialsType, isNull(c.isbn13, '') as isbn13, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		strSql = strSql & "	, CONVERT(VARCHAR(10), isNull(sellSTDate, getdate()), 23) as sellSTDate "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_auction_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_auction_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_auction_category_New as tm on am.depthCode = tm.depthCode "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_auction_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		IF (CUPJODLVVALID) then
'		    strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		ELSE
		    strSql = strSql & " and (i.deliveryType<>9)"
	    END IF
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.itemdiv not in ('21', '23', '30') "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "				'플라워/화물배송/해외직구
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
'		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)>=" & CMAXMARGIN & ")"
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &") ) "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"' and i.mwdiv <> 'M')) "	'등록제외 카테고리
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_etcmall.dbo.tbl_auction_regItem WHERE auctionStatCD >= 3) "	''등록완료이상은 등록안됨.										'롯데등록상품 제외
		strSql = strSql & " and cm.mapCnt is Not Null "
		strSql = strSql & "		"	& addSql											'카테고리 매칭 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CAuctionItem
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
			If (IsNULL(rsget("basicImage600")) or (rsget("basicImage600")="")) Then
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
			ELSE
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic600/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage600")
			End If
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FAuctionStatCD		= rsget("auctionStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.Fsocname_kor		= rsget("socname_kor")
				FOneItem.FDepthCode			= rsget("depthCode")
				FOneItem.FbasicimageNm 		= rsget("basicimage")

'				FOneItem.FNotinCate 		= rsget("notinCate")
'				FOneItem.FSafeAuthType 		= rsget("SafeAuthType")
'				FOneItem.FAuthItemTypeCode 	= rsget("AuthItemTypeCode")
'				FOneItem.FOverlap 			= rsget("overlap")
				FOneItem.FIsChildrenCate 	= rsget("isChildrenCate")
				FOneItem.FIsLifeCate 		= rsget("isLifeCate")
				FOneItem.FIsElecCate 		= rsget("isElecCate")
				FOneItem.FRawMaterialsType 	= rsget("RawMaterialsType")
				FOneItem.FIsbn13 			= rsget("isbn13")
				FOneItem.FSellSTDate		= rsget("sellSTDate")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FOutmallstandardMargin	= rsget("outmallstandardMargin")
		End If
		rsget.Close
	End Sub

	'// 미등록 옵션(등록용)
	Public Sub getAuctionNotOptOneItem
		Dim strSql, addSql, i
		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, J.AuctionGoodNo, isnull(J.APIadditem, 'N') as APIadditem, isnull(J.APIaddopt, 'N') as APIaddopt "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_auction_regItem as J on i.itemid = J.itemid"
		strSql = strSql & " WHERE 1=1 "
		strSql = strSql & " and J.itemid = '"&FRectItemID&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CAuctionItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FitemDiv			= rsget("itemdiv")
				FOneItem.FoptionCnt			= rsget("optionCnt")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.ForgPrice			= rsget("orgPrice")
				FOneItem.ForgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FAuctionGoodNo		= rsget("AuctionGoodNo")
				FOneItem.FAPIadditem		= rsget("APIadditem")
				FOneItem.FAPIaddopt			= rsget("APIaddopt")
		End If
		rsget.Close
	End Sub

	Public Sub getAuctionEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
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
		strSql = strSql & "	, m.auctionGoodNo, m.auctionprice, m.auctionSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
'		strSql = strSql & "	, isnull(tm.notinCate, '') as notinCate, tm.SafeAuthType, isnull(tm.AuthItemTypeCode, '') as AuthItemTypeCode, tm.isChildrenCate, tm.overlap, tm.RawMaterialsType "
		strSql = strSql & "	, tm.isChildrenCate, tm.isLifeCate, tm.isElecCate, tm.RawMaterialsType, am.depthCode, isNull(c.isbn13, '') as isbn13, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		strSql = strSql & "	, CONVERT(VARCHAR(10), isNull(sellSTDate, getdate()), 23) as sellSTDate "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & "		or i.deliveryType in ('7','6') "
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
'		strSql = strSql & "		or ((i.sailyn = 'N') and ( Round(((i.sellcash-i.buycash)/i.sellcash)*100,0) < "&CMAXMARGIN&" )) "

		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "

		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"' and i.mwdiv <> 'M')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_auction_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_auction_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_auction_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_auction_category_New as tm on am.depthCode = tm.depthCode "
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and m.APIadditem = 'Y' "
		strSql = strSql & " and m.APIaddopt = 'Y' "
		strSql = strSql & " and m.APIaddgosi = 'Y' "
		strSql = strSql & " and m.auctionStatCD = 7 "
		strSql = strSql & addSql
		strSql = strSql & " and m.auctionGoodNo is Not Null "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CAuctionItem
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
			If (IsNULL(rsget("basicImage600")) or (rsget("basicImage600")="")) Then
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
			ELSE
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic600/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage600")
			End If
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FauctionGoodNo		= rsget("auctionGoodNo")
				FOneItem.FAuctionprice		= rsget("auctionprice")
				FOneItem.FAuctionSellYn		= rsget("auctionSellYn")

	            FOneItem.FoptionCnt         = rsget("optionCnt")
	            FOneItem.FregedOptCnt       = rsget("regedOptCnt")
	            FOneItem.FaccFailCNT        = rsget("accFailCNT")
	            FOneItem.FlastErrStr        = rsget("lastErrStr")
	            FOneItem.Fdeliverytype      = rsget("deliverytype")
	            FOneItem.FrequireMakeDay    = rsget("requireMakeDay")

	            FOneItem.FinfoDiv       = rsget("infoDiv")
	            FOneItem.Fsafetyyn      = rsget("safetyyn")
	            FOneItem.FsafetyDiv     = rsget("safetyDiv")
	            FOneItem.FsafetyNum     = rsget("safetyNum")
	            FOneItem.FmaySoldOut    = rsget("maySoldOut")
	            FOneItem.Fregitemname    = rsget("regitemname")
                FOneItem.FregImageName		= rsget("regImageName")
                FOneItem.FbasicImageNm		= rsget("basicimage")

'				FOneItem.FNotinCate 		= rsget("notinCate")
'				FOneItem.FSafeAuthType 		= rsget("SafeAuthType")
'				FOneItem.FAuthItemTypeCode 	= rsget("AuthItemTypeCode")
'				FOneItem.FOverlap 			= rsget("overlap")
				FOneItem.FIsChildrenCate 	= rsget("isChildrenCate")
				FOneItem.FIsLifeCate 		= rsget("isLifeCate")
				FOneItem.FIsElecCate 		= rsget("isElecCate")
				FOneItem.FRawMaterialsType 	= rsget("RawMaterialsType")
				FOneItem.FDepthCode			= rsget("depthCode")
				FOneItem.FIsbn13 			= rsget("isbn13")
				FOneItem.FSellSTDate		= rsget("sellSTDate")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FOutmallstandardMargin	= rsget("outmallstandardMargin")
		End If
		rsget.Close
	End Sub

End Class

'옥션 상품코드 얻기
Function getAuctionGoodno(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 Auctiongoodno FROM db_etcmall.dbo.tbl_auction_regitem WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		getAuctionGoodno = rsget("Auctiongoodno")
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

Function getNationName2Code(iname, byref inationname)
	Dim sqlStr , retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT top 1 code, nationname" & VBCRLF
	sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_auction_Nation " & VBCRLF
	sqlStr = sqlStr & " WHERE nationname='"&html2db(iname)&"'"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
		retVal = rsget("code")
		inationname = rsget("nationname")
	end if
	rsget.Close

	If (retVal = "") Then
		retVal="240"
		inationname = "E.T.C"
	End If

	getNationName2Code = retVal
End Function

function replaceRst(checkvalue)
	dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, "&", "&amp;")
    v = replace(v, """", "&quot;")
    v = replace(v, "최대", "")			'2017-01-31 김진영 수정.."최대" 가 금칙어로 지정됨
	'v = Replace(v,"<br>","&#xA;")
	'v = Replace(v,"</br>","&#xA;")
	'v = Replace(v,"<br />","&#xA;")
	v = Replace(v,"<","&lt;")
	v = Replace(v,">","&gt;")
    replaceRst = v
end function

Function getAllRegChk(iitemid)
	Dim sqlStr , retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT Count(*) as cnt " & VBCRLF
	sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_auction_regItem " & VBCRLF
	sqlStr = sqlStr & " WHERE itemid='"&iitemid&"'"
	sqlStr = sqlStr & " and APIadditem = 'Y' "
	sqlStr = sqlStr & " and APIaddopt = 'Y' "
	sqlStr = sqlStr & " and APIaddgosi = 'Y' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("cnt") = 0 Then
		getAllRegChk = "N"
	Else
		getAllRegChk = "Y"
	End If
	rsget.Close
End Function

Function getAllRegChk2(iitemid)
	Dim sqlStr , retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT Count(*) as cnt " & VBCRLF
	sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_auction_regItem " & VBCRLF
	sqlStr = sqlStr & " WHERE itemid='"&iitemid&"'"
	sqlStr = sqlStr & " and APIadditem = 'Y' "
	sqlStr = sqlStr & " and APIaddopt = 'Y' "
	sqlStr = sqlStr & " and APIaddgosi = 'Y' "
	sqlStr = sqlStr & " and auctionStatCD = 7 "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("cnt") = 0 Then
		getAllRegChk2 = "N"
	Else
		getAllRegChk2 = "Y"
	End If
	rsget.Close
End Function

%>
