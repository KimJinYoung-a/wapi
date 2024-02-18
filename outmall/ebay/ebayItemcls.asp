<!-- #include virtual="/outmall/ebay/inc_gubunChk.asp"-->
<%
CONST CMAXMARGIN = 15
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST APIURL = "http://api.11st.co.kr/rest"
CONST APISSLURL = "https://sa.esmplus.com"
CONST APIkey = "a2319e071dbc304243ee60abd07e9664"
CONST CDEFALUT_STOCK = 99999

Class CEbayItem
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
	Public FBrandName
	Public FBrandNameKor
	Public FItemsize
	Public FItemsource
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
	Public FCateCode
	Public FSDCategoryCode
	Public Fcdmkey
	Public Fcddkey
	Public FSt11GoodNo
	Public FSt11price
	Public FSt11SellYn
	Public FIsbn13

	Public FSafeDiv
	Public FIsNeed
	Public FDepth1Code
	Public FAdultType

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

	'최대 구매 수량
	Public Function getLimitEbayEa()
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
		getLimitEbayEa = ret
	End Function

	'// 11st 판매여부 반환
	Public Function getEbaySellyn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getEbaySellyn = "Y"
			Else
				getEbaySellyn = "N"
			End If
		Else
			getEbaySellyn = "N"
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

    public function getItemNameFormat(v)
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
		If v = 1 Then
	        buf = LeftB(buf, 50)
		End If
        getItemNameFormat = buf
    end function

    Public Function GetSourcearea()
		If IsNULL(Fsourcearea) or (Fsourcearea="") then
			GetSourcearea = "."
		Else
			GetSourcearea = Fsourcearea
		End if
    End function

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
						If (cntType <> cntOpt) OR (cntOpt > 2) Then		'3중 옵션 지원안함
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

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	Public Function getEbayContParamToReg(obj, vGubun)
		Dim strRst, strSQL, tmpContent, gubunStr
		If vGubun = "A" Then
			gubunStr = "auction"
		Else
			gubunStr = "gmarket"
		End If

		strRst = ("<div align=""center"">")
		strRst = strRst & ("<p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_"&gubunStr&".jpg></p><br />")
		strRst = strRst & ("<div style=""width:100%; max-width:700px; margin:0; padding:0; margin-bottom:14px; padding-bottom:6px; background:url(http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_namebg.png) left bottom no-repeat;"">")
		strRst = strRst & ("<table cellpadding=""0"" cellspacing=""0"" width=""100%"">")
		strRst = strRst & ("<tr>")
		strRst = strRst & ("<th style=""vertical-align:middle; width:73px; height:42px; text-align:center; margin:0; padding:3px 0 0 0;""><img src=""http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_nametit.png"" alt=""상품명"" style=""vertical-align:top; display:inline;""/></th>")
		strRst = strRst & ("<td style=""width:627px; vertical-align:middle; text-align:left; font-size:14px; line-height:1.2; color:#000; font-weight:bold; font-family:dotum, dotumche, '돋움', sans-serif; margin:0; padding:4px 0 0 0;"">")
		strRst = strRst & ("<p style=""letter-spacing:-0.03em; margin:0; padding:12px 10px;"">")
		strRst = strRst & getItemNameFormat(2)
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
				strRst = strRst & (Fitemcontent & "<br />")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "<br />")
			Case Else
				strRst = strRst & (nl2br(Fitemcontent) & "<br />")
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
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%""><br />")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		'#기본 상품 설명이미지
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%""><br />")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%""><br />")

		'#배송 주의사항
		strRst = strRst & ("<br /><img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_"&gubunStr&".jpg>")
		strRst = strRst & ("</div>")
		tmpContent = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		if Not(rsget.EOF or rsget.BOF) then
			strRst = rsget("textVal")
			strRst = "<div align=""center""><p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_"&gubunStr&".jpg></p><br />" & strRst & "<br /><img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_"&gubunStr&".jpg></div>"
			tmpContent = strRst
		End If
		rsget.Close

		Set obj("itemAddtionalInfo")("descriptions") = jsObject()
			Set obj("itemAddtionalInfo")("descriptions")("kor") = jsObject()
				obj("itemAddtionalInfo")("descriptions")("kor")("type") = 2				'#상품상세정보타입 | 1 contentID(추후제공), 2 html
				obj("itemAddtionalInfo")("descriptions")("kor")("contentId") = ""		'상품상세정보 코드 | 상품상세정보타입이 1일 때 필수
				obj("itemAddtionalInfo")("descriptions")("kor")("html")	= tmpContent	'#상품상세정보 html | iframe, Script 불가
	End Function

	'// 검색어
	Public Function getItemKeyword()
		Dim arrRst, arrRst2, q, Keyword1, strRst
		If trim(Fkeywords) = "" Then Exit Function
		Fkeywords  = replace(Fkeywords,"%", "")
		Fkeywords  = replace(Fkeywords,"/", ",")
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

	Public Function getEbayImageParameter(obj)
		Dim strRst, strSQL, i, imgAdds, spImage

		Set obj("itemAddtionalInfo")("images") = jsObject()
			obj("itemAddtionalInfo")("images")("basicImgURL") = FbasicImage&"/10x10/thumbnail/600x600/quality/85/"			'#상품 기본이미지 | 최소 600x600 권장 1000x1000

 		strSQL = ""
		strSQL = strSQL & " SELECT TOP 2 gubun,ImgType,addimage_400,addimage_600,addimage_1000 "
		strSQL = strSQL & " FROM db_item.[dbo].tbl_item_addimage "
		strSQL = strSQL & " WHERE IMGTYPE = 0 "
		strSQL = strSQL & " AND itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			imgAdds = ""
			For i=1 to rsget.RecordCount
				imgAdds = imgAdds & "http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & ","
				rsget.MoveNext
			Next
			If Right(imgAdds,1) = "," Then
				imgAdds = Left(imgAdds, Len(imgAdds) - 1)
			End If
			spImage = Split(imgAdds, ",")

			If isArray(spImage) Then
				If Ubound(spImage) >= 0 Then
					obj("itemAddtionalInfo")("images")("addtionalImg1URL") = spImage(0)&"/10x10/thumbnail/600x600/quality/85/"
					If Ubound(spImage) = 1 Then
						obj("itemAddtionalInfo")("images")("addtionalImg2URL") = spImage(1)&"/10x10/thumbnail/600x600/quality/85/"
					Else
						obj("itemAddtionalInfo")("images")("addtionalImg2URL") = null
					End If
				End If
			End If
		Else
			obj("itemAddtionalInfo")("images")("addtionalImg1URL") = null
			obj("itemAddtionalInfo")("images")("addtionalImg2URL") = null
		End If
		rsget.Close
	End Function

	Public Function fnCertCodes(iGubun, icertNo, icertDiv, itype)
		Dim strSql, addSql, tmpVal
		If iGubun = "ELEC" Then
			addSql = addSql & " and r.safetyDiv in ('10', '20', '30') "
		ElseIf iGubun = "LIFE" Then
			addSql = addSql & " and r.safetyDiv in ('40', '50', '60') "
		Else
			addSql = addSql & " and r.safetyDiv in ('70', '80', '90') "
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP 1 r.certNum "
		strSql = strSql & "	,Case When r.safetyDiv in ('10', '40', '70') THEN 0 "
		strSql = strSql & "		  When r.safetyDiv in ('20', '50', '80') THEN 1 "
		strSql = strSql & " 	  When r.safetyDiv in ('30', '60', '90') THEN 2 end as safetyStr "
		strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg as r " & vbcrlf
		strSql = strSql & " WHERE r.itemid='"&FItemid&"' "
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			icertNo		= rsget("certNum")
			icertDiv	= rsget("safetyStr")
			tmpVal		= "Y"
		Else
			icertNo		= ""
			icertDiv	= ""
			tmpVal		= "N"
		End If
		rsget.Close

		If tmpVal = "Y" Then
			If icertDiv = 2 Then
				itype = 2
			Else
				itype = 0
			End If
		Else
			itype = 1
		End If
	End Function

	Public Function getEbayCertInfoParameter(obj)
		Dim certNo, certDiv, vType
		Set obj("itemAddtionalInfo")("certInfo") = jsObject()
			obj("itemAddtionalInfo")("certInfo")("gmkt") = null										'(G마켓용) 인증정보코드
			obj("itemAddtionalInfo")("certInfo")("iac") = null										'(옥션용) 인증정보코드 - 의료기기, 방송통신기기, 식품제조가공업, 건강기능식품, 친환경인증 등
			Set obj("itemAddtionalInfo")("certInfo")("safetyCerts") = jsObject()
				Call fnCertCodes("CHILD", certNo, certDiv, vType)
				Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child") = jsObject()
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child")("type") = vType										'#통합인증대상 상품 아닐경우 "인증대상아님"으로 입력 | 0 인증대상, 1 인증대상아님, 2 상품상세별도표기
				If vType = 1 Then
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child")("details")= null
				Else
					Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child")("details") = jsArray()
						Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child")("details")(null) = jsObject()
							obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child")("details")(null)("certId") = certNo			'통합어린이 인증코드
							obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child")("details")(null)("certTargetCode") = certDiv	'통합어린이인증품목 | 0 안전인증, 1 안전확인, 3 공급자적합성확인
				End If

				Call fnCertCodes("ELEC", certNo, certDiv, vType)
				Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric") = jsObject()
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("type") = vType										'#통합인증대상 상품 아닐경우 "인증대상아님"으로 입력 | 0 인증대상, 1 인증대상아님, 2 상품상세별도표기
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("mandatorySafetySign") = "UnknownOrNone"			'병행수입여부 | BuyingAgent : 구매대행, ParallelImport 병행수입, UnknownOrNone : 해당사항없음
				If vType = 1 Then
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("details") = null
				Else
					Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("details") = jsArray()
						Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("details")(null) = jsObject()
							obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("details")(null)("certId") = certNo			'통합전기 인증코드
							obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("details")(null)("certTargetCode") = certDiv'통합전기인증품목 | 0 안전인증, 1 안전확인, 3 공급자적합성확인
				End If

				Call fnCertCodes("LIFE", certNo, certDiv, vType)
				Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life") = jsObject()
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("type") = vType											'#통합인증대상 상품 아닐경우 "인증대상아님"으로 입력 | 0 인증대상, 1 인증대상아님, 2 상품상세별도표기
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("mandatorySafetySign") = "UnknownOrNone"				'병행수입여부 | BuyingAgent : 구매대행, ParallelImport 병행수입, UnknownOrNone : 해당사항없음
				If vType = 1 Then
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("details") = null
				Else
					Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("details") = jsArray()
						Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("details")(null) = jsObject()
							obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("details")(null)("certId") = certNo				'통합생활용품 인증코드
							obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("details")(null)("certTargetCode") = certDiv	'통합생활용품인증품목 | 0 안전인증, 1 안전확인, 3 공급자적합성확인
				End If
				Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("harmful") = jsObject()
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("harmful")("type") = 2											'통합위해루려제품인증타입 | 설정없으면 디폴트로 상세설명표기로 설정, 0 인증대상, 1 인증대상아님, 2 상품상세별도표기
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("harmful")("certId") = null										'통합자가검사번호 | type > 0일 떄 필수
	End Function

	Public Function get11stOptionParam()
		Dim strSql, strRst1, strRst2, strRst3, i, optNm, optDc, chkMultiOpt, optLimit, arrList1, arrList2, optCode, j, optionname, optaddprice, tmpoName
		Dim voptsellyn, visusing
		strRst1 = ""
		strRst2 = ""
		strRst3 = ""
		chkMultiOpt = false
		If FoptionCnt > 0 Then
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql,dbget,1
		    if not rsget.Eof then
		    	chkMultiOpt = true
		        arrList1 = rsget.getRows()
		    end if
		    rsget.close
		End If
		strRst1 = strRst1 & "	<optSelectYn>Y</optSelectYn>"									'선택형 옵션 여부 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요.
		strRst1 = strRst1 & "	<txtColCnt>1</txtColCnt>"										'고정값 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 옵션을 등록하실 경우 1 고정값을 주셔야 합니다.
'		If chkMultiOpt = True Then
'		strRst1 = strRst1 & "	<optionAllQty>"&getLimit11stEa&"</optionAllQty>"				'멀티옵션 일괄재고수량 설정 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. "상품상세 옵션값 노출 방식 선택"을 생략하실 경우 등록순 옵션이 노출됩니다. "멀티옵션" 방식이 아닌 "싱글옵션" 방식 일 경우는 Element는 생략해주셔야 합니다. 멀티옵션은 옵션별 재고 수량 설정이 api 에서는 불가합니다. 일괄설정만 가능.
'		strRst1 = strRst1 & "	<optionAllAddPrc>0</optionAllAddPrc>"							'멀티옵션 옵션가 0원 설정 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. "상품상세 옵션값 노출 방식 선택"을 생략하실 경우 등록순 옵션이 노출됩니다. "멀티옵션" 방식이 아닌 "싱글옵션" 방식 일 경우는 Element는 생략해주셔야 합니다. 멀티옵션은 옵션별 옵션가 설정이 api 에서는 불가합니다. 0원만 입력 가능
'		strRst1 = strRst1 & "	<optionAllAddWght/>"											'멀티옵션 일괄옵션추가무게 설정 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. "상품상세 옵션값 노출 방식 선택"을 생략하실 경우 등록순 옵션이 노출됩니다. "멀티옵션" 방식이 아닌 "싱글옵션" 방식 일 경우는 Element는 생략해주셔야 합니다. 멀티옵션은 옵션별 옵션무게 설정이 api 에서는 불가합니다. 일괄설정만 가능.
'		End If
		strRst1 = strRst1 & "	<prdExposeClfCd>00</prdExposeClfCd>"							'상품상세 옵션값 노출 방식 선택 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 00 : 등록순, 01 : 옵션값 가나다순, 02 : 옵션값 가나다 역순, 03 : 옵션가격 낮은 순, 04 : 옵션가격 높은 순
		strRst1 = strRst1 & "	<optUpdateYn>Y</optUpdateYn>"
'		strRst1 = strRst1 & "	<optMixYn/>"													'전체옵션 조합여부 |"옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. Y : 정의한 전체 옵션값이 조합되어 멀티옵션으로 등록, N : 옵션 매핑Key에 존재하는(선택 된) 값으로만 멀티 옵션 등록
		If chkMultiOpt = True Then
			strSql = "select typeseq, optionTypeName, optionKindName, optaddPrice from db_item.[dbo].[tbl_item_option_Multiple] where itemid = " & FItemid & " ORDER BY Typeseq, kindSeq "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If not rsget.Eof Then
				arrList2 = rsget.getRows()
			End If
			rsget.close

			For i = 0 To Ubound(arrList1, 2)
				strRst2 = strRst2 & "	<ProductRootOption>"											'ProductRootOption |"옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요.
				strRst2 = strRst2 & "		<colTitle>"&arrList1(2,i)&"</colTitle>"						'옵션명 | 40Byte 까지만 입력가능하며 특수 문자[',",%,&,<,>,#,†]는 입력할 수 없습니다.
				For j = 0 to Ubound(arrList2, 2)
					If arrList1(1,i) = arrList2(0,j) then

						arrList2(2,j) = replace(arrList2(2,j), "&", "+")			'2017-06-05 김진영..혜리 대리 요청으로 &->+로 수정

						strRst2 = strRst2 & "		<ProductOption>"
						strRst2 = strRst2 & "			<colOptPrice>0</colOptPrice>"					'옵션가 | 기본 판매가의 +100%/-50%까지 설정하실 수 있습니다. 옵션가격이 0원인 상품이 반드시 1개 이상 있어야 합니다.
						strRst2 = strRst2 & "			<colValue0>"&arrList2(2,j)&"</colValue0>"		'옵션값 | 50Byte 까지만 입력가능하며 특수 문자[',",%,&,<,>,#,†]는 입력할 수 없습니다. 한 상품안에서 옵션값은 중복이 될수 없습니다.
						strRst2 = strRst2 & "		</ProductOption>"
					end if
				Next
				strRst2 = strRst2 & "	</ProductRootOption>"
			Next

			strSql = "Select itemoption, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice, isUsing, optsellyn "
			strSql = strSql & " From [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
'			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "		'2017-05-17 김진영 수정
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				strRst2 = strRst2 & "	<ProductOptionExt>"												'ProductOptionExt | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요.
				Do until rsget.EOF
					optCode		= FItemid&"_"&rsget("itemoption")
					optaddprice = rsget("optaddprice")
				    optLimit	= rsget("optLimit")
				    tmpoName	= db2html(rsget("optionname"))
				    visUsing= rsget("isUsing")
				    voptsellyn= rsget("optsellyn")

					optionname = ""
					For i = 0 To Ubound(arrList1, 2)
						If Ubound(Split(tmpoName, ",")) > 0 Then				'2017-06-15 김진영 추가
							optionname = optionname & arrList1(2,i) &":"&Split(tmpoName, ",")(i) &"†"
						End If
					Next

					If Right(optionname,1) = "†" Then
						optionname = Left(optionname, Len(optionname) - 1)
					End If

				    optLimit = optLimit - 5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
			    	If voptsellyn <> "Y" Then optLimit = 0	'2017-05-17 김진영 수정
			    	If visUsing <> "Y" Then optLimit = 0	'2017-05-17 김진영 수정

			'		If optLimit > 0 Then		'2017-05-17 김진영 수정
						strRst2 = strRst2 & "		<ProductOption>"
						strRst2 = strRst2 & "			<useYn>"&chkiif(optLimit>0,"Y","N")&"</useYn>"
						strRst2 = strRst2 & "			<colOptPrice>"&optaddprice&"</colOptPrice>"				'옵션가 | 기본 판매가의 +100%/-50%까지 설정하실 수 있습니다. 옵션가격이 0원인 상품이 반드시 1개 이상 있어야 합니다.
						strRst2 = strRst2 & "			<colCount>"&optLimit&"</colCount>"			'옵션재고수량 | 멀티옵션일 경우는 일괄설정이 되므로 입력하시면 안됩니다. 옵션상태(useYn)가 N일 때만 0 입력 가능합니다.
						strRst2 = strRst2 & "			<colSellerStockCd>"&optCode&"</colSellerStockCd>"		'셀러재고번호 | 셀러가 사용하는 재고번호
						strRst2 = strRst2 & "			<optionMappingKey>"&optionname&"</optionMappingKey>"	'옵션매핑Key | 멀티옵션의 조합된 옵션을 매핑하기 위한 Key(예: 옵션명1:옵션값1†옵션명2:옵션값2)
						strRst2 = strRst2 & "		</ProductOption>"
			'		End If
					rsget.MoveNext
				Loop
				strRst2 = strRst2 & "	</ProductOptionExt>"
			End If
			rsget.Close
		Else
			strSql = "Select itemoption, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
			strSql = strSql & " From [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				If db2Html(rsget("optionTypeName"))<>"" Then
					optNm = Replace(db2Html(rsget("optionTypeName")),":","")
				Else
					optNm = "옵션"
				End If
				strRst1 = strRst1 & "	<colTitle>"&optNm&"</colTitle>"
				Do until rsget.EOF
					optCode		= FItemid&"_"&rsget("itemoption")
					optaddprice = rsget("optaddprice")
				    optLimit	= rsget("optLimit")
				    optionname	= db2html(rsget("optionname"))
				    optionname = replace(optionname, "&", "+")			'2017-06-05 김진영..혜리 대리 요청으로 &->+로 수정
				    optionname = replace(optionname, ",", "+")
				    optLimit = optLimit - 5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
					If optLimit > 0 Then
						strRst2 = strRst2 & "	<ProductOption>"								'ProductOption | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요.
						strRst2 = strRst2 & "		<useYn>Y</useYn>"							'옵션상태 | 멀티옵션일 경우는 지원하지 않는 기능입니다. Y : 사용함, N : 품절
						strRst2 = strRst2 & "		<colOptPrice>"&optaddprice&"</colOptPrice>"	'옵션가 | 기본 판매가의 +100%/-50%까지 설정하실 수 있습니다. 옵션가격이 0원인 상품이 반드시 1개 이상 있어야 합니다.
						strRst2 = strRst2 & "		<colValue0>"&optionname&"</colValue0>"		'옵션값 | 50Byte 까지만 입력가능하며 특수 문자[',",%,&,<,>,#,†]는 입력할 수 없습니다. 한 상품안에서 옵션값은 중복이 될수 없습니다
						strRst2 = strRst2 & "		<colCount>"&optLimit&"</colCount>"			'옵션재고수량 | 멀티옵션일 경우는 일괄설정이 되므로 입력하시면 안됩니다. 옵션상태(useYn)가 N일 때만 0 입력 가능합니다.
						strRst2 = strRst2 & "		<colSellerStockCd>"&optCode&"</colSellerStockCd>"	'셀러재고번호 | 셀러가 사용하는 재고번호
						strRst2 = strRst2 & "	</ProductOption>"
					End If
					rsget.MoveNext
				Loop
			End If
			rsget.Close
		End If

		If FitemDiv = "06" Then
			strRst3 = strRst3 & "	<ProductCustOption>"										'옵션등록 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 구매자작성형 옵션의 등록 최대 5개까지 등록 가능
			strRst3 = strRst3 & "		<colOptName>텍스트를 입력하세요</colOptName>"				'구매자 작성형 옵션명 | 옵션명 최대 공백포함 한글10자/영문(숫자)20자)까지 입력가능하며 특수 문자[',",%,&,<,>,#,†]는 입력할 수 없습니다.
			strRst3 = strRst3 & "		<colOptUseYn>Y</colOptUseYn>"							'옵션 사용 여부 | Y : 사용함, N : 사용안함
			strRst3 = strRst3 & "	</ProductCustOption>"
		End If
		get11stOptionParam = strRst1 & strRst2 & strRst3
'rw get11stOptionParam
'response.end
	End Function

	Public Function getEbayInfoCdParameter(obj)
		Dim strSQL
		Dim mallinfodiv, mallinfoCd, infoContent, certNum
		strSQL = ""
		strSQL = strSQL & " SELECT TOP 1 isNull(r.certNum, '') as certNum "
		strSQL = strSQL & "	,Case When r.safetyDiv in ('10', '40', '70') THEN 'SafeCert' "
		strSQL = strSQL & "		  When r.safetyDiv in ('20', '50', '80') THEN 'SafeCheck' "
		strSQL = strSQL & " 	  When r.safetyDiv in ('30', '60', '90') THEN 'SupplierCheck' end as safetyStr "
		strSQL = strSQL & " ,convert(date, f.certDate) as certDate, f.modelName " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_safetycert_tenReg as r " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.[dbo].[tbl_safetycert_info] as f on r.itemid = f.itemid " & vbcrlf
		strSQL = strSQL & " WHERE r.itemid='"&FItemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			certNum		= rsget("certNum")
		End If
		rsget.Close

		strSQL = ""
		strSQL = strSQL & " SELECT top 100 M.* , " & vbcrlf
		If certNum = "" Then
			strSQL = strSQL & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN  IC.safetyNum " & vbcrlf
		Else
			If certNum = "x" Then
				certNum = "해당없음"
			End If
			strSQL = strSQL & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN '"& certNum &"' " & vbcrlf
		End If
		strSql = strSql & "		 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00001') THEN '상세정보 별도표기' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='10000') THEN '관련법 및 소비자분쟁해결기준에 따름' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035'  " & vbcrlf
		strSQL = strSQL & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd  " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"'  " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='" & FItemid &"' " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'ebay' and IC.itemid='"&FItemid&"'"
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			mallinfodiv = CInt(rsget("mallinfodiv"))
			Set obj("itemAddtionalInfo")("officialNotice") = jsObject()
				obj("itemAddtionalInfo")("officialNotice")("officialNoticeNo") = mallinfodiv	'#상품정보고시 상품군코드
				Set obj("itemAddtionalInfo")("officialNotice")("details") = jsArray()
			Do until rsget.EOF
				mallinfoCd  = rsget("mallinfoCd")
				infoContent = rsget("infoContent")
					Set obj("itemAddtionalInfo")("officialNotice")("details")(null) = jsObject()
						obj("itemAddtionalInfo")("officialNotice")("details")(null)("officialNoticeItemelementCode") = mallinfoCd	'#상품정보고시 항목코드
						obj("itemAddtionalInfo")("officialNotice")("details")(null)("value") = infoContent							'#상품정보고시 값
						obj("itemAddtionalInfo")("officialNotice")("details")(null)("isExtraMark") = false							'상품정보고시 추가입력여부 | true : 추가입력고시, false : 추가 입력없음
				rsget.MoveNext
			Loop
		End If
		rsget.Close
	End Function

	'기본정보 등록 XML
	Public Function getEbayItemRegParameter(vGubun)
		Dim strRst
		Dim obj
		Set obj = jsObject()
'			Set obj("isSell") = jsObject()
'				obj("isSell")(Chkiif(vGubun="A", "iac", "gmkt")) = false						'#판매상태변경 | true 판매, false 판매중지, 판매중지로 1개월 유지시 삭제됨

			Set obj("itemBasicInfo") = jsObject()
				Set obj("itemBasicInfo")("goodsName") = jsObject()
					obj("itemBasicInfo")("goodsName")("eng") = null								'영문상품명
					obj("itemBasicInfo")("goodsName")("chi") = null								'중문상품명
					obj("itemBasicInfo")("goodsName")("jpn") = null								'일문상품명
					obj("itemBasicInfo")("goodsName")("kor") = ""&getItemNameFormat(1)&""			'#검색용 국문상품명
					obj("itemBasicInfo")("goodsName")("promotion") = null						'프로모션용 국문상품명
				Set obj("itemBasicInfo")("category") = jsObject()
					Set obj("itemBasicInfo")("category")("site") = jsArray()
						Set obj("itemBasicInfo")("category")("site")(null) = jsObject()
							obj("itemBasicInfo")("category")("site")(null)("siteType") = Chkiif(vGubun="A", "1", "2")	'#G마켓/옥션 카테고리 등록을 위한 사이트 선택 | 1 옥션, 2 G마켓
							obj("itemBasicInfo")("category")("site")(null)("catCode") = ""& FCateCode &""				'#G마켓/옥션에서 제공하는 최하위(Leaf)카테고리 코드 등록
					Set obj("itemBasicInfo")("category")("shop") = jsArray()
						obj("itemBasicInfo")("category")("shop") = null
'					Set obj("itemBasicInfo")("category")("shop") = jsArray()
'						Set obj("itemBasicInfo")("category")("shop")(null) = jsObject()
'							obj("itemBasicInfo")("category")("shop")(null)("siteType") = ""			'미니샵 카테고리 사이트 구분
'							obj("itemBasicInfo")("category")("shop")(null)("largeCatCode") = ""		'미니샵 대카테고리코드
'							obj("itemBasicInfo")("category")("shop")(null)("middleCatCode") = ""	'미니샵 대카테고리코드
'							obj("itemBasicInfo")("category")("shop")(null)("smallCatCode") = ""		'미니샵 대카테고리코드
					Set obj("itemBasicInfo")("category")("esm") = jsObject()
						obj("itemBasicInfo")("category")("esm")("catCode") = ""& FSDCategoryCode &""	'#ESM카테고리코드등록

 				If FIsbn13 <> "" Then
					Set obj("itemBasicInfo")("book") = jsObject()
						obj("itemBasicInfo")("book")("isUseIsbnCode") = true						'(도서상품명)ISBN코드 사용여부
						obj("itemBasicInfo")("book")("isbnCode") = ""&FIsbn13&""					'(도서상품명)ISBN코드
						obj("itemBasicInfo")("book")("price") = null								'(도서상품명)참고가격
						obj("itemBasicInfo")("book")("attributeCode") = null						'(도서상품명/G마켓전용)추가등록 카테고리
				End If
					Set obj("itemBasicInfo")("catalog") = jsObject()
						obj("itemBasicInfo")("catalog")("modelName") = null							'모델명
						obj("itemBasicInfo")("catalog")("brandNo") = 0								'브랜드코드
						obj("itemBasicInfo")("catalog")("barCode") = null							'바코드
						Set obj("itemBasicInfo")("catalog")("epinCode") = jsArray()
							obj("itemBasicInfo")("catalog")("epinCode")(null) = 0					'ESM 상품분류코드 | 현재 API 제공하지 않아 null로 호출

			Set obj("itemAddtionalInfo") = jsObject()
				Set obj("itemAddtionalInfo")("buyableQuantity") = jsObject()
					obj("itemAddtionalInfo")("buyableQuantity")("type") = 0							'#구매수량제한 타입 | 0 : 구매수량제한없음, 1 : 1회당 최대 구매수량, 2 : ID당 최대 구매수량, 3 : 기간당 최대 구매수량
					obj("itemAddtionalInfo")("buyableQuantity")("qty") = null						'최대구매수량 | 구매수량제한 타입이 1~3일 때 필수
					obj("itemAddtionalInfo")("buyableQuantity")("unitDate") = null					'제한기간 | 구매수량제한 타입이 3일 때 필수
				Set obj("itemAddtionalInfo")("price") = jsObject()									'#옥션/G마켓 판매가격 | 10원단위로 등록
					obj("itemAddtionalInfo")("price")(Chkiif(vGubun="A", "Iac", "Gmkt")) = Clng(GetRaiseValue(MustPrice/10)*10)
				Set obj("itemAddtionalInfo")("stock") = jsObject()									'#옥션/G마켓 재고수량 | 1~99999까지 입력가능, 옵션등록시 옵션재고관리(true)로 선택할 경우 본판매수량은 입력해도 무시되고 옵션의 합산재고로 산정됨
					obj("itemAddtionalInfo")("stock")(Chkiif(vGubun="A", "Iac", "Gmkt")) = getLimitEbayEa()
				Set obj("itemAddtionalInfo")("sellingPeriod") = jsObject()									'#옥션/G마켓 재고수량 | 1~99999까지 입력가능, 옵션등록시 옵션재고관리(true)로 선택할 경우 본판매수량은 입력해도 무시되고 옵션의 합산재고로 산정됨
					obj("itemAddtionalInfo")("sellingPeriod")(Chkiif(vGubun="A", "Iac", "Gmkt")) = 90
					obj("itemAddtionalInfo")("managedCode") = ""& FItemid &""						'#판매자 상품코드
				Set obj("itemAddtionalInfo")("recommendedOpts") = jsObject()
					obj("itemAddtionalInfo")("recommendedOpts")("type") = 0							'#추천옵션 사용여부 | 0 옵션미사용, 1 선택형(최대20개), 2 2개조합형
					obj("itemAddtionalInfo")("recommendedOpts")("isStockManage") = false			'옵션재고관리 | 추천옵션 사용일 시 필수
					obj("itemAddtionalInfo")("recommendedOpts")("independent") = null				'(선택형/조합형 관련)
					obj("itemAddtionalInfo")("recommendedOpts")("combination") = null				'(선택형/조합형 관련)
					obj("itemAddtionalInfo")("inventoryCode") = null								'(G마켓용)G마켓 인벤토리 코드
				Set obj("itemAddtionalInfo")("sellerShop") = jsObject()
					obj("itemAddtionalInfo")("sellerShop")("catCode") = FtenCateLarge & FtenCateMid & FtenCateSmall	'#판매자 카테고리코드
					obj("itemAddtionalInfo")("sellerShop")("catName") = FtenCateSmall				'#판매자 카테고리명
					obj("itemAddtionalInfo")("sellerShop")("brandCode") = FMakerId					'#판매자 브랜드코드
					obj("itemAddtionalInfo")("sellerShop")("brandName") = FMakerName				'#판매자 브랜드명
					obj("itemAddtionalInfo")("expiryDate") = null									'유효일
					obj("itemAddtionalInfo")("manufacturedDate") = null								'제조일
				Set obj("itemAddtionalInfo")("origin") = jsObject()
					obj("itemAddtionalInfo")("origin")("goodsType") = 1								'#원산지상품 타입 | 0 원산지표시대상아님(식품이외), 1 상세설명참조, 2 가공품, 3 농산물, 4 수산물
					obj("itemAddtionalInfo")("origin")("type") = 5									'#원산지지역 타입 | 0 없음, 1 국내산, 2 수입산, 5 기타 | 상세설명참조일경우 0~5 다 가능, 단 상품조회시 상세설명참조는 0으로 내려감
					obj("itemAddtionalInfo")("origin")("code") = null								'원산지지역 코드
					obj("itemAddtionalInfo")("origin")("isMultipleOrigin") = false					'#복수원산지 여부 | true 복수원산지 상품, false 단일원산지 상품
					obj("itemAddtionalInfo")("capacity") = null
'				Set obj("itemAddtionalInfo")("capacity") = jsObject()
'					obj("itemAddtionalInfo")("capacity")("vol") = null								'(옥션상품용)용량/규격 값
'					obj("itemAddtionalInfo")("capacity")("unit") = null								'(옥션상품용)용량/규격 단위

				Set obj("itemAddtionalInfo")("shipping") = jsObject()
					obj("itemAddtionalInfo")("shipping")("type") = 1								'#배송방식 타입 입력 | G마켓은 무조건 1번만 사용가능 / 옥션 3번 선택시 일반우편, 퀵서비스 방문수령중 선택 필요 | 1 택배소포, 2 화물배달, 3 판매자직접배송
					obj("itemAddtionalInfo")("shipping")("companyNo") = 10013						'#택배사코드 | 10013 CJ대한통운
					Set obj("itemAddtionalInfo")("shipping")("policy") = jsObject()
						obj("itemAddtionalInfo")("shipping")("policy")("placeNo") = 210824			'#출하지번호
						obj("itemAddtionalInfo")("shipping")("policy")("feeType") = 1				'#배송비 타입
						Set obj("itemAddtionalInfo")("shipping")("policy")("bundle") = jsObject()
							obj("itemAddtionalInfo")("shipping")("policy")("bundle")("deliveryTmplId") = 2356837 '#묶음배송비정책번호
						Set obj("itemAddtionalInfo")("shipping")("policy")("each") = jsObject()
							obj("itemAddtionalInfo")("shipping")("policy")("each")("feeType") = 0		'상품별배송비 타입 | 현재 제공하지 않아 무조건 0번으로 입력 | 0 묶음배송비사용, 1 무료, 2 유료, 3 조건부무료, 4 수량별차등
							obj("itemAddtionalInfo")("shipping")("policy")("each")("feePayType") = 0	'상품별배송비지불방법 | 착불인지 선결제여부인지 입력(추후제공)
							obj("itemAddtionalInfo")("shipping")("policy")("each")("fee") = 0			'상품별배송비금액 | 무료인 경우 0 입력(추후제공)
							obj("itemAddtionalInfo")("shipping")("policy")("each")("baseFee") = 0		'상품별배송비조건부 금액 (추후제공)
					Set obj("itemAddtionalInfo")("shipping")("returnAndExchange") = jsObject()
						obj("itemAddtionalInfo")("shipping")("returnAndExchange")("addrNo") = 490970			'(반품주소) 판매자주소번호
						obj("itemAddtionalInfo")("shipping")("returnAndExchange")("shippingCompany") = "0008"	'반품교환택배사코드
						obj("itemAddtionalInfo")("shipping")("returnAndExchange")("fee") = 2500					'반품/교환 편도배송비
					Set obj("itemAddtionalInfo")("shipping")("dispatchPolicyNo") = jsObject()
					'''''''''''''''''''''''아래 건 기존에 없음 추가해야 함''''''''''''''''''''''''''''''''''''''
					obj("itemAddtionalInfo")("shipping")("dispatchPolicyNo")(Chkiif(vGubun="A", "Iac", "Gmkt")) = Chkiif(vGubun="A", 587470, 587465)	'#(옥션/G마켓) 발송타입정책번호
					obj("itemAddtionalInfo")("shipping")("generalPost") = null						'#(옥션용)일반우편 제공 여부 및 요금관련
					obj("itemAddtionalInfo")("shipping")("visitAndTake") = null						'#방문수령 제공여부
					obj("itemAddtionalInfo")("shipping")("quickService") = null						'#퀵서비스 제공여부
				Call getEbayInfoCdParameter(obj)	'#상품정보고시 관련
				obj("itemAddtionalInfo")("isAdultProduct") = Chkiif(IsAdultItem()="Y", true, false)			'#성인상품여부 | true : 성인상품, false : 일반상품
				obj("itemAddtionalInfo")("isYouthNotAvailable") = Chkiif(IsAdultItem()="Y", true, false)	'#청소년구매불가여부 | 상품이미지가 노출여부 | true : 청소년구매불가상품, false : 일반상품
				obj("itemAddtionalInfo")("isVatFree") = Chkiif(FVatInclude="N", true, false)				'#부과세 여부 | true : 면세상품, false : 과세상품
				Call getEbayCertInfoParameter(obj)	'#상품정보고시 관련
				Call getEbayImageParameter(obj)		'#상품이미지 관련
				obj("itemAddtionalInfo")("weight") = 0												'(G마켓용) 상품무게(단위:kg)
				Call getEbayContParamToReg(obj, vGubun)
				obj("itemAddtionalInfo")("addonService") = null										'추가구성 관련
			Set obj("addtionalInfo") = jsObject()
				Set obj("addtionalInfo")("sellerDiscount") = jsObject()
					obj("addtionalInfo")("sellerDiscount")("isUse") = false							'#판매자할인 사용여부 | true 할인적용, false 할인미적용
					Set obj("addtionalInfo")("sellerDiscount")(Chkiif(vGubun="A", "iac", "gmkt")) = jsObject()
						obj("addtionalInfo")("sellerDiscount")(Chkiif(vGubun="A", "iac", "gmkt"))("type") = 0 '할인타입 | 판매자할인 사용여부 true일경우 필수 0 사용안함, 1 정액, 2 정률
				Set obj("addtionalInfo")("siteDiscount") = jsObject()
					obj("addtionalInfo")("siteDiscount")(Chkiif(vGubun="A", "iac", "gmkt")) = true		'#G마켓/옥션에서 부담하는 사이트 할인을 적용할지 여부 | true 적용, false 미적용
					obj("addtionalInfo")("gift") = null
					Set obj("addtionalInfo")("pcs") = jsObject()
						obj("addtionalInfo")("pcs")("isUse") = true									'#가격비교사이트 노출여부 | true 등록(노출됨), false 등록하지않음(미노출)
						If vGubun="A" Then
						obj("addtionalInfo")("pcs")("isUseIacPcsCoupon") = false					'#(옥션용)가격비교사이트 쿠폰적용여부
						Else
						obj("addtionalInfo")("pcs")("isUseGmkPcsCoupon") = false					'#(G마켓용)가격비교사이트 쿠폰적용여부 | G마켓은 한번 설정하면 변경불가(추후제공예정)
						End If
					Set obj("addtionalInfo")("overseaSales") = jsObject()
						obj("addtionalInfo")("overseaSales")("isAgree") = false						'#(G마켓용)해외판매여부 | true 진행, false 진행안함

'		response.write obj.jsString
'		response.end
		getEbayItemRegParameter = obj.jsString
	End Function
End Class

Class CEbay
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
            addSql = addSql & " WHERE (optCnt-optNotSellCnt < 1) "
			addSql = addSql & " OR (optAddCNT > 0) "
            addSql = addSql & " )"
		End If

		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum "
		strSql = strSql & "	, isNULL(R.auctionStatCD,-9) as auctionStatCD "
		strSql = strSql & "	, UC.socname_kor, am.SDCategoryCode, am.cateCode "
		strSql = strSql & "	, isNull(c.isbn13, '') as isbn13 "
		strSql = strSql & "	, CONVERT(VARCHAR(10), isNull(sellSTDate, getdate()), 23) as sellSTDate "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_ebay_cate_mapping "
		strSql = strSql & "		WHERE gubun = 'A' "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ebay_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small and gubun='A' "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_auction1010_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		ELSE
		    strSql = strSql & " and (i.deliveryType<>9)"
	    END IF
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.itemdiv <> '21' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "						'플라워/화물배송
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'등록제외 카테고리
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_etcmall.dbo.tbl_auction1010_regItem WHERE auctionStatCD >= 3) "	''등록완료이상은 등록안됨.										'롯데등록상품 제외
		strSql = strSql & " and cm.mapCnt is Not Null "
		strSql = strSql & "		"	& addSql											'카테고리 매칭 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CEbayItem
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

				FOneItem.FSDCategoryCode	= rsget("SDCategoryCode")
				FOneItem.FcateCode			= rsget("cateCode")
				FOneItem.FbasicimageNm 		= rsget("basicimage")

				FOneItem.FIsbn13 			= rsget("isbn13")
'				FOneItem.FSellSTDate		= rsget("sellSTDate")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub

	Public Sub get11stEditOneItem
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
		strSql = strSql & "	, m.st11GoodNo, m.st11price, m.st11SellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, UC.socname_kor, am.cateCode, isNULL(m.st11StatCD,-9) as st11StatCD, tm.safeDiv, tm.isNeed, tm.depth1Code "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv = '21' "
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"

		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "

		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_11st_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_category as tm on am.depthCode = tm.depthCode "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.st11GoodNo is Not Null "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new C11stItem
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
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSt11GoodNo		= rsget("st11GoodNo")
				FOneItem.FSt11price			= rsget("st11price")
				FOneItem.FSt11SellYn		= rsget("st11SellYn")

				FOneItem.FMakerName			= db2html(rsget("makername"))
				FOneItem.FBrandName			= db2html(rsget("brandname"))
				FOneItem.FBrandNameKor		= db2html(rsget("socname_kor"))
				If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
					FOneItem.FMakerName		= FOneItem.FBrandName
				End If

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
	            FOneItem.Fregitemname   	= rsget("regitemname")
                FOneItem.FregImageName		= rsget("regImageName")
                FOneItem.FbasicImageNm		= rsget("basicimage")
                FOneItem.FDepthCode			= rsget("depthCode")
                FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.FSt11StatCD		= rsget("st11StatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")

				FOneItem.FSafeDiv 			= rsget("safeDiv")
				FOneItem.FIsNeed 			= rsget("isNeed")
				FOneItem.FDepth1Code 		= rsget("depth1Code")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub

End Class

'11번가 상품코드 얻기
Function get11stGoodno(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 st11goodno FROM db_etcmall.dbo.tbl_11st_regitem WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		get11stGoodno = rsget("st11goodno")
	End If
	rsget.Close
End Function

'11번가 상품코드/상품가 얻기
Function get11stGoodno2(iitemid, ist11goodno, byref MustPrice)

	Dim strRst, strSql
	Dim sellcash, orgprice, buycash, saleyn, tmpPrice, vdeliverytype, ispecialPrice
	Dim GetTenTenMargin, st11goodno

	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.sellcash, i.buycash, i.orgprice, i.sailyn, r.st11goodno, i.deliverytype, isnull(mi.mustPrice, 0) as specialPrice "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_etcmall.dbo.tbl_11st_regitem as r on i.itemid = r.itemid "
	strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_outmall_mustPriceItem] as mi "
	strSql = strSql & " 	on i.itemid = mi.itemid "
	strSql = strSql & " 	and mi.mallgubun = '11st1010' "
	strSql = strSql & " 	and (GETDATE() >= mi.startDate and GETDATE() <= mi.endDate ) "
	strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		sellcash	= rsget("sellcash")
		orgprice	= rsget("orgprice")
		buycash		= rsget("buycash")
		saleyn		= rsget("sailyn")
		st11goodno	= rsget("st11goodno")
		vdeliverytype = rsget("deliverytype")
		ispecialPrice = rsget("specialPrice")
	Else
		get11stGoodno2 = ""
		Exit Function
		response.end
	End If
	rsget.close

	If ispecialPrice <> "0" Then
		tmpPrice = ispecialPrice
	Else
		GetTenTenMargin = CLng((10000 - buycash / sellcash * 100 * 100) / 100)
	'	If (vdeliverytype = 2) OR (vdeliverytype = 9) Then
	'		If (GetTenTenMargin < CMAXMARGIN) OR (saleyn = "Y" AND sellcash < 10000) Then
	'			tmpPrice = orgprice
	'		Else
	'			tmpPrice = sellcash
	'		End If
	'	Else
			If (GetTenTenMargin < CMAXMARGIN) Then
				tmpPrice = orgprice
			Else
				tmpPrice = sellcash
			End If
	'	End If
	End If
	MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
	ist11goodno = st11goodno
End Function

'11번가 상품코드, 옵션 수 얻기
Function get11stGoodno3(iitemid, ist11goodno, byref opCnt)
	Dim strSql, st11goodno, optioncnt

	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.optioncnt, r.st11goodno "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_etcmall.dbo.tbl_11st_regitem as r on i.itemid = r.itemid "
	strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		opCnt		= rsget("optioncnt")
		ist11goodno	= rsget("st11goodno")
	Else
		get11stGoodno3 = ""
		Exit Function
		response.end
	End If
	rsget.close
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
    v = replace(v, """", "&quot;")
	'v = replace(v, "&", "&amp;")

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
