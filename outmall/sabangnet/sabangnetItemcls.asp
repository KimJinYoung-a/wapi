<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "sabangnet"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST sabangnetAPIURL = "http://r.sabangnet.co.kr"
CONST sabangnetID = "tenbyten"
CONST sabangnetAPIKEY = "PTxNV3d9CXPXBNu60X72EbSNYTJd5955b"
CONST CDEFALUT_STOCK = 999
CONST wapiURL = "http://wapi.10x10.co.kr"

Class CSabangnetItem
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
	Public FIcon1Image
	Public FListimage
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public FSafetyNum
	Public FSafetydiv
	Public Fitemcontent
	Public FSabangnetStatCD
	Public Fdeliverfixday
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FinfoDiv
	Public Fsafetyyn
	Public FMaySoldOut
	Public Fregitemname
	Public FregImageName
	Public FbasicImageNm
	Public FItemsize
	Public FItemsource
	Public FBrandCode
	Public Fsocname_kor
	Public FDepthCode
	Public FDepth4Code
	Public FSabangnetGoodNo
	Public FSabangnetprice
	Public FSabangnetSellYn
	Public FMayLimitSoldout
	Public FMwdiv

	Function RightCommaDel(ostr)
		Dim restr
		restr = ""
		If IsNULL(ostr) Then Exit Function
		restr = Trim(ostr)
		If (Right(restr,1)=",") Then restr = Left(restr,Len(restr)-1)
		RightCommaDel = restr
	End Function

	Public Function IsFreeBeasong()
		IsFreeBeasong = False
		If (FdeliveryType=2) or (FdeliveryType=4) or (FdeliveryType=5) then				'2(텐무), 4,5(업무)
			IsFreeBeasong = True
		End If

		If (FdeliveryType=9) Then														'업체조건
			IsFreeBeasong = False
		End If
		If (MustPrice >= 50000) Then IsFreeBeasong = True
    End Function

	'// 품절여부
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold <= CMAXLIMITSELL))
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
		Dim GetTenTenMargin, sqlStr, specialPrice, tmpPrice, vBigPrice, vSmallPrice, ownItemCnt
		specialPrice = 0
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

		If specialPrice <> 0 Then
			MustPrice = specialPrice
		ElseIf ownItemCnt > 0 Then
			MustPrice = Forgprice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If GetTenTenMargin < CMAXMARGIN Then
				tmpPrice = Forgprice
			Else
				tmpPrice = FSellCash
			End If
			MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
		End If
	End Function

	'// 사방넷 판매여부 반환
	Public Function getSabangnetSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getSabangnetSellYn = "Y"
			Else
				getSabangnetSellYn = "N"
			End If
		Else
			getSabangnetSellYn = "N"
		End If
	End Function

	Public Function getSourcearea()
		Dim arrAreaName, i
		arrAreaName = Array("America", "Australia", "Belgium", "Brazil", "Chile", "CHINA", "ITALY", "KOREA", "Mexico", "Norway", "Thailand", "과테말라", "국내산", "그루지야(조지아)", "그리스", "기타국가", "나이지리아", "남아프리카공화국", "네덜란드", "네팔", "노르웨이", "뉴질랜드", "니카라과", "대만", "대한민국", "덴마크", "도미니카", "도미니카공화국", "독일", "라오스", "라트비아", "러시아", "러시아공화국", "레바논", "루마니아", "리비아", "리투아니아", "마다가스카르", "마카오", "말레이시아", "말레이지아", "멕시코", "모로코", "모리셔스", "몰다비아", "몰도바", "몰타", "몽골", "미국", "미국/일본", "미국OEM", "미얀마", "바레인", "방글라데시", "베네수엘라", "베트남", "벨기에", "보스니아", "복합원산지", "볼리비아", "북한", "불가리아", "브라질", "사우디아라비아", "상세설명참조", "세네갈", "세르비아", "수입산", "스리랑카", "스웨덴", "스위스", "스코틀랜드", "스페인", "슬로바키아", "슬로베니아", "싱가포르", "아랍에미레이트", "아랍에미리트", "아르메니아", "아르헨티나", "아이티", "아일랜드", "아프리카", "알바니아", "에스토니아", "에콰도르", "엘살바도르", "영국", "오스트레일리아", "오스트리아", "온두라스", "외국산", "요르단", "우간다", "우루과이", "우즈베키스탄", "우크라이나", "원양산", "유럽연합(EU)", "이디오피아", "이라크", "이란", "이스라엘", "이집트", "이탈리아", "이태리", "인도", "인도네시아", "인도네시아OEM", "인디아", "일본", "일본/태국", "중국", "중국/대만", "중국/말레이시아", "중국/미얀마", "중국/베트남", "중국/인도", "중국/인도네시아", "중국/태국", "중국/필리핀", "중국OEM", "중국개성", "중국외복수국가", "지부티", "체코", "칠레", "캄보디아", "캐나다", "케냐", "켈리포니아", "콜롬비아", "쿠웨이트", "크로아티아", "타이완", "타일랜드", "태국", "터키", "튀니지", "파라과이", "파키스탄", "페루", "포르투갈", "폴란드", "프랑스", "프랑스/미국", "프랑스/중국", "핀란드", "필리핀", "한국/중국", "한국/중국/미국", "헝가리", "호주", "홍콩")

		For i =0 To Ubound(arrAreaName)
			If Trim(arrAreaName(i)) = Trim(FSourcearea) Then
				getSourcearea = Trim(arrAreaName(i))
				Exit For
			End If
		Next

		If FSourcearea = "한국" Then
			getSourcearea = "대한민국"
		End If

		If getSourcearea = "" Then
			getSourcearea = "기타국가"
		End If
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
				End If
				rsget.Close
			End If
		Else
			getiszeroWonSoldOut = "N"
		End If
	End Function

	Public Function getSabangnetContParamToReg()
		Dim strRst, strSQL, infoContRst
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '돋움','돋움' }</style><br>"
		strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_iPark.jpg'></p><br>"
		If Fitemsize <> "" Then
			strRst = strRst & "- 사이즈 : " & Fitemsize & "<br>"
		End if

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

		strSQL = ""
		strSQL = strSQL & " SELECT c.infoCd, c.infoItemName, "
		strSQL = strSQL & " CASE WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035' "
		strSQL = strSQL & " 	WHEN c.infotype='T' and c.infoItemName = '품질보증기준' THEN '관련법 및 소비자분쟁해결기준에 따름' "
		strSql = strSql & " 	WHEN LEN(isNull(F.infocontent, '')) < 2 THEN '상품 상세 참고' " & vbcrlf
		strSQL = strSQL & " ELSE F.infocontent END AS infocontent "
		strSQL = strSQL & " FROM db_item.dbo.tbl_item_contents IC "
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_infoCode c ON ic.infoDiv=c.infoDiv "
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON F.infoCd = c.infoCd and F.itemid='"& Fitemid &"' "
		strSQL = strSQL & " WHERE IC.itemid='"& Fitemid &"' "
		strSQL = strSQL & " ORDER BY convert(int, c.infoCd) ASC "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			infoContRst = ""
			infoContRst = infoContRst & "<table align=""center"">"
			infoContRst = infoContRst & "	<colgroup>"
			infoContRst = infoContRst & "		<col style=""width: 30%;"">"
			infoContRst = infoContRst & "		<col style=""width: 70%;"">"
			infoContRst = infoContRst & "	</colgroup>"
			infoContRst = infoContRst & "	<tbody>"
			Do until rsget.EOF
				infoContRst = infoContRst & "<tr>"
				infoContRst = infoContRst & "	<td scope=""row"">"&rsget("infoItemName")&"</td>"
				infoContRst = infoContRst & "	<td>"&rsget("infocontent")&"</td>"
				infoContRst = infoContRst & "</tr>"
				rsget.MoveNext
			Loop
			infoContRst = infoContRst & "	</tbody>"
			infoContRst = infoContRst & "</table>"
			strRst = strRst & infoContRst
		End If
		rsget.Close

		'#배송 주의사항
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_sabangnet.jpg"">")
		getSabangnetContParamToReg = strRst
	End Function

	Public Function getSabangnetOptParamtoREG()
		Dim buf, sqlStr, i, tmpVAL, limitYCnt, limitNCnt
		Dim vitemoption, voptionname, voptsellyn, voptaddprice, voptLimit, optStatus
    	buf = ""
    	tmpVAL = ""
    	limitYCnt = 0
    	limitNCnt = 0

		If FOptionCnt > 0 Then
			sqlStr = ""
			sqlStr = sqlStr & " SELECT itemoption, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
			sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item_option "
			sqlStr = sqlStr & " WHERE isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				Do until rsget.EOF
					vitemoption 		= rsget("itemoption")
					voptionname 		= Replace(rsget("optionname"), ",", "_")
					voptsellyn 			= rsget("optsellyn")
					voptaddprice		= rsget("optaddprice")
					voptLimit			= rsget("optLimit")
					voptLimit = voptLimit-5
					If (voptLimit < 1) Then voptLimit = 0
					If (FLimitYN <> "Y") Then voptLimit = CDEFALUT_STOCK

					If ((voptsellyn <> "Y") OR (voptLimit = 0)) Then
						optStatus = "004"			'004 : 품절
					Else
						optStatus = "002"			'002 : 공급중
					End If

					tmpVAL = tmpVAL & voptionname &"^^"& voptLimit &"^^"& voptaddprice &"^^"& vitemoption &"^^EA^^"& optStatus & ","
					If (voptLimit = 0) Then
						limitNCnt = limitNCnt + 1
					Else
						limitYCnt = limitYCnt + 1
					End If
					rsget.MoveNext
				Loop
			End If
			rsget.Close
			tmpVAL = RightCommaDel(tmpVAL)

			If FOptioncnt > 0 Then
				If limitYCnt = 0 Then
					FMayLimitSoldout = "Y"
				Else
					FMayLimitSoldout = "N"
				End If
			End If

			buf = buf & "		<CHAR_1_NM><![CDATA[옵션]]></CHAR_1_NM>"
			buf = buf & "		<CHAR_1_VAL><![CDATA["&tmpVAL&"]]></CHAR_1_VAL>"
		Else
			buf = buf & "		<CHAR_1_NM><![CDATA[단품]]></CHAR_1_NM>"
			If Flimityn = "Y" Then
				voptLimit = FLimitNo - FLimitSold -5
			Else
				voptLimit = CDEFALUT_STOCK
			End If

			If voptLimit < 1 Then
				voptLimit = 0
			End If

			buf = buf & "		<CHAR_1_VAL><![CDATA[단품^^"&voptLimit&"]]></CHAR_1_VAL>"
		End If
		buf = buf & "		<CHAR_2_NM><![CDATA[]]></CHAR_2_NM>"
		buf = buf & "		<CHAR_2_VAL><![CDATA[]]></CHAR_2_VAL>"
		getSabangnetOptParamtoREG = buf
	End Function

	Public Function getSabangnetAddImageParam()
		Dim strRst, strSQL, i
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If
		strRst = ""
		strRst = strRst & "		<IMG_PATH><![CDATA["&FbasicImage&"]]></IMG_PATH>"					'#대표이미지 | 예 : http://gs4333.CO.KR/product_image/a0000769/200907/image20_700.jpg
		strRst = strRst & "		<IMG_PATH1><![CDATA["&FbasicImage&"]]></IMG_PATH1>"					'#종합몰(JPG)이미지 | 예 : http://gs4333.CO.KR/product_image/a0000769/200907/image20_700.jpg  (종합몰(JPG)이미지 (500x500 ~ 700x700))
		strRst = strRst & "		<IMG_PATH2><![CDATA[]]></IMG_PATH2>"								'부가이미지2
		strRst = strRst & "		<IMG_PATH3><![CDATA["&FIcon1Image&"]]></IMG_PATH3>"					'#부가이미지3 | 예 : http://gs4333.CO.KR/product_image/a0000769/200907/image20_700.jpg  (11번가목록이미지 (300*300))
		strRst = strRst & "		<IMG_PATH4><![CDATA[]]></IMG_PATH4>"								'부가이미지4
		strRst = strRst & "		<IMG_PATH5><![CDATA[]]></IMG_PATH5>"								'부가이미지5
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					If (isNull(rsget("addimage_600")) OR rsget("addimage_600") = "") Then
						strRst = strRst & "		<IMG_PATH"&i+5&"><![CDATA[http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")&"]]></IMG_PATH"&i+5&">"	'부가이미지 6~10 | 쇼핑몰 추가이미지(1~5)
					Else
						strRst = strRst & "		<IMG_PATH"&i+5&"><![CDATA[http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "_600/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_600")&"]]></IMG_PATH"&i+5&">"	'부가이미지 6~10 | 쇼핑몰 추가이미지(1~5)
					End If
				End If
				rsget.MoveNext
				If i>=5 Then Exit For
			Next
		End If
		rsget.Close
		getSabangnetAddImageParam = strRst
	End Function

	Public Function getSabangnetCertInfoToReg
		Dim buf, strSql, safetyDiv, safetyId, certNum, certOrganName, certmakerName, isRegCert, certDiv
		strSql = ""
		strSql = strSql & " select top 1 i.itemid, t.safetyDiv "
		strSql = strSql & " ,Case When t.safetyDiv = '10' THEN '전기용품_안전인증' "
		strSql = strSql & " 	When t.safetyDiv = '20' THEN '전기용품_안전확인신고' "
		strSql = strSql & " 	When t.safetyDiv = '30' THEN '전기용품_공급자적합성확인' "
		strSql = strSql & " 	When t.safetyDiv = '40' THEN '생활제품_안전인증' "
		strSql = strSql & " 	When t.safetyDiv = '50' THEN '생활제품_안전확인신고' "
		strSql = strSql & " 	When t.safetyDiv = '60' THEN '생활제품_공급자적합성확인' "
		strSql = strSql & " 	When t.safetyDiv = '70' THEN '어린이제품_안전인증' "
		strSql = strSql & " 	When t.safetyDiv = '80' THEN '어린이제품_안전확인신고' "
		strSql = strSql & " 	When t.safetyDiv = '90' THEN '어린이제품_공급자적합성확인' end as safetyId "
		strSql = strSql & " , t.certNum, f.certOrganName, f.makerName, f.certDiv "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
		strSql = strSql & " WHERE i.itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			safetyDiv		= rsget("safetyDiv")
			safetyId		= rsget("safetyId")
			certNum			= rsget("certNum")
			certOrganName	= rsget("certOrganName")
			certmakerName	= rsget("makerName")
			certDiv			= rsget("certDiv")
			isRegCert		= "Y"
		Else
			isRegCert		= "N"
		End If
		rsget.Close

		buf = ""
		buf = buf & "		<CERTNO><![CDATA["&certNum&"]]></CERTNO>"								'인증번호 | 전기용품, 유아 안전용품 등 안전검사를 거쳐야 하는 상품의 경우 해당기관에서 부여한 인증번호를 입력합니다
		buf = buf & "		<AVLST_DM></AVLST_DM>"													'인증유효 시작일 | 숫자8자리 입력하세요 예:20100401
		buf = buf & "		<AVLED_DM></AVLED_DM>"													'인증유효 마지막일 | 숫자8자리 입력하세요 예:20100401
		buf = buf & "		<ISSUEDATE></ISSUEDATE>"												'발급일자 | 숫자8자리 입력하세요 예:20100401
		buf = buf & "		<CERTDATE></CERTDATE>"													'인증일자 | 숫자8자리 입력하세요 예:20100401
		buf = buf & "		<CERT_AGENCY><![CDATA["&certOrganName&"]]></CERT_AGENCY>"				'인증기관 | 예 : 한국기업인증연구원
		buf = buf & "		<CERTFIELD><![CDATA["&certDiv&"]]></CERTFIELD>"							'인증분야 | 예 : 규격인증
		getSabangnetCertInfoToReg = buf
	End Function

	' Public Function getSabangnetCertInfoToReg
	' 	Dim buf, safetydivName, certNo, ssgCERTFIELD
	' 	If (FSafetyyn = "Y") and (Trim(FSafetyNum) <> "") Then
	' 		certNo = Trim(FSafetyNum)
	' 		Select Case FSafetydiv
	' 			Case "10"
	' 				safetydivName = "국가통합인증(KC마크)"
	' 				ssgCERTFIELD = "안전인증_안전인증"
	' 			Case "20"
	' 				safetydivName = "전기용품 안전인증"
	' 				ssgCERTFIELD = "전파인증"
	' 			Case "30"
	' 				safetydivName = "KPS 안전인증 표시"
	' 				ssgCERTFIELD = "안전인증_안전인증"
	' 			Case "40"
	' 				safetydivName = "KPS 자율안전 확인 표시"
	' 				ssgCERTFIELD = "안전인증_안전확인"
	' 			Case "50"
	' 				safetydivName = "KPS 어린이 보호포장 표시"
	' 				ssgCERTFIELD = "안전인증_안전인증"
	' 		End Select
	' 		'##SSG에 입력해야 되는 네이밍
	' 		'안전인증_안전인증
	' 		'안전인증_안전확인
	' 		'안전인증_공급자적합성
	' 		'전파인증
	' 		'위해우려
	' 	End If

	' 	If ssgCERTFIELD = "" Then
	' 		ssgCERTFIELD = "어린이제품^^안전인증"
	' 		certNo = "없음^^없음"
	' 	End If

	' 	buf = ""
	' 	buf = buf & "		<CERTNO><![CDATA["&certNo&"]]></CERTNO>"								'인증번호 | 전기용품, 유아 안전용품 등 안전검사를 거쳐야 하는 상품의 경우 해당기관에서 부여한 인증번호를 입력합니다
	' 	buf = buf & "		<AVLST_DM></AVLST_DM>"													'인증유효 시작일 | 숫자8자리 입력하세요 예:20100401
	' 	buf = buf & "		<AVLED_DM></AVLED_DM>"													'인증유효 마지막일 | 숫자8자리 입력하세요 예:20100401
	' 	buf = buf & "		<ISSUEDATE></ISSUEDATE>"												'발급일자 | 숫자8자리 입력하세요 예:20100401
	' 	buf = buf & "		<CERTDATE></CERTDATE>"													'인증일자 | 숫자8자리 입력하세요 예:20100401
	' 	buf = buf & "		<CERT_AGENCY><![CDATA[]]></CERT_AGENCY>"								'인증기관 | 예 : 한국기업인증연구원
	' 	buf = buf & "		<CERTFIELD><![CDATA["&ssgCERTFIELD&"]]></CERTFIELD>"					'인증분야 | 예 : 규격인증
	' 	getSabangnetCertInfoToReg = buf
	' End Function

	Public Function getSabangnetItemInfoCdToReg
		Dim strSql, buf, lp
		Dim mallinfoCd, infoContent, rsMallinfoDiv
		strSql = ""
		strSql = strSql & " SELECT TOP 100 M.* , "
		strSql = strSql & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN 'Y' "
		strSql = strSql & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'N') THEN 'N' "
		strSql = strSql & " 	 WHEN (M.infoCd='00001') AND (IC.safetyyn= 'Y') THEN IC.safetyNum "
		strSql = strSql & " 	 WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035'  "
		strSql = strSql & " 	 WHEN LEN(isNull(F.infocontent, '')) < 2 THEN '상품 상세 참고' " & vbcrlf
		strSql = strSql & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent "
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M  "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv  "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd  "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"'  "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='"&FItemID&"' "
		strSql = strSql & " WHERE M.mallid = 'sabangnet' and IC.itemid='"&FItemID&"'"
		strSql = strSql & " ORDER BY convert(int, mallinfocd) ASC "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			rsMallinfoDiv = rsget("mallinfoDiv")
			If rsMallinfoDiv = "47" Then
				rsMallinfoDiv = "36"
			ElseIf rsMallinfoDiv = "" Then
				rsMallinfoDiv = "37"
			End If

			buf = ""
			buf = buf & "		<PROP_EDIT_YN>Y</PROP_EDIT_YN>"										'속성수정여부 | "속성정보 수정여부를 Y or N로 입력합니다. Y입력시 속성정보(속성분류코드, 속성값)를 수정 처리합니다."
			buf = buf & "		<PROP1_CD>0"& rsMallinfoDiv &"</PROP1_CD>"					'속성분류코드 | "속성분류코드를 숫자 3자리 형식으로 입력합니다. 속성분류코드는 상품속성코드 조회 API나 사방넷 상품관리화면의 속성분류표를 참고하시기 바랍니다. 예: 의류는 001을 입력합니다."
			Do until rsget.EOF
				infoContent = rsget("infocontent")
				mallinfocd = rsget("mallinfocd")
				buf = buf & "		<PROP_VAL"&mallinfocd&"><![CDATA["&infoContent&"]]></PROP_VAL"&mallinfocd&">"	'속성값 | "속성분류코드에 따른 속성순번1인 속성명에 해당하는 속성값을 입력합니다. 속성값(1 ~ 20)은  입력순서대로 처리되므로, 속성순번에 주의하시기 바랍니다.(속성값이 없을 경우, 공란으로 처리하시기 바랍니다.) 예 : 의류 001의 속성명1은 제품 소재이며, 속성값1에 면,나일론 등을 입력합니다."
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		If (session("ssBctID")="kjy8517") and FItemId = "1882712" Then
			buf = ""
			buf = buf & "		<PROP_EDIT_YN>Y</PROP_EDIT_YN>"				'속성수정여부 | "속성정보 수정여부를 Y or N로 입력합니다. Y입력시 속성정보(속성분류코드, 속성값)를 수정 처리합니다."
			buf = buf & "		<PROP1_CD>008</PROP1_CD>"					'속성분류코드 | "속성분류코드를 숫자 3자리 형식으로 입력합니다. 속성분류코드는 상품속성코드 조회 API나 사방넷 상품관리화면의 속성분류표를 참고하시기 바랍니다. 예: 의류는 001을 입력합니다."
			buf = buf & "		<PROP_VAL1><![CDATA[상세페이지 참조]]></PROP_VAL1>"
			buf = buf & "		<PROP_VAL2><![CDATA[상세페이지 참조]]></PROP_VAL2>"
			buf = buf & "		<PROP_VAL3><![CDATA[상세페이지 참조]]></PROP_VAL3>"
			buf = buf & "		<PROP_VAL4><![CDATA[상세페이지 참조]]></PROP_VAL4>"
			buf = buf & "		<PROP_VAL5><![CDATA[상세페이지 참조]]></PROP_VAL5>"
			buf = buf & "		<PROP_VAL6><![CDATA[상세페이지 참조]]></PROP_VAL6>"
			buf = buf & "		<PROP_VAL7><![CDATA[상세페이지 참조]]></PROP_VAL7>"
			buf = buf & "		<PROP_VAL8><![CDATA[상세페이지 참조]]></PROP_VAL8>"
			buf = buf & "		<PROP_VAL9><![CDATA[상세페이지 참조]]></PROP_VAL9>"
			buf = buf & "		<PROP_VAL10><![CDATA[상세페이지 참조]]></PROP_VAL10>"
			buf = buf & "		<PROP_VAL11><![CDATA[상세페이지 참조]]></PROP_VAL11>"
			buf = buf & "		<PROP_VAL12><![CDATA[상세페이지 참조]]></PROP_VAL12>"
			buf = buf & "		<PROP_VAL13><![CDATA[상세페이지 참조]]></PROP_VAL13>"
			buf = buf & "		<PROP_VAL14><![CDATA[상세페이지 참조]]></PROP_VAL14>"
			buf = buf & "		<PROP_VAL15><![CDATA[상세페이지 참조]]></PROP_VAL15>"
		End If
		getSabangnetItemInfoCdToReg = buf
	End Function

	'상품 등록 XML
	Public Function getSabangnetItemRegParameter(isReg, ichgSellyn)
		Dim strRst, tmpStatus, vMwdiv
		If isReg = False Then
			'4(완전품절)로 보냈을 때 만약 연결된 쇼핑몰이 있다면 몰에 따라 삭제가 될 수 있다함
			'3(일시중지)으로 보내면 자동화 서비스 사용시 판매/품절을 입맛에 맞게 변경가능 답변 받음
			tmpStatus = 2
		Else
			If ichgSellyn = "Y" Then
				tmpStatus = 2
			Else
				tmpStatus = 3
			End If
		End If

		Select Case FMwdiv
			Case "M"	vMwdiv = "3"
			Case Else	vMwdiv = "1"
		End Select

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<SABANG_GOODS_REGI>"
		strRst = strRst & "	<HEADER>"
		strRst = strRst & "		<SEND_COMPAYNY_ID>"&sabangnetID&"</SEND_COMPAYNY_ID>"				'#사방넷 로그인 아이디
		strRst = strRst & "		<SEND_AUTH_KEY>"&sabangnetAPIKEY&"</SEND_AUTH_KEY>"					'#사방넷에서 발급 받은 인증키
		strRst = strRst & "		<SEND_DATE>"&Replace(Date(), "-", "")&"</SEND_DATE>"				'#전송일자 | YYYYMMDD
		strRst = strRst & "		<SEND_GOODS_CD_RT>Y</SEND_GOODS_CD_RT>"								'자체코드 반환여부 | 통신 성공시 결과에 자체코드 표시함 (Y : 반환, NULL : 없음)
		strRst = strRst & "	</HEADER>"
		strRst = strRst & "	<DATA>"
		strRst = strRst & "		<GOODS_NM><![CDATA["&getItemNameFormat&"]]></GOODS_NM>"				'#상품명 | 한글기준 50자리까지 사용가능하며 , HTML 태그 사용은 불가합니다.
		strRst = strRst & "		<GOODS_KEYWORD></GOODS_KEYWORD>"									'상품약어 | 간략한 상품명으로써 택배송장 출력과 물류 담당자의 빠른 인식을 위하여 사용할 수 있습니다. ( 단, "NULL"이면 수정안함)
		strRst = strRst & "		<MODEL_NM></MODEL_NM>"												'모델명 | 상품의 모델명을 정확히 기재합니다. ( 30자리까지 )
		strRst = strRst & "		<MODEL_NO></MODEL_NO>"												'모델No | 상품의 모델No.를 정확히 기재합니다. ( 30자리까지 )
		strRst = strRst & "		<BRAND_NM><![CDATA["&chkIIF(trim(FSocname_kor)="" or isNull(FSocname_kor),"상품설명 참조",FSocname_kor)&"]]></BRAND_NM>"	'브랜드명 | 브랜드명을 기재합니다.
		strRst = strRst & "		<COMPAYNY_GOODS_CD><![CDATA["& FItemid &"]]></COMPAYNY_GOODS_CD>"	'#자체상품코드 | 자사에서 사용하는 상품코드를 기재합니다. ( 30자리까지 )
		strRst = strRst & "		<GOODS_SEARCH><![CDATA["&RightCommaDel(Trim(getItemKeyword()))&"]]></GOODS_SEARCH>"	'사이트검색어 | 쇼핑몰로 상품정보 전송시 사용될 사이트검색어를 콤마(,)로 구분하여 입력합니다.( 단, "NULL"이면 수정안함)
		strRst = strRst & "		<GOODS_GUBUN><![CDATA["&vMwdiv&"]]></GOODS_GUBUN>"					'#상품구분 | 상품의 구분을 숫자로 입력합니다. 1.위탁상품 2.제조상품 3.사입상품 4.직영상품
		strRst = strRst & "		<CLASS_CD1><![CDATA["&FTenCateLarge&"]]></CLASS_CD1>"				'#대분류코드 | 사방넷에 등록된 대분류코드를 입력합니다.( 단, "NULL"이면 수정안함)
		strRst = strRst & "		<CLASS_CD2><![CDATA["&FTenCateMid&"]]></CLASS_CD2>"					'#중분류코드 | 사방넷에 등록된 중분류코드를 입력합니다.( 단, "NULL"이면 수정안함)
		strRst = strRst & "		<CLASS_CD3><![CDATA["&FTenCateSmall&"]]></CLASS_CD3>"				'#소분류코드 | 사방넷에 등록된 소분류코드를 입력합니다.( 단, "NULL"이면 수정안함)
		strRst = strRst & "		<CLASS_CD3><![CDATA[]]></CLASS_CD3>"								'세분류코드 | 사방넷에 등록된 세분류코드를 입력합니다.( 단, "NULL"이면 수정안함)
		strRst = strRst & "		<PARTNER_ID><![CDATA[]]></PARTNER_ID>"								'매입처ID | 매입처의 ID를 기재합니다.(대/소문자 정확히 구분해야 함)
		strRst = strRst & "		<DPARTNER_ID><![CDATA[]]></DPARTNER_ID>"							'물류처ID | 물류처의 ID를 기재합니다.(대/소문자 정확히 구분해야 함)
		strRst = strRst & "		<MAKER><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)&"]]></MAKER>"	'제조사 | 제조회사의 명칭을 정확히 기재합니다. ( 30자리까지 )
		strRst = strRst & "		<ORIGIN><![CDATA["&getSourcearea&"]]></ORIGIN>"						'#원산지(제조국) | 예:중국,사방넷 원산지 표를 참고하시어 표에 기재되어 있는 원산지 명으로 기입해주세요. 원산지가 등록되어 있지 않는 경우 콜센터로 요청하시기 바랍니다 ( "NULL" 이거나 "없는값" 인 경우 "기타" 로 입력됨)
		strRst = strRst & "		<MAKE_YEAR><![CDATA[]]></MAKE_YEAR>"								'생산연도 | 상품이 생산된 년도를 숫자 4자리로 입력합니다. 예 : 2009
		strRst = strRst & "		<MAKE_DM><![CDATA["& replace(Date(), "-", "") &"]]></MAKE_DM>"									'제조일자 | 상품이 제조된 일자를 숫자 8자리로 입력합니다. 예 : 20100101
		strRst = strRst & "		<GOODS_SEASON>7</GOODS_SEASON>"										'#시즌 | 계절의 구분을 숫자로 입력합니다. 1.봄 2.여름 3.가을 4.겨울 5.FW 6.SS 7.해당없음  ( 단, "NULL"이면 수정안함)
		strRst = strRst & "		<SEX>4</SEX>"														'#남녀구분 | 남여구분을 숫자로 입력합니다. 1.남성용 2.여성용 3.공용 4.해당없음 ( 단, "NULL"이면 수정안함)
		strRst = strRst & "		<STATUS>"&tmpStatus&"</STATUS>"										'#상품상태 | 상품의 공급상태에 대한 구분코드를 기재합니다. 1.대기중 2.공급중 3.일시중지 4.완전품절 5.미사용 6.삭제
		strRst = strRst & "		<DELIV_ABLE_REGION>1</DELIV_ABLE_REGION>"							'판매지역 | 판매가능지역을 숫자로 입력합니다. 1.전국 2.전국(도서제외) 3.수도권 4.기타
		strRst = strRst & "		<TAX_YN>"&Chkiif(Fvatinclude="Y", "1", "2")&"</TAX_YN>"				'#과세구분 | 과세여부를 숫자로 입력합니다. 1.과세 2.면세 3.자료없음 4.비과세
		strRst = strRst & "		<DELV_TYPE>"&Chkiif(IsFreeBeasong = True, "1", "3")&"</DELV_TYPE>"	'#배송비구분 | 배송비 구분을 숫자로 입력합니다. 1.무료 2.착불 3.선결제 4.착불/선결제
		strRst = strRst & "		<DELV_COST><![CDATA["&CHKIIF(IsFreeBeasong=False,"3000","0")&"]]></DELV_COST>"	'배송비 | 배송비를 숫자로 입력합니다. 첫글자는 반드시 '(ENTER좌측Key)로 시작해야하며 숫자사이에 콤마(,)가 들어가면 안됩니다.
		strRst = strRst & "		<BANPUM_AREA></BANPUM_AREA>"										'반품지구분 | 매입처의 복수의 반품지중 해당하는 순서를 기재합니다. 예 : 1, 공백일경우 기본주소가 적용됩니다.
		strRst = strRst & "		<GOODS_COST><![CDATA["&Clng(GetRaiseValue(FBuycash/10)*10)&"]]></GOODS_COST>"	'#원가 | 입력시 첫글자는 반드시 ( ‘ ) 아포스트로피(ENTER좌측Key)로 시작해야 하며 숫자 사이에 ( , ) 콤마가 들어가면 안됩니다.
		strRst = strRst & "		<GOODS_PRICE><![CDATA["& Clng(MustPrice/10)*10 &"]]></GOODS_PRICE>"	'#판매가 | 입력시 첫글자는 반드시 ( ‘ ) 아포스트로피(ENTER좌측Key)로 시작해야 하며 숫자 사이에 ( , ) 콤마가 들어가면 안됩니다.
		strRst = strRst & "		<GOODS_CONSUMER_PRICE><![CDATA["&Clng(FOrgPrice/10)*10&"]]></GOODS_CONSUMER_PRICE>"	'#TAG가(소비자가) | 입력시 첫글자는 반드시 ( ‘ ) 아포스트로피(ENTER좌측Key)로 시작해야 하며 숫자 사이에 ( , ) 콤마가 들어가면 안됩니다.
		strRst = strRst & getSabangnetOptParamtoREG()
		strRst = strRst & getSabangnetAddImageParam()
		strRst = strRst & "		<GOODS_REMARKS><![CDATA["&getSabangnetContParamToReg()&"]]></GOODS_REMARKS>"	'#상품상세설명 | 상품상세(HTML)을 기재합니다.
		strRst = strRst & getSabangnetCertInfoToReg()
		strRst = strRst & "		<MATERIAL><![CDATA[]]></MATERIAL>"									'식품재료/원산지 | "※식품의 재료와 원산지 구분은 /(슬러시)로 표기하며 추가 입력 시 ,(콤마)로 구분하여 추가할 재료와 원산지를 입력합니다. 판매식품 돼지갈비 예: 갈비/호수산,양념/국내산 "
		'############################################################
		'중요! : STOCK_USE_YN를 Y로 보내면 OPT_TYPE를 9로 전송해야함
		'		STOCK_USE_YN를 N으로 보내면 OPT_TYPE를 2로 전송 가능
		strRst = strRst & "		<STOCK_USE_YN><![CDATA[N]]></STOCK_USE_YN>"							'#재고관리사용여부 | "재고관리 사용여부를 Y or N로 입력합니다.  Y입력시 [재고관리] 메뉴에서 해당상품에 대한 입/출고가 가능하며, 쇼핑몰에 상품연동시 재고수량으로 연동됩니다. N입력시 [재고관리] 메뉴에서 [출고관리(주문)] 메뉴만 사용가능하며, 쇼핑몰에 상품연동시 가상재고로 연동됩니다. 단품별 가상재고 입력은 [상품관리] >> [단품대량수정] 에서 입력가능합니다. "
		strRst = strRst & "		<OPT_TYPE><![CDATA[2]]></OPT_TYPE>"									'#옵션수정여부 | "상품수정시 등록된 옵션의 내용을 모두 지우고 새로 등록하는 옵션입니다. 9: 옵션의 내용을 지우지 않는다. 사방넷 재고관리를 이용하는 업체인경우 옵션의 내용을 지우게 되면 기존에 가지고 있던 옵션코드값을 잃어버리기 때문에 재고관리에 큰문제가 발생하므로 재고관리와 연결되어 있는 업체라면 옵션지우지 않음을 선택하셔야 합니다. 한번 9로 선택된 상품은 다른 값선택이 불가능합니다. 2: 등록된 옵션의 내용을 모두 지우고 새로 옵션을 구성한다. 값을 2로 보내게 되면 그 상품에 적용되어 있는 옵션의 내용을 모두 삭제하고 보내신 내용으로 옵션을 재구성합니다. 이때, 이전에 옵션의 수정가능여부에 대해서 9로 선택하셨다면, 2로 변경은 불가능합니다. 단, 이전에 옵션의 수정가능여부에 2를 선택하셨다가 9로 선택은 가능합니다."
		'############################################################
		strRst = strRst & getSabangnetItemInfoCdToReg()
		strRst = strRst & "		<PACK_CODE_STR><![CDATA[]]></PACK_CODE_STR>"						'추가상품그룹코드 | 사방넷에 입력되어 구성된 추가상품의 그룹을 기재합니다. 예 : G001,G004,G201 (7개의 그룹이 입력 가능함)
		strRst = strRst & "		<GOODS_NM_EN><![CDATA[]]></GOODS_NM_EN>"							'영문 상품명 | 영문 100자리까지 사용가능하며 , HTML 태그 사용은 불가합니다.
		strRst = strRst & "		<GOODS_NM_PR><![CDATA[]]></GOODS_NM_PR>"							'출력 상품명 | 한글기준 50자리까지 사용가능하며 , HTML 태그 사용은 불가합니다.
		strRst = strRst & "		<GOODS_REMARKS2><![CDATA[]]></GOODS_REMARKS2>"						'추가 상품상세설명_1 | 상품 추가상세(HTML)을 기재합니다. (단, "DEL" 입력시 저장된 추가상세설명1을 삭제합니다.)
		strRst = strRst & "		<GOODS_REMARKS3><![CDATA[]]></GOODS_REMARKS3>"						'추가 상품상세설명_2 | 상품 추가상세(HTML)을 기재합니다. (단, "DEL" 입력시 저장된 추가상세설명2를 삭제합니다.)
		strRst = strRst & "		<GOODS_REMARKS4><![CDATA[]]></GOODS_REMARKS4>"						'추가 상품상세설명_3 | 상품 추가상세(HTML)을 기재합니다. (단, "DEL" 입력시 저장된 추가상세설명3을 삭제합니다.)
		strRst = strRst & "		<IMPORTNO><![CDATA[]]></IMPORTNO>"									'수입신고번호 | 상품 수입신고번호를 기재합니다. (12345-12-123456U)
		strRst = strRst & "		<GOODS_COST2><![CDATA[]]></GOODS_COST2>"							'원가2 | 원가2는 상품송신,주문매핑,매출집계,정산등에 이용되지 않으며, 관리상 참고를 위한 가격입니다.
		strRst = strRst & "		<ORIGIN2><![CDATA[]]></ORIGIN2>"									'원산지 상세지역 | 원산지 상세 정보를 입력하세요.
		strRst = strRst & "		<EXPIRE_DM><![CDATA[]]></EXPIRE_DM>"								'유효일자 | 숫자8자리 입력하세요 예:20100401
		strRst = strRst & "		<SUPPLY_SAVE_YN><![CDATA[N]]></SUPPLY_SAVE_YN>"						'합포제외여부 | 합포장 제외 여부를 Y or N로 입력하세요. "Y" 입력시 합포장 제외 항목에 체크됩니다.
		strRst = strRst & "		<DESCRITION><![CDATA[]]></DESCRITION>"								'관리자메모 | 관리자 메모 내용을 입력하세요
		strRst = strRst & "	</DATA>"
		strRst = strRst & "</SABANG_GOODS_REGI>"
		getSabangnetItemRegParameter = strRst
	End Function

	'상품 요약 수정 XML
	Public Function getSabangnetSimpleEditItemParameter(ichgSellyn)
		Dim strRst, tmpStatus

		If ichgSellyn <> "" Then
			Select Case ichgSellyn
				Case "Y"	tmpStatus = 2
				Case "N"	tmpStatus = 3
			End Select
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<SABANG_GOODS_REGI>"
		strRst = strRst & "	<HEADER>"
		strRst = strRst & "		<SEND_COMPAYNY_ID>"&sabangnetID&"</SEND_COMPAYNY_ID>"				'#사방넷 로그인 아이디
		strRst = strRst & "		<SEND_AUTH_KEY>"&sabangnetAPIKEY&"</SEND_AUTH_KEY>"					'#사방넷에서 발급 받은 인증키
		strRst = strRst & "		<SEND_DATE>"&Replace(Date(), "-", "")&"</SEND_DATE>"				'#전송일자 | YYYYMMDD
		strRst = strRst & "		<SEND_GOODS_CD_RT>Y</SEND_GOODS_CD_RT>"								'자체코드 반환여부 | 통신 성공시 결과에 자체코드 표시함 (Y : 반환, NULL : 없음)
		strRst = strRst & "	</HEADER>"
		strRst = strRst & "	<DATA>"
		strRst = strRst & "		<GOODS_NM><![CDATA["&getItemNameFormat&"]]></GOODS_NM>"				'#상품명 | 한글기준 50자리까지 사용가능하며 , HTML 태그 사용은 불가합니다.
		strRst = strRst & "		<COMPAYNY_GOODS_CD><![CDATA["& FItemid &"]]></COMPAYNY_GOODS_CD>"	'#자체상품코드 | 자사에서 사용하는 상품코드를 기재합니다. ( 30자리까지 )
		strRst = strRst & "		<STATUS>"&tmpStatus&"</STATUS>"
		strRst = strRst & "		<GOODS_COST><![CDATA["&Clng(GetRaiseValue(FBuycash/10)*10)&"]]></GOODS_COST>"	'#원가 | 입력시 첫글자는 반드시 ( ‘ ) 아포스트로피(ENTER좌측Key)로 시작해야 하며 숫자 사이에 ( , ) 콤마가 들어가면 안됩니다.
		strRst = strRst & "		<GOODS_PRICE><![CDATA["& Clng(MustPrice/10)*10 &"]]></GOODS_PRICE>"	'#판매가 | 입력시 첫글자는 반드시 ( ‘ ) 아포스트로피(ENTER좌측Key)로 시작해야 하며 숫자 사이에 ( , ) 콤마가 들어가면 안됩니다.
		strRst = strRst & "		<GOODS_CONSUMER_PRICE><![CDATA["&Clng(FOrgPrice/10)*10&"]]></GOODS_CONSUMER_PRICE>"	'#TAG가(소비자가) | 입력시 첫글자는 반드시 ( ‘ ) 아포스트로피(ENTER좌측Key)로 시작해야 하며 숫자 사이에 ( , ) 콤마가 들어가면 안됩니다.
		strRst = strRst & "	</DATA>"
		strRst = strRst & "</SABANG_GOODS_REGI>"
		getSabangnetSimpleEditItemParameter = strRst
	End Function

	'쇼핑몰별 DATA수정 XML
	Public Function getSabangnetShoppingMallEditParameter
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<SABANG_GOODS_REGI>"
		strRst = strRst & "	<HEADER>"
		strRst = strRst & "		<SEND_COMPAYNY_ID>"&sabangnetID&"</SEND_COMPAYNY_ID>"				'#사방넷 로그인 아이디
		strRst = strRst & "		<SEND_AUTH_KEY>"&sabangnetAPIKEY&"</SEND_AUTH_KEY>"					'#사방넷에서 발급 받은 인증키
		strRst = strRst & "		<SEND_DATE>"&Replace(Date(), "-", "")&"</SEND_DATE>"				'#전송일자 | YYYYMMDD
		strRst = strRst & "		<SEND_GOODS_CD_RT>Y</SEND_GOODS_CD_RT>"								'자체코드 반환여부 | 통신 성공시 결과에 자체코드 표시함 (Y : 반환, NULL : 없음)
		strRst = strRst & "	</HEADER>"
		strRst = strRst & "	<DATA>"
		strRst = strRst & "		<MALL_CODE>shop0060</MALL_CODE>"			'#쇼핑몰CODE | 쇼핑몰 코드를 기재합니다. (사방넷메뉴 A>2 쇼핑몰관리(지원) 메뉴 참조)
		strRst = strRst & "		<COMPAYNY_GOODS_CD><![CDATA["& FItemid &"]]></COMPAYNY_GOODS_CD>"	'#자체상품코드 | 자사에서 사용하는 상품코드를 기재합니다. ( 30자리까지 )
		strRst = strRst & "		<MALL_PROP1_CD>008</MALL_PROP1_CD>"	'속성분류코드 | 속성분류코드를 숫자 3자리 형식으로 입력합니다. 속성분류코드는 상품속성코드 조회 API나 사방넷 상품관리화면의 속성분류표를 참고하시기 바랍니다. 예: 의류는 001을 입력합니다.
		strRst = strRst & "	</DATA>"
		strRst = strRst & "</SABANG_GOODS_REGI>"
		getSabangnetShoppingMallEditParameter = strRst
	End Function

End Class

Class CSabangnet
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
	Public Sub getSabangnetNotRegOneItem
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
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isnull(c.safetyyn, '') as safetyyn, isnull(c.safetyNum, '') as safetyNum, isnull(c.safetydiv, '') as safetydiv "
		strSql = strSql & "	, isNULL(R.sabangnetStatCD,-9) as sabangnetStatCD "
		strSql = strSql & "	, UC.socname_kor, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_sabangnet_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7')"

		'2020-10-27 김진영..사방넷은 1만원 미만도 등록하게 해달라심..by 소정
		' IF (CUPJODLVVALID) then
		' 	strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		' ELSE
		'     strSql = strSql & " and (i.deliveryType<>9)"
	    ' END IF

		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "					'플라워/화물배송/해외직구
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
'		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and i.itemdiv not in ('21', '23', '30') "
		strSql = strSql & " and i.itemdiv <> '06' "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and isnull(R.sabangnetGoodNo, '') = '' "
		strSql = strSql & "		"	& addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CSabangnetItem
				FOneItem.FItemid			= rsget("itemid")
				FOneItem.FTenCateLarge		= rsget("cate_large")
				FOneItem.FTenCateMid		= rsget("cate_mid")
				FOneItem.FTenCateSmall		= rsget("cate_small")
				FOneItem.FItemname			= db2html(rsget("itemname"))
				FOneItem.FItemDiv			= rsget("itemdiv")
				FOneItem.FSmallImage		= rsget("smallImage")
				FOneItem.FMakerid			= rsget("makerid")
				FOneItem.FRegdate			= rsget("regdate")
				FOneItem.FLastUpdate		= rsget("lastUpdate")
				FOneItem.FOrgPrice			= rsget("orgPrice")
				FOneItem.FOrgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FSellYn			= rsget("sellYn")
				FOneItem.FSaleYn			= rsget("sailyn")
				FOneItem.FIsUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.FKeywords			= rsget("keywords")
				FOneItem.FVatinclude        = rsget("vatinclude")
				FOneItem.FOrderComment		= db2html(rsget("ordercomment"))
				FOneItem.FOptionCnt			= rsget("optionCnt")
				If isnull(rsget("basicImage600")) or rsget("basicImage600") = "" Then
					FOneItem.FBasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				Else
					FOneItem.FBasicImage		= "http://webimage.10x10.co.kr/image/basic600/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage600")
				End If
				FOneItem.FMainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FMainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.FIcon1Image		= "http://webimage.10x10.co.kr/image/icon1/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon1Image")
				FOneItem.FIcon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.FListimage			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
				FOneItem.FSourcearea		= db2html(rsget("sourcearea"))
				FOneItem.FMakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.FItemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetydiv			= rsget("safetydiv")
				FOneItem.FSabangnetStatCD	= rsget("sabangnetStatCD")
				FOneItem.FDeliverfixday		= rsget("deliverfixday")
				FOneItem.FDeliverytype		= rsget("deliverytype")
				FOneItem.FSocname_kor		= rsget("socname_kor")
				FOneItem.FBasicimageNm 		= rsget("basicimage")
				FOneItem.FItemsize 			= rsget("itemsize")
				FOneItem.FItemsource 		= rsget("itemsource")
				FOneItem.FMwdiv 			= rsget("mwdiv")
		End If
		rsget.Close
	End Sub

	Public Sub getSabangnetEditOneItem
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
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(C.safetyyn,'N') as safetyyn, isnull(c.safetyNum, '') as safetyNum, isNULL(C.safetyDiv, '') as safetyDiv "
		strSql = strSql & "	, m.sabangnetGoodNo, m.sabangnetprice, m.sabangnetSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr "
		strSql = strSql & "	, C.infoDiv, m.sabangnetStatCD, UC.socname_kor, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType = 7"
'2020-10-27 김진영..사방넷은 1만원 미만도 등록하게 해달라심..by 소정
'		strSql = strSql & "		or ((i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & " 	or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & " 	or i.itemdiv = '06' "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' "
'		strSql = strSql & "		or i.cate_large = '999' "
		strSql = strSql & "		or i.cate_large=''"
		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_sabangnet_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and m.sabangnetStatCD = 7 "
		strSql = strSql & addSql
		strSql = strSql & " and m.sabangnetGoodNo is Not Null "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CSabangnetItem
				FOneItem.FItemid			= rsget("itemid")
				FOneItem.FTenCateLarge		= rsget("cate_large")
				FOneItem.FTenCateMid		= rsget("cate_mid")
				FOneItem.FTenCateSmall		= rsget("cate_small")
				FOneItem.FItemname			= db2html(rsget("itemname"))
				FOneItem.FItemDiv			= rsget("itemdiv")
				FOneItem.FSmallImage		= rsget("smallImage")
				FOneItem.FMakerid			= rsget("makerid")
				FOneItem.FRegdate			= rsget("regdate")
				FOneItem.FLastUpdate		= rsget("lastUpdate")
				FOneItem.FOrgPrice			= rsget("orgPrice")
				FOneItem.FOrgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FSellYn			= rsget("sellYn")
				FOneItem.FSaleYn			= rsget("sailyn")
				FOneItem.FIsUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.FKeywords			= rsget("keywords")
				FOneItem.FVatinclude        = rsget("vatinclude")
				FOneItem.FOrderComment		= db2html(rsget("ordercomment"))
				FOneItem.FOptionCnt			= rsget("optionCnt")
				If isnull(rsget("basicImage600")) or rsget("basicImage600") = "" Then
					FOneItem.FBasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				Else
					FOneItem.FBasicImage		= "http://webimage.10x10.co.kr/image/basic600/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage600")
				End If
				FOneItem.FMainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FMainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.FIcon1Image		= "http://webimage.10x10.co.kr/image/icon1/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon1Image")
				FOneItem.FIcon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.FListimage			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
				FOneItem.FSourcearea		= db2html(rsget("sourcearea"))
				FOneItem.FMakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.FItemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetydiv			= rsget("safetydiv")
				FOneItem.FSabangnetStatCD	= rsget("sabangnetStatCD")
				FOneItem.FDeliverfixday		= rsget("deliverfixday")
				FOneItem.FDeliverytype		= rsget("deliverytype")
				FOneItem.FSocname_kor		= rsget("socname_kor")
				FOneItem.FBasicimageNm 		= rsget("basicimage")
				FOneItem.FItemsize 			= rsget("itemsize")
				FOneItem.FItemsource 		= rsget("itemsource")
				FOneItem.FMwdiv 			= rsget("mwdiv")
				FOneItem.FmaySoldOut 		= rsget("maySoldOut")
		End If
		rsget.Close
	End Sub

	Public Sub getSabangnetSimpleEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.*, m.sabangnetGoodNo, m.sabangnetprice, m.sabangnetSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType = 7"
'2020-10-27 김진영..사방넷은 1만원 미만도 등록하게 해달라심..by 소정
'		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & " 	or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & " 	or i.itemdiv = '06' "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' "
'		strSql = strSql & "		or i.cate_large = '999' "
		strSql = strSql & "		or i.cate_large=''"
		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_sabangnet_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and m.sabangnetStatCD = 7 "
		strSql = strSql & addSql
		strSql = strSql & " and m.sabangnetGoodNo is Not Null "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CSabangnetItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FMakerid			= rsget("makerid")
				FOneItem.FItemname			= db2html(rsget("itemname"))
				FOneItem.FSabangnetGoodNo	= rsget("sabangnetGoodNo")
				FOneItem.FSabangnetprice	= rsget("sabangnetprice")
				FOneItem.FSabangnetSellYn	= rsget("sabangnetSellYn")
	            FOneItem.FOptionCnt         = rsget("optionCnt")
	            FOneItem.FRegedOptCnt       = rsget("regedOptCnt")
				FOneItem.FOrgPrice			= rsget("orgPrice")
				FOneItem.FOrgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FSellYn			= rsget("sellYn")
				FOneItem.FSaleYn			= rsget("sailyn")
				FOneItem.FIsUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
	            FOneItem.FMaySoldOut		= rsget("maySoldOut")
		End If
		rsget.Close
	End Sub
End Class

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
%>