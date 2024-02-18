<%
CONST CMAXMARGIN = 15			'' MaxMagin임.. '(롯데닷컴 11%)
CONST CMAXLIMITSELL = 5        '' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CMALLNAME = "lotteCom"
CONST CDEFALUT_STOCK = 99       '' 재고관리 수량 기본 99 (한정 아닌경우)

Class CLotteItem
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
	Public FbasicImage
	Public FbasicimageNm
	Public FregImageName
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FitemGbnKey
	Public FLotteStatcd
	Public FLotteGoodNo
	Public FLotteTmpGoodNo
	Public FLotteSellYn
	Public FLottePrice
	Public FregedOptCnt
	Public FaccFailCNT
	Public FlastErrStr
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FmaySoldOut

	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public Fsocname_kor
	Public FAdultType

	'// 품절여부
	public function IsSoldOut()
		If ((FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))) Then
			IsSoldOut = "Y"
		Else
			IsSoldOut = "N"
		End If
	end function

	'// 품절여부
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Public Function getBasicImage()
		If IsNULL(FbasicImageNm) or (FbasicImageNm="") Then Exit function
		getBasicImage = FbasicImageNm
	End Function

	Public Function isImageChanged()
		Dim ibuf : ibuf = getBasicImage
		If InStr(ibuf,"-") < 1 Then
			isImageChanged = FALSE
			Exit Function
		End If
		isImageChanged = ibuf <> FregImageName
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, tmpPrice, sqlStr, specialPrice
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
			tmpPrice = specialPrice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If GetTenTenMargin < CMAXMARGIN Then
				tmpPrice = Forgprice
			Else
				tmpPrice = FSellCash
			End If
		End If
		MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
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

    public function getLotteReturnDLVCode
        dim sqlStr
        dim retReturnCode
        getLotteReturnDLVCode = "333126"    ''기본 반품지 코드(물류)
		Exit function '' 물류로 재통합 2013/02/25
        if Not (Fdeliverytype="2" or Fdeliverytype="7" or Fdeliverytype="9") then Exit function  ''자체배송이면 물류


'        ''임시 막음
'        if (Fcorp_dlvp_sn<>"72125") and (Fcorp_dlvp_sn<>"113045") and (Fcorp_dlvp_sn<>"113044") and (Fcorp_dlvp_sn<>"114747") then
'            Fcorp_dlvp_sn = "72125"
'        end if

        '' 반품주소지가 같은지 검사 방식으로 변경
        sqlStr = " select R.returnCode from db_item.dbo.tbl_OutMall_BrandReturnCode R"
        sqlStr = sqlStr & "  	Join db_temp.dbo.tbl_jaehyumall_returnInfo T"
        sqlStr = sqlStr & "	on R.makerid='"&Fmakerid&"'"
        sqlStr = sqlStr & "	and R.returnCode=T.returnCode"
        sqlStr = sqlStr & "	Join db_partner.dbo.tbl_partner p"
        sqlStr = sqlStr & "	on p.id=R.makerid"
        sqlStr = sqlStr & " where replace(T.returnAddress,' ','')=replace(replace(p.return_zipCode,'-','') +  p.return_address + p.return_address2,' ','')"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if Not(rsget.EOF or rsget.BOF) then
		    retReturnCode = rsget("returnCode")
        end if
        rsget.close

        if isNULL(retReturnCode) then Exit function
        if (retReturnCode="") then Exit function

        getLotteReturnDLVCode = CStr(retReturnCode)
    end function

	'// 롯데닷컴 판매여부 반환
	public function getLotteSellYn()
		'판매상태 (10:판매진행, 20:품절)
		if FsellYn="Y" and FisUsing="Y" then
			if FLimitYn="N" or (FLimitYn="Y" and FLimitNo-FLimitSold > CMAXLIMITSELL) then
				getLotteSellYn = "Y"
			else
				getLotteSellYn = "N"
			end if
		else
			getLotteSellYn = "N"
		end if
	end Function

    public function getLimitLotteEa()
        dim ret
        ret = FLimitNo-FLimitSold-5
		'2013-10-10 진영 추가..한정 중 1000개가 넘으면 최대 999로 넣겠다고 롯데에서 요청메일 옴
		If ret > 1000 Then
			 ret = 999
		End If

        if (ret<1) then ret=0
        getLimitLotteEa = ret
    end function

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
        buf = replace(FItemName,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","프로")
        buf = replace(buf,"[무료배송]","")
        buf = replace(buf,"[무료 배송]","")
        getItemNameFormat = buf
    end function

	'// 검색어
	Public Function getItemKeyword()
		Dim arrRst, arrRst2, q, p, r, divBound1, divBound2, divBound3, Keyword1, Keyword2, Keyword3, strRst
		If trim(Fkeywords) = "" Then Exit Function
        Fkeywords = replace(Fkeywords,"/",",") ''2015/06/22 ::981002,952467,952466
		If Len(Fkeywords) > 50 Then
			arrRst = Split(Fkeywords,",")
'#########2016-08-10 김진영 하단 주석 후 개선된 소스로 적용#########
'			If Ubound(arrRst) = 0 then
'				'구분이 공백일 경우
'				arrRst2 = split(arrRst(0)," ")
'				If Ubound(arrRst2) > 0 then
'					arrRst = split(Fkeywords," ")
'				End If
'			End If

			If Ubound(arrRst) = 0 then
				'구분이 공백일 경우
				If Ubound(split(arrRst(0)," ")) > 0 then
					arrRst2 = split(arrRst(0)," ")
					arrRst = split(Fkeywords," ")
				ElseIf Ubound(split(arrRst(0),";")) > 0 then
					arrRst2 = split(arrRst(0),";")
					arrRst = split(Fkeywords,";")
				End If
			End If
'##############################################################

			If Ubound(arrRst) = 2 Then	'2015-06-29 김진영 수정 : ex)769674
				Keyword1 = arrRst(0)
				Keyword2 = arrRst(1)
				Keyword3 = arrRst(2)
			Else
				'키워드 1
				divBound1 = Cint(Ubound(arrRst)/3)
				For q = 0 to divBound1
					If lenb(Keyword1) > 80 Then
						Exit For
					End If
					Keyword1 = Keyword1&arrRst(q)&","
				Next
				If Right(keyword1,1) = "," Then
					keyword1 = Left(keyword1,Len(keyword1)-1)
				End If

				'키워드 2
				divBound2 = divBound1 + 1
				For p = divBound2 to divBound2 + divBound1
					Keyword2 = Keyword2&arrRst(p)&","
					If lenb(Keyword2) > 80 Then
						Exit For
					End If
				Next
				If Right(keyword2,1) = "," Then
					keyword2 = Left(keyword2,Len(keyword2)-1)
				End If

				'키워드 3
				divBound3 = divBound2 + divBound1
				For r = divBound3 to Ubound(arrRst)
					Keyword3 = Keyword3&arrRst(r)&","
					If lenb(Keyword3) > 80 Then
						Exit For
					End If
				Next
				If Right(keyword3,1) = "," Then
					keyword3 = Left(keyword3,Len(keyword3)-1)
				End If
			End If

			strRst = ""
			strRst = strRst & "&sch_kwd_1_nm=" & Server.URLEncode(Keyword1)
			strRst = strRst & "&sch_kwd_2_nm=" & Server.URLEncode(Keyword2)
			strRst = strRst & "&sch_kwd_3_nm=" & Server.URLEncode(Keyword3)
			getItemKeyword = strRst
		Else
			strRst = ""
			strRst = strRst & "&sch_kwd_1_nm="&Server.URLEncode(Fkeywords)
			strRst = strRst & "&sch_kwd_2_nm="
			strRst = strRst & "&sch_kwd_3_nm="
			getItemKeyword = strRst
		End If
	End Function

	'// 상품등록: 옵션 파라메터 생성(상품등록용)
	public function getLotteOptionParamToReg()
		dim strSql, strRst, i, optYn, optNm, optDc, chkMultiOpt, optLimit
		chkMultiOpt = false
		optYn = "N"

		if FoptionCnt>0 then
			'// 이중옵션일 때
			'#옵션명 생성
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget

			optNm = ""
			if Not(rsget.EOF or rsget.BOF) then
				chkMultiOpt = true
				optYn = "Y"
				Do until rsget.EOF
					optNm = optNm & Replace(db2Html(rsget("optionTypeName")),":","")
					rsget.MoveNext
					if Not(rsget.EOF) then optNm = optNm & ":"
				Loop
			end if
			rsget.Close

			'#옵션내용 생성
			if chkMultiOpt then
				strSql = "Select optionname, (optlimitno-optlimitsold) as optLimit "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				'''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) " ''일단 입력
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				optDc = ""
				if Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit-5
					    if (optLimit<1) then optLimit=0
					    if (FLimitYN<>"Y") then optLimit=CDEFALUT_STOCK   ''2013/06/12 재고관리여부 모두 Y로 변경 되므로

						optDc = optDc & Replace(Replace(db2Html(rsget("optionname")),":",""),"'","") & "," & optLimit
						rsget.MoveNext
						if Not(rsget.EOF) then optDc = optDc & ":"
					Loop
				end if
				rsget.Close
			end if


			'// 단일옵션일 때
			if Not(chkMultiOpt) then
				strSql = "Select optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				if Not(rsget.EOF or rsget.BOF) then
					optYn = "Y"
					if db2Html(rsget("optionTypeName"))<>"" then
						optNm = Replace(db2Html(rsget("optionTypeName")),":","")
					else
						optNm = "옵션"
					end if
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit-5
					    if (optLimit<1) then optLimit=0
					    if (FLimitYN<>"Y") then optLimit=CDEFALUT_STOCK   ''2013/06/12 재고관리여부 모두 Y로 변경 되므로

						optDc = optDc & Replace(Replace(Replace(db2Html(rsget("optionname")),":",""),",",""),"'","") & "," & optLimit
						rsget.MoveNext
						if Not(rsget.EOF) then optDc = optDc & ":"
					Loop
				end if
				rsget.Close
			end if
		end if

		strRst = strRst & "&item_mgmt_yn=" & optYn						'단품관리여부(옵션)
		strRst = strRst & "&opt_nm=" & server.URLEncode(optNm)			'옵션명
		strRst = strRst & "&item_list=" & server.URLEncode(optDc)		'옵션상세

		getLotteOptionParamToReg = strRst
	end function

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	public function getLotteItemContParamToReg()
		dim strRst, strSQL
		strRst = Server.URLEncode("<div align=""center"">")
		'2014-01-17 10:00 김진영 탑 이미지 추가
		strRst = strRst & Server.URLEncode("<p><a href=""http://www.lotte.com/display/viewDispShop.lotte?disp_no=5293948"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_lotteCom.jpg""></a></p>")

		If ForderComment <> "" Then
			strRst = strRst & Server.URLEncode("- 주문시 유의사항 :<br>" & Fordercomment & "<br>")
		End If

		'#기본 상품설명
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & Server.URLEncode(oLotteitem.FOneItem.Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & Server.URLEncode(nl2br(oLotteitem.FOneItem.Fitemcontent) & "<br>")
			Case Else
				strRst = strRst & Server.URLEncode(nl2br(ReplaceBracket(oLotteitem.FOneItem.Fitemcontent)) & "<br>")
		End Select

		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		if Not(rsget.EOF or rsget.BOF) then
			Do Until rsget.EOF
				if rsget("imgType")="1" then
					strRst = strRst & Server.URLEncode("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0""><br>")
				end if
				rsget.MoveNext
			Loop
		end if

		rsget.Close

		'#기본 상품 설명이미지
		if ImageExists(FmainImage) then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage & """ border=""0""><br>")
		if ImageExists(FmainImage2) then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage2 & """ border=""0""><br>")

		'#배송 주의사항
		strRst = strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg"">")

		strRst = strRst & Server.URLEncode("</div>")

		getLotteItemContParamToReg = "&dtl_info_fcont=" & strRst

		''660877 db_item.dbo.tbl_OutMall_etcLink 쿼리 후 값이 존재하면 시작
		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		if Not(rsget.EOF or rsget.BOF) then
			strRst = Server.URLEncode(""&rsget("textVal")&"")
			strRst = Server.URLEncode("<div align=""center""><p><a href=""http://www.lotte.com/display/viewDispShop.lotte?disp_no=5293948"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_lotteCom.jpg""></a></p>") & strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg""></div>")
			getLotteItemContParamToReg = "&dtl_info_fcont=" & strRst
		End If
		rsget.Close
		''660877 db_item.dbo.tbl_OutMall_etcLink 쿼리 후 값이 존재하면 끝

'		if (FItemID="502049") or (FItemID="660877") then
'		    ''getLotteItemContParamToReg = "&dtl_info_fcont="&Server.URLEncode("<div></div>")     ''
'		    getLotteItemContParamToReg = ""                                           '' 연동안하려면 파라메터 없이 또는빈값
'		end if
	end function

	'// 상품등록: MD상품군 및 전시 카테고리 파라메터 생성(상품등록용)
	public function getLotteCateParamToReg()
		dim strSql, strRst, i, ogrpCode
		strSql = "Select top 6 c.groupCode, m.dispNo, c.disptpcd "
		strSql = strSql & " from db_item.dbo.tbl_lotte_cate_mapping as m "
		strSql = strSql & " 	join db_temp.dbo.tbl_lotte_Category as c "
		strSql = strSql & " 		on m.DispNO=c.DispNO "
		strSql = strSql & " where tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " 	and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " 	and tenCateSmall='" & FtenCateSmall & "' "
	    strSql = strSql & " 	and c.isusing='Y'"
		strSql = strSql & " order by (CASE WHEN c.disptpcd='12' THEN 'ZZ' ELSE c.disptpcd END) desc"           ''''//전문몰을 기본 카테고리로..
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
		    ogrpCode = rsget("groupCode")
			strRst = "&md_gsgr_no=" & ogrpCode				            'md상품군 코드 (md상품군 코드 같은 카테고리만 등록가능)
			i = 0
			Do until rsget.EOF
'			If FItemid = "1545080" Then
				If (rsget("disptpcd")="12") then                        ''전문몰 카테고리를 기본으로 하라고.. /2012/06/14
				    strRst = strRst & "&disp_no_b=" & rsget("dispNo")		'전시카테고리(선택)
				Else
				    IF (ogrpCode=rsget("groupCode")) then
    				    strRst = strRst & "&disp_no=" & rsget("dispNo") 	'전시카테고리(필수)
    				End IF
			    End If
'			Else
'				If (rsget("disptpcd")="12") then                        ''전문몰 카테고리를 기본으로 하라고.. /2012/06/14
'				    strRst = strRst & "&disp_no=" & rsget("dispNo")		'기본 전시카테고리
'				Else
'				    IF (ogrpCode=rsget("groupCode")) then
'						strRst = strRst & "&disp_no_b=" & rsget("dispNo") 	'추가 전시카테고리
'					End IF
'				End If
'			End If
				rsget.MoveNext
				i = i + 1
			Loop
		End If
		rsget.Close
		getLotteCateParamToReg = strRst
	end function

	'// 텐바이텐 상품옵션 검사
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
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold > "&CMAXLIMITSELL&")) "
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
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold > "&CMAXLIMITSELL&")) "
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

    ''국산여부
    public Function isMadeInKorea(iVal)
        isMadeInKorea = False
        if isNULL(iVal) then Exit Function
        iVal = Trim(iVal)

        if (iVal="한국") or (iVal="대한민국") or (iVal="KOREA") or (iVal="KOREA 한국") or (iVal="KOREA 국내산") then
            isMadeInKorea = True
        end if

        if (iVal="국산") or (iVal="국내제작") or (iVal="한국OEM") or (iVal="국내제작(한국)") or (iVal="국내") then
            isMadeInKorea = True
        end if

        if (iVal="한국 / 코지아트") then
            isMadeInKorea = True
        end if
    end Function

	public function getLotteGoodDLVDtParams()
        dim strRst
        strRst = ""

		If Fmakerid = "bugangit10" Then		 '2013-07-26 현아주임님 요청..특정 브랜드는 무조건 일반으로 돌리라고 해서..
		    strRst = strRst & "&dlv_goods_sct_cd=01"
    		strRst = strRst & "&dlv_dday=3"
    		getLotteGoodDLVDtParams = strRst
    		Exit Function
		End If

		If Fitemid = "305876" OR Fitemid = "303848" OR Fitemid = "305877" OR Fitemid = "305878" Then
		    strRst = strRst & "&dlv_goods_sct_cd=03"
    		strRst = strRst & "&dlv_dday=10"
    		getLotteGoodDLVDtParams = strRst
			Exit Function
		End If

        if (FtenCateLarge="055") or (FtenCateLarge="040") then ''가구/패브릭 15일로
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=15"
		elseif (FtenCateLarge="080") or (FtenCateLarge="100") then  ''우먼/베이비 5일
		    strRst = strRst & "&dlv_goods_sct_cd=03" 																						'배송상품구분		(*:주문제작03)
		    strRst = strRst & "&dlv_dday=5"
		elseif ((FtenCateLarge="045") and (FtenCateMid="001" or FtenCateMid="004")) then  ''수납/생활> 옷/이불수납 or 주방수납 10일 - 현아씨요청 2013/01/22
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
		elseif ((FtenCateLarge="025") and (FtenCateMid="107")) then  ''디지털 > 기타 스마트기기 케이스  10일 - 현아씨요청 2013/01/22
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
	    elseif ((FtenCateLarge="050") and (FtenCateMid="777")) then   ''홈/데코 > 거울   - 미희씨요청 2013/03/08
	        strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
		elseif ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="001")) then    ''HOME > 수납/생활 > 보관/정리용품 > 수납장 			주문제작15일 045&cdm=002&cds=001
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=15"
		elseif ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="002")) then    ''HOME > 수납/생활 > 보관/정리용품 > 틈새수납장			주문제작10일
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="005")) then    ''HOME > 수납/생활 > 보관/정리용품 > 잡지꽂이 			주문제작10일
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
		elseif ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="001")) then    ''HOME > 수납/생활 > 데코수납 > 우드박스 				주문제작10일
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="007")) then    ''HOME > 수납/생활 > 데코수납 > 인터폰박스 			               주문제작10일
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="050") and (FtenCateMid="060") and (FtenCateSmall="070")) then    ''HOME > 홈/데코 > 소품박스/바구니 > 인터폰박스			주문제작10일 cdl=050&cdm=060&cds=070
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="110") and (FtenCateMid="090") and (FtenCateSmall="040")) then    ''HOME > 감성채널 > DIY > 나무로만들기 				주문제작10일 110&cdm=090&cds=040
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="045") and (FtenCateMid="010")) then   ''수납/생활 > 디자인선반  - 미희씨요청 2013/03/08
	        strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
'		elseif (FtenCateLarge="025")  then  ''디지털 10일 - 미희씨요청 2013/01/17
'		    strRst = strRst & "&dlv_goods_sct_cd=03" 																						'배송상품구분		(*:주문제작03)
'		    strRst = strRst & "&dlv_dday=10"
		elseif ((FitemDiv="06") or (FitemDiv="16")) then    ''주문(후)제작상품
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    if (FrequireMakeDay>7) then
		        strRst = strRst & "&dlv_dday="&CStr(FrequireMakeDay)
		    elseif (FrequireMakeDay<1) then
		        strRst = strRst & "&dlv_dday=7"
		    else
		        strRst = strRst & "&dlv_dday="&(FrequireMakeDay+1)
		    end if
		else
		    strRst = strRst & "&dlv_goods_sct_cd=01" 																						'배송상품구분		(*:일반상품)
    		strRst = strRst & "&dlv_dday=3" 																								'배송기일			(*:3일이내)
    	end if

    	getLotteGoodDLVDtParams = strRst
    end function

	'// 상품등록: 상품추가이미지 파라메터 생성(상품등록용)
	public function getLotteAddImageParamToReg()
		dim strRst, strSQL, i

		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		if Not(rsget.EOF or rsget.BOF) then
			for i=1 to rsget.RecordCount
				if rsget("imgType")="0" then
					strRst = strRst & "&img_url" & i & "=" & Server.URLEncode("http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400"))
				end if
				rsget.MoveNext
				if i>=5 then Exit For
			next
		end if

		rsget.Close

		getLotteAddImageParamToReg = strRst
	end function

	'2012/11/02 김진영 생성 상품품목정보 파라메타
	Public Function getLotteItemInfoCdToReg()
		Dim strRst, strSQL
		Dim anjunInfo, mallinfoDiv, mallinfoCdAll,mallinfoCd, infoCDVal, psourceArea
        Dim bufTxt : bufTxt=""
		Dim certNum, safetyDiv, isTenReg
		isTenReg = "N"

		strSQL = ""
		strSQL = strSQL & " SELECT TOP 1 certNum, safetyDiv " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_safetycert_tenReg " & vbcrlf
		strSQL = strSQL & " WHERE itemid='"&FItemID&"' " & vbcrlf
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			certNum = rsget("certNum")
			safetyDiv = rsget("safetyDiv")
			isTenReg = "Y"
		End If
		rsget.Close

		If isTenReg = "Y" Then
			anjunInfo = anjunInfo & "&sft_cert_tgt_yn=Y"
			If (safetyDiv = "10") OR (safetyDiv = "40") Then
				anjunInfo = anjunInfo & "&kps_1_no="
				anjunInfo = anjunInfo & "&kps_3_no="&Server.URLEncode(certNum)
			ElseIf (safetyDiv = "20") OR (safetyDiv = "50") Then
				anjunInfo = anjunInfo & "&kps_1_no="
				anjunInfo = anjunInfo & "&kps_2_no="&Server.URLEncode(certNum)
			ElseIf (safetyDiv = "70") OR (safetyDiv = "80") Then
				anjunInfo = anjunInfo & "&kps_1_no="
				anjunInfo = anjunInfo & "&kps_5_no="&Server.URLEncode(certNum)
			Else
				anjunInfo = ""
			End If
		Else
			''안전인증정보
			If (Fsafetyyn="Y" and FsafetyDiv<>0) Then
				anjunInfo = anjunInfo & "&sft_cert_tgt_yn=Y"
				If (FsafetyDiv=10) Then
					anjunInfo = anjunInfo & "&kps_1_no="&Server.URLEncode(FsafetyNum)
				Elseif (FsafetyDiv=20) Then
					anjunInfo = anjunInfo & "&kps_1_no="
					anjunInfo = anjunInfo & "&kps_2_no="&Server.URLEncode(FsafetyNum)
				Elseif (FsafetyDiv=30) Then
					anjunInfo = anjunInfo & "&kps_1_no="
					anjunInfo = anjunInfo & "&kps_3_no="&Server.URLEncode(FsafetyNum)
				Elseif (FsafetyDiv=40) Then
					anjunInfo = anjunInfo & "&kps_1_no="
					anjunInfo = anjunInfo & "&kps_4_no="&Server.URLEncode(FsafetyNum)
				Elseif (FsafetyDiv=50) Then
					anjunInfo = anjunInfo & "&kps_1_no="
					anjunInfo = anjunInfo & "&kps_5_no="&Server.URLEncode(FsafetyNum)
				Else
					anjunInfo = ""
				End if
			End If
		End If

'		strSql = ""
'		strSql = strSql & " SELECT top 100 M.* " & vbcrlf
'		strSql = strSql & " ,isNULL(CASE WHEN M.infocd='00000' then 'N' " '''db_item.dbo.[fn_LotteCom_SaftyFormat](IC.safetyyn,IC.safetyDiv,IC.safetyNum) " & vbcrlf
'		strSql = strSql & " 	  WHEN M.infocd='99999' then M.infoETC"& vbcrlf
'		strSql = strSql & " 	  WHEN M.infocd='00005' then '해당없음'"& vbcrlf
'		strSql = strSql & " 	  WHEN M.infocd='00006' then '상품 상세 참고'"& vbcrlf
'		strSql = strSql & " 	  WHEN c.infotype='C' and F.chkDiv='N' THEN '해당없음' " & vbcrlf
'		strSql = strSql & " 	  WHEN c.infotype='P' THEN replace(c.infoDesc,'1644-6030','1644-6035') " & vbcrlf
'		'2014-07-14 16:07 김진영 하단 추가. 조아름 요청 "품질보증기준" 앞에 텍스트 문구 추가
'		strSql = strSql & " 	  WHEN (c.infoItemName= '품질보증기준') and (isnull(ET.itemid, '') <> '') THEN '품질보증기준 관련법 및 소비자 분쟁해결 기준에 따름, ' + F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
'		strSql = strSql & "  ELSE F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
'		strSql = strSql & "  END,'') as infoCDVal " & vbcrlf
'		strSql = strSql & "  , L.shortVal, ET.itemid as ETText " & vbcrlf
'		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
'		strSql = strSql & " Join db_item.dbo.tbl_item_contents IC " & vbcrlf
'		strSql = strSql & " on IC.infoDiv=M.mallinfoDiv " & vbcrlf
'		strSql = strSql & " left Join db_item.dbo.tbl_item_infoCode c " & vbcrlf
'		strSql = strSql & " on M.infocd=c.infocd " & vbcrlf
'		strSql = strSql & " Left Join db_item.dbo.tbl_item_infoCont F " & vbcrlf
'		strSql = strSql & " on M.infocd=F.infocd and F.itemid=" & FItemid & vbcrlf
'		strSql = strSql & " left join db_item.dbo.tbl_item_infoCont F2 " & vbcrlf
'		strSql = strSql & " on M.infocdAdd=F2.infocd and F2.itemid=" & FItemid & vbcrlf
'		strSql = strSql & " left join db_item.dbo.tbl_OutMall_etcLink as L " & vbcrlf
'		strSql = strSql & " on ((L.mallid = M.mallid) OR (isnull(L.mallid,'') = '')) and L.itemid =" & FItemid & vbcrlf
'		strSql = strSql & " and L.linkgbn='infoDiv21Lotte'" & vbcrlf
'		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_Outmall_etcTextItem as ET on IC.itemid = ET.itemid and ET.mallid = '"&CMALLNAME&"' " & vbcrlf
'		strSql = strSql & " where M.mallid = '"&CMALLNAME&"' AND IC.itemid=" & FItemid
'		strSql = strSql & " order by M.mallinfoCd"

		strSql = ""
		strSql = strSql & " SELECT top 100 M.* " & vbcrlf
		strSql = strSql & " ,isNULL(CASE WHEN M.infocd='99999' then M.infoETC"& vbcrlf
		strSql = strSql & " 	  WHEN M.infocd='00005' then '해당없음'"& vbcrlf
		strSql = strSql & " 	  WHEN M.infocd='00006' then '상품 상세 참고'"& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0106' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0507' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0701' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0703' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0801' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0803' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0901' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0903' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1001' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1003' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1007' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1101' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1107' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1201' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1206' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1210' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1301' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1306' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1401' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1407' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1409' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1501' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1601' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1602' and F.chkDiv='Y' THEN 'Y'+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1602' and F.chkDiv='N' THEN 'N'+'//번호없음'" & vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1604' and F.chkDiv='Y' THEN F.infocontent+'//'+F.infocontent " & vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1604' and F.chkDiv='N' THEN '해당없음//해당없음' " & vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1608' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1701' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1801' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1803' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1805' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1901' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2001' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2002' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2004' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2008' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2102' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2103' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2104' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2202' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2203' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2204' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2208' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2209' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2301' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2303' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2306' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2310' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2401' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2501' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2502' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2607' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='3501' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='3504' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='C' and F.chkDiv='N' THEN '"&trim(html2db(Fmakername))&"//해당없음'" & vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='C' and F.chkDiv='Y' THEN '"&trim(html2db(Fmakername))& "//" &trim(html2db(Fmakername))&"'" & vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='P' THEN '텐바이텐 고객행복센터//1644-6035' " & vbcrlf
		strSql = strSql & " 	  WHEN (c.infoItemName= '품질보증기준') and (isnull(ET.itemid, '') <> '') THEN '품질보증기준 관련법 및 소비자 분쟁해결 기준에 따름, ' + F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
		strSql = strSql & " 	  WHEN LEN(F.infocontent + isNULL(F2.infocontent,'')) < 2 THEN '상품 상세 참고' " & vbcrlf
		strSql = strSql & "  ELSE F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
		strSql = strSql & "  END,'') as infoCDVal " & vbcrlf
		strSql = strSql & "  , L.shortVal, ET.itemid as ETText " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents IC on IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c on M.infocd=c.infocd " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F on M.infocd=F.infocd and F.itemid=" & FItemid & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd=F2.infocd and F2.itemid=" & FItemid & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_OutMall_etcLink as L on ((L.mallid = M.mallid) OR (isnull(L.mallid,'') = '')) and L.itemid =" & FItemid & vbcrlf
		strSql = strSql & " and L.linkgbn='infoDiv21Lotte'" & vbcrlf
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_Outmall_etcTextItem as ET on IC.itemid = ET.itemid and ET.mallid = '"&CMALLNAME&"' " & vbcrlf
		strSql = strSql & " WHERE M.mallid = '"&CMALLNAME&"' AND IC.itemid=" & FItemid
		strSql = strSql & " and M.infoCd not in ('00000', '99999') "
		strSql = strSql & " ORDER BY M.mallinfoCd"
'response.write strSql
'response.end
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		Dim mat_name, mat_percent, mat_place, material

		psourceArea = ""
		If Not(rsget.EOF or rsget.BOF) then
			mallinfoDiv = "&ec_goods_artc_cd="&Server.URLEncode(rsget("mallinfoDiv"))
			Do until rsget.EOF
			    mallinfoCd = rsget("mallinfoCd")
			    infoCDVal  = rsget("infoCDVal")

			    IF mallinfoCd="2105" OR mallinfoCd="2003" Then
			    	If isNull(rsget("shortVal")) = FALSE Then
						material = Split(rsget("shortVal"),"!!^^")
						mat_name	= material(0)
						mat_percent	= material(1)
						mat_place	= material(2)

				        bufTxt = "&mmtr_nm="&Server.URLEncode(""&mat_name&"")
				        bufTxt = bufTxt&"&cmps_rt="&Server.URLEncode(""&mat_percent&"")
				        bufTxt = bufTxt&"&mmtr_orpl_nm="&Server.URLEncode(""&mat_place&"")
			    	End If
			    Else
			        if (mallinfoCd="3503") then
			            '' 품목(35 기타인경우) 원산지값 저장
			            psourceArea = rsget("infoCDVal")
			        end if

			        if (mallinfoCd="3504") then  ''국산인경우 해당없음으로 넣어달라고 함. 수입인경우만 표시 // 2014-07-14 14:30 3504일 때 "브랜드네임,해당없음" 유미희 처리요청
			            if (isMadeInKorea(psourceArea)) then
			                infoCDVal = Fmakername & "//해당없음"
    			        end if
			        end if
        		    mallinfoCdAll = mallinfoCdAll & "&"&mallinfoCd&"=" &Server.URLEncode(infoCDVal)
    			End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		strRst = anjunInfo & mallinfoDiv & mallinfoCdAll & bufTxt
		getLotteItemInfoCdToReg = strRst
	End Function

    Public Function getLotteOptionParamToEdit()
		Dim ret : ret = ""
		Dim i
		Dim strSql, arrRows, iErrStr
		Dim isOptionExists, availOptCnt
		Dim mayOptionCnt : mayOptionCnt = 0
		Dim item_sale_stat_cd,outmalloptcode, optLimit
		Dim item_noStr, item_sale_stat_cdStr, inv_qtyStr
		Dim optValidExists : optValidExists = false
		Dim preMaxOutmalloptcode : preMaxOutmalloptcode=-1

		availOptCnt = 0
		strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_Lotte '"&CMallName&"'," & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) then
			arrRows = rsget.getRows
		End If
		rsget.close

		isOptionExists = isArray(arrRows)
		If (isOptionExists) Then
			mayOptionCnt = UBound(ArrRows,2)
			mayOptionCnt = mayOptionCnt + 1
		End If

		ret = ""
		If (Not isOptionExists) Then
			If (FLimitYn = "Y") Then
				ret = ret & "&inv_mgmt_yn=Y"
				ret = ret & "&item_no=0"
				ret = ret & "&item_sale_stat_cd=10"
				ret = ret & "&inv_qty="&getLimitLotteEa()			'재고관리 할 경우 단품이 있어야..
			Else
				''2013/06/12 재고관리여부 모두 Y로 변경
				ret = ret & "&inv_mgmt_yn=Y" 						'재고관리여부
				ret = ret & "&item_no=0"							''(추가/아래라인까지..)2013-07-16 김진영 이 항목 필수..없으니 재고수정이 안 됨
				ret = ret & "&item_sale_stat_cd=10"
				ret = ret & "&inv_qty="&CDEFALUT_STOCK				''2013/06/12추가
			End If
		Else
			If FLimitYn = "Y" Then
				ret = ret&"&inv_mgmt_yn=Y"
			Else
				ret = ret&"&inv_mgmt_yn=Y"							''2013/06/12 수정
			End If

			For i = 0 To UBound(ArrRows,2)
				item_sale_stat_cd = "10"							''10:판매진행,20:품절,30:판매종료
				outmalloptcode = ArrRows(2,i)
				If IsNULL(outmalloptcode) then
					outmalloptcode=preMaxOutmalloptcode+1
				Else
					If (preMaxOutmalloptcode > outmalloptcode) Then
						preMaxOutmalloptcode = preMaxOutmalloptcode
					Else
						preMaxOutmalloptcode = outmalloptcode
					End if
				End If

				If FLimitYn="Y" then
					optLimit = ArrRows(4,i) - 5
					'2013-10-10 진영 추가..한정 중 1000개가 넘으면 최대 999로 넣겠다고 롯데에서 요청메일 옴
					If optLimit > 1000 Then
						optLimit = 999
					End If
				Else
					optLimit = CDEFALUT_STOCK
				End If
				If (optLimit < 1) Then optLimit = 0
				If (ArrRows(6,i) = "N") or (ArrRows(7,i)="N") Then item_sale_stat_cd="20"
				If (FLimitYn = "Y") and (optLimit < 1) Then item_sale_stat_cd="20"

				If ((ArrRows(11,i) = "1") and (ArrRows(12,i) = "1")) or (ArrRows(13,i) = "1") Then
					optLimit = 0
					item_sale_stat_cd = "20"
				End If

				item_noStr = item_noStr & "&item_no="&outmalloptcode
				item_sale_stat_cdStr = item_sale_stat_cdStr & "&item_sale_stat_cd="&item_sale_stat_cd
				inv_qtyStr = inv_qtyStr & "&inv_qty="&optLimit

				'If (item_sale_stat_cd = "10") Then optValidExists=TRUE
				If (item_sale_stat_cd = "10") Then
					'optValidExists=TRUE
					availOptCnt = availOptCnt + 1
				End If
            Next
            ret = ret&item_noStr&item_sale_stat_cdStr&inv_qtyStr
        End If
		''rw ret
		If (Not isOptionExists) Then   ''옵션이 없으면.
			If getLotteSellYn = "Y" Then										 																'판매상태			(*:10:판매,20:품절)
				ret = ret & "&sale_stat_cd=10"
			Else
				FSellyn = "N"
				ret = ret & "&sale_stat_cd=20"
			End If
		Else
			'If (optValidExists) and (getLotteSellYn = "Y") Then  ''판매중 이고 옵션 판매가능이면.
			If (availOptCnt > 0) and (getLotteSellYn = "Y") Then  ''판매중 이고 옵션 판매가능이면.
				ret = ret & "&sale_stat_cd=10"
			Else
				'rw "None Exists Valid Option"
				FSellyn="N"
				ret = ret & "&sale_stat_cd=20"
			End if
		End If
		getLotteOptionParamToEdit = ret
	End Function

	'// 상품등록 파라메터 생성
	public Function getLotteComItemRegParameter(isEdit)
		dim strRst
		strRst = "subscriptionId=" & lotteAuthNo																						'롯데닷컴 인증번호	(*)
		if (isEdit) then
		   strRst = strRst & "&goods_req_no="&FLotteTmpGoodNo
		end if
		strRst = strRst & "&brnd_no=" & tenBrandCd																						'브랜드코드			(*)
		strRst = strRst & "&goods_nm=" & Server.URLEncode(Trim(getItemNameFormat))
		strRst = strRst & getItemKeyword												'키워드 //2013-08-09 김진영 수정
		strRst = strRst & "&pmct_fix_cd=2"																					 			'원가고정여부		(*:마진율고정)
		strRst = strRst & "&pur_shp_cd=3" 		'' 2=>3 특정																			'매출형태			(*:판매분매입)
		strRst = strRst & "&sale_shp_cd=10" 																							'판매형태			(*:정상)
		strRst = strRst & "&sale_prc=" & cLng(GetRaiseValue(FSellCash/10)*10)															'판매가(원판매가)	(*) 소비자가로 설정한 이유..?
		strRst = strRst & "&mrgn_rt=12" 																								'마진율				(*:11%) ==> 2013/01/01 12%
		strRst = strRst & "&pur_prc=" & cLng(FSellCash*0.88)																			'공급가				(*)
		strRst = strRst & "&tdf_sct_cd=1" 																								'과면세코드			(*:과세)
		strRst = strRst & "&card_dsct_yn=Y"																					 			'롯데카드할인여부	(*:사용)
		if getLotteSellYn="Y" then										 																'판매상태			(*:10:판매,20:품절)
			strRst = strRst & "&sale_stat_cd=10"
		else
			strRst = strRst & "&sale_stat_cd=20"
		end if

		IF (FLimitYn="Y") then
		    strRst = strRst & "&inv_mgmt_yn=Y"
		    if FoptionCnt=0 then
		        strRst = strRst & "&inv_qty="&getLimitLotteEa()
		    end if
		ELSE
		    ''2013/06/12 재고관리여부 모두 Y로 변경
    		strRst = strRst & "&inv_mgmt_yn=Y" 																							'재고관리여부		(*:관리안함)
    		if FoptionCnt=0 then
    		    strRst = strRst & "&inv_qty="&CDEFALUT_STOCK
    		end if
    	END IF

		if FitemDiv="06" then
			strRst = strRst & "&add_choc_tp_cd_20=" & Server.URLEncode("주문제작상품") 													'문구입력형옵션
		end if
		strRst = strRst & "&dlv_proc_tp_cd=1" 																							'배송유형			(*:직송)
		strRst = strRst & "&box_pkg_yn=Y" 																								'배송Box포장		(*:포장가능)
		strRst = strRst & "&fut_msg_yn=N" 																								'선물포장여부		(*:불가)
		strRst = strRst & "&shop_fwd_msg_yn=N"	 																						'매장전달			(*:사용안함)
		strRst = strRst & "&dlv_mean_cd=10" 																							'배송수단			(*:택배)

    	strRst = strRst & getLotteGoodDLVDtParams
		strRst = strRst & "&dlvp_stn_grp_cd=01" 																						'배송가능지역		(*:전국)
		strRst = strRst & "&byr_age_lmt_cd="&Chkiif(IsAdultItem() = "Y", "19", "0")&"" 													'구입자나이제한		(*:0:전체, 12:12세이상, 15:15세이상, 19:19세이상)
		strRst = strRst & "&exch_rtgs_sct_cd=21" 																						'교환반품구분		(*:분담교환)
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd																					'배송정책번호		(*:전체)
		strRst = strRst & "&corp_dlvp_sn=" & getLotteReturnDLVCode 																			 			'반품배송지코드		(*:도봉동물류센터)
		strRst = strRst & "&dcom_asgn_rtgs_hdc_use_yn=Y"                                                                 ''반품교환 지정택배 사용여부  dcom_asgn_rtgs_hdc_use_yn
		strRst = strRst & "&orpl_nm=" & Server.URLEncode(chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea))	'원산지				(*)
		strRst = strRst & "&mfcp_nm=" & Server.URLEncode(chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername))		'제조사				(*)
		strRst = strRst & "&img_url=" & Server.URLEncode(FbasicImage)																	'대표이미지URL		(*)
		'strRst = strRst & "&img_url=" & Server.URLEncode(FbasicImage&"/10x10/thumbnail/600x600/quality/85/")							'대표이미지URL		(*)
		strRst = strRst & "&attd_mtr_cont=" & Server.URLEncode(ForderComment)															'주문시 주의사항(-)
		strRst = strRst & "&brnd_intro_cont=" & server.URLEncode("Design Your Life! 새로운 일상을 만드는 감성생활브랜드 텐바이텐")		'브랜드 설명
		strRst = strRst & "&corp_goods_no=" & Fitemid																					'업체상품번호(텐텐상품코드)
		strRst = strRst & getLotteCateParamToReg()																						'MD상품군 및 해당 전시카테고리
		strRst = strRst & getLotteOptionParamToReg()																					'상품옵션 여부 및 옵션내용
		strRst = strRst & getLotteItemContParamToReg()																					'상품상세설명		(*)
		strRst = strRst & getLotteAddImageParamToReg()																					'상품 추가 이미지
        strRst = strRst & getLotteItemInfoCdToReg()   ''진영
		'결과 반환
		getLotteComItemRegParameter = strRst
	end Function

	public Function getLotteComItemEditParameter2()
		Dim strRst
		strRst = getLotteComItemRegParameter(true)
		getLotteComItemEditParameter2 = strRst
	end function

	'// 상품수정 파라메터 생성
	Public Function getLotteComItemEditParameter()
		dim strRst
		strRst = "subscriptionId=" & lotteAuthNo																						'롯데닷컴 인증번호	(*)
		strRst = strRst & "&goods_no=" & FLotteGoodNo																					'롯뎃닷컴 상품번호	(*)
		strRst = strRst & "&brnd_no=" & tenBrandCd																						'브랜드코드			(*)
		strRst = strRst & "&goods_nm=" & Server.URLEncode(Trim(getItemNameFormat))				'상품명				(*)
		strRst = strRst & getItemKeyword																								'키워드 //2013-08-09 김진영 수정
		strRst = strRst & "&pur_shp_cd=2" 																								'매출형태			(*:판매분매입)

		strRst = strRst & "&dlv_proc_tp_cd=1" 																							'배송유형			(*:직송)
		strRst = strRst & "&box_pkg_yn=Y" 																								'배송Box포장		(*:포장가능)
		strRst = strRst & "&fut_msg_yn=N" 																								'선물포장여부		(*:불가)
		strRst = strRst & "&shop_fwd_msg_yn=N"	 																						'매장전달			(*:사용안함)
		strRst = strRst & "&dlv_mean_cd=10" 																							'배송수단			(*:택배)

    	strRst = strRst & getLotteGoodDLVDtParams
		strRst = strRst & "&dlvp_stn_grp_cd=01" 																						'배송가능지역		(*:전국)
		strRst = strRst & "&byr_age_lmt_cd="&Chkiif(IsAdultItem() = "Y", "19", "0")&"" 													'구입자나이제한		(*:0:전체, 12:12세이상, 15:15세이상, 19:19세이상)
		strRst = strRst & "&exch_rtgs_sct_cd=21" 																						'교환반품구분		(*:분담교환)
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd																					'배송정책번호		(*:전체)
		strRst = strRst & "&corp_dlvp_sn=" & getLotteReturnDLVCode 																			 			'반품배송지코드		(*:도봉동물류센터)
		strRst = strRst & "&dcom_asgn_rtgs_hdc_use_yn=Y"                                                                                               ''반품교환 지정택배 사용여부 2013/02/26
		strRst = strRst & "&orpl_nm=" & Server.URLEncode(chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea))	'원산지				(*)
		strRst = strRst & "&mfcp_nm=" & Server.URLEncode(chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername))		'제조사				(*)
'		strRst = strRst & "&img_url=" & Server.URLEncode(FbasicImage)		'대표이미지URL		(*)	// 2018-12-14 김진영 주석처리..이미지API신규생성
		'strRst = strRst & "&img_url=" & Server.URLEncode(FbasicImage&"/10x10/thumbnail/600x600/quality/85")							'대표이미지URL		(*)
		strRst = strRst & "&brnd_intro_cont=" & server.URLEncode("Design Your Life! 새로운 일상을 만드는 감성생활브랜드 텐바이텐")		'브랜드 설명
		if (Fitemid="409536") then
		    strRst = strRst & "&attd_mtr_cont="
		else
    		strRst = strRst & "&attd_mtr_cont=" & Server.URLEncode(ForderComment)															'주문시 주의사항
    	end if
		strRst = strRst & "&corp_goods_no=" & Fitemid																					'업체상품번호(텐텐상품코드)
		strRst = strRst & getLotteItemContParamToReg()																					'상품상세설명		(*)
'		strRst = strRst & getLotteAddImageParamToReg()						'상품 추가 이미지	// 2018-12-14 김진영 주석처리..이미지API신규생성
        strRst = strRst & getLotteOptionParamToEdit()
		strRst = strRst & getLotteItemInfoCdToReg()																						'상품품목정보/2012-11-02김진영 생성
		'결과 반환
		getLotteComItemEditParameter = strRst
	end Function

	'2014/12/15 김진영 생성 상품품목정보수정 파라메타
	Public Function getLotteItemInfoCdToEdt()
		Dim strRst, strSQL, strRst2
		Dim anjunInfo, mallinfoCdAll,mallinfoCd, infoCDVal, psourceArea
        Dim bufTxt : bufTxt=""

		strSql = ""
		strSql = strSql & " SELECT top 100 M.* " & vbcrlf
		strSql = strSql & " ,isNULL(CASE WHEN M.infocd='99999' then M.infoETC"& vbcrlf
		strSql = strSql & " 	  WHEN M.infocd='00005' then '해당없음'"& vbcrlf
		strSql = strSql & " 	  WHEN M.infocd='00006' then '상품 상세 참고'"& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0106' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0507' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0701' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0703' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0801' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0803' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0901' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0903' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1001' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1003' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1007' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1101' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1107' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1201' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1206' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1210' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1301' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1306' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1401' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1407' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1409' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1501' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1601' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1602' and F.chkDiv='Y' THEN 'Y'+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1602' and F.chkDiv='N' THEN 'N'+'//번호없음'" & vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1604' and F.chkDiv='Y' THEN F.infocontent+'//'+F.infocontent " & vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1604' and F.chkDiv='N' THEN '해당없음//해당없음' " & vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1608' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1701' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1801' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1803' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1805' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1901' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2001' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2002' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2004' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2008' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2102' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2103' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2104' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2202' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2203' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2204' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2208' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2209' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2301' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2303' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2306' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2310' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2401' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2501' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2502' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2607' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='3501' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='3504' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='C' and F.chkDiv='N' THEN '"&trim(Fmakername)&"//해당없음'" & vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='C' and F.chkDiv='Y' THEN '"&trim(Fmakername)& "//" &trim(Fmakername)&"'" & vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='P' THEN '텐바이텐 고객행복센터//1644-6035' " & vbcrlf
		strSql = strSql & " 	  WHEN (c.infoItemName= '품질보증기준') and (isnull(ET.itemid, '') <> '') THEN '품질보증기준 관련법 및 소비자 분쟁해결 기준에 따름, ' + F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
		strSql = strSql & " 	  WHEN LEN(F.infocontent + isNULL(F2.infocontent,'')) < 2 THEN '상품 상세 참고' " & vbcrlf
		strSql = strSql & "  ELSE F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
		strSql = strSql & "  END,'') as infoCDVal " & vbcrlf
		strSql = strSql & "  , L.shortVal, ET.itemid as ETText " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents IC on IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c on M.infocd=c.infocd " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F on M.infocd=F.infocd and F.itemid=" & FItemid & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd=F2.infocd and F2.itemid=" & FItemid & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_OutMall_etcLink as L on ((L.mallid = M.mallid) OR (isnull(L.mallid,'') = '')) and L.itemid =" & FItemid & vbcrlf
		strSql = strSql & " and L.linkgbn='infoDiv21Lotte'" & vbcrlf
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_Outmall_etcTextItem as ET on IC.itemid = ET.itemid and ET.mallid = '"&CMALLNAME&"' " & vbcrlf
		strSql = strSql & " WHERE M.mallid = '"&CMALLNAME&"' AND IC.itemid=" & FItemid
		strSql = strSql & " and M.infoCd not in ('00000', '99999') "
		strSql = strSql & " ORDER BY M.mallinfoCd"
'		rw strSql
'		response.end
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		Dim mat_name, mat_percent, mat_place, material
		psourceArea = ""
		If Not(rsget.EOF or rsget.BOF) then
			strRst2 = "&ec_goods_artc_cd="&Server.URLEncode(rsget("mallinfoDiv"))
			Do until rsget.EOF
			    mallinfoCd = rsget("mallinfoCd")
			    infoCDVal  = rsget("infoCDVal")

			    IF mallinfoCd="2105" OR mallinfoCd="2003" Then
			    	If isNull(rsget("shortVal")) = FALSE Then
						material = Split(rsget("shortVal"),"!!^^")
						mat_name	= material(0)
						mat_percent	= material(1)
						mat_place	= material(2)

				        bufTxt = "&mmtr_nm="&Server.URLEncode(""&mat_name&"")
				        bufTxt = bufTxt&"&cmps_rt="&Server.URLEncode(""&mat_percent&"")
				        bufTxt = bufTxt&"&mmtr_orpl_nm="&Server.URLEncode(""&mat_place&"")
			    	End If
			    Else
			        if (mallinfoCd="3503") then
			            '' 품목(35 기타인경우) 원산지값 저장
			            psourceArea = rsget("infoCDVal")
			        end if

			        if (mallinfoCd="3504") then  ''국산인경우 해당없음으로 넣어달라고 함. 수입인경우만 표시 // 2014-07-14 14:30 3504일 때 "브랜드네임,해당없음" 유미희 처리요청
			            if (isMadeInKorea(psourceArea)) then
			                infoCDVal = Fmakername & "//해당없음"
    			        end if
			        end if
        		    mallinfoCdAll = mallinfoCdAll & "&"&mallinfoCd&"=" &Server.URLEncode(infoCDVal)
    			End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		strRst = "subscriptionId=" & lotteAuthNo
		strRst = strRst & "&goods_no=" & FLotteGoodNo
		strRst2 = strRst2 & "&chg_caus_cont=" & Server.URLEncode("api 상품품목 변경")
		getLotteItemInfoCdToEdt = strRst & bufTxt & strRst2 & mallinfoCdAll
	End Function

	'2018/12/14 김진영 생성 이미지 수정 파라메타
	Public Function getLotteItemImageEdit
		Dim strRst
		strRst = "subscriptionId=" & lotteAuthNo								'롯데닷컴 인증번호	(*)
		strRst = strRst & "&goods_no=" & FLotteGoodNo							'롯데닷컴 상품번호	(*)
		strRst = strRst & "&img_url=" & Server.URLEncode(FbasicImage)			'대표이미지URL		(*)
		strRst = strRst & getLotteAddImageParamToReg()							'상품 추가 이미지
		getLotteItemImageEdit = strRst
	End Function
End Class

Class CLotte
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

	Public Sub getLotteNotRegOneItem
		Dim strSql, addSql, i
		if FRectItemID<>"" then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"

			'''2013-07-25 김진영 옵션 추가금액 있는경우, 옵션금액 팝업에서 설정한 것만
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & " select itemid from ("
            addSql = addSql & "     select o.itemid"
            addSql = addSql & " 	,count(*) as optCNT"
            addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	from db_item.dbo.tbl_item_option as o "
            addSql = addSql & " 	left join db_item.dbo.tbl_lotte_regItem as RR on o.itemid = RR.itemid and RR.itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	where o.itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and o.isusing='Y'"
'            addSql = addSql & " 	and isnull(RR.optAddPrcRegType,'') = '0'"		'2016-11-15 김진영 주석처리
            addSql = addSql & " 	group by o.itemid"
            addSql = addSql & " ) T"
            addSql = addSql & " where optAddCNT>0"
            addSql = addSql & " or (optCnt-optNotSellCnt<1)"
            addSql = addSql & " )"
		end if

		If FRectItemID = "1401874" Then
			addSql = addSql & "or i.itemid = '1401874' "
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay, UC.socname_kor "
		strSql = strSql & "	, isNULL(R.lotteStatcd,-9) as lotteStatcd "
		strSql = strSql & "	, C.infoDiv,isNULL(C.safetyyn,'N') as safetyyn,isNULL(C.safetyDiv,0) as safetyDiv,C.safetyNum " '' 품목정보 및 안전인증정보. 추가 20121102
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotte_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " JOIN db_user.dbo.tbl_user_c UC on i.makerid = UC.userid"
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_lotte_regItem R on i.itemid=R.itemid"
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "																				'플라워/화물배송 상품 제외
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and i.itemdiv <> '21' "
		strSql = strSql & "	and UC.isExtUsing <> 'N'"
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
'		strSql = strSql & " and (i.sellcash<>0 and ((i.sellcash-i.buycash)/i.sellcash)*100>=" & CMAXMARGIN & ")"
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'등록제외 카테고리
		strSql = strSql & "	and i.itemid not in (Select itemid From db_item.dbo.tbl_lotte_regItem  where lottestatCD not in ('00','10') ) "				'롯데등록상품 제외  ,'10' -- 오류
		strSql = strSql & "	and cm.mapCnt is Not Null "	& addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CLotteItem
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
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
                FOneItem.FLotteStatcd     	= rsget("lotteStatcd")

                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.Fsafetyyn			= rsget("safetyyn")
                FOneItem.FsafetyDiv			= rsget("safetyDiv")
                FOneItem.FsafetyNum			= rsget("safetyNum")
				FOneItem.Fsocname_kor		= rsget("socname_kor")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub

	Public Sub getLotteEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

        ''//연동 제외상품
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt<getdate()"
        addSql = addSql & "     and edDt>getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay, uc.socname_kor "
		strSql = strSql & "	, m.LotteGoodNo, m.LotteTmpGoodNo, m.LotteSellYn, isNULL(m.regedOptCnt,0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, m.Lotteprice, m.regImagename "
		strSql = strSql & "	, C.infoDiv,isNULL(C.safetyyn,'N') as safetyyn,isNULL(C.safetyDiv,0) as safetyDiv,C.safetyNum "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or ((i.sailyn='N') and (i.deliveryType=9) and (i.sellcash<10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or m.optAddPrcCnt > 0"
		strSql = strSql & "		or i.itemdiv = '21'"
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut"
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " join db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " join db_item.dbo.tbl_lotte_regItem as m on i.itemid=m.itemid "
		strSql = strSql & " left Join (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotte_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " left join db_user.dbo.tbl_user_c uc on i.makerid=uc.userid"
		strSql = strSql & " Where 1=1"
		strSql = strSql & addSql
		strSql = strSql & " and isNULL(m.LotteTmpGoodNo,m.LotteGoodNo) is Not Null "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CLotteItem
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
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FregImageName 		= rsget("regImagename")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FLotteGoodNo		= rsget("LotteGoodNo")
				FOneItem.FLotteTmpGoodNo	= rsget("LotteTmpGoodNo")
				FOneItem.FLotteSellYn		= rsget("LotteSellYn")
				FOneItem.FLottePrice		= rsget("Lotteprice")
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
                FOneItem.Fsocname_kor    = rsget("socname_kor")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub
End Class

Function GetRaiseValue(value)
    If Fix(value) < value Then
    GetRaiseValue = Fix(value) + 1
    Else
    GetRaiseValue = Fix(value)
    End If
End Function

'// 상품이미지 존재여부 검사
function ImageExists(byval iimg)
	if (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") then
		ImageExists = false
	else
		ImageExists = true
	end if
end function

Function getLotteGoodno(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 lotteGoodno FROM db_item.dbo.tbl_lotte_regitem WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getLotteGoodno = rsget("lotteGoodno")
	End If
	rsget.Close
End Function
%>
