<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "halfclub"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST APIURL = "http://api.tricycle.co.kr"
CONST UPCHECODE = "A5703"								'업체코드
CONST APIKEY = "B6D75816-1F35-4450-8B9B-71137B9212F9"	'API KEY
CONST CDEFALUT_STOCK = 999

Class CHalfclubItem
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
	Public FHalfClubStatCD
	Public Fdeliverfixday
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FmaySoldOut
	Public FHalfclubGoodno
	Public Fregitemname
	Public FregImageName
	Public FbasicImageNm
	Public Fsocname_kor
	Public FDepthCode
	Public FBrandCode
	Public FNeedInfoDiv
	Public FItemweight
	Public Fcdmkey
	Public Fcddkey

	'// 품절여부
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	end function

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
		rsget.Open sqlStr,dbget,1
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

	Public Function getMatchingInfoDiv(halfClubInfoDiv)
		Dim mappingDiv
		Select Case halfClubInfoDiv
			Case "C01"		mappingDiv = "01"
			Case "C02"		mappingDiv = "02"
			Case "C03"		mappingDiv = "03"
			Case "C04"		mappingDiv = "04"
			Case "C05"		mappingDiv = "05"
			Case "C06"		mappingDiv = "06"
			Case "C07"		mappingDiv = "17"
			Case "C08"		mappingDiv = "18"
			Case "C09"		mappingDiv = "19"
			Case "C10"		mappingDiv = "23"
			Case "C11"		mappingDiv = "25"
			Case "C12"		mappingDiv = "26"
			Case "C13"		mappingDiv = "08"
			Case "C14"		mappingDiv = "21"
			Case "C20"		mappingDiv = "35"
		End Select
		If FinfoDiv = mappingDiv Then
			getMatchingInfoDiv = "Y"
		Else
			getMatchingInfoDiv = "N"
		End If
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, tmpPrice, sqlStr, specialPrice
		sqlStr = ""
		sqlStr = sqlStr & " SELECT mustPrice "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] "
		sqlStr = sqlStr & " WHERE mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " and itemid = '"& Fitemid &"' "
		sqlStr = sqlStr & " and getdate() >= startDate and getdate() <= endDate "
		rsget.Open sqlStr,dbget,1
		If Not(rsget.EOF or rsget.BOF) Then
			specialPrice = rsget("mustPrice")
		End If
		rsget.Close

		If specialPrice <> "" Then
			MustPrice = specialPrice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If (GetTenTenMargin < CMAXMARGIN) Then
				tmpPrice = Forgprice
			Else
				tmpPrice = FSellCash
			End If
			MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
		End If
	End Function

	'최대 구매 수량
	Public Function getLimitHalfClubEa()
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
		getLimitHalfClubEa = ret
	End Function

	'// 하프클럽 판매여부 반환
	Public Function gethalfclubSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				gethalfclubSellYn = "Y"
			Else
				gethalfclubSellYn = "N"
			End If
		Else
			gethalfclubSellYn = "N"
		End If
	End Function

	Public Function getItemidYear()
		If Clng(Fitemid) <= 1199999 Then
			getItemidYear = "2014"
		ElseIf Clng(Fitemid) >= 1200000 AND Clng(Fitemid) <= 1399999 Then
			getItemidYear = "2015"
		ElseIf Clng(Fitemid) >= 1400000 AND Clng(Fitemid) <= 1599999 Then
			getItemidYear = "2016"
		ElseIf Clng(Fitemid) >= 1600000 AND Clng(Fitemid) <= 1799999 Then
			getItemidYear = "2017"
		ElseIf Clng(Fitemid) >= 1800000 Then
			getItemidYear = Year(Date())
		End If
	End Function

    public function getItemNameFormat()
        dim buf
		If application("Svr_Info") = "Dev" Then
			buf = "[TEST상품] "&FItemName
		Else
			buf = "["&FBrandNameKor&"] "&FItemName
		End If
        buf = replace(buf,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","프로")
        buf = replace(buf,"&","＆")
        buf = replace(buf,"[무료배송]","")
        buf = replace(buf,"[무료 배송]","")
        buf = LeftB(buf, 100)
        getItemNameFormat = buf
    end function

    public function getOptionNameFormat(v)
        dim buf
        buf = replace(v,"&"," ")
        buf = replace(buf,"(","")
        buf = replace(buf,")","")
        buf = replace(buf,"/","")
        buf = replace(buf,"-","")
        buf = replace(buf,"+","_")
		buf = replace(buf,"[","")
		buf = replace(buf,"]","")
        getOptionNameFormat = buf
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
				rsget.Open strSql,dbget,1

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
				rsget.Open strSql,dbget,1
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
	Public Function getHalfClubContParamToReg()
		Dim strRst, strSQL,strtextVal
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '돋움','돋움' }</style>"
		strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_halfclub.jpg'></p><br>"

		If ForderComment <> "" Then
			strRst = strRst & "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
		End If

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
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%""><br>")

		'#배송 주의사항
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_halfclub.jpg"">")
		getHalfClubContParamToReg = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " WHERE mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			strtextVal = rsget("textVal")
			strRst = ""
			strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '돋움','돋움' }</style>"
			strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_halfclub.jpg'></p><br>"
			strRst = strRst & Replace(Replace(strtextVal,"",""),"","")
			strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_halfclub.jpg"">")
			getHalfClubContParamToReg = strRst
		End If
		rsget.Close
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

	Public Function getHalfClubAddImageParam()
		Dim strRst, strSQL, i, k, tmpCnt, addImgUrl
		strRst = ""
		strRst = strRst & " <ImgURL_Base>"&FbasicImage&"/10x10/thumbnail/500!x500!/quality/85/"&"</ImgURL_Base>"			'#상품 기본 큰 이미지 URL(350 이상)
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
					strRst = strRst & "	<ImgURL_Other"&i&">http://webimage.10x10.co.kr/image/"&addImgUrl&"/10x10/thumbnail/500!x500!/quality/85/</ImgURL_Other"&i&">"					'추가 이미지 1 URL
					tmpCnt = tmpCnt + 1
				End If
				rsget.MoveNext
				If i>=3 Then Exit For
			Next
		End If
'rw tmpCnt
'response.end
		If tmpCnt < 3 Then
			For k = tmpCnt + 1 to 3
				strRst = strRst & "	<ImgURL_Other"&k&" />"
			Next
		End If
		rsget.Close
		getHalfClubAddImageParam = strRst
	End Function

	Public Function getHalfClubOptParamtoREG()
		Dim strSql, strRst, vItemOption, vOptionName, vOptAddPrice, vOptLimit, i
		strRst = ""
		vOptAddPrice		= 0
		strSql = ""
		strSql = strSql & " SELECT TOP 900 i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(100),o.optionname) as optionname" & VBCRLF
		strSql = strSql & " , o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, o.optaddprice, (optlimitno-optlimitsold) as optLimit " & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_item as i " & VBCRLF
		strSql = strSql & " JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' and o.optsellyn='Y' " & VBCRLF
		strSql = strSql & " WHERE i.itemid = "&Fitemid
		strSql = strSql & " and o.optionname = (SELECT db_etcmall.[dbo].[RemoveSpecialChars](o.optionname)) "
		strSql = strSql & " ORDER BY o.itemoption ASC "
		rsget.Open strSql, dbget, 1
		If Not(rsget.EOF or rsget.BOF) Then
			strRst = strRst & "			<OptionInfo>"															'옵션 정보 시작 엘리먼트
			For i = 1 to rsget.RecordCount
				vItemOption	 		= rsget("itemoption")
				vOptionName 		= db2Html(rsget("optionname"))
				vOptAddPrice		= rsget("optaddprice")
				vOptLimit			= rsget("optLimit")
				vOptLimit			= vOptLimit - 5
				If (vOptLimit < 1) Then vOptLimit = 0
				If (FLimitYN <> "Y") Then vOptLimit = CDEFALUT_STOCK

				strRst = strRst & "				<Option>"
				strRst = strRst & "					<OptCd>"&vItemOption&"</OptCd>"								'#옵션코드
				strRst = strRst & "					<OptNm><![CDATA["&getOptionNameFormat(vOptionName)&"]]></OptNm>"								'#옵션명
				strRst = strRst & "					<OptPri>"&vOptAddPrice&"</OptPri>"							'옵션가 (신규 파라미터)
				strRst = strRst & "					<InvQty>"&vOptLimit&"</InvQty>"								'옵션 재고 수량(판매 중지 시 수량 0)
				strRst = strRst & "				</Option>"
				rsget.MoveNext
			Next
			strRst = strRst & "			</OptionInfo>"
		Else
			strRst = strRst & "			<OptionInfo>"
			strRst = strRst & "				<Option>"
			strRst = strRst & "					<OptCd>0000</OptCd>"								'#옵션코드
			strRst = strRst & "					<OptNm>단일상품</OptNm>"								'#옵션명
			strRst = strRst & "					<OptPri>0</OptPri>"									'옵션가 (신규 파라미터)
			strRst = strRst & "					<InvQty>"&getLimitHalfClubEa()&"</InvQty>"			'옵션 재고 수량(판매 중지 시 수량 0)
			strRst = strRst & "				</Option>"
			strRst = strRst & "			</OptionInfo>"
		End If
		rsget.Close
		getHalfClubOptParamtoREG = strRst
	End Function

	Public Function getHalfClubItemInfoCdParameter()
		Dim strRst
		Dim strSql, buf, isMatchInfoDiv
		Dim mallinfoCd, infoContent, mallinfodiv, vType

		isMatchInfoDiv = getMatchingInfoDiv(FNeedInfoDiv)
		If isMatchInfoDiv = "N" Then
			strSql = ""
			strSql = strSql & " SELECT TOP 100 mallinfoCd, infoContent "
			strSql = strSql & " FROM db_etcmall.[dbo].[tbl_halfclub_fakeInfoCodeMap] "
			strSql = strSql & " WHERE mallinfoDiv = '"&FNeedInfoDiv&"' "
		Else
			strSql = ""
			strSql = strSql & " SELECT top 100 M.* , "
			strSql = strSql & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND isnull(IC.safetyNum, '') <> ''  THEN IC.safetyNum "
			strSql = strSql & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND isnull(IC.safetyNum, '') = ''  THEN tr.certNum "
			strSql = strSql & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '해당없음' "
			strSql = strSql & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'I') THEN '상세설명에 표기' "
			strSql = strSql & " 	 WHEN (M.infoCd='10000') THEN '관련법 및 소비자분쟁해결기준에 따름' "
			strSql = strSql & " 	 WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035' "
			strSql = strSql & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent "
			strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M "
			strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv "
			strSql = strSql & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid "
			strSql = strSql & " LEFT JOIN ( "
			strSql = strSql & "  SELECT TOP 1 itemid, certNum FROM db_item.dbo.tbl_safetycert_tenReg where itemid = '"&FItemID&"' "
			strSql = strSql & " ) as tr on I.itemid = tr.itemid "
			strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd "
			strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"' "
			strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='"&FItemID&"' "
			strSql = strSql & " WHERE M.mallid = 'halfclub' and IC.itemid='"&FItemID&"' "
		End If

		rsget.Open strSql,dbget,1
		If Not(rsget.EOF or rsget.BOF) then
			buf = buf & " 		<NotiInfo>"
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
			    If Not (IsNULL(infoContent)) AND (infoContent <> "") Then
			    	infoContent = replaceRst(replace(infoContent, chr(31), ""))
			    	infoContent = replace(infoContent, "/", " ")
			    Else
			    	infoContent = "상세설명 참고"
				End If
				buf = buf & "			<Noti>"
				buf = buf & "				<NotiNv><![CDATA["&mallinfoCd&"]]></NotiNv>"
				buf = buf & "				<NotiValue><![CDATA["&infoContent&"]]></NotiValue>"
				buf = buf & "			</Noti>"
				rsget.MoveNext
			Loop
			buf = buf & "		</NotiInfo>"
		End If
		rsget.Close
		getHalfClubItemInfoCdParameter = buf
		'rw buf
	End Function

	Public Function getCertInfoParam()
		Dim strRst, strSql, i, arrRows, notarrRows, newCertNo, nLp, newDiv, tCode, SafeCertTarget
		Dim buf
		strSql = ""
		strSql = strSql & " SELECT TOP 5 certNum, safetyDiv " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg " & vbcrlf
		strSql = strSql & " WHERE itemid='"&FItemID&"' " & vbcrlf
		rsget.Open strSql,dbget,1
		If Not(rsget.EOF or rsget.BOF) then
			arrRows = rsget.getRows()
		Else
			notarrRows = "Y"
		End If
		rsget.Close

		If notarrRows = "" Then		'전안법 적용된 데이터라면 제대로 꼽기
			If FsafetyYn = "Y" Then
				SafeCertTarget = "RequireCert"
				For nLp =0 To UBound(arrRows,2)
			    	newDiv = ""
					Select Case arrRows(1,nLp)
						Case "10"				'전기용품 > 안전인증
							newDiv = "Electric"
							tCode = "SafeCert"
						Case "20"				'전기용품 > 안전확인 신고
							newDiv = "Electric"
							tCode = "SafeCheck"
						Case "30"				'전기용품 > 공급자 적합성 확인
							newDiv = "Electric"
							tCode = "SupplierCheck"
						Case "40"				'생활제품 > 안전인증
							newDiv = "Living"
							tCode = "SafeCert"
						Case "50"				'생활제품 > 자율안전확인
							newDiv = "Living"
							tCode = "SafeCheck"
						Case "60"				'생활제품 > 안전품질표시
							newDiv = "Living"
							tCode = "SupplierCheck"
						Case "70"				'어린이제품 > 안전인증
							newDiv = "Child"
							tCode = "SafeCert"
						Case "80"				'어린이제품 > 안전확인
							newDiv = "Child"
							tCode = "SafeCheck"
						Case "90"				'어린이제품 > 공급자 적합성 확인
							newDiv = "Child"
							tCode = "SupplierCheck"
					End Select

					newCertNo = arrRows(0,nLp)
					If newCertNo = "x" Then
						newCertNo = ""
					End If

					strRst = strRst & "	<CertInfo>"
					strRst = strRst & "		<TargetCode>"&tCode&"</TargetCode>"
					strRst = strRst & "		<SafeCertType>"&newDiv&"</SafeCertType>"
					strRst = strRst & "		<CertNum><![CDATA["&newCertNo&"]]></CertNum>"				'인증번호
					strRst = strRst & "	</CertInfo>"
				Next
			Else
				SafeCertTarget = "NotCert"
				strRst = strRst & "	<CertInfo>"
				strRst = strRst & "		<TargetCode>SafeCert</TargetCode>"
				strRst = strRst & "		<SafeCertType>NONE</SafeCertType>"
				strRst = strRst & "		<CertNum></CertNum>"				'인증번호
				strRst = strRst & "	</CertInfo>"
			End If
		Else
			If FsafetyYn = "Y" AND FSafetyNum <> "" Then
'				SafeCertTarget = "RequireCert"
'				Select Case FsafetyDiv
'					Case "10"			'[공산품] 안전인증
'						newDiv = "Living"
'						tCode = "SafeCert"
'					Case "20"			'[전기용품] 안전인증
'						newDiv = "Electric"
'						tCode = "SafeCert"
'					Case "30"			'[공산품] 안전/품질표시
'						newDiv = "Living"
'						tCode = "SupplierCheck"
'					Case "40"			'[공산품] 자율안전확인
'						newDiv = "Living"
'						tCode = "SafeCheck"
'					Case "50"			'[공산품] 어린이보호포장
'						newDiv = "Child"
'						tCode = "ProtectedPackage"
'				End Select
'				strRst = strRst & "	<CertInfo>"
'				strRst = strRst & "		<TargetCode>"&tCode&"</TargetCode>"
'				strRst = strRst & "		<SafeCertType>"&newDiv&"</SafeCertType>"
'				strRst = strRst & "		<CertNum><![CDATA["&FSafetyNum&"]]></CertNum>"				'인증번호
'				strRst = strRst & "	</CertInfo>"
				SafeCertTarget = "NotCert"
				strRst = strRst & "	<CertInfo>"
				strRst = strRst & "		<TargetCode>SafeCert</TargetCode>"
				strRst = strRst & "		<SafeCertType>NONE</SafeCertType>"
				strRst = strRst & "		<CertNum></CertNum>"				'인증번호
				strRst = strRst & "	</CertInfo>"
			Else
				SafeCertTarget = "NotCert"
				strRst = strRst & "	<CertInfo>"
				strRst = strRst & "		<TargetCode>SafeCert</TargetCode>"
				strRst = strRst & "		<SafeCertType>NONE</SafeCertType>"
				strRst = strRst & "		<CertNum></CertNum>"				'인증번호
				strRst = strRst & "	</CertInfo>"
			End If
		End If

		buf = ""
		buf = buf & " 			<SafeCertTarget>"&SafeCertTarget&"</SafeCertTarget>"
		If SafeCertTarget <> "NotCert" Then
			buf = buf & "			<CertInfos>"
			buf = buf & strRst
			buf = buf & "			</CertInfos>"
		End If
		getCertInfoParam = buf
	End Function

	Public Function getItemweight()
		Dim itemweight
		If FItemweight <> 0 Then
			itemweight = FItemweight / 1000
			If itemweight < 0.01 Then
				itemweight = 0
			End If
		Else
			itemweight = 0
		End If
		getItemweight = itemweight
	End Function

	'기본정보 등록 XML
	Public Function getHalfClubItemRegParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
		strRst = strRst & "<soap12:Header>"
		strRst = strRst & "	<SOAPHeaderAuth xmlns=""http://api.tricycle.co.kr/"">"
		strRst = strRst & " 	<User_ID>"&UPCHECODE&"</User_ID>"
		strRst = strRst & " 	<User_PWD>"&APIKEY&"</User_PWD>"
		strRst = strRst & " </SOAPHeaderAuth>"
		strRst = strRst & "</soap12:Header>"
		strRst = strRst & "<soap12:Body>"
		strRst = strRst & "	<Set_GoodsRegister xmlns=""http://api.tricycle.co.kr/"">"
		strRst = strRst & "		<req_Goods>"
		strRst = strRst & "			<PCode>"&FItemid&"</PCode>"												'#상품 코드
		strRst = strRst & "			<CategoryCd></CategoryCd>"												'카테고리 코드
		strRst = strRst & "			<CategoryNm></CategoryNm>"												'카테고리 명
		strRst = strRst & "			<BrdCd>"&FBrandCode&"</BrdCd>"											'#브랜드 코드(하프클럽 기준)
		strRst = strRst & "			<BrdNm><![CDATA[텐바이텐]]></BrdNm>"										'브랜드 명
		strRst = strRst & "			<Item_BCode></Item_BCode>"												'하프클럽 기준 대분류 코드
		strRst = strRst & "			<Item_BName></Item_BName>"												'하프클럽 기준 대분류 명
		strRst = strRst & "			<Item_MCode></Item_MCode>"												'하프클럽 기준 중분류 코드
		strRst = strRst & "			<Item_MName></Item_MName>"												'하프클럽 기준 중분류 명
		strRst = strRst & "			<Item_SCode>"&FDepthCode&"</Item_SCode>"								'#하프클럽 기준 소분류 코드
		strRst = strRst & "			<Item_SName></Item_SName>"												'하프클럽 기준 소분류 명
		strRst = strRst & "			<MakeYear>"&getItemidYear()&"</MakeYear>"								'#상품 제조년도
		strRst = strRst & "			<PrdNm><![CDATA["&getItemNameFormat()&"]]></PrdNm>"						'#상품명
		strRst = strRst & "			<Pri>"&Clng(GetRaiseValue(ForgPrice/10)*10)&"</Pri>"					'#상품 정상가(택가) / 쩐단위 입력안됨 수정(2018-10-25 진영)
		strRst = strRst & "			<SalPri>"&Clng(GetRaiseValue(MustPrice()/10)*10)&"</SalPri>"			'#상품 판매가 / 쩐단위 입력안됨 수정(2018-10-25 진영)
		strRst = strRst & "			<PrdDescInfo><![CDATA["&getHalfClubContParamToReg()&"]]></PrdDescInfo>"	'#상품 상세 설명
		strRst = strRst & "			<CopyInfo></CopyInfo>"													'상품 카피명
		strRst = strRst & "			<Nation><![CDATA["&Fsourcearea&"]]></Nation>"							'#상품 원산지
		strRst = strRst & getHalfClubAddImageParam()
		strRst = strRst & "			<PrdWeight>"&getItemweight()&"</PrdWeight>"								'상품 무게(단위 : kg)
		strRst = strRst & "			<SalOut>a</SalOut>"														'#상품 상태(판매중 : a, 일시품절 : b, 판매종료 : k)
		strRst = strRst & "			<ImageUpdate>Y</ImageUpdate>"											'이미지 등록 여부(Y : 등록, N : 미등록)
		strRst = strRst & getHalfClubOptParamtoREG()
		strRst = strRst & getHalfClubItemInfoCdParameter()
		strRst = strRst & getCertInfoParam()
		strRst = strRst & "			<IsConversion></IsConversion>"											'환금성 상품 여부 (1 : 환금성상품, 0 : 환금성 상품 아님)
		strRst = strRst & "		</req_Goods>"
		strRst = strRst & "	</Set_GoodsRegister>"
		strRst = strRst & "</soap12:Body>"
		strRst = strRst & "</soap12:Envelope>"
		getHalfClubItemRegParameter = strRst
'response.write replace(strRst, "UTF-8","EUC-KR")
'response.write replace(strRst, "?xml","aaaass")
'response.end
	End Function

	'상품상태 변경 XML
	Public Function getHalfClubItemEditParameter(ichgSellyn)
		Dim strRst, SalOut
		Select Case ichgSellyn
			Case "Y"	SalOut = "a"
			'Case "N"	SalOut = "b"	'일시품절
			Case "N"	SalOut = "k"	'판매종료..2019-03-18 16:29 김진영..b로 처리시 오류 발생 x 그러나 실제 품절이 되지 않는 Case 발생하여 k로 변경
		End Select

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
		strRst = strRst & "<soap12:Header>"
		strRst = strRst & "	<SOAPHeaderAuth xmlns=""http://api.tricycle.co.kr/"">"
		strRst = strRst & " 	<User_ID>"&UPCHECODE&"</User_ID>"
		strRst = strRst & " 	<User_PWD>"&APIKEY&"</User_PWD>"
		strRst = strRst & " </SOAPHeaderAuth>"
		strRst = strRst & "</soap12:Header>"
		strRst = strRst & "<soap12:Body>"
		strRst = strRst & "	<Set_GoodsRegister xmlns=""http://api.tricycle.co.kr/"">"
		strRst = strRst & "		<req_Goods>"
		strRst = strRst & "			<PCode>"&FItemid&"</PCode>"												'#상품 코드
		strRst = strRst & "			<CategoryCd></CategoryCd>"												'카테고리 코드
		strRst = strRst & "			<CategoryNm></CategoryNm>"												'카테고리 명
		strRst = strRst & "			<BrdCd>"&FBrandCode&"</BrdCd>"											'#브랜드 코드(하프클럽 기준)
		strRst = strRst & "			<BrdNm><![CDATA[텐바이텐]]></BrdNm>"										'브랜드 명
		strRst = strRst & "			<Item_BCode></Item_BCode>"												'하프클럽 기준 대분류 코드
		strRst = strRst & "			<Item_BName></Item_BName>"												'하프클럽 기준 대분류 명
		strRst = strRst & "			<Item_MCode></Item_MCode>"												'하프클럽 기준 중분류 코드
		strRst = strRst & "			<Item_MName></Item_MName>"												'하프클럽 기준 중분류 명
		strRst = strRst & "			<Item_SCode>"&FDepthCode&"</Item_SCode>"								'#하프클럽 기준 소분류 코드
		strRst = strRst & "			<Item_SName></Item_SName>"												'하프클럽 기준 소분류 명
		strRst = strRst & "			<MakeYear>"&getItemidYear()&"</MakeYear>"								'#상품 제조년도
		strRst = strRst & "			<PrdNm><![CDATA["&getItemNameFormat()&"]]></PrdNm>"						'#상품명
		strRst = strRst & "			<Pri>"&Clng(GetRaiseValue(ForgPrice/10)*10)&"</Pri>"					'#상품 정상가(택가) / 쩐단위 입력안됨 수정(2018-10-25 진영)
		strRst = strRst & "			<SalPri>"&Clng(GetRaiseValue(MustPrice()/10)*10)&"</SalPri>"			'#상품 판매가 / 쩐단위 입력안됨 수정(2018-10-25 진영)
		strRst = strRst & "			<PrdDescInfo><![CDATA["&getHalfClubContParamToReg()&"]]></PrdDescInfo>"	'#상품 상세 설명
		strRst = strRst & "			<CopyInfo></CopyInfo>"													'상품 카피명
		strRst = strRst & "			<Nation><![CDATA["&Fsourcearea&"]]></Nation>"							'#상품 원산지
		strRst = strRst & "			<PrdWeight>"&getItemweight()&"</PrdWeight>"								'상품 무게(단위 : kg)
		strRst = strRst & "			<SalOut>"&SalOut&"</SalOut>"											'#상품 상태(판매중 : a, 일시품절 : b, 판매종료 : k)
		strRst = strRst & getHalfClubAddImageParam()
		strRst = strRst & "			<ImageUpdate>"&Chkiif(isImageChanged, "Y", "N")&"</ImageUpdate>"	'이미지 등록 여부(Y : 등록, N : 미등록)
		If ichgSellyn = "N" Then
			strRst = strRst & "			<OptionInfo />"
		Else
			strRst = strRst & getHalfClubOptParamtoREG()
		End If
		strRst = strRst & getHalfClubItemInfoCdParameter()
		strRst = strRst & getCertInfoParam()
		strRst = strRst & "		</req_Goods>"
		strRst = strRst & "	</Set_GoodsRegister>"
		strRst = strRst & "</soap12:Body>"
		strRst = strRst & "</soap12:Envelope>"
		getHalfClubItemEditParameter = strRst
		' if session("ssBctID")="icommang" or session("ssBctID")="kjy8517" then
		' 	response.write replace(strRst, "UTF-8","EUC-KR")
		' 	response.write replace(strRst, "?xml","aaaass")
		' End If
'response.end
	End Function

	'상품가격 변경 XML
	Public Function getHalfClubPriceParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "<soap:Header>"
		strRst = strRst & "	<SOAPHeaderAuth xmlns=""http://api.tricycle.co.kr/"">"
		strRst = strRst & " 	<User_ID>"&UPCHECODE&"</User_ID>"
		strRst = strRst & " 	<User_PWD>"&APIKEY&"</User_PWD>"
		strRst = strRst & "	</SOAPHeaderAuth>"
		strRst = strRst & "</soap:Header>"
		strRst = strRst & "<soap:Body>"
		strRst = strRst & "	<Set_Good_Price_Change xmlns=""http://api.tricycle.co.kr/"">"
		strRst = strRst & "		<gar ResultCode="""" ResultMsg="""">"
		strRst = strRst & "			<PCode>"&FItemid&"</PCode>"
		strRst = strRst & "			<goodpriinfo>"
		strRst = strRst & "				<PCode>"&FItemid&"</PCode>"
		strRst = strRst & "				<Margin>13</Margin>"
		strRst = strRst & "				<Pri>"&Clng(GetRaiseValue(ForgPrice/10)*10)&"</Pri>"					'#상품 정상가(택가) / 쩐단위 입력안됨 수정(2018-10-25 진영)
		strRst = strRst & "				<SalPri>"&Clng(GetRaiseValue(MustPrice()/10)*10)&"</SalPri>"			'#상품 판매가 / 쩐단위 입력안됨 수정(2018-10-25 진영)
		strRst = strRst & "			</goodpriinfo>"
		strRst = strRst & "		</gar>"
		strRst = strRst & "	</Set_Good_Price_Change>"
		strRst = strRst & "</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
		getHalfClubPriceParameter = strRst
'response.write replace(strRst, "UTF-8","EUC-KR")
'response.write replace(strRst, "?xml","aaaass")
'response.end
	End Function
End Class

Class CHalfclub
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
	Public Sub getHalfClubNotRegOneItem
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
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum, c.safetyDiv, c.infodiv "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	, isNULL(R.HalfClubStatCD,-9) as HalfClubStatCD "
		strSql = strSql & "	, UC.socname_kor, am.depthCode, am.brandCode, am.needInfoDiv "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.[tbl_halfclub_cate_mapping] "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.[tbl_halfclub_cate_mapping] as am on am.tenCateLarge = i.cate_large and am.tenCateMid = i.cate_mid and am.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_halfclub_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " WHERE i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		ELSE
		    strSql = strSql & " and (i.deliveryType<>9)"
	    END IF
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.itemdiv not in ('21', '06', '08') "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "						'플라워/화물배송
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)>=" & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_etcmall.dbo.tbl_halfclub_regItem WHERE HalfClubStatCD >= 3) "	''등록완료이상은 등록안됨.	'11st등록상품 제외
		strSql = strSql & " and cm.mapCnt is Not Null "'	카테고리 매칭 상품만
		strSql = strSql & addSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHalfClubItem
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
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.FMakerName			= db2html(rsget("makername"))
				FOneItem.FBrandName			= db2html(rsget("brandname"))
				FOneItem.FBrandNameKor		= db2html(rsget("socname_kor"))
				If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
					FOneItem.FMakerName		= FOneItem.FBrandName
				End If
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetyDiv			= rsget("safetyDiv")
				FOneItem.FHalfClubStatCD		= rsget("HalfClubStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.FDepthCode			= rsget("depthCode")
				FOneItem.FBrandCode			= rsget("brandCode")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FInfodiv 			= rsget("infodiv")
				FOneItem.FNeedInfoDiv 		= rsget("needInfoDiv")
				FOneItem.FItemweight 		= rsget("itemweight")
		End If
		rsget.Close
	End Sub

	Public Sub gethalfclubEditPriceOneItem()
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHalfClubItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.ForgPrice			= rsget("orgPrice")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
		End If
		rsget.Close
	End Sub

	Public Sub gethalfclubEditOneItem()
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum, c.safetyDiv, c.infodiv "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	,isNULL(m.HalfClubStatCD,-9) as HalfClubStatCD "
		strSql = strSql & "	, UC.socname_kor, am.depthCode, am.brandCode, am.needInfoDiv "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '06', '08', '09') "
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " LEFT JOIN ( "
		strSql = strSql & " 	SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & " 	FROM db_etcmall.dbo.tbl_halfclub_cate_mapping "
		strSql = strSql & " 	GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.[tbl_halfclub_cate_mapping] as am on am.tenCateLarge = i.cate_large and am.tenCateMid = i.cate_mid and am.tenCateSmall = i.cate_small "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_halfclub_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.HalfClubGoodNo is Not Null "		'등록 상품만
		strSql = strSql & " and m.HalfclubStatCD = '7' "				'승인완료된 애들만 수정이 된다함..TEST 해봐야 됨
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHalfClubItem
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
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.FMakerName			= db2html(rsget("makername"))
				FOneItem.FBrandName			= db2html(rsget("brandname"))
				FOneItem.FBrandNameKor		= db2html(rsget("socname_kor"))
				If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
					FOneItem.FMakerName		= FOneItem.FBrandName
				End If
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetyDiv			= rsget("safetyDiv")
				FOneItem.FHalfClubStatCD		= rsget("HalfClubStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.FDepthCode			= rsget("depthCode")
				FOneItem.FBrandCode			= rsget("brandCode")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FInfodiv 			= rsget("infodiv")
				FOneItem.FNeedInfoDiv 		= rsget("needInfoDiv")
				FOneItem.FItemweight 		= rsget("itemweight")
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
