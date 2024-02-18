<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "interpark"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST interparkAPIURL = "http://ipss1.interpark.com"
CONST CDEFALUT_STOCK = 999
CONST wapiURL = "http://wapi.10x10.co.kr"

Class CInterparkitem
	Public Fitemid
	Public Fitemname
	Public FMakerid
	Public Fbuycash
	Public Fsellcash
	Public Forgsellcash
	Public Fsourcearea
	Public Foptioncnt
	Public FRegdate
	Public Fsellyn
	Public Flimityn
	Public Flimitno
	Public Flimitsold
	Public Fcate_large
	Public Fcate_mid
	Public Fcate_small
	Public FMakerName
	Public FBrandName
	Public FBrandNameKor
	Public Fkeywords
	Public Fitemoption
	Public FItemOptionTypeName
	Public FItemOptionName
	Public Fbasicimage
	Public FregImageName
	Public Fmainimage
	Public Fmainimage2
	Public FInfoImage
	Public Fordercomment
	Public FItemContent
	Public Fvatinclude
	Public Finterparkdispcategory
	Public Fitemsize
	Public Fitemsource
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold
	Public Foptaddprice
	Public FLastUpdate
	Public FSellEndDate
	Public FInfoImage1
	Public FInfoImage2
	Public FInfoImage3
	Public FInfoImage4
	Public FAddImage1
	Public FAddImage2
	Public FAddImage3
	Public FAddImage4
	Public FItemDiv
	Public Fisusing
	Public FInterparkPrdNo
	Public FmayiParkSellYn
	Public FdeliveryType
	Public FdefaultfreeBeasongLimit
	Public FSailYn
	Public FOrgPrice
	Public Finterparkregdate
	Public Fdeliverfixday
	Public Ffreight_min
	Public Ffreight_max
	Public FlastErrStr
	Public Fmayiparkprice
	Public FregOptCnt
	Public FMaySoldOut
	Public FbasicimageNm
	Public FAdultType

	Public FMayLimitSoldout
	Public FOrderMaxNum

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "999" Then
			getOrderMaxNum = 999
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
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>"&CMAXLIMITSELL&")) "
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
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>"&CMAXLIMITSELL&")) "
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

	Function GetRaiseValue(value)
		If Fix(value) < value Then
			GetRaiseValue = Fix(value) + 1
		Else
			GetRaiseValue = Fix(value)
		End If
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, tmpPrice, sqlStr, specialPrice, outmallstandardMargin, ownItemCnt
		sqlStr = ""
		sqlStr = sqlStr & " SELECT isnull(outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		sqlStr = sqlStr & " FROM db_partner.dbo.tbl_partner_addInfo "
		sqlStr = sqlStr & " WHERE partnerid = '"& CMALLNAME &"'  "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			outmallstandardMargin	= rsget("outmallstandardMargin")
		End If
		rsget.Close

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

		If specialPrice <> "" Then
			tmpPrice = specialPrice
		ElseIf ownItemCnt > 0 Then
			tmpPrice = Forgprice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If GetTenTenMargin < outmallstandardMargin Then
				tmpPrice = Forgprice
			Else
				tmpPrice = FSellCash
			End If
		End If
		MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
	End Function

	Function RightCommaDel(ostr)
		Dim restr
		restr = ""
		If IsNULL(ostr) Then Exit Function
		restr = Trim(ostr)
		If (Right(restr,1)=",") Then restr = Left(restr,Len(restr)-1)
		RightCommaDel = restr
	End Function

	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold <= CMAXLIMITSELL))
	End Function

	'// 품절여부
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	Function getiszeroWonSoldOut(iitemid, ilimityn)
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
					If (ilimityn = "Y") AND (goptlimitno - goptlimitsold > CMAXLIMITSELL) Then
						i = i + 1
					End If
					rsget.MoveNext
				Loop

				If (ilimityn = "Y") AND (i = 0) Then
					getiszeroWonSoldOut = "Y"
				ElseIf (ilimityn = "Y") AND (i > 0) Then
					getiszeroWonSoldOut = "N"
				Else
					getiszeroWonSoldOut = "N"
				End If
			Else
				getiszeroWonSoldOut = "Y"
			End If
			rsget.Close
		End If

		If getiszeroWonSoldOut = "Y" Then		'0원 품절시 상품삭제 큐에 쌓기
			sqlStr = ""
			sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_outmall_API_Que "
			sqlStr = sqlStr & " (mallid, apiAction, itemid, priority, lastUserid) "
			sqlStr = sqlStr & " VALUES ('interpark', 'DELETE', '"&iitemid&"', 10, 'system') "
			dbget.Execute sqlStr
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

	Public Function getItemNameFormat()
		Dim buf
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
		buf = "[텐바이텐] " & Replace(Replace(Replace(Replace(Replace(FBrandNameKor & " " & CStr(buf),"'",""),Chr(34),""),"<",""),">",""),"^","")
		getItemNameFormat = buf
	End Function

    Public Function GetSourcearea()
		If IsNULL(Fsourcearea) or (Fsourcearea="") then
			GetSourcearea = "."
		Else
			GetSourcearea = Fsourcearea
		End if
    End function

    Public Function GetInterParkSaleStatTp
		If (IsSoldOut) Then
			if (FSellyn = "S") then
				GetInterParkSaleStatTp = "05"       ''품절(02)     SellYN-S
			Else
				If (Fisusing = "N") Then
					GetInterParkSaleStatTp = "03"   ''판매중지
				Else
					GetInterParkSaleStatTp = "02"   ''"03"   ''판매중지(03) SellYN-N  //02로 수정 2013/09/02
				End if
			End If
		ElseIf FMaySoldout = "Y" Then
			GetInterParkSaleStatTp = "02"
		Else
			GetInterParkSaleStatTp = "01"
		End If
    End Function

	Public Function GetInterParkLmtQty()
		CONST CLIMIT_SOLDOUT_NO = 5
		''Max 99999 -> 1000
		If (Flimityn = "Y") Then
			If (Flimitno-Flimitsold) < CLIMIT_SOLDOUT_NO then
				GetInterParkLmtQty = 0
			Else
				GetInterParkLmtQty = Flimitno - Flimitsold - CLIMIT_SOLDOUT_NO
			End If
		Else
			GetInterParkLmtQty = 999
		End if
	End Function

    Public Function GetSellEndDateStr()
		GetSellEndDateStr = "99991231"
		If IsNULL(FSellEndDate) Then Exit Function
		FSellEndDate = Replace(Left(CStr(FSellEndDate),10),"-","")
    End Function

	Public Function IsTruckReturnDlvExists
		IsTruckReturnDlvExists = false
		If (FItemID = 240488) then
			IsTruckReturnDlvExists = false
			Exit Function
		End If

		If IsNULL(Ffreight_max) Then Exit Function
		If CStr(Ffreight_max = "") Then Exit Function

		IsTruckReturnDlvExists = (Fdeliverfixday="X") and (Ffreight_max>0)
	End Function

	Public Function getTruckReturnDlvPrice
		getTruckReturnDlvPrice = 0

		If (FItemID=240488) then
			getTruckReturnDlvPrice = 50000
			Exit Function
		End If

		getTruckReturnDlvPrice = CLNG(Ffreight_max*2)   '' 인터파크 프런트상 편도로 되있으나.. 현아씨 2배로?
	End Function

	Public Function getInterparkContParamToReg()
		Dim strRst, strSQL
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '돋움','돋움' }</style>"
		strRst = strRst & "<p align='center'><a href='http://www.interpark.com/display/sellerAllProduct.do?_method=main&sc.entrNo=3000010614&sc.supplyCtrtSeq=2&mid1=middle&mid2=seller&mid3=001#N_E_B_50_1_~' target='_blank'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_iPark.jpg'></a></p><br>"
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

		'#배송 주의사항
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg"">")
		getInterparkContParamToReg = strRst
	End Function

	Public Function GetInterParkentrPoint()
		GetInterParkentrPoint = CLng(Fsellcash*0.01)
		If (GetInterParkentrPoint < 10) Then GetInterParkentrPoint = 0
		If (Fsellcash < 1000) Then GetInterParkentrPoint = 0	'천원미만의상품은 아이포인트 등록이 불가합니다.
		GetInterParkentrPoint = 0	'2013/02/07 아이포인트제외
	End Function

	Public Function getInterparkOptParamtoREG
		Dim sqlStr, optLimit, itemoption
		Dim optArrRows, optlp, optlpName, optlpCode, optlpSu, optlpUsing, optlpStr, buf, optstr
		Dim ioptNameBuf, ioptCodeBuf, ioptAddPrice, ioptLimitNo, ioptTypeName

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
				If optLimit > 0 Then
			        ioptTypeName	= Replace(Replace(Trim(rsget("optionTypeName"))," ",""),"수량","갯수")
					ioptCodeBuf		= ioptCodeBuf & rsget("itemoption") & ","
					ioptNameBuf		= ioptNameBuf & Replace(Replace(Replace(Replace(Trim(rsget("optionname")),",",".")," ",""),"<","("),">",")") & ","  ''옵션내용에 공백 있으면 안됨.//선택형 옵션 데이터에 공백이
					ioptAddPrice	= ioptAddPrice & CStr(rsget("optaddprice")) & ","
					ioptLimitNo		= ioptLimitNo & CStr(optLimit) & ","
				End If
				rsget.MoveNext
			Loop
		Else
			getInterparkOptParamtoREG = ""
			rsget.Close
			Exit Function
		End If
		rsget.Close

		ioptNameBuf		= RightCommaDel(ioptNameBuf)
	    ioptCodeBuf		= RightCommaDel(ioptCodeBuf)
	    ioptAddPrice	= RightCommaDel(ioptAddPrice)
	    ioptLimitNo		= RightCommaDel(ioptLimitNo)

	    If (ioptTypeName="") then ioptTypeName="옵션명"
	    optstr = ioptTypeName & "<" & ioptNameBuf & ">"

        If (ioptLimitNo <> "") Then
            optstr = optstr & "수량<" & ioptLimitNo & ">"
        End If
        optstr = optstr & "추가금액<" & ioptAddPrice & ">"
        optstr = optstr & "옵션코드<" & ioptCodeBuf & ">"
		optstr = Replace(optstr, VbTab, "")

		If Fitemdiv = "06" Then
			buf = buf & "		<optPrirTp><![CDATA[01]]></optPrirTp>"
			buf = buf & "		<prdOption><![CDATA[{" & optstr & "}]]></prdOption>"
		Else
			buf = buf & "		<prdOption><![CDATA[" & optstr & "]]></prdOption>"
		End If
		getInterparkOptParamtoREG = buf
	End Function

	Public Function getInterparkOptParamtoEDT
		Dim sqlStr, optLimit, itemoption, limitNCnt, limitYCnt
		Dim optArrRows, optlp, optlpName, optlpCode, optlpSu, optlpUsing, optlpStr, buf, optstr
		Dim ioptNameBuf, ioptCodeBuf, ioptAddPrice, ioptLimitNo, ioptTypeName
		Dim notUbound
		Dim notCArr
		notUbound = ""
		limitYCnt = 0
		limitYCnt = 0

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 i.optioncnt, isnull(T.regedoptcnt, 0) as regedoptcnt "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_interpark_reg_Item as T on i.itemid = T.itemid "
		sqlStr = sqlStr & " WHERE i.itemid = '"&FItemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			FOptionCnt = rsget("optioncnt")
			FregOptCnt = rsget("regedoptcnt")
		End If
		rsget.Close

		buf = ""
		If Fitemdiv = "06" Then
			buf = buf & "		<optPrirTp><![CDATA[01]]></optPrirTp>"
		End If
		If FOptionCnt = 0 AND FregOptCnt > 0 Then	'현재 단품인데, 등록당시에 옵션이 있었을 경우
			sqlStr = ""
		    sqlStr = sqlStr & " SELECT itemoption, outmallOptName "
		    sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption"
		    sqlStr = sqlStr & " WHERE itemid='"&FItemID&"' "
		    sqlStr = sqlStr & " and mallid = '"&CMALLNAME&"' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		    if not rsget.Eof then
		        optArrRows = rsget.getRows()
		    Else
		    	notUbound = "Y"
		    end if
		    rsget.close

			If notUbound = "" Then
			    For optlp =0 To UBound(optArrRows,2)
			    	optlpName	= optlpName & optArrRows(1,optlp) & ","
			    	optlpCode	= optlpCode & optArrRows(0,optlp) & ","
			    	optlpSu		= optlpSu & "0,"
			    	optlpUsing	= optlpUsing & "N,"
				Next
				optlpName	= RightCommaDel(optlpName)
				optlpCode	= RightCommaDel(optlpCode)
				optlpSu		= RightCommaDel(optlpSu)
				optlpUsing	= RightCommaDel(optlpUsing)

				optlpName	= "옵션<" & optlpName & ">"
				optlpCode	= "옵션코드<" & optlpCode & ">"
				optlpSu		= "수량<" & optlpSu & ">"
				optlpUsing	= "사용여부<" & optlpUsing & ">"
				optlpStr = optlpName & optlpSu & optlpCode & optlpUsing
				optlpStr = Replace(optlpStr, VbTab, "")
				If Fitemdiv = "06" Then
					buf = buf & "		<prdOption><![CDATA[{" & optlpStr & "}]]></prdOption>"
				Else
					buf = buf & "		<prdOption><![CDATA[" & optlpStr & "]]></prdOption>"
				End If
				getInterparkOptParamtoEDT = buf
			Else
				sqlStr = ""
				sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_interpark_reg_item "
				sqlStr = sqlStr & " SET regedOptCnt = 0 "
				sqlStr = sqlStr & " WHERE itemid =" & FItemid
				dbget.Execute sqlStr
				getInterparkOptParamtoEDT = ""
			End If
		Else										'그 외
			If FOptionCnt = 0 Then
				getInterparkOptParamtoEDT = ""
			ElseIf FItemid = "1422765" Then
				If FOptionCnt <> FregOptCnt Then
					sqlStr = ""
				    sqlStr = sqlStr & " SELECT itemoption, outmallOptName "
				    sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption"
				    sqlStr = sqlStr & " WHERE itemid='"&FItemID&"' "
				    sqlStr = sqlStr & " and mallid = '"&CMALLNAME&"' "
					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				    if not rsget.Eof then
				        optArrRows = rsget.getRows()
				    Else
				    	notUbound = "Y"
				    end if
				    rsget.close

					If notUbound = "" Then
					    For optlp =0 To UBound(optArrRows,2)
					    	optlpName	= optlpName & optArrRows(1,optlp) & ","
					    	optlpCode	= optlpCode & optArrRows(0,optlp) & ","
					    	optlpSu		= optlpSu & "0,"
					    	optlpUsing	= optlpUsing & "N,"
						Next
						optlpName	= RightCommaDel(optlpName)
						optlpCode	= RightCommaDel(optlpCode)
						optlpSu		= RightCommaDel(optlpSu)
						optlpUsing	= RightCommaDel(optlpUsing)

						optlpName	= "옵션<" & optlpName & ">"
						optlpCode	= "옵션코드<" & optlpCode & ">"
						optlpSu		= "수량<" & optlpSu & ">"
						optlpUsing	= "사용여부<" & optlpUsing & ">"
						optlpStr = optlpName & optlpSu & optlpCode & optlpUsing
						optlpStr = Replace(optlpStr, VbTab, "")
'						If Fitemdiv = "06" Then
							buf = buf & "		<optPrirTp><![CDATA[01]]></optPrirTp>"
							buf = buf & "		<prdOption><![CDATA[{" & optlpStr & "}]]></prdOption>"
'						Else
'							buf = buf & "		<prdOption><![CDATA[" & optlpStr & "]]></prdOption>"
'						End If
						getInterparkOptParamtoEDT = buf
					Else
						sqlStr = ""
						sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_interpark_reg_item "
						sqlStr = sqlStr & " SET regedOptCnt = 0 "
						sqlStr = sqlStr & " WHERE itemid =" & FItemid
						dbget.Execute sqlStr
						getInterparkOptParamtoEDT = ""
					End If
				End If
			Else
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

				        ioptTypeName	= Replace(Replace(Trim(rsget("optionTypeName"))," ",""),"수량","갯수")
						ioptCodeBuf		= ioptCodeBuf & rsget("itemoption") & ","
						ioptNameBuf		= ioptNameBuf & Replace(Replace(Replace(Replace(Trim(rsget("optionname")),",",".")," ",""),"<","("),">",")") & ","  ''옵션내용에 공백 있으면 안됨.//선택형 옵션 데이터에 공백이
						ioptAddPrice	= ioptAddPrice & CStr(rsget("optaddprice")) & ","
						ioptLimitNo		= ioptLimitNo & CStr(optLimit) & ","
						If (optLimit = 0) Then
							optlpUsing	= optlpUsing & "N,"
							limitNCnt = limitNCnt + 1
						Else
							optlpUsing	= optlpUsing & "Y,"
							 limitYCnt =  limitYCnt + 1
						End If
						rsget.MoveNext
					Loop
				End If
				rsget.Close

				If FOptioncnt > 0 Then
					If limitYCnt = 0 Then
						FMayLimitSoldout = "Y"
					Else
						FMayLimitSoldout = "N"
					End If
				End If

				ioptNameBuf		= RightCommaDel(ioptNameBuf)
			    ioptCodeBuf		= RightCommaDel(ioptCodeBuf)
			    ioptAddPrice	= RightCommaDel(ioptAddPrice)
			    ioptLimitNo		= RightCommaDel(ioptLimitNo)
			    optlpUsing		= RightCommaDel(optlpUsing)

			    If (ioptTypeName="") then ioptTypeName="옵션명"
			    optstr = ioptTypeName & "<" & ioptNameBuf & ">"

                If (ioptLimitNo <> "") Then
                    optstr = optstr & "수량<" & ioptLimitNo & ">"
                End If
                optstr = optstr & "추가금액<" & ioptAddPrice & ">"
                optstr = optstr & "옵션코드<" & ioptCodeBuf & ">"
                optstr = optstr & "사용여부<" & optlpUsing & ">"
				optstr = Replace(optstr, VbTab, "")
				If Fitemdiv = "06" Then
					buf = buf & "		<prdOption><![CDATA[{" & optstr & "}]]></prdOption>"
				Else
					buf = buf & "		<prdOption><![CDATA[" & optstr & "]]></prdOption>"
				End If
			End If
		End If
		getInterparkOptParamtoEDT = buf
	End Function

	'// 검색어
	Public Function getItemKeyword()
		Dim keywordsBuf, keywordsStr, k
		keywordsBuf = Split(Fkeywords,",")
		For k = 0 to 2
			If UBound(keywordsBuf)> k Then keywordsStr = keywordsStr & Trim(keywordsBuf(k)) & ","
		Next
		keywordsStr = "텐바이텐," & keywordsStr
		keywordsStr = RightCommaDel(keywordsStr)
		If (FItemID = 486220) or (FItemID = 486222) Then
			keywordsStr=""
		End If
		keywordsStr = Replace(keywordsStr,"'","")

		If stringCount(keywordsStr) > 100 Then
			keywordsStr = chrbyte("keywordsStr",100,"N")
		End If
		getItemKeyword = keywordsStr
	End Function

	Function stringCount(strString)
		Dim intPos, chrTemp, intLength
		'문자열 길이 초기화
		intLength = 0
		intPos = 1

		'문자열 길이만큼 돈다
		while ( intPos <= Len( strString ) )
			'문자열을 한문자씩 비교한다
			chrTemp = ASC(Mid( strString, intPos, 1))
			if chrTemp < 0 then '음수값(-)이 나오면 한글임
				intLength = intLength + 2 '한글일 경우 2바이트를 더한다
			else
				intLength = intLength + 1 '한글이 아닐경우 1바이트를 더한다
			end If
			intPos = intPos + 1
		wend
		stringCount = intLength
	End function

	Public Function isImageChanged()
		Dim ibuf : ibuf = getBasicImage
		If InStr(ibuf,"-") < 1 Then
			isImageChanged = FALSE
			Exit Function
		End If
		isImageChanged = ibuf <> FregImageName
	End Function

	Public Function getBasicImage()
		If IsNULL(FbasicImageNm) or (FbasicImageNm="") Then Exit function
		getBasicImage = FbasicImageNm
	End Function

	Public Function IsFreeBeasong()
		IsFreeBeasong = False
		If (FdeliveryType = 2) or (FdeliveryType = 4) or (FdeliveryType = 5) Then
			IsFreeBeasong = True
		End If

		If (FSellcash >= 50000) Then IsFreeBeasong = True
	End Function

    Public Function getOrderCommentStr()
		Dim reStr
		reStr = ""
		If Not IsNULL(Fordercomment) Then
			If Len(Fordercomment) < 2 Then
				reStr = ""
			Else
				reStr = "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
			End If
		End If
		getOrderCommentStr = reStr
    End Function

    Public Function getInterparkAddImageParam()
    	Dim strRst, strSQL, i
    	strRst = ""
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					strRst = strRst & "http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")&","
				End If
				rsget.MoveNext
				If i >= 4 Then Exit For
			Next
		End If
		rsget.Close
		getInterparkAddImageParam = RightCommaDel(strRst)
    End Function

	Public Function getInterparkItemsafetyReg
		Dim strSql, buf, safetyDiv, safetyNum, safetyYn, infoDiv, isElecCate, isLifeCate, isChildrenCate
		Dim certYN, bufLife, bufElec, bufChild, arrRows, notarrRows
		Dim newSafetyDiv, nLp, newDiv, newCertNo
		buf = ""
		bufLife = ""
		bufElec = ""
		bufChild = ""
		strSql = ""
		strSql = strSql & " SELECT TOP 1 " & vbcrlf
		strSql = strSql & " c.itemid, c.safetyYn, c.safetyDiv, isNULL(c.safetyNum, '') as safetyNum, c.infoDiv " & vbcrlf
		strSql = strSql & " , isnull(t.electric, '') as isElecCate, isnull(t.industrial, '') as isLifeCate, isnull(t.child, '') as isChildrenCate " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_item as i " & vbcrlf
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid " & vbcrlf
		strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_interpark_cate_mapping] m on i.cate_large = m.tenCateLarge and i.cate_mid = m.tenCateMid and i.cate_small = m.tenCateSmall " & vbcrlf
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_interpark_category as t on m.CateKey = t.dispNo " & vbcrlf
		strSql = strSql & " WHERE i.itemid='"&FItemID&"' " & vbcrlf
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			safetyYn		= rsget("safetyYn")
			safetyDiv		= rsget("safetyDiv")
			safetyNum		= rsget("safetyNum")
			isElecCate		= rsget("isElecCate")
			isLifeCate		= rsget("isLifeCate")
			isChildrenCate	= rsget("isChildrenCate")
		End If
		rsget.Close

		strSql = ""
		strSql = strSql & " SELECT TOP 5 certNum, safetyDiv " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg " & vbcrlf
		strSql = strSql & " WHERE itemid='"&FItemID&"' " & vbcrlf
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			arrRows = rsget.getRows()
		Else
			notarrRows = "Y"
		End If
		rsget.Close

		If (isElecCate = "Y") OR (isLifeCate = "Y") OR (isChildrenCate = "Y") Then
			If notarrRows = "" Then		'전안법 적용된 데이터라면 제대로 꼽기
				If safetyYn = "Y" Then
					For nLp =0 To UBound(arrRows,2)
				    	newDiv = ""
						Select Case arrRows(1,nLp)
							Case "10"		newDiv = "0201"		'전기용품 > 안전인증
							Case "20"		newDiv = "0202"		'전기용품 > 안전확인 신고
							Case "30"		newDiv = "0203"		'전기용품 > 공급자 적합성 확인
							Case "40"		newDiv = "0101"		'생활제품 > 안전인증
							Case "50"		newDiv = "0102"		'생활제품 > 자율안전확인
							Case "60"		newDiv = "0104"		'생활제품 > 안전품질표시
							Case "70"		newDiv = "0401"		'어린이제품 > 안전인증
							Case "80"		newDiv = "0402"		'어린이제품 > 안전확인
							Case "90"		newDiv = "0403"		'어린이제품 > 공급자 적합성 확인
						End Select

						newCertNo = arrRows(0,nLp)
						If newCertNo = "x" Then
							newCertNo = ""
						End If

						If newDiv = "0201" OR newDiv = "0202" OR newDiv = "0203" Then
					    	bufElec = bufElec & "			<certInfo>"
					    	bufElec = bufElec & "				<certKind><![CDATA["&newDiv&"]]></certKind>"
					    	bufElec = bufElec & "				<certNo><![CDATA["&newCertNo&"]]></certNo>"
					    	bufElec = bufElec & "			</certInfo>"
						ElseIf newDiv = "0101" OR newDiv = "0102" OR newDiv = "0104" Then
					    	bufLife = bufLife & "			<certInfo>"
					    	bufLife = bufLife & "				<certKind><![CDATA["&newDiv&"]]></certKind>"
					    	bufLife = bufLife & "				<certNo><![CDATA["&newCertNo&"]]></certNo>"
					    	bufLife = bufLife & "			</certInfo>"
						ElseIf newDiv = "0401" OR newDiv = "0402" OR newDiv = "0403" Then
					    	bufChild = bufChild & "			<certInfo>"
					    	bufChild = bufChild & "				<certKind><![CDATA["&newDiv&"]]></certKind>"
					    	bufChild = bufChild & "				<certNo><![CDATA["&newCertNo&"]]></certNo>"
					    	bufChild = bufChild & "			</certInfo>"
						End If
					Next

					If (isElecCate = "Y") AND (isLifeCate = "Y") Then					'전기랑 생활이 둘다 있는 인팍 카테고리
						If bufElec <> "" OR bufLife <> "" Then
					    	buf = buf & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
					    	buf = buf & "		<prdCertDetail>"
							If bufElec <> "" Then
								buf = buf & bufElec
							ElseIf bufLife <> "" Then
								buf = buf & bufLife
							End If
							buf = buf & "		</prdCertDetail>"
						Else
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'상품 설명 내 표기
						End If
					ElseIf (isElecCate = "Y") AND (isChildrenCate = "Y") Then
						If bufElec <> "" OR bufChild <> "" Then
					    	buf = buf & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
					    	buf = buf & "		<prdCertDetail>"
							If bufElec <> "" Then
								buf = buf & bufElec
							ElseIf bufChild <> "" Then
								buf = buf & bufChild
							End If
							buf = buf & "		</prdCertDetail>"
						Else
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'상품 설명 내 표기
						End If
					ElseIf (isLifeCate = "Y") AND (isChildrenCate = "Y") Then
						If bufLife <> "" OR bufChild <> "" Then
					    	buf = buf & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
					    	buf = buf & "		<prdCertDetail>"
							If bufLife <> "" Then
								buf = buf & bufLife
							ElseIf bufChild <> "" Then
								buf = buf & bufChild
							End If
							buf = buf & "		</prdCertDetail>"
						Else
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'상품 설명 내 표기
						End If
					ElseIf (isElecCate = "Y") Then
						If bufElec = "" Then
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'상품 설명 내 표기
						Else
					    	buf = buf & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
					    	buf = buf & "		<prdCertDetail>"
							buf = buf & bufElec
							buf = buf & "		</prdCertDetail>"
						End If
					ElseIf (isLifeCate = "Y") Then
						If bufLife = "" Then
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'상품 설명 내 표기
						Else
					    	buf = buf & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
					    	buf = buf & "		<prdCertDetail>"
							buf = buf & bufLife
							buf = buf & "		</prdCertDetail>"
						End If
					ElseIf (isChildrenCate = "Y") Then
						If bufChild = "" Then
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'상품 설명 내 표기
						Else
					    	buf = buf & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
					    	buf = buf & "		<prdCertDetail>"
							buf = buf & bufChild
							buf = buf & "		</prdCertDetail>"
						End If
					End If
				Else
					buf = buf & "		<prdCertStatus><![CDATA[N]]></prdCertStatus>"
				End If
			Else						'전안법 적용 안 된 데이터라면 구버전 데이터 꼽기
				If safetyNum = "" OR safetyYn = "N" then	'인증번호가 없거나 인증이 아니라면
					certYN = "N"
				Else
					certYN = "Y"
				End If

				If certYN = "N" Then
					buf = buf & "		<prdCertStatus><![CDATA[N]]></prdCertStatus>"
				Else
				    If safetyDiv = "10" Then		'우리쪽 국가통합인증(KC마크)
				    	bufLife = bufLife & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
				    	bufLife = bufLife & "		<prdCertDetail>"
				    	bufLife = bufLife & "			<certInfo>"
				    	bufLife = bufLife & "				<certKind><![CDATA[0101]]></certKind>"					'생활용품] 안전인증확인
				    	bufLife = bufLife & "				<certNo><![CDATA["&safetyNum&"]]></certNo>"
				    	bufLife = bufLife & "			</certInfo>"
				    	bufLife = bufLife & "		</prdCertDetail>"
				    ElseIf safetyDiv = "20" Then	'우리쪽 전기용품 안전인증
				    	bufElec = bufElec & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
				    	bufElec = bufElec & "		<prdCertDetail>"
				    	bufElec = bufElec & "			<certInfo>"
				    	bufElec = bufElec & "				<certKind><![CDATA[0201]]></certKind>"					'[전기용품] 안전인증확인
				    	bufElec = bufElec & "				<certNo><![CDATA["&safetyNum&"]]></certNo>"
				    	bufElec = bufElec & "			</certInfo>"
				    	bufElec = bufElec & "		</prdCertDetail>"
					ElseIf safetyDiv = "30" Then	'우리쪽 KPS 안전인증 표시
				    	bufLife = bufLife & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
				    	bufLife = bufLife & "		<prdCertDetail>"
				    	bufLife = bufLife & "			<certInfo>"
				    	bufLife = bufLife & "				<certKind><![CDATA[0104]]></certKind>"					'[생활용품] 안전품질표시
				    	bufLife = bufLife & "				<certNo><![CDATA["&safetyNum&"]]></certNo>"
				    	bufLife = bufLife & "			</certInfo>"
				    	bufLife = bufLife & "		</prdCertDetail>"
					ElseIf safetyDiv = "40" Then	'우리쪽 KPS 자율안전 확인 표시
				    	bufLife = bufLife & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
				    	bufLife = bufLife & "		<prdCertDetail>"
				    	bufLife = bufLife & "			<certInfo>"
				    	bufLife = bufLife & "				<certKind><![CDATA[0102]]></certKind>"					'[생활용품] 자율안전확인
				    	bufLife = bufLife & "				<certNo><![CDATA["&safetyNum&"]]></certNo>"
				    	bufLife = bufLife & "			</certInfo>"
				    	bufLife = bufLife & "		</prdCertDetail>"
					ElseIf safetyDiv = "50" Then	'우리쪽 KPS 어린이 보호포장 표시
				    	bufLife = bufLife & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
				    	bufLife = bufLife & "		<prdCertDetail>"
				    	bufLife = bufLife & "			<certInfo>"
				    	bufLife = bufLife & "				<certKind><![CDATA[0103]]></certKind>"					'[생활용품] 어린이보호포장
				    	bufLife = bufLife & "				<certNo><![CDATA["&safetyNum&"]]></certNo>"
				    	bufLife = bufLife & "			</certInfo>"
				    	bufLife = bufLife & "		</prdCertDetail>"
					End If

					If (isElecCate = "Y") AND (isLifeCate = "Y") Then					'전기랑 생활이 둘다 있는 인팍 카테고리
						If bufElec <> "" Then
							buf = buf & bufElec
						ElseIf bufLife <> "" Then
							buf = buf & bufLife
						Else
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'상품 설명 내 표기
						End If
					ElseIf (isElecCate = "Y") AND (isChildrenCate = "Y") Then
						If bufElec <> "" Then
							buf = buf & bufElec
						Else
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'상품 설명 내 표기
						End If
					ElseIf (isLifeCate = "Y") AND (isChildrenCate = "Y") Then
						If bufLife <> "" Then
							buf = buf & bufLife
						Else
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'상품 설명 내 표기
						End If
					ElseIf (isElecCate = "Y") Then
						If bufElec = "" Then
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'상품 설명 내 표기
						Else
							buf = buf & bufElec
						End If
					ElseIf (isLifeCate = "Y") Then
						If bufLife = "" Then
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'상품 설명 내 표기
						Else
							buf = buf & bufLife
						End If
					ElseIf (isChildrenCate = "Y") Then
						buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'상품 설명 내 표기
					End If
				End If
			End If
		Else
			buf = buf & "		<prdCertStatus><![CDATA[N]]></prdCertStatus>"
		End If
		getInterparkItemsafetyReg = buf
'rw buf
	End Function

	'진영 상품품목관리 코드 관련 2012-11-12 생성
    Public Function getInterparkItemInfoCdToReg()
		Dim strSql, buf
		Dim mallinfoCd,infoContent,infotype
		'''IC.safetyyn => isNULL(IC.safetyyn,'N')
		'2014-05-15김진영 00002 추가
		strSql = "EXEC [db_etcmall].[dbo].[usp_API_Interpark_InfoCodeMap_Get] " & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) then
			buf = buf & "<prdinfoNoti>"
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infotype	= rsget("infotype")
			    infoContent = rsget("infoContent")

			    If Not (IsNULL(infoContent)) AND (infoContent <> "") Then
			    	infoContent = replace(infoContent, chr(31), "")
				End If

				buf = buf & "<info>"
				buf = buf & "	<infoSubNo><![CDATA["&mallinfoCd&"]]></infoSubNo>"
				buf = buf & "	<infoCd>"&infotype&"</infoCd>"
				buf = buf & "	<infoTx><![CDATA["&infoContent&"]]></infoTx>"
				buf = buf & "</info>"
				rsget.MoveNext
			Loop
			buf = buf & "</prdinfoNoti>"
		End If
		rsget.Close
		getInterparkItemInfoCdToReg = buf
    End Function

	Public Function getInterparkItemRegParameter()
		Dim strRst
	    strRst = ""
	    strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr"" ?>"
	    strRst = strRst & "<result>"
	    strRst = strRst & "	<title>Interpark Product API</title>"
	    strRst = strRst & "	<description>상품 등록</description>"
	    strRst = strRst & "	<item>"
		strRst = strRst & "		<prdStat>01</prdStat>"																'상품상태 - 새상품:01, 중고상품:02, 반품상품:03
		strRst = strRst & "		<shopNo>0000100000</shopNo>"														'인터파크 상점번호 (default - 0000100000)  | 상점번호 API업체 고정
		strRst = strRst & "		<omDispNo>" & Trim(Finterparkdispcategory) & "</omDispNo>" 							'인터파크 전시코드
	    strRst = strRst & "		<prdNm><![CDATA["&getItemNameFormat&"]]></prdNm>"									'상품명 - 한글 60자 (영문/숫자 120자)
	    strRst = strRst & "		<hdelvMafcEntrNm><![CDATA["&CStr(FMakerName)&"]]></hdelvMafcEntrNm>"				'제조업체명
	    strRst = strRst & "		<prdOriginTp><![CDATA["&GetSourcearea&"]]></prdOriginTp>"							'원산지
	    strRst = strRst & "		<taxTp>"&Chkiif(Fvatinclude="Y", "01", "02")&"</taxTp>"								'부가면세상품 - 과세상품:01, 면세상품:02, 영세상품:03
	    strRst = strRst & "		<ordAgeRstrYn>"&Chkiif(IsAdultItem() = "Y", "Y", "N")&"</ordAgeRstrYn>"				'성인용품여부 - 성인용품:Y, 일반용품:N
		strRst = strRst & "		<saleStatTp>01</saleStatTp>"														'판매중:01, 품절:02, 판매중지:03, 일시품절:05, 예약판매:09, 상품삭제:98
	    strRst = strRst & "		<saleUnitcost>"&MustPrice&"</saleUnitcost>"											'판매가
		strRst = strRst & "		<saleLmtQty>"&GetInterParkLmtQty&"</saleLmtQty>"									'판매수량 - 99999 개 이하로 입력
		strRst = strRst & "		<saleStrDts>"&Replace(Left(CStr(now()),10),"-","")&"</saleStrDts>"					'판매시작일 - yyyyMMdd => 호출당시 날짜
		strRst = strRst & "		<saleEndDts>"&GetSellEndDateStr&"</saleEndDts>"										'판매종료일 - yyyyMMdd => 99991231 (고정값)
'2017-11-27 김진영 수정..전 상품 N으로
'		strRst = strRst & "		<proddelvCostUseYn>"&Chkiif(Fdeliverytype="4", "Y", "N")&"</proddelvCostUseYn>"		'상품배송비사용여부 - 상품배송비사용:Y, 업체배송비정책사용:N
		strRst = strRst & "		<proddelvCostUseYn>N</proddelvCostUseYn>"											'상품배송비사용여부 - 상품배송비사용:Y, 업체배송비정책사용:N
	If IsTruckReturnDlvExists Then
		strRst = strRst & "		<prdrtnCostUseYn>Y</prdrtnCostUseYn>"												'상품 반품택배비 사용여부 - 상품반품택배비사용:Y, 업체반품택배비사용:N
		strRst = strRst & "		<rtndelvCost>"&getTruckReturnDlvPrice&"</rtndelvCost>"								'상품 반품택배비. prdrtnCostUseYn 가 'Y' 일 경우 필수임
	End If
		strRst = strRst & "		<prdBasisExplanEd><![CDATA["&getInterparkContParamToReg&"]]></prdBasisExplanEd>"	'상품설명
		strRst = strRst & "		<zoomImg><![CDATA["&Fbasicimage&"]]></zoomImg>"										'대표이미지 - 대표이미지 URL, 영문/숫자 조합, JPG와 GIF만 가능
		strRst = strRst & "		<prdKeywd><![CDATA["&getItemKeyword&"]]></prdKeywd>"								'쇼핑태그 - 최대 4개까지, 콤마로 구분
		strRst = strRst & "		<brandNm><![CDATA["&Fbrandname&"]]></brandNm>"										'브랜드명
		strRst = strRst & "		<entrPoint>"&GetInterParkentrPoint&"</entrPoint>"									'업체POINT - 업체부여 포인트 금액 입력, 판매가의 최대 10%까지 가능
		strRst = strRst & "		<perordRstrQty>"& getOrderMaxNum &"</perordRstrQty>"								'1회당 주문 제한 수량
		strRst = strRst & "		<minOrdQty>1</minOrdQty>"															'최소구매수량 - 1개 이상 입력
		strRst = strRst & getInterparkOptParamtoREG
	If (Fitemdiv = "06") Then
		strRst = strRst & "		<inOpt>주문제작문구</inOpt>"														'입력형 옵션. ex) 사은품을 입력하세요.
	End If
'2017-11-27 김진영 수정..전 상품 N으로 설정했기에 하단 필드 주석
'	If (Fdeliverytype = "4") Then
'		strRst = strRst & "		<delvCost>0</delvCost>"																'배송비 -상품 배송비 선택일때 필수, 0이면 무료배송
'	End If
		strRst = strRst & "		<delvAmtPayTpCom>"&Chkiif(FdeliveryType = "7", "01", "02")&"</delvAmtPayTpCom>"		'배송비 결제 방식 - 착불:01, 선불:02, 선불전환착불가능:03 상품배송비를 사용할 경우 필수, 무료배송일때:02
    	strRst = strRst & "		<delvCostApplyTp>02</delvCostApplyTp>"												'배송비 적용 방식 - 개당:01, 무조건:02
	If (IsFreeBeasong) Then
		strRst = strRst & "		<freedelvStdCnt>1</freedelvStdCnt>"													'무료배송기준 수량 - 기준수량 입력 사용하지 않을 경우 0
	End If
		strRst = strRst & "		<jejuetcDelvCostUseYn>Y</jejuetcDelvCostUseYn>"										'제주도서산간배송비사용여부 - Y : 등록/수정, N : 사용안함
		strRst = strRst & "		<jejuDelvCost>3000</jejuDelvCost>"													'제주배송비 - jejuetcDelvCostUseYn가 Y일때 제주배송비와 도서산간비 둘 중 하나는 필수, 0이면 제주배송비 0원, null이면 사용안함
		strRst = strRst & "		<etcDelvCost>3000</etcDelvCost>"													'도서산간배송비 - jejuetcDelvCostUseYn가 Y일때 제주배송비와 도서산간비 둘 중 하나는 필수, 0이면 도서산간배송비 0원, null이면 사용안함
		strRst = strRst & "		<spcaseEd><![CDATA[" & getOrderCommentStr & "]]></spcaseEd>"						'특이사항
		strRst = strRst & "		<pointmUseYn>N</pointmUseYn>"														'포인트몰등록여부 - 포인트몰상품:Y, 일반상품:N 단 500원 미만 상품은 등록이 불가능합니다. || 2013/02/07 이제 인팍에서 아이포인트 사용률이 줄어서 당분간 진행안하기로 했거든요~ 그 비용을 광고비로 사용하기로 해서 오늘부터 아이포인트를 다 빼주시면 될 거 같습니다~   컨펌 받은 내용이고 인팍 쪽에서도 오늘 요청 들어갔다고 합니당~
		strRst = strRst & "		<ippSubmitYn>Y</ippSubmitYn>"														'가격비교등록여부
		strRst = strRst & "		<originPrdNo>"&CStr(FItemID)&"</originPrdNo>"										'상품번호
		strRst = strRst & "		<asInfo>상세페이지참조</asInfo>"													'A/S정보 | 2016-08-11 김진영 추가..신규개설필드
		strRst = strRst & "		<detailImg>"&getInterparkAddImageParam&"</detailImg>"								'상세이미지 - 상세이미지 URL, 영문/숫자 조합, JPG와 GIF만 가능 최대 4개의 이미지까지, 콤마(,)로 구분하여 등록.
		strRst = strRst & getInterparkItemsafetyReg()
 		strRst = strRst & getInterparkItemInfoCdToReg()
	    strRst = strRst & "	</item>"
	    strRst = strRst & "</result>"
		getInterparkItemRegParameter = strRst
	End Function

	Public Function getInterparkItemEditParameter()
	    Dim strRst
	    strRst = ""
	    strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr"" ?>"
	    strRst = strRst & "<result>"
	    strRst = strRst & "	<title>Interpark Product API</title>"
	    strRst = strRst & "	<description>상품 수정</description>"
	    strRst = strRst & "	<item>"
	    strRst = strRst & "		<prdNo>"&FInterparkPrdNo&"</prdNo>"													'인터파크 상품번호
	    strRst = strRst & "		<prdStat>01</prdStat>"																'상품상태 - 새상품:01, 중고상품:02, 반품상품:03
	    strRst = strRst & "		<prdNm><![CDATA["&getItemNameFormat&"]]></prdNm>"									'상품명
	    strRst = strRst & "		<hdelvMafcEntrNm><![CDATA["&CStr(FMakerName)&"]]></hdelvMafcEntrNm>"				'제조업체명
	    strRst = strRst & "		<prdOriginTp><![CDATA["&GetSourcearea&"]]></prdOriginTp>"							'원산지
	    strRst = strRst & "		<taxTp>"&Chkiif(Fvatinclude="Y", "01", "02")&"</taxTp>"								'부가면세상품 - 과세상품:01, 면세상품:02, 영세상품:03
	    strRst = strRst & "		<ordAgeRstrYn>"&Chkiif(IsAdultItem() = "Y", "Y", "N")&"</ordAgeRstrYn>"				'성인용품여부 - 성인용품:Y, 일반용품:N
	    strRst = strRst & "		<saleStatTp>"&GetInterParkSaleStatTp&"</saleStatTp>"								'판매중:01, 품절:02, 판매중지:03, 일시품절:05, 예약판매:09, 상품삭제:98
	    strRst = strRst & "		<saleUnitcost>"&MustPrice&"</saleUnitcost>"											'판매가
		strRst = strRst & "		<saleLmtQty>"&GetInterParkLmtQty&"</saleLmtQty>"									'판매수량 - 99999 개 이하로 입력
		'2018-04-17 김진영..수정시엔 아래 필드 주석
		'2018-11-01 김진영..수정시 saleStrDts 필드없다고 오류 발생으로 주석 제거
		strRst = strRst & "		<saleStrDts>"&Replace(Left(CStr(now()),10),"-","")&"</saleStrDts>"						'판매시작일 - yyyyMMdd => 호출당시 날짜
		strRst = strRst & "		<saleEndDts>"&GetSellEndDateStr&"</saleEndDts>"										'판매종료일 - yyyyMMdd => 99991231 (고정값)
'2017-11-27 김진영 수정..전 상품 N으로
'		strRst = strRst & "		<proddelvCostUseYn>"&Chkiif(Fdeliverytype="4", "Y", "N")&"</proddelvCostUseYn>"		'상품배송비사용여부 - 상품배송비사용:Y, 업체배송비정책사용:N
		strRst = strRst & "		<proddelvCostUseYn>N</proddelvCostUseYn>"											'상품배송비사용여부 - 상품배송비사용:Y, 업체배송비정책사용:N
	If IsTruckReturnDlvExists Then
		strRst = strRst & "		<prdrtnCostUseYn>Y</prdrtnCostUseYn>"												'상품 반품택배비 사용여부 - 상품반품택배비사용:Y, 업체반품택배비사용:N
		strRst = strRst & "		<rtndelvCost>"&getTruckReturnDlvPrice&"</rtndelvCost>"								'상품 반품택배비. prdrtnCostUseYn 가 'Y' 일 경우 필수임
	End If
		strRst = strRst & "		<prdBasisExplanEd><![CDATA["&getInterparkContParamToReg&"]]></prdBasisExplanEd>"	'상품설명
		strRst = strRst & "		<zoomImg><![CDATA["&Fbasicimage&"]]></zoomImg>"										'대표이미지 - 대표이미지 URL, 영문/숫자 조합, JPG와 GIF만 가능
		strRst = strRst & "		<prdKeywd><![CDATA["&getItemKeyword&"]]></prdKeywd>"								'쇼핑태그 - 최대 4개까지, 콤마로 구분
		strRst = strRst & "		<brandNm><![CDATA["&Fbrandname&"]]></brandNm>"										'브랜드명
		strRst = strRst & "		<entrPoint>"&GetInterParkentrPoint&"</entrPoint>"									'업체POINT - 업체부여 포인트 금액 입력, 판매가의 최대 10%까지 가능
		strRst = strRst & "		<perordRstrQty>"& getOrderMaxNum &"</perordRstrQty>"								'1회당 주문 제한 수량
		strRst = strRst & "		<minOrdQty>1</minOrdQty>"															'최소구매수량 - 1개 이상 입력
		strRst = strRst & getInterparkOptParamtoEDT
	If (Fitemdiv = "06") Then
		strRst = strRst & "		<inOpt>주문제작문구</inOpt>"														'입력형 옵션. ex) 사은품을 입력하세요.
	End If
'2017-11-27 김진영 수정..전 상품 N으로 설정했기에 하단 필드 주석
'	If (Fdeliverytype = "4") Then
'		strRst = strRst & "		<delvCost>0</delvCost>"																'배송비 -상품 배송비 선택일때 필수, 0이면 무료배송
'	End If
		strRst = strRst & "		<delvAmtPayTpCom>"&Chkiif(FdeliveryType = "7", "01", "02")&"</delvAmtPayTpCom>"		'배송비 결제 방식 - 착불:01, 선불:02, 선불전환착불가능:03 상품배송비를 사용할 경우 필수, 무료배송일때:02
    	strRst = strRst & "		<delvCostApplyTp>02</delvCostApplyTp>"												'배송비 적용 방식 - 개당:01, 무조건:02
	If (IsFreeBeasong) Then
		strRst = strRst & "		<freedelvStdCnt>1</freedelvStdCnt>"													'무료배송기준 수량 - 기준수량 입력 사용하지 않을 경우 0
	End If
		strRst = strRst & "		<jejuetcDelvCostUseYn>Y</jejuetcDelvCostUseYn>"										'제주도서산간배송비사용여부 - Y : 등록/수정, N : 사용안함
		strRst = strRst & "		<jejuDelvCost>3000</jejuDelvCost>"													'제주배송비 - jejuetcDelvCostUseYn가 Y일때 제주배송비와 도서산간비 둘 중 하나는 필수, 0이면 제주배송비 0원, null이면 사용안함
		strRst = strRst & "		<etcDelvCost>3000</etcDelvCost>"													'도서산간배송비 - jejuetcDelvCostUseYn가 Y일때 제주배송비와 도서산간비 둘 중 하나는 필수, 0이면 도서산간배송비 0원, null이면 사용안함
		strRst = strRst & "		<spcaseEd><![CDATA[" & getOrderCommentStr & "]]></spcaseEd>"						'특이사항
		strRst = strRst & "		<pointmUseYn>N</pointmUseYn>"														'포인트몰등록여부 - 포인트몰상품:Y, 일반상품:N 단 500원 미만 상품은 등록이 불가능합니다. || 2013/02/07 이제 인팍에서 아이포인트 사용률이 줄어서 당분간 진행안하기로 했거든요~ 그 비용을 광고비로 사용하기로 해서 오늘부터 아이포인트를 다 빼주시면 될 거 같습니다~   컨펌 받은 내용이고 인팍 쪽에서도 오늘 요청 들어갔다고 합니당~
		strRst = strRst & "		<ippSubmitYn>Y</ippSubmitYn>"														'가격비교등록여부
		strRst = strRst & "		<originPrdNo>"&CStr(FItemID)&"</originPrdNo>"										'상품번호
		strRst = strRst & "		<asInfo>상세페이지참조</asInfo>"													'A/S정보 | 2016-08-11 김진영 추가..신규개설필드
	If isImageChanged Then
		strRst = strRst & "		<detailImg>"&getInterparkAddImageParam&"</detailImg>"								'상세이미지 - 상세이미지 URL, 영문/숫자 조합, JPG와 GIF만 가능 최대 4개의 이미지까지, 콤마(,)로 구분하여 등록.
		strRst = strRst & "		<imgUpdateYn>Y</imgUpdateYn>"														'이미지수정여부 - 대표이미지,상세이미지의 수정여부를 결정 합니다.Y : 이미지 수정 필요N : 이미지 수정 불필요(기본값 : N)대표이미지나 상세이미지 중에 하나만이라도 수정이 필요한 경우 Y로 설정해야 합니다.
	End If
		strRst = strRst & getInterparkItemsafetyReg()
 		strRst = strRst & getInterparkItemInfoCdToReg()
	    strRst = strRst & "	</item>"
	    strRst = strRst & "</result>"
		getInterparkItemEditParameter = strRst
	End Function
End Class

Class CInterpark
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

	Public Sub getInterparkNotRegOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & " ,c.makername, uc.socname_kor, uc.defaultfreeBeasongLimit "
		strSql = strSql & " ,c.keywords, c.ordercomment, c.itemcontent, c.sourcearea "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & " ,c.usinghtml, m.CateKey "
        strSql = strSql & " ,isNULL(c.freight_min,0) as freight_min, isNULL(c.freight_max,0) as freight_max "
        strSql = strSql & " ,isNULL(s.regImageName,'') as regImageName"
		strSql = strSql & " FROM [db_item].[dbo].tbl_interpark_reg_item s, [db_item].[dbo].tbl_item i "
		strSql = strSql & " LEFT JOIN [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		strSql = strSql & " LEFT JOIN [db_etcmall].[dbo].tbl_interpark_cate_mapping m on i.cate_large = m.tenCateLarge and i.cate_mid = m.tenCateMid and i.cate_small = m.tenCateSmall "
	    strSql = strSql & " LEFT JOIN [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid "
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE s.itemid = i.itemid"
		strSql = strSql & " and s.itemid in ("
		strSql = strSql & " 	SELECT TOP " & CStr(FPageSize * FCurrPage) & " s.itemid "
		strSql = strSql & " 	FROM [db_item].[dbo].tbl_interpark_reg_item s, [db_item].[dbo].tbl_item i, [db_etcmall].[dbo].tbl_interpark_cate_mapping p "
		strSql = strSql & " 	WHERE s.itemid = i.itemid"
		strSql = strSql & "		and s.interparkregdate is NULL"
		strSql = strSql & "		and i.basicimage is not null"
		strSql = strSql & "		and i.cate_large <> '' "
		strSql = strSql & "		and i.cate_large <> '999' "
		strSql = strSql & "		and i.sellcash > 0"
		strSql = strSql & " 	and rtrim(ltrim(isNull(i.deliverfixday, ''))) = '' "		'택배(일반)
		strSql = strSql & "		and i.itemdiv in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
		strSql = strSql & " 	and i.isusing = 'Y' "
	    strSql = strSql & " 	and i.isExtusing = 'Y'"
	    strSql = strSql & " 	and i.sellyn='Y'"           '''판매중인 상품만 등록. // 조건 추가 2011-11-02
	    strSql = strSql & "		and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = '"&CMALLNAME&"')"	'등록제외브랜드
	    strSql = strSql & "		and i.itemid NOT IN (SELECT itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = '"&CMALLNAME&"')"		'등록제외상품
		strSql = strSql & "		and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'등록제외 카테고리
	    strSql = strSql & " 	and i.deliverytype not in ('7','6') "   '''착불 등록 제외 // 조건 추가 2011-11-02
'	    strSql = strSql & " 	and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000)))"
		strSql = strSql & "		and i.cate_large= p.tenCateLarge "
		strSql = strSql & "		and i.cate_mid = p.tenCateMid "
		strSql = strSql & "		and i.cate_small = p.tenCateSmall "
	    strSql = strSql & "		and p.CateKey is Not NULL"   '' 전시코드
		strSql = strSql & "		and 'Y' = case	when i.sailyn = 'Y' "
		strSql = strSql & " 					AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &") ) "
		strSql = strSql & " 						OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & " 					) THEN 'Y' "
		strSql = strSql & " 					WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) THEN 'Y' ELSE 'N' END "
		If FRectItemID <> "" Then
			strSql = strSql & "		and s.itemid in (" & FRectItemID & ")"
		End If
		strSql = strSql & " )"
		strSql = strSql & " and uc.isExtusing <> 'N'"
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CInterparkItem
				FOneItem.Fitemid					= rsget("itemid")
				FOneItem.Fitemname					= LeftB(db2html(rsget("itemname")),255)
				FOneItem.FMakerid					= rsget("makerid")
				FOneItem.Fsellcash					= rsget("sellcash")
				FOneItem.Forgsellcash				= rsget("orgprice")
				FOneItem.Fsourcearea				= LeftB(db2html(rsget("sourcearea")),64)
				FOneItem.FRegdate					= rsget("regdate")
				FOneItem.Fsellyn					= rsget("sellyn")
				FOneItem.Flimityn					= rsget("limityn")
				FOneItem.Flimitno					= rsget("limitno")
				FOneItem.Flimitsold					= rsget("limitsold")
				FOneItem.Fcate_large				= rsget("cate_large")
				FOneItem.Fcate_mid					= rsget("cate_mid")
				FOneItem.Fcate_small				= rsget("cate_small")
				FOneItem.FMakerName					= db2html(rsget("makername"))
				FOneItem.FBrandName					= db2html(rsget("brandname"))
				FOneItem.Foptioncnt					= rsget("optioncnt")
				FOneItem.FBrandNameKor = db2html(rsget("socname_kor"))
			If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
				FOneItem.FMakerName					= FOneItem.FBrandName
			End If
				FOneItem.Fkeywords					= db2html(rsget("keywords"))
				FOneItem.Fbasicimage				= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FregImageName				= rsget("regImageName")
				FOneItem.Fmainimage					= rsget("mainimage")
			If IsNULL(FOneItem.FInfoImage) Then
				FOneItem.FInfoImage				= ",,,,"
			End If
				FOneItem.Fordercomment				= db2html(rsget("ordercomment"))
				FOneItem.FItemContent				= db2html(rsget("itemcontent"))
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,"♂","")
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,"","")
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,"","")
				FOneItem.Fsourcearea				= db2html(rsget("sourcearea"))
				FOneItem.Fvatinclude				= rsget("vatinclude")
			If (rsget("usinghtml") = "N") Then
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,vbcrlf,"<br>")
			End If
				FOneItem.Finterparkdispcategory		= rsget("CateKey")
				FOneItem.Fitemsize					= db2html(rsget("itemsize"))
				FOneItem.Fitemsource				= db2html(rsget("itemsource"))
				FOneItem.FLastUpdate				= rsget("LastUpdate")
				FOneItem.FSellEndDate				= rsget("sellenddate")
				FOneItem.FItemDiv					= rsget("ItemDiv")
				FOneItem.Fisusing					= rsget("isusing")
				FOneItem.FSailYn					= rsget("sailyn")
				FOneItem.FOrgPrice					= rsget("orgprice")
				FOneItem.FdeliveryType				= rsget("deliveryType")
				FOneItem.FdefaultfreeBeasongLimit	= rsget("defaultfreeBeasongLimit")
				FOneItem.Fdeliverfixday				= rsget("deliverfixday")
				FOneItem.Ffreight_min				= rsget("freight_min")
				FOneItem.Ffreight_max				= rsget("freight_max")
				FOneItem.FAdultType 				= rsget("adulttype")
				FOneItem.FOrderMaxNum 				= rsget("orderMaxNum")
		End If
		rsget.Close
	End Sub

	Public Sub getInterparkEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & " ,c.makername, uc.socname_kor, uc.defaultfreeBeasongLimit "
		strSql = strSql & " ,c.keywords, c.ordercomment, c.itemcontent, c.sourcearea "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & " ,s.interparkPrdNo, s.mayiParkSellYn "
        strSql = strSql & " ,c.usinghtml, m.CateKey ,s.interparkregdate "
		strSql = strSql & " ,isNULL(c.freight_min,0) as freight_min, isNULL(c.freight_max,0) as freight_max "
		strSql = strSql & " ,isNULL(s.regImageName,'') as regImageName, isNULL(s.lastErrStr,'') as lastErrStr, s.mayiparkprice "
		strSql = strSql & " ,(SELECT COUNT(*) as regOptCnt FROM db_item.dbo.tbl_outmall_regedoption as RO WHERE RO.itemid = s.itemid and RO.mallid = 'interpark') as regOptCnt "
		strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or rtrim(ltrim(isNull(i.deliverfixday, ''))) <> '' "
		strSql = strSql & " 	or i.itemdiv not in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
		strSql = strSql & "		or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & "		or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM [db_item].[dbo].tbl_interpark_reg_item s, [db_item].[dbo].tbl_item i "
		strSql = strSql & " LEFT JOIN [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid "
        strSql = strSql & " LEFT JOIN [db_etcmall].[dbo].tbl_interpark_cate_mapping m on i.cate_large = m.tenCateLarge and i.cate_mid = m.tenCateMid and i.cate_small = m.tenCateSmall "
		strSql = strSql & " LEFT JOIN [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid "
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE s.itemid=i.itemid"
		strSql = strSql & addSql
		strSql = strSql & " ORDER BY i.itemid "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CInterparkItem
				FOneItem.Fitemid				= rsget("itemid")
				FOneItem.Fitemname				= LeftB(db2html(rsget("itemname")),255)
				FOneItem.FMakerid				= rsget("makerid")
				FOneItem.Fbuycash				= rsget("buycash")
				FOneItem.Fsellcash				= rsget("sellcash")
				FOneItem.Forgsellcash			= rsget("orgprice")
				FOneItem.Fsourcearea			= LeftB(db2html(rsget("sourcearea")),64)
				FOneItem.Foptioncnt				= rsget("optioncnt")
				FOneItem.FRegdate				= rsget("regdate")
				FOneItem.Fsellyn				= rsget("sellyn")
				FOneItem.Flimityn				= rsget("limityn")
				FOneItem.Flimitno				= rsget("limitno")
				FOneItem.Flimitsold				= rsget("limitsold")
				FOneItem.Fcate_large			= rsget("cate_large")
				FOneItem.Fcate_mid				= rsget("cate_mid")
				FOneItem.Fcate_small			= rsget("cate_small")
				FOneItem.FMakerName				= db2html(rsget("makername"))
				FOneItem.FBrandName				= db2html(rsget("brandname"))
				FOneItem.FBrandNameKor			= db2html(rsget("socname_kor"))
				If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
					FOneItem.FMakerName			= FOneItem.FBrandName
				End If
				FOneItem.Fkeywords				= db2html(rsget("keywords"))
				FOneItem.Fbasicimage			= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FregImageName			= rsget("regImageName")
				FOneItem.Fmainimage				= rsget("mainimage")
				FOneItem.Fmainimage2			= rsget("mainimage2")
				If IsNULL(FOneItem.FInfoImage) Then
					FOneItem.FInfoImage			= ",,,,"
				End If
				FOneItem.Fordercomment			= db2html(rsget("ordercomment"))
				FOneItem.FItemContent			= db2html(rsget("itemcontent"))
				FOneItem.FItemContent			= replace(FOneItem.FItemContent,"♂","")
				FOneItem.FItemContent			= replace(FOneItem.FItemContent,"","")
				FOneItem.FItemContent			= replace(FOneItem.FItemContent,"","")
				FOneItem.Fsourcearea			= db2html(rsget("sourcearea"))
				FOneItem.Fvatinclude			= rsget("vatinclude")
				If (rsget("usinghtml") = "N") Then
					FOneItem.FItemContent		= replace(FOneItem.FItemContent,vbcrlf,"<br>")
				End If
                FOneItem.Finterparkdispcategory	= rsget("CateKey")
				FOneItem.Fitemsize				= db2html(rsget("itemsize"))
				FOneItem.Fitemsource			= db2html(rsget("itemsource"))
				FOneItem.FLastUpdate			= rsget("LastUpdate")
				FOneItem.FSellEndDate			= rsget("sellenddate")
				FOneItem.FItemDiv				= rsget("ItemDiv")
				FOneItem.Fisusing				= rsget("isusing")
				FOneItem.FInterparkPrdNo		= rsget("InterparkPrdNo")
				FOneItem.FmayiParkSellYn		= rsget("mayiParkSellYn")
				FOneItem.FdeliveryType			= rsget("deliveryType")
				FOneItem.FdefaultfreeBeasongLimit	= rsget("defaultfreeBeasongLimit")
				FOneItem.FSailYn				= rsget("sailyn")
                FOneItem.FOrgPrice				= rsget("orgprice")
                FOneItem.Finterparkregdate		= rsget("interparkregdate")
                FOneItem.Fdeliverfixday			= rsget("deliverfixday")
                FOneItem.Ffreight_min			= rsget("freight_min")
                FOneItem.Ffreight_max			= rsget("freight_max")
                FOneItem.FlastErrStr			= rsget("lastErrStr")
                FOneItem.Fmayiparkprice			= rsget("mayiparkprice")
                FOneItem.FregOptCnt				= rsget("regOptCnt")
                FOneItem.FMaySoldOut			= rsget("maySoldOut")
                FOneItem.FbasicimageNm 			= rsget("basicimage")
				FOneItem.FAdultType 			= rsget("adulttype")
				FOneItem.FOrderMaxNum 			= rsget("orderMaxNum")
		End If
		rsget.Close
	End Sub

	Public Sub getInterparkNotRegScheduleOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP 100 i.* "
		strSql = strSql & " ,c.makername, uc.socname_kor, uc.defaultfreeBeasongLimit "
		strSql = strSql & " ,c.keywords, c.ordercomment, c.itemcontent, c.sourcearea "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & " ,c.usinghtml, m.CateKey "
		strSql = strSql & " ,isNULL(c.freight_min,0) as freight_min, isNULL(c.freight_max,0) as freight_max "
		strSql = strSql & " ,isNULL(s.regImageName,'') as regImageName "
		strSql = strSql & " FROM [db_item].[dbo].tbl_item i "
		strSql = strSql & " INNER JOIN [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid "
		strSql = strSql & " LEFT JOIN [db_etcmall].[dbo].tbl_interpark_cate_mapping m on i.cate_large = m.tenCateLarge and i.cate_mid = m.tenCateMid and i.cate_small = m.tenCateSmall "
		strSql = strSql & " LEFT JOIN [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid "
		strSql = strSql & " LEFT JOIN [db_item].[dbo].tbl_interpark_reg_item s on i.itemid = s.itemid "
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " where 1=1 "
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.cate_large <> ''  "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and rtrim(ltrim(isNull(i.deliverfixday, ''))) = '' "		'택배(일반)
		strSql = strSql & "	and i.itemdiv in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
		strSql = strSql & " and i.isusing = 'Y'  "
		strSql = strSql & " and i.isExtusing = 'Y' "
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = '"&CMALLNAME&"') "
		strSql = strSql & " and i.itemid NOT IN (SELECT itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = '"&CMALLNAME&"') "
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "	and 'Y' = case	when i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &") ) "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) THEN 'Y' ELSE 'N' END "
		strSql = strSql & " and i.deliverytype not in ('7','6') "
'		strSql = strSql & " and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000))) "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold> 5 ))) "
		strSql = strSql & " and isnull(s.interParkPrdNo, '') = '' "
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CInterparkItem
				FOneItem.Fitemid					= rsget("itemid")
				FOneItem.Fitemname					= LeftB(db2html(rsget("itemname")),255)
				FOneItem.FMakerid					= rsget("makerid")
				FOneItem.Fsellcash					= rsget("sellcash")
				FOneItem.Forgsellcash				= rsget("orgprice")
				FOneItem.Fsourcearea				= LeftB(db2html(rsget("sourcearea")),64)
				FOneItem.FRegdate					= rsget("regdate")
				FOneItem.Fsellyn					= rsget("sellyn")
				FOneItem.Flimityn					= rsget("limityn")
				FOneItem.Flimitno					= rsget("limitno")
				FOneItem.Flimitsold					= rsget("limitsold")
				FOneItem.Fcate_large				= rsget("cate_large")
				FOneItem.Fcate_mid					= rsget("cate_mid")
				FOneItem.Fcate_small				= rsget("cate_small")
				FOneItem.FMakerName					= db2html(rsget("makername"))
				FOneItem.FBrandName					= db2html(rsget("brandname"))
				FOneItem.Foptioncnt					= rsget("optioncnt")
				FOneItem.FBrandNameKor = db2html(rsget("socname_kor"))
			If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
				FOneItem.FMakerName					= FOneItem.FBrandName
			End If
				FOneItem.Fkeywords					= db2html(rsget("keywords"))
				FOneItem.Fbasicimage				= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FregImageName				= rsget("regImageName")
				FOneItem.Fmainimage					= rsget("mainimage")
			If IsNULL(FOneItem.FInfoImage) Then
				FOneItem.FInfoImage				= ",,,,"
			End If
				FOneItem.Fordercomment				= db2html(rsget("ordercomment"))
				FOneItem.FItemContent				= db2html(rsget("itemcontent"))
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,"♂","")
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,"","")
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,"","")
				FOneItem.Fsourcearea				= db2html(rsget("sourcearea"))
				FOneItem.Fvatinclude				= rsget("vatinclude")
			If (rsget("usinghtml") = "N") Then
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,vbcrlf,"<br>")
			End If

				FOneItem.Finterparkdispcategory		= rsget("CateKey")
				FOneItem.Fitemsize					= db2html(rsget("itemsize"))
				FOneItem.Fitemsource				= db2html(rsget("itemsource"))
				FOneItem.FLastUpdate				= rsget("LastUpdate")
				FOneItem.FSellEndDate				= rsget("sellenddate")
				FOneItem.FItemDiv					= rsget("ItemDiv")
				FOneItem.Fisusing					= rsget("isusing")
				FOneItem.FSailYn					= rsget("sailyn")
				FOneItem.FOrgPrice					= rsget("orgprice")
		'2012-11-09 진영 수정(다이어리 상품이면 무료배송
'		2017-11-27 진영 수정 / 조아름 대리님 요청 (다이어리 상품도 30000원 이상 무배, 미만일 때 2500으로 수정요청
'			If (IsNull(rsget("DyItemid")) = "False" and CLng(rsget("sellcash")) > 13000) AND ((rsget("cate_large") = "010") AND (rsget("cate_mid") = "010") OR (rsget("cate_large") = "010") AND (rsget("cate_mid") = "020") OR (rsget("cate_large") = "010") AND (rsget("cate_mid") = "030") ) Then
'				FOneItem.FdeliveryType				= "4"
'			Else
				FOneItem.FdeliveryType				= rsget("deliveryType")
'			End If
				FOneItem.FdefaultfreeBeasongLimit	= rsget("defaultfreeBeasongLimit")
				FOneItem.Fdeliverfixday				= rsget("deliverfixday")
				FOneItem.Ffreight_min				= rsget("freight_min")
				FOneItem.Ffreight_max				= rsget("freight_max")
				FOneItem.FAdultType 				= rsget("adulttype")
		End If
		rsget.Close
	End Sub
End Class

Function getInterparkPrdno(iitemid)
	Dim sqlStr, retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT isNULL(interparkPrdNo,'') as interparkPrdNo "&VbCRLF
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_interpark_reg_Item "&VbCRLF
	sqlStr = sqlStr & " WHERE itemid="&iitemid
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		retVal = rsget("interparkPrdNo")
	End if
	rsget.Close
	If IsNULL(retVal) Then retVal=""
	getInterparkPrdno = retVal
End Function

'// 상품이미지 존재여부 검사
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function
%>
