<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "11stmy"
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CDEFALUT_STOCK = 9999
CONST my11stAPIURL = "http://api.11street.my/rest"
CONST apiKEY = "31d6989cb7c076d3aae4c4e4970dabca"

Class CMy11stItem
	Public FItemid
	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public Fitemname
	Public FNotdb2HTMLitemname
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
	Public FMy11stGoodNo
	Public FMy11stprice
	Public FMy11stSellYn
	Public FregedOptCnt
	Public FAccFailCNT
	Public FMaySoldOut
	Public Fregitemname
	Public FLastErrStr
	Public FRequireMakeDay
	Public FSafetyyn
	Public FSafetyDiv
	Public FSafetyNum
	Public FMy11stStatCD
	Public FinfoDiv
	Public FDeliveryType
	Public FSocname_kor
	Public FbasicimageNm
	Public FRegImageName
	Public FTransItemname
	Public FItemweight
	Public FCateKey
	Public FOptRecordCnt
	Public FExchangeRate
	Public FMultiplerate
	Public FMaySellPrice
	Public FTransSourcearea
	Public FAreaCode11st

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Function IsMayLimitSoldout
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

	'// 품절여부
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	end function

    public function getBasicImage()
        if IsNULL(FbasicImageNm) or (FbasicImageNm="") then Exit function
        getBasicImage = FbasicImageNm
    end function

	Public Function MustPrice()
		Dim formula
		formula = ((FOrgprice * FMultiplerate) / FExchangeRate)		'( 원판매가 * 20% ) / 환율
		MustPrice = CDbl(FormatNumber(formula ,2))
	End Function

	'// 11번가 판매여부 반환
	Public Function getMy11stSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FSellYn="Y" and FIsUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getMy11stSellYn = "Y"
			Else
				getMy11stSellYn = "N"
			End If
		Else
			getMy11stSellYn = "N"
		End If
	End Function

	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Function GetRaiseValue(value)
		If Fix(value) < value Then
			GetRaiseValue = Fix(value) + 1
		Else
			GetRaiseValue = Fix(value)
		End If
	End Function

	Public Function isImageChanged()
		Dim ibuf : ibuf = getBasicImage
		If InStr(ibuf,"-") < 1 Then
			isImageChanged = FALSE
			Exit Function
		End If
		isImageChanged = ibuf <> FRegImageName
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

    Public Function getMy11stAddImageParam()
    	Dim strRst, strSQL, i
    	strRst = ""
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=2 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					strRst = strRst & "	<prdImage0"&i&"><![CDATA[http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")&"]]></prdImage0"&i&">"
				End If
				rsget.MoveNext
				If i >= 4 Then Exit For
			Next
		End If
		rsget.Close
		getMy11stAddImageParam = strRst
    End Function

	Public Function getMy11stSaleItemParam
		Dim strRst, strSQL, i
		strRst = ""
		If date() <= "2017-04-06" Then

			If (Fitemid = "1396455") OR (Fitemid = "1452878") OR (Fitemid = "1378681") OR (Fitemid = "1259852") Then
				strRst = strRst & "	<cuponcheck>Y</cuponcheck>"								'?			'Enable/Disable discount program | Y : Use, N : Not Use
				strRst = strRst & "	<dscAmtPercnt>30</dscAmtPercnt>"						'?			'Discount value
				strRst = strRst & "	<cupnDscMthdCd>02</cupnDscMthdCd>"						'?			'Discount unit code | 01 : RM (in ringgit), 02 : % (in percentage)
				strRst = strRst & "	<cupnIssStartDy>02/03/2017</cupnIssStartDy>"			'?			'Start date of Limit the discount program (DD/MM/YYYY)
				strRst = strRst & "	<cupnIssEndDy>06/04/2017</cupnIssEndDy>"				'?			'End date of Limit the discount program (DD/MM/YYYY)
				strRst = strRst & "	<cupnUseLmtDyYn>Y</cupnUseLmtDyYn>"						'?			'Limit the discount program (in days) | Y : Use, N : Not Use
			ElseIf (Fitemid = "1130741") OR (Fitemid = "1419077") OR (Fitemid = "1304358") OR (Fitemid = "589447") Then
				strRst = strRst & "	<cuponcheck>Y</cuponcheck>"								'?			'Enable/Disable discount program | Y : Use, N : Not Use
				strRst = strRst & "	<dscAmtPercnt>20</dscAmtPercnt>"						'?			'Discount value
				strRst = strRst & "	<cupnDscMthdCd>02</cupnDscMthdCd>"						'?			'Discount unit code | 01 : RM (in ringgit), 02 : % (in percentage)
				strRst = strRst & "	<cupnIssStartDy>02/03/2017</cupnIssStartDy>"			'?			'Start date of Limit the discount program (DD/MM/YYYY)
				strRst = strRst & "	<cupnIssEndDy>06/04/2017</cupnIssEndDy>"				'?			'End date of Limit the discount program (DD/MM/YYYY)
				strRst = strRst & "	<cupnUseLmtDyYn>Y</cupnUseLmtDyYn>"						'?			'Limit the discount program (in days) | Y : Use, N : Not Use
			ElseIf (Fitemid = "1576077") OR (Fitemid = "1350012") OR (Fitemid = "1291761") OR (Fitemid = "1594408") OR (Fitemid = "1594407") OR (Fitemid = "1594406") OR (Fitemid = "1594405") OR (Fitemid = "1594404") OR (Fitemid = "1541342") OR (Fitemid = "1541344") OR (Fitemid = "1541334") OR (Fitemid = "1342237") OR (Fitemid = "1541336") OR (Fitemid = "1350015") OR (Fitemid = "1485267") OR (Fitemid = "1291762") OR (Fitemid = "1350014") OR (Fitemid = "1350013") OR (Fitemid = "1485266") OR (Fitemid = "1291763") OR (Fitemid = "1350016") OR (Fitemid = "1541339") Then
				strRst = strRst & "	<cuponcheck>Y</cuponcheck>"								'?			'Enable/Disable discount program | Y : Use, N : Not Use
				strRst = strRst & "	<dscAmtPercnt>15</dscAmtPercnt>"						'?			'Discount value
				strRst = strRst & "	<cupnDscMthdCd>02</cupnDscMthdCd>"						'?			'Discount unit code | 01 : RM (in ringgit), 02 : % (in percentage)
				strRst = strRst & "	<cupnIssStartDy>02/03/2017</cupnIssStartDy>"			'?			'Start date of Limit the discount program (DD/MM/YYYY)
				strRst = strRst & "	<cupnIssEndDy>06/04/2017</cupnIssEndDy>"				'?			'End date of Limit the discount program (DD/MM/YYYY)
				strRst = strRst & "	<cupnUseLmtDyYn>Y</cupnUseLmtDyYn>"						'?			'Limit the discount program (in days) | Y : Use, N : Not Use
			Else
				If (Fitemid <> "1629288") OR (Fitemid <> "1629289") OR (Fitemid <> "1629287") OR (Fitemid <> "1615290") OR (Fitemid <> "1629290") OR (Fitemid <> "1629286") OR (Fitemid <> "1615312") OR (Fitemid <> "1615289") OR (Fitemid <> "1617211") OR (Fitemid <> "1615288") OR (Fitemid <> "1615313") Then
					strRst = strRst & "	<cuponcheck>Y</cuponcheck>"								'?			'Enable/Disable discount program | Y : Use, N : Not Use
					strRst = strRst & "	<dscAmtPercnt>10</dscAmtPercnt>"						'?			'Discount value
					strRst = strRst & "	<cupnDscMthdCd>02</cupnDscMthdCd>"						'?			'Discount unit code | 01 : RM (in ringgit), 02 : % (in percentage)
					strRst = strRst & "	<cupnIssStartDy>15/02/2017</cupnIssStartDy>"			'?			'Start date of Limit the discount program (DD/MM/YYYY)
					strRst = strRst & "	<cupnIssEndDy>06/04/2017</cupnIssEndDy>"				'?			'End date of Limit the discount program (DD/MM/YYYY)
					strRst = strRst & "	<cupnUseLmtDyYn>Y</cupnUseLmtDyYn>"						'?			'Limit the discount program (in days) | Y : Use, N : Not Use
				End If
			End If
		End If

'		strRst = ""
'		'strRst = strRst & "	<cuponcheck>N</cuponcheck>"							'?			'Enable/Disable discount program | Y : Use, N : Not Use
'		strRst = strRst & "	<cuponcheck>Y</cuponcheck>"								'?			'Enable/Disable discount program | Y : Use, N : Not Use
''		If (Fitemid = "1497868") OR (Fitemid = "1497868") OR (Fitemid = "1497868") OR (Fitemid = "1497868") Then
''			strRst = strRst & "	<dscAmtPercnt>15</dscAmtPercnt>"					'?			'Discount value
''		Else
'			strRst = strRst & "	<dscAmtPercnt>10</dscAmtPercnt>"					'?			'Discount value
''		End If
'		strRst = strRst & "	<cupnDscMthdCd>02</cupnDscMthdCd>"						'?			'Discount unit code | 01 : RM (in ringgit), 02 : % (in percentage)
'		strRst = strRst & "	<cupnIssStartDy>03/08/2016</cupnIssStartDy>"			'?			'Start date of Limit the discount program (DD/MM/YYYY)
'		strRst = strRst & "	<cupnIssEndDy>10/08/2016</cupnIssEndDy>"				'?			'End date of Limit the discount program (DD/MM/YYYY)
'		strRst = strRst & "	<cupnUseLmtDyYn>Y</cupnUseLmtDyYn>"						'?			'Limit the discount program (in days) | Y : Use, N : Not Use
		getMy11stSaleItemParam = strRst
	End Function

	Public Function getMy11stContParamToReg()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
		'2014-01-17 10:00 김진영 탑 이미지 추가
		strRst = strRst & ("<a href=""http://www.11street.my/store/MiniMallAction/getMiniMallHome.do?sellerHmpgUrl=10x10"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_11stmy.jpg""></a></div><br>")
		strRst = strRst & ("<div align=""center"">")
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
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_11stmy.jpg"">")
		strRst = strRst & ("</div>")
		getMy11stContParamToReg = strRst
	End Function

	'최대 구매 수량
	Public Function getLimitMy11stEa()
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
		getLimitMy11stEa = ret
	End Function

	Public Function getMy11stOptParamtoReg(vMy11stgoodno)
		Dim strRst, strSql, vCount, i, optLimit, optDc, optaddprice, formula, optRateaddprice
		strRst = ""
		If FOptioncnt = 0 Then
			strRst = strRst & "	<optSelCnt>0</optSelCnt>"					'Count the Select type option | 0 : 0EA, 1 : 1EA(Required productSelOption and colTitle), 2 : 2EA(Required productSelOption and colTitle)
			FOptRecordCnt = 0
		Else
			strRst = strRst & "	<optSelCnt>1</optSelCnt>"					'Count the Select type option | 0 : 0EA, 1 : 1EA(Required productSelOption and colTitle), 2 : 2EA(Required productSelOption and colTitle)
			strSql = ""
			strSql = strSql & " SELECT o.itemoption, o.isusing, o.optsellyn, o.optaddprice, mo.optionTypeName, mo.optionname, (o.optlimitno-o.optlimitsold) as optLimit "
			strSql = strSql & " from db_item.dbo.tbl_item_option as o "
			strSql = strSql & " JOIN [db_item].[dbo].[tbl_item_multiLang_option] as mo on o.itemid = mo.itemid and o.itemoption = mo.itemoption and mo.countryCd = 'EN' "
			If vMy11stgoodno <> "" Then		'2016-08-29 김진영 수정
				strSql = strSql & " JOIN db_item.dbo.tbl_OutMall_regedoption as R on o.itemid = R.itemid and o.itemoption = R.itemoption and mallid = '11stmy' "
			End If
			strSql = strSql & " WHERE o.itemid = '"&FItemid&"' "
			strSql = strSql & " and o.isusing = 'Y' "
			strSql = strSql & " and o.optsellyn = 'Y' "
			strSql = strSql & " and mo.isusing = 'Y' "
			rsget.Open strSql,dbget,1
			vCount = rsget.RecordCount
			FOptRecordCnt = rsget.RecordCount
			If Not(rsget.EOF or rsget.BOF) then
				For i = 0 to vCount - 1
					If i = 0 Then
						strRst = strRst & "	<colTitle><![CDATA["&rsget("optionTypeName")&"]]></colTitle>"	'Option name
					End If

					optLimit = rsget("optLimit")
					optLimit = optLimit - 5
					If (optLimit < 1) Then optLimit = 0
					If (Flimityn <> "Y") Then optLimit = 9999
					optDc		= db2Html(rsget("optionname"))
					optaddprice	= rsget("optaddprice")

					If optaddprice > 0 Then
						formula = ((optaddprice * FMultiplerate) / FExchangeRate)		'( 옵션가 * 20% ) / 환율
						optRateaddprice = CDbl(FormatNumber(formula ,2))
					Else
						optRateaddprice = 0
					End If

					strRst = strRst & "	<productSelOption>"													'Select type option(Node)
					strRst = strRst & "		<colCount>"&optLimit&"</colCount>"								'	Count of Option
					strRst = strRst & "		<colValue0><![CDATA["&optDc&"]]></colValue0>"					'	Option value
					strRst = strRst & "		<optPrc>"&optRateaddprice&"</optPrc>"							'	Additional Price | 옵션추가금액 어쩔?
				If optLimit > 0 Then
					strRst = strRst & "		<optStatus>01</optStatus>"										'	Display of option | 01 : display, 02 : Not display
				Else
					strRst = strRst & "		<optStatus>02</optStatus>"										'	Display of option | 01 : display, 02 : Not display
				End If
					strRst = strRst & "		<optWght>0</optWght>"											'	Additional Weight | 추가 무게 어쩔?
					strRst = strRst & "	</productSelOption>"
					rsget.moveNext
				Next
			End If
		End If
		getMy11stOptParamtoReg = strRst
	End Function

	Public Function getMy11stOptParamtoEDT
		Dim strRst, strSql, vCount, i, optLimit, optDc, optaddprice, formula, optRateaddprice
		strRst = ""
		If FOptioncnt = 0 Then
			strRst = strRst & "	<optSelCnt>0</optSelCnt>"					'Count the Select type option | 0 : 0EA, 1 : 1EA(Required productSelOption and colTitle), 2 : 2EA(Required productSelOption and colTitle)
		Else
			strRst = strRst & "	<optSelCnt>1</optSelCnt>"					'Count the Select type option | 0 : 0EA, 1 : 1EA(Required productSelOption and colTitle), 2 : 2EA(Required productSelOption and colTitle)
			strSql = ""
			strSql = strSql & " SELECT o.itemoption, o.isusing, o.optsellyn, o.optaddprice, mo.optionTypeName, mo.optionname, (o.optlimitno-o.optlimitsold) as optLimit "
			strSql = strSql & " from db_item.dbo.tbl_item_option as o "
			strSql = strSql & " JOIN [db_item].[dbo].[tbl_item_multiLang_option] as mo on o.itemid = mo.itemid and o.itemoption = mo.itemoption and mo.countryCd = 'EN' "
			strSql = strSql & " WHERE o.itemid = '"&FItemid&"' "
			strSql = strSql & " and o.isusing = 'Y' "
			strSql = strSql & " and o.optsellyn = 'Y' "
			strSql = strSql & " and mo.isusing = 'Y' "
			rsget.Open strSql,dbget,1
			vCount = rsget.RecordCount
			FOptRecordCnt = rsget.RecordCount
			If Not(rsget.EOF or rsget.BOF) then
				For i = 0 to vCount - 1
					If i = 0 Then
						strRst = strRst & "	<colTitle><![CDATA["&rsget("optionTypeName")&"]]></colTitle>"	'Option name
					End If

					optLimit = rsget("optLimit")
					optLimit = optLimit - 5
					If (optLimit < 1) Then optLimit = 0
					If (Flimityn <> "Y") Then optLimit = 9999
					optDc		= db2Html(rsget("optionname"))
					optaddprice	= rsget("optaddprice")

					If optaddprice > 0 Then
						formula = ((optaddprice * FMultiplerate) / FExchangeRate)		'( 옵션가 * 20% ) / 환율
						optRateaddprice = CDbl(FormatNumber(formula ,2))
					Else
						optRateaddprice = 0
					End If

					strRst = strRst & "	<productSelOption>"													'Select type option(Node)
					strRst = strRst & "		<colCount>"&optLimit&"</colCount>"								'	Count of Option
					strRst = strRst & "		<colValue0><![CDATA["&optDc&"]]></colValue0>"					'	Option value
					strRst = strRst & "		<optPrc>"&optRateaddprice&"</optPrc>"			'?				'	Additional Price | 옵션추가금액 어쩔?
				If optLimit > 0 Then
					strRst = strRst & "		<optStatus>01</optStatus>"										'	Display of option | 01 : display, 02 : Not display
				Else
					strRst = strRst & "		<optStatus>02</optStatus>"										'	Display of option | 01 : display, 02 : Not display
				End If
					strRst = strRst & "		<optWght>0</optWght>"							'?				'	Additional Weight | 추가 무게 어쩔?
					strRst = strRst & "	</productSelOption>"
					rsget.moveNext
				Next
			End If
		End If
		getMy11stOptParamtoEDT = strRst
	End Function

	'옵션 수정 XML
	Public Function getMy11stOptEditXML
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst & "	<Product>"
	If FItemdiv = "06" Then
		strRst = strRst & "	<optWrtCnt>1</optWrtCnt>"											'주문제작문구 옵션 수 | 0 : 0EA, 1 : 1EA(Required productWrtOption), 2 : 2EA (Required productWrtOption)
		strRst = strRst & "	<productWrtOption>"
		strRst = strRst & "		<colValue0>Please leave a sentence for carving service.</colValue0>"
		strRst = strRst & "	</productWrtOption>"
	Else
		strRst = strRst & "	<optWrtCnt>0</optWrtCnt>"											'주문제작문구 옵션 수
	End If
		strRst = strRst & getMy11stOptParamtoEDT
		strRst = strRst & "	</Product>"
		getMy11stOptEditXML = strRst
	End Function

	Public Function getMy11stItemRegXML(imy11stGoodno)
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		strRst = strRst & "<Product>"
		If imy11stGoodno <> "" Then
			strRst = strRst & "	<prdNo>"&imy11stGoodno&"</prdNo>"								'수정시 필수 | 업체상품코드
		End If
		strRst = strRst & "	<selMthdCd>01</selMthdCd>"											'#Sales Type | 01 : Ready Stock
		strRst = strRst & "	<dispCtgrNo>"&FCateKey&"</dispCtgrNo>"								'#Category ID
		strRst = strRst & "	<prdTypCd>01</prdTypCd>"											'#Service Type | 01 : General Product, 25 : e-voucher
		strRst = strRst & "	<prdNm><![CDATA["&FTransItemname&"]]></prdNm>"						'#ItemName
		strRst = strRst & "	<prdStatCd>01</prdStatCd>"											'#Item conditions | 01 : New
		strRst = strRst & "	<prdWght>"&CDBL(FormatNumber((FitemWeight/1000),3))&"</prdWght>"	'#Item weight in kilograms | 소수점 3자리까지
		strRst = strRst & "	<minorSelCnYn>Y</minorSelCnYn>"										'#Minors can buy (Under 18 years old).If set as N, no image is displayed to guest user and user under than 18 years old.
		strRst = strRst & "	<prdImage01><![CDATA["&Fbasicimage&"]]></prdImage01>"				'#Main(Representative) Image URL
		strRst = strRst & getMy11stAddImageParam()												'Additional Image 1 URL, Additional Image 2 URL, Additional Image 3 URL
		strRst = strRst & "	<htmlDetail><![CDATA["&getMy11stContParamToReg&"]]></htmlDetail>"	'#Item’s Detailed Description(html format supported)
'		strRst = strRst & "	<advrtStmt></advrtStmt>"											'Item’s advertisement information. e.g: “Hot item’s of this week”
		strRst = strRst & "	<orgnTypCd>02</orgnTypCd>"											'Code of Country of Origin | 01 : Domestic, 02 : Overseas
		strRst = strRst & "	<orgnTypDtlsCd>"&FAreaCode11st&"</orgnTypDtlsCd>"					'Regional Code of Country of Origin
		strRst = strRst & "	<sellerPrdCd>"&FItemid&"</sellerPrdCd>"								'Seller’s Item/Product Code
		strRst = strRst & "	<reviewDispYn>Y</reviewDispYn>"										'Whether display the Item Comments/ Review or not | Y : display, N : Not display
		strRst = strRst & "	<reviewOptDispYn>Y</reviewOptDispYn>"								'Enable/Disable the Review/Comment in the item’s page | Y : Enable, N : Disable
		strRst = strRst & "	<selTermUseYn>N</selTermUseYn>"										'#Whether to use sales period(Y/N) | Y : Use, N : Not Use
'		strRst = strRst & "	<selPrdClfFpCd></selPrdClfFpCd>"									'selTermUseYn가 Y면 필수 (판매기간인 듯)
'		strRst = strRst & "	<aplBgnDy></aplBgnDy>"												'판매기간설정시 시작일
'		strRst = strRst & "	<aplEndDy></aplEndDy>"												'판매기간설정시 종료일
'		strRst = strRst & "	<wrhsPlnDy></wrhsPlnDy>"											'Due date of Stock (DD/MM/YYYY)
		strRst = strRst & "	<selPrc>"&FMaySellPrice&"</selPrc>"									'#Product’s price | 소수점 2자리까지
		strRst = strRst & "	<prdSelQty>"&getLimitMy11stEa&"</prdSelQty>"						'#Set the product stocks
		strRst = strRst & getMy11stSaleItemParam()												'할인관리
'		strRst = strRst & "	<pointYN>Y</pointYN>"												'Enable/Disable credit giveaway | Y : Enable, N : Disable
'		strRst = strRst & "	<pointValue>100</pointValue>"										'Credit value to give to buyer
'		strRst = strRst & "	<spplWyCd>02</spplWyCd>"											'Credit unit code | 01 : % (in percentage), 02 : RM (in ringgit)
	If FItemdiv = "06" Then
		strRst = strRst & "	<optWrtCnt>1</optWrtCnt>"											'주문제작문구 옵션 수 | 0 : 0EA, 1 : 1EA(Required productWrtOption), 2 : 2EA (Required productWrtOption)
		strRst = strRst & "	<productWrtOption>"
		strRst = strRst & "		<colValue0>Please leave a sentence for carving service.</colValue0>"
		strRst = strRst & "	</productWrtOption>"
	Else
		strRst = strRst & "	<optWrtCnt>0</optWrtCnt>"											'주문제작문구 옵션 수
	End If
		strRst = strRst & getMy11stOptParamtoReg(imy11stGoodno)
'		strRst = strRst & "	<ProductComponent>"													'Additional setting for Item
'		strRst = strRst & "		<addCompPrc>5.50</addCompPrc>"									'	Additional setting for price
'		strRst = strRst & "		<addPrdGrpNm>Add Product 1</addPrdGrpNm>"						'	Additional Item name
'		strRst = strRst & "		<addPrdWght>0.234</addPrdWght>"									'	Additional Weight setting
'		strRst = strRst & "		<addUseYn>Y</addUseYn>"											'	Status | Y : display, N : Not display
'		strRst = strRst & "		<compPrdNm>1G</compPrdNm>"										'	Additional Data type
'		strRst = strRst & "		<compPrdQty>5</compPrdQty>"										'	Additional Product count
'		strRst = strRst & "	</ProductComponent>"
'		strRst = strRst & "	<ProductComponent>"
'		strRst = strRst & "		<addCompPrc>5.50</addCompPrc>"
'		strRst = strRst & "		<addPrdGrpNm>Add Product 2</addPrdGrpNm>"
'		strRst = strRst & "		<addPrdWght>0.234</addPrdWght>"
'		strRst = strRst & "		<addUseYn>Y</addUseYn>"
'		strRst = strRst & "		<compPrdNm>2G</compPrdNm>"
'		strRst = strRst & "		<compPrdQty>5</compPrdQty>"
'		strRst = strRst & "	</ProductComponent>"
'		strRst = strRst & "	<ProductComponent>"
'		strRst = strRst & "		<addCompPrc>5.50</addCompPrc>"
'		strRst = strRst & "		<addPrdGrpNm>Add Product 3</addPrdGrpNm>"
'		strRst = strRst & "		<addPrdWght>0.234</addPrdWght>"
'		strRst = strRst & "		<addUseYn>Y</addUseYn>"
'		strRst = strRst & "		<compPrdNm>3G</compPrdNm>"
'		strRst = strRst & "		<compPrdQty>5</compPrdQty>"
'		strRst = strRst & "	</ProductComponent>"
		strRst = strRst & "	<selMinLimitTypCd>00</selMinLimitTypCd>"							'Minimum purchase quantity code (see code reference) | 00 : Not use, 01 : Per order
'		strRst = strRst & "	<selMinLimitQty></selMinLimitQty>"									'Amount of minimum purchase quantity
		strRst = strRst & "	<selLimitTypCd>00</selLimitTypCd>"									'Maximum purchase amount code (see code reference) | 00 : Not use, 01 : Per order, 02 : Per person(ID)
'		strRst = strRst & "	<selLimitQty></selLimitQty>"										'Amount of maximum purchase quantity
		strRst = strRst & "	<asDetail><![CDATA[Please contact us by email (email : csglobal@10x10.co.kr ) or use the Customer questions board.]]></asDetail>"	'#After service information. Could be the address or contact info of the aftersales Service
		strRst = strRst & "	<dlvMthCd>01</dlvMthCd>"											'#Shipping Method | 01 : Courier service(택배), 02 : Direct shipping(직접)
		strRst = strRst & "	<dlvCstInstBasiCd>12</dlvCstInstBasiCd>"							'#Delivery type | 01 : Free, 11 : Shipping rate by product(Required ProductDlvTariff and productDlvPrmt), 12 : Bundle shipping fee (default setting of seller's ProductDlvTariff and productDlvPrmt)
'		strRst = strRst & "	<ProductDlvTariff>"										'?			'Delivery Tariff infomation
'		strRst = strRst & "		<addPrc>1000.12</addPrc>"										'	Additional Delivery Price
'		strRst = strRst & "		<addWght>1000.120</addWght>"									'	Additional Delivery Weight
'		strRst = strRst & "		<basePrc>1000</basePrc>"										'	Basic Delivery Price
'		strRst = strRst & "		<baseWght>1000</baseWght>"										'	Basic Delivery Weight
'		strRst = strRst & "		<dlvDstnCd>01</dlvDstnCd>"										'	Shipping Area | 01 : WestMalaysia, 02 : Sabah/Labuan, 03 : Sarawak
'		strRst = strRst & "	</ProductDlvTariff>"
'		strRst = strRst & "	<ProductDlvTariff>"
'		strRst = strRst & "		<addPrc>1000.12</addPrc>"
'		strRst = strRst & "		<addWght>1000.120</addWght>"
'		strRst = strRst & "		<basePrc>1000</basePrc>"
'		strRst = strRst & "		<baseWght>1000</baseWght>"
'		strRst = strRst & "		<dlvDstnCd>02</dlvDstnCd>"
'		strRst = strRst & "	</ProductDlvTariff>"
'		strRst = strRst & "	<ProductDlvTariff>"
'		strRst = strRst & "		<addPrc>1000.12</addPrc>"
'		strRst = strRst & "		<addWght>1000.120</addWght>"
'		strRst = strRst & "		<basePrc>1000</basePrc>"
'		strRst = strRst & "		<baseWght>1000</baseWght>"
'		strRst = strRst & "		<dlvDstnCd>03</dlvDstnCd>"
'		strRst = strRst & "	</ProductDlvTariff>"
'		strRst = strRst & "	<productDlvPrmt>"										'안해도됨	'Delivery Tariff promotion
'		strRst = strRst & "		<cfsDscAmt>1000</cfsDscAmt>"									'	Delivery Tariff Discount Price
'		strRst = strRst & "		<dlvPrmtCd>CD</dlvPrmtCd>"										'	Delivery Tariff promotion code | NA : Not use, CD : Discount of Conditional, CF : Free of Conditional
'		strRst = strRst & "		<pdBuyAmt>1000</pdBuyAmt>"										'	Buy Price
'		strRst = strRst & "	</productDlvPrmt>"
		strRst = strRst & "	<rtngExchDetail><![CDATA[No exchange/refund for international shipment. Please contact us by email (email : csglobal@10x10.co.kr )]]></rtngExchDetail>"		'#Return/Exchange information
		strRst = strRst & "	<suplDtyfrPrdClfCd>02</suplDtyfrPrdClfCd>"				'02라함		'#GST code | 01:Standard rate, 02 : Exempted rate, 03 : Zero rate, 04 : Flat rate
'		strRst = strRst & "	<suplDtyfrPrdClfRate>2.56</suplDtyfrPrdClfRate>"		'02므로무시	'GST Flat rate percent | 선택적 필수
'		strRst = strRst & "	<createCd></createCd>"												'Code Assigned to a Company Listing Products. If you are an affiliate partner that lists products on behalf of seller, you need to get this code issued first. You can get the code from 11street administrator.
		strRst = strRst & "</Product>"
		getMy11stItemRegXML = strRst
	End Function
End Class

Class CMy11st
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
	Public FRectGubun

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

	Public Sub getmy11stNotRegOneItem
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & " select itemid from ("
            addSql = addSql & "     select itemid"
            addSql = addSql & " 	,count(*) as optCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	from db_item.dbo.tbl_item_option"
            addSql = addSql & " 	where itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and isusing='Y'"
            addSql = addSql & " 	group by itemid"
            addSql = addSql & " ) T"
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, isNULL(R.my11stStatCD,-9) as my11stStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum, uc.socname_kor "
		strSql = strSql & " ,isNULL(R.regImageName,'') as regImageName, m.itemname as transItemname, m.sourcearea as transSourcearea "
		strSql = strSql & "	, isnull(bm.CateKey, '') as CateKey, ex.exchangeRate, ex.multiplerate, uu.orgprice as maySellPrice, m.areaCode11st "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_multiLang_price uu on i.itemid = uu.itemid and uu.sitename = '11STMY' "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		strSql = strSql & " JOIN db_item.dbo.tbl_exchangeRate as ex on uu.sitename = ex.sitename and ex.countryLangCD = m.countrycd "
		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_my11st_cate_mapping] as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_my11st_regItem] R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1  "
		strSql = strSql & " and i.isusing = 'Y' "
		strSql = strSql & " and i.itemdiv <> '21' "
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.deliverOverseas = 'Y' "		'해외배송상품 Y
		strSql = strSql & " and i.itemweight <> 0 "				'무게는 0보다 커야
'		strSql = strSql & " and i.mwdiv in ('m', 'w') "			'매입 or 위탁
'		strSql = strSql & " and i.deliverytype in (1 ,4) "		'텐배 or 텐무료배
		strSql = strSql & " and i.itemid not in (select itemid from db_item.[dbo].[tbl_item_option] Where optaddprice > 0 group by itemid) "		'옵션 중 추가금액 제외
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and isnull(R.my11stStatCD, 0) < 3 "
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & addSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CMy11stItem
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
                FOneItem.FMy11stStatCD		= rsget("my11stStatCD")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FCateKey			= rsget("CateKey")
                FOneItem.FSocname_kor		= rsget("socname_kor")
                FOneItem.FbasicimageNm 		= rsget("basicimage")
                FOneItem.FRegImageName 		= rsget("regImageName")
                FOneItem.FTransItemname 	= rsget("transItemname")
                FOneItem.FItemweight 		= rsget("itemweight")
                FOneItem.FExchangeRate 		= rsget("exchangeRate")
                FOneItem.FMultiplerate 		= rsget("multiplerate")
                FOneItem.FMaySellPrice 		= rsget("maySellPrice")
                FOneItem.FTransSourcearea 	= rsget("transSourcearea")
                FOneItem.FAreaCode11st 		= rsget("areaCode11st")
		End If
		rsget.Close
	End Sub

	Public Sub getmy11stlEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, m.my11stGoodNo, m.my11stprice, m.my11stSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName, ml.itemname as transItemname, ml.sourcearea as transSourcearea  "
		strSql = strSql & "	, isnull(bm.CateKey, '') as CateKey, ex.exchangeRate, ex.multiplerate, uu.orgprice as maySellPrice, ml.areaCode11st "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
        strSql = strSql & "	,(CASE WHEN i.isusing = 'N' "
		strSql = strSql & "		or i.sellyn <> 'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or ((i.LimitYn = 'Y') and (i.LimitNo - i.LimitSold <= "&CMAXLIMITSELL&")) "
		strSql = strSql & "		or i.deliverOverseas <> 'Y' "
		strSql = strSql & "		or i.itemweight = 0 "
		strSql = strSql & "		or i.itemdiv = '21' "
'		strSql = strSql & " 	or (i.mwdiv not in ('m', 'w')) "
'		strSql = strSql & " 	or (i.deliverytype not in (1 ,4)) "
		strSql = strSql & " 	or i.itemid in (select itemid from db_item.[dbo].[tbl_item_option] Where optaddprice > 0 group by itemid) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_my11st_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_multiLang_price uu on i.itemid = uu.itemid and uu.sitename = '11STMY' "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_item_multiLang] as ml on i.itemid = ml.itemid and ml.countrycd = 'EN'  "
		strSql = strSql & " JOIN db_item.dbo.tbl_exchangeRate as ex on uu.sitename = ex.sitename and ex.countryLangCD = ml.countrycd  "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_my11st_cate_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and m.my11stStatCD = 7 "
		strSql = strSql & addSql
		strSql = strSql & " and m.my11stGoodNo is Not Null "									'#등록 상품만
'rw strSql
'response.end
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CMy11stItem
				FOneItem.Fitemid				= rsget("itemid")
				FOneItem.FtenCateLarge			= rsget("cate_large")
				FOneItem.FtenCateMid			= rsget("cate_mid")
				FOneItem.FtenCateSmall			= rsget("cate_small")
				FOneItem.Fitemname				= db2html(rsget("itemname"))
				FOneItem.FNotdb2HTMLitemname	= rsget("itemname")
				FOneItem.FitemDiv				= rsget("itemdiv")
				FOneItem.FsmallImage			= rsget("smallImage")
				FOneItem.Fmakerid				= rsget("makerid")
				FOneItem.Fregdate				= rsget("regdate")
				FOneItem.FlastUpdate			= rsget("lastUpdate")
				FOneItem.ForgPrice				= rsget("orgPrice")
				FOneItem.ForgSuplyCash			= rsget("orgSuplyCash")
				FOneItem.FSellCash				= rsget("sellcash")
				FOneItem.FBuyCash				= rsget("buycash")
				FOneItem.FsellYn				= rsget("sellYn")
				FOneItem.FsaleYn				= rsget("sailyn")
				FOneItem.FisUsing				= rsget("isusing")
				FOneItem.FLimitYn				= rsget("LimitYn")
				FOneItem.FLimitNo				= rsget("LimitNo")
				FOneItem.FLimitSold				= rsget("LimitSold")
				FOneItem.Fkeywords				= rsget("keywords")
				FOneItem.ForderComment			= db2html(rsget("ordercomment"))
				FOneItem.FoptionCnt				= rsget("optionCnt")
				FOneItem.FbasicImage			= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage				= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2			= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Fsourcearea			= rsget("sourcearea")
				FOneItem.Fmakername				= rsget("makername")
				FOneItem.FUsingHTML				= rsget("usingHTML")
				FOneItem.Fitemcontent			= db2html(rsget("itemcontent"))
				FOneItem.FMy11stGoodNo			= rsget("my11stGoodNo")
				FOneItem.FMy11stprice			= rsget("my11stprice")
				FOneItem.FMy11stSellYn			= rsget("my11stSellYn")

	            FOneItem.FoptionCnt				= rsget("optionCnt")
	            FOneItem.FregedOptCnt			= rsget("regedOptCnt")
	            FOneItem.FaccFailCNT			= rsget("accFailCNT")
	            FOneItem.FlastErrStr			= rsget("lastErrStr")
	            FOneItem.Fdeliverytype			= rsget("deliverytype")
	            FOneItem.FrequireMakeDay		= rsget("requireMakeDay")

	            FOneItem.FinfoDiv				= rsget("infoDiv")
	            FOneItem.Fsafetyyn				= rsget("safetyyn")
	            FOneItem.FsafetyDiv				= rsget("safetyDiv")
	            FOneItem.FsafetyNum				= rsget("safetyNum")
	            FOneItem.FmaySoldOut			= rsget("maySoldOut")
	            FOneItem.Fregitemname			= rsget("regitemname")
                FOneItem.FregImageName			= rsget("regImageName")
                FOneItem.FbasicImageNm			= rsget("basicimage")
                FOneItem.FCateKey				= rsget("CateKey")
                FOneItem.FTransItemname 		= rsget("transItemname")
                FOneItem.FItemweight 			= rsget("itemweight")

                FOneItem.FExchangeRate 			= rsget("exchangeRate")
                FOneItem.FMultiplerate 			= rsget("multiplerate")
				FOneItem.FMaySellPrice 			= rsget("maySellPrice")
				FOneItem.FTransSourcearea 		= rsget("transSourcearea")
				FOneItem.FAreaCode11st 			= rsget("areaCode11st")
		End If
		rsget.Close
	End Sub

	Function getchangeOptionNameCnt(iitemid)
		Dim strSql
		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as cnt "
		strSql = strSql & " FROM [db_item].[dbo].[tbl_item_multiLang_option] "
		strSql = strSql & " WHERE isnull(optionname, '') <> '' "
		strSql = strSql & " and isnull(optionTypeName, '') <> '' "
		strSql = strSql & " and isusing = 'Y' "
		strSql = strSql & " and itemid = '"&iitemid&"' "
		strSql = strSql & " and countryCd = 'EN' "
		rsget.Open strSql, dbget, 1
			getchangeOptionNameCnt = rsget("cnt")
		rsget.Close
	End Function
End Class

'// 상품이미지 존재여부 검사
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function

Function getMy11stGoodNo(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 my11stGoodNo FROM db_etcmall.[dbo].[tbl_my11st_regItem] WHERE itemid = '"&iitemid&"' "
	rsget.Open strSql, dbget, 1
	If not rsget.EOF Then
		getMy11stGoodNo = rsget("my11stGoodNo")
	Else
		getMy11stGoodNo = ""
	End If
	rsget.Close
End Function

Function getMy11stRatePrice(iitemid, byref vOrgprice, byref vExchangeRate, byref vMultiplerate, byref vMaySellPrice)
	Dim strSql, formula, exchangeRate, multiplerate
	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.orgprice, p.orgprice as maySellprice, e.exchangeRate, e.multiplerate "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_item.dbo.tbl_item_multiLang_price as p on i.itemid = p.itemid and p.sitename = '11STMY' "
	strSql = strSql & " JOIN db_item.dbo.tbl_exchangeRate as e on p.sitename = e.sitename "
	strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
	rsget.Open strSql, dbget, 1
	If not rsget.EOF Then
		vOrgprice		= rsget("orgprice")
		vExchangeRate	= rsget("exchangeRate")	'환율
		vMultiplerate	= rsget("multiplerate")	'배수
		vMaySellPrice	= rsget("maySellprice")	'해외판매가
	End If
	rsget.Close
End Function
%>