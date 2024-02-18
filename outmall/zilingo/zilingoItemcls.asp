<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "zilingo"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST zilingoAPIURL = "https://api.sellers.zilingo.com"
CONST zilingoSELLERID = "SEL2532759781"
CONST zilingoAPIKEY = "SAK-1508201128366-FAC9E976DEFB43D15F372829546DC"
CONST zilingoLOCALE = "en"
CONST CDEFALUT_STOCK = 999

Class CZilingoItem
	Public Fitemid
	Public FZilingoGoodNo
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
	Public FZilingoSkuGoodNo
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
	Public FListimage
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public FSafetyNum
	Public Fitemcontent
	Public FZilingoStatCD
	Public Fdeliverfixday
	Public Fdeliverytype
	Public FCateKey
	Public FRegImageName
	Public FmaySoldOut
	Public FDisplayDate
	Public Fregitemname
	Public FbasicImageNm
	Public FChgItemname
	Public FChgItemContent
	Public FChgItemsource
	Public FChgItemsize
	Public FChgSourcearea
    Public FItemweight
    Public FExchangeRate
    Public FMultiplerate
    Public FMaySellPrice
    Public FWonprice
	Public FBrandCode
	Public Fsocname_kor
	Public FDepthCode
	Public FDepth4Code
	Public FGmarketGoodNo
	Public FGmarketprice
	Public FGmarketSellYn

	'// 품절여부
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold <= CMAXLIMITSELL))
	end function

	Public Function MustPrice()
		Dim GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
	End Function

	'// 지마켓 판매여부 반환
	Public Function getGmarketSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getGmarketSellYn = "Y"
			Else
				getGmarketSellYn = "N"
			End If
		Else
			getGmarketSellYn = "N"
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

	Public Function fnAttributes(vOption)
		Dim strSql
		strSql = ""
		strSql = strSql & " SELECT TOP 1 attributes "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_zilingo_attr_mapping] "
		strSql = strSql & " WHERE itemid = '"&FItemid&"' "
		strSql = strSql & " and itemoption = '"&vOption&"' "
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			fnAttributes = rsget("attributes")
		End If
		rsget.Close
	End Function

	Public Function getAttributeChoiceIds(vAttr)
		Dim strRst, arrAttributes, attributes, tmpAttributes, spTmpAttributes
		Dim i

		If vAttr <> "" Then
			tmpAttributes = Split(vAttr, ",")
			If isArray(tmpAttributes) Then
				For i=0 To Ubound(tmpAttributes)
					spTmpAttributes = Split(tmpAttributes(i), "||")(0)
					If (spTmpAttributes <> "COLORS") AND (spTmpAttributes <> "SIZES") AND (spTmpAttributes <> "CAPACITIES") Then
						arrAttributes = arrAttributes & """" & Split(tmpAttributes(i), "||")(1) & ""","
					End If
				Next
				If Right(arrAttributes,1) = "," Then
					arrAttributes = Left(arrAttributes, Len(arrAttributes) - 1)
				End If

				strRst = "	""attributeChoiceIds"": ["&arrAttributes&"],"
			End If
		End If
		getAttributeChoiceIds = strRst
	End Function
	
	Public Function getZilingoContParamToReg()
		Dim strRst
		strRst = ""
		If FChgItemContent <> "" Then
			strRst = strRst & jsEncode(FChgItemContent & vbcrlf & vbcrlf & vbcrlf)
		End If

		If FChgItemsource <> "" Then
			strRst = strRst & jsEncode("Material : " & FChgItemsource & vbcrlf)
		End If

		If FChgItemsize <> "" Then
			strRst = strRst & jsEncode("Size : " & FChgItemsize & vbcrlf)
		End If

		If FChgSourcearea <> "" Then
			strRst = strRst & jsEncode("Origin : " & FChgSourcearea)
		End If

		getZilingoContParamToReg = strRst
	End Function

	Public Function getEtcIds(vAttr)
		Dim arrAttributes, attributes, tmpAttributes, spTmpAttributes
		Dim i, colorStr, sizeStr, capaStr

		If vAttr <> "" Then
			tmpAttributes = Split(vAttr, ",")
			If isArray(tmpAttributes) Then
				For i=0 To Ubound(tmpAttributes)
					spTmpAttributes = Split(tmpAttributes(i), "||")(0)
					If (spTmpAttributes = "COLORS") Then
						colorStr = "		""colorId"": """&Split(tmpAttributes(i), "||")(1)&""","
					ElseIf (spTmpAttributes = "SIZES") Then
						sizeStr = "		""sizeId"": """&Split(tmpAttributes(i), "||")(1)&""","
					ElseIf (spTmpAttributes = "CAPACITIES") Then
						capaStr = "		""capacityId"": """&Split(tmpAttributes(i), "||")(1)&""","
					End If
				Next
			End If
		End If
		getEtcIds = colorStr & sizeStr & capaStr
	End Function

	Public Function getZilingoAddImageParam
		Dim strSql, imgLists, strRst
		strSql = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		
		imgLists = ""
		imgLists = imgLists & """"&FbasicImage&""","
		If Not(rsget.EOF or rsget.BOF) Then
			'For i=2 to rsget.RecordCount
			For i=0 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					imgLists = imgLists & """http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & ""","
				End If
				rsget.MoveNext
				If i>=7 Then Exit For
			Next
		End If
		rsget.Close
		If Right(imgLists,1) = "," Then
			imgLists = Left(imgLists, Len(imgLists) - 1)
		End If

		strRst = "		""imageUrls"": ["&imgLists&"],"
		getZilingoAddImageParam = strRst
	End Function

	Public Function getLimitEA(vOption)
		Dim strSql, itemEa
		Dim rsOptlimitno, rsOptlimitsold, rsIsusing, rsOptsellyn

		If vOption = "0000" Then
			If FLimitYn = "Y" Then
				itemEa = FLimitNo - FLimitSold - 5
			Else
				itemEa = CDEFALUT_STOCK
			End If
		Else
			strSql = ""
			strSql = strSql & " SELECT TOP 1 isusing, optsellyn, optlimitno, optlimitsold "
			strSql = strSql & " FROM db_item.dbo.tbl_item_option "
			strSql = strSql & " WHERE itemid = '"&Fitemid&"' "
			strSql = strSql & " and itemoption = '"&vOption&"' "
			rsget.Open strSql,dbget,1
			If not rsget.EOF Then
				rsIsusing		= rsget("isusing")
				rsOptsellyn		= rsget("optsellyn")
				rsOptlimitno	= rsget("optlimitno")
				rsOptlimitsold	= rsget("optlimitsold")
			End If
			rsget.Close

			If FLimitYn = "Y" Then
				itemEa = rsOptlimitno - rsOptlimitsold - 5
			Else
				itemEa = CDEFALUT_STOCK
			End If
		End If

		If itemEa < 1 Then
			getLimitEA = 0
		Else
			getLimitEA = itemEa
		End If
	End Function

	Public Function getZilingoItemRegJSON(iitemname, ioption, iquantity)
		Dim strRst, attrStr
		attrStr		= fnAttributes(ioption)
		iquantity	= getLimitEA(ioption)

		strRst = ""
		strRst = strRst & "{"
		strRst = strRst & "	""name"": """&iitemname&""","							'name of the product
		'strRst = strRst & "	""description"": """&FChgItemContent&""","			'description of the product
		strRst = strRst & "	""description"": """&getZilingoContParamToReg&""","		'description of the product
		strRst = strRst & "	""primaryImageUrl"": """&FbasicImage&""","				'url of primary image to be shown at the product displayn
		strRst = strRst & "	""subCategory"": {"
		strRst = strRst & "		""subCategoryIdType"": ""ZILINGO_SUBCAT"","
		strRst = strRst & "		""id"": """&FCateKey&""""
		strRst = strRst & "	},"
		strRst = strRst & getAttributeChoiceIds(attrStr)
		strRst = strRst & "	""skus"": [{"
		strRst = strRst & getEtcIds(attrStr)
		strRst = strRst & "		""quantity"": "&iquantity&","
		strRst = strRst & getZilingoAddImageParam()
		strRst = strRst & "		""price"" : {"
		strRst = strRst & "			""amount"": "&FWonprice&","
		strRst = strRst & "			""currency"": ""KRW"","
		strRst = strRst & "			""sellerSKUId"": """&FItemid&"_"&ioption&""""
		strRst = strRst & "		}"
		strRst = strRst & "	}]"
		strRst = strRst & "}"
		getZilingoItemRegJSON = strRst
	End Function

	Public Function getZilingoPriceJSON()
		Dim strRst
		strRst = ""
		strRst = strRst & "{"
		strRst = strRst & "	""productId"": """&FZilingoGoodNo&""","
		strRst = strRst & "	""pricePerUnit"": {"
		strRst = strRst & "		""amount"": "&FWonprice&","
		strRst = strRst & "		""currency"": ""KRW"""
		strRst = strRst & "	}"
		strRst = strRst & "}"
		getZilingoPriceJSON = strRst
	End Function

	Public Function getZilingoPriceBySkuNoJSON()
		Dim strRst
		strRst = ""
		strRst = strRst & "{"
		strRst = strRst & "	""productId"": """&FZilingoGoodNo&""","
		strRst = strRst & "	""skuPrices"": [{"
		strRst = strRst & "		""zilingoSKUId"": """&FZilingoSkuGoodNo&""","
		strRst = strRst & "			""pricePerUnit"": {"
		strRst = strRst & "				""amount"": "&FWonprice&","
		strRst = strRst & "				""currency"": ""KRW"""
		strRst = strRst & "			}"
		strRst = strRst & "	}]"
		strRst = strRst & "}"
		getZilingoPriceBySkuNoJSON = strRst
	End Function
End Class

Class CZilingo
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
	Public FRectItemOption

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
	Public Sub getZilingoNotRegOneItem
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
		strSql = strSql & "	,isNULL(R.zilingoStatCD,-9) as zilingoStatCD "
		strSql = strSql & " ,isnull(m.itemname, '') as chgItemname, m.itemContent as chgitemContent, isnull(m.itemsource, '') as chgitemsource, isnull(m.itemsize, '') as chgitemsize, isnull(m.sourcearea, '') as chgsourcearea "
		strSql = strSql & "	,isnull(c.CateKey, '') as CateKey, ex.exchangeRate, ex.multiplerate, uu.orgprice as maySellPrice, uu.wonprice "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		If FRectItemoption <> "0000" Then
			strSql = strSql & " JOIN db_item.[dbo].[tbl_item_option] as o on i.itemid = o.itemid and o.itemoption = '"&FRectItemoption&"' "
		End If
		strSql = strSql & " JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_multiLang_price uu on i.itemid = uu.itemid and uu.sitename = 'ZILINGO' "
		strSql = strSql & " JOIN db_item.dbo.tbl_exchangeRate as ex on uu.sitename = ex.sitename and ex.countryLangCD = m.countrycd "
		If FRectItemoption <> "0000" Then
			strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_zilingo_regItem] R on i.itemid = R.itemid and R.itemoption = '"&FRectItemoption&"' "
		Else
			strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_zilingo_regItem] R on i.itemid = R.itemid"
		End If
		strSql = strSql & " left JOIN db_etcmall.[dbo].[tbl_zilingo_cate_mapping] as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small  "
		strSql = strSql & " WHERE 1 = 1  "
		strSql = strSql & " and i.isusing = 'Y' "
		strSql = strSql & " and i.itemdiv <> '21' "
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.deliverOverseas = 'Y' "		'해외배송상품 Y
'		strSql = strSql & " and i.itemweight <> 0 "				'무게는 0보다 커야 /  2018-01-29 김진영..질링고의 경우 무게에 상관없이 무료 배송으로 진행한다함 주석처리
		strSql = strSql & " and i.itemid not in (select itemid from db_item.[dbo].[tbl_item_option] Where optaddprice > 0 group by itemid) "		'옵션 중 추가금액 제외
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and isnull(R.zilingoStatCD, 0) < 3 "
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & addSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CZilingoItem
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
				FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.FoptionCnt			= rsget("optionCnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
                FOneItem.FZilingoStatCD		= rsget("zilingoStatCD")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FCateKey			= rsget("CateKey")
                FOneItem.FbasicimageNm 		= rsget("basicimage")
                FOneItem.FChgItemname 		= rsget("chgItemname")
                FOneItem.FChgItemContent 	= rsget("chgitemContent")
                FOneItem.FChgItemsource	 	= rsget("chgitemsource")
                FOneItem.FChgItemsize	 	= rsget("chgitemsize")
                FOneItem.FChgSourcearea	 	= rsget("chgsourcearea")
                FOneItem.FItemweight 		= rsget("itemweight")
                FOneItem.FExchangeRate 		= rsget("exchangeRate")
                FOneItem.FMultiplerate 		= rsget("multiplerate")
                FOneItem.FMaySellPrice 		= rsget("maySellPrice")
                FOneItem.FWonprice	 		= rsget("wonprice")
		End If
		rsget.Close
	End Sub

	'// 가격 수정 리스트
	Public Sub getZilingoPriceOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	,R.zilingoGoodNo ,isNULL(R.zilingoStatCD,-9) as zilingoStatCD "
		strSql = strSql & "	, ex.exchangeRate, ex.multiplerate, uu.orgprice as maySellPrice, uu.wonprice, i.orgprice, R.zilingoSkuGoodNo "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		If FRectItemoption <> "0000" Then
			strSql = strSql & " JOIN db_item.[dbo].[tbl_item_option] as o on i.itemid = o.itemid and o.itemoption = '"&FRectItemoption&"' "
		End If
		strSql = strSql & " JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_multiLang_price uu on i.itemid = uu.itemid and uu.sitename = 'ZILINGO' "
		strSql = strSql & " JOIN db_item.dbo.tbl_exchangeRate as ex on uu.sitename = ex.sitename and ex.countryLangCD = m.countrycd "
		strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_zilingo_regItem] R on i.itemid = R.itemid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and R.zilingoGoodNo is Not Null "						'#등록 상품만
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CZilingoItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FZilingoGoodNo		= rsget("zilingoGoodNo")
                FOneItem.FExchangeRate 		= rsget("exchangeRate")
                FOneItem.FMultiplerate 		= rsget("multiplerate")
                FOneItem.FMaySellPrice 		= rsget("maySellPrice")
                FOneItem.FWonprice	 		= rsget("wonprice")
                FOneItem.FOrgprice	 		= rsget("orgprice")
                FOneItem.FZilingoSkuGoodNo	= rsget("zilingoSkuGoodNo")
		End If
		rsget.Close
	End Sub

	Function fnMaySoldout(iitemid, iitemoption)
		Dim strSql
		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as cnt "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN [db_item].[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = 'ZILINGO' "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
		strSql = strSql & " LEFT JOIN db_item.[dbo].[tbl_item_multiLang_option] as mo on mo.itemid = o.itemid and mo.itemoption = IsNULL(o.itemoption,'0000') and mo.countryCd = 'EN' "
		strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
		strSql = strSql & " and o.itemoption = '"&iitemoption&"' "
		strSql = strSql & " and 'N' = (CASE WHEN i.isusing = 'N' "
		strSql = strSql & " OR i.sellyn <> 'Y' "
		strSql = strSql & " OR o.optsellyn <> 'Y' "
		strSql = strSql & " OR o.isusing <> 'Y' "
		strSql = strSql & " OR (i.limityn='Y' and o.optlimitno-o.optlimitsold < 5) "
		strSql = strSql & " OR o.optaddprice > 0 "
		strSql = strSql & " OR isnull(mo.optionname, '') = '' "
		strSql = strSql & " OR isnull(mo.optionTypeName, '') = '' "
		strSql = strSql & " THEN 'Y' ELSE 'N' END) "
		rsget.Open strSql, dbget, 1
		If rsget("cnt") = 0 Then
			fnMaySoldout = "Y"
		Else
			fnMaySoldout = "N"
		End If
		rsget.Close
	End Function

	Function fnZilingoItemname(iitemid, iitemoption, ichgitemname)
		Dim strSql
		If iitemoption = "0000" Then
			fnZilingoItemname = ichgitemname
		Else
			strSql = ""
			strSql = strSql & " SELECT optionname FROM db_item.[dbo].[tbl_item_multiLang_option] WHERE itemid = '"&iitemid&"' and itemoption = '"&iitemoption&"' and countryCd = 'EN' "
			rsget.Open strSql, dbget, 1
			If not rsget.EOF Then
				fnZilingoItemname = ichgitemname & "_" & rsget("optionname")
			End If
			rsget.Close
		End If
	End Function

End Class

'질링고 임시 상품코드 얻기
Function getTmpZilingoGoodNo(iitemid, iitemoption)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 zilingoTmpGoodNo FROM db_etcmall.dbo.tbl_zilingo_regitem WHERE itemid = '"&iitemid&"' and itemoption = '"&iitemoption&"' "
	rsget.Open strSql, dbget, 1
		getTmpZilingoGoodNo = rsget("zilingoTmpGoodNo")
	rsget.Close
End Function

'질링고 재고 상품코드 얻기
Function getSKUZilingoGoodNo(iitemid, iitemoption)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 zilingoSkuGoodNo FROM db_etcmall.dbo.tbl_zilingo_regitem WHERE itemid = '"&iitemid&"' and itemoption = '"&iitemoption&"' "
	rsget.Open strSql, dbget, 1
	If not rsget.EOF Then
		getSKUZilingoGoodNo = rsget("zilingoSkuGoodNo")
	End If
	rsget.Close
End Function

'질링고 재고 상품코드 얻기
Function getSKUZilingoGoodNo2(iitemid, iitemoption, iquantity)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 zilingoSkuGoodNo, quantity FROM db_etcmall.dbo.tbl_zilingo_regitem WHERE itemid = '"&iitemid&"' and itemoption = '"&iitemoption&"' "
	rsget.Open strSql, dbget, 1
	If not rsget.EOF Then
		iquantity				= rsget("quantity")
		getSKUZilingoGoodNo2	= rsget("zilingoSkuGoodNo")
	End If
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

Function jsEncode(str)
	Dim charmap(127), haystack()
	charmap(8)  = "\b"
	charmap(9)  = "\t"
	charmap(10) = "\n"
	charmap(12) = "\f"
	charmap(13) = "\r"
	charmap(34) = "\"""
	charmap(47) = "\/"
	charmap(92) = "\\"

	Dim strlen : strlen = Len(str) - 1
	ReDim haystack(strlen)

	Dim i, charcode
	For i = 0 To strlen
		haystack(i) = Mid(str, i + 1, 1)

		charcode = AscW(haystack(i)) And 65535
		If charcode < 127 Then
			If Not IsEmpty(charmap(charcode)) Then
				haystack(i) = charmap(charcode)
			ElseIf charcode < 32 Then
				haystack(i) = "\u" & Right("000" & Hex(charcode), 4)
			End If
		Else
			haystack(i) = "\u" & Right("000" & Hex(charcode), 4)
		End If
	Next

	jsEncode = Join(haystack, "")
End Function
%>