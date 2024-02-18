<%
CONST CMAXMARGIN = 5
CONST CMALLNAME = "shopify"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 1									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST shopifySELLERID = "e4854ac50618f4c0d9b76a478be2d183"
CONST shopifyAPIKEY = "122a50a2918b1da442ada0a15e962330"
Const shopifyAPIURL = "https://10x10-co-kr.myshopify.com"
CONST shopifyLOCALE = "en"
CONST CDEFALUT_STOCK = 999
CONST CcurrencyUnit = "USD"
CONST CcountryCd = "EN"
CONST C_PREFIX_ITEMNAME ="test:"

Class CshopifyItem
	Public Fitemid
	Public FshopifyGoodNo
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
	Public FshopifySkuGoodNo
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
	Public FshopifyStatCD
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
	
	Public FDepthCode
	Public FDepth4Code
	
	Public Fsocname
    Public Fproduct_type
    
    
    Public Function getshopifyEditJson()
        Dim strRst, attrStr
        
        strRst = ""
		strRst = strRst & "{"
        strRst = strRst & " ""product"": {"
        strRst = strRst & " ""id"": """&FshopifyGoodNo&""","
        strRst = strRst & " ""title"": """&C_PREFIX_ITEMNAME & Fitemname&""","
        strRst = strRst & " ""body_html"": """&getShopifyContParamToReg&""","  
        strRst = strRst & " ""vendor"": """&FSocName&""","
        strRst = strRst & " ""product_type"": """&Fproduct_type&""","  '' 1depth cateName En
        strRst = strRst & " ""tags"": """&getShopifyTags&""", "
        strRst = strRst & getShopifyVariantsOptionToEdit  ''옵션.
        strRst = strRst & getShopifyImagesToEdit  ''이미지.
        strRst = strRst & " ""published"": true "  ''Create a new unpublished product
        strRst = strRst & " }"
        strRst = strRst & "}"

		getshopifyEditJson = strRst
		
    end function
    
	Public Function getshopifyItemRegJSON()
		Dim strRst, attrStr
		'iquantity	= getLimitEA(ioption)
        
        strRst = ""
		strRst = strRst & "{"
        strRst = strRst & " ""product"": {"
        strRst = strRst & " ""title"": """&C_PREFIX_ITEMNAME & Fitemname&""","
        strRst = strRst & " ""body_html"": """&getShopifyContParamToReg&""","  
        strRst = strRst & " ""vendor"": """&FSocName&""","
        strRst = strRst & " ""product_type"": """&Fproduct_type&""","  '' 1depth cateName En
        strRst = strRst & " ""tags"": """&getShopifyTags&""", "
        strRst = strRst & getShopifyVariantsOptionToReg  ''옵션.
        strRst = strRst & getShopifyImagesToReg  ''이미지.
        strRst = strRst & " ""published"": true "  ''Create a new unpublished product
        strRst = strRst & " }"
        strRst = strRst & "}"

		getshopifyItemRegJSON = strRst
	End Function

    public function getShopifyVariantsOptionToEdit()
    
    end function

    public function getShopifyVariantsOptionToReg()
	    '' 옵션이 있을경우
		Dim strSql, RetJson, isOptionExists
		Dim variantsBody, optionTypeName, optionsBody
		dim limitea, optsellyn
      
		strSql = ""
		strSql = strSql & " SELECT i.itemid,isNULL(o.itemoption,'0000') itemoption,isNULL(mo.optionTypeName,'Title') optionTypeName,isNULL(mo.optionname,'Default Title') optionname,isNULL(mo.countryCd,'"&CcountryCd&"') countryCd ,p.orgprice, p.orgprice as comparePrice "
		strSql = strSql & "  , [db_storage].[dbo].[uf_getTenBarCodeType]('10',i.itemid,isNULL(o.itemoption,'0000')) [SKUcode] "
		strSql = strSql & "  , isNULL(i.itemWeight,0) as itemWeight"
		strSql = strSql & "  , (CASE WHEN isNULL(o.itemoption,'0000')='0000' THEN "
		strSql = strSql & "      (CASE WHEN i.limityn='Y' THEN i.limitno-i.limitsold ELSE "&CDEFALUT_STOCK&" END)"
		strSql = strSql & "   ELSE        "
		strSql = strSql & "      (CASE WHEN o.optlimityn='Y' THEN o.optlimitno-o.optlimitsold ELSE "&CDEFALUT_STOCK&" END)"
		strSql = strSql & "   END)  as limitea "
        strSql = strSql & "  ,(CASE WHEN i.isusing = 'N'  "
        strSql = strSql & "  OR i.sellyn <> 'Y'  "
        strSql = strSql & "  OR isNULL(o.optsellyn,'Y') <> 'Y'  "
        strSql = strSql & "  OR isNULL(o.isusing,'Y') <> 'Y'  "
        strSql = strSql & "  OR isNULL(o.optaddprice,0) <> 0  "
        ''strSql = strSql & "  OR (i.limityn='Y' and o.optlimitno-o.optlimitsold < "&CMAXLIMITSELL&")  "
        ''strSql = strSql & "  OR isnull(mo.optionname, '') = ''  "
        ''strSql = strSql & "  OR isnull(mo.optionTypeName, '') = ''  "
        strSql = strSql & "  THEN 'N' ELSE 'Y' END)  as optsellyn "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN [db_item].[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = '"&CMALLNAME&"' "
		strSql = strSql & " left JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
		strSql = strSql & " Join [db_item].[dbo].[tbl_item_multiLang_price] p on i.itemid = p.itemid and p.sitename = 'shopify'  and currencyUnit='"&CcurrencyUnit&"'"
		strSql = strSql & " left JOIN db_item.[dbo].[tbl_item_multiLang_option] as mo on mo.itemid = o.itemid and mo.itemoption = IsNULL(o.itemoption,'0000') and mo.countryCd = '"&CcountryCd&"' "
		strSql = strSql & " WHERE i.itemid = '"&Fitemid&"' "
		
'rw strSql
'response.end

		rsget.CursorLocation = adUseClient
        rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		if Not rsget.Eof then
		    isOptionExists = true
		    variantsBody = ""
		    optionsBody  = ""
		    Do until rsget.EOF
		        limitea = rsget("limitea")
		        optsellyn  = optsellyn
		        if (optsellyn="N") then limitea = 0
		            
		        if (variantsBody<>"") then variantsBody = variantsBody & ","
				variantsBody = variantsBody &"{"
                variantsBody = variantsBody & " ""option1"": """&rsget("optionname")&""", "
                variantsBody = variantsBody & " ""price"": """&rsget("orgprice")&""", "
                variantsBody = variantsBody & " ""compare_at_price"": """&rsget("orgprice")&""", "
                variantsBody = variantsBody & " ""grams"": """&rsget("itemWeight")&""", "
                variantsBody = variantsBody & " ""inventory_quantity"": """&limitea&""", "
                variantsBody = variantsBody & " ""sku"": """&rsget("SKUcode")&""", "
                variantsBody = variantsBody & " ""inventory_management"": ""shopify"", "
                variantsBody = variantsBody & " ""inventory_policy"": ""deny"", "
                variantsBody = variantsBody & " ""barcode"": """&rsget("SKUcode")&""" "
                variantsBody = variantsBody & " }"
                
                '' 단일 옵션만 쓰자.
                if (optionTypeName="") then
                    optionTypeName = rsget("optionTypeName")
                end if
                
                if optionsBody<>"" then optionsBody = optionsBody & ","
                optionsBody = optionsBody & " """&rsget("optionname")&""" "
                
				rsget.MoveNext
			Loop
		end if
		rsget.Close
	
	    RetJson = " ""variants"": ["
        RetJson = RetJson & variantsBody
        RetJson = RetJson & "],"
        RetJson = RetJson &" ""options"": ["
        RetJson = RetJson &" {"
        RetJson = RetJson &" ""name"": """&optionTypeName&""", "
        RetJson = RetJson &" ""values"": ["
        RetJson = RetJson & optionsBody
        RetJson = RetJson &" ]"
        RetJson = RetJson &" }"
        RetJson = RetJson &" ]," 
	    
	    '' 없을경우 디폴트로 생성됨. => 옵션없으면 Title / Default Title
	    if (NOT isOptionExists) then 
	        RetJson = ","
	    end if
	    
	    getShopifyVariantsOptionToReg = RetJson
    end Function

    public function getShopifyImagesToEdit()
    
    end function

    public function getShopifyImagesToReg()
	    Dim strSql, imgLists, strRst, addimgName, addimgURL, ii
		strSql = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		
		imgLists = ""
		ii=0
		If Not(rsget.EOF or rsget.BOF) Then
			Do until rsget.EOF
			    if(ii<7) then
    				If rsget("imgType") = "0" Then
    				    addimgName = rsget("addimage_600")
    				    if isNULL(addimgName) or (addimgName="") then
    				        addimgName = rsget("addimage_400")
    				        if isNULL(addimgName) or (addimgName="") then
    				            addimgURL = ""
    				        else
    				            addimgURL = "http://owebimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & addimgName
    				        end if
    				    else
    				        addimgURL = "http://owebimage.10x10.co.kr/image/add" & rsget("gubun") & "_600/" & GetImageSubFolderByItemid(Fitemid) & "/" & addimgName
    				    end if
    				    
    				    if (addimgURL<>"") then
        				    imgLists = imgLists & "  ,{""src"": """&addimgURL&"""}"
        				end if
    				End If
    			end if
				rsget.MoveNext
				ii=ii+1
				
			loop
		End If
		rsget.Close
		If Right(imgLists,1) = "," Then
			imgLists = Left(imgLists, Len(imgLists) - 1)
		End If

        strRst = " ""images"": ["
        strRst = strRst &  "{"
        strRst = strRst &  "  ""src"": """&FbasicImage&""""
        strRst = strRst &  "}"
        strRst = strRst &  imgLists
        strRst = strRst & "],"
        
		getShopifyImagesToReg = strRst
    end function
    
    Public Function getshopifyContParamToReg()
		Dim strRst
		strRst = ""
		If FChgItemContent <> "" Then
			strRst = strRst & jsEncode(FChgItemContent & "<br><br>")
		End If

		If FChgItemsource <> "" Then
			strRst = strRst & jsEncode("Material : " & FChgItemsource & "<br>")
		End If

		If FChgItemsize <> "" Then
			strRst = strRst & jsEncode("Size : " & FChgItemsize & "<br>")
		End If

		If FChgSourcearea <> "" Then
			strRst = strRst & jsEncode("Origin : " & FChgSourcearea)
		End If

		getshopifyContParamToReg = strRst
	End Function

    ''태그 	
	Public function getShopifyTags()
	    getShopifyTags = ""
    end function

    '// 품절여부
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold <= CMAXLIMITSELL))
	end function
	
'''-----------------------------------------------    

	Public Function MustPrice()
		Dim GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
	End Function





''	Public Function getLimitEA(vOption)
''		Dim strSql, itemEa
''		Dim rsOptlimitno, rsOptlimitsold, rsIsusing, rsOptsellyn
''
''		If vOption = "0000" Then
''			If FLimitYn = "Y" Then
''				itemEa = FLimitNo - FLimitSold - 5
''			Else
''				itemEa = CDEFALUT_STOCK
''			End If
''		Else
''			strSql = ""
''			strSql = strSql & " SELECT TOP 1 isusing, optsellyn, optlimitno, optlimitsold "
''			strSql = strSql & " FROM db_item.dbo.tbl_item_option "
''			strSql = strSql & " WHERE itemid = '"&Fitemid&"' "
''			strSql = strSql & " and itemoption = '"&vOption&"' "
''			rsget.Open strSql,dbget,1
''			If not rsget.EOF Then
''				rsIsusing		= rsget("isusing")
''				rsOptsellyn		= rsget("optsellyn")
''				rsOptlimitno	= rsget("optlimitno")
''				rsOptlimitsold	= rsget("optlimitsold")
''			End If
''			rsget.Close
''
''			If FLimitYn = "Y" Then
''				itemEa = rsOptlimitno - rsOptlimitsold - 5
''			Else
''				itemEa = CDEFALUT_STOCK
''			End If
''		End If
''
''		If itemEa < 1 Then
''			getLimitEA = 0
''		Else
''			getLimitEA = itemEa
''		End If
''	End Function

    


'	Public Function getshopifyPriceJSON()
'		Dim strRst
'		strRst = ""
'		strRst = strRst & "{"
'		strRst = strRst & "	""productId"": """&FshopifyGoodNo&""","
'		strRst = strRst & "	""pricePerUnit"": {"
'		strRst = strRst & "		""amount"": "&FWonprice&","
'		strRst = strRst & "		""currency"": ""KRW"""
'		strRst = strRst & "	}"
'		strRst = strRst & "}"
'		getshopifyPriceJSON = strRst
'	End Function
'
'	Public Function getshopifyPriceBySkuNoJSON()
'		Dim strRst
'		strRst = ""
'		strRst = strRst & "{"
'		strRst = strRst & "	""productId"": """&FshopifyGoodNo&""","
'		strRst = strRst & "	""skuPrices"": [{"
'		strRst = strRst & "		""shopifySKUId"": """&FshopifySkuGoodNo&""","
'		strRst = strRst & "			""pricePerUnit"": {"
'		strRst = strRst & "				""amount"": "&FWonprice&","
'		strRst = strRst & "				""currency"": ""KRW"""
'		strRst = strRst & "			}"
'		strRst = strRst & "	}]"
'		strRst = strRst & "}"
'		getshopifyPriceBySkuNoJSON = strRst
'	End Function
End Class

Class Cshopify
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
	Public Sub getshopifyNotRegOneItem
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
		strSql = strSql & "	,isNULL(R.shopifyStatCD,-9) as shopifyStatCD "
		strSql = strSql & " ,isnull(m.itemname, '') as chgItemname, m.itemContent as chgitemContent, isnull(m.itemsource, '') as chgitemsource, isnull(m.itemsize, '') as chgitemsize, isnull(m.sourcearea, '') as chgsourcearea "
		''strSql = strSql & "	,isnull(c.CateKey, '') as CateKey"
		strSql = strSql & "	, ex.exchangeRate, ex.multiplerate, uu.orgprice as maySellPrice, uu.wonprice "
		strSql = strSql & "	, c.socname, isNULL(ca.catename_e,'') as catename_e"
		
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_user.dbo.tbl_user_c c on i.makerid=c.userid" 
		'If FRectItemoption <> "0000" Then
		'	strSql = strSql & " JOIN db_item.[dbo].[tbl_item_option] as o on i.itemid = o.itemid and o.itemoption = '"&FRectItemoption&"' "
		'End If
		strSql = strSql & " JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_multiLang_price uu on i.itemid = uu.itemid and uu.sitename = 'shopify' "
		strSql = strSql & " JOIN db_item.dbo.tbl_exchangeRate as ex on uu.sitename = ex.sitename and ex.countryLangCD = m.countrycd "
			strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_shopify_regItem] R on i.itemid = R.itemid"
		strSql = strSql & " left join db_item.[dbo].[tbl_display_cate] ca"
		strSql = strSql & " on i.dispcate1=ca.catecode"
		strSql = strSql & " WHERE 1 = 1  "
		strSql = strSql & " and i.isusing = 'Y' "
		strSql = strSql & " and i.itemdiv <> '21' "
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.deliverOverseas = 'Y' "		'해외배송상품 Y
		strSql = strSql & " and i.itemweight > 0 "				'무게는 0보다 커야 /  2018-01-29 김진영..질링고의 경우 무게에 상관없이 무료 배송으로 진행한다함 주석처리
		strSql = strSql & " and i.itemid not in (select itemid from db_item.[dbo].[tbl_item_option] Where optaddprice > 0 group by itemid) "		'옵션 중 추가금액 제외
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and isnull(R.shopifyStatCD, 0) < 3 "
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & addSql
		
		rsget.CursorLocation = adUseClient
        rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CshopifyItem
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
                FOneItem.FshopifyStatCD		= rsget("shopifyStatCD")
                FOneItem.FDeliveryType		= rsget("deliveryType")
              '  FOneItem.FCateKey			= rsget("CateKey")
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
                
                FOneItem.FSocName           = rsget("socname")
                FOneItem.Fproduct_type      = rsget("catename_e")
		End If
		rsget.Close
	End Sub

    '// 등록상품 수정용
    Public Sub getshopifyEditOneItem
		Dim strSql, addSql, i

		strSql = ""
		strSql = strSql & " SELECT TOP 1  i.* "
		strSql = strSql & "	,R.shopifyGoodNo"
		strSql = strSql & "	,isNULL(R.shopifyStatCD,-9) as shopifyStatCD "
		strSql = strSql & " ,isnull(m.itemname, '') as chgItemname, m.itemContent as chgitemContent, isnull(m.itemsource, '') as chgitemsource, isnull(m.itemsize, '') as chgitemsize, isnull(m.sourcearea, '') as chgsourcearea "
		''strSql = strSql & "	,isnull(c.CateKey, '') as CateKey"
		strSql = strSql & "	, ex.exchangeRate, ex.multiplerate, uu.orgprice as maySellPrice, uu.wonprice "
		strSql = strSql & "	, c.socname, isNULL(ca.catename_e,'') as catename_e"
		
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_user.dbo.tbl_user_c c on i.makerid=c.userid" 
		strSql = strSql & " JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_multiLang_price uu on i.itemid = uu.itemid and uu.sitename = 'shopify' "
		strSql = strSql & " JOIN db_item.dbo.tbl_exchangeRate as ex on uu.sitename = ex.sitename and ex.countryLangCD = m.countrycd "
		strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_shopify_regItem] R on i.itemid = R.itemid"
		strSql = strSql & " left join db_item.[dbo].[tbl_display_cate] ca"
		strSql = strSql & " on i.dispcate1=ca.catecode "
		strSql = strSql & " WHERE 1 = 1  "
		strSql = strSql & " and i.isusing = 'Y' "
		strSql = strSql & " and i.itemdiv <> '21' "
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.deliverOverseas = 'Y' "		'해외배송상품 Y
		strSql = strSql & " and i.itemweight > 0 "				'무게는 0보다 커야 /  2018-01-29 김진영..질링고의 경우 무게에 상관없이 무료 배송으로 진행한다함 주석처리
		strSql = strSql & " and i.itemid not in (select itemid from db_item.[dbo].[tbl_item_option] Where optaddprice > 0 group by itemid) "		'옵션 중 추가금액 제외
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.sellcash > 0 "
		'strSql = strSql & " and i.sellyn = 'Y' "
		'strSql = strSql & " and isnull(R.shopifyStatCD, 0) < 3 "
		'strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and i.itemid="&FRectItemID
		
		rsget.CursorLocation = adUseClient
        rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CshopifyItem
				FOneItem.FItemid			= rsget("itemid")
				FOneItem.FshopifyGoodNo		= rsget("shopifyGoodNo")
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
                FOneItem.FshopifyStatCD		= rsget("shopifyStatCD")
                FOneItem.FDeliveryType		= rsget("deliveryType")
              '  FOneItem.FCateKey			= rsget("CateKey")
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
                
                FOneItem.FSocName           = rsget("socname")
                FOneItem.Fproduct_type      = rsget("catename_e")
                
		End If
		rsget.Close
	End Sub
    
    
    
    
	'// 가격 수정 리스트
	Public Sub getshopifyPriceOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	,R.shopifyGoodNo ,isNULL(R.shopifyStatCD,-9) as shopifyStatCD "
		strSql = strSql & "	, ex.exchangeRate, ex.multiplerate, uu.orgprice as maySellPrice, uu.wonprice, i.orgprice, R.shopifySkuGoodNo "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		If FRectItemoption <> "0000" Then
			strSql = strSql & " JOIN db_item.[dbo].[tbl_item_option] as o on i.itemid = o.itemid and o.itemoption = '"&FRectItemoption&"' "
		End If
		strSql = strSql & " JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_multiLang_price uu on i.itemid = uu.itemid and uu.sitename = 'shopify' "
		strSql = strSql & " JOIN db_item.dbo.tbl_exchangeRate as ex on uu.sitename = ex.sitename and ex.countryLangCD = m.countrycd "
		strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_shopify_regItem] R on i.itemid = R.itemid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and R.shopifyGoodNo is Not Null "						'#등록 상품만
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CshopifyItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FshopifyGoodNo		= rsget("shopifyGoodNo")
                FOneItem.FExchangeRate 		= rsget("exchangeRate")
                FOneItem.FMultiplerate 		= rsget("multiplerate")
                FOneItem.FMaySellPrice 		= rsget("maySellPrice")
                FOneItem.FWonprice	 		= rsget("wonprice")
                FOneItem.FOrgprice	 		= rsget("orgprice")
                FOneItem.FshopifySkuGoodNo	= rsget("shopifySkuGoodNo")
		End If
		rsget.Close
	End Sub

''	Function fnMaySoldout(iitemid, iitemoption)
''		Dim strSql
''		strSql = ""
''		strSql = strSql & " SELECT COUNT(*) as cnt "
''		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
''		strSql = strSql & " JOIN [db_item].[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = 'shopify' "
''		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
''		strSql = strSql & " LEFT JOIN db_item.[dbo].[tbl_item_multiLang_option] as mo on mo.itemid = o.itemid and mo.itemoption = IsNULL(o.itemoption,'0000') and mo.countryCd = 'EN' "
''		strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
''		strSql = strSql & " and o.itemoption = '"&iitemoption&"' "
''		strSql = strSql & " and 'N' = (CASE WHEN i.isusing = 'N' "
''		strSql = strSql & " OR i.sellyn <> 'Y' "
''		strSql = strSql & " OR o.optsellyn <> 'Y' "
''		strSql = strSql & " OR o.isusing <> 'Y' "
''		strSql = strSql & " OR (i.limityn='Y' and o.optlimitno-o.optlimitsold < 5) "
''		strSql = strSql & " OR o.optaddprice > 0 "
''		strSql = strSql & " OR isnull(mo.optionname, '') = '' "
''		strSql = strSql & " OR isnull(mo.optionTypeName, '') = '' "
''		strSql = strSql & " THEN 'Y' ELSE 'N' END) "
''		
''		rsget.CursorLocation = adUseClient
''        rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
''		If rsget("cnt") = 0 Then
''			fnMaySoldout = "Y"
''		Else
''			fnMaySoldout = "N"
''		End If
''		rsget.Close
''	End Function

''	Function fnshopifyItemname(iitemid, iitemoption, ichgitemname)
''		Dim strSql
''		If iitemoption = "0000" Then
''			fnshopifyItemname = ichgitemname
''		Else
''			strSql = ""
''			strSql = strSql & " SELECT optionname FROM db_item.[dbo].[tbl_item_multiLang_option] WHERE itemid = '"&iitemid&"' and itemoption = '"&iitemoption&"' and countryCd = 'EN' "
''			rsget.Open strSql, dbget, 1
''			If not rsget.EOF Then
''				fnshopifyItemname = ichgitemname & "_" & rsget("optionname")
''			End If
''			rsget.Close
''		End If
''	End Function

End Class

'shopify 상품코드 얻기
Function getShopifyGoodNo(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 shopifyGoodNo FROM db_etcmall.dbo.tbl_shopify_regitem WHERE itemid = '"&iitemid&"'" ''' and itemoption = '"&iitemoption&"' "
	rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		getShopifyGoodNo = rsget("shopifyGoodNo")
	rsget.Close
End Function

'shopify 재고 상품코드 얻기
'Function getSKUshopifyGoodNo(iitemid, iitemoption)
'	Dim strSql
'	strSql = ""
'	strSql = strSql & " SELECT TOP 1 shopifySkuGoodNo FROM db_etcmall.dbo.tbl_shopify_regitem WHERE itemid = '"&iitemid&"' and itemoption = '"&iitemoption&"' "
'	rsget.Open strSql, dbget, 1
'	If not rsget.EOF Then
'		getSKUshopifyGoodNo = rsget("shopifySkuGoodNo")
'	End If
'	rsget.Close
'End Function


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