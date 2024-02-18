<%
CONST CMAXMARGIN = 14.9
CONST CMALLNAME = "homeplus"
CONST CUPJODLVVALID = TRUE		''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5			'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.

Class CHomeplusItem
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
	Public FHomeplusGoodNo
	Public FHomeplusprice		
	Public FHomeplusSellYn	
	Public FregedOptCnt       
	Public FaccFailCNT        
	Public FlastErrStr        
	Public Fdeliverytype      
	Public FrequireMakeDay    
	Public FinfoDiv       	
	Public Fsafetyyn      	
	Public FsafetyDiv     	
	Public FsafetyNum     	
	Public FmaySoldOut    	

	Public FHomeplusStatCD
	Public FhDIVISION
	Public FhGROUP
	Public FhDEPT
	Public FhCLASS
	Public FhSUBCLASS
	Public FdepthCode
	Public FbrandDepthCode

	Public MustPrice
	Public FItemOption
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold
	Public Fregitemname

	Public Function getOptionLimitNo()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (IsOptionSoldOut) Then
			getOptionLimitNo = 0
		Else
			If (Foptlimityn = "Y") Then
				If (Foptlimitno - Foptlimitsold < CLIMIT_SOLDOUT_NO) Then
					getOptionLimitNo = 0
				Else
					getOptionLimitNo = Foptlimitno - Foptlimitsold - CLIMIT_SOLDOUT_NO
				End If
			Else
				getOptionLimitNo = 999
			End if
		End If
	End Function

	Public Function IsOptionSoldOut()
		CONST CLIMIT_SOLDOUT_NO = 5
		IsOptionSoldOut = false
		If (FItemOption = "0000") Then Exit Function
		IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno - Foptlimitsold < CLIMIT_SOLDOUT_NO))
	End Function

	'// Homeplus 판매여부 반환
	Public Function getHomeplusSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getHomeplusSellYn = "Y"
			Else
				getHomeplusSellYn = "N"
			End If
		Else
			getHomeplusSellYn = "N"
		End If
	End Function

	public function GetHomeplusLmtQty()
		CONST CLIMIT_SOLDOUT_NO = 5
		If (Flimityn="Y") then
			If (Flimitno - Flimitsold) < CLIMIT_SOLDOUT_NO Then
				GetHomeplusLmtQty = 0
			Else
				GetHomeplusLmtQty = Flimitno - Flimitsold - CLIMIT_SOLDOUT_NO
			End If
		Else
			GetHomeplusLmtQty = 999
		End If
	End Function

    Function getHomeplusSuplyPrice(optaddprice)
		getHomeplusSuplyPrice= cLng((MustPrice+optaddprice)*0.89)
    End Function

	'// 품절여부
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	'// 품절여부
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	'// 검색어
	Public Function getItemKeyword()
		Dim p, strRst, arrData, arrTmp
		If trim(Fkeywords) = "" Then Exit Function
		strRst = ""
		Fkeywords = replace(Fkeywords, ",,", ",")

		If instr(Fkeywords, ",") > 1 Then
			arrData = Split(Fkeywords, ",")
			arrTmp = FnDistinctData(arrData)
			strRst = "<TAGS>"
			For p=0 to Ubound(arrTmp)-1
				strRst = strRst & "<item><![CDATA["&arrTmp(p)&"]]></item>"
			Next
			strRst = strRst & "</TAGS>"
		End If
		getItemKeyword = strRst
	End Function

	'배열내의 중복값 제거
	Function FnDistinctData(ByVal aData)
		Dim dicObj, items, returnValue
		Set dicObj = CreateObject("Scripting.dictionary")
			dicObj.removeall
			dicObj.CompareMode = 0
			'loop를 돌면서 기존 배열에 있는지 검사 후 Add
			For Each items In aData
				If not dicObj.Exists(items) Then dicObj.Add items, items
			Next

			returnValue = dicObj.keys
		Set dicObj = Nothing
		FnDistinctData = returnValue
	End Function

	'// 상품등록: 옵션 파라메터 생성(상품등록용)
	Public Function getHomeplusOptionParamToReg
		Dim strSql, strRst, itemSu, itemoption, optionname, optaddprice
		Dim GetTenTenMargin, i
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		strRst = ""
		optaddprice		= 0
		strSql = ""
		strSql = strSql & " SELECT top 900 i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(40),o.optionname) as optionname" & VBCRLF
		strSql = strSql & " , o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, i.deliverfixday, o.optaddprice " & VBCRLF
		strSql = strSql & " ,DATALENGTH(o.optionname) as optnmLen" & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_item as i " & VBCRLF
		strSql = strSql & " LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF
		strSql = strSql & " WHERE i.itemid = "&Fitemid
		strSql = strSql & " ORDER BY o.optaddprice ASC, o.itemoption ASC "
'rw strSql
'response.end
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				If rsget.RecordCount = 1 AND IsNull(rsget("itemoption")) Then  ''단일상품
					FItemOption = "0000"
					optionname = DdotFormat(chrbyte(getItemNameFormat,40,""),20)
					itemSu = GetHomeplusLmtQty
				Else
					FItemOption 	= rsget("itemoption")
					optionname 		= rsget("optionname")
					Foptsellyn 		= rsget("optsellyn")
					Foptlimityn 	= rsget("optlimityn")
					Foptlimitno 	= rsget("optlimitno")
					Foptlimitsold 	= rsget("optlimitsold")
					optaddprice		= rsget("optaddprice")
					itemSu = getOptionLimitNo

					If rsget("optnmLen")>100 then
					    optionname=DdotFormat(optionname,50)
					End If
				End If
				strRst = strRst &"<ITEM>"
				strRst = strRst &"	<s_ITEMNO>"&FItemOption&"</s_ITEMNO>"							'##*업체 아이템번호 / 업체의 해당 아이템(옵션) 번호 나중에 ProductResult값에서 비교하기 위해 입력하여 준다.
				strRst = strRst &"	<i_SIZE>1</i_SIZE>"												'##*Size(Amos) / 1부터 시작 1,2,3,4…….)해당 사이즈 정보는 업체에서 저장해놓으시기 바랍니다. 다른 API에서 사용됩니다. I_ITEMNO+I_SIZE가 키 값으로 사용 되어 집니다.
				strRst = strRst &"	<s_OPTION_NAME><![CDATA["&optionname&"]]></s_OPTION_NAME>"		'##*옵션명
				strRst = strRst &"	<i_STOCK_TYPE>1</i_STOCK_TYPE>"									'재고관리 / 1: WEB 관리 3: 관리 안 함(Default)재고관리를 할 경우 1번 선택
				strRst = strRst &"	<i_LIBQTY>"&itemSu&"</i_LIBQTY>"								'재고수량 / 재고관리에 3번을 선택한 경우 값은 무시된다
				strRst = strRst &"	<f_RETAILPRICE>"&MustPrice+optaddprice&"</f_RETAILPRICE>"		'*판매가
				strRst = strRst &"	<f_BUYPRICE>"&getHomeplusSuplyPrice(optaddprice)&"</f_BUYPRICE>"'*공급가(VAT포함)
'				strRst = strRst &"	<i_ACCUMULATION_RATE></i_ACCUMULATION_RATE>"						'상품별적립율 / 상품별 FMC적립율
'				strRst = strRst &"	<d_RELEASE_DATE></d_RELEASE_DATE>"									'출시일자 / 출시일자 (YYYYMMDD)
				strRst = strRst &"</ITEM>"
				rsget.MoveNext
			Next
		End If
		rsget.Close
		getHomeplusOptionParamToReg = strRst
	End Function

	'// 상품수정: 옵션 파라메터 생성(상품수정용)
	Public Function getHomeplusOptionParamToEDT
		Dim strSql, sRst, itemSu, itemoption, optionname, optaddprice
		Dim GetTenTenMargin, i, arrRows, sellstat
		Dim isOptionExists, notitemId, notmakerid
		Dim optiontypename, optLimit, optlimityn, isUsing, optsellyn, preged, optNameDiff, forceExpired, oopt, ooptCd, DelOpt

		strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_homeplus 'homeplus'," & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			arrRows = rsget.getRows
		End If
		rsget.close
		isOptionExists = isArray(arrRows)

		strSql = "SELECT COUNT(*) as cnt FROM db_temp.dbo.tbl_jaehyumall_not_in_itemid where mallgubun = 'homeplus' and itemid =" & Fitemid
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			notitemId = rsget("cnt")
		End If
		rsget.close

		strSql = "SELECT COUNT(*) as cnt FROM db_item.dbo.tbl_item as i join [db_temp].dbo.tbl_jaehyumall_not_in_makerid as m on i.makerid = m.makerid where i.itemid = "& Fitemid&" and m.mallgubun = 'homeplus'"
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			notmakerid = rsget("cnt")
		End If
		rsget.close

		If (isOptionExists) Then
			For i = 0 To UBound(ArrRows,2)
				itemoption			= ArrRows(1,i)
				optiontypename		= ArrRows(2,i)
'				 optionname			= Replace(Replace(db2Html(ArrRows(3,i)),":",""),",","")
				optionname			= Replace(db2Html(ArrRows(3,i)),":","")					'2015-05-15 11:16 수정 / 복합옵션의 콤마를 replace해서 수정됐었음..위 optionname 주석함
				optLimit			= ArrRows(4,i)
				optlimityn			= ArrRows(5,i)
				isUsing				= ArrRows(6,i)
				optsellyn			= ArrRows(7,i)
				preged				= ArrRows(11,i)
				optNameDiff			= ArrRows(12,i)
				forceExpired		= ArrRows(13,i)
				oopt				= ArrRows(14,i)
				ooptCd				= ArrRows(15,i)
				DelOpt				= ArrRows(16,i)
				optaddprice			= ArrRows(17,i)

				If IsSoldOut Then
					sellstat = 2
				Else
					If itemoption = "0000" AND UBound(ArrRows,2) = 0 Then
						optionname = oopt
						itemSu = GetHomeplusLmtQty
					Else
						If (optlimityn = "Y") Then
							If optLimit <= 5 Then
								itemSu = 0
							Else
								itemSu = optLimit - 5
							End If
						Else
							itemSu = 999
						End if
	
						If (DelOpt = 1) OR (isUsing = "N") OR (optsellyn = "N") OR (notitemId > 0) OR (notmakerid > 0) Then
							sellstat = 2
						Else
							sellstat = 1
						End If
					End If
					optionname = DdotFormat(optionname,50)
	
					GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
					If GetTenTenMargin < CMAXMARGIN Then
						MustPrice = Forgprice
					Else
						MustPrice = FSellCash
					End If
				End If

'rw itemoption
'rw ooptCd
'rw optionname
'rw itemSu
'rw MustPrice+optaddprice
'rw getHomeplusSuplyPrice(optaddprice)
'rw sellstat
'rw "------------"
				sRst = sRst &"<ITEM>"
				sRst = sRst &"	<s_ITEMNO>"&itemoption&"</s_ITEMNO>"							'*업체 아이템번호 / 업체의 해당 아이템(옵션) 번호 나중에 ProductResult값에서 비교하기 위해 입력하여 준다.
				If preged = 1 Then
					sRst = sRst &"	<i_ITEMNO>"&ooptCd&"</i_ITEMNO>"							'아이템번호 / 수정되는 아이템이면 해당 값을 반드시 입력하여 주시기 바랍니다 신규 추가되는 아이템의 경우에는 입력하지 마세요
				End If
				sRst = sRst &"	<i_SIZE>1</i_SIZE>"												'*Size(Amos) / 하단의 예제 참조(AK 몰은 아이템 리스트의 순번1부터 시작 1,2,3,4…….)해당 사이즈 정보는 업체에서 저장해놓으시기 바랍니다. 다른 API에서 사용됩니다.
				sRst = sRst &"	<s_OPTION_NAME><![CDATA["&optionname&"]]></s_OPTION_NAME>"		'*옵션명
				sRst = sRst &"	<i_STOCK_TYPE>1</i_STOCK_TYPE>"									'재고관리 / 1: WEB 관리 3: 관리 안 함(Default)재고관리를 할 경우 1번 선택
				sRst = sRst &"	<i_LIBQTY>"&itemSu&"</i_LIBQTY>"								'재고수량 / 재고관리에 3번을 선택한 경우 값은 무시된다
				sRst = sRst &"	<f_RETAILPRICE>"&MustPrice+optaddprice&"</f_RETAILPRICE>"		'*판매가 / 공급가 정보가 이전에 입력한 적이 있는 공급가인 경우 판매가는 이전 공급가의 판매가에 맞춰집니다..API 연동 상품의 경우 제휴 마진율이 정하여져 있으므로 마진율을 임의로 변경하지 마시기 바랍니다.
				sRst = sRst &"	<f_BUYPRICE>"&getHomeplusSuplyPrice(optaddprice)&"</f_BUYPRICE>"'*공급가(VAT포함)
				If preged = 1 Then
					sRst = sRst &"	<i_STATUS>"&sellstat&"</i_STATUS>"							'판매 중/판매중지 | 1: 판매중 2:판매중지, 신규 추가되는 아이템은 자동으로 판매중으로 처리됩니다. 수정되는 아이템의 경우에만 이 필드를 사용합니다.
				End If
'				sRst = sRst &"	<ACCUMULATION_RATE></ACCUMULATION_RATE>"						'상품별적립율 / 상품별 FMC적립율
'				sRst = sRst &"	<RELEASE_DATE></RELEASE_DATE>"									'출시일자 / 출시일자 (YYYYMMDD)
				sRst = sRst &"</ITEM>"
			Next
		End If
'response.end
		getHomeplusOptionParamToEDT = sRst
	End Function

	'// 상품등록: 상품추가이미지 파라메터 생성(상품등록용)
	Public Function getHomeplusAddImageParamToReg()
		Dim strRst, strSQL, i, strRst2
		strRst = ""
		strRst2 = ""
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If

		strRst = strRst &"<s_IMG_BIG>"&FbasicImage&"</s_IMG_BIG>"		'*기본이미지 URL | HTTP URL 형식. 해당 이미지는 외부에서 다운로드 가능한 URL이어야 한다(IP 로 기술 권장, 도메인은 문의)
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		'보조이미지 URL | HTTP URL 형식. 여러 개를 등록할 수 있다. 해당 이미지는 외부에서 다운로드 가능한 URL 이어야 한다(IP로 기술 권장, 도메인은 문의)
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					strRst2 = strRst2 &"	<item>http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") &"</item>"
				End If
				rsget.MoveNext
				If i >= 5 Then Exit For
			Next

			If strRst2 <> "" Then
				strRst2 = "<s_IMG_SKCS1>"&strRst2&"</s_IMG_SKCS1>"
			End If
		End If
		rsget.Close
		getHomeplusAddImageParamToReg = strRst&strRst2
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

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	Public Function getHomeplusItemContParamToReg()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
		strRst = strRst & ("<p><center><a href=""http://direct.homeplus.co.kr/app.exhibition.category.Category.ghs?comm=usr.category.inf&ctg_id=133459"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_homeplus.jpg""></a></center></p><br>")
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
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_common.jpg"">")

		strRst = strRst & ("</div>")
		getHomeplusItemContParamToReg = strRst
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

	'// 상품수정: 상품추가이미지 파라메터 생성(상품수정용)
	Public Function getHomeplusAddImageParamToEDT()
		Dim strRst, strSQL, i
		strRst = ""
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If

		strRst = strRst &"<BASIC>"&FbasicImage&"</BASIC>"		'*기본이미지 URL | HTTP URL 형식. 해당 이미지는 외부에서 다운로드 가능한 URL이어야 한다(IP 로 기술 권장, 도메인은 문의)
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		'보조이미지 URL | HTTP URL 형식. 여러 개를 등록할 수 있다. 해당 이미지는 외부에서 다운로드 가능한 URL 이어야 한다(IP로 기술 권장, 도메인은 문의)
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					strRst = strRst &"		<EXTRA>http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") &"</EXTRA>"
				End If
				rsget.MoveNext
				If i >= 5 Then Exit For
			Next
		End If
		rsget.Close
		getHomeplusAddImageParamToEDT = strRst
	End Function

	'// 상품등록 XML 생성
	Public Function getHomeplusItemRegXML()
		Dim strRst
		'전송 구분 및 반복리스트 건수
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:createNewProduct xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<Product>"
		strRst = strRst & "				<PRODUCT_CODE>"&FItemid&"</PRODUCT_CODE>"				'##*업체상품코드 | 업체에서 제공하는 해당 상품에 대한 Unique한 식별 코드(API 상품 수정을 통하여 수정불가)
		strRst = strRst & "				<s_POS_NAME><![CDATA["&Trim(getItemNameFormat)&"]]></s_POS_NAME>"	'##*상품명(Web) | 웹 판매 상품명
'		strRst = strRst & "				<s_PREFIX>[텐바이텐]</s_PREFIX>"						'##앞 문구 | 상품명 앞에 붙는 문구
		strRst = strRst & "				<s_DESIGN></s_DESIGN>"									'디자인
		strRst = strRst & "				<s_MAK_CORP><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)&"]]></s_MAK_CORP>"	'##*제조사
		strRst = strRst & "				<s_ORIGN>"&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea)&"</s_ORIGN>"		'##*원산지
		strRst = strRst & "				<DIVISION>"&FhDIVISION&"</DIVISION>"	'##*기준카테고리 DIVISION | 최상위 분류코드
		strRst = strRst & "				<GROUP>"&FhGROUP&"</GROUP>"				'##*기준카테고리 GROUP | DIVISION 하위 분류 코드
		strRst = strRst & "				<DEPT>"&FhDEPT&"</DEPT>"				'##*기준카테고리 DEPT | GROUP 하위 분류 코드
		strRst = strRst & "				<CLASS>"&FhCLASS&"</CLASS>"				'##*기준카테고리 CLASS | DEPT 하위 분류 코드
		strRst = strRst & "				<SUBCLASS>"&FhSUBCLASS&"</SUBCLASS>"	'##*기준카테고리 SUBCLASS | CLASS 하위 분류 코드
		strRst = strRst & "				<s_STORENO>"							'##*전시카테고리 | String[] | 전시등록 카테고리 복수 개를 등록할 수 있다. 실제 상품이 전시될 카테고리.
		If (FbrandDepthCode <> "") AND (FbrandDepthCode <> "0") Then
		strRst = strRst & "					<item>"&FbrandDepthCode&"</item>"
		End If
		If (FdepthCode <> "") AND (FdepthCode <> "0") Then
		strRst = strRst & "					<item>"&FdepthCode&"</item>"
		End If
		strRst = strRst & "				</s_STORENO>"
		strRst = strRst & "				<s_BRANDNO><item>134079</item></s_BRANDNO>"	'##브랜드카테고리 | String[] | 브랜드 카테고리 복수 개를 등록할 수 있다
		strRst = strRst & "				<s_STUFF></s_STUFF>"					'소재
		strRst = strRst & "				<i_DES_KIND>1</i_DES_KIND>"				'##상품설명종류 | 0:TEXT (Default) 1:HTML
		strRst = strRst & "				<s_DES><![CDATA["&getHomeplusItemContParamToReg&"]]></s_DES>"	'##*상품상세설명
		strRst = strRst & getHomeplusAddImageParamToReg							'##*이미지정보
		strRst = strRst & "				<d_SDATE>"&DATE()&"</d_SDATE>"			'##*판매시작일 | YYYY-MM-DD
		strRst = strRst & "				<i_TAXCODE>"&CHKIIF(FVatInclude="N","0","1")&"</i_TAXCODE>"		'##*과세유무 | 0: 비과세, 1:과세
		strRst = strRst & "				<ITEMS>"&getHomeplusOptionParamToReg&"</ITEMS>"					'*ITEM(옵션) | ITEM 정보. 상품에 옵션항목이 없더라도 한 개라도 입력하여야 한다.
		strRst = strRst & "				<c_HARMFUL_YN>N</c_HARMFUL_YN>"			'##성인상품여부 | Y: 성인상품, N: 성인상품 아님(Default)
		strRst = strRst & getItemKeyword										'##검색 유의어 | 상품검색 시 상품명 이외에 해당 상품이 검색되도록 검색 유사어 지정
		strRst = strRst & "				<c_COOP_SEND_YN>Y</c_COOP_SEND_YN>"		'##가격비교사이트 노출여부 | 가격비교 사이트에 해당 상품이 노출될 지 여부..Y: 가격비교사이트 노출, N: 가격비교사이트 비 노출(default)
'		strRst = strRst & "				<DELIVERY_SEQ></DELIVERY_SEQ>"			'하위업체코드 | 업체 벤더 별 하위 업체코드 필수 값이 아니며, 미 입력 시 기본배송 하위업체 코드로 자동입력 하위업체 코드 등록 시 하위업체 코드 등록됨
		strRst = strRst & "				<FIELD_SKIP>false</FIELD_SKIP>"			'##상품정보제공고시 필드정보 생략여부 | true이면 생략 false이면 생략 안 함 false일 경우 FIELDS 데이터를 정확히 입력하여 전송 하여야 한다
		strRst = strRst & getHomeplusItemInfoCdToReg							'##상품정보제공고시 필드정보 | 상품정보제공 고시를 위한 필드정보
		strRst = strRst & "			</Product>"
		strRst = strRst & "		</m:createNewProduct>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
'response.write strRst
'response.end
		getHomeplusItemRegXML = strRst
	End Function

	'// 상품수정 XML 생성
	Public Function getHomeplusItemEditXML()
		Dim strRst
		'전송 구분 및 반복리스트 건수
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:updateProduct xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<Product>"
		strRst = strRst & "				<i_STYLE>"&FHomeplusGoodno&"</i_STYLE>"				'*스타일번호 | 상품등록 시 리턴 한 업체상품코드정보와 매핑 되는 홈플러스 상품(스타일)번호
		strRst = strRst & "				<PRODUCT_CODE>"&FItemid&"</PRODUCT_CODE>"				'##*업체상품코드 | 업체에서 제공하는 해당 상품에 대한 Unique한 식별 코드(API 상품 수정을 통하여 수정불가)
		strRst = strRst & "				<s_POS_NAME><![CDATA["&Trim(getItemNameFormat)&"]]></s_POS_NAME>"	'##*상품명(Web) | 웹 판매 상품명
'		strRst = strRst & "				<s_PREFIX>[텐바이텐]</s_PREFIX>"						'##앞 문구 | 상품명 앞에 붙는 문구
		strRst = strRst & "				<s_DESIGN></s_DESIGN>"									'디자인
		strRst = strRst & "				<s_MAK_CORP><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)&"]]></s_MAK_CORP>"	'##*제조사
		strRst = strRst & "				<s_ORIGN>"&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea)&"</s_ORIGN>"		'##*원산지
		strRst = strRst & "				<DIVISION>"&FhDIVISION&"</DIVISION>"	'##*기준카테고리 DIVISION | 최상위 분류코드
		strRst = strRst & "				<GROUP>"&FhGROUP&"</GROUP>"				'##*기준카테고리 GROUP | DIVISION 하위 분류 코드
		strRst = strRst & "				<DEPT>"&FhDEPT&"</DEPT>"				'##*기준카테고리 DEPT | GROUP 하위 분류 코드
		strRst = strRst & "				<CLASS>"&FhCLASS&"</CLASS>"				'##*기준카테고리 CLASS | DEPT 하위 분류 코드
		strRst = strRst & "				<SUBCLASS>"&FhSUBCLASS&"</SUBCLASS>"	'##*기준카테고리 SUBCLASS | CLASS 하위 분류 코드
		strRst = strRst & "				<s_STORENO>"							'##*전시카테고리 | String[] | 전시등록 카테고리 복수 개를 등록할 수 있다. 실제 상품이 전시될 카테고리.
		If FbrandDepthCode <> "" Then
		strRst = strRst & "					<item>"&FbrandDepthCode&"</item>"
		End If
		If FdepthCode <> "" Then
		strRst = strRst & "					<item>"&FdepthCode&"</item>"
		End If
		strRst = strRst & "				</s_STORENO>"
		strRst = strRst & "				<s_BRANDNO><item>134079</item></s_BRANDNO>"	'##브랜드카테고리 | String[] | 브랜드 카테고리 복수 개를 등록할 수 있다
		strRst = strRst & "				<s_STUFF></s_STUFF>"					'소재
		strRst = strRst & "				<i_DES_KIND>1</i_DES_KIND>"				'##상품설명종류 | 0:TEXT (Default) 1:HTML
		strRst = strRst & "				<s_DES><![CDATA["&getHomeplusItemContParamToReg&"]]></s_DES>"	'##*상품상세설명
		strRst = strRst & getHomeplusAddImageParamToReg							'##*이미지정보
		strRst = strRst & "				<i_IMAGE_UPDATE>1</i_IMAGE_UPDATE>"		'0 : 이미지 업데이트 안됨 1: 이미지 갱신 필요
		strRst = strRst & "				<d_SDATE>"&DATE()&"</d_SDATE>"			'##*판매시작일 | YYYY-MM-DD
		strRst = strRst & "				<c_HARMFUL_YN>N</c_HARMFUL_YN>"			'##성인상품여부 | Y: 성인상품, N: 성인상품 아님(Default)
		strRst = strRst & getItemKeyword										'##검색 유의어 | 상품검색 시 상품명 이외에 해당 상품이 검색되도록 검색 유사어 지정
		strRst = strRst & "				<c_COOP_SEND_YN>Y</c_COOP_SEND_YN>"		'##가격비교사이트 노출여부 | 가격비교 사이트에 해당 상품이 노출될 지 여부..Y: 가격비교사이트 노출, N: 가격비교사이트 비 노출(default)
		strRst = strRst & "				<s_BRAND></s_BRAND>"					'홈플러스 에서 지정하여 주는 브랜드 이름 값을 넣어준다.
'		strRst = strRst & "				<DELIVERY_SEQ></DELIVERY_SEQ>"			'하위업체코드 | 업체 벤더 별 하위 업체코드 필수 값이 아니며, 미 입력 시 기본배송 하위업체 코드로 자동입력 하위업체 코드 등록 시 하위업체 코드 등록됨
		strRst = strRst & "				<FIELD_SKIP>false</FIELD_SKIP>"			'##상품정보제공고시 필드정보 생략여부 | true이면 생략 false이면 생략 안 함 false일 경우 FIELDS 데이터를 정확히 입력하여 전송 하여야 한다
		strRst = strRst & getHomeplusItemInfoCdToReg							'##상품정보제공고시 필드정보 | 상품정보제공 고시를 위한 필드정보
		strRst = strRst & "			</Product>"
		strRst = strRst & "		</m:updateProduct>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
'response.write strRst
'response.end
		getHomeplusItemEditXML = strRst
	End Function

	'// 상품 이미지 수정 XML 생성
	Public Function getHomeplusItemEditImgXML
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:updateImage xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<I_STYLENO>"&FHomeplusGoodno&"</I_STYLENO>"
		strRst = strRst & getHomeplusAddImageParamToEDT							'##*이미지정보
		strRst = strRst & "		</m:updateImage>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
		getHomeplusItemEditImgXML = strRst
	End Function

	Public Function getHomeplusItemEditOPTXML
		Dim strRst
		'전송 구분 및 반복리스트 건수
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:updateProductItem xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<I_STYLENO>"&FHomeplusGoodno&"</I_STYLENO>"		'*스타일번호
		strRst = strRst & getHomeplusOptionParamToEDT								'*아이템 | 추가/수정 될 아이템(옵션)정보.추가 아이템 정보의 I_SIZE는 기존 등록된 I_SIZE와 달라야 합니다.
		strRst = strRst & "		</m:updateProductItem>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
		getHomeplusItemEditOPTXML = strRst
	End Function

	Public Function fngetMustPrice
		Dim strRst, GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			fngetMustPrice = Forgprice
		Else
			fngetMustPrice = FSellCash
		End If
	End Function

	Public Function getHomeplusItemInfoCdToReg()
		Dim buf, strSQL, mallinfoCd, infoContent
		strSQL = ""
		strSQL = strSQL & " SELECT top 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE WHEN (M.infoCdAdd='00000') AND (F.chkdiv ='N') THEN '0' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00000') AND (F.chkdiv ='Y') THEN '1' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00007') AND (F.chkdiv ='N') THEN '0' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00007') AND (F.chkdiv ='Y') THEN '1' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00002') THEN '상세페이지참고' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='99999') THEN '의류' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00016') THEN '1' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='10000') THEN '공정거래위원회 고시(소비자분쟁해결기준)에 의거하여 보상해 드립니다.' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00001') THEN I.itemname " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00003') AND ((IC.safetyyn= 'N') OR IC.safetyyn= '') THEN '0' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00003') AND (IC.safetyyn= 'Y') THEN '1' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00021') AND ((IC.safetyyn= 'N') OR IC.safetyyn= '') THEN 'N' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00021') AND (IC.safetyyn= 'Y') THEN 'Y' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00004') AND (IC.safetyyn= 'Y') AND (M.mallinfocd <> '125018') THEN '본 제품은 KC 안전인증을 필함' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00004') AND (IC.safetyyn= 'Y') AND (M.mallinfocd= '125018') THEN '화장품법에 따른 식품의약품안전청 심사를 필함' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00005') AND (IC.safetyyn= 'Y') THEN IC.safetyNum " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00005') AND ((IC.safetyyn= 'N') OR IC.safetyyn= '') THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00008') THEN '61502' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00011') THEN '61201' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00009') THEN '61301' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00014') THEN '61401' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00017') AND (F.chkdiv ='Y') THEN '본 제품은 광고사전심의를 필함' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00019') AND (F.chkdiv ='Y') THEN '식품위생법에 따른 수입신고를 필함' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00020') AND (F.chkdiv ='Y') THEN '' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00018') AND (F.chkdiv ='Y') THEN infocontent  " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00006') THEN '0' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035'  " & vbcrlf
		strSQL = strSQL & " ELSE convert(varchar(500),F.infocontent) END AS infocontent  " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd  " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"'  " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'homeplus' and IC.itemid='"&FItemid&"'  " & vbcrlf
		strSQL = strSQL & " and not (F.chkdiv ='N' and (M.mallinfocd in ('134005', '133006', '130005', '113011', '101012', '102008', '107010', '108010', '103008', '104007', '105008', '106008', '135007', '131004', '131013', '131014', '132006', '115013', '115015', '115005', '116013', '111009'))) " & vbcrlf
		strSQL = strSQL & " and not (((IC.safetyyn= 'N') OR IC.safetyyn= '') and (M.mallinfocd in ('113016', '113017', '101003', '101004', '107015', '107016', '108017', '108018', '103003', '103004', '104003', '104004', '105003', '105004', '106003', '106004', '135003', '135004', '131010', '131011', '125018', '125019', '116017', '116018'))) " & vbcrlf
		rsget.Open strSQL,dbget,1
		If Not(rsget.EOF or rsget.BOF) then
			buf = buf & "<FIELDS>"
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
			    buf = buf &"	<item>"
				buf = buf & " 		<FILED_ID>"&mallinfoCd&"</FILED_ID>"
				buf = buf & " 		<VALUE><![CDATA["&infoContent&"]]></VALUE>"
				buf = buf &" 	</item>"
				rsget.MoveNext
			Loop
			buf = buf & "</FIELDS>"
		End If
		rsget.Close
		getHomeplusItemInfoCdToReg = buf
	End Function


	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

End Class

Class CHomeplus
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectMakerid
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

	Public Sub getHomeplusNotRegOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			''' 옵션 추가금액 있는경우 등록 불가. //옵션 전체 품절인 경우 등록 불가.
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
            'addSql = addSql & " where optAddCNT>0"
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"

            ''' 2013/05/29 특정품목 등록 불가 (화장품, 식품류)
            addSql = addSql & " and isNULL(c.infodiv,'') not in ('','18','20','21','22')"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, isNULL(R.homeplusStatCD,-9) as homeplusStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(pm.hDIVISION, '') as hDIVISION, isnull(pm.hGROUP, '') as hGROUP, isnull(pm.hDEPT, '') as hDEPT, isnull(pm.hCLASS, '') as hCLASS, isnull(pm.hSUBCLASS, '') as hSUBCLASS, isnull(pm.hCATEGORY_ID, '') as hCATEGORY_ID "
		strSql = strSql & "	, isnull(hm.depthCode, '') as depthCode, isnull(bm.depthCode, '') as brandDepthCode "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_homeplus_brandCategory_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as pm on pm.tenCateLarge=i.cate_large and pm.tenCateMid=i.cate_mid and pm.tenCateSmall=i.cate_small and c.infodiv = pm.infodiv "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_cate_mapping as hm on hm.tenCateLarge=i.cate_large and hm.tenCateMid=i.cate_mid and hm.tenCateSmall=i.cate_small and c.infodiv = hm.infodiv "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_brandCategory_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_regItem R on i.itemid=R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.deliverytype not in ('7')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000)))"
		ELSE
		    strSql = strSql & "	and (i.deliveryType <> 9)"
	    END IF
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "												'플라워/화물배송 상품 제외
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' and i.itemdiv not in ('06', '16') "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (Select makerid From db_temp.dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (Select itemid From db_temp.dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and i.itemid not in (Select itemid From db_etcmall.dbo.tbl_homeplus_regItem where homeplusStatCD>3) "
		strSql = strSql & "	and uc.isExtUsing='Y'"  ''20130304 브랜드 제휴사용여부 Y만.
		strSql = strSql & addSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHomeplusItem
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
                FOneItem.FHomeplusStatCD	= rsget("HomeplusStatCD")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.FhDIVISION			= rsget("hDIVISION")
                FOneItem.FhGROUP			= rsget("hGROUP")
                FOneItem.FhDEPT				= rsget("hDEPT")
                FOneItem.FhCLASS			= rsget("hCLASS")
                FOneItem.FhSUBCLASS			= rsget("hSUBCLASS")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FdepthCode			= rsget("depthCode")
                FOneItem.FbrandDepthCode	= rsget("brandDepthCode")
		End If
		rsget.Close
	End Sub

	Public Sub getHomeplusEditOneItem
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
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, m.HomeplusGoodNo, m.Homeplusprice, m.HomeplusSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(pm.hDIVISION, '') as hDIVISION, isnull(pm.hGROUP, '') as hGROUP, isnull(pm.hDEPT, '') as hDEPT, isnull(pm.hCLASS, '') as hCLASS, isnull(pm.hSUBCLASS, '') as hSUBCLASS, isnull(pm.hCATEGORY_ID, '') as hCATEGORY_ID "
		strSql = strSql & "	, isnull(hm.depthCode, '') as depthCode, isnull(bm.depthCode, '') as brandDepthCode "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.itemdiv = '06' or i.itemdiv = '16' "
		strSql = strSql & "		or isNULL(c.infodiv,'') in ('','18','20','21','22') "
		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_Homeplus_regitem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as pm on pm.tenCateLarge=i.cate_large and pm.tenCateMid=i.cate_mid and pm.tenCateSmall=i.cate_small and c.infodiv = pm.infodiv "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_cate_mapping as hm on hm.tenCateLarge=i.cate_large and hm.tenCateMid=i.cate_mid and hm.tenCateSmall=i.cate_small and c.infodiv = hm.infodiv "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_brandCategory_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.HomeplusGoodNo is Not Null "									'#등록 상품만
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHomeplusItem
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
				FOneItem.FHomeplusGoodNo	= rsget("HomeplusGoodNo")
				FOneItem.FHomeplusprice		= rsget("Homeplusprice")
				FOneItem.FHomeplusSellYn	= rsget("HomeplusSellYn")
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
                FOneItem.FhDIVISION			= rsget("hDIVISION")
                FOneItem.FhGROUP			= rsget("hGROUP")
                FOneItem.FhDEPT				= rsget("hDEPT")
                FOneItem.FhCLASS			= rsget("hCLASS")
                FOneItem.FhSUBCLASS			= rsget("hSUBCLASS")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FdepthCode			= rsget("depthCode")
                FOneItem.FbrandDepthCode	= rsget("brandDepthCode")
                FOneItem.Fregitemname		= rsget("regitemname")
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

Function getHomplusGoodNo(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 homeplusgoodno FROM db_etcmall.dbo.tbl_homeplus_regitem WHERE itemid = '"&iitemid&"' "
	rsget.Open strSql, dbget, 1
		getHomplusGoodNo = rsget("homeplusgoodno")
	rsget.Close
End Function
%>