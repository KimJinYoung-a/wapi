<%
CONST CMAXMARGIN = 14.9
CONST CMALLNAME = "ezwel"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CEzwelMARGIN = 10									'이지웰페어 마진 10%
CONST cspCd		= "10040413"							'CP업체코드(이지웰 발급)
CONST crtCd		= "8e5a6dbdd27efb49fc600c293884ef47"	'보안코드(이지웰 발급)
CONST cspDlvrId	= "10040413"							'배송처코드

Class CEzwelItem
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
	Public FezwelStatCD
	Public FinfoDiv
	Public FDeliveryType
	Public FdepthCode
	Public FbasicimageNm

	Public Function MustPrice()
		Dim GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
	End Function

	'// Ezwel 판매여부 반환
	Public Function getEzwelSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getEzwelSellYn = "Y"
			Else
				getEzwelSellYn = "N"
			End If
		Else
			getEzwelSellYn = "N"
		End If
	End Function

    Function getEzwelAddSuplyPrice(addprice)
		getEzwelAddSuplyPrice = CLNG((addprice)  * (100-CEzwelMARGIN) / 100)
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
'				strSql = strSql & " 	and optaddprice=0 "
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

	Public Function IsFreeBeasong()
		IsFreeBeasong = False
		If (FdeliveryType=2) or (FdeliveryType=4) or (FdeliveryType=5) then				'2(텐무), 4,5(업무)
			IsFreeBeasong = True
		End If
'		If (FSellcash>=30000) then IsFreeBeasong=True
		If (FdeliveryType=9) Then														'업체조건
'			If (Clng(FSellcash) >= Clng(FdefaultfreeBeasongLimit)) then
'				IsFreeBeasong=True
'			End If
			IsFreeBeasong = False
		End If
    End Function

	'상품설명 파라메터 생성
	Public Function getEzwelItemContParam()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
		strRst = strRst & ("<p><center><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_ezwel.jpg""></center></p><br>")
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
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ezwel.jpg"">")
		strRst = strRst & ("</div>")
		getEzwelItemContParam = strRst
	End Function

	'// 상품등록: 상품추가이미지 파라메터 생성
	Public Function getEzwelAddImageParam()
		Dim strRst, strSQL, i
		strRst = ""
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If

		strRst = strRst &"	<imgPath>"&FbasicImage&"</imgPath>"		'메인이미지경로 | ex)http://www.ezwel.com/img/goods1.gif
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		'추가이미지경로1~3
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					strRst = strRst &"	<imgPath"&i&">http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") &"</imgPath"&i&">"
				End If
				rsget.MoveNext
				If i >= 3 Then Exit For
			Next

		End If
		rsget.Close
		getEzwelAddImageParam = strRst
	End Function

	'상품품목정보
    public function getEzwelItemInfoCd()
		Dim buf1, buf2, buf3, strSQL, mallinfoCd, infoContent, mallinfodiv
		strSQL = ""
		strSQL = strSQL & " SELECT top 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN IC.safetyNum " & vbcrlf
		strSql = strSql & "		 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00001') THEN '상세페이지참고' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='10000') THEN '공정거래위원회 고시(소비자분쟁해결기준)에 의거하여 보상해 드립니다.' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035'  " & vbcrlf
		strSQL = strSQL & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd  " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"'  " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='" & FItemid &"' " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'ezwel' and IC.itemid='"&FItemid&"'  " & vbcrlf
		rsget.Open strSQL,dbget,1
		mallinfodiv = "10" & rsget("mallinfodiv")
		If Not(rsget.EOF or rsget.BOF) then
			buf1 = "<goodsGrpCd>"&mallinfodiv&"</goodsGrpCd>"		'##*상품고시 코드 | 별도첨부 
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
				buf2 = buf2 & " 		<arrLayoutDesc><![CDATA["&Server.URLEncode(infoContent)&"]]></arrLayoutDesc>"
				buf2 = buf2 & " 		<arrLayoutSeq>"&mallinfoCd&"</arrLayoutSeq>"
				rsget.MoveNext
			Loop
			buf3 = buf1 & buf2
		End If
		rsget.Close
		getEzwelItemInfoCd = buf3
	End Function

   Public Function getEzwelOptionParam()
		Dim strSql, strRst, i, optLimit
    	Dim buf, optDc, itemsu, addprice, addbuyprice, optTaxCk, optTax, optUsingCk, optUsing

    	buf = ""
		If FoptionCnt>0 then
			strSql = ""
			strSql = strSql &  "SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
			strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
			rsget.Open strSql,dbget,1

			optDc = ""
			optLimit = ""
			If FVatInclude = "N" Then
				optTaxCk = "N"
			Else
			 	optTaxCk = "Y"
			End If

			If Not(rsget.EOF or rsget.BOF) Then
				Do until rsget.EOF
				    optLimit = rsget("optLimit")
				    optLimit = optLimit-5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = 999   ''2013/06/12 재고관리여부 모두 Y로 변경 되므로
					optUsingCk = "Y"
					optDc = optDc & Server.URLEncode(rpTxt(db2Html(rsget("optionname"))))
					
					itemsu = itemsu & optLimit
					addprice = addprice & rsget("optaddprice")
					addbuyprice = addbuyprice & getEzwelAddSuplyPrice(rsget("optaddprice"))
					optTax = optTax & optTaxCk
					optUsing = optUsing & optUsingCk

					rsget.MoveNext
					If Not(rsget.EOF) Then
						 optDc	= optDc & "|"
						 itemsu = itemsu & "|"
						 addprice = addprice & "|" 
						 addbuyprice = addbuyprice & "|"
						 optTax	= optTax & "|"
						 optUsing = optUsing & "|"
					End If
				Loop
			End If
			rsget.Close
			buf = buf & "		<useYn>Y</useYn>"												'상품옵션사용여부 | 옵션이 있을경우(Y) 없을경우(N)
			buf = buf & "		<arrOptionCdNm>"&Server.URLEncode("선택")&"</arrOptionCdNm>"	'상품옵션명
			buf = buf & "		<arrOptionContent>"&optDc&"</arrOptionContent>"					'상품옵션 내용
			buf = buf & "		<arrOptionUseYn>Y</arrOptionUseYn>"								'옵션별에 따른 사용여부 | Y:N
			buf = buf & "		<arrOptionAddAmt>"&itemsu&"</arrOptionAddAmt>"					'*(옵션이 존재하는 경우만) | 상품옵션 수량 | Default: 10000 
			buf = buf & "		<arrOptionAddPrice>"&addprice&"</arrOptionAddPrice>"			'상품옵션추가가격
			buf = buf & "		<arrOptionAddBuyPrice>"&addbuyprice&"</arrOptionAddBuyPrice>"	'공급가
			buf = buf & "		<arrOptionAddTaxYn>"&optTax&"</arrOptionAddTaxYn>"				'과세여부 | 과세(Y), 면세(N), 영세(숫자 0) 
			buf = buf & "		<arrOptionFullUseYn>"&optUsing&"</arrOptionFullUseYn>"			'옵션 상세별에 따른 사용여부 |||    Y|Y|Y:N|N:N 
		Else
			buf = buf & "		<useYn>N</useYn>"												'상품옵션사용여부 | 옵션이 있을경우(Y) 없을경우(N)
		End If
		getEzwelOptionParam = buf
    End Function

	'상품등록/수정 XML 생성
	Public Function getEzwelItemRegXML(ezwelMethod)
		Dim strRst
		Dim EzwelStatus
		Select Case ezwelMethod
			Case "Reg"			EzwelStatus = "1001"
			Case "SellY"		EzwelStatus = "1002"
			Case "SellN"		EzwelStatus = "1005"
		End Select
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "	<dataSet>"
		strRst = strRst & "		<cspCd>"&cspCd&"</cspCd>"					'##*CP 업체코드 | 이지웰 발급(고정값)
		If ezwelMethod <> "Reg" Then
		strRst = strRst & "		<goodsCd>"&FEzwelGoodno&"</goodsCd>"		'##*값이 존재하면 수정 존재하지 않으면 입력 | 상품코드 | 이지웰 상품코드
		End If
		strRst = strRst & "		<cspGoodsCd>"&FItemid&"</cspGoodsCd>"		'##업체상품코드
		strRst = strRst & "		<goodsNm><![CDATA["&Server.URLEncode(Trim(getItemNameFormat))&"]]></goodsNm>"	'##*상품명
		strRst = strRst & "		<taxYn>"&CHKIIF(FVatInclude="N","N","Y")&"</taxYn>"							'##*과세여부 | 과세(Y), 면세(N), 영세(숫자 0)
		strRst = strRst & "		<goodsStatus>"&EzwelStatus&"</goodsStatus>"									'##상품상태 | 등록(1001), 판매중(1002), 판매중지(1005), 삭제(1006)
		strRst = strRst & "		<dlvrPrice>"&CHKIIF(IsFreeBeasong=False,"2500","0")&"</dlvrPrice>"			'##배송가격
		strRst = strRst & "		<dlvrPriceApplYn>"&CHKIIF(IsFreeBeasong=True,"Y","P")&"</dlvrPriceApplYn>"	'##*착불/선결제/무료 | 무료: Y/ 소비자부담:N /착불만: A /선결제만:P 
		strRst = strRst & "		<realSalePrice>"&Clng(GetEzwel10wonDown(MustPrice/10)*10)&"</realSalePrice>"	'##*판매가
		strRst = strRst & "		<normalSalePrice>"&Clng(GetRaiseValue(ForgPrice/10)*10)&"</normalSalePrice>"'##*정상(시중)가
		strRst = strRst & "		<brandNm><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("상품설명 참조"),Server.URLEncode(rpTxt(Fmakername)))&"]]></brandNm>"	'##브랜드명
		strRst = strRst & "		<buyPrice>"&GetEzwelBuyPrice(Clng(GetEzwel10wonDown(MustPrice/10)*10))&"</buyPrice>"		'##*공급가(매입가)
		strRst = strRst & "		<modelNum>"&FItemid&"</modelNum>"										'상품모델
		strRst = strRst & "		<orginNm><![CDATA["&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),Server.URLEncode("상품설명 참조"),Server.URLEncode(Fsourcearea))&"]]></orginNm>"	'##원산지
		strRst = strRst & "		<mafcNm><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("상품설명 참조"),Server.URLEncode(rpTxt(Fmakername)))&"]]></mafcNm>"		'##제조사
		strRst = strRst & "		<enterAmt>10000</enterAmt>"						'##*입고수량 | Default: 10000
		strRst = strRst & "		<cspDlvrId>"&cspDlvrId&"</cspDlvrId>"			'##출고지ID | 이지웰 발급(고정값)
		strRst = strRst & "		<goodsDesc><![CDATA["&Server.URLEncode(getEzwelItemContParam())&"]]></goodsDesc>"	'##상품설명
		If (ezwelMethod <> "Reg") Then		'2014-12-02 김진영 추가 | 이미지 전송 시간 오래걸림
			If isImageChanged Then
				strRst = strRst & getEzwelAddImageParam()
			End If
		Else
			strRst = strRst & getEzwelAddImageParam()	
		End If
		strRst = strRst & "		<ctgCd>"&FDepthCode&"</ctgCd>"					'##*관리카테고리 | 별도첨부
		strRst = strRst & "		<dispCtgCd>"&FDepthCode&"</dispCtgCd>"			'##*전시 카테고리 | 별도첨부
		strRst = strRst & getEzwelItemInfoCd()									'##상품정보제공고시 필드정보 | 상품정보제공 고시를 위한 필드정보
		strRst = strRst & getEzwelOptionParam()
		strRst = strRst & "		<marginRate>"&CEzwelMARGIN&"</marginRate>"		'##현아대리님 10%라고 답변 | *마진률 | 9.0 
		strRst = strRst & "</dataSet>"
		getEzwelItemRegXML = strRst
		'response.write strRst
		'response.end
	End Function

End Class

Class CEzwel
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


	Public Sub getEzwellNotRegOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			'옵션 전체 품절인 경우 등록 불가.
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
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"
            addSql = addSql & " and isNULL(c.infodiv,'') not in ('','18','20','21','22')"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, isNULL(R.ezwelStatCD,-9) as ezwelStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(bm.depthCode, '') as depthCode "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_Newcate_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_regItem R on i.itemid=R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.deliverytype not in ('7')"
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "												'플라워/화물배송 상품 제외
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' and i.itemdiv not in ('06', '16') "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (Select makerid From db_etcmall.dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (Select itemid From db_etcmall.dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and i.itemid not in (Select itemid From db_etcmall.dbo.tbl_ezwel_regItem where ezwelStatCD>3) "
		strSql = strSql & "	and uc.isExtUsing='Y'"  ''20130304 브랜드 제휴사용여부 Y만.
		strSql = strSql & addSql																				'카테고리 매칭 상품만
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CEzwelItem
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
                FOneItem.FezwelStatCD		= rsget("ezwelStatCD")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FdepthCode			= rsget("depthCode")
                FOneItem.FbasicimageNm 		= rsget("basicimage")
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

Function GetEzwel10wonDown(value)
   	GetEzwel10wonDown = Fix(value/10)*10
End Function

Function rpTxt(checkvalue)
	Dim v
	v = checkvalue
	if Isnull(v) then Exit function
	
    On Error resume Next
    v = replace(v, "&", "&amp;")
    v = Replace(v, """", "&quot;")
    v = Replace(v, "'", "&apos;")
    v = replace(v, "<", "&lt;")
    v = replace(v, ">", "&gt;")
    v = replace(v, ":", "：")
    rpTxt = v
End Function

Function GetEzwelBuyPrice(value)
   	GetEzwelBuyPrice = Clng(value - (value / 10))
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