<%
'' 배송정책  3만원 이하 2500
CONST CMAXMARGIN = 14.9			'' MaxMagin임.. '(롯데iMall 10%)
CONST CMAXLIMITSELL = 5        '' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CMALLNAME = "lotteimall"
CONST CLTIMALLMARGIN = 11       ''마진 11%
CONST CHEADCOPY = "Design Your Life! 새로운 일상을 만드는 감성생활브랜드 텐바이텐" ''생활 감성채널 텐바이텐
CONST CPREFIXITEMNAME ="[텐바이텐]"
CONST CitemGbnKey ="K1099999" ''상품구분키 ''하나로 통일
CONST CUPJODLVVALID = TRUE   ''업체 조건배송 등록 가능여부

CONST ENTP_CODE = "011799"                                    '' 협력사코드
CONST MD_CODE   = "0168"                                      '' MD_Code
CONST BRAND_CODE   = "1099329"                                '' 롯데에 받아야함
CONST BRAND_NAME   = "텐바이텐(10x10)"                        '' 롯데에 받아야함
CONST MAKECO_CODE  = "9999"                                   '' 롯데에 받아야함
CONST CDEFALUT_STOCK = 999       '' 재고관리 수량 기본 999 (한정 아닌경우)

Class CLotteiMallItem
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
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FLTiMallGoodNo
	Public FLTiMallTmpGoodNo
	Public FLtiMallPrice
	Public FLtiMallSellYn
	Public FregedOptCnt
	Public FaccFailCNT
	Public FlastErrStr
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FmaySoldOut
	Public FitemGbnKey
	Public FLtiMallStatCD

	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum

	Public FRegedOptionname
	Public FRegedItemname
	Public FNewitemname
	Public FItemnameChange
	Public FItemoption
	Public FOptisusing
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold
	Public FOptaddprice
	Public FRealSellprice
	Public FNewItemid
	Public FOptionname
	Public FAdultType

	'// 품절여부
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	'// 품절여부
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Public Function IsOptionSoldOut()
		CONST CLIMIT_SOLDOUT_NO = 5
		IsOptionSoldOut = false
		IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno - Foptlimitsold <= CLIMIT_SOLDOUT_NO))
	End Function

	Public Function isDiffName
		isDiffName = false
		If (Fitemname <> FRegedItemname) OR (FOptionname <> FRegedOptionname) Then
			isDiffName = True
		End If
	End Function

	Public Function getRealItemname
		If FitemnameChange = "" Then
			getRealItemname = FNewitemname
		Else
			getRealItemname = FItemnameChange
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

	Function getItemNameFormat()
		Dim buf
		buf = replace(FItemName,"'","")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","프로")
		buf = replace(buf,"[무료배송]","")
		buf = replace(buf,"[무료 배송]","")
		getItemNameFormat = buf
	End Function

	Function getItemOptNameFormat()
		Dim buf
		buf = replace(getRealItemname,"'","")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","프로")
		buf = replace(buf,"[무료배송]","")
		buf = replace(buf,"[무료 배송]","")
		getItemOptNameFormat = buf
	End Function

	Public Function getLTiMallSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FsellYn = "Y" and FisUsing = "Y" Then
			If FLimitYn="N" or (FLimitYn = "Y" and FLimitNo - FLimitSold > CMAXLIMITSELL) Then
				getLTiMallSellYn = "Y"
			Else
				getLTiMallSellYn = "N"
			End if
		Else
			getLTiMallSellYn = "N"
		End If
	End Function

	'// 검색어
	Public Function getItemKeyword()
		Dim arrRst, arrRst2, q, p, r, divBound1, divBound2, divBound3, Keyword1, Keyword2, Keyword3, strRst
		If trim(Fkeywords) = "" Then Exit Function
		'2015-05-06 김진영 수정 내용중 %가 들어가면 에러남
		Fkeywords  = replace(Fkeywords,"%", "％")
		Fkeywords  = replace(Fkeywords,chr(13), "")
		Fkeywords  = replace(Fkeywords,chr(10), "")
		Fkeywords  = replace(Fkeywords,chr(9), " ")

		If Len(Fkeywords) > 50 Then
			arrRst = Split(Fkeywords,",")
			If Ubound(arrRst) = 0 then
				'구분이 공백일 경우
				arrRst2 = split(arrRst(0)," ")
				If Ubound(arrRst2) > 0 then
					arrRst = split(Fkeywords," ")
				'2013-10-22 김진영 수정..ex)826121, 826124
				Else
					'구분이 세미콜론일 경우
					arrRst2 = split(arrRst(0),";")
					If Ubound(arrRst2) > 0 then
						arrRst = split(Fkeywords,";")
					End If
				End If
			End If

			If Ubound(arrRst) = 2 Then	'2015-06-29 김진영 수정 : ex)769674
				Keyword1 = arrRst(0)
				Keyword2 = arrRst(1)
				Keyword3 = arrRst(2)
			Else
				'키워드 1
				divBound1 = Cint(Ubound(arrRst)/3)
				For q = 0 to divBound1
					Keyword1 = Keyword1&arrRst(q)&","
				Next
				If Right(keyword1,1) = "," Then
					keyword1 = Left(keyword1,Len(keyword1)-1)
				End If
				'키워드 2
				divBound2 = divBound1 + 1
				For p = divBound2 to divBound2 + divBound1
					Keyword2 = Keyword2&arrRst(p)&","
				Next
				If Right(keyword2,1) = "," Then
					keyword2 = Left(keyword2,Len(keyword2)-1)
				End If

				'키워드 3
				divBound3 = divBound2 + divBound1
				For r = divBound3 to Ubound(arrRst)
					Keyword3 = Keyword3&arrRst(r)&","
				Next
				If Right(keyword3,1) = "," Then
					keyword3 = Left(keyword3,Len(keyword3)-1)
				End If
			End If

			strRst = ""
			strRst = strRst & "&sch_kwd_1_nm="&Keyword1
			strRst = strRst & "&sch_kwd_2_nm="&Keyword2
			strRst = strRst & "&sch_kwd_3_nm="&Keyword3
			getItemKeyword = strRst
		Else
			strRst = ""
			strRst = strRst & "&sch_kwd_1_nm="&Fkeywords
			strRst = strRst & "&sch_kwd_2_nm="
			strRst = strRst & "&sch_kwd_3_nm="
			getItemKeyword = strRst
		End If
	End Function

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

	'// 상품등록: MD상품군 및 전시 카테고리 파라메터 생성(상품등록용)
	Public Function getLotteiMallCateParamToReg()
		Dim strSql, strRst, i, ogrpCode
		strSql = ""
		strSql = strSql & " SELECT TOP 6 c.groupCode, m.dispNo, c.disptpcd "
		strSql = strSql & " FROM db_item.dbo.tbl_lotteimall_cate_mapping as m "
		strSql = strSql & " INNER JOIN db_temp.dbo.tbl_lotteimall_Category as c on m.DispNO = c.DispNO "
		strSql = strSql & " WHERE tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " and tenCateSmall='" & FtenCateSmall & "' "
	    strSql = strSql & " and c.isusing='Y'"
		strSql = strSql & " and c.dispLrgNm = '텐바이텐' "
		strSql = strSql & " ORDER BY disptpcd ASC "           ''''//일반몰을 기본 카테고리로..
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
		    ogrpCode = rsget("groupCode")
			strRst = "&md_gsgr_no=" & ogrpCode
			i = 0
			Do until rsget.EOF
				If (rsget("disptpcd")="10") then
				    strRst = strRst & "&disp_no=" & rsget("dispNo")			'기본 전시카테고리
'				Else
'				    IF (ogrpCode=rsget("groupCode")) then
'					    strRst = strRst & "&disp_no_b=" & rsget("dispNo") 	'추가 전시카테고리
'					End IF
			    End If
				rsget.MoveNext
				i = i + 1
			Loop
		End If
		rsget.Close
		getLotteiMallCateParamToReg = strRst
'rw strRst
'response.end
	End Function

	Public Function getLotteiMallGoodDLVDtParams()
		dim strRst
		strRst = ""
		If (FtenCateLarge="055") or (FtenCateLarge="040") then ''가구/패브릭 15일로
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=15"
		ElseIf (FtenCateLarge="080") or (FtenCateLarge="100") then  ''우먼/베이비 5일
			strRst = strRst & "&dlv_goods_sct_cd=03" 																						'배송상품구분		(*:주문제작03)
			strRst = strRst & "&dlv_dday=5"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="001" or FtenCateMid="004")) then  ''수납/생활> 옷/이불수납 or 주방수납 10일 - 현아씨요청 2013/01/22
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="025") and (FtenCateMid="107")) then  ''디지털 > 기타 스마트기기 케이스  10일 - 현아씨요청 2013/01/22
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="050") and (FtenCateMid="777")) then   ''홈/데코 > 거울   - 미희씨요청 2013/03/08
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="001")) then    ''HOME > 수납/생활 > 보관/정리용품 > 수납장 			주문제작15일 045&cdm=002&cds=001
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=15"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="002")) then    ''HOME > 수납/생활 > 보관/정리용품 > 틈새수납장			주문제작10일
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="005")) then    ''HOME > 수납/생활 > 보관/정리용품 > 잡지꽂이 			주문제작10일
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="001")) then    ''HOME > 수납/생활 > 데코수납 > 우드박스 				주문제작10일
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="007")) then    ''HOME > 수납/생활 > 데코수납 > 인터폰박스 			               주문제작10일
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="050") and (FtenCateMid="060") and (FtenCateSmall="070")) then    ''HOME > 홈/데코 > 소품박스/바구니 > 인터폰박스			주문제작10일 cdl=050&cdm=060&cds=070
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="110") and (FtenCateMid="090") and (FtenCateSmall="040")) then    ''HOME > 감성채널 > DIY > 나무로만들기 				주문제작10일 110&cdm=090&cds=040
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="010")) then   ''수납/생활 > 디자인선반  - 미희씨요청 2013/03/08
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
'		ElseIf (FtenCateLarge="025")  then  ''디지털 10일 - 미희씨요청 2013/01/17
'		    strRst = strRst & "&dlv_goods_sct_cd=03" 																						'배송상품구분		(*:주문제작03)
'		    strRst = strRst & "&dlv_dday=10"
		ElseIf ((FitemDiv="06") or (FitemDiv="16")) then    ''주문(후)제작상품
			strRst = strRst & "&dlv_goods_sct_cd=03"
			If (FrequireMakeDay>7) then
				    strRst = strRst & "&dlv_dday="&CStr(FrequireMakeDay)
			ElseIf (FrequireMakeDay<1) then
				    strRst = strRst & "&dlv_dday=7"
			Else
				    strRst = strRst & "&dlv_dday="&(FrequireMakeDay+1)
			End If
		Else
			strRst = strRst & "&dlv_goods_sct_cd=01" 																						'배송상품구분		(*:일반상품)
			strRst = strRst & "&dlv_dday=3" 																								'배송기일			(*:3일이내)
		End If
		getLotteiMallGoodDLVDtParams = strRst
	End Function

	'// 상품등록: 옵션 파라메터 생성(상품등록용)
	Public function getLotteiMallOptionParamToReg()
		dim strSql, strRst, i, optYn, optNm, optDc, chkMultiOpt, optLimit, optDanPoomCD
		chkMultiOpt = false
		optYn = "N"

		strRst = strRst & "&item_mgmt_yn=" & optYn				'단품관리여부(옵션)
		strRst = strRst & "&opt_nm="							'옵션명
		strRst = strRst & "&item_list="							'옵션상세
		getLotteiMallOptionParamToReg = strRst
	End Function

	'// 상품등록: 상품추가이미지 파라메터 생성(상품등록용)
	Public Function getLotteiMallAddImageParamToReg()
		Dim strRst, strSQL, i
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				If rsget("imgType")="0" then
					strRst = strRst & "&img_url" & i & "=" & Server.URLEncode("http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400"))
				End If
				rsget.MoveNext
				If i >= 5 Then Exit For
			Next
		End If
		rsget.Close
		getLotteiMallAddImageParamToReg = strRst
	End Function

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	Public Function getLotteiMallItemContParamToReg()
		Dim strRst, strSQL, strtextVal
		strRst = Server.URLEncode("<div align=""center"">")
		'2014-01-17 10:00 김진영 탑 이미지 추가
		strRst = strRst & Server.URLEncode("<p><a href=""http://www.lotteimall.com/display/viewDispShop.lotte?disp_no=5100455"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_Ltimall.jpg""></a></p><br>")
		ForderComment = replace(ForderComment,"&nbsp;"," ")
		ForderComment = replace(ForderComment,"&nbsp"," ")
		ForderComment = replace(ForderComment,"&"," ")
		ForderComment = replace(ForderComment,chr(13)," ")
		ForderComment = replace(ForderComment,chr(10)," ")
		ForderComment = replace(ForderComment,chr(9)," ")
		If ForderComment <> "" Then
			strRst = strRst & "- 주문시 유의사항 :<br>" & URLEncodeUTF8(Fordercomment) & "<br>"
		End If

		'#기본 상품설명
		oiMall.FOneItem.Fitemcontent = replace(oiMall.FOneItem.Fitemcontent,"&nbsp;"," ")
		oiMall.FOneItem.Fitemcontent = replace(oiMall.FOneItem.Fitemcontent,"&nbsp"," ")
		oiMall.FOneItem.Fitemcontent = replace(oiMall.FOneItem.Fitemcontent,"&"," ")
		oiMall.FOneItem.Fitemcontent = replace(oiMall.FOneItem.Fitemcontent,chr(13)," ")
		oiMall.FOneItem.Fitemcontent = replace(oiMall.FOneItem.Fitemcontent,chr(10)," ")
		oiMall.FOneItem.Fitemcontent = replace(oiMall.FOneItem.Fitemcontent,chr(9)," ")
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & URLEncodeUTF8(oiMall.FOneItem.Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & URLEncodeUTF8(oiMall.FOneItem.Fitemcontent & "<br>")
			Case Else
				strRst = strRst & URLEncodeUTF8(oiMall.FOneItem.Fitemcontent & "<br>")
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
					strRst = strRst & Server.URLEncode("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0""><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#기본 상품 설명이미지
		If ImageExists(FmainImage) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage & """ border=""0""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage2 & """ border=""0""><br>")

		'#배송 주의사항
		strRst = strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_LTimall.jpg"">")
		strRst = strRst & Server.URLEncode("</div>")
		getLotteiMallItemContParamToReg = "&dtl_info_fcont=" & strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " WHERE mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strtextVal = Server.URLEncode(rsget("textVal"))
			strRst = Server.URLEncode("<div align=""center""><p><a href=""http://www.lotteimall.com/display/viewDispShop.lotte?disp_no=5100455"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_Ltimall.jpg""></a></p><br>") & strtextVal & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_LTimall.jpg""></div>")
			getLotteiMallItemContParamToReg = "&dtl_info_fcont=" & strRst
		End If
		rsget.Close
	End Function

	Public Function getLotteiMallItemInfoCdToReg()
		Dim anjunInfo
        ''안전인증정보(애매함)
        ''2017-01-06 17:15 수정김진영..아이몰에서 sft_cert_org_cd 안 쓴다고 함..아래 주석에서 ->로 주석 표기
		If (Fsafetyyn="Y" and FsafetyDiv<>0) Then
			If (FsafetyDiv=10) Then											'국가통합인증(KC마크)
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=31"					'KS인증 -> KS인증
			Elseif (FsafetyDiv=20) Then										'전기용품 안전인증
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=21"					'전기용품안전인증 -> 안전인증
			Elseif (FsafetyDiv=30) Then										'KPS 안전인증 표시
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=21"					'전기용품안전인증 -> 안전인증
			Elseif (FsafetyDiv=40) Then										'KPS 자율안전 확인 표시
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=22"					'전기용품자율안전확인신고 -> 안전확인
			Elseif (FsafetyDiv=50) Then										'KPS 어린이 보호포장 표시
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=31"					'KS인증 -> KS인증
			Else
				anjunInfo = ""
			End if
			anjunInfo = anjunInfo & "&sft_cert_no="&Server.URLEncode(FsafetyNum)
		End If

		Dim strRst, strSQL
		Dim mallinfoDiv,mallinfoCd,infoContent, mallinfoCdAll, bufTxt
		Dim rsgetmallinfoDiv, newInfodiv

		If Finfodiv = "47" OR Finfodiv = "48" Then
			newInfodiv = "1" + CSTR(Finfodiv)
		Else
			newInfodiv = ""
		End If

		'동일모델의 출시년월 뽑는 쿼리
		Dim YM, ConvertYM, SD
		strSQL = ""
		strSQL = strSQL & " SELECT top 1 F.infocontent, IC.safetyDiv " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		If newInfodiv = "" Then
			strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		Else
			strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv= '"& newInfodiv &"'  " & vbcrlf
		End If
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd AND F.itemid='"&Fitemid&"' " & vbcrlf
		strSQL = strSQL & " where IC.itemid='"&Fitemid&"' and M.mallinfocd = '10011' " & vbcrlf
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly

		If Not(rsget.EOF or rsget.BOF) then
			YM = rsget("infocontent")
			SD = rsget("safetyDiv")
		Else
			YM = "X"
			SD = "X"
		End If
		rsget.Close

		If YM <> "X" Then
		    YM = replace(YM,".","")
		    YM = replace(YM,"/","")
		    YM = replace(YM,"-","")
		    YM = replace(YM," ","")
		    YM = TRIM(YM)

			If isNumeric(Ym) Then
				ConvertYM = Clng(YM)
			Else
				ConvertYM = "X"
			End If
		Else
			ConvertYM = YM
		End If

		strSQL = ""
		strSQL = strSQL & " SELECT TOP 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE " & vbcrlf

		If SD = "10" Then
			'출시년월의 값이 없는 경우
			If ConvertYM = "X" Then
				strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (LEFT(IC.safetyDiv,3)='KCC') THEN IC.safetyNum " & vbcrlf
			'출시년월의 값이 있는 경우
			Else
				'출시년월이 2012년 7월 이전인 경우
				If ConvertYM < 201207 Then
					strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (LEFT(IC.safetyDiv,3)='KCC') THEN '해당없음' " & vbcrlf	 '(맵핑코드가 KCC인증이고), (10x10에서 안전인증코드여부가 Y, 구분이 KC(10), 201207전)일 때
				'출시년월이 2012년 7월 이후인 경우
				ElseIf ConvertYM >= 201207 Then
					strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (LEFT(IC.safetyDiv,3)='KCC') THEN IC.safetyNum " & vbcrlf '(맵핑코드가 KCC인증이고), (10x10에서 안전인증코드여부가 Y, 구분이 KC(10), 201207후)일 때
				End If
			End If
		End If
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10063') THEN IC.safetyNum " & vbcrlf		'10206은 KC인증
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') <> 'Y') AND (M.mallinfoCd= '10063') THEN '해당없음'  " & vbcrlf		'10206은 KC인증
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10205') THEN IC.safetyNum " & vbcrlf		'10206은 KC인증
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') <> 'Y') AND (M.mallinfoCd= '10205') THEN '해당없음'  " & vbcrlf		'10206은 KC인증
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10206') THEN 'KC 안전인증 필'  " & vbcrlf	'10206은 KC인증
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') <> 'Y') AND (M.mallinfoCd= '10206') THEN '해당없음'  " & vbcrlf		'10206은 KC인증
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'N') THEN '해당없음'  " & vbcrlf		'(맵핑코드가 KCC인증이고), (10x10에서 안전인증코드여부가 N)일 때
		strSQL = strSQL & " 	 WHEN M.infoCd='00001' THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN M.infoCd='00002' THEN '원산지와 동일' " & vbcrlf
		strSQL = strSQL & " 	 WHEN F.chkDiv='Y' AND (M.infoCd='19008') THEN '제공함' " & vbcrlf				'귀금속의 가공지
		strSQL = strSQL & " 	 WHEN F.chkDiv='N' AND (M.infoCd='19008') THEN '제공하지 않음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN F.chkDiv='Y' AND (M.infoCd='18008') THEN '기능성 심사 필' " & vbcrlf		'화장품의 기능성 화장품 여부
		strSQL = strSQL & " 	 WHEN F.chkDiv='N' AND (M.infoCd='18008') THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN F.chkDiv='Y' AND (M.infoCd='17008') THEN '식품위생법에 따른 수입신고필함' " & vbcrlf		'식품위생법에 따른 수입신고 여부	20130215진영 추가
		strSQL = strSQL & " 	 WHEN F.chkDiv='N' AND (M.infoCd='17008') THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='M' THEN replace(F.infocontent,'.','') " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='C' AND F.chkDiv='N' THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN replace(F.infocontent,'1644-6030','1644-6035') " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infoCd='02004' and F.infocontent='' then '해당없음' " & vbcrlf
		strSql = strSql & " 	 WHEN LEN(F.infocontent) < 2 THEN '상품 상세 참고' " & vbcrlf
		strSQL = strSQL & " 	 ELSE F.infocontent " & vbcrlf
		strSQL = strSQL & " END AS infoContent, L.shortVal " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		If newInfodiv = "" Then
			strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		Else
			strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv= '"& newInfodiv &"'  " & vbcrlf
		End If
		strSQL = strSQL & " JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		strSQL = strSQL & " JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd AND F.itemid='"&Fitemid&"' " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_OutMall_etcLink as L on L.mallid = M.mallid and L.linkgbn='infoDiv21Lotte' and L.itemid ='"&FItemid&"' " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'lotteimall' AND IC.itemid='"&Fitemid&"' " & vbcrlf
		strSQL = strSQL & " ORDER BY M.infocd ASC"
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		Dim mat_name, mat_percent, mat_place, material

		If Not(rsget.EOF or rsget.BOF) then
			rsgetmallinfoDiv = rsget("mallinfoDiv")
			mallinfoDiv = "&ec_goods_artc_cd="&Server.URLEncode(rsget("mallinfoDiv"))						'상품품목코드
			Do until rsget.EOF
				mallinfoCd = rsget("mallinfoCd")
				infoContent  = rsget("infoContent")
				If isnull(infoContent) Then
					infoContent = ""
				End If
				infoContent  = replace(infoContent,"%", "％")
				infoContent  = replace(infoContent,chr(13), "")
				infoContent  = replace(infoContent,chr(10), "")
				infoContent  = replace(infoContent,chr(9), " ")
				If mallinfoCd="10085" Then
					If isNull(rsget("shortVal")) = FALSE Then
						material = Split(rsget("shortVal"),"!!^^")
						mat_name	= material(0)
						mat_percent	= material(1)
						mat_place	= material(2)

						bufTxt = "&mmtr_nm="&mat_name														'주원료명
						bufTxt = bufTxt&"&cmps_rt="&mat_percent												'함량
						bufTxt = bufTxt&"&mmtr_orpl_nm="&mat_place											'원료원산지
					End If
				End If
				mallinfoCdAll = mallinfoCdAll & "&"&mallinfoCd&"=" &infoContent								'상품품목별 항목정보
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		Dim isTarget, certParams
		certParams = getNewCertParams(rsgetmallinfoDiv)

		'strRst = anjunInfo & mallinfoDiv & mallinfoCdAll & bufTxt		'전안법 적용전
		strRst = certParams & mallinfoDiv & mallinfoCdAll & bufTxt		'전안법 적용
		getLotteiMallItemInfoCdToReg = strRst
	End Function

	Public Function getNewCertParams(imallinfoDiv)
		Dim strSql, certNum, safetyDiv, isTarget
		Dim tgtSeq, sctCd, certNo, isRegCert
		Dim strRst
		strRst = ""
		Select Case imallinfoDiv
			Case "106", "107", "108", "109", "110", "111", "112", "113", "114", "115", "117", "119", "123", "124", "125", "126", "131", "132", "135"
				isTarget = "Y"		'안전인증대상품목 O
			Case Else
				isTarget = "N"		'안전인증대상품목 X
		End Select
		If isTarget = "Y" Then
			strSql = ""
			strSql = strSql & " SELECT TOP 1 isnull(certNum, '') as certNum, isnull(safetyDiv, '') as safetyDiv " & vbcrlf
			strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg " & vbcrlf
			strSql = strSql & " WHERE itemid='"&FItemID&"' " & vbcrlf
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				certNum		= rsget("certNum")
				safetyDiv	= rsget("safetyDiv")
				isRegCert	= "Y"
			Else
				isRegCert 	= "N"
			End If
			rsget.Close

			If isRegCert = "Y" Then
				tgtSeq = 1
				If imallinfoDiv = "119" Then
					tgtSeq = "2"
				End If

				strRst = strRst & "&sft_cert_tgt_seq="&tgtSeq&""
				Select Case safetyDiv
					Case "10", "40", "70"
						strRst = strRst & "&sft_cert_sct_cd=21"
					Case "20", "50", "80"
						strRst = strRst & "&sft_cert_sct_cd=22"
					Case "30", "60", "90"
						strRst = strRst & "&sft_cert_sct_cd=23"
				End Select
				strRst = strRst & "&sft_cert_no=" & certNum
			Else
				Select Case imallinfoDiv
					Case "106", "115", "117", "124", "125", "126", "131", "132", "135"
						tgtSeq = "2"	'안전인증대상 or 비대상 둘 다 가능
					Case "107", "108", "109", "110", "111", "112", "113", "114", "123"
						tgtSeq = "1"	'전상품
					Case Else
						tgtSeq = "0"	'안전인증대상품목 X
				End Select

				If imallinfoDiv = "119" Then
					tgtSeq = "1"
				End If

				strRst = strRst & "&sft_cert_tgt_seq="&tgtSeq&""
				strRst = strRst & "&sft_cert_sct_cd="
				strRst = strRst & "&sft_cert_no="
			End If
		End If
		getNewCertParams = strRst
	End Function

	'// 롯데아이몰 판매여부 반환
	Public Function getLotteiMallSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold > CMAXLIMITSELL) then
				getLotteiMallSellYn = "Y"
			Else
				getLotteiMallSellYn = "N"
			End If
		Else
			getLotteiMallSellYn = "N"
		End If
	End Function

	Public Function getLimitLotteEa()
		Dim ret
		ret = FLimitNo - FLimitSold - 5
		If (ret < 1) Then ret = 0
		getLimitLotteEa = ret
	End Function

	Public Function getOptionLimitNo()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (IsOptionSoldOut) Then
			getOptionLimitNo = 0
		Else
			If (FLimitYn = "Y") Then
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

    Public Function getLotteiMallItemEditParameter2()
		Dim strRst
		strRst = getLotteiMallItemRegParameter(true)
		getLotteiMallItemEditParameter2 = strRst
    End Function

	'// 상품등록 파라메터 생성
	Public Function getLotteiMallItemRegParameter(isEdit)
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo											'(*)사용자인증키
		If (isEdit) Then
		   strRst = strRst & "&goods_req_no="&FLTiMallTmpGoodNo
		End If
		strRst = strRst & "&brnd_no=" & BRAND_CODE											'(*)브랜드코드
		strRst = strRst & "&goods_nm=" & Trim(getItemOptNameFormat)							'(*)전시상품명
		strRst = strRst & getItemKeyword
		strRst = strRst & "&mdl_no="															'모델번호(?)
		strRst = strRst & "&pur_shp_cd=3" 													'(*)매출형태(1.직매입, 4.특정, 3.특정판매)	롯데닷컴은 2(판매분매입)로 설정되어있음..아이몰엔 2가 없는데..그래서 4로 놓긴했는데; ''3일듯: 현아 확인
		strRst = strRst & "&sale_shp_cd=10" 												'(*)판매형태코드(10:정상)
		strRst = strRst & "&sale_prc=" & cLng(GetRaiseValue(FRealSellprice/10)*10)			'(*)판매가
		strRst = strRst & "&mrgn_rt="&CLTIMALLMARGIN 										'(*)마진율(7/1일 시스템 개편하면서 11로 바뀐다함..)
		strRst = strRst & "&tdf_sct_cd="&CHKIIF(FVatInclude="N","2","1")					'(*)과면세코드(1:과세, 2:면세)
		strRst = strRst & getLotteiMallCateParamToReg()										'(*)MD상품군 및 해당 전시카테고리(상품수정에서 카테고리 변경이 안 됨..2013-07-02 전시카테고리 수정API로 수정
		If (FLimitYn="Y") then
		    strRst = strRst & "&inv_mgmt_yn=Y"												'(*)재고관리여부(롯데닷컴처럼 변형) 2013-06-24 김진영
			strRst = strRst & "&inv_qty="&getOptionLimitNo()								'재고수량
		Else
			strRst = strRst & "&inv_mgmt_yn=Y" 												'(*)재고관리여부(롯데닷컴처럼 변형) 2013-06-24 김진영
		    strRst = strRst & "&inv_qty="&CDEFALUT_STOCK									'디폴트 수량 99로
		End If
		strRst = strRst & getLotteiMallOptionParamToReg()									'옵션명 및 옵션상세 :: 단품번호 추가
		strRst = strRst & "&add_choc_tp_cd_10="												'날짜선택형옵션
		If FitemDiv = "06" Then
			strRst = strRst & "&add_choc_tp_cd_20=주문제작상품"						 		'입력형옵션
		End If

		If FitemDiv="06" or FitemDiv="16" then
			strRst = strRst & "&exch_rtgs_sct_cd=10"										'교환/반품여부 10:불가능 / 20:가능
		Else
			strRst = strRst & "&exch_rtgs_sct_cd=20"										'교환/반품여부 10:불가능 / 20:가능
		End If

		strRst = strRst & "&dlv_proc_tp_cd=1" 												'(*)배송유형(1:업체배송, 3:센터배송, 4:센터경유, 6:e-쿠폰배송)
		strRst = strRst & "&gift_pkg_yn=N" 													'(*)선물포장여부
		strRst = strRst & "&dlv_mean_cd=10" 												'(*)배송수단(10:택배 ,11:명절퀵배송 ,40:현장수령 ,50:DHL ,60:해외우편 ,70:일반우편 ,80:등기우편)
		strRst = strRst & getLotteiMallGoodDLVDtParams										'(*)배송상품구분 및 배송기일
		strRst = strRst & "&imps_rgn_info_val="												'배송불가지역(10:서울,수도권, 21:지방, 22:도서지역, 23:인천영종도, 30:제주) 여러개의경우:(콜론)으로 구분하여 전송 한개라도 콜론으로 전송
		strRst = strRst & "&byr_age_lmt_cd="&Chkiif(IsAdultItem() = "Y", "19", "0")&"" 		'(*)구입자나이제한(0:전체, 19:19세이상)
		If Fitemid = "407171" or Fitemid = "788038" or Fitemid = "785541" or Fitemid = "785540" or Fitemid = "785542" or Fitemid = "679670" or Fitemid = "620503" or Fitemid = "590196" or Fitemid = "221081" Then
		strRst = strRst & "&dlv_polc_no=" & tenDlvFreeCd									'(*)배송정책번호(???) tenDlvCd는 inc_dailyAuthCheck.asp에서 정의 (API_TEST에서 따옴)
		Else
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd										'(*)배송정책번호(???) tenDlvCd는 inc_dailyAuthCheck.asp에서 정의 (API_TEST에서 따옴)
		End If
		strRst = strRst & "&corp_dlvp_sn=525713"						 					'(*)반품지(???) (API_TEST에서 따옴)
		strRst = strRst & "&corp_rls_pl_sn=525712"						 					'(*)출고지(???) (API_TEST에서 따옴)
		strRst = strRst & "&orpl_nm=" & chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea)	'(*)원산지
		strRst = strRst & "&mfcp_nm=" & chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)		'(*)제조사
		strRst = strRst & "&impr_nm="						 								'판매자(???)
		strRst = strRst & "&img_url=" & FbasicImage											'(*)대표이미지URL
		strRst = strRst & getLotteiMallAddImageParamToReg()									'부가이미지URL
		strRst = strRst & getLotteiMallItemContParamToReg()									'(*)상세설명
		strRst = strRst & "&md_ntc_2_FCONT="												'MD공지
		strRst = strRst & "&brnd_intro_cont=Design Your Life! 새로운 일상을 만드는 감성생활브랜드 텐바이텐"		'브랜드 설명
'2013-10-10 김진영 수정..주의사항 땜시 상품등록/수정오류 났었음
		ForderComment = replace(ForderComment,"&nbsp;"," ")
		ForderComment = replace(ForderComment,"&nbsp"," ")
		ForderComment = replace(ForderComment,"&"," ")
		ForderComment = replace(ForderComment,chr(13)," ")
		ForderComment = replace(ForderComment,chr(10)," ")
		ForderComment = replace(ForderComment,chr(9)," ")
		strRst = strRst & "&att_mtr_cont=" &URLEncodeUTF8(ForderComment)					'주의사항
		strRst = strRst & "&as_cont="														'AS정보
		strRst = strRst & "&gft_nm="														'사은품명
		strRst = strRst & "&gft_aply_strt_dtime="											'사은품시작일시
		strRst = strRst & "&gft_aply_end_dtime="											'사은품종료일시
		strRst = strRst & "&gft_fcont="														'사은품정보
		strRst = strRst & "&corp_goods_no=" & FNewItemid									'업체상품번호
		strRst = strRst & "&sum_pkg_psb_yn=Y"												'합포장가능여부(자체배송만Y ,N) ==> 우선은 Y로..
	    strRst = strRst & getLotteiMallItemInfoCdToReg()   ''진영
		getLotteiMallItemRegParameter = strRst
	End Function

	Public Function getLotteiMallOptionParamToEdit()
		Dim ret : ret = ""
		ret = ""
	    If (FLimitYn="Y") Then
		    ret = ret & "&inv_mgmt_yn=Y"
		    ret = ret & "&inv_qty="&getOptionLimitNo()
		    ret = ret & "&item_sale_stat_cd=10"
		Else
			ret = ret & "&inv_mgmt_yn=Y"
			ret = ret & "&inv_qty="&CDEFALUT_STOCK
		    ret = ret & "&item_sale_stat_cd=10"
		END IF

		If (getLTiMallSellYn = "Y") and (getOptionLimitNo > 0) Then										 																'판매상태			(*:10:판매,20:품절)
			ret = ret & "&sale_stat_cd=10"
		Else
			FSellyn = "N"
			ret = ret & "&sale_stat_cd=20"
		End If
		getLotteiMallOptionParamToEdit = ret
	End Function

	'// 상품수정 파라메터 생성
	Public Function getLotteiMallItemEditParameter()
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo											'(*)사용자인증키
		strRst = strRst & "&goods_no=" & FLtiMallGoodNo										'(*)롯데아이몰 상품번호
		strRst = strRst & "&brnd_no=" & BRAND_CODE											'(*)브랜드코드
		strRst = strRst & getItemKeyword
		strRst = strRst & "&mdl_no="															'모델번호(?)
		strRst = strRst & "&pur_shp_cd=3" 													'(*)매출형태(1.직매입, 4.특정, 3.특정판매)	롯데닷컴은 2(판매분매입)로 설정되어있음..아이몰엔 2가 없는데..그래서 4로 놓긴했는데; ''3일듯: 현아 확인
		strRst = strRst & "&tdf_sct_cd="&CHKIIF(FVatInclude="N","2","1")					'(*)과면세코드(1:과세, 2:면세)
		strRst = strRst & getLotteiMallCateParamToReg()										'(*)해당 전시카테고리(MD상품군 파라매타도 넘기는 데 괜찮을지 몰겠음..매뉴얼엔 MD상품군 넘기는 파라매타가 없음..진영맘대로)
		strRst = strRst & getLotteiMallOptionParamToEdit()
		strRst = strRst & "&add_choc_tp_cd_10="												'날짜선택형옵션
		If FitemDiv = "06" Then
			strRst = strRst & "&add_choc_tp_cd_20=주문제작상품"						 		'입력형옵션
		End If

		If FitemDiv="06" or FitemDiv="16" then
			strRst = strRst & "&exch_rtgs_sct_cd=10"																					'교환/반품여부 10:불가능 / 20:가능
		Else
			strRst = strRst & "&exch_rtgs_sct_cd=20"																					'교환/반품여부 10:불가능 / 20:가능
		End If
		strRst = strRst & "&dlv_proc_tp_cd=1" 												'(*)배송유형(1:업체배송, 3:센터배송, 4:센터경유, 6:e-쿠폰배송)
		strRst = strRst & "&gift_pkg_yn=N" 													'(*)선물포장여부
		strRst = strRst & "&dlv_mean_cd=10" 												'(*)배송수단(10:택배 ,11:명절퀵배송 ,40:현장수령 ,50:DHL ,60:해외우편 ,70:일반우편 ,80:등기우편)
		strRst = strRst & getLotteiMallGoodDLVDtParams										'(*)배송상품구분 및 배송기일
		strRst = strRst & "&imps_rgn_info_val="													'배송불가지역(10:서울,수도권, 21:지방, 22:도서지역, 23:인천영종도, 30:제주) 여러개의경우:(콜론)으로 구분하여 전송 한개라도 콜론으로 전송
		strRst = strRst & "&byr_age_lmt_cd="&Chkiif(IsAdultItem() = "Y", "19", "0")&"" 		'(*)구입자나이제한(0:전체, 19:19세이상)
		If Fitemid = "407171" or Fitemid = "788038" or Fitemid = "785541" or Fitemid = "785540" or Fitemid = "785542" or Fitemid = "679670" or Fitemid = "620503" or Fitemid = "590196" or Fitemid = "221081" Then
		strRst = strRst & "&dlv_polc_no=" & tenDlvFreeCd									'(*)배송정책번호(???) tenDlvCd는 inc_dailyAuthCheck.asp에서 정의 (API_TEST에서 따옴)
		Else
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd										'(*)배송정책번호(???) tenDlvCd는 inc_dailyAuthCheck.asp에서 정의 (API_TEST에서 따옴)
		End If
		strRst = strRst & "&corp_dlvp_sn=525713"						 					'(*)반품지(???) (API_TEST에서 따옴)
		strRst = strRst & "&corp_rls_pl_sn=525712"						 					'(*)출고지(???) (API_TEST에서 따옴)
		strRst = strRst & "&orpl_nm=" & chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea)	'(*)원산지
		strRst = strRst & "&mfcp_nm=" & chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)		'(*)제조사
		strRst = strRst & "&impr_nm="						 									'판매자(???)
		strRst = strRst & "&img_url=" & FbasicImage											'(*)대표이미지URL
		strRst = strRst & getLotteiMallAddImageParamToReg()									'부가이미지URL
		strRst = strRst & getLotteiMallItemContParamToReg()									'(*)상세설명
		strRst = strRst & "&md_ntc_2_FCONT="													'MD공지
		strRst = strRst & "&brnd_intro_cont=Design Your Life! 새로운 일상을 만드는 감성생활브랜드 텐바이텐"		'브랜드 설명
'2013-10-10 김진영 수정..주의사항 땜시 상품등록/수정오류 났었음
		ForderComment = replace(ForderComment,"&nbsp;"," ")
		ForderComment = replace(ForderComment,"&nbsp"," ")
		ForderComment = replace(ForderComment,"&"," ")
		ForderComment = replace(ForderComment,chr(13)," ")
		ForderComment = replace(ForderComment,chr(10)," ")
		ForderComment = replace(ForderComment,chr(9)," ")
		strRst = strRst & "&attd_mtr_cont=" &URLEncodeUTF8(ForderComment)						'주의사항
		strRst = strRst & "&as_cont="															'AS정보
		strRst = strRst & "&gft_nm="															'사은품명
		strRst = strRst & "&gft_aply_strt_dtime="												'사은품시작일시
		strRst = strRst & "&gft_aply_end_dtime="												'사은품종료일시
		strRst = strRst & "&gft_fcont="															'사은품정보
		strRst = strRst & "&corp_goods_no=" & FNewItemid										'업체상품번호
		strRst = strRst & "&sum_pkg_psb_yn=Y"												'합포장가능여부(자체배송만Y ,N) ==> 우선은 Y로..
	    strRst = strRst & getLotteiMallItemInfoCdToReg()   ''진영
		'결과 반환
		getLotteiMallItemEditParameter = strRst
	End Function


End Class

Class CLotteiMall
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
	Public FRectIdx

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

	Public Sub getLtimallNotRegOneItem
		Dim strSql, addSql, i
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.*, o.itemoption, o.isusing as optisusing, o.optsellyn, o.optlimitno, o.optlimitsold, o.optaddprice, o.optionname  "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, '"&CitemGbnKey&"' as itemGbnKey"
		strSql = strSql & "	, isNULL(R.LtiMallStatCD,-9) as LtiMallStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(M.newitemname, '') as newitemname, isnull(M.itemnameChange, '') as itemnameChange "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
		strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M on M.itemid = o.itemid and M.itemoption = o.itemoption and M.mallid = 'lotteimall' and M.idx = '"&FRectIdx&"' "
		strSql = strSql & " JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotteimall_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " JOIN db_user.dbo.tbl_user_c UC on i.makerid = UC.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ltimallAddOption_regItem as R on R.midx = M.idx "
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000)))"
		ELSE
		    strSql = strSql & "	and (i.deliveryType <> 9)"
	    END IF
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.itemdiv not in ('21', '30') "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "					'플라워/화물배송/해외직구 상품 제외
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & "	and UC.isExtUsing <> 'N'"
		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'등록제외 카테고리
		strSql = strSql & "	and M.idx not in (Select midx From db_etcmall.dbo.tbl_ltimallAddOption_regItem where LtiMallStatCD > 3 ) "			''LtiMallStatCD>=3 등록완료이상은 등록안됨.										'롯데등록상품 제외
		strSql = strSql & " and o.optsellyn = 'Y' "
		strSql = strSql & " and (o.optlimityn = 'N' or ((o.optlimityn = 'Y') and (o.optlimitno - o.optlimitsold >="&CMAXLIMITSELL&"))) "
		strSql = strSql & " and isNULL(c.infodiv,'') not in ('','18','20','22')"
		strSql = strSql & addSql																				'카테고리 매칭 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CLotteiMallItem
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
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
                FOneItem.FitemGbnKey        = rsget("itemGbnKey")
                FOneItem.FLtiMallStatCD     = rsget("LtiMallStatCD")

                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.Fsafetyyn			= rsget("safetyyn")
                FOneItem.FsafetyDiv			= rsget("safetyDiv")
                FOneItem.FsafetyNum			= rsget("safetyNum")

                FOneItem.FNewitemname		= rsget("newitemname")
                FOneItem.FItemnameChange	= rsget("itemnameChange")
                FOneItem.FItemoption		= rsget("itemoption")
                FOneItem.FOptisusing		= rsget("optisusing")
                FOneItem.FOptsellyn			= rsget("optsellyn")
                FOneItem.FOptlimitno		= rsget("optlimitno")
                FOneItem.FOptlimitsold		= rsget("optlimitsold")
                FOneItem.FOptaddprice		= rsget("optaddprice")
                FOneItem.FRealSellprice		= rsget("sellcash") + rsget("optaddprice")
                FOneItem.FNewItemid			= CStr(rsget("itemid")) & CStr(rsget("itemoption"))
                FOneItem.FOptionname		= rsget("optionname")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub

	Public Sub getLtimallEditOneItem
		Dim strSql, addSql, i
        ''//연동 제외상품
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt < getdate()"
        addSql = addSql & "     and edDt > getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.*, o.itemoption, o.isusing as optisusing, o.optsellyn, o.optlimitno, o.optlimitsold, o.optaddprice, o.optionname "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, R.LtiMallGoodNo, R.LtiMallTmpGoodNo, R.LtiMallprice, R.LtiMallSellYn "
		strSql = strSql & "	, R.accFailCNT, R.lastErrStr "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(M.newitemname, '') as newitemname, isnull(M.itemnameChange, '') as itemnameChange "
		strSql = strSql & "	, M.optionname as regedOptionname, M.itemname as regedItemname  "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N' or o.itemoption is null "
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '30') "
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or isNULL(c.infodiv,'') in ('','18','20','22')"  ''화장품, 식품류 제외
		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
		strSql = strSql & "		or (('1'+c.infoDiv in ('107', '108', '109', '110', '111', '112', '113', '114', '123')) and isnull(TR.certNum, '') = '') "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut"
		strSql = strSql & "	FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] as m "
		strSql = strSql & "	JOIN db_item.dbo.tbl_item as i on m.itemid = i.itemid "
		strSql = strSql & "	JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid  "
		strSql = strSql & "	JOIN db_etcmall.dbo.tbl_ltimallAddOption_regItem as R on R.midx = M.idx  "
		strSql = strSql & " LEFT JOIN ("
		strSql = strSql & " 	SELECT TOP 1 itemid, isnull(certNum, '') as certNum"
		strSql = strSql & " 	FROM db_item.dbo.tbl_safetycert_tenReg "
		strSql = strSql & " 	WHERE isnull(certNum, '') <> ''"
		strSql = strSql & " ) as TR on i.itemid = TR.itemid"
		strSql = strSql & "	LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid  "
		strSql = strSql & "	LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotteimall_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small  "
		strSql = strSql & "	LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid and M.itemid = o.itemid and M.itemoption = o.itemoption  "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and M.mallid = 'lotteimall' and M.idx = '"&FRectIdx&"' "
		strSql = strSql & addSql
		strSql = strSql & " and isNULL(R.LtiMallTmpGoodNo, R.LtiMallGoodNo) is Not Null "									'#등록 상품만

'		strSql = ""
'		strSql = strSql & " SELECT TOP " & FPageSize & " i.*, o.itemoption, o.isusing as optisusing, o.optsellyn, o.optlimitno, o.optlimitsold, o.optaddprice, o.optionname "
'		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
'		strSql = strSql & "	, R.LtiMallGoodNo, R.LtiMallTmpGoodNo, R.LtiMallprice, R.LtiMallSellYn "
'		strSql = strSql & "	, R.accFailCNT, R.lastErrStr "
'		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
'		strSql = strSql & "	, isnull(M.newitemname, '') as newitemname, isnull(M.itemnameChange, '') as itemnameChange "
'		strSql = strSql & "	, M.optionname as regedOptionname, M.itemname as regedItemname  "
'        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
'		strSql = strSql & "		or i.isExtUsing='N'"
'		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
'		strSql = strSql & "		or i.sellyn<>'Y'"
'		strSql = strSql & "		or i.deliveryType = 7"
'		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
'		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
'		strSql = strSql & "		or isNULL(c.infodiv,'') in ('','18','20','22')"  ''화장품, 식품류 제외
'		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
'		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
'		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
'		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut"
'		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
'		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
'		strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
'		strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M on M.itemid = o.itemid and M.itemoption = o.itemoption and M.mallid = 'lotteimall' and M.idx = '"&FRectIdx&"' "
'		strSql = strSql & " JOIN db_etcmall.dbo.tbl_ltimallAddOption_regItem as R on R.midx = M.idx "
'		strSql = strSql & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotteimall_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
'		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
'		strSql = strSql & " WHERE 1 = 1"
'		strSql = strSql & addSql
'		strSql = strSql & " and isNULL(R.LtiMallTmpGoodNo, R.LtiMallGoodNo) is Not Null "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CLotteiMallItem
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
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FLTiMallGoodNo		= rsget("LtiMallGoodNo")
				FOneItem.FLtiMallTmpGoodNo	= rsget("LtiMallTmpGoodNo")
				FOneItem.FLtiMallPrice		= rsget("LtiMallprice")
				FOneItem.FLtiMallSellYn		= rsget("LtiMallSellYn")
				FOneItem.Fvatinclude        = rsget("vatinclude")
                FOneItem.FoptionCnt         = rsget("optionCnt")
                FOneItem.FaccFailCNT        = rsget("accFailCNT")
                FOneItem.FlastErrStr        = rsget("lastErrStr")
                ''FOneItem.Fcorp_dlvp_sn      = rsget("returnCode")
                FOneItem.Fdeliverytype      = rsget("deliverytype")
                FOneItem.FrequireMakeDay    = rsget("requireMakeDay")

                FOneItem.FinfoDiv       = rsget("infoDiv")
                FOneItem.Fsafetyyn      = rsget("safetyyn")
                FOneItem.FsafetyDiv     = rsget("safetyDiv")
                FOneItem.FsafetyNum     = rsget("safetyNum")
                FOneItem.FmaySoldOut    = rsget("maySoldOut")

                FOneItem.FNewitemname		= rsget("newitemname")
                FOneItem.FItemnameChange	= rsget("itemnameChange")
                FOneItem.FItemoption		= rsget("itemoption")
                FOneItem.FOptisusing		= rsget("optisusing")
                FOneItem.FOptsellyn			= rsget("optsellyn")
                FOneItem.FOptlimitno		= rsget("optlimitno")
                FOneItem.FOptlimitsold		= rsget("optlimitsold")
                FOneItem.FOptaddprice		= rsget("optaddprice")
                FOneItem.FRealSellprice		= rsget("sellcash") + rsget("optaddprice")
                If not isnull(rsget("itemoption")) Then
                	FOneItem.FNewItemid			= CStr(rsget("itemid")) & CStr(rsget("itemoption"))
            	End If
                FOneItem.FOptionname		= rsget("optionname")
	            FOneItem.FRegedOptionname	= rsget("regedOptionname")
	            FOneItem.FRegedItemname		= rsget("regedItemname")
				FOneItem.FAdultType 		= rsget("adulttype")
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

Function getLtimallGoodno(iidx)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 LtimallGoodNo FROM db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] WHERE midx = '"&iidx&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		getLtimallGoodno = rsget("ltimallGoodno")
	rsget.Close
End Function
%>