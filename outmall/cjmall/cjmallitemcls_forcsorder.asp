<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "cjmall"
CONST CMAXLIMITSELL = 5        '' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CCJMALLMARGIN = 12       ''마진 12%...// 왜 12? // 2013-11-05 김진영..12->15로 수정 =>12로 수정 유미희.(2013/11/21)
CONST CitemGbnKey ="K1099999" ''상품구분키 ''하나로 통일
CONST CUPJODLVVALID = True   ''업체 조건배송 등록 가능여부

CONST CVENDORID = 411378					'협력업체코드
CONST CVENDORCERTKEY = "CJ03074113780"		'인증키
CONST CUNIQBRANDCD = 24049000				'브랜드코드
CONST MD_CODE = "5103"						'MD_Code

Class cjmallItem
	Public Fitemid
	Public Fitemname
	Public FsmallImage
	Public Fmakerid
	Public Fregdate
	Public FlastUpdate
	Public ForgPrice
	Public Fsellcash
	Public Fbuycash
	Public FsellYn
	Public Fsaleyn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public Fdeliverytype
	Public FoptionCnt
	Public FcjmallRegdate
	Public FcjmallLastUpdate
	Public FcjmallPrdNo
	Public FcjmallPrice
	Public FcjmallSellYn
	Public FregUserid
	Public FcjmallStatCd
	Public FregedOptCnt
	Public FrctSellCNT
	Public FaccFailCNT
	Public FlastErrStr
	Public FCateMapCnt
	Public FcdmKey
	Public FdefaultfreeBeasongLimit
	'카테고리
	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FDispNo
	Public FDispNm
	Public FDispLrgNm
	Public FDispMidNm
	Public FDispSmlNm
	Public FDispThnNm
	Public FCateIsUsing
	Public Fdisptpcd
	Public FisUsing

	'상품분류
	Public Finfodiv
	Public Ficnt
	Public FCddKey

	Public Fcdd_Name
	Public Fcdl_Name
	Public Fcdm_Name
	Public Fcds_Name
	Public FPrdDivIsUsing

	Public FRectMode
	Public FRectItemID

	Public FCdm
	Public FCdd

	'상품등록 매칭
	Public FitemDiv
	Public ForgSuplyCash
	Public Fkeywords
	Public Fvatinclude
	Public ForderComment
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FitemGbnKey
	Public Fdeliverfixday

	Public FItemOption
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold
	Public Fsocname_kor
	Public FmaySoldOut

	Public MustPrice

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

	'// 품절여부
	Public Function IsSoldOut()
		ISsoldOut = (FSellyn <> "Y") or ((FLimitYn = "Y") and (FLimitNo - FLimitSold < 1))
	End Function

    public Function IsCjFreeBeasong()
        IsCjFreeBeasong = False
    end Function

	function getLimitEa()
		Dim ret : ret = (FLimitno - FLimitSold)
		If (ret < 1) Then ret = 0
		getLimitEa = ret
	end function

	Function getLimitHtmlStr()
		If IsNULL(FLimityn) Then Exit Function

		If (FLimityn="Y") Then
			getLimitHtmlStr = "<font color=blue>한정:"&getLimitEa&"</font>"
		End If
	End Function

	Public Function getcjmallStatName
	    If IsNULL(FcjmallStatCd) then FcjmallStatCd=-1
		Select Case FcjmallStatCd
			CASE -9 : getcjmallStatName = "미등록"
			CASE -2 : getcjmallStatName = "<font color=red>반려</font>"
			CASE -1 : getcjmallStatName = "등록실패"
			CASE 0 : getcjmallStatName = "<font color=blue>등록예정</font>"
			CASE 1 : getcjmallStatName = "전송시도"
			CASE 3 : getcjmallStatName = "승인대기"
			CASE 7 : getcjmallStatName = ""
			CASE ELSE : getcjmallStatName = FcjmallStatCd
		End Select
	End Function

	Function getDispGubunNm()
		getDispGubunNm = getDisptpcdName
	End Function

	Public Function getDisptpcdName
		If (Fdisptpcd="B") Then
			getDisptpcdName = "<font color='blue'>전문</font>"
		Elseif (Fdisptpcd = "D") Then
			getDisptpcdName = "일반"
		Else
			getDisptpcdName = Fdisptpcd
		End if
	End Function

	'화물배송 관련
	Public Function getdeliverfixday()
		If (Fdeliverfixday = "C") or (Fdeliverfixday = "X") Then
			getdeliverfixday = 20
		Else
			getdeliverfixday = 10
		End If
	End Function

	'//cjmall 등록상태 반환
	Public Function getcjItemStatCd()
	    getcjItemStatCd = getcjmallStatName
	End Function

    Function getCJmallSuplyPrice(optaddprice)
'        getCJmallSuplyPrice = CLNG(FSellCash * (100-CCJMALLMARGIN) / 100)
		'하단은 CJ메뉴얼에 적힌 내용
		'* 마진율 확인요함
		'1. 과세상품 : 매입원가(VAT제외) = Round(판매가/1.1 - 0.1 * (판매가/1.1)), 0)
		'2. 면세상품 : 매입원가(VAT제외) = Round(판매가 - 0.1 * 판매가, 0)
		If FVatInclude = "Y" Then		'과세
			getCJmallSuplyPrice = Round((MustPrice+optaddprice) /1.1 - (CCJMALLMARGIN/100) * ((MustPrice+optaddprice)/1.1))
		Else							'면세
			getCJmallSuplyPrice = Round((MustPrice+optaddprice) - (CCJMALLMARGIN/100) * (MustPrice+optaddprice))
		End If
    End Function

    Function getCJmallSuplyPrice2()
'        getCJmallSuplyPrice2 = CLNG(FSellCash * (100-CCJMALLMARGIN) / 100)
		'하단은 CJ메뉴얼에 적힌 내용
		'* 마진율 확인요함
		'1. 과세상품 : 매입원가(VAT제외) = Round(판매가/1.1 - 0.1 * (판매가/1.1)), 0)
		'2. 면세상품 : 매입원가(VAT제외) = Round(판매가 - 0.1 * 판매가, 0)
		If FVatInclude = "Y" Then		'과세
			getCJmallSuplyPrice2 = Round((MustPrice) /1.1 - (CCJMALLMARGIN/100) * ((MustPrice)/1.1))
		Else							'면세
			getCJmallSuplyPrice2 = Round((MustPrice) - (CCJMALLMARGIN/100) * (MustPrice))
		End If
    End Function

    public function getDeliverytypeName
        if (Fdeliverytype="9") then
            getDeliverytypeName = "<font color='blue'>[조건 "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
        elseif (Fdeliverytype="7") then
            getDeliverytypeName = "<font color='red'>[업체착불]</font>"
        elseif (Fdeliverytype="2") then
            getDeliverytypeName = "<font color='blue'>[업체]</font>"
        else
            getDeliverytypeName = ""
        end if
    end function

	public function GetCJLmtQty()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (Flimityn="Y") then
			If (Flimitno - Flimitsold) < CLIMIT_SOLDOUT_NO Then
				GetCJLmtQty = 0
			Else
				GetCJLmtQty = Flimitno - Flimitsold - CLIMIT_SOLDOUT_NO
			End If
		Else
			GetCJLmtQty = 999
		End If
	End Function

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

    ''주문제작 여부
    Public Function getzCostomMadeInd()
		dim ret, CMadeInd
        ret = (Fitemdiv="06" or Fitemdiv="16")
        ret = ret or (FtenCateLarge="010" and FtenCateMid="070" and FtenCateSmall="070")	'디자인문구	스탬프	주문제작
		ret = ret or (FtenCateLarge="035" and FtenCateMid="016" and FtenCateSmall="010")	'여행/취미	드라이브	주차판
		ret = ret or (FtenCateLarge="040")													'가구
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="001")	'수납/생활	보관/정리용품	수납장
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="002")	'수납/생활	보관/정리용품	틈새수납장
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="005")	'수납/생활	보관/정리용품	잡지꽂이
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="010")	'수납/생활	보관/정리용품	벽걸이수납함
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="010")	'수납/생활	보관/정리용품	벽걸이수납함
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="019")	'수납/생활	보관/정리용품	이동식수납장
		ret = ret or (FtenCateLarge="045" and FtenCateMid="003")							'수납/생활	데스크수납
		ret = ret or (FtenCateLarge="045" and FtenCateMid="006")							'수납/생활	데코수납
		ret = ret or (FtenCateLarge="045" and FtenCateMid="007" and FtenCateSmall="008")	'수납/생활	키즈수납	키즈 서랍장
		ret = ret or (FtenCateLarge="050" and FtenCateMid="010" and FtenCateSmall="050")	'홈/데코	조명	이니셜/메세지조명
		ret = ret or (FtenCateLarge="050" and FtenCateMid="030" and FtenCateSmall="010")	'홈/데코	장식소품	이니셜장식
		ret = ret or (FtenCateLarge="050" and FtenCateMid="045" and FtenCateSmall="120")	'홈/데코	홈갤러리	수작업 주문제작
		ret = ret or (FtenCateLarge="055" and FtenCateMid="070")							'패브릭 > 침구세트
		ret = ret or (FtenCateLarge="055" and FtenCateMid="080")							'패브릭 > 커튼
		ret = ret or (FtenCateLarge="055" and FtenCateMid="090")							'패브릭 > 쿠션/방석
		ret = ret or (FtenCateLarge="055" and FtenCateMid="100")							'패브릭 > 매트/러그
		ret = ret or (FtenCateLarge="055" and FtenCateMid="110")							'패브릭 > 패브릭소품
		ret = ret or (FtenCateLarge="055" and FtenCateMid="120")							'패브릭 > 침구단품
		ret = ret or (FtenCateLarge="060" and FtenCateMid="130")							'키친 > 작가 생활자기
		ret = ret or (FtenCateLarge="070" and FtenCateMid="160")							'가방/슈즈/쥬얼리 > 쥬얼리
		ret = ret or (FtenCateLarge="090" and FtenCateMid="070" and FtenCateSmall="010")	'Men > 쥬얼리/잡화 > 시계/쥬얼리
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="020")	'베이비 > 가구/침구/수납 > 데코스티커/벽지
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="040")	'베이비 > 가구/침구/수납 > 수납함/책꽂이
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="050")	'베이비 > 가구/침구/수납 > 의자
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="060")	'베이비 > 가구/침구/수납 > 조명/액자
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="066")	'베이비 > 가구/침구/수납 > 테이블/책상
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="070")	'베이비 > 가구/침구/수납 > 안전용품
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="100")	'베이비 > 가구/침구/수납 > 아기침대
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="110")	'베이비 > 가구/침구/수납 > 플레이매트
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="120")	'베이비 > 가구/침구/수납 > 블랑켓/아기담요
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="130")	'베이비 > 가구/침구/수납 > 모빌
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="140")	'베이비 > 가구/침구/수납 > 쿠션/침구/커튼
		If ret Then
			CMadeInd = "Y"
		Else
			CMadeInd = "N"
		End If
        getzCostomMadeInd = CMadeInd
    End Function

    ''리드타임 얻기
    Public Function getzLeadTime()
		If (FtenCateLarge="040") or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="001") or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="002")	or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="005") or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="010")	or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="019") or (FtenCateLarge="045" and FtenCateMid="003")	or (FtenCateLarge="045" and FtenCateMid="006") or (FtenCateLarge="045" and FtenCateMid="007" and FtenCateSmall="008")	or (FtenCateLarge="055" and FtenCateMid="070") or (FtenCateLarge="055" and FtenCateMid="080")	or (FtenCateLarge="055" and FtenCateMid="090") or (FtenCateLarge="055" and FtenCateMid="100")	or (FtenCateLarge="055" and FtenCateMid="110") or (FtenCateLarge="055" and FtenCateMid="120")	or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="040") or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="050")	or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="066") or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="100")	or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="120") or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="140") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="020") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="060") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="070") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="110") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="130") OR (FtenCateLarge = "050" and FtenCateMid = "120" and FtenCateSmall = "080") OR (FtenCateLarge = "050" and FtenCateMid = "045" and FtenCateSmall = "100") OR (FtenCateLarge = "070" and FtenCateMid = "070") OR (FtenCateLarge = "070" and FtenCateMid = "160") Then
			getzLeadTime = "15"
		ElseIf (FtenCateLarge = "010" and FtenCateMid = "070" and FtenCateSmall = "070") OR (FtenCateLarge="035" and FtenCateMid="016" and FtenCateSmall="010") OR (FtenCateLarge="050" and FtenCateMid="010" and FtenCateSmall="050") OR (FtenCateLarge="050" and FtenCateMid="030" and FtenCateSmall="010") OR (FtenCateLarge="050" and FtenCateMid="045" and FtenCateSmall="120") OR (FtenCateLarge="060" and FtenCateMid="130") OR (FtenCateLarge="070" and FtenCateMid="160") OR (FtenCateLarge="090" and FtenCateMid="070" and FtenCateSmall="010") Then
			getzLeadTime = "03"
		End If
	End Function

	'// 상품등록: 옵션 파라메터 생성(상품등록용)
	Public Function getCJOptionParamToReg()
		Dim strSql, strRst, itemSu, itemoption, validSellno, optionname, fixday, optaddprice
		Dim GetTenTenMargin, i
		'2013-07-24 김진영//텐텐마진이 CJMALL의 마진보다 작을 때 orgprice로 전송 시작
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		'2013-07-24 김진영//텐텐마진이 CJMALL의 마진보다 작을 때 orgprice로 전송 끝

		optaddprice		= 0
		strSql = ""
		strSql = strSql & " SELECT top 900 i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(40),o.optionname) as optionname" & VBCRLF
		strSql = strSql & " , o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, i.deliverfixday, o.optaddprice " & VBCRLF
		strSql = strSql & " ,DATALENGTH(o.optionname) as optnmLen" & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_item as i " & VBCRLF
		strSql = strSql & " LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF
		strSql = strSql & " WHERE i.itemid = "&Fitemid
		strSql = strSql & " ORDER BY o.optaddprice ASC, o.itemoption ASC "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				If rsget.RecordCount = 1 AND IsNull(rsget("itemoption")) Then  ''단일상품
					FItemOption = "0000"
					optionname = DdotFormat(chrbyte(getItemNameFormat,40,""),20)
					itemSu = GetCJLmtQty
					optaddprice		= 0
				Else
					FItemOption 	= rsget("itemoption")
					optionname 		= rsget("optionname")
					Foptsellyn 		= rsget("optsellyn")
					Foptlimityn 	= rsget("optlimityn")
					Foptlimitno 	= rsget("optlimitno")
					Foptlimitsold 	= rsget("optlimitsold")
					optaddprice		= rsget("optaddprice")
					itemSu = getOptionLimitNo

					if rsget("optnmLen")>40 then
					    optionname=DdotFormat(optionname,20)
					end if
				End If

				If rsget("deliverfixday") = "C" OR rsget("deliverfixday") = "X" Then
					fixday = "60"
				Else
					fixday = "20"
				End If
				strRst = strRst &"	<tns:unit>"
				''strRst = strRst &"		<tns:unitNm><![CDATA["&DDotFormat(optionname, 16)&"]]></tns:unitNm>"	'단품정보 - 단품상세(옵션명을 텍스트로 넘기면 됨)
				strRst = strRst &"		<tns:unitNm><![CDATA["&optionname&"]]></tns:unitNm>"
				strRst = strRst &"		<tns:unitRetail>"&FSellCash+optaddprice&"</tns:unitRetail>"				'단품정보 - 판매가
				strRst = strRst &"		<tns:unitCost>"&getCJmallSuplyPrice(optaddprice)&"</tns:unitCost>"					'단품정보 - 매입원가
				strRst = strRst &"		<tns:availableQty>"&itemSu&"</tns:availableQty>"						'단품정보 - 공급가능수량 (상품 재고 파악이 안되는경우는 999같은 숫자를 넣습니다.)
			If getzCostomMadeInd = "Y" Then
				strRst = strRst &"		<tns:leadTime>"&getzLeadTime()&"</tns:leadTime>"						'단품정보 - 리드타임 (* 출고리드타임 예외등록 해야 함, 리드타임 기준 협의요함 00 : 오늘배송, 01 : 당일출고, 02 : 익일출고, 03 : 2일후출고, 04 : 4일, 05 : 5일, 06 : 6일.....)
'			ElseIf Left(FCddkey,2) = "35" OR Left(FCddkey,2) = "37" Then											'상품등록시 대분류값(35 전기전자/37 정보통신)일경우 리드타임의 값은 '02' 등록만 가능하도록 처리되어있습니다.
'				strRst = strRst &"		<tns:leadTime>02</tns:leadTime>"										'단품정보 - 리드타임 (* 출고리드타임 예외등록 해야 함, 리드타임 기준 협의요함 00 : 오늘배송, 01 : 당일출고, 02 : 익일출고, 03 : 2일후출고, 04 : 4일, 05 : 5일, 06 : 6일.....)
			Else
				strRst = strRst &"		<tns:leadTime>03</tns:leadTime>"										'단품정보 - 리드타임 (* 출고리드타임 예외등록 해야 함, 리드타임 기준 협의요함 00 : 오늘배송, 01 : 당일출고, 02 : 익일출고, 03 : 2일후출고, 04 : 4일, 05 : 5일, 06 : 6일.....)
			End If
				strRst = strRst &"		<tns:unitApplyRsn>"&fixday&"</tns:unitApplyRsn>"						'단품정보 - 적용사유 (10 : 적용안함, 20 : 상품포장, 30 : 상품생산, 40 : 입고검사, 50 : 출고검사, 60 : 설치상품)
				strRst = strRst &"		<tns:startSaleDt>"&FormatDate(now(), "0000-00-00")&"</tns:startSaleDt>"	'단품정보 - 판매시작일자
				strRst = strRst &"		<tns:endSaleDt>9999-12-30</tns:endSaleDt>"								'단품정보 - 판매종료일자 (판매상태수정에서..)
			If (application("Svr_Info")="Dev") OR (Fitemid="899506") Then
				strRst = strRst &"		<tns:vpn>"&rsget("itemid")&"_Q"&FItemOption&"</tns:vpn>"				'단품정보 - 협력사상품코드(899506만 Q라는 문자삽입)
			Else
				strRst = strRst &"		<tns:vpn>"&rsget("itemid")&"_"&FItemOption&"</tns:vpn>"					'단품정보 - 협력사상품코드
			End If
				strRst = strRst &"	</tns:unit>"
				rsget.MoveNext
			Next
		End If
		rsget.Close
		getCJOptionParamToReg = strRst
	End Function

	'// 상품등록: MD상품군 및 전시 카테고리 파라메터 생성(상품등록용)
	Public Function getCjCateParamToReg()
		Dim strSql, strRst, i
		strSql = ""
		strSql = strSql & " SELECT top 100 c.CateKey "
		strSql = strSql & " FROM db_outmall.dbo.tbl_cjmall_cate_mapping as m "
		strSql = strSql & " JOIN db_outmall.dbo.tbl_cjMall_Category as c on m.CateKey = c.CateKey "
		strSql = strSql & " WHERE tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " and tenCateSmall='" & FtenCateSmall & "' "
		strSql = strSql & " ORDER BY c.cateGbn ASC " ''B : 브랜드 / D : 일반
		rsCTget.Open strSql,dbCTget,1
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			strRst = ""
			i = 0
			Do until rsCTget.EOF
				If i = 0 Then
					strRst = strRst &"		<tns:mallCtg>"
					strRst = strRst &"			<tns:mainInd>Y</tns:mainInd>"
					strRst = strRst &"			<tns:ctgName>" & rsCTget("CateKey") & "</tns:ctgName>"
					strRst = strRst &"		</tns:mallCtg>"
				Else
					strRst = strRst &"		<tns:mallCtg>"
					strRst = strRst &"			<tns:ctgName>" & rsCTget("CateKey") & "</tns:ctgName>"
					strRst = strRst &"		</tns:mallCtg>"
				End If
				rsCTget.MoveNext
				i = i + 1
			Loop
		End If
		rsCTget.Close
		getCjCateParamToReg = strRst
	End Function

	'상품품목정보
    public function getCjmallItemInfoCdToReg()
		Dim strSql, buf, addSql
		Dim mallinfoCd,infoContent,infotype, infocd, mallinfodiv
		Dim chkInfodiv, chkCdmKey

		strSql = ""
		strSql = strSql & " SELECT top 1 PD.infodiv, PD.cdmKey " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_item as i  " & vbcrlf
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid  " & vbcrlf
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small and c.infodiv = PD.infodiv " & vbcrlf
		strSql = strSql & " WHERE i.itemid ='"&FItemID&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			chkInfodiv	= rsget("infodiv")
			chkCdmKey	= rsget("cdmKey")
		End If
		rsget.Close

		If chkInfodiv = "01" and chkCdmKey = "1006" Then
			addSql = " and M.infocd <> '00000'  "
		End If

		strSql = ""
		strSql = strSql & " SELECT top 100 M.* , " & vbcrlf
		strSql = strSql & "		CASE " & vbcrlf
		'strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND left(isNULL(IC.safetyNum,''),3) = 'KCC' AND (IC.infoDiv not in ('06','23')) THEN 'Y' " & vbcrlf
		'strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND left(isNULL(IC.safetyNum,''),3) <> 'KCC' AND (IC.infoDiv not in ('06','23')) THEN 'N' " & vbcrlf
		'strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (IC.infoDiv in ('06','23')) THEN 'Y' " & vbcrlf
		'strSql = strSql & "			WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN 'N' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN 'Y' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN 'N' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='J' and F.chkDiv='Y' THEN 'Y' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='J' and F.chkDiv='N' THEN 'N' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='P' THEN 'I' " & vbcrlf
		strSql = strSql & "		ELSE 'I' " & vbcrlf
		strSql = strSql & " END AS infoType, " & vbcrlf
		strSql = strSql & "		CASE " & vbcrlf
'		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND left(isNULL(IC.safetyNum,''),3) = 'KCC' AND (IC.infoDiv not in ('06','23')) THEN IC.safetyNum " & vbcrlf
'		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND left(isNULL(IC.safetyNum,''),3) <> 'KCC' AND (IC.infoDiv not in ('06','23')) THEN '해당없음' " & vbcrlf
'		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (IC.infoDiv in ('06','23')) THEN IC.safetyNum " & vbcrlf
'		strSql = strSql & "			WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '해당없음' " & vbcrlf
        strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN IC.safetyNum " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '해당없음' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00001') THEN '해당없음' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00002') AND (M.mallinfoCd='25044') THEN '해당없음' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00003') THEN '상세내역참고' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='J' and F.chkDiv='N' THEN '해당없음' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='P' AND c.infoCd <> '22009' THEN '텐바이텐 고객행복센터 1644-6035' " & vbcrlf
		'strSql = strSql & "		WHEN c.infotype='P' THEN replace(F.infocontent,'1644-6030','1644-6035') " & vbcrlf
		strSql = strSql & "		ELSE convert(varchar(500),F.infocontent) " & vbcrlf
		strSql = strSql & " END AS infocontent " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"' " & vbcrlf
		strSql = strSql & " WHERE M.mallid = 'cjmall' and IC.itemid='"&FItemID&"' " & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infotype	= rsget("infotype")
			    infoContent = rsget("infoContent")
				infocd		= rsget("infocd")
				mallinfodiv = rsget("mallinfodiv")

                if (mallinfodiv="02") and (mallinfoCd="25012") and (infoContent="") then  '' 구두/굽높이
                    infoContent="해당없음"
                end if

                If (FItemID = "674455" OR FItemID = "881879")  AND (mallinfoCd = "25008" OR mallinfoCd = "25013") Then	'2013-06-25 김진영 수정(업체가 내용을 -(하이픈)으로 등록한 경우..이런 경우가 생길때마다 분기줘야 될 것 같음
                	infoContent = "상세참조"
                End If

				buf = buf &"	<tns:goodsReport>"
				buf = buf &"		<tns:pedfId>"&mallinfoCd&"</tns:pedfId>"
				buf = buf &"		<tns:html><![CDATA["&infoContent&"]]></tns:html>"
				buf = buf &"	</tns:goodsReport>"
				rsget.MoveNext
			Loop
		End If
		rsget.Close

'2014-06-09 김진영 하단 주석 제거 / db_outmall.dbo.tbl_OutMall_infoCodeMap에 하단 코드(25066) 삽입완료
'		if chkInfodiv = "19" and chkCdmKey = "8504" Then  ''보석/장신구 : 시계 인경우 25066 필요
'            buf = buf &"	<tns:goodsReport>"
'			buf = buf &"		<tns:pedfId>25066</tns:pedfId>"
'			buf = buf &"		<tns:html><![CDATA[상세참조]]></tns:html>"
'			buf = buf &"	</tns:goodsReport>"
'        end if

		getCjmallItemInfoCdToReg = buf
	End Function

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	Public Function getCJItemContParamToReg()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
		'2014-01-17 10:00 김진영 탑 이미지 추가
		strRst = strRst & ("<p><a href=""http://10x10.cjmall.com/ctg/specialshop_brand/main.jsp?ctg_id=292240"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_cjmall.jpg""></a></p><br>")
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
		getCJItemContParamToReg = strRst
		''2013-06-10 김진영 추가(롯데닷컴처럼 상품이미지가 길면 엑박나오는 현상)
		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','cjmall') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF  '' mallid='cjmall' => mallid in ('','cjmall')
		rsCTget.Open strSQL, dbCTget
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			strRst = rsCTget("textVal")
			strRst = "<div align=""center""><p><a href=""http://10x10.cjmall.com/ctg/specialshop_brand/main.jsp?ctg_id=292240"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_cjmall.jpg""></a></p><br>" & strRst & "<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_common.jpg""></div>"
			getCJItemContParamToReg = strRst
		End If
		rsCTget.Close
	End Function

	'// 상품등록: 상품추가이미지 파라메터 생성(상품등록용)
	Public Function getCJAddImageParamToReg()
		Dim strRst, strSQL, i
		strRst = ""
		If application("Svr_Info")="Dev" Then
			FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
		End If

		strRst = strRst &"	<tns:image>"
		strRst = strRst &"		<tns:imageMain>"&FbasicImage&"</tns:imageMain>"
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					strRst = strRst &"		<tns:imageSub"&i&">http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") &"</tns:imageSub"&i&">"
				End If
				rsget.MoveNext
				If i >= 5 Then Exit For
			Next
		End If
		rsget.Close
		strRst = strRst &"	</tns:image>"
		getCJAddImageParamToReg = strRst
	End Function

	'// 상품상태수정시 옵션이 추가된 경우
	Public Function getCJOptionParamToEdit()
		Dim strSql, strRst, itemSu, itemoption, validSellno, optionname, fixday, optaddprice
		Dim GetTenTenMargin, i
		'2013-07-24 김진영//텐텐마진이 CJMALL의 마진보다 작을 때 orgprice로 전송 시작
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		'2013-07-24 김진영//텐텐마진이 CJMALL의 마진보다 작을 때 orgprice로 전송 끝

		optaddprice = 0
		strSql = ""
		strSql = strSql & " SELECT top 900 i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(40),o.optionname) as optionname, o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, isnull(R.outmallOptCode, '') as outmallOptCode, i.deliverfixday, isnull(o.optaddprice,'') as optaddprice " & VBCRLF
		strSql = strSql & " ,DATALENGTH(o.optionname) as optnmLen" & VBCRLF
		strSql = strSql & " FROM db_AppWish.dbo.tbl_item as i " & VBCRLF
		strSql = strSql & " JOIN db_AppWish.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF ''LEFT Join => Join
		strSql = strSql & " LEFT JOIN [db_outmall].[dbo].tbl_OutMall_regedoption as R on i.itemid = R.itemid and R.itemoption = o.itemoption and R.mallid='cjmall' " & VBCRLF
		strSql = strSql & " WHERE i.itemid = "&Fitemid
		strSql = strSql & " ORDER BY o.optaddprice ASC, o.itemoption ASC "
		rsCTget.Open strSql, dbCTget, 1
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			For i = 1 to rsCTget.RecordCount
				If rsCTget("outmallOptCode") = "" Then
					itemSu = getOptionLimitNo
					FItemOption 	= rsCTget("itemoption")
					optionname 		= rsCTget("optionname")
					Foptsellyn 		= rsCTget("optsellyn")
					Foptlimityn 	= rsCTget("optlimityn")
					Foptlimitno 	= rsCTget("optlimitno")
					Foptlimitsold 	= rsCTget("optlimitsold")
					optaddprice		= rsCTget("optaddprice")
					If rsCTget("deliverfixday") = "C" OR rsCTget("deliverfixday") = "X" Then
						fixday = "60"
					Else
						fixday = "20"
					End If

                    if rsCTget("optnmLen")>40 then
					    optionname=DdotFormat(optionname,20)
					end if

					If itemSu <> 0 Then
						strRst = strRst &"	<tns:unit>"
						strRst = strRst &"		<tns:unitNm><![CDATA["&optionname&"]]></tns:unitNm>"					'단품정보 - 단품상세(옵션명을 텍스트로 넘기면 됨)
						strRst = strRst &"		<tns:unitRetail>"&FSellCash+optaddprice&"</tns:unitRetail>"				'단품정보 - 판매가
						strRst = strRst &"		<tns:unitCost>"&getCJmallSuplyPrice(optaddprice)&"</tns:unitCost>"		'단품정보 - 매입원가
						strRst = strRst &"		<tns:availableQty>"&itemSu&"</tns:availableQty>"						'단품정보 - 공급가능수량 (상품 재고 파악이 안되는경우는 999같은 숫자를 넣습니다.)
						If getzCostomMadeInd = "Y" Then
							strRst = strRst &"		<tns:leadTime>"&getzLeadTime()&"</tns:leadTime>"					'단품정보 - 리드타임 (* 출고리드타임 예외등록 해야 함, 리드타임 기준 협의요함 00 : 오늘배송, 01 : 당일출고, 02 : 익일출고, 03 : 2일후출고, 04 : 4일, 05 : 5일, 06 : 6일.....)
	        			Else
	        				strRst = strRst &"		<tns:leadTime>03</tns:leadTime>"									'단품정보 - 리드타임 (* 출고리드타임 예외등록 해야 함, 리드타임 기준 협의요함 00 : 오늘배송, 01 : 당일출고, 02 : 익일출고, 03 : 2일후출고, 04 : 4일, 05 : 5일, 06 : 6일.....)
	        			End If
						strRst = strRst &"		<tns:unitApplyRsn>"&fixday&"</tns:unitApplyRsn>"						'단품정보 - 적용사유 (10 : 적용안함, 20 : 상품포장, 30 : 상품생산, 40 : 입고검사, 50 : 출고검사, 60 : 설치상품)
						strRst = strRst &"		<tns:startSaleDt>"&FormatDate(now(), "0000-00-00")&"</tns:startSaleDt>"	'단품정보 - 판매시작일자
						strRst = strRst &"		<tns:endSaleDt>9999-12-30</tns:endSaleDt>"								'단품정보 - 판매종료일자 (판매상태수정에서..)
						strRst = strRst &"		<tns:vpn>"&rsCTget("itemid")&"_"&FItemOption&"</tns:vpn>"				'단품정보 - 협력사상품코드
						strRst = strRst &"	</tns:unit>"
					End If
				End If
				rsCTget.MoveNext
			Next
		End If
		rsCTget.Close
		getCJOptionParamToEdit = strRst
	End Function

	'// CJMALL 판매여부 반환
	Public Function getCjmallSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold > CMAXLIMITSELL) then
				getCjmallSellYn = "Y"
			Else
				getCjmallSellYn = "N"
			End If
		Else
			getCjmallSellYn = "N"
		End If
	End Function

	'상품 등록 XML
	Public Function getCjmallItemRegXML
		Dim strRst
		Dim ioriginCode, ioriginname
		Dim makercompCode, makercompName
		ioriginCode 	= getOriginName2Code(Fsourcearea, ioriginname) 		'원산지코드
		makercompCode	= getmakerName2Code(Fsocname_kor, makercompName)	'제조사코드
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_01' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_01.xsd'>"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"									'!!!협력업체코드
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"					'!!!인증키
		strRst = strRst &"<tns:good>"
		strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"												'!!!상품분류체계 - 가등록채널구분(30:인터넷, 40:카탈로그)
		strRst = strRst &"	<tns:tGrpCd>"&FCddKey&"</tns:tGrpCd>"										'!!!상품분류체계 - 상품분류
		strRst = strRst &"	<tns:uniqBrandCd>"&CUNIQBRANDCD&"</tns:uniqBrandCd>"						'!!!상품분류체계 - 브랜드(텐바이텐:24049000)
		strRst = strRst &"	<tns:giftInd>Y</tns:giftInd>"											    '!!!상품분류체계 - 상품구분 (Y=일반판매상품, N=사은품)
		strRst = strRst &"	<tns:uniqMkrNatCd>"&ioriginCode&"</tns:uniqMkrNatCd>"						'!!!상품분류체계 - 제조국
		strRst = strRst &"	<tns:uniqMkrCompCd>"&makercompCode&"</tns:uniqMkrCompCd>"					'!!!상품분류체계 - 제조사
'		strRst = strRst &"	<tns:ingredient></tns:ingredient>"											'상품분류체계 - 주원료명	(샘플 페이지에는 누락)
'		strRst = strRst &"	<tns:zingredientOrigin></tns:zingredientOrigin>"							'상품분류체계 - 원료원산지	(샘플 페이지에는 누락) // 상품분류(대분류)가 식품일때만 원산지 필수(라던데..;;)
	If Fitemid = "899506" Then
		strRst = strRst &"	<tns:mdCode>5066</tns:mdCode>"
	Else
		strRst = strRst &"	<tns:mdCode>"&MD_CODE&"</tns:mdCode>"										'!!!MD코드						(있는 샘플도 있고, 누락된 샘플도 있음) 현아씨 문의 (텐바이텐 으로 가능)
	End If
		strRst = strRst &"	<tns:itemDesc><![CDATA["&DDotFormat(getItemNameFormat, 100)&"]]></tns:itemDesc>"			'!!!기본정보 - 상품명(120자 제약) (샘플에 CDATA없던거 추가)
		strRst = strRst &"	<tns:zLocalBolDesc><![CDATA["&DDotFormat(getItemNameFormat, 10)&"]]></tns:zLocalBolDesc>"	'!!!기본정보 - 운송장명(40자 제약)
		strRst = strRst &"	<tns:zlocalCcDesc><![CDATA["&DDotFormat(getItemNameFormat, 5)&"]]></tns:zlocalCcDesc>"		'!!!기본정보 - SMS상품명(20자 제약)
		strRst = strRst &"	<tns:vatCode>"&CHKIIF(FVatInclude="N","E","S")&"</tns:vatCode>"			 	'!!!기본정보 - 과세형태 (S:과세, E:면세, N:비과세, Z:영세)
		strRst = strRst &"	<tns:zDeliveryType>20</tns:zDeliveryType>"									'!!!기본정보 - 배송구분 (10:센터배송, 20:협력사배송, 30:직택배, 35:직택배Ⅱ, 40:직송, 99:배송없음)
		strRst = strRst &"	<tns:zShippingMethod>"&getdeliverfixday&"</tns:zShippingMethod>"			'!!!기본정보 - 배송유형 (10:택배배송, 20:설치상품, 30:배달서비스, 40:우편/등기배송) ''화물배송 확인
		strRst = strRst &"	<tns:courier>22</tns:courier>"												'!!!기본정보 - 택배사 (메인택배사 하나 지정 후 고정값 등록)(11:현대택배, 12:대한통운, 15:한진택배, 22:CJGLS, 29:CJHTH, 87:동부익스프레스) CJ택배 코드로 등록
		strRst = strRst &"	<tns:deliveryHomeCost>2500</tns:deliveryHomeCost>"							'기본정보 - 배송비 (배송구분이 협력사배송, 직송일 경우 필수 입력)
		strRst = strRst &"	<tns:zreturnNotReqInd>10</tns:zreturnNotReqInd>"							'기본정보 - 회수구분 (배송구분에 따라 필수/옵션)
'		strRst = strRst &"	<tns:zJointPackingQty></tns:zJointPackingQty>"								'기본정보 - 합포장단위 (배송구분에 따라 필수/옵션) (샘플페이지에는 누락)
		strRst = strRst &"	<tns:zCostomMadeInd>"&getzCostomMadeInd()&"</tns:zCostomMadeInd>"			'!!!기본정보 - 주문제작여부 (Y=주문제작, N=주문제작안함)) ''' 주문제작상품, 주문후제작상품 =>Y
		strRst = strRst &"	<tns:stockMgntLevel>2</tns:stockMgntLevel>"									'기본정보 - 재고관리레벨 (1=판매코드,2=단품코드)
'		strRst = strRst &"	<tns:leadtime></tns:leadtime>"												'기본정보 - 리드타임 (1. 프라자는 NULL셋팅 2.재고관리레벨이 "판매코드"일때 필수) (샘플페이지에는 누락)
'		strRst = strRst &"	<tns:leadtimeChgRsn></tns:leadtimeChgRsn>"									'기본정보 - 적용사유 (1. 프라자는 NULL셋팅 2.재고관리레벨이 "판매코드"일때 필수) (샘플페이지에는 누락)
		strRst = strRst &"	<tns:lowpriceInd>"&CHKIIF(IsCjFreeBeasong=False,"Y","N")&"</tns:lowpriceInd>"	'!!!기본정보 - 유료배송여부 (Y=유료배송,N=무료배송)        '' 확인.
		strRst = strRst &"	<tns:delayShipRewardIind>N</tns:delayShipRewardIind>"						'기본정보 - 지연보상여부 (Y=지연보상,N=지연보상안함)
'		strRst = strRst &"	<tns:packingMethod></tns:packingMethod>"									'기본정보 - 입고형태 (센터배송인 경우만 입력)
'		strRst = strRst &"	<tns:zOrderMaxQty></tns:zOrderMaxQty>"										'기본정보 - 1회최대주문수량 (고객당 1회 최대 주문가능 수량. 미입력시 제한없음
'		strRst = strRst &"	<tns:zDayOrderMaxQty></tns:zDayOrderMaxQty>"								'기본정보 - 1일최대주문수량 (고객당 일일 최대 주문가능 수량. 미입력시 제한없음)
		strRst = strRst &"	<tns:reserveDayInd>Y</tns:reserveDayInd>"									'기본정보 - 예약배송방식 (* 디폴트: YN-주문즉시 출하지시 Y-최초공급가능일 출하지시_Default)
		strRst = strRst &"	<tns:zContactSeqNo>"&chkiif(application("Svr_Info")="Dev","10002","10002")&"</tns:zContactSeqNo>"		'기본정보 - 협력사담당자
		strRst = strRst &"	<tns:zSupShipSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zSupShipSeqNo>"		'기본정보 - 출하지
		strRst = strRst &"	<tns:zReturnSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zReturnSeqNo>"			'기본정보 - 회수지
		strRst = strRst &"	<tns:zAsSupShipSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zAsSupShipSeqNo>"	'기본정보 - AS출하지
		strRst = strRst &"	<tns:zAsReturnSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zAsReturnSeqNo>"		'기본정보 - AS회수지
		strRst = strRst & getCJOptionParamToReg															'단품정보
		strRst = strRst &"	<tns:mallitem>"
		strRst = strRst &"		<tns:mallItemDesc><![CDATA["&"텐바이텐 " & Fsocname_kor & " "&DDotFormat(getItemNameFormat, 186)&"]]></tns:mallItemDesc>"	'!!!CJmall상품정보 - CJmall상품명 , 텐바이텐 브랜드명 추가
		strRst = strRst &"		<tns:keyword><![CDATA["&"텐바이텐;"&replace(Fkeywords,",",";")&"]]></tns:keyword>"						'!!!CJmall상품정보 - 검색키워드
		strRst = strRst & getCjCateParamToReg															'!!!메인카테고리여부(Y=카테고리,N=카테고리아님) // CJmall카테고리(세)
		strRst = strRst &"	</tns:mallitem>"
'		strRst = strRst &"	<tns:cert>"																			'QC에러 해결하려면 아래정보가 필요한듯..(2013-06-04 김진영)
'		strRst = strRst &"		<tns:certCode>350504</tns:certCode>"											'품질인증정보 - 항목코드
'		strRst = strRst &"		<tns:certNo>YU11100-12001</tns:certNo>"											'품질인증정보 - 인증번호 - 길이제약(50)
'		strRst = strRst &"		<tns:issueDate>2012-06-04</tns:issueDate>"										'품질인증정보 - 발급일자
'		strRst = strRst &"		<tns:certDate>2012-06-05</tns:certDate>"         								'품질인증정보 - 인증일자
'		strRst = strRst &"		<tns:avlStartDate>2012-06-04</tns:avlStartDate>"								'품질인증정보 - 유효기간(FROM)
'		strRst = strRst &"		<tns:avlEndDate>2013-06-04</tns:avlEndDate>"      								'품질인증정보 - 유효기간(TO)
'		strRst = strRst &"		<tns:itemModel>item</tns:itemModel>"        									'품질인증정보 - 상품명 및 모델명	-길이제약(200)
'		strRst = strRst &"		<tns:orgCode>전기인증</tns:orgCode>"            								'품질인증정보 - 인증검사기관명		-길이제약(200)
'		strRst = strRst &"		<tns:certField>전기제품</tns:certField>"        								'품질인증정보 - 인증분야			-길이제약(200)
'		strRst = strRst &"		<tns:originCode>원산지</tns:originCode>"     									'품질인증정보 - 원산지(제조국)
'		strRst = strRst &"		<tns:certSpec>세부</tns:certSpec>"          									'품질인증정보 - 세부사항			-길이제약(2000)
'		strRst = strRst &"	</tns:cert>"
		strRst = strRst & getCjmallItemInfoCdToReg()													'상품기술서
		strRst = strRst &"	<tns:goodsReport>"
		strRst = strRst &"		<tns:pedfId>91059</tns:pedfId>"
		strRst = strRst &"		<tns:html>"
		strRst = strRst &"			<![CDATA["&getCJItemContParamToReg&"]]>"
		strRst = strRst &"		</tns:html>"
		strRst = strRst &"	</tns:goodsReport>"
														'daebeak	대백상품추가정보 빠져있음
		strRst = strRst & getCJAddImageParamToReg		'!!!이미지정보
		strRst = strRst &"</tns:good>"
		strRst = strRst &"</tns:ifRequest>"
		getCjmallItemRegXML = strRst
'		response.write strRst
'		response.end
	End Function

	'상품 상태 변경 XML
	Public Function getcjmallItemSellStatusDTXML(cmd)
		Dim stopYN, strRst

		If cmd = "N" Then
			stopYN = "I"
		ElseIf cmd = "Y" Then
			stopYN = "A"
		End If

		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_03' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_03.xsd'>"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"					'!!!협력업체코드
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"	'!!!인증키
		strRst = strRst &"<tns:itemStates>"
		strRst = strRst &"	<tns:typeCd>01</tns:typeCd>"								'!!!01=판매코드,02=단품코드)
		strRst = strRst &"	<tns:itemCd_zip>"&FcjmallPrdNo&"</tns:itemCd_zip>"
		strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
		strRst = strRst &"	<tns:packInd>"&stopYN&"</tns:packInd>"						'!!!A-진행, I-일시중단
		strRst = strRst &"</tns:itemStates>"
		strRst = strRst &"</tns:ifRequest>"
		getcjmallItemSellStatusDTXML = strRst
	End Function

	'정보 수정 XML
	Public Function getcjmallItemModXML()
		Dim strRst
		Dim ioriginCode, ioriginname
		ioriginCode = getOriginName2Code(Fsourcearea, ioriginname)
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_02"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_02.xsd"">"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"												'!!!협력업체코드
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"								'!!!인증키
		strRst = strRst &"<tns:good>"
		strRst = strRst &"	<tns:sItem>"&FcjmallPrdNo&"</tns:sItem>"												'!!!판매상품코드(홈쇼핑)
	If Fitemid = "899506" Then
		strRst = strRst &"	<tns:loc>110</tns:loc>"																	'!!!상품분류체계 - 등록채널구분(공동구매)
	Else
		strRst = strRst &"	<tns:loc>30</tns:loc>"																	'!!!상품분류체계 - 등록채널구분(store포맷)
	End If
		strRst = strRst &"	<tns:zLocalBolDesc><![CDATA["&DDotFormat(getItemNameFormat, 10)&"]]></tns:zLocalBolDesc>"		'!!!기본정보 - 운송장명
		strRst = strRst &"	<tns:zlocalCcDesc><![CDATA["&DDotFormat(getItemNameFormat, 5)&"]]></tns:zlocalCcDesc>"			'!!!기본정보 - SMS상품명
		strRst = strRst &"	<tns:zContactSeqNo>"&chkiif(application("Svr_Info")="Dev","10002","10002")&"</tns:zContactSeqNo>"		'!!!기본정보 - 협력사담당자
		strRst = strRst &"	<tns:zSupShipSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zSupShipSeqNo>"		'!!!기본정보 - 출하지
		strRst = strRst &"	<tns:zReturnSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zReturnSeqNo>"			'!!!기본정보 - 회수지
		strRst = strRst &"	<tns:zAsSupShipSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zAsSupShipSeqNo>"	'!!!기본정보 - AS출하지
		strRst = strRst &"	<tns:zAsReturnSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zAsReturnSeqNo>"		'!!!기본정보 - AS회수지
        strRst = strRst &"	<tns:lowpriceInd>"&CHKIIF(IsCjFreeBeasong=False,"Y","N")&"</tns:lowpriceInd>"	'!!!기본정보 - 유료배송여부 (Y=유료배송,N=무료배송)        '' 확인.
		strRst = strRst & getCJOptionParamToEdit                                                                      '' 확인해 볼것 ''864806
		strRst = strRst &"	<tns:mallitem>"
		strRst = strRst &"		<tns:mallItemDesc><![CDATA["&"텐바이텐 " & Fsocname_kor & " "&DDotFormat(getItemNameFormat, 186)&"]]></tns:mallItemDesc>"	'!!!CJmall상품정보 - CJmall상품명
		strRst = strRst &"	</tns:mallitem>"
		strRst = strRst & getCjmallItemInfoCdToReg()													'상품기술서
		strRst = strRst &"	<tns:goodsReport>"
		strRst = strRst &"		<tns:pedfId>91059</tns:pedfId>"
		strRst = strRst &"		<tns:html>"
		strRst = strRst &"			<![CDATA["&getCJItemContParamToReg&"]]>"
		strRst = strRst &"		</tns:html>"
		strRst = strRst &"	</tns:goodsReport>"
		strRst = strRst & getCJAddImageParamToReg		'!!!이미지정보
		strRst = strRst &"</tns:good>"
		strRst = strRst &"</tns:ifRequest>"
		getcjmallItemModXML = strRst
	End Function

	'단품 수정 XML
	Public Function getcjmallOptSellModXML
		Dim sqlStr, arrRows, i
		Dim itemoption, optiontypename, optionname, optLimit, optlimityn, isUsing, optsellyn, preged, optNameDiff, forceExpired, oopt, ooptCd, YtoN, NtoY, DelOpt
		Dim validSellno, strRst
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_03' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_03.xsd'>"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"					'!!!협력업체코드
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"	'!!!인증키

		sqlStr = "exec db_outmall.dbo.sp_Ten_OutMall_optEditParamList_cjmall 'cjmall'," & iitemid
		rsCTget.CursorLocation = adUseClient
		rsCTget.CursorType = adOpenStatic
		rsCTget.LockType = adLockOptimistic
		rsCTget.Open sqlStr, dbCTget
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			arrRows = rsCTget.getRows
		End If
		rsCTget.close

		If FMaySoldOut = "Y" Then
			strRst = strRst &"<tns:itemStates>"
			strRst = strRst &"<tns:typeCd>01</tns:typeCd>"						'01=판매코드,02=단품코드
			strRst = strRst &"<tns:itemCd_zip>"&Fcjmallprdno&"</tns:itemCd_zip>"
			strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
			strRst = strRst &"<tns:packInd>I</tns:packInd>"						'A-진행, I-일시중단
			strRst = strRst &"</tns:itemStates>"
		ElseIf FMaySoldOut = "N" Then
			strRst = strRst &"<tns:itemStates>"
			strRst = strRst &"<tns:typeCd>01</tns:typeCd>"						'01=판매코드,02=단품코드
			strRst = strRst &"<tns:itemCd_zip>"&Fcjmallprdno&"</tns:itemCd_zip>"
			strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
			strRst = strRst &"<tns:packInd>A</tns:packInd>"						'A-진행, I-일시중단
			strRst = strRst &"</tns:itemStates>"
		End If

		For i = 0 To UBound(ArrRows,2)
			itemoption		= ArrRows(1,i)
			optiontypename	= ArrRows(2,i)
			optionname		= Replace(Replace(db2Html(ArrRows(3,i)),":",""),",","")
			optLimit		= ArrRows(4,i)
			optlimityn		= ArrRows(5,i)
			isUsing			= ArrRows(6,i)
			optsellyn		= ArrRows(7,i)
			preged			= (ArrRows(11,i)=1)
			optNameDiff		= (ArrRows(12,i)=1)
			forceExpired	= (ArrRows(13,i)=1)
			oopt			= ArrRows(14,i)
			ooptCd			= ArrRows(15,i)
			YtoN			= (ArrRows(16,i)=1)
			NtoY			= (ArrRows(17,i)=1)
			DelOpt			= (ArrRows(18,i)=1)
			If FMaySoldOut = "Y" Then
				strRst = strRst &"<tns:itemStates>"
				strRst = strRst &"<tns:typeCd>02</tns:typeCd>"						'01=판매코드,02=단품코드
				strRst = strRst &"<tns:itemCd_zip>"&ooptCd&"</tns:itemCd_zip>"
				strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
				strRst = strRst &"<tns:packInd>I</tns:packInd>"						'A:진행, I:일시중단
				strRst = strRst &"</tns:itemStates>"
			ElseIf (forceExpired) or (optNameDiff) or (DelOpt) or (isUsing="N") or (optsellyn="N") or (optlimityn = "Y" AND optLimit <= 5) Then			'한정이고 수량이 5개 이하인 경우 // (isUsing="N") or (optsellyn="N") or 추가 2013/05/31..''2013-12-04 13:30 김진영..optLimit < 5를 optLimit <= 5로 수정
				strRst = strRst &"<tns:itemStates>"
				strRst = strRst &"<tns:typeCd>02</tns:typeCd>"						'01=판매코드,02=단품코드
				strRst = strRst &"<tns:itemCd_zip>"&ooptCd&"</tns:itemCd_zip>"
				strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
				strRst = strRst &"<tns:packInd>I</tns:packInd>"						'A:진행, I:일시중단
				strRst = strRst &"</tns:itemStates>"
		    ElseIf (preged) and (ooptCd <> "") Then
				strRst = strRst &"<tns:itemStates>"
				strRst = strRst &"<tns:typeCd>02</tns:typeCd>"						'01=판매코드,02=단품코드
				strRst = strRst &"<tns:itemCd_zip>"&ooptCd&"</tns:itemCd_zip>"
				strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
				strRst = strRst &"<tns:packInd>A</tns:packInd>"						'A:진행, I:일시중단
				strRst = strRst &"</tns:itemStates>"
			End If
		Next
		strRst = strRst &"</tns:ifRequest>"
		getcjmallOptSellModXML = strRst
	End Function

	'단품 수량 수정 XML
	Public Function getcjmallItemQTYXML
		Dim sqlStr, oneOpt, j
		Dim arrRows, i, strRst, validSellno
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_05"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_05.xsd"">"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"					'!!!협력업체코드
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"	'!!!인증키

		sqlStr = ""
		sqlStr = sqlStr & " select isnull(o.itemoption, '') as itemoption, r.outmallOptCode, r.outmallOptName "
		sqlStr = sqlStr & " from [db_outmall].[dbo].tbl_OutMall_regedoption as r "
		sqlStr = sqlStr & " left join [db_AppWish].[dbo].tbl_item_option as o on r.itemid = o.itemid and r.itemoption = o.itemoption "
		sqlStr = sqlStr & " where r.mallid = '"&CMALLNAME&"' and r.itemid="&Fitemid
		rsCTget.Open sqlStr, dbCTget
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			oneOpt = rsCTget.getRows
		End If
		rsCTget.close

		If (UBound(oneOpt ,2) = "0") and (oneOpt(2,0) = "단일상품") Then
			strRst = strRst &"<tns:ltSupplyPlans>"
			strRst = strRst &"	<tns:unitCd>"&oneOpt(1,0)&"</tns:unitCd>"
			strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
			strRst = strRst &"	<tns:strDt>"&FormatDate(now(), "0000-00-00")&"</tns:strDt>"
			If GetCJLmtQty = 0 Then
				strRst = strRst &"	<tns:endDt>"&FormatDate(now(), "0000-00-00")&"</tns:endDt>"
			Else
				strRst = strRst &"	<tns:endDt>9999-12-30</tns:endDt>"
			End If
			strRst = strRst &"	<tns:availSupQty>"&chkiif(GetCJLmtQty=0,"1",GetCJLmtQty)&"</tns:availSupQty>"
			strRst = strRst &"</tns:ltSupplyPlans>"
		Else
			sqlStr = ""
			sqlStr = sqlStr & " SELECT o.itemoption, o.optionTypeName, o.optionname, isnull(R.outmallOptCode, '') as outmallOptCode, (o.optlimitno-o.optlimitsold) as optLimit, o.optlimityn, o.isUsing, o.optsellyn " & VBCRLF
			sqlStr = sqlStr & " FROM [db_AppWish].[dbo].tbl_item_option o " & VBCRLF
			sqlStr = sqlStr & " left join [db_outmall].[dbo].tbl_OutMall_regedoption R on o.itemid=R.itemid and o.itemoption=R.itemoption and R.mallid='"&CMALLNAME&"' " & VBCRLF
			sqlStr = sqlStr & " where R.outmallOptCode <> '' and o.itemid="&Fitemid
			rsCTget.Open sqlStr, dbCTget
			If Not(rsCTget.EOF or rsCTget.BOF) Then
				arrRows = rsCTget.getRows
			End If
			rsCTget.close

			If isArray(arrRows) Then
				For i = 0 To UBound(ArrRows,2)
					validSellno = 999				'최대 999로 강제지정
					If (FSellyn <> "Y") or ((arrRows(5,i) = "Y") and (arrRows(4,i) < 1)) or (arrRows(6,i) <> "Y") or (arrRows(7,i) <> "Y") Then
						validSellno = 0
					End If

					If (arrRows(5,i) = "Y") Then
						validSellno = arrRows(4,i)
					End If

					If (validSellno < CMAXLIMITSELL) Then validSellno = 0
					If (arrRows(5,i) = "Y") and (validSellno > 0) Then
						validSellno = validSellno - CMAXLIMITSELL
					End If
					If (validSellno < 1) then validSellno = 0
					If IsSoldOut Then validSellno = 0

					strRst = strRst &"<tns:ltSupplyPlans>"
					strRst = strRst &"	<tns:unitCd>"&arrRows(3,i)&"</tns:unitCd>"
					strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
					strRst = strRst &"	<tns:strDt>"&FormatDate(now(), "0000-00-00")&"</tns:strDt>"
					If validSellno = 0 Then
						strRst = strRst &"	<tns:endDt>"&FormatDate(now(), "0000-00-00")&"</tns:endDt>"
					Else
						strRst = strRst &"	<tns:endDt>9999-12-30</tns:endDt>"
					End If
					strRst = strRst &"	<tns:availSupQty>"&chkiif(validSellno=0,"1",validSellno)&"</tns:availSupQty>"
					strRst = strRst &"</tns:ltSupplyPlans>"
				Next
			End If
		End If
		strRst = strRst &"</tns:ifRequest>"
		getcjmallItemQTYXML = strRst
	End Function

	'판매 가격 수정 XML
	Function getcjmallItemSellPriceModXML()
		Dim strRst, sqlStr, i, GetTenTenMargin
		'2013-07-24 김진영//텐텐마진이 CJMALL의 마진보다 작을 때 orgprice로 전송 시작
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		'2013-07-24 김진영//텐텐마진이 CJMALL의 마진보다 작을 때 orgprice로 전송 끝

		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_04"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_04.xsd"">"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"						'!!!협력업체코드
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"		'!!!인증키
        strRst = strRst &"<tns:itemPrices>"
		strRst = strRst &"	<tns:typeCD>01</tns:typeCD>"									'01이면 판매코드 / 02면 단품코드
		strRst = strRst &"	<tns:itemCD_ZIP>"&FcjmallPrdNo&"</tns:itemCD_ZIP>"
		strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
		strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
		strRst = strRst &"	<tns:newUnitRetail>"&MustPrice&"</tns:newUnitRetail>"
		strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice2&"</tns:newUnitCost>"
		strRst = strRst &"</tns:itemPrices>"
		strRst = strRst &"</tns:ifRequest>"
		getcjmallItemSellPriceModXML = strRst
	End Function

	'단품 가격 수정 XML
	Public Function getcjmallOptionPriceModXML()
		Dim strRst, sqlStr, arrrows, chkOption, i, optAddPRcExists, GetTenTenMargin
		optAddPRcExists = false
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT distinct o.itemid, o.optAddPrice,  ro.outmallOptCode, o.itemoption"
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item_option o "
		sqlStr = sqlStr & " JOIN [db_outmall].[dbo].tbl_OutMall_regedoption ro on o.itemid=ro.itemid and ro.mallid ='"&CMALLNAME&"' and ro.itemoption = o.itemoption "
		sqlStr = sqlStr & " WHERE o.itemid = '"&Fitemid&"' "
		sqlStr = sqlStr & " GROUP BY o.itemid, o.optAddPrice, ro.outmallOptCode, o.itemoption"
		sqlStr = sqlStr & " ORDER BY o.optAddPrice, o.itemoption"
		rsCTget.Open sqlStr, dbCTget
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			arrrows = rsCTget.getRows
			chkOption = True
		Else
			chkOption = False
		End If
		rsCTget.close

		if (chkOption) then
			For i = 0 To UBound(ArrRows,2)
				optAddPRcExists = optAddPRcExists or (arrRows(1,i)>0)
			Next
		end if

		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_04"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_04.xsd"">"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"						'!!!협력업체코드
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"		'!!!인증키

		If (Not optAddPRcExists) OR (chkOption = False) Then
			strRst = strRst &"<tns:itemPrices>"
			strRst = strRst &"	<tns:typeCD>01</tns:typeCD>"								'01이면 판매코드 / 02면 단품코드
			strRst = strRst &"	<tns:itemCD_ZIP>"&FcjmallPrdNo&"</tns:itemCD_ZIP>"
			strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
			strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
			strRst = strRst &"	<tns:newUnitRetail>"&MustPrice&"</tns:newUnitRetail>"
			strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice2&"</tns:newUnitCost>"
			strRst = strRst &"</tns:itemPrices>"
		Else
			If chkOption = True Then
			strRst = strRst &"<tns:itemPrices>"
			strRst = strRst &"	<tns:typeCD>01</tns:typeCD>"								'01이면 판매코드 / 02면 단품코드
			strRst = strRst &"	<tns:itemCD_ZIP>"&FcjmallPrdNo&"</tns:itemCD_ZIP>"
			strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
			strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
			strRst = strRst &"	<tns:newUnitRetail>"&MustPrice&"</tns:newUnitRetail>"
			strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice2&"</tns:newUnitCost>"
			strRst = strRst &"</tns:itemPrices>"
    			For i = 0 To UBound(ArrRows,2)
					strRst = strRst &"<tns:itemPrices>"
					strRst = strRst &"	<tns:typeCD>02</tns:typeCD>"						'01이면 판매코드 / 02면 단품코드
					strRst = strRst &"	<tns:itemCD_ZIP>"&arrRows(2,i)&"</tns:itemCD_ZIP>"
					strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
					strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
					strRst = strRst &"	<tns:newUnitRetail>"&MustPrice+arrRows(1,i)&"</tns:newUnitRetail>"
					strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice(arrRows(1,i))&"</tns:newUnitCost>"
					strRst = strRst &"</tns:itemPrices>"
					optAddPRcExists = optAddPRcExists or (arrRows(1,i)>0)
				Next
			End If
		End If
		strRst = strRst &"</tns:ifRequest>"
		getcjmallOptionPriceModXML = strRst
	End Function
End Class

Class CCjmall
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
	Public FRectItemName
	Public FRectCJMallPrdNo
	Public FRectCDL
	Public FRectCDM
	Public FRectCDS
	Public FRectItemID
	Public FRectEventid
	Public FRectExtNotReg
	Public FRectOrdType
	Public FRectMatchCate
	Public FRectPrdDivMatch
	''Public FRectMatchCateNotCheck
	Public FRectSellYn
	Public FRectLimitYn
	Public FRectSailYn
	Public FRectonlyValidMargin
	Public FRectFailCntExists
	Public FRectoptAddprcExists
	Public FRectoptAddprcExistsExcept
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectisMadeHand
    Public FRectCjSell10x10Soldout
    Public FRectexpensive10x10
	Public FRectExtSellYn
	public FRectOnlyNotUsingCheck
	public FRectdiffPrc
	public FRectCjmallYes10x10No
	public FRectCjmallNo10x10Yes
	public FRectInfoDiv
	public FRectFailCntOverExcept
	public FRectIsOption
	public FRectIsReged
	public FRectNotinmakerid
	public FRectReqEdit
	public FRectOPTCntEqual

	'카테고리
	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FRectDspNo

	Public FRectMode

	Public Finfodiv
	Public FCateName
	Public FsearchName
	Public FRectdisptpcd

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

	'// 텐바이텐-cjmall 카테고리
	Public Sub getTencjmallCateList
		Dim sqlStr, addSql, i

		If FRectCDL<>"" Then
			addSql = addSql & " and s.code_large='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and s.code_mid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and s.code_small='" & FRectCDS & "'"
		End if

		If FRectIsMapping = "Y" Then
			addSql = addSql & " and T.CateKey is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and T.CateKey is null "
		End if

		If FRectKeyword<>"" Then
			Select Case FRectSDiv
				Case "CCD"	'cjmall 전시코드 검색
					addSql = addSql & " and T.CateKey='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(텐바이텐 소분류명)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.D_Name,cc.L_Name,cc.M_Name,cc.S_Name, cc.isusing, cc.CateGbn "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_cjMall_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_cjMall_Category as cc on cc.CateKey = cm.CateKey  "  & VBCRLF
		If FRectdisptpcd <> "" Then
            sqlStr = sqlStr & " and cc.CateGbn='"&FRectdisptpcd&"'"
        End If
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " 	s.code_large,s.code_mid,s.code_small " & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm  "  & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "  & VBCRLF
		sqlStr = sqlStr & " ,code_nm as small_nm "  & VBCRLF
		sqlStr = sqlStr & " ,T.CateKey as DispNo ,T.D_Name as DispNm, T.L_Name as DispLrgNm, T.M_Name as DispMidNm, T.S_Name as DispSmlNm ,T.IsUsing as CateIsUsing,T.cateGbn as disptpcd "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.D_Name,cc.L_Name,cc.M_Name,cc.S_Name, cc.isusing, cc.CateGbn  "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_cjMall_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_cjMall_Category as cc on cc.CateKey = cm.CateKey  "  & VBCRLF
		If FRectdisptpcd <> "" Then
            sqlStr = sqlStr & " and cc.CateGbn='"&FRectdisptpcd&"'"
        End If
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		sqlStr = sqlStr & " ORDER BY s.code_large,s.code_mid,s.code_small, T.CateGbn  ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new cjmallItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FDispNo			= rsget("DispNo")
					FItemList(i).FDispNm			= db2html(rsget("DispNm"))
					FItemList(i).FDispLrgNm			= db2html(rsget("DispLrgNm"))
					FItemList(i).FDispMidNm			= db2html(rsget("DispMidNm"))
					FItemList(i).FDispSmlNm			= db2html(rsget("DispSmlNm"))
					FItemList(i).Fdisptpcd			= rsget("disptpcd")
	                FItemList(i).FCateIsUsing		= rsget("CateIsUsing")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// 텐바이텐-cjmall 상품분류
	Public Sub getTencjmallprdDivList
		Dim sqlStr, addSql, i

		If FRectCDL<>"" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End if

		If Finfodiv <> "" Then
			addSql = addSql & " and c.infodiv='" & Finfodiv & "'"
		End if

		If FRectIsMapping <> "" Then
			If FRectIsMapping = "Y" Then
				addSql = addSql & " and isnull(P.CddKey, '') <> '' "
			ElseIf FRectIsMapping = "N" Then
				addSql = addSql & " and isnull(P.CddKey, '') = '' "
			End If
		End if

		If FCateName <> "" AND FsearchName <> "" Then
			Select Case FCateName
				Case "cdlnm"
					addSql = addSql & " and v.nmlarge like '%" & FsearchName & "%'"
				Case "cdmnm"
					addSql = addSql & " and v.nmmid like '%" & FsearchName & "%'"
				Case "cdsnm"
					addSql = addSql & " and v.nmsmall like '%" & FsearchName & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM  ( " & VBCRLF
		sqlStr = sqlStr & " 	SELECT c.infodiv, i.cate_large, i.cate_mid, i.cate_small " & VBCRLF
		sqlStr = sqlStr & " 	, v.nmlarge, v.nmmid, v.nmsmall , count(*) as icnt " & VBCRLF
		sqlStr = sqlStr & "		,P.CddKey ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name,P.IsUsing as PrdDivIsUsing, P.infodiv as Pinfodiv "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_item i " & VBCRLF
		sqlStr = sqlStr & " 	INNER JOIN db_item.dbo.tbl_item_contents c on i.itemid = C.itemid " & VBCRLF
		sqlStr = sqlStr & " 	LEFT JOIN db_item.dbo.vw_category v	on i.cate_large = v.cdlarge and i.cate_mid = v.cdmid and i.cate_small = v.cdsmall " & VBCRLF
		sqlStr = sqlStr & "		LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 		SELECT dm.CddKey, dm.tenCateLarge,dm.tenCateMid, dm.tenCateSmall, pv.cdd_Name, pv.cdl_Name, pv.cdm_Name, pv.cds_Name, pv.isusing, dm.infodiv "  & VBCRLF
		sqlStr = sqlStr & " 		FROM db_etcmall.dbo.tbl_cjMall_prdDiv_mapping as dm  "  & VBCRLF
		sqlStr = sqlStr & " 		JOIN db_etcmall.dbo.tbl_cjMall_prdDiv as pv on dm.CddKey = pv.cdd  "  & VBCRLF
		sqlStr = sqlStr & " 	) P on P.tenCateLarge=i.cate_large and P.tenCateMid=i.cate_mid and P.tenCateSmall=i.cate_small and P.infodiv = c.infodiv   "  & VBCRLF
		sqlStr = sqlStr & " 	WHERE i.sellyn = 'Y' and v.nmlarge is not null and isNULL(c.infodiv,'')<>'' "&addsql&" " & VBCRLF
		sqlStr = sqlStr & " 	GROUP BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small, v.nmlarge, v.nmmid, v.nmsmall,P.CddKey ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.IsUsing, P.infodiv " & VBCRLF
		sqlStr = sqlStr & " ) as T " & VBCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " c.infodiv, i.cate_large, i.cate_mid, i.cate_small " & VBCRLF
		sqlStr = sqlStr & " , v.nmlarge, v.nmmid, v.nmsmall , count(*) as icnt " & VBCRLF
		sqlStr = sqlStr & " ,P.CddKey ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.IsUsing as PrdDivIsUsing, P.infodiv as Pinfodiv "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_item.dbo.tbl_item_contents c on i.itemid = C.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.vw_category v	on i.cate_large = v.cdlarge and i.cate_mid = v.cdmid and i.cate_small = v.cdsmall " & VBCRLF
		sqlStr = sqlStr & "	LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT dm.CddKey, dm.tenCateLarge,dm.tenCateMid, dm.tenCateSmall, pv.cdd_Name, pv.cdl_Name, pv.cdm_Name, pv.cds_Name, pv.isusing, dm.infodiv "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_cjMall_prdDiv_mapping as dm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_cjMall_prdDiv as pv on dm.CddKey = pv.cdd  "  & VBCRLF
		sqlStr = sqlStr & " ) P on P.tenCateLarge=i.cate_large and P.tenCateMid=i.cate_mid and P.tenCateSmall=i.cate_small and P.infodiv = c.infodiv  "  & VBCRLF
		sqlStr = sqlStr & " WHERE i.sellyn = 'Y' and v.nmlarge is not null and isNULL(c.infodiv,'')<>'' "&addsql&" " & VBCRLF
		sqlStr = sqlStr & " GROUP BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small, v.nmlarge, v.nmmid, v.nmsmall,P.CddKey ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.IsUsing, P.infodiv " & VBCRLF
		sqlStr = sqlStr & " ORDER BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new cjmallItem
					FItemList(i).Finfodiv		= rsget("infodiv")
					FItemList(i).FtenCateLarge	= rsget("cate_large")
					FItemList(i).FtenCateMid	= rsget("cate_mid")
					FItemList(i).FtenCateSmall	= rsget("cate_small")
					FItemList(i).FtenCDLName	= rsget("nmlarge")
					FItemList(i).FtenCDMName	= rsget("nmmid")
					FItemList(i).FtenCDSName	= rsget("nmsmall")
					FItemList(i).Ficnt			= rsget("icnt")
					FItemList(i).FCddKey		= rsget("CddKey")
					FItemList(i).Fcdd_Name		= rsget("cdd_Name")
					FItemList(i).Fcdl_Name		= rsget("cdl_Name")
					FItemList(i).Fcdm_Name		= rsget("cdm_Name")
					FItemList(i).Fcds_Name		= rsget("cds_Name")
					FItemList(i).FPrdDivIsUsing	= rsget("PrdDivIsUsing")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function getTencjmallOneprdDiv
		Dim sqlStr, addSql, addsql2

		If FRectCDL<>"" Then
			addSql = addSql & " and v.cdlarge='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and v.cdmid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and v.cdsmall='" & FRectCDS & "'"
		End if

		If Finfodiv <> "" Then
			addSql2 = addSql2 & " and p.infodiv='" & Finfodiv & "'"
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top 1 p.CddKey, p.infodiv, p.CdmKey, p.tenCateLarge, p.tenCateMid, p.tenCateSmall, v.nmlarge, v.nmmid, v.nmsmall, T.cdd_NAME " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.vw_category as v " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_cjMall_prdDiv_mapping p on p.tenCateLarge = v.cdlarge and p.tenCateMid = v.cdmid and p.tenCateSmall = v.cdsmall " & addsql2
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_cjMall_prdDiv as T on p.cddKey = T.cdd " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addsql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		If not rsget.EOF Then
			Set FItemList(0) = new cjmallItem
				FItemList(0).Finfodiv		= rsget("infodiv")
				FItemList(0).FtenCateLarge	= rsget("tenCateLarge")
				FItemList(0).FtenCateMid	= rsget("tenCateMid")
				FItemList(0).FtenCateSmall	= rsget("tenCateSmall")
				FItemList(0).FtenCDLName	= rsget("nmlarge")
				FItemList(0).FtenCDMName	= rsget("nmmid")
				FItemList(0).FtenCDSName	= rsget("nmsmall")
				FItemList(0).FCddKey		= rsget("CddKey")
				FItemList(0).Fcdd_Name		= rsget("cdd_NAME")
		End If
		rsget.Close
	End Function

	'// cjmall 카테고리
	Public Sub getcjmallCategoryList
		Dim sqlStr, addSql, i

		If FRectDspNo <> "" Then
			addSql = addSql & " and c.cateKey = " & FRectDspNo
		End If

		If FRectKeyword <> "" Then
			Select Case FRectSDiv
				Case "CCD"	'cjmall 전시코드 검색
					addSql = addSql & " and c.cateKey='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명
					addSql = addSql & " and (c.D_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.S_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.M_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.L_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " )"
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(c.cateKey) as cnt, CEILING(CAST(Count(c.cateKey) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_cjMall_Category as c " & VBCRLF
		sqlStr = sqlStr & " WHERE c.M_Name like '%텐바이텐%' " & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage) & " c.* " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_cjMall_Category as c " & VBCRLF
		sqlStr = sqlStr & " WHERE c.M_Name like '%텐바이텐%' " & addSql
		sqlStr = sqlStr & " ORDER BY c.cateKey ASC"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new cjmallItem
					FItemList(i).FDispNo		= rsget("cateKey")
					FItemList(i).FDispNm		= db2html(rsget("D_Name"))
					FItemList(i).FDispLrgNm		= db2html(rsget("L_Name"))
					FItemList(i).FDispMidNm		= db2html(rsget("M_Name"))
					FItemList(i).FDispSmlNm		= db2html(rsget("S_Name"))
					FItemList(i).FDispThnNm		= db2html(rsget("D_Name"))
					FItemList(i).FisUsing		= rsget("isUsing")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// cjmall 상품분류
	Public Sub getcjmallPrdDivList
		Dim sqlStr, addSql, i

		If Finfodiv <> "" Then
			addSql = addSql & " and m.infodiv = '" & Finfodiv & "'"
		End If

		If FsearchName <> "" Then
			addSql = addSql & " and p.cdd_NAME like '%" & FsearchName & "%'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " from db_etcmall.dbo.tbl_cjmall_PrddivMid_map as m " & VBCRLF
		sqlStr = sqlStr & " inner join db_etcmall.dbo.tbl_cjMall_prdDiv as p on m.cjcdm = p.cdm " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage) & " m.infodiv, p.cdm, p.cdd, p.cdl_NAME, p.cdm_NAME, p.cds_NAME, p.cdd_NAME, p.isusing  " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_cjmall_PrddivMid_map as m " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_etcmall.dbo.tbl_cjMall_prdDiv as p on m.cjcdm = p.cdm " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY m.infodiv, p.cdm, p.cdd"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new cjmallItem
					FItemList(i).Finfodiv		= rsget("infodiv")
					FItemList(i).FCdm			= rsget("cdm")
					FItemList(i).FCdd			= rsget("cdd")
					FItemList(i).FDispNm		= db2html(rsget("cdd_NAME"))
					FItemList(i).FDispLrgNm		= db2html(rsget("cdl_NAME"))
					FItemList(i).FDispMidNm		= db2html(rsget("cdm_NAME"))
					FItemList(i).FDispSmlNm		= db2html(rsget("cds_NAME"))
					FItemList(i).FDispThnNm		= db2html(rsget("cdd_NAME"))
					FItemList(i).Fisusing		= rsget("isusing")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'등록 상품 리스트
	Public Sub getCjmallRegedItemList
		Dim i, sqlStr, addSql
		'브랜드검색
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'상품명 검색
		If FRectItemName <> "" Then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		End if

		'Cjmall상품번호 검색
        If (FRectCJMallPrdNo <> "") then
            If Right(Trim(FRectCJMallPrdNo) ,1) = "," Then
            	FRectItemid = Replace(FRectCJMallPrdNo,",,",",")
            	addSql = addSql & " and J.cjmallPrdNo in (" + Left(FRectCJMallPrdNo,Len(FRectCJMallPrdNo)-1) + ")"
            Else
				FRectCJMallPrdNo = Replace(FRectCJMallPrdNo,",,",",")
            	addSql = addSql & " and J.cjmallPrdNo in (" + FRectCJMallPrdNo + ")"
            End If
        End If

		'카테고리 검색
		If FRectCDL <> "" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if
		If FRectCDM <> "" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if
		If FRectCDS <> "" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End If

		'등록여부 검색
		Select Case FRectExtNotReg
			Case "M"	'미등록
			    addSql = addSql & " and J.itemid is NULL  and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
			Case "Q"	''등록실패
				addSql = addSql & " and J.cjmallStatCd = -1"
			Case "J"	'등록예정이상
				addSql = addSql & " and J.cjmallStatCd >= 0"
			Case "W"	'등록예정
				addSql = addSql & " and J.cjmallStatCd = 0"
		    Case "A"	'전송시도중오류
				addSql = addSql & " and J.cjmallStatCd = 1"
			Case "F"	'등록완료(임시)
			    addSql = addSql & " and J.cjmallStatCd = 3"
			Case "D"	'등록완료(전시)
			    addSql = addSql & " and J.cjmallStatCd = 7"
				addSql = addSql & " and J.cjmallPrdNo is Not Null"
			Case "R"	'수정요망		'스케줄링에서 사용
				addSql = addSql & " and J.cjmallStatCd = 7"
				addSql = addSql & " and J.cjmallLastUpdate < i.lastupdate"
				addSql = addSql & " and isnull(J.cjmallprdno, '') <> '' "
		End Select

		'미등록 라디오버튼 클릭 시
		Select Case FRectIsReged
			Case "N"	'등록예정이상
			    addSql = addSql & " and J.itemid is NULL  and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
		End Select

		'판매여부 검색
		Select Case FRectSellYn
			Case "Y"	addSql = addSql & " and i.sellYn='Y'"			'판매
			Case "N"	addSql = addSql & " and i.sellYn in ('S','N')"	'품절
		End Select

		'텐바이텐 한정여부 검색
		If FRectLimitYn <> "" Then
			addSql = addSql & " and i.limitYn = '" & FRectLimitYn & "'"
		End If

		'텐바이텐 세일여부 검색
		If FRectSailYn <> "" Then
			addSql = addSql & " and i.sailYn = '" & FRectSailYn & "'"
		End If

		'역마진 및 마진 CMAXMARGIN 이상 검색
		If (FRectonlyValidMargin <> "") Then
			IF (FRectonlyValidMargin = "Y") Then
				addSql = addSql & " and i.sellcash<>0"
				addSql = addSql & " and i.sellcash - i.buycash > 0 "
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&CMAXMARGIN & VbCrlf
			Else
				addSql = addSql & " and i.sellcash<>0"
				addSql = addSql & " and i.sellcash - i.buycash > 0 "
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<"&CMAXMARGIN & VbCrlf
			End If
		End If

		'주문제작 여부 검색
		If FRectisMadeHand <> "" Then
			If (FRectisMadeHand = "Y") Then
				addSql = addSql & " and i.itemdiv in ('06', '16')" & VbCrlf
			Else
				addSql = addSql & " and i.itemdiv not in ('06', '16')" & VbCrlf
			End If
		End if

		'옵션 여부 검색
		If FRectIsOption <> "" Then
			If FRectIsOption = "optAll" Then			'옵션전체
				addSql = addSql & " and i.optioncnt > 0"
			ElseIf FRectIsOption = "optaddpricey" Then	'추가금액Y
				addSql = addSql & " and i.optioncnt > 0"
 				addSql = addSql & " and J.optAddPrcCnt > 0"
			ElseIf FRectIsOption = "optaddpricen" Then	'추가금액N
				addSql = addSql & " and i.optioncnt > 0"
				addSql = addSql & " and isNULL(J.optAddPrcCnt,0)=0"
			ElseIf FRectIsOption = "optN" Then			'단품
				addSql = addSql & " and i.optioncnt = 0"
			End If
		End If

		'텐바이텐 품목정보 검색
		If (FRectInfoDiv <> "") then
			If (FRectInfoDiv = "YY") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')<>''"
			ElseIf (FRectInfoDiv = "NN") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')=''"
			Else
				addSql = addSql & " and ct.infodiv = '"&FRectInfoDiv&"'"
			End If
		End If

		'텐바이텐 등록제외 브랜드 제외 검색
		If (FRectNotinmakerid <> "") then
			If (FRectNotinmakerid = "Y") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
			End If
		End If

		'CJ몰 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.cjmallSellYn <> 'X'"
			Else
				addSql = addSql & " and J.cjmallSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'등록수정오류
'		If (FRectFailCntExists <> "") Then
'			addSql = addSql & " and J.accFailCNT>0"
'		End If

		'CJ몰 카테고리 매칭 여부
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(c.mapCnt, '') <> ''"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(c.mapCnt, '') = ''"
		End Select

		'분류매칭 검색
		Select Case FRectPrdDivMatch
			Case "Y"	'매칭완료
				addSql = addSql & " and IsNull(pd.cddkey, '') <> '' "
			Case "N"	'미매칭
				addSql = addSql & " and IsNull(pd.cddkey, '') = '' "
		End Select

		'등록수정오류상품
		Select Case FRectFailCntExists
			Case "Y"	'오류1회이상
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'오류0회
				addSql = addSql & " and J.accFailCNT=0"
		End Select

        'cj가격 < 10x10 가격
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.cjmallPrice is Not Null and J.cjmallPrice < i.sellcash"
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.cjmallPrice is Not Null and i.sellcash <> J.cjmallPrice "
		End If

		'cj판매 10x10 품절
		If (FRectCjmallYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.cjmallSellyn='Y'"
		End If

		'CJ품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectCjmallNo10x10Yes <> "" Then
			addSql = addSql & " and (J.cjmallSellyn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'수정요망상품보기(최종업데이트일 기준)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.cjmallLastUpdate < i.lastupdate "
		End If

		'스케줄링에서 사용 실패횟수 제한
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'스케줄링에서 사용 라스트업데이트 기준 수정
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.cjmallLastUpdate, 10) "
		End If

		'########################################################    리스트 갯수 시작 ########################################################
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_cjmall_regitem as J on J.itemid = i.itemid "
		Else
			sqlStr = sqlStr & "	JOIN db_item.dbo.tbl_cjmall_regitem as J on J.itemid = i.itemid "
		End If
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small and ct.infodiv = PD.infodiv "
		sqlStr = sqlStr & " WHERE 1 = 1 and isnull(uc.userid, '') <> '' "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.cjmallPrdNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or i.deliveryType = 7 "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
				If FRectOPTCntEqual = "Y" Then		'스케줄링에서 사용
					sqlStr = sqlStr & " and i.optioncnt = J.regedoptcnt "
				End If
			End If
		Else
			sqlStr = sqlStr & " and i.isusing='Y' " & VBCRLF
			sqlStr = sqlStr & " and i.basicimage is not null " & VBCRLF
			sqlStr = sqlStr & " and i.itemdiv<50 " & VBCRLF  '''and i.itemdiv<>'08'
			sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
			sqlStr = sqlStr & " and i.cate_large<>'' " & VBCRLF
			sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'주문제작
			End If
			sqlStr = sqlStr & "	and uc.isExtUsing='Y'"
			sqlStr = sqlStr & " and not (i.deliverytype='9' and uc.defaultfreeBeasongLimit < 10000)"	'조건배송이며 10000원 미만 제외
			sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//제휴몰 판매만 허용
			sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//착불배송 상품 제거
			sqlStr = sqlStr & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"	'//조건배송 10000원 이상

			'20130514 채현아 주임 요청 카테고리 뺴기
			'sqlStr = sqlStr & "	and i.cate_large <> '080'"
			'sqlStr = sqlStr & "	and i.cate_large <> '090'"
			'sqlStr = sqlStr & "	and i.cate_large <> '070'"
			'sqlStr = sqlStr & "	and i.cate_large <> '100'"
			'sqlStr = sqlStr & "	and i.cate_large <> '075'"
			'2013-12-31 채현아 주임 요청 뷰티카테고리 중 일부카테고리만 오픈 ==> 뷰티-다이어트-운동기구, 뷰티-다이어트-체중계/만보기, 뷰티-그루밍 뷰티-기타, 뷰티-뷰티기기-헤어기기 시작
			sqlStr = sqlStr & " and (i.cate_large + i.cate_mid not in ('075001', '075002', '075003', '075004', '075005', '075006', '075009', '075010', '075012', '075013', '075016', '075018') ) "
			sqlStr = sqlStr & " and (i.cate_large + i.cate_mid + i.cate_small not in ('075020001', '075020005', '075020006', '075020007', '075020008', '075021001', '075021002', '075021003', '075014001', '075014002') ) "
			'2013-12-31 채현아 주임 요청 뷰티카테고리 중 일부카테고리만 오픈 ==> 뷰티-다이어트-운동기구, 뷰티-다이어트-체중계/만보기, 뷰티-그루밍 뷰티-기타, 뷰티-뷰티기기-헤어기기 끝
			'2014-11-14 채현아 대리 요청 감성채널/카메라 오픈
			'sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110010')"
			sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110030')"
			sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110040')"
			sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110060')"
			sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110050')"
			'20130514 채현아 주임 요청 카테고리 뺴기 끝
		End If
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If


		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt, c.mapCnt "
		sqlStr = sqlStr & "	, J.cjmallRegdate, J.cjmallLastUpdate, J.cjmallPrdNo, J.cjmallPrice, J.cjmallSellYn, J.regUserid, IsNULL(J.cjmallStatCd,-9) as cjmallStatCd "
		sqlStr = sqlStr & "	, J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr, PD.infodiv, PD.cdmKey, PD.cddkey, UC.defaultfreeBeasongLimit, i.itemdiv "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_cjmall_regitem as J on J.itemid = i.itemid "
		Else
			sqlStr = sqlStr & "	JOIN db_item.dbo.tbl_cjmall_regitem as J on J.itemid = i.itemid "
		End If
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small and ct.infodiv = PD.infodiv "
		sqlStr = sqlStr & " WHERE 1 = 1 and isnull(uc.userid, '') <> '' "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then	'// 미등록도 아니고 등록실패도 아니면 조건 없음
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.cjmallPrdNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or i.deliveryType = 7 "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
				If FRectOPTCntEqual = "Y" Then		'스케줄링에서 사용
					sqlStr = sqlStr & " and i.optioncnt = J.regedoptcnt "
				End If
			End If
		Else
			sqlStr = sqlStr & " and i.isusing='Y' " & VBCRLF
			sqlStr = sqlStr & " and i.basicimage is not null " & VBCRLF
			sqlStr = sqlStr & " and i.itemdiv<50 " & VBCRLF  '''and i.itemdiv<>'08'
			sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
			sqlStr = sqlStr & " and i.cate_large<>'' " & VBCRLF
			sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'주문제작
			End If
			sqlStr = sqlStr & "	and uc.isExtUsing='Y'"
			sqlStr = sqlStr & " and not (i.deliverytype='9' and uc.defaultfreeBeasongLimit < 10000)"	'조건배송이며 10000원 미만 제외
			sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//제휴몰 판매만 허용
			sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//착불배송 상품 제거
			sqlStr = sqlStr & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"	'//조건배송 10000원 이상

			'20130514 채현아 주임 요청 카테고리 뺴기
			'sqlStr = sqlStr & "	and i.cate_large <> '080'"
			'sqlStr = sqlStr & "	and i.cate_large <> '090'"
			'sqlStr = sqlStr & "	and i.cate_large <> '070'"
			'sqlStr = sqlStr & "	and i.cate_large <> '100'"
			'sqlStr = sqlStr & "	and i.cate_large <> '075'"
			'2013-12-31 채현아 주임 요청 뷰티카테고리 중 일부카테고리만 오픈 ==> 뷰티-다이어트-운동기구, 뷰티-다이어트-체중계/만보기, 뷰티-그루밍 뷰티-기타, 뷰티-뷰티기기-헤어기기 시작
			sqlStr = sqlStr & " and (i.cate_large + i.cate_mid not in ('075001', '075002', '075003', '075004', '075005', '075006', '075009', '075010', '075012', '075013', '075016', '075018') ) "
			sqlStr = sqlStr & " and (i.cate_large + i.cate_mid + i.cate_small not in ('075020001', '075020005', '075020006', '075020007', '075020008', '075021001', '075021002', '075021003', '075014001', '075014002') ) "
			'2013-12-31 채현아 주임 요청 뷰티카테고리 중 일부카테고리만 오픈 ==> 뷰티-다이어트-운동기구, 뷰티-다이어트-체중계/만보기, 뷰티-그루밍 뷰티-기타, 뷰티-뷰티기기-헤어기기 끝
			'2014-11-14 채현아 대리 요청 감성채널/카메라 오픈
			'sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110010')"
			sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110030')"
			sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110040')"
			sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110060')"
			sqlStr = sqlStr & "	and (i.cate_large + i.cate_mid <> '110050')"
			'20130514 채현아 주임 요청 카테고리 뺴기 끝
		End If
		sqlStr = sqlStr & addSql

		If FRectExtNotReg = "M" Then
			sqlStr = sqlStr & " ORDER BY i.itemid DESC"
		ElseIf FRectIsReged = "N" Then
			sqlStr = sqlStr & " ORDER BY i.itemid DESC"
		Else
			IF (FRectOrdType = "B") Then
				sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC"
			ElseIf (FRectOrdType = "BM") Then
				sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.itemid DESC"
			ElseIf (FRectOrdType = "PM") Then
				sqlStr = sqlStr & " ORDER BY J.lastPriceCheckDate ASC, J.cjmallLastupdate ASC"
			ElseIf (FRectOrdType = "LU") Then
				sqlStr = sqlStr & " ORDER BY i.lastupdate DESC, i.itemscore DESC, i.itemid DESC "
			Else
				sqlStr = sqlStr & " ORDER BY J.itemid DESC"
		    End If
	    End If

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do Until rsget.EOF
				Set FItemList(i) = new cjmallItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= rsget("itemname")
					FItemList(i).FsmallImage		= rsget("smallImage")
				If Not(FItemList(i).FsmallImage = "" OR isNull(FItemList(i).FsmallImage)) Then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
				Else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				End If
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).Fsellcash			= rsget("sellcash")
					FItemList(i).Fbuycash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).Fsaleyn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).Fdeliverytype		= rsget("deliverytype")
					FItemList(i).FoptionCnt			= rsget("optionCnt")
					FItemList(i).FcjmallRegdate		= rsget("cjmallRegdate")
					FItemList(i).FcjmallLastUpdate	= rsget("cjmallLastUpdate")
					FItemList(i).FcjmallPrdNo		= rsget("cjmallPrdNo")
					FItemList(i).FcjmallPrice		= rsget("cjmallPrice")
					FItemList(i).FcjmallSellYn		= rsget("cjmallSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FcjmallStatCd		= rsget("cjmallStatCd")
					FItemList(i).FregedOptCnt		= rsget("regedOptCnt")
					FItemList(i).FrctSellCNT		= rsget("rctSellCNT")
					FItemList(i).FaccFailCNT		= rsget("accFailCNT")
					FItemList(i).FlastErrStr		= rsget("lastErrStr")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
					FItemList(i).Finfodiv			= rsget("infodiv")
					FItemList(i).FcdmKey			= rsget("cdmKey")
					FItemList(i).FcddKey			= rsget("cddKey")
					FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
					FItemList(i).Fitemdiv 			= rsget("itemdiv")
				i = i + 1
				rsget.MoveNext
			Loop
		End If
		rsget.Close
	End Sub

	'등록되지 말아야 될 상품..
	Public Sub getCjmallreqExpireItemList
		Dim sqlStr, addSql, i

		'스케줄링에서 사용 라스트업데이트 기준 수정
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.cjmallLastUpdate, 10) "
		End If

        If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End if

		'텐바이텐 상품번호 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		''2013/05/29 추가
		If (FRectInfoDiv <> "") Then
			If (FRectInfoDiv = "YY") then
				addSql = addSql & " and isNULL(ct.infodiv,'')<>''"
			Elseif (FRectInfoDiv = "NN") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')=''"
			Else
				addSql = addSql & " and ct.infodiv='"&FRectInfoDiv&"'"
			End if
		End If

		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.cjmallSellYn <> 'X'"
			Else
				addSql = addSql & " and J.cjmallSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_item.dbo.tbl_cjmall_regitem as J on J.itemid = i.itemid and J.cjmallprdNo is not null " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_item.dbo.tbl_item_contents as t on i.itemid = t.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small and t.infodiv = PD.infodiv "
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리
		sqlStr = sqlStr & "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & "     or i.deliverytype in ('7') "
        sqlStr = sqlStr & "     or ((i.deliveryType=9) and (i.sellcash<10000))"
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr & "		or uc.isExtUsing='N'"
        sqlStr = sqlStr & "	)"
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
		sqlStr = sqlStr & "     where getdate() between stdt and eddt"
        sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
        sqlStr = sqlStr & "     and linkgbn='donotEdit'"
        sqlStr = sqlStr & " )"
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage " & VBCRLF
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash " & VBCRLF
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt, c.mapCnt " & VBCRLF
		sqlStr = sqlStr & "	, J.cjmallRegdate, J.cjmallLastUpdate, J.cjmallPrdNo, J.cjmallPrice, J.cjmallSellYn, J.regUserid, IsNULL(J.cjmallStatCd,-9) as cjmallStatCd  " & VBCRLF
		sqlStr = sqlStr & "	, J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr, PD.infodiv, PD.cdmKey, PD.cddkey, UC.defaultfreeBeasongLimit " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_item.dbo.tbl_cjmall_regitem as J on J.itemid = i.itemid and J.cjmallprdNo is not null " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_item.dbo.tbl_item_contents as t on i.itemid = t.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small and t.infodiv = PD.infodiv "
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리
		sqlStr = sqlStr & "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & "     or i.deliverytype in ('7') "
        sqlStr = sqlStr & "     or ((i.deliveryType=9) and (i.sellcash<10000))"
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr & "		or uc.isExtUsing='N'"
        sqlStr = sqlStr & "	)"
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
		sqlStr = sqlStr & "     where getdate() between stdt and eddt"
        sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
        sqlStr = sqlStr & "     and linkgbn='donotEdit'"
        sqlStr = sqlStr & " )"
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " order by J.regdate desc, i.itemid desc "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				Set FItemList(i) = new cjmallItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= rsget("itemname")
					FItemList(i).FsmallImage		= rsget("smallImage")
				If Not(FItemList(i).FsmallImage = "" OR isNull(FItemList(i).FsmallImage)) Then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
				Else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				End If
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).Fsellcash			= rsget("sellcash")
					FItemList(i).Fbuycash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).Fsaleyn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).Fdeliverytype		= rsget("deliverytype")
					FItemList(i).FoptionCnt			= rsget("optionCnt")
					FItemList(i).FcjmallRegdate		= rsget("cjmallRegdate")
					FItemList(i).FcjmallLastUpdate	= rsget("cjmallLastUpdate")
					FItemList(i).FcjmallPrdNo		= rsget("cjmallPrdNo")
					FItemList(i).FcjmallPrice		= rsget("cjmallPrice")
					FItemList(i).FcjmallSellYn		= rsget("cjmallSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FcjmallStatCd		= rsget("cjmallStatCd")
					FItemList(i).FregedOptCnt		= rsget("regedOptCnt")
					FItemList(i).FrctSellCNT		= rsget("rctSellCNT")
					FItemList(i).FaccFailCNT		= rsget("accFailCNT")
					FItemList(i).FlastErrStr		= rsget("lastErrStr")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
					FItemList(i).Finfodiv			= rsget("infodiv")
					FItemList(i).FcdmKey			= rsget("cdmKey")
					FItemList(i).FcddKey			= rsget("cddKey")
					FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	End Sub

	'// 미등록 상품 목록(등록용)
	Public Sub getCJMallNotRegItemList
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			''' 옵션 추가금액 있는경우 등록 불가. //옵션 전체 품절인 경우 등록 불가.
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & "	SELECT itemid FROM ("
            addSql = addSql & "     SELECT itemid"
            addSql = addSql & " 	,count(*) as optCNT"
			addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	FROM db_AppWish.dbo.tbl_item_option"
            addSql = addSql & " 	WHERE itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and isusing='Y'"
            addSql = addSql & " 	GROUP BY itemid"
            addSql = addSql & " ) T"
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"
		End If

		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, '"&CitemGbnKey&"' as itemGbnKey "
		strSql = strSql & "	, isNULL(R.cjmallStatCD,-9) as cjmallStatCD "
		strSql = strSql & "	, UC.socname_kor, isnull(PD.cdmkey, '') as cdmkey, isnull(PD.cddkey, '') as cddkey "
		strSql = strSql & " FROM db_AppWish.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_AppWish.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_outmall.dbo.tbl_cjmall_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_cjmall_regitem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_AppWish.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small and c.infodiv = PD.infodiv "
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		ELSE
		    strSql = strSql & " and (i.deliveryType<>9)"
	    END IF
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "						'플라워/화물배송
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)>=" & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_outmall].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_outmall].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_item.dbo.tbl_cjmall_regitem WHERE cjmallStatCD >= 3) "	''등록완료이상은 등록안됨.										'롯데등록상품 제외
		strSql = strSql & " and cm.mapCnt is Not Null "
		strSql = strSql & "		"	& addSql											'카테고리 매칭 상품만
'rw strSql
'response.end
		rsCTget.Open strSql,dbCTget,1
		FResultCount = rsCTget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsCTget.EOF Then
			Do until rsCTget.EOF
				Set FItemList(i) = new cjmallItem
					FItemList(i).Fitemid			= rsCTget("itemid")
					FItemList(i).FtenCateLarge		= rsCTget("cate_large")
					FItemList(i).FtenCateMid		= rsCTget("cate_mid")
					FItemList(i).FtenCateSmall		= rsCTget("cate_small")
					FItemList(i).Fitemname			= db2html(rsCTget("itemname"))
					FItemList(i).FitemDiv			= rsCTget("itemdiv")
					FItemList(i).FsmallImage		= rsCTget("smallImage")
					FItemList(i).Fmakerid			= rsCTget("makerid")
					FItemList(i).Fregdate			= rsCTget("regdate")
					FItemList(i).FlastUpdate		= rsCTget("lastUpdate")
					FItemList(i).ForgPrice			= rsCTget("orgPrice")
					FItemList(i).ForgSuplyCash		= rsCTget("orgSuplyCash")
					FItemList(i).FSellCash			= rsCTget("sellcash")
					FItemList(i).FBuyCash			= rsCTget("buycash")
					FItemList(i).FsellYn			= rsCTget("sellYn")
					FItemList(i).FsaleYn			= rsCTget("sailyn")
					FItemList(i).FisUsing			= rsCTget("isusing")
					FItemList(i).FLimitYn			= rsCTget("LimitYn")
					FItemList(i).FLimitNo			= rsCTget("LimitNo")
					FItemList(i).FLimitSold			= rsCTget("LimitSold")
					FItemList(i).Fkeywords			= rsCTget("keywords")
					FItemList(i).Fvatinclude        = rsCTget("vatinclude")
					FItemList(i).ForderComment		= db2html(rsCTget("ordercomment"))
					FItemList(i).FoptionCnt			= rsCTget("optionCnt")
					FItemList(i).FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("basicImage")
					FItemList(i).FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("mainimage")
					FItemList(i).FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("mainimage2")
					FItemList(i).Fsourcearea		= db2html(rsCTget("sourcearea"))
					FItemList(i).Fmakername			= db2html(rsCTget("makername"))
					FItemList(i).FUsingHTML			= rsCTget("usingHTML")
					FItemList(i).Fitemcontent		= db2html(rsCTget("itemcontent"))
					FItemList(i).FitemGbnKey        = rsCTget("itemGbnKey")
					FItemList(i).FcjmallStatCD		= rsCTget("cjmallStatCD")
					FItemList(i).FRectMode			= FRectMode
					FItemList(i).Fdeliverfixday		= rsCTget("deliverfixday")
					FItemList(i).Fdeliverytype		= rsCTget("deliverytype")
					FItemList(i).Fsocname_kor		= rsCTget("socname_kor")
					FItemList(i).Fcdmkey			= rsCTget("cdmkey")
					FItemList(i).Fcddkey			= rsCTget("cddkey")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub getCjmallEditedItemList
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		ElseIf FRectNotJehyu="Y" Then
			'제휴몰 상품이 아닌것
			addSql = " and i.isExtUsing='N' "
		Else
			'수정된 상품만
			addSql = " and m.cjmallLastUpdate < i.lastupdate"
		End If

        ''//연동 제외상품
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & " 	SELECT itemid FROM db_outmall.dbo.tbl_jaehyumall_not_edit_itemid"
        addSql = addSql & " 	WHERE stDt < getdate()"
        addSql = addSql & " 	and edDt > getdate()"
        addSql = addSql & " 	and mallgubun = '"&CMALLNAME&"'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, m.cjmallPrdNo, m.cjmallSellYn, m.accFailCnt, m.lastErrStr, UC.socname_kor "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
'		strSql = strSql & "		or i.itemdiv = '06' or i.itemdiv = '16' "
		strSql = strSql & " 	or i.deliveryType = '7' "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_outmall].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_outmall].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_AppWish.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_AppWish.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_cjmall_regitem as m on i.itemid = m.itemid "
		''If (FRectMatchCateNotCheck<>"on") Then
		IF (FRectMatchCate="Y") THEN '' eastone 수정 2013/09/01
			strSql = strSql & " INNER JOIN  (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt FROM db_outmall.dbo.tbl_cjmall_cate_mapping GROUP BY tenCateLarge, tenCateMid, tenCateSmall ) as cm "
			strSql = strSql & " 	on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
			strSql = strSql & " INNER JOIN  (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as PmapCnt FROM db_etcmall.dbo.tbl_cjMall_prdDiv_mapping  GROUP BY tenCateLarge, tenCateMid, tenCateSmall ) as Pm "
			strSql = strSql & " 	on Pm.tenCateLarge = i.cate_large and Pm.tenCateMid = i.cate_mid and Pm.tenCateSmall = i.cate_small "
    	End If
		strSql = strSql & " LEFT JOIN db_AppWish.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " WHERE 1 = 1 " & addSql
		strSql = strSql & " and m.cjmallPrdNo is Not Null "									'#등록 상품만
		rsCTget.Open strSql,dbCTget,1
		FResultCount = rsCTget.RecordCount
		Redim preserve FItemList(FResultCount)
		i=0
		If not rsCTget.EOF Then
			Do until rsCTget.EOF
				Set FItemList(i) = new CjmallItem
					FItemList(i).Fitemid			= rsCTget("itemid")
					FItemList(i).FtenCateLarge		= rsCTget("cate_large")
					FItemList(i).FtenCateMid		= rsCTget("cate_mid")
					FItemList(i).FtenCateSmall		= rsCTget("cate_small")
					FItemList(i).Fitemname			= db2html(rsCTget("itemname"))
					FItemList(i).FitemDiv			= rsCTget("itemdiv")
					FItemList(i).FsmallImage		= rsCTget("smallImage")
					FItemList(i).Fmakerid			= rsCTget("makerid")
					FItemList(i).Fregdate			= rsCTget("regdate")
					FItemList(i).FlastUpdate		= rsCTget("lastUpdate")
					FItemList(i).ForgPrice			= rsCTget("orgPrice")
					FItemList(i).ForgSuplyCash		= rsCTget("orgSuplyCash")
					FItemList(i).FSellCash			= rsCTget("sellcash")
					FItemList(i).FBuyCash			= rsCTget("buycash")
					FItemList(i).FsellYn			= rsCTget("sellYn")
					FItemList(i).FsaleYn			= rsCTget("sailyn")
					FItemList(i).FisUsing			= rsCTget("isusing")
					FItemList(i).FLimitYn			= rsCTget("LimitYn")
					FItemList(i).FLimitNo			= rsCTget("LimitNo")
					FItemList(i).FLimitSold			= rsCTget("LimitSold")
					FItemList(i).Fkeywords			= rsCTget("keywords")
					FItemList(i).ForderComment		= db2html(rsCTget("ordercomment"))
					FItemList(i).FoptionCnt			= rsCTget("optionCnt")
					FItemList(i).FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("basicImage")
					FItemList(i).FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("mainimage")
					FItemList(i).FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("mainimage2")
					FItemList(i).Fsourcearea		= db2html(rsCTget("sourcearea"))
					FItemList(i).Fmakername			= db2html(rsCTget("makername"))
					FItemList(i).FUsingHTML			= rsCTget("usingHTML")
					FItemList(i).Fitemcontent		= db2html(rsCTget("itemcontent"))
					FItemList(i).FcjmallPrdNo		= rsCTget("cjmallPrdNo")
					FItemList(i).FcjmallSellYn		= rsCTget("cjmallSellYn")
	                FItemList(i).Fvatinclude        = rsCTget("vatinclude")
	                FItemList(i).Fsocname_kor		= rsCTget("socname_kor")
					FItemList(i).FmaySoldOut    	= rsCTget("maySoldOut")
					FItemList(i).FaccFailCnt    	= rsCTget("accFailCnt")
					FItemList(i).FlastErrStr    	= rsCTget("lastErrStr")
					i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
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
%>