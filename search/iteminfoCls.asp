<%
'#######################################################
'	History	: 김정인 생성
'			  2008.03.18 정윤정 수정 - 클래스 분리
'			  2008.04.13 한용민 추가
'             2008.08.27 서동석 업체 착불 배송 관련 추가
'	Description :상품관련 함수 모음
'#######################################################

'#=========================================#
'# 카테고리 상품 아이템                    #
'#=========================================#

CLASS CCategoryPrdItem

	'// 필수 변수  //

	dim FItemID
	dim FItemName
	dim FSellcash
	dim FOrgPrice
	dim fEval_excludeyn
	dim FNewitem

	dim FMakerID
	dim FBrandName
	dim FBrandName_kor
	dim FBrandLogo
	dim FBrandUsing
	dim FisBestBrand
	dim FUserDiv

	dim FItemDiv
	dim FMakerName
	dim FOrgMakerID

	dim FMileage
	dim FSourceArea
	dim FDeliverytype

	dim FcdL
	dim FcdM
	dim FcdS
	dim FcateCode
	dim FCateName
	dim FcateCd1
	dim FcateCd2
	dim FcateCd3
	dim FcateDepth
	dim FarrCateCd

	dim Freviewcnt


	dim FcolorCode
	dim FcolorName

	dim FLimitNo
	dim FLimitSold
	dim fsailprice
	dim FImageBasic
	dim FImageBasic600		'600px이미지
	dim FImageBasic1000		'1000px이미지
	dim FImageMask
	dim FImageMask1000		'1000px이미지
	dim FImageList
	dim FImageList120
	dim FImageSmall
	dim FImageBasicIcon
	dim FImageMaskIcon
	dim FImageIcon1	'신상품리스트, 할인리스트에서 사용(200x200)
	dim FImageIcon2
	dim FImageIcon3
	dim FImageIcon4
	dim FImageIcon5
	dim FIcon1Image
	dim FIcon2Image

	'// 텐텐 기본 이미지 추가(2015.01.21 원승현)
	Dim Ftentenimage
	Dim Ftentenimage50
	Dim Ftentenimage200
	Dim Ftentenimage400
	Dim Ftentenimage600
	Dim Ftentenimage1000

	'// 상품상세설명 동영상 추가(2016.02.17 원승현)
	Dim FvideoUrl
	Dim FvideoWidth
	Dim FvideoHeight
	Dim Fvideogubun
	Dim FvideoType
	Dim FvideoFullUrl


	dim FOrderComment
	dim Fdeliverarea
	dim FItemSource
	dim FItemSize
	dim FItemWeight
	dim FdeliverOverseas

	dim Fkeywords
	dim FUsingHTML
	dim FItemContent

	dim Fisusing
	dim FStreetUsing

	dim FRegDate

	dim FReipgodate
	dim FSpecialbrand


	dim Fdgncomment
	dim FDesignerComment

	dim FLimitYn
	dim FSellYn
	dim FItemScore

	dim Fitemgubun

	dim FSaleYn
	dim FTenOnlyYn		'텐바이텐 독점상품여부(2011.04.14)

	dim FEvalcnt
	dim FEvalcnt_Photo
	dim FfavCount
	dim FQnaCnt
	dim FOptionCnt
	dim FAvgDlvDate

	dim FAddimageGubun
	dim FAddimageSmall
	dim FAddImageType
	dim FAddimage
	dim FAddimage600
	dim FAddimage1000
	dim FIsExistAddimg

	dim Ffreeprizeyn '?

	dim FReipgoitemyn
	dim FSpecialUserItem

	dim Fitemcouponyn
	dim FItemCouponType
	dim FItemCouponValue
	dim FItemCouponExpire
	dim FCurrItemCouponIdx

	dim FAvailPayType               '결제 방식 지정 0-일반 ,1-실시간(선착순)
	dim FDefaultFreeBeasongLimit    '업체 개별배송시 배송비 무료 적용값
	dim FDefaultDeliverPay		    ' 업체 개별배송시 배송비
	dim FRequireMakeDay				'주문제작상품의 제작 소요일(2011.04.14)

	Dim FsafetyYN		'안전인증대상
	Dim FsafetyDiv		'안전인증구분 '10 ~ 50
	Dim FsafetyNum	'안전인증번호
	Dim FcertNum	'신안전인증번호
	Dim FcertDiv	'안전인증구분텍스트
	Dim FcertUid	'안전인증구분 인증기관 고유번호

	public FPoints
	public FPoint_fun
	public FPoint_dgn
	public FPoint_prc
	public FPoint_stf
	public Fuserid
	public Fcontents
	public FImageMain
	public FImageMain2			'상품설명2 이미지 추가(2011.04.14)
	public FImageMain3			'상품설명3 이미지 추가(2013.07.31)
	public FlinkURL

	public FCurrRank
	public FLastRank

	public FPojangOk			'선물포장 가능 여부

	public FBRWriteRegdate		'베스트리뷰용
	public FUseGood
	public FUseETC

	public FplusSalePro			''세트구매 할인율.
	public FisJust1day			'Just 1day 상품 여부

	'스타일라이프용
	public FStyleCd1
	public FStyleCd1Nm
	public FStyleCd2
	public FStyleCd2Nm
	public FStyleCd3
	public FStyleCd3Nm
	public fOrderNo

	'hotcateitem 2012-04-04
	Public Fidx
	Public Fitemseq
	Public Fcdmname
	Public Fcdsname
	Public Fsailyn

	'상품상세 추가 2012-11-01
	Public FInfoname
	Public FInfoContent
	Public FinfoCode

	Public ForderMinNum
	Public ForderMaxNum

	'2013 리뉴얼 카테고리메인용
	Public FDisp
	Public Ftype
	Public Fcode
	Public Ftitle
	Public Fsubcopy
	Public Fimgurl
	Public Ficon

	'2013 popular wish
	Public FInCount
	Public FRegTime
	Public FEvaluate
	Public FMyCount
	
	'/브랜드 페이지용
	public fdetailidx
	public fmasteridx
	public fsortNo
	public Flastupdate
	public fregadminid
	public flastadminid
	public fevt_code

	'/2014 Gift
	public FtalkCnt
	public FdayCnt
	public FthemeCnt
	
	'/상품상세추가
	public FLimitDispYn
	
	public fdevice
	public Fsdate
	public Fedate

	'/2015 내 주문 상품
	public Forderserial
	public ForderDate
	public ForderOption
	public ForderOptionName
	public ForderCnt

	'브랜드 공지 추가2017-01-31 유태욱
	public FBrandNoticeGubun
	public FBrandNoticeTitle
	public FBrandNoticeText

	'/루키관련
	public Frecentsellcount

	'딜 대표상품 코드 추가
	Public FDealMasterItemID
	Public FItemOptionCnt
	
	public function IsRookieItem()
		IsRookieItem = false
		if (Not IsNewItem) then Exit function
		
		IsRookieItem = (Frecentsellcount>=20)
	end function

	public function IsStreetAvail() ' !
		IsStreetAvail = (FStreetUsing="Y") and (Fuserdiv<10)
	end function


	'// 원 판매 가격  '!
	public Function getOrgPrice()
		if FOrgPrice=0 then
			getOrgPrice = FSellCash
		else
			getOrgPrice = FOrgPrice
		end if
	end Function

	'// 세일포함 실제가격  '!
	public Function getRealPrice()

		getRealPrice = FSellCash


		if (IsSpecialUserItem()) then
			getRealPrice = getSpecialShopItemPrice(FSellCash)
		end if
	end Function

	'//상품코드  '!
	public Function FProductCode()
		 FProductCode = formatCode(FItemid)
	end Function

	'// 상품명
	public Function getCuttingItemName()
		if Len(FItemName)>18 then
			getCuttingItemName=Left(FItemName,18) + "..."
		else
			getCuttingItemName=FItemName
		end if
	end Function

	'// 상품 설명  '?
	public Function GetCuttingItemContents()
		''## 이상은 잘라버림.
		dim reStr
		reStr = LeftB(Fitemcontent,120)
		reStr = replace(reStr,"<P>","")
		reStr = replace(reStr,"<p>","")
		reStr = replace(reStr,"<br>",Chr(2))
		reStr = Left(reStr,100)
		reStr = replace(reStr,Chr(2),"&nbsp;")
		GetCuttingItemContents = reStr + "..."
	end Function

	'// 우수회원샵 상품 여부 '!
	public Function IsSpecialUserItem()
	    dim uLevel
	    uLevel = GetLoginUserLevel()
		IsSpecialUserItem = (FSpecialUserItem>0) and (uLevel>1 and uLevel<>5)
	end Function

	'// 판매종료  여부 '! '2008/07/07 추가
	public Function IsSoldOut()

		'isSoldOut = (FSellYn="N")
		IF FLimitNo<>"" and FLimitSold<>"" Then
			isSoldOut = (FSellYn<>"Y") or ((FLimitYn = "Y") and (clng(FLimitNo)-clng(FLimitSold)<1))
		Else
			isSoldOut = (FSellYn<>"Y")
		End If
	end Function

	'// 딜 판매종료  여부 '! '2017/11/17 추가
	public Function isDealSoldout() 
		isDealSoldout = (FSellYn="N")
	end Function

	'//일시품절 여부 '2008/07/07 추가 '!
	public Function isTempSoldOut()

		isTempSoldOut = (FSellYn="S")

	end Function

	'// 세일 상품 여부 '!
	public Function IsSaleItem()
	    IsSaleItem = ((FSaleYn="Y") and (FOrgPrice-FSellCash>0)) or (IsSpecialUserItem)
	end Function

	'//	한정 여부 '!
	public Function IsLimitItem()
			IsLimitItem= (FLimitYn="Y") and (FLimitDispYn="Y" or isNull(FLimitDispYn))
	end Function

	'//	한정 여부 (표시여부와 상관없는 실제 상품 한정여부)
	public Function IsLimitItemReal()
			IsLimitItemReal= (FLimitYn="Y")
	end Function

	'// 신상품 여부 '!
	public Function IsNewItem()
			IsNewItem =	(datediff("d",FRegdate,now())<= 14)
	end Function

	'// 무료 배송 쿠폰 여부 '?
	public function IsFreeBeasongCoupon()
		IsFreeBeasongCoupon = Fitemcoupontype="3"
	end function

	'// 상품 쿠폰 여부  '!
	public Function IsCouponItem()
			IsCouponItem = (FItemCouponYN="Y")
	end Function

	'// 사은품 증정 상품 여부 '?
	public Function IsGiftItem()
			IsGiftItem	= (FFreePrizeYN ="Y")
	end Function

	'// 재입고 상품 여부
	public Function isReipgoItem()
		isReipgoItem = (datediff("d",FReIpgoDate,now())<= 14)
	end Function

	'// 마일리지샵 아이템 여부 '!
	public Function IsMileShopitem()
		IsMileShopitem = (FItemDiv="82")
	end Function

	'// 텐바이텐 독점상품 여부 '!
	public Function IsTenOnlyitem()
		IsTenOnlyitem = (FTenOnlyYn="Y")
	end Function

	'// 텐바이텐 포장가능 상품 여부
	public Function IsPojangitem()
		IsPojangitem = (FPojangOk="Y" and IsTenBeasong)
	end Function

	'// 한정 상품 남은 수량 '!
	public Function FRemainCount()
		if IsSoldOut then
			FRemainCount=0
		else
			FRemainCount=(clng(FLimitNo) - clng(FLimitSold))
		end if
	End Function

	'// 상품 문의 받기 '!
	public Function IsSpecialBrand()
		IsSpecialBrand = FSpecialBrand="Y"
	End Function

	'// 할인가
	public Function getDiscountPrice()
		dim tmp

		if (FDiscountRate<>1) then
			tmp = cstr(FSellcash * FDiscountRate)
			getDiscountPrice = round(tmp / 100) * 100
		else
			getDiscountPrice = FSellcash
		end if
	end Function

	'// 할인율 '!
	public Function getSalePro()
		if FOrgprice=0 then
			getSalePro = 0 & "%"
		else
			getSalePro = CLng((FOrgPrice-getRealPrice)/FOrgPrice*100) & "%"
		end if
	end Function

	'// 쿠폰 적용가
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function

	'// 쿠폰 할인가 '?
	public Function GetCouponDiscountPrice()
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*getRealPrice/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

	end Function

	'// 상품 쿠폰 내용  '!
	public function GetCouponDiscountStr()

		Select Case Fitemcoupontype
			Case "1"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "%"
			Case "2"
				GetCouponDiscountStr = formatNumber(Fitemcouponvalue,0) + "원 할인"
			Case "3"
				GetCouponDiscountStr ="무료배송"
			Case Else
				GetCouponDiscountStr = Fitemcoupontype
		End Select

	end function


	public function GetCouponDiscountStr_new()

		Select Case Fitemcoupontype
			Case "1"
				GetCouponDiscountStr_new =CStr(Fitemcouponvalue) + "%"
			Case "2"
				GetCouponDiscountStr_new = formatNumber(Fitemcouponvalue,0) + "원 할인"
			Case "3"
				GetCouponDiscountStr_new =""
			Case Else
				GetCouponDiscountStr_new = Fitemcoupontype
		End Select

	end function


	'// 무료 배송 여부
	public Function IsFreeBeasong()
		if (getRealPrice()>=getFreeBeasongLimitByUserLevel()) then
			IsFreeBeasong = true
		else
			IsFreeBeasong = false
		end if

		if (FDeliverytype="2") or (FDeliverytype="4") or (FDeliverytype="5") or (FDeliverytype="6") then
			IsFreeBeasong = true
		end if

		''//착불 배송은 무료배송이 아님
		if (FDeliverytype="7") then
		    IsFreeBeasong = false
		end if
	end Function

	'// 해외 배송 여부(텐배 + 해외여부 + 상품무게)
	public Function IsAboardBeasong()
		if FdeliverOverseas="Y" and FItemWeight>0 and (FDeliverytype="1" or FDeliverytype="3" or FDeliverytype="4") then
			IsAboardBeasong = true
		else
			IsAboardBeasong = false
		end if
	end function

	'// 텐바이텐 배송 여부
	public Function IsTenBeasong()
		IsTenBeasong = false
		if (FDeliverytype="1" or FDeliverytype="3" or FDeliverytype="4") then
			IsTenBeasong = true
		end if
	end function

	''// 업체별 배송비 부과 상품(업체 조건 배송)
	public Function IsUpcheParticleDeliverItem()
	    IsUpcheParticleDeliverItem = (FDefaultFreeBeasongLimit>0) and (FDefaultDeliverPay>0) and (FDeliveryType="9")
	end function

	''// 업체착불 배송여부
	public Function IsUpcheReceivePayDeliverItem()
	    IsUpcheReceivePayDeliverItem = (FDeliveryType="7")
	end function

	public function getDeliverNoticsStr()
	    getDeliverNoticsStr = ""
	    if (IsUpcheParticleDeliverItem) then
	        getDeliverNoticsStr = FBrandName & "(" & FBrandName_kor & ") 제품으로만" & "<br>"
	        getDeliverNoticsStr = getDeliverNoticsStr & FormatNumber(FDefaultFreeBeasongLimit,0) & "원 이상 구매시 무료배송 됩니다."
	        getDeliverNoticsStr = getDeliverNoticsStr & "배송비(" & FormatNumber(FDefaultDeliverPay,0) & "원)"
	    elseif (IsUpcheReceivePayDeliverItem) then
	        getDeliverNoticsStr = "착불 배송비는 지역에 따라 차이가 있습니다. "
            getDeliverNoticsStr = getDeliverNoticsStr & " 상품설명의 '배송안내'를 꼭 읽어보세요." & "<br>"
	    end if
	end function

	' 사용자 등급별 무료 배송 가격  '?
	public Function getFreeBeasongLimitByUserLevel()
		dim ulevel

		''쇼핑에서는 사용자레벨에 상관없이 3만 / 업체 개별배송 5만 장바구니에서만 체크
		if (FDeliverytype="9") then
		    If (IsNumeric(FDefaultFreeBeasongLimit)) and (FDefaultFreeBeasongLimit<>0) then
		        getFreeBeasongLimitByUserLevel = FDefaultFreeBeasongLimit
		    else
		        getFreeBeasongLimitByUserLevel = 50000
		    end if
		else
		    getFreeBeasongLimitByUserLevel = 30000
		end if

	end Function

    '// 옵션 존재여부 옵션 갯수로 체크
    public function IsItemOptionExists()
        IsItemOptionExists = (FOptioncnt>0)
    end function

	'// 배송구분 : 무료배송은 따로 처리  '!
	public Function GetDeliveryName()
		Select Case FDeliverytype
			Case "1"
				GetDeliveryName="텐바이텐배송"
			Case "2"
				if FMakerid="goodovening" then
					GetDeliveryName="업체배송"
				else
					GetDeliveryName="업체무료배송"
				end if
			'Case "3"
			'		GetDeliveryName="텐바이텐배송"
			Case "4"
					GetDeliveryName="텐바이텐배송"
			Case "5"
					GetDeliveryName="업체무료배송"
			Case "6"
					GetDeliveryName="현장수령상품"
			Case "7"
				GetDeliveryName="업체착불배송"
			Case "9"
				if Not IsFreeBeasong then
					GetDeliveryName="업체조건배송"
				else
					GetDeliveryName="업체무료배송"
				end if
			Case Else
				GetDeliveryName="텐바이텐배송"
		End Select
	end Function


	'// 무이자 이미지 & 레이어  '!
	public Function getInterestFreeImg()
			if getRealPrice>=50000 then
				getInterestFreeImg="<div class='clicklayer' class='relative'>" & vbCrLf &_
									"<img class='btn img' src='http://fiximage.10x10.co.kr/web2012/product/product_desc_title03_1.png' style='cursor:pointer'/>" & vbCrLf &_
									"	<div class='layer credit-card'>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_ld.gif' /></div> <span class='black_11px_bold'>롯데카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_sh.gif' /></div> <span class='black_11px_bold'>신한카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_hd.gif' /></div> <span class='black_11px_bold'>현대카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_keb.gif' /></div> <span class='black_11px_bold'>국민카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_bc.gif' /></div> <span class='black_11px_bold'>비씨카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_ss.gif' /></div> <span class='black_11px_bold'>삼성카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf &_
									"		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_kb.gif' /></div> <span class='black_11px_bold'>국민카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf &_
									"	</div>" & vbCrLf &_
									"</div>"
				'// 2013년 1월 1일부로 모든 카드 무이자혜택 제거
				getInterestFreeImg = ""

				'//2013년 1,2월 무이자 안내
				if date()>="2013-01-07" and date()<="2013-02-28" then
					getInterestFreeImg="<div class='clicklayer' class='relative'>" & vbCrLf
					getInterestFreeImg= getInterestFreeImg & "<img class='btn img' src='http://fiximage.10x10.co.kr/web2012/product/product_desc_title03_1.png' style='cursor:pointer'/>" & vbCrLf
					getInterestFreeImg= getInterestFreeImg & "	<div class='layer credit-card' style='border:3px solid #DDD;'>" & vbCrLf
					if date()>="2013-01-07" and date()<="2013-02-17" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_ss.gif' /></div> <span class='black_11px_bold'>삼성카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf
					if date()>="2013-01-09" and date()<="2013-02-17" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_sh.gif' /></div> <span class='black_11px_bold'>신한카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf
					if date()>="2013-01-11" and date()<="2013-02-17" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_ld.gif' /></div> <span class='black_11px_bold'>롯데카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf
					if date()>="2013-01-11" and date()<="2013-02-17" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_hd.gif' /></div> <span class='black_11px_bold'>현대카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf
					if date()>="2013-01-12" and date()<="2013-02-28" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_kb.gif' /></div> <span class='black_11px_bold'>국민카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf
					if date()>="2013-01-12" and date()<="2013-02-17" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_keb.gif' /></div> <span class='black_11px_bold'>외환카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf
					if date()>="2013-01-12" and date()<="2013-02-28" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_nh.gif' /></div> <span class='black_11px_bold'>농협카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf
					if date()>="2013-02-01" and date()<="2013-02-28" then getInterestFreeImg= getInterestFreeImg & "		<div><img src='http://fiximage.10x10.co.kr/web2012/product/card_bc.gif' /></div> <span class='black_11px_bold'>비씨카드</span>&nbsp;5만원↑ / 2,3개월<br/>" & vbCrLf
					getInterestFreeImg= getInterestFreeImg & "	</div>" & vbCrLf
					getInterestFreeImg= getInterestFreeImg & "</div>"
				end if
			end if
	end Function


    ''// 세트구매 할인가격
    public function GetPLusSalePrice()
        if (FplusSalePro>0) then
            GetPLusSalePrice = getRealPrice-CLng(getRealPrice*FplusSalePro/100)
        else
            GetPLusSalePrice = getRealPrice
        end if
    end function


	public function GetLevelUpCount()

		if (FCurrRank<FLastRank) then
			GetLevelUpCount = CStr(FLastRank-FCurrRank)
		elseif (FCurrRank=FLastRank) and (FLastRank=0) then
			GetLevelUpCount = ""
		elseif (FCurrRank=FLastRank) then
			GetLevelUpCount = ""
		elseif (FCurrRank>FLastRank) and (FLastRank=0) then
			GetLevelUpCount = ""
		else
			GetLevelUpCount = CStr(FCurrRank-FLastRank)
			if FCurrRank-FLastRank>=FCurrPos then
				GetLevelUpCount = ""
			end if
		end if
	end function

	public function GetLevelUpArrow()
		if (FCurrRank<FLastRank) then
			GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2009/bestaward/award_up.gif' width='7' height='4' align='absmiddle'> <font class='verdanared'><b>" & GetLevelUpCount() & "</b></font>"
		elseif (FCurrRank=FLastRank) and (FLastRank=0) then
			GetLevelUpArrow = ""
			'##기존 GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2008/award/s_arrow_new.gif' width='9' height='5'>"
		elseif (FCurrRank=FLastRank) then
			GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2009/bestaward/award_none.gif' width='6' height='2' align='absmiddle'> <font class='eng11px00'><b>0</b></font>"
		elseif (FCurrRank>FLastRank) and (FLastRank=0) then
			GetLevelUpArrow = ""
			'##기존 GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2008/award/s_arrow_new.gif' width='9' height='5'>"
		else
			GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2009/bestaward/award_down.gif' width='7' height='4' align='absmiddle'> <font class='verdanabk'><b>" & GetLevelUpCount() & "</b></font>"
			if FCurrRank-FLastRank>=FCurrPos then
				GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2009/bestaward/award_none.gif' width='6' height='2' align='absmiddle'> <font class='eng11px00'><b>0</b></font>"
			end if
		end if
	end Function

	public function isBestRankItem()
		isBestRankItem = false
		if not(FCurrRank="" or isNull(FCurrRank)) then
			if FCurrRank<=1000 then
				isBestRankItem = true
			end if
		end if
	end function

	'// 안전인증정보 여부
	public Function IsSafetyYN()
		if FsafetyYN="Y"  then
			IsSafetyYN = true
		else
			IsSafetyYN = false
		end if
	end Function

	'// 안전인증정보 마크
	public Function IsSafetyDIV()
		if FsafetyDIV="10"  then
			IsSafetyDIV = "국가통합인증(KC마크)"
		ElseIf FsafetyDIV="20"  then
			IsSafetyDIV = "전기용품 안전인증"
		ElseIf FsafetyDIV="30"  then
			IsSafetyDIV = "KPS 안전인증 표시"
		ElseIf FsafetyDIV="40"  then
			IsSafetyDIV = "KPS 자율안전 확인 표시"
		ElseIf FsafetyDIV="50"  then
			IsSafetyDIV = "KPS 어린이 보호포장 표시"
		end if
		
		'### 실섭에 없어서 소스맞추려고 뺴놨습니다.
		'ElseIf FsafetyDIV="60"  then
		'	IsSafetyDIV = "KCC인증(구MIC인증)"
	end function
	
	
	public Function fnRealAllPrice()
		'####### 쿠폰 할인 모두 다 계산하여 1가지로 나타냄. 할인&쿠폰 중 쿠폰이 우위.
		Dim vPrice
		vPrice = FSellCash
		IF FSaleyn = "Y" AND FItemcouponyn = "Y" Then
			vPrice = GetCouponAssignPrice
		Else
			If FItemcouponyn = "Y" Then
				vPrice = GetCouponAssignPrice
			End If
		End If
		fnRealAllPrice = vPrice
	End Function

    ''여행상품 //2016/04/15 추가
    public function IsTravelItem()
        IsTravelItem = False
        if FItemDiv="18" then
			IsTravelItem = true
		end if
    end function

	Private Sub Class_Initialize()
        FplusSalePro = 0
        Frecentsellcount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

end CLASS
%>
