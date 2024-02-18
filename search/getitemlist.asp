<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.charset = "utf-8" %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/util/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/search/searchcls.asp" -->
<!-- #include virtual="/search/iteminfoCls.asp" -->
<%

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("13.125.145.40","13.125.12.181","52.79.73.145","61.252.133.88","192.168.1.70","61.252.133.81","192.168.1.81","192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

function getBrandNameKrBySearchKeyList(iitemid,ibrandname,ikeylist)
    dim orginKey : orginKey=ikeylist
    dim pos1 , pos2, ikeylistArr
    dim ret : ret = ""
    
    getBrandNameKrBySearchKeyList = ret
    pos1 = InStr(ikeylist,CStr(iitemid))
    pos2 = InStr(ikeylist,CStr(ibrandname))
    if (pos1<1) or (pos2<1) then 
        exit function
    end if
    
    ikeylist = Trim(LEFT(ikeylist,pos1-1))
    ikeylist = TRIM(Mid(ikeylist,pos2+LEN(ibrandname),512))
    
    pos2 = InStr(ikeylist,CStr(ibrandname))
    if (pos2>0) then 
        ikeylist = TRIM(Mid(ikeylist,pos2+LEN(ibrandname),512))
    end if
    
    
    if (pos2>0) then 
        pos2 = InStr(ikeylist,CStr(ibrandname))
        if (pos2>0) then 
            ikeylist = TRIM(Mid(ikeylist,pos2+LEN(ibrandname),512))
        end if
    end if
    
    getBrandNameKrBySearchKeyList = ikeylist
    exit function
    
    'ikeylistArr = split(ikeylist ," ")
    
    'if isArray(ikeylistArr) then
    '    getBrandNameKrBySearchKeyList = ikeylistArr(Ubound(ikeylistArr))
    'end if
end function

dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    response.write ref
    response.end
end if

dim reBuf
dim q : q = request("q")
q = RepWord(q,"[^ㄱ-ㅎㅏ-ㅣ가-힣a-zA-Z0-9.&%\-\_\s]","")

if (LEN(q)<1) then
    response.write "no param"
    response.end    
end if

dim ListDiv : ListDiv = "search"
dim SellScope : SellScope = "Y" ' 판매여부 Y:판매상품만  ELSE 일시품절 포함.
dim SortMet : SortMet="be" ''"be"  2018/07/18 be로 수정

dim CurrPage : CurrPage=1
dim PageSize : PageSize=30  ''2018/07/18
dim ScrollCount : ScrollCount=10
'// 총 검색수 산출
dim oTotalCnt
set oTotalCnt = new SearchItemCls
oTotalCnt.FRectSearchTxt = q
''oTotalCnt.FRectExceptText = ExceptText
''oTotalCnt.FRectSearchItemDiv = SearchItemDiv
''oTotalCnt.FRectSearchCateDep = SearchCateDep
oTotalCnt.FListDiv = ListDiv
oTotalCnt.FSellScope=SellScope
oTotalCnt.getTotalCount

'// 상품검색
dim oDoc,i
set oDoc = new SearchItemCls
oDoc.FCurrPage = CurrPage
oDoc.FPageSize = PageSize
oDoc.FScrollCount = ScrollCount

oDoc.FRectSearchTxt = q
oDoc.FRectSortMethod	= SortMet
oDoc.FListDiv = ListDiv
oDoc.FSellScope=SellScope

'oDoc.FRectPrevSearchTxt = PrevSearchText
'oDoc.FRectExceptText = ExceptText
'oDoc.FRectSearchFlag = searchFlag
'oDoc.FRectSearchItemDiv = SearchItemDiv
'oDoc.FRectSearchCateDep = SearchCateDep
'oDoc.FRectCateCode	= dispCate
'oDoc.FRectMakerid	= makerid
'oDoc.FminPrice	= minPrice
'oDoc.FmaxPrice	= maxPrice
'oDoc.FdeliType	= deliType
'oDoc.FLogsAccept = LogsAccept
'oDoc.FRectColsSize = ColsSize
'oDoc.FcolorCode = colorCD
'oDoc.FstyleCd = styleCd
'oDoc.FattribCd = attribCd
'oDoc.FarrCate=arrCate

oDoc.getSearchList

reBuf = oTotalCnt.FTotalcount&vbCRLF

'' 상품코드, 상품명, 브랜드명, 한글 브랜드명
IF oDoc.FResultCount >0 then
    For i=0 To oDoc.FResultCount -1
        reBuf = reBuf &oDoc.FITemList(i).FItemid&"||"
        reBuf = reBuf &oDoc.FITemList(i).FItemName&"||"
        reBuf = reBuf &oDoc.FITemList(i).FBrandName&"||"
        reBuf = reBuf &getBrandNameKrBySearchKeyList(oDoc.FITemList(i).FItemid,oDoc.FITemList(i).FBrandName,oDoc.FITemList(i).FKeyWords)&"||"
        reBuf = reBuf &oDoc.FItemList(i).FImageBasic&"||"
        reBuf = reBuf &oDoc.FItemList(i).FAddimage&"||"
        reBuf = reBuf &oDoc.FITemList(i).FKeyWords&vbCRLF
    Next
end if

SET oDoc=Nothing
SET oTotalCnt=Nothing


response.write reBuf
%>
