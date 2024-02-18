<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #Include Virtual = "/lib/util/gridFunction.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/datamart/DataMartItemsalecls.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<%
function IsVaildAdmin()
  IsVaildAdmin = false
  IsVaildAdmin = (session("ssBctId")<>"")
  IsVaildAdmin = IsVaildAdmin and ((session("ssBctDiv")<=9) or (session("ssBctDiv")=101) or (session("ssBctDiv")=111) or (session("ssBctDiv")=112) or (session("ssBctDiv")=201) or (session("ssBctDiv")=301))
end function

if (Not IsVaildAdmin) then
    response.write returnXMLResultObjArr(cmd,"N","S_ERR","세션이 종료되었습니다. "&Vbcrlf&Vbcrlf&"다시 로그인해 주세요.","")
    response.end
end if



dim i, retVal, arrColName, iobj, buf
dim cmd         : cmd = RequestCheckVar(request("cmd"),32)
dim page        : page = RequestCheckVar(request("page"),10)
dim pagesize    : pagesize = RequestCheckVar(request("pagesize"),10)
dim research    : research = request("research")

if page="" then page=1
if pagesize="" then pagesize=20

dim orderserial : orderserial = requestCheckvar(request("orderserial"),32)
dim oReport

dim ojumun
dim jumundiv    : jumundiv = RequestCheckVar(request("jumundiv"),10)
dim jumunsite   : jumunsite = RequestCheckVar(request("jumunsite"),32)
dim searchtype  : searchtype  = requestCheckVar(request("searchtype"),32)
dim searchrect  : searchrect  = requestCheckVar(request("searchrect"),32)

dim ckipkumdiv4 : ckipkumdiv4 = request("ckipkumdiv4")
dim ckipkumdiv2 : ckipkumdiv2 = request("ckipkumdiv2")

if (cmd="simpleOrderList") then
    set ojumun = new CJumunMaster
    if (jumundiv="flowers") then
    	ojumun.FRectIsFlower = "Y"
    elseif (jumundiv="minus") then
        ojumun.FRectIsMinus = "Y"
    elseif (jumundiv="foreign") then
        ojumun.FRectIsForeign = "Y"
    elseif (jumundiv="military") then
        ojumun.FRectIsMilitary = "Y"
    end if
    
    ojumun.FRectRegStart = RequestCheckVar(request("stdt"),10)
    ojumun.FRectRegEnd = RequestCheckVar(request("eddt"),10)
    
    if request("ckdelsearch")<>"on" then
    	ojumun.FRectDelNoSearch="on"
    end if
    
    
    if searchtype="01" then
    	ojumun.FRectBuyname = searchrect
    elseif searchtype="02" then
    	ojumun.FRectReqName = searchrect
    elseif searchtype="03" then
    	ojumun.FRectUserID = searchrect
    elseif searchtype="04" then
    	ojumun.FRectIpkumName = searchrect
    elseif searchtype="06" then
    	ojumun.FRectSubTotalPrice = searchrect
    end if
    
    ojumun.FPageSize = pagesize
    ojumun.FRectIpkumDiv4 = ckipkumdiv4
    ojumun.FRectIpkumDiv2 = ckipkumdiv2
    ojumun.FRectOrderSerial = orderserial
    
    ojumun.FCurrPage = page
    ojumun.SearchJumunList

    arrColName = Array("ORDERSERIAL","COMPANY","SITE","USERID","BUYNAME","REQNAME","SUBTOTALPRICE","TOTALSUM","ACCOUNTDIVNAME","IPKUMDIVNAME","CANCELYN","REGDATE")
    
    set iobj = new CTnGridData
    
    iobj.FPageSize  = ojumun.FPageSize
    iobj.FCurrPage  = ojumun.FCurrPage
    iobj.FTotalCount = ojumun.FTotalCount
    iobj.FTotalSum  = ojumun.FSubTotal
    iobj.FAvgSum    = CLNG(ojumun.FAvgTotal)
    
    
    for i=0 to ojumun.FResultCount-1
        iobj.AddData i,ojumun.FMasterItemList(i).FOrderSerial,"ORDERSERIAL"  
        IF (ojumun.FMasterItemList(i).FDlvcountryCode="ZZ") then
            iobj.AddData i,"군","DLVCODE"  
        Else
            iobj.AddData i,ojumun.FMasterItemList(i).FDlvcountryCode,"DLVCODE"  
        End IF
        iobj.AddData i,ojumun.FMasterItemList(i).FSitename,"SITE"  
        if ojumun.FMasterItemList(i).FSitename<>"10x10" then
		    iobj.AddData i,ojumun.FMasterItemList(i).FAuthCode,"USERID" 
		else
		    iobj.AddData i,ojumun.FMasterItemList(i).FUserID,"USERID" 
		    'iobj.AddData i,"<a href=""?searchfield=userid&userid="& ojumun.FMasterItemList(i).FUserID&"""><font color=""" & ojumun.FMasterItemList(i).GetUserLevelColor & """>" & ojumun.FMasterItemList(i).FUserID & "</font></a>","USERID" 
		    'iobj.AddData i,"<font color=""" & ojumun.FItemList(i).GetUserLevelColor & """>" & ojumun.FItemList(i).FUserID & "</font>","USERID" 
		end if
		iobj.AddData i,ojumun.FMasterItemList(i).FBuyName,"BUYNAME"  
        iobj.AddData i,ojumun.FMasterItemList(i).FReqName,"REQNAME"  
        iobj.AddData i,FormatNumber(ojumun.FMasterItemList(i).FTotalSum,0),"TOTALSUM"  
		iobj.AddData i,FormatNumber(ojumun.FMasterItemList(i).FSubTotalPrice,0),"SUBTOTALPRICE" 
		
		iobj.AddData i,ojumun.FMasterItemList(i).JumunMethodName,"ACCOUNTDIVNAME"
        iobj.AddData i,ojumun.FMasterItemList(i).IpkumDivName,"IPKUMDIVNAME"
        iobj.AddData i,ojumun.FMasterItemList(i).CancelYnName,"VALIDGUBUN"  
        iobj.AddData i,Left(ojumun.FMasterItemList(i).FRegDate,10),"REGDATE"
        
        
'        IF(ojumun.FItemList(i).IsForeignDeliver) then
'            iobj.AddData i,"해외","JUMUNDIV"
'        else
'            iobj.AddData i,ojumun.FItemList(i).GetJumunDivName,"JUMUNDIV"
'        end if
         
    next
    response.write returnXMLResultObjArr(cmd,"N","S_OK","",iobj)
    set ojumun = Nothing
    set iobj = Nothing
    
elseif (cmd="orderMasterlist") then
    
    dim searchfield : searchfield = request("searchfield")
    dim userid      : userid 		= requestCheckvar(request("userid"),32)
    dim username    : username 	= requestCheckvar(request("username"),32)
    dim userhp      : userhp 		= requestCheckvar(request("userhp"),32)
    dim etcfield    : etcfield 	= requestCheckvar(request("etcfield"),32)
    dim etcstring   : etcstring 	= requestCheckvar(request("etcstring"),32)
    dim checkYYYYMMDD : checkYYYYMMDD = requestCheckvar(request("checkYYYYMMDD"),32)
    
    if (research="") and (checkYYYYMMDD="") then checkYYYYMMDD="Y"
    
    set ojumun = new COrderMaster
    ojumun.FPageSize = pagesize
    ojumun.FCurrPage = page
    
    if (checkYYYYMMDD="Y") then
    	ojumun.FRectRegStart = RequestCheckVar(request("stdt"),10)
    	ojumun.FRectRegEnd = RequestCheckVar(request("eddt"),10)
    end if
    
    if (request("checkJumunDiv") = "Y") then
        if (jumundiv="flowers") then
        	ojumun.FRectIsFlower = "Y"
        elseif (jumundiv="minus") then
            ojumun.FRectIsMinus = "Y"
        elseif (jumundiv="foreign") then
            ojumun.FRectIsForeign = "Y"
        end if
    end if
    
    if (request("checkJumunSite") = "Y") then
    	ojumun.FRectExtSiteName = jumunsite
    end if
    
    
    if (searchfield = "orderserial") then
            '주문번호
            ojumun.FRectOrderSerial = orderserial
    elseif (searchfield = "userid") then
            '고객아이디
            ojumun.FRectUserID = userid
    elseif (searchfield = "username") then
            '구매자명
            ojumun.FRectBuyname = username
    elseif (searchfield = "userhp") then
            '구매자핸드폰
            ojumun.FRectBuyHp = userhp
    elseif (searchfield = "etcfield") then
            '기타조건
            if etcfield="01" then
            	ojumun.FRectBuyname = etcstring
            elseif etcfield="02" then
            	ojumun.FRectReqName = etcstring
            elseif etcfield="03" then
            	ojumun.FRectUserID = etcstring
            elseif etcfield="04" then
            	ojumun.FRectIpkumName = etcstring
            elseif etcfield="06" then
            	ojumun.FRectSubTotalPrice = etcstring
            elseif etcfield="07" then
            	ojumun.FRectBuyPhone = etcstring
            elseif etcfield="08" then
            	ojumun.FRectReqHp = etcstring
            elseif etcfield="09" then
            	ojumun.FRectReqSongjangNo = etcstring
            elseif etcfield="10" then
            	ojumun.FRectReqPhone = etcstring
            end if
    end if
    
    ''검색조건 없을때 최근 N건 검색
    ojumun.QuickSearchOrderList
    
    '' 과거 6개월 이전 내역 검색
    if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
        ojumun.FRectOldOrder = "on"
        ojumun.QuickSearchOrderList
        
        'if (ojumun.FResultCount>0) then
        '    AlertMsg = "6개월 이전 주문입니다."
        'end if
    end if
    arrColName = Array("VALIDGUBUN","JUMUNDIV" ,"ORDERSERIAL","SITE","USERID","BUYNAME","REQNAME","TOTALSUM","COUPONSUM","MILETOTALPRICE","ETCSUM","SUBTOTALPRICE","ACCOUNTDIVNAME","IPKUMDIVNAME","REGDATE","IPKUMDATE","BALJUDATE")
    
    
    set iobj = new CTnGridData
    
    iobj.FPageSize  = ojumun.FPageSize
    iobj.FCurrPage  = ojumun.FCurrPage
    iobj.FTotalCount = ojumun.FTotalCount
    
    for i=0 to ojumun.FResultCount-1
        
        iobj.AddData i,ojumun.FItemList(i).CancelYnName,"VALIDGUBUN"  
        IF(ojumun.FItemList(i).IsForeignDeliver) then
            iobj.AddData i,"해외","JUMUNDIV"
        else
            iobj.AddData i,ojumun.FItemList(i).GetJumunDivName,"JUMUNDIV"
        end if
         
        ''iobj.AddData i,"<a href=""?orderserial="&ojumun.FItemList(i).FOrderSerial&""">"&ojumun.FItemList(i).FOrderSerial&"</a>","ORDERSERIAL"  
        iobj.AddData i,ojumun.FItemList(i).FOrderSerial,"ORDERSERIAL"  
        iobj.AddData i,ojumun.FItemList(i).FSitename,"SITE"  
        if ojumun.FItemList(i).FSitename<>"10x10" then
		    iobj.AddData i,ojumun.FItemList(i).FAuthCode,"USERID" 
		else
		    iobj.AddData i,"<a href=""?searchfield=userid&userid="& ojumun.FItemList(i).FUserID&"""><font color=""" & ojumun.FItemList(i).GetUserLevelColor & """>" & ojumun.FItemList(i).FUserID & "</font></a>","USERID" 
		    'iobj.AddData i,"<font color=""" & ojumun.FItemList(i).GetUserLevelColor & """>" & ojumun.FItemList(i).FUserID & "</font>","USERID" 
		    
		end if
		    
         
        iobj.AddData i,ojumun.FItemList(i).FBuyName,"BUYNAME"  
        iobj.AddData i,ojumun.FItemList(i).FReqName,"REQNAME"  
        iobj.AddData i,FormatNumber(ojumun.FItemList(i).FTotalSum,0),"TOTALSUM"  
        iobj.AddData i,FormatNumber(ojumun.FItemList(i).Ftencardspend,0),"COUPONSUM"  
        iobj.AddData i,FormatNumber(ojumun.FItemList(i).Fmiletotalprice,0),"MILETOTALPRICE"  
        iobj.AddData i,FormatNumber(ojumun.FItemList(i).Fallatdiscountprice+ ojumun.FItemList(i).Fspendmembership,0),"ETCSUM"  
        iobj.AddData i,FormatNumber(ojumun.FItemList(i).FSubTotalPrice,0),"SUBTOTALPRICE" 
        iobj.AddData i,ojumun.FItemList(i).JumunMethodName,"ACCOUNTDIVNAME"
        iobj.AddData i,ojumun.FItemList(i).IpkumDivName,"IPKUMDIVNAME"
        iobj.AddData i,Left(ojumun.FItemList(i).FRegDate,10),"REGDATE"
        iobj.AddData i,Left(ojumun.FItemList(i).Fipkumdate,10),"IPKUMDATE"
        iobj.AddData i,Left(ojumun.FItemList(i).Fbaljudate,10),"BALJUDATE"
        
    next
    response.write returnXMLResultObjArr(cmd,"N","S_OK","",iobj)
    set ojumun = Nothing
    set iobj = Nothing
elseif (cmd="orderDetaillist") then
    
    dim oorderdetail
    dim buf1, buf2
    set oorderdetail = new COrderMaster
    oorderdetail.FRectOrderSerial = orderserial
    oorderdetail.QuickSearchOrderDetail
    
    if (oorderdetail.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
        oorderdetail.FRectOldOrder = "on"
        oorderdetail.QuickSearchOrderDetail
    end if

    set iobj = new CTnGridData

    iobj.FPageSize   = pagesize
    iobj.FCurrPage   = page
    iobj.FTotalCount = oorderdetail.FResultCount
    
    for i=0 to oorderdetail.FResultCount-1
        
        if (oorderdetail.FItemList(i).Fitemid=0) then
            buf1 = "배송비"
            buf2 = oorderdetail.BeasongCD2Name(oorderdetail.FItemList(i).Fitemoption)
        else
            buf1 = oorderdetail.FItemList(i).FItemName
            buf2 = oorderdetail.FItemList(i).FItemOptionName
        end if
        
        iobj.AddData i,oorderdetail.FItemList(i).CancelStateStr,"VALIDGUBUN"  
        iobj.AddData i,oorderdetail.FItemList(i).GetStateName,"STATE"  
        iobj.AddData i,oorderdetail.FItemList(i).Fitemid,"ITEMID"  
        iobj.AddData i,"<IMG src="""&oorderdetail.FItemList(i).FSmallImage&""">","ITEMIMAGE"  
        ''iobj.AddData i,oorderdetail.FItemList(i).FSmallImage,"ITEMIMAGE"  
        iobj.AddData i,oorderdetail.FItemList(i).Fmakerid,"BRANDID" 
        iobj.AddData i,buf1,"ITEMNAME"  
        iobj.AddData i,buf2,"ITEMOPTIONNAME"  
        iobj.AddData i,oorderdetail.FItemList(i).FItemNo,"ITEMNO"  
        iobj.AddData i,oorderdetail.FItemList(i).Forgprice,"ORGCOST"  
        iobj.AddData i,oorderdetail.FItemList(i).Fitemcost,"ITEMCOST"  
        iobj.AddData i,oorderdetail.FItemList(i).Fupcheconfirmdate,"CONFIRMDATE"  
        iobj.AddData i,oorderdetail.FItemList(i).Fbeasongdate,"CHULGODATE"  
        iobj.AddData i,oorderdetail.FItemList(i).Fsongjangno,"DLVINFO"  
    Next
    
    response.write returnXMLResultObjArr(cmd,"N","S_OK","",iobj)
    set oorderdetail = Nothing
    set iobj = Nothing
elseif (cmd="channelSellsum") then
    
    set oReport = new CDatamartItemSale
    oReport.FRectStartDate = RequestCheckVar(request("stdt"),10)
    oReport.FRectEndDate = RequestCheckVar(request("eddt"),10)
    oReport.FRectDateGubun = "M"
    oReport.FRectIncludeMinus = RequestCheckVar(request("ckMinus"),10)
    oReport.FRectCD1 = RequestCheckVar(request("cdL"),10)
    oReport.FRectCD2 = RequestCheckVar(request("cdM"),10)
    
    oReport.SearchMallSellrePortChannel
    
    set iobj = new CTnGridData
    
    iobj.FPageSize   = oReport.FPageSize
    iobj.FCurrPage   = oReport.FCurrPage
    iobj.FTotalCount = oReport.FTotalCount
    
    for i=0 to oReport.FResultCount-1
        
        iobj.AddData i,oReport.FItemList(i).FcateName,"CATEGUBUN"  
        iobj.AddData i,"","GRAPH"  
        iobj.AddData i,FormatNumber(oReport.FItemList(i).Fsellcnt,0)&"건","ORDERCNT"  
        iobj.AddData i,FormatNumber(oReport.FItemList(i).Fselltotal,0)&"원","SELLSUM"  
        iobj.AddData i,FormatNumber(oReport.FItemList(i).Fbuytotal,0)&"원","BUYSUM"  
        iobj.AddData i,FormatNumber(oReport.FItemList(i).Fselltotal-oReport.FItemList(i).Fbuytotal,0)&"원","GAINSUM"  
        if oreport.FItemList(i).Fselltotal<>0 then
	  	    iobj.AddData i,100-CLng(oreport.FItemList(i).Fbuytotal/oreport.FItemList(i).Fselltotal*100*100)/100 & "%","GAINPRO"  
	    end if 
	    iobj.AddData i,"<IMG src=""http://scm.10x10.co.kr/images/icon_search.jpg"">","BIGO"
        
'        if ojumun.FItemList(i).FSitename<>"10x10" then
'		    iobj.AddData i,ojumun.FItemList(i).FAuthCode,"USERID" 
'		else
'		    iobj.AddData i,"<a href=""?searchfield=userid&userid="& ojumun.FItemList(i).FUserID&"""><font color=""" & ojumun.FItemList(i).GetUserLevelColor & """>" & ojumun.FItemList(i).FUserID & "</font></a>","USERID" 
'		end if
'		    
'         
'        iobj.AddData i,ojumun.FItemList(i).FBuyName,"BUYNAME"  
'        iobj.AddData i,ojumun.FItemList(i).FReqName,"REQNAME"  
'        iobj.AddData i,FormatNumber(ojumun.FItemList(i).FTotalSum,0),"TOTALSUM"  
'        iobj.AddData i,FormatNumber(ojumun.FItemList(i).Ftencardspend,0),"COUPONSUM"  
'        iobj.AddData i,FormatNumber(ojumun.FItemList(i).Fmiletotalprice,0),"MILETOTALPRICE"  
'        iobj.AddData i,FormatNumber(ojumun.FItemList(i).Fallatdiscountprice+ ojumun.FItemList(i).Fspendmembership,0),"ETCSUM"  
'        iobj.AddData i,FormatNumber(ojumun.FItemList(i).FSubTotalPrice,0),"SUBTOTALPRICE" 
'        iobj.AddData i,ojumun.FItemList(i).JumunMethodName,"ACCOUNTDIVNAME"
'        iobj.AddData i,ojumun.FItemList(i).IpkumDivName,"IPKUMDIVNAME"
'        iobj.AddData i,Left(ojumun.FItemList(i).FRegDate,10),"REGDATE"
'        iobj.AddData i,Left(ojumun.FItemList(i).Fipkumdate,10),"IPKUMDATE"
'        iobj.AddData i,Left(ojumun.FItemList(i).Fbaljudate,10),"BALJUDATE"
        
    next
    response.write returnXMLResultObjArr(cmd,"N","S_OK","",iobj)
    
    SET oReport=Nothing
    set iobj = Nothing
else
    response.write returnXMLResultObjArr(cmd,"N","S_ERR","지정되지 않았습니다." + cmd,"")
end if
%>
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->