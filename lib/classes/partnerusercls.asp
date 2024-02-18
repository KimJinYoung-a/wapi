<%
function drawPartner3plCompany(selectBoxName,selectedId,onChange)
   dim tmp_str,query1

	%><select class="select" name="<%=selectBoxName%>" <%=onChange%> ><option value=''>-선택-</option><%

	query1 = " select top 100 t.tplcompanyid, t.tplcompanyname " & vbCrLf
	query1 = query1 & " from [db_partner].[dbo].[tbl_partner_tpl] t " & vbCrLf
	query1 = query1 & " order by t.tplcompanyname " & vbCrLf
	rsget.CursorLocation = adUseClient
	rsget.Open query1, dbget, adOpenForwardOnly, adLockReadOnly

   if not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
		   tmp_str = ""
           if Lcase(selectedId) = Lcase(rsget("tplcompanyid")) then
               tmp_str = " selected"
           end if

		   response.write("<option value='"&rsget("tplcompanyid")&"' "&tmp_str&">"&db2html(rsget("tplcompanyname"))&"</option>")
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function

class COffContractInfoItem
	public Fshopid
	public Fchargediv
	public Fdefaultmargin
	public Fshopname
	public FShopdiv

	public function GetChargeDivName()
		if Fchargediv="2" then
			GetChargeDivName = "특정"
		elseif Fchargediv="4" then
			GetChargeDivName = "출고정산"
		elseif Fchargediv="5" then
			GetChargeDivName = "출고정산"
		elseif Fchargediv="6" then
			GetChargeDivName = "업체특정"
		elseif Fchargediv="8" then
			GetChargeDivName = "매입"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class COffContractInfo
	public FItemList()
	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectDesignerID

	public function GetSpecialChargeDivName(byval ishopid)
		dim i

		for i=LBound(FItemList) to UBound(FItemList)
			if IsObject(FItemList(i)) then
				if FItemList(i).Fshopid=ishopid then
					GetSpecialChargeDivName = FItemList(i).GetChargeDivName
				end if
			end if
		next

	end function

	public function GetSpecialDefaultMargin(byval ishopid)
		dim i

		for i=LBound(FItemList) to UBound(FItemList)
			if IsObject (FItemList(i)) then
				if FItemList(i).Fshopid=ishopid then
					GetSpecialDefaultMargin = FItemList(i).Fdefaultmargin
				end if
			end if
		next

	end function

    public Sub GetOffMajorContractInfo()
        dim sqlStr, i
		sqlStr = "select d.shopid,d.chargediv,d.defaultmargin,u.shopname,u.shopdiv"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr + " where d.shopid=u.userid"
		sqlStr = sqlStr + " and u.shopdiv in ('2','4','6','8','10','12')"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
        sqlStr = sqlStr + " order by convert(int,shopdiv)"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)
		i=0
		do until rsget.eof
			set FItemList(i) = new COffContractInfoItem
			FItemList(i).Fshopid = rsget("shopid")
			FItemList(i).Fchargediv = rsget("chargediv")
			FItemList(i).Fdefaultmargin = rsget("defaultmargin")
			FItemList(i).Fshopname = db2html(rsget("shopname"))
			FItemList(i).FShopdiv = rsget("shopdiv")
			i=i+1
			rsget.movenext
		loop
		rsget.Close
	end sub

	public Sub GetPartnerOffContractInfo()
		dim sqlStr, i
		sqlStr = "select d.shopid,d.chargediv,d.defaultmargin,u.shopname,u.shopdiv"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr + " where d.shopid=u.userid"
		sqlStr = sqlStr + " and u.isusing='Y'"
		sqlStr = sqlStr + " and makerid='" + FRectDesignerID + "'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)
		i=0
		do until rsget.eof
			set FItemList(i) = new COffContractInfoItem
			FItemList(i).Fshopid = rsget("shopid")
			FItemList(i).Fchargediv = rsget("chargediv")
			FItemList(i).Fdefaultmargin = rsget("defaultmargin")
			FItemList(i).Fshopname = db2html(rsget("shopname"))
			FItemList(i).FShopdiv = rsget("shopdiv")
			i=i+1
			rsget.movenext
		loop
		rsget.Close
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CPartnerGroupItem
	public Fgroupid
	public Fgroupid_old
	public Fcompany_name
	public Fcompany_no
	public Fceoname
	public Fcompany_uptae
	public Fcompany_upjong
	public Fcompany_zipcode
	public Fcompany_address
	public Fcompany_address2
	public Fcompany_tel
	public Fcompany_fax
	public Freturn_zipcode				'업체 사무실주소로 전용
	public Freturn_address				'업체 사무실주소로 전용
	public Freturn_address2				'업체 사무실주소로 전용
	public Fjungsan_gubun
	public Fjungsan_bank
	public Fjungsan_date
	public Fjungsan_date_off     '' 추가.
	public Fjungsan_acctname
	public Fjungsan_acctno
	public Fmanager_name
	public Fmanager_phone
	public Fmanager_hp
	public Fmanager_email
	public Fdeliver_name
	public Fdeliver_phone
	public Fdeliver_hp
	public Fdeliver_email
	public Fjungsan_name
	public Fjungsan_phone
	public Fjungsan_hp
	public Fjungsan_email
	public Fregdate
	public Flastupdate
	public FComment

	public FBrandList

	public Fdefaultsongjangdiv
	public FPrtidx
	public Fpopularid
    public FpartnerCnt

    public FerpUsing
    public FerpCust_CD
    public FerpCUST_USE_CD


    public function getPartnerIdInfoStr()
        getPartnerIdInfoStr = ""
        if IsNULL(Fpopularid) or (Fpopularid="") or (FpartnerCnt<1) then
            exit function
        end if

        getPartnerIdInfoStr = Fpopularid

        if (FpartnerCnt>1) then
            getPartnerIdInfoStr = getPartnerIdInfoStr & " (외 " & CStr(FpartnerCnt-1) & "건)"
        end if
    end function

    public function getBrandListHTML()
        dim buf : buf = getBrandList
        dim splited
        if InStr(buf,",")>0 then
           splited = split(buf,",")
           buf = ""
           for i=Lbound(splited) to Ubound(splited)
               buf=buf+splited(i)+","
               if (((i+1) mod 5)=0) then buf=buf+"<br>"
           next
        end if

        getBrandListHTML = buf
    end function

	public function getBrandList()
		if Right(FBrandList,1)="," then
			getBrandList = Left(FBrandList,Len(FBrandList)-1)
		else
			getBrandList = FBrandList
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class

class CPartnerGroup
	public FItemList()
	public FOneItem

	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage

	public FRectGroupid
	public FrectDesigner
	public Frectconame
    public FRectSocno
	public FRectceoname

	'/admin/member/grouplist.asp
	public Sub GetGroupInfoListByBrand
		dim sqlStr,i , sqlsearch

		if FrectDesigner<>"" then
			sqlsearch = sqlsearch + " and p.id like '%" + FrectDesigner + "%'"
		end if
		if FRectceoname <> "" then
			sqlsearch = sqlsearch & " and g.ceoname like '%"&FRectceoname&"%'"
		end if

		sqlStr = "select count(g.groupid) as cnt "
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner_group g ,"
		sqlStr = sqlStr + " [db_partner].[dbo].tbl_partner p "
		sqlStr = sqlStr + " where g.groupid=p.groupid " & sqlsearch
		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " g.*, T.popularid, T.cnt as partnerCnt "
		sqlStr = sqlStr + " , s.defaultsongjangdiv, c.prtidx "
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p,"
		sqlStr = sqlStr + " [db_partner].[dbo].tbl_partner_group g"
		sqlStr = sqlStr + "     left join ("
		sqlStr = sqlStr + "     select (select top 1 id from [db_partner].[dbo].tbl_partner where groupid = p.groupid order by regdate desc) as popularid "
		sqlStr = sqlStr + "				, p.groupid, count(id) cnt  "
		sqlStr = sqlStr + "     from [db_partner].[dbo].tbl_partner p"
		sqlStr = sqlStr + "     where p.isusing='Y'"
		sqlStr = sqlStr + "     group by p.groupid"
		sqlStr = sqlStr + "     ) T on g.groupid=T.groupid"
		sqlStr = sqlStr + "   left outer join db_partner.dbo.tbl_partner as s "
		sqlStr = sqlStr + "			on s.id = T.popularid  "
		sqlStr = sqlStr + "   left outer join db_user.dbo.tbl_user_c as c "
		sqlStr = sqlStr + "			on c.userid = T.popularid  "
		sqlStr = sqlStr + " where g.groupid=p.groupid " & sqlsearch
		sqlStr = sqlStr + " order by g.groupid"
		 'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		''올림.
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CPartnerGroupItem

				FItemList(i).Fgroupid         = rsget("groupid")
				FItemList(i).Fcompany_name    = db2html(rsget("company_name"))
				FItemList(i).Fcompany_no      = rsget("company_no")
				FItemList(i).Fceoname         = db2html(rsget("ceoname"))
				FItemList(i).Fcompany_uptae   = db2html(rsget("company_uptae"))
				FItemList(i).Fcompany_upjong  = db2html(rsget("company_upjong"))
				FItemList(i).Fcompany_zipcode = rsget("company_zipcode")
				FItemList(i).Fcompany_address = db2html(rsget("company_address"))
				FItemList(i).Fcompany_address2= db2html(rsget("company_address2"))
				FItemList(i).Fcompany_tel     = rsget("company_tel")
				FItemList(i).Fcompany_fax     = rsget("company_fax")
				FItemList(i).Freturn_zipcode  = rsget("return_zipcode")
				FItemList(i).Freturn_address  = db2html(rsget("return_address"))
				FItemList(i).Freturn_address2 = db2html(rsget("return_address2"))
				FItemList(i).Fjungsan_gubun   = rsget("jungsan_gubun")
				FItemList(i).Fjungsan_bank    = rsget("jungsan_bank")
				FItemList(i).Fjungsan_date    = rsget("jungsan_date")
				FItemList(i).Fjungsan_date_off= rsget("jungsan_date_off")
				FItemList(i).Fjungsan_acctname= db2html(rsget("jungsan_acctname"))
				FItemList(i).Fjungsan_acctno  = rsget("jungsan_acctno")
				FItemList(i).Fmanager_name    = db2html(rsget("manager_name"))
				FItemList(i).Fmanager_phone   = rsget("manager_phone")
				FItemList(i).Fmanager_hp      = rsget("manager_hp")
				FItemList(i).Fmanager_email   = db2html(rsget("manager_email"))
				FItemList(i).Fdeliver_name    = db2html(rsget("deliver_name"))
				FItemList(i).Fdeliver_phone   = rsget("deliver_phone")
				FItemList(i).Fdeliver_hp      = rsget("deliver_hp")
				FItemList(i).Fdeliver_email   = db2html(rsget("deliver_email"))
				FItemList(i).Freturn_zipcode  = rsget("return_zipcode")
				FItemList(i).Freturn_address  = rsget("return_address")
				FItemList(i).Freturn_address2  = rsget("return_address2")
				FItemList(i).Fjungsan_name    = db2html(rsget("jungsan_name"))
				FItemList(i).Fjungsan_phone   = rsget("jungsan_phone")
				FItemList(i).Fjungsan_hp      = rsget("jungsan_hp")
				FItemList(i).Fjungsan_email   = db2html(rsget("jungsan_email"))
				FItemList(i).Fregdate         = rsget("regdate")
				FItemList(i).Flastupdate      = rsget("lastupdate")
         FItemList(i).Fpopularid       = rsget("popularid")
         FItemList(i).FpartnerCnt      = rsget("partnerCnt")
				 FItemList(i).Fdefaultsongjangdiv= rsget("defaultsongjangdiv")
				FItemList(i).FPrtidx       = format00(4,rsget("prtidx"))
				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
	end sub


	'//admin/member/grouplist.asp
	public Sub GetGroupInfoList
		dim sqlStr,i , sqlsearch

		if Frectconame<>"" then
			sqlsearch = sqlsearch + " and g.company_name like '%" + Frectconame + "%' "
		end if

		if FRectSocno<>"" then
		    sqlsearch = sqlsearch + " and Replace(g.company_no,'-','')='" + Replace(FRectSocno,"-","") + "'"
		end if

        if FRectGroupid<>"" then
		    sqlsearch = sqlsearch + " and g.groupid ='" + FRectGroupid + "'"
		end if

		if FRectceoname <> "" then
			sqlsearch = sqlsearch & " and g.ceoname like '%"&FRectceoname&"%'"
		end if

		sqlStr = "select count(groupid) as cnt "
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner_group g"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch

        'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " g.*, T.popularid, T.cnt as partnerCnt "
		sqlStr = sqlStr + " , s.defaultsongjangdiv, c.prtidx "
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner_group g"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + "     select (select top 1 id from [db_partner].[dbo].tbl_partner where groupid = p.groupid order by regdate desc) as popularid "
		sqlStr = sqlStr + "				, p.groupid, count(id) cnt  "
		sqlStr = sqlStr + "     from [db_partner].[dbo].tbl_partner p"
		sqlStr = sqlStr + "     where p.isusing='Y'"

		if Frectconame<>"" then
			sqlStr = sqlStr + " and p.groupid in (select groupid from [db_partner].[dbo].tbl_partner_group where company_name like '%" + Frectconame + "%') "
		end if

		sqlStr = sqlStr + "     group by p.groupid"
		sqlStr = sqlStr + "     ) T on g.groupid=T.groupid"
		sqlStr = sqlStr + "   left outer join db_partner.dbo.tbl_partner as s "
		sqlStr = sqlStr + "			on s.id = T.popularid  "
		sqlStr = sqlStr + "   left outer join db_user.dbo.tbl_user_c as c "
		sqlStr = sqlStr + "			on c.userid = T.popularid  "
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by g.groupid desc"

        ''response.write sqlStr & "<Br>"

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		''올림.
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CPartnerGroupItem

				FItemList(i).Fgroupid         = rsget("groupid")
				FItemList(i).Fcompany_name    = db2html(rsget("company_name"))
				FItemList(i).Fcompany_no      = rsget("company_no")
				FItemList(i).Fceoname         = db2html(rsget("ceoname"))
				FItemList(i).Fcompany_uptae   = db2html(rsget("company_uptae"))
				FItemList(i).Fcompany_upjong  = db2html(rsget("company_upjong"))
				FItemList(i).Fcompany_zipcode = rsget("company_zipcode")
				FItemList(i).Fcompany_address = db2html(rsget("company_address"))
				FItemList(i).Fcompany_address2= db2html(rsget("company_address2"))
				FItemList(i).Fcompany_tel     = rsget("company_tel")
				FItemList(i).Fcompany_fax     = rsget("company_fax")
				FItemList(i).Freturn_zipcode  = rsget("return_zipcode")
				FItemList(i).Freturn_address  = db2html(rsget("return_address"))
				FItemList(i).Freturn_address2 = db2html(rsget("return_address2"))
				FItemList(i).Fjungsan_gubun   = rsget("jungsan_gubun")
				FItemList(i).Fjungsan_bank    = rsget("jungsan_bank")
				FItemList(i).Fjungsan_date    = rsget("jungsan_date")
				FItemList(i).Fjungsan_date_off= rsget("jungsan_date_off")
				FItemList(i).Fjungsan_acctname= db2html(rsget("jungsan_acctname"))
				FItemList(i).Fjungsan_acctno  = rsget("jungsan_acctno")
				FItemList(i).Fmanager_name    = db2html(rsget("manager_name"))
				FItemList(i).Fmanager_phone   = rsget("manager_phone")
				FItemList(i).Fmanager_hp      = rsget("manager_hp")
				FItemList(i).Fmanager_email   = db2html(rsget("manager_email"))
				FItemList(i).Fdeliver_name    = db2html(rsget("deliver_name"))
				FItemList(i).Fdeliver_phone   = rsget("deliver_phone")
				FItemList(i).Fdeliver_hp      = rsget("deliver_hp")
				FItemList(i).Fdeliver_email   = db2html(rsget("deliver_email"))
				FItemList(i).Freturn_zipcode  = rsget("return_zipcode")
				FItemList(i).Freturn_address  = rsget("return_address")
				FItemList(i).Freturn_address2  = rsget("return_address2")

				FItemList(i).Fjungsan_name    = db2html(rsget("jungsan_name"))
				FItemList(i).Fjungsan_phone   = rsget("jungsan_phone")
				FItemList(i).Fjungsan_hp      = rsget("jungsan_hp")
				FItemList(i).Fjungsan_email   = db2html(rsget("jungsan_email"))
				FItemList(i).Fregdate         = rsget("regdate")
				FItemList(i).Flastupdate      = rsget("lastupdate")
        FItemList(i).Fpopularid       = rsget("popularid")
        FItemList(i).FpartnerCnt      = rsget("partnerCnt")
			 	FItemList(i).Fdefaultsongjangdiv= rsget("defaultsongjangdiv")
			 	FItemList(i).FPrtIdx      		= format00(4,rsget("prtidx"))
				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
	end sub

	'//common/offshop/beasong/popupchejumunsms_off.asp		'//admin/offshop/popupchejumunsms_off.asp
	public Sub GetOneGroupInfo
		dim sqlStr
		sqlStr = "select top 1 g.*,b.CUST_USE_CD from [db_partner].[dbo].tbl_partner_group g"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_TMS_BA_CUST b"
	    sqlStr = sqlStr + " on g.erpCust_cd=b.CUST_CD"
		sqlStr = sqlStr + " where g.groupid='" + html2db(FRectGroupid) + "'"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		set FOneItem = new CPartnerGroupItem
		if Not rsget.Eof then
			FOneItem.Fgroupid         = db2html(rsget("groupid"))
			FOneItem.Fcompany_name    = db2html(rsget("company_name"))
			FOneItem.Fcompany_no      = rsget("company_no")
			FOneItem.Fceoname         = db2html(rsget("ceoname"))
			FOneItem.Fcompany_uptae   = db2html(rsget("company_uptae"))
			FOneItem.Fcompany_upjong  = db2html(rsget("company_upjong"))
			FOneItem.Fcompany_zipcode = rsget("company_zipcode")
			FOneItem.Fcompany_address = db2html(rsget("company_address"))
			FOneItem.Fcompany_address2= db2html(rsget("company_address2"))
			FOneItem.Fcompany_tel     = rsget("company_tel")
			FOneItem.Fcompany_fax     = rsget("company_fax")
			FOneItem.Freturn_zipcode  = rsget("return_zipcode")					'업체 사무실 주소로 전용
			FOneItem.Freturn_address  = db2html(rsget("return_address"))
			FOneItem.Freturn_address2 = db2html(rsget("return_address2"))
			FOneItem.Fjungsan_gubun   = rsget("jungsan_gubun")
			FOneItem.Fjungsan_bank    = rsget("jungsan_bank")
			FOneItem.Fjungsan_date    = rsget("jungsan_date")
			FOneItem.Fjungsan_date_off= rsget("jungsan_date_off")
			FOneItem.Fjungsan_acctname= db2html(rsget("jungsan_acctname"))
			FOneItem.Fjungsan_acctno  = rsget("jungsan_acctno")
			FOneItem.Fmanager_name    = db2html(rsget("manager_name"))
			FOneItem.Fmanager_phone   = rsget("manager_phone")
			FOneItem.Fmanager_hp      = rsget("manager_hp")
			FOneItem.Fmanager_email   = db2html(rsget("manager_email"))
			FOneItem.Fdeliver_name    = db2html(rsget("deliver_name"))
			FOneItem.Fdeliver_phone   = rsget("deliver_phone")
			FOneItem.Fdeliver_hp      = rsget("deliver_hp")
			FOneItem.Fdeliver_email   = db2html(rsget("deliver_email"))
			FOneItem.Fjungsan_name    = db2html(rsget("jungsan_name"))			'정산담당자
			FOneItem.Fjungsan_phone   = rsget("jungsan_phone")
			FOneItem.Fjungsan_hp      = rsget("jungsan_hp")
			FOneItem.Fjungsan_email   = db2html(rsget("jungsan_email"))
			FOneItem.Fregdate         = rsget("regdate")
			FOneItem.Flastupdate      = rsget("lastupdate")
			FOneItem.FerpUsing        = rsget("erpUsing")
            FOneItem.FerpCust_CD      = rsget("erpCust_CD")                 ''한사업자번호가 여러개 따져 있을경우 ERP에 한 코드로 물리기 위함 또는 기존 따진 ERP코드에 물리기위함.
            FOneItem.FerpCUST_USE_CD  = rsget("CUST_USE_CD")
		end if
		rsget.close

        dim bufStr
		if FOneItem.Fgroupid<>"" then
			sqlStr = "select id, isusing from [db_partner].[dbo].tbl_partner"
			sqlStr = sqlStr + " where groupid='" + FRectGroupid + "'"
			'sqlStr = sqlStr + " and isusing='Y'"
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				do until rsget.eof
				    if rsget("isusing")="Y" then
				        bufStr = rsget("id")
				    else
				        bufStr = "<font color='#BBBBBB'>" & rsget("id") & "</font>"
				    end if

					FOneItem.FBrandList = FOneItem.FBrandList + bufStr + ","
					rsget.movenext
				loop
			rsget.close
		end if
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0
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
end class

class COutBrandItem
	public Fyyyymm
	public Fmakerid
	public Fmakername
	public Fmakerlevel
	public Fnewitemcount

	public Flastonjungsansum
	public Flastoffjungsansum
	public Flastminuscnt
	public Flastminussum

	public Fusingitemcount
	public Fregdate

	public Fbrandregdate
	public Fmaeipdiv
	public Fdefaultmargine
	public Femail

	public Fisusing
	public Fisextusing
	public Fstreetusing
	public Fextstreetusing
	public Fspecialbrand

	public Fcurrentusingitemcnt
	public Fcurrentsellitemcnt
    public Foffcurrentusingitemcnt
    public Fetccurrentusingitemcnt
	public Fmduserid
    public Fpartnerusing
    public FlastsellDateON
    public FlastsellDateOF
    public Flastgrouplogindate
    public FLastPartnerLogindate
    public Fgroupid
    public Fcompany_name
    public Fcompany_no

    public function IsReqOutProcessBrand()
        dim BaseDT : BaseDT = dateadd("m",1,Fyyyymm+"-01")

        if (Fisusing="N") and (Fisextusing="N") and (Fstreetusing="N") and (Fspecialbrand="N") and (Fcurrentusingitemcnt=0)  and (Foffcurrentusingitemcnt=0) then
            IsReqOutProcessBrand = false
            Exit function
        end if

        IsReqOutProcessBrand = (isNULL(FlastsellDateON) or FlastsellDateON<dateAdd("yyyy",-1,BaseDT)) '' 최종판매월ON 1년이전
        IsReqOutProcessBrand = IsReqOutProcessBrand AND (isNULL(FlastsellDateOF) or FlastsellDateOF<dateAdd("yyyy",-1,BaseDT)) '' 최종판매월OF 1년이전

        IsReqOutProcessBrand = IsReqOutProcessBrand AND (Fnewitemcount<1) ''신상품
        IsReqOutProcessBrand = IsReqOutProcessBrand AND (Fcurrentsellitemcnt<1) ''판매상품수ON

        IsReqOutProcessBrand = IsReqOutProcessBrand AND (Fbrandregdate<dateAdd("m",-6,BaseDT)) '' 브랜드 등록일 6개월

        ''IsReqOutProcessBrand = IsReqOutProcessBrand AND (Foffcurrentusingitemcnt<1) ''판매(사용)상품수OF
    end function

    public function IsReqBrandScmClose()
        dim BaseDT : BaseDT = dateadd("m",1,Fyyyymm+"-01")

        IsReqBrandScmClose = (NOT IsReqOutProcessBrand)
        IsReqBrandScmClose = IsReqBrandScmClose and ((Fisusing="N") and (Fisextusing="N") and (Fstreetusing="N") and (Fspecialbrand="N") and (Fcurrentusingitemcnt=0))
        IsReqBrandScmClose = IsReqBrandScmClose and (isNULL(FLastPartnerLogindate) or FLastPartnerLogindate<Fyyyymm+"-01")
        IsReqBrandScmClose = IsReqBrandScmClose and (Fpartnerusing="Y")
        IsReqBrandScmClose = IsReqBrandScmClose and (isNULL(FlastsellDateON) or FlastsellDateON<dateAdd("yyyy",-1,BaseDT)) '' 최종판매월ON 1년이전
        IsReqBrandScmClose = IsReqBrandScmClose and (isNULL(FlastsellDateOF) or FlastsellDateOF<dateAdd("yyyy",-1,BaseDT)) '' 최종판매월OF 1년이전
    end function

	public function GetMWUName()
		if Fmaeipdiv="M" then
			GetMWUName = "매입"
		elseif Fmaeipdiv="W" then
			GetMWUName = "특정"
		elseif Fmaeipdiv="U" then
			GetMWUName = "업체"
		end if
	end function

	public function GetMwName()
		if Fmaeipdiv="M" then
			GetMwName = "매입"
		elseif Fmaeipdiv="W" then
			GetMwName = "특정"
		end if
	end function

	public function GetMwColor()
		if Fmaeipdiv="M" then
			GetMwColor = "#FF4444"
		elseif Fmaeipdiv="W" then
			GetMwColor = "#4444FF"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class

class CDesignerUserItem
	public FUserID
	public FSocNo
	public FSocName
	public FSocMail
	public FSocUrl
	public FSocPhone
	public FSocFax
	public FIsUsing
	public FIsB2B
	public FUserDiv
	public FUserDivName
	public FIsExtUsing


	public function Is10x10Using()
		Is10x10Using = false
		if IsNull(FIsUsing) or FIsUsing="N" then Exit function
		if FIsUsing="Y" then
			Is10x10Using = true
		end if
	end function

	public function IsExtUsing()
		IsExtUsing = false
		if IsNull(FIsExtUsing) or FIsExtUsing="N" then Exit function
		if FIsExtUsing="Y" then
			IsExtUsing = true
		end if
	end function

	public function IsB2B()
		IsB2B = false
		if IsNull(FIsB2B) or FIsB2B="N" then Exit function
		if FIsB2B="Y" then
			IsB2B = true
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CPartnerUserItem
	public Fid
	public Fpassword
	public Fdiscountrate
	public Fcompany_name
	public Faddress
	public Ftel
	public Ffax
	public Fbigo
	public Furl
	public Fmanager_name
	public Fmanager_address
	public Fcommission
	public Femail
	public Fuserdiv
	public Fcatecode
	public Fisusing
	public Fbuseo
	public Fpart
	public Fcposition
	public Fintro
	public Fmsn
	public Fbirthday
	public Fuserimg


	public Fonlyflg
	public Fartistflg
	public Fkdesignflg
	public FstandardCateCode

	public FVatinclude
	public Fmaeipdiv		'기본계약구분
	public Fdefaultmargine	'기본마진
	public FM_margin		'별도 매입시마진
	public FW_margin		'별도 특정시마진
	public FU_margin		'별도 업체배송시마진

	public Fpid

	public Fcompany_no
	public Fzipcode
	public Fceoname
	public Fmanager_phone
	public Fmanager_hp
	public Fdeliver_name
	public Fdeliver_phone
	public Fdeliver_hp
	public Fdeliver_email
	public Fjungsan_name
	public Fjungsan_phone
	public Fjungsan_hp
	public Fjungsan_email
	public Fjungsan_gubun
	public Fjungsan_bank
	public Fjungsan_date
	public Fjungsan_date_off
	public Fjungsan_date_frn

	public Fjungsan_acctname
	public Fjungsan_acctno

	public Fcompany_upjong
	public Fcompany_uptae

	public FGroupId
	public FSubId
	public Fppass

	public Fsocname
	public Fsocname_kor

	public Fisextusing
	public Fspecialbrand
	public FPrtIdx

	public Fstreetusing
	public Fextstreetusing
	public FTotalitemcount
	Public Fsocicon
	public Fsoclog
	public Ftitleimgurl
	public Fdgncomment
	public Fsamebrand

	public Fmduserid
	public Fregdate

	public Fpartnerusing

    public FdefaultFreeBeasongLimit         ''기본무료배송비기준
    public FdefaultDeliverPay               ''기본배송비
    public FdefaultDeliveryType             ''기본배송정책

    public Fdefaultsongjangdiv

    public Ftakbae_name
    public Ftakbae_tel

    public Flec_yn
    public Fdiy_yn
    public Flec_margin
    public Fmat_margin
    public Fdiy_margin
    public Fitemid

    public Foffcatecode
    public Foffmduserid

    public Freturn_zipcode
    public Freturn_address
    public Freturn_address2
    public FpurchaseType
    public FpurchasetypeNm
    public FsellType

    public FselltypeNm
    public FpcUserDiv

    public FsellBizCd
    public FsellBizNm

    public FpadminUrl
    public FpadminId
    public FpadminPwd
    public FpmallSellType
    public FpmallSellTypeNm
    public FpcomType
    public FpcomTypeNm
    public Ftaxevaltype
    public FtaxevaltypeNm
    public Fetcjungsantype
    public Ftplcompanyid        ''2013/10/31 추가
	public FlastInfoChgDT

    '' 매입처인지.
    public function isbuyingPartner()
        isbuyingPartner = (FpcUserDiv="9999_02") or (FpcUserDiv="9999_14")
    end function

    '' 매장출고처 인지
    public function isShopPartner()
        isShopPartner = (FpcUserDiv="501_21") or (FpcUserDiv="503_21")
    end function

    '' 온라인 제휴사 인지
    public function isOnlinePartner()
        isOnlinePartner = (FpcUserDiv="999_50")
    end function

    '' 기타출고처 인지.
    public function isEtcSellPartner()
        isEtcSellPartner = (FpcUserDiv="900_21")
    end function

    public function getCommissionPro()
        if isNULL(Fcommission) then
            getCommissionPro = 0
            Exit function
        end if

        if (Fcommission="") then
            getCommissionPro = 0
            Exit function
        end if

        getCommissionPro = CLNG(Fcommission*100.0*100.0)/100
    end function

	public function getSocIconUrl()
		getSocIconUrl = webImgUrl + "/image/brandicon/" + Fsocicon
	end function

	public function getSocLogoUrl()
		getSocLogoUrl = webImgUrl + "/image/brandlogo/" + Fsoclog
	end function

	public function getTitleImgUrl()
		getTitleImgUrl = webImgUrl + "/image/brandlogo/" + Ftitleimgurl
	end function

	public function getRackCode()
		getRackCode = format00(4,FPrtIdx)
	end function

	public function GetBrandDivName()
		if Fuserdiv="02" then
			GetBrandDivName = "디자인업체"
		elseif Fuserdiv="03" then
			GetBrandDivName = "플라워업체"
		elseif Fuserdiv="04" then
			GetBrandDivName = "패션업체"
		elseif Fuserdiv="05" then
			GetBrandDivName = "쥬얼리업체"
		elseif Fuserdiv="06" then
			GetBrandDivName = "케어업체"
		elseif Fuserdiv="07" then
			GetBrandDivName = "애견업체"
		elseif Fuserdiv="08" then
			GetBrandDivName = "보드게임"
		elseif Fuserdiv="13" then
			GetBrandDivName = "여행몰업체"
		elseif Fuserdiv="14" then
			GetBrandDivName = "강사"
		else
			GetBrandDivName = Fuserdiv
		end if
	end function

	public function GetUserDivName()
		if Fuserdiv="02" then
			GetUserDivName = "매입처"
		elseif Fuserdiv="14" then
			GetUserDivName = "강사"
		elseif Fuserdiv="21" then
			GetUserDivName = "출고처"
	    elseif Fuserdiv="50" then
			GetUserDivName = "제휴사(온라인)"
		elseif Fuserdiv="95" then
			GetUserDivName = "사용안함"
		else
			GetUserDivName = Fuserdiv
		end if
	end function

	public function GetCateCodeName()
		if Fcatecode="10" then
			GetCateCodeName = "문구,사무"
		elseif Fcatecode="15" then
			GetCateCodeName = "인테리어,리빙데코"
		elseif Fcatecode="20" then
			GetCateCodeName = "취미,여가,뷰티"
		elseif Fcatecode="25" then
			GetCateCodeName = "주방,욕실,생활"
		elseif Fcatecode="30" then
			GetCateCodeName = "패션,잡화"
		elseif Fcatecode="35" then
			GetCateCodeName = "보석,액세서리"
		elseif Fcatecode="40" then
			GetCateCodeName = "키덜트,얼리"
		elseif Fcatecode="45" then
			GetCateCodeName = "선물샵"
		elseif Fcatecode="50" then
			GetCateCodeName = "플라워샵"
		elseif Fcatecode="94" then
			GetCateCodeName = "아카데미 DIY"
		elseif Fcatecode="95" then
			GetCateCodeName = "아카데미 강좌"
		else
			GetCateCodeName = Fuserdiv
		end if
	end function

	public function GetMWUName()
		if Fmaeipdiv="M" then
			GetMWUName = "매입"
		elseif Fmaeipdiv="W" then
			GetMWUName = "특정"
		elseif Fmaeipdiv="U" then
			GetMWUName = "업체배송"
		end if
	end function

	public function GetMaeipDivName()
		if Fmaeipdiv="M" then
			GetMaeipDivName = "매입"
		elseif Fmaeipdiv="W" then
			GetMaeipDivName = "특정"
		end if
	end function

	''// 업체별 배송비 부과 상품(업체 조건 배송)
	public Function IsUpcheParticleDeliverItem()
	    IsUpcheParticleDeliverItem = (FDefaultFreeBeasongLimit>0) and (FDefaultDeliverPay>0) and (FdefaultDeliveryType="9")
	end function

	''// 업체착불 배송여부
	public Function IsUpcheReceivePayDeliverItem()
	    IsUpcheReceivePayDeliverItem = (FdefaultDeliveryType="7")
	end function

	'// 무료 배송 여부
	public Function IsFreeBeasong()

		if (FdefaultDeliveryType="2") or (FdefaultDeliveryType="4") or (FdefaultDeliveryType="5") then
			IsFreeBeasong = true
		end if

		''//착불 배송은 무료배송이 아님
		if (FdefaultDeliveryType="7") then
		    IsFreeBeasong = false
		end if
	end Function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CPartnerUser
	public FPartnerList()
	public FOneItem

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectDesignerID
	public FRectDesignerName
	public FRectDesignerDiv
	public FRectIsUsing
	public FRectIsB2BUsing
	public FRectIsExtUsing
	public FRectOrder
    public FRectPCuserDiv

    public FPass_yn
	public Fpassword

	public Fcompany_name
	public Faddress
	public Ftel
	public Ffax
	public Fbigo
	public Furl
	public Fmanager_name
	public Fmanager_address
	public Fcommission
	public Femail
	public Fuserdiv
	public Fcatecode

	public Fisusing
	public Fisextusing
	public Fstreetusing
	public Fextstreetusing
	public Fspecialbrand

	public FVatinclude
	public Fmaeipdiv
	public Fdefaultmargine

	public Fpid
	public FRectoffcatecode
	public Fcompany_no
	public Fzipcode
	public Fceoname
	public Fmanager_phone
	public Fmanager_hp
	public Fdeliver_name
	public Fdeliver_phone
	public Fdeliver_hp
	public Fdeliver_email
	public Fjungsan_name
	public Fjungsan_phone
	public Fjungsan_hp
	public Fjungsan_email
	public Fjungsan_gubun
	public Fjungsan_bank
	public Fjungsan_date
	public Fjungsan_acctname
	public Fjungsan_acctno
	public FRectManagerhp
	public Fcompany_upjong
	public Fcompany_uptae
	public FRectManageremail
	public FRectInitial
	public FRectUserDiv
	public FRectYYYYMM
	public FRectUserDivUnder
	public FRectMdUserID
	public FRectCatecode
	public FRectmakerlevel
	public FRectCompanyName
	public FRectManagerName
	public FRectoffmduserid
	public FRectCompanyNo
    public FRectSOCName
    public Fitemid
    public FRectGroupid

    public FRectPartnerIsUsing
    public FRectnewbrandgbn
    public FRectOutReqBrand

	Private Sub Class_Initialize()
		'redim preserve FPartnerList(0)
		redim  FPartnerList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub GetGroupList()
		dim sqlStr

	end sub

    ''201010 추가
    public sub GetAcademyPartnerList()
        dim sqlStr, i

        sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " c.userid, "
		sqlStr = sqlStr + " c.vatinclude, c.maeipdiv, c.defaultmargine, c.socname, c.socname_kor,"
		sqlStr = sqlStr + " c.isusing, c.isextusing, c.specialbrand, c.prtidx,c.streetusing,c.extstreetusing,c.userdiv,c.catecode,"
		sqlStr = sqlStr + " p.company_name,c.regdate,"
		sqlStr = sqlStr + " p.email, p.address, p.manager_address,"
		sqlStr = sqlStr + " p.tel, p.fax, p.url, p.manager_name, p.id as pid,"
		sqlStr = sqlStr + " p.isusing as partnerusing,"
		sqlStr = sqlStr + " p.company_no, p.zipcode, p.ceoname, p.manager_phone,"
		sqlStr = sqlStr + " p.manager_hp, p.deliver_name, p.deliver_phone, "
		sqlStr = sqlStr + " p.deliver_hp, p.deliver_email, p.jungsan_name, "
		sqlStr = sqlStr + " p.jungsan_phone, p.jungsan_hp, p.jungsan_email,"
		sqlStr = sqlStr + " p.jungsan_gubun, p.jungsan_bank, p.jungsan_date,"
		sqlStr = sqlStr + " p.jungsan_acctname, p.jungsan_acctno,"
		sqlStr = sqlStr + " p.company_upjong, p.company_uptae, IsNULL(p.groupid,'') as groupid, p.subid, p.password as ppass"
		sqlStr = sqlStr + " ,U.lec_yn, U.diy_yn, U.lec_margin, U.mat_margin, U.diy_margin, U.diy_dlv_gubun, U.defaultFreeBeasongLimit, U.defaultDeliveryPay "
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner p "
		sqlStr = sqlStr + "     on c.userid=p.id"
		sqlStr = sqlStr + "     left join [ACADEMYDB].[db_academy].[dbo].tbl_lec_user U "
		sqlStr = sqlStr + "     on c.userid=U.lecturer_id"
		sqlStr = sqlStr + " where c.userid<>''" + vbCrlf
		sqlStr = sqlStr + " and c.userdiv ='14'" + vbCrlf

        if FRectIsUsing="on" then
            sqlStr = sqlStr + " and c.isusing='Y'"
        end if

		if FRectMdUserID<>"" then
			sqlStr = sqlStr + " and c.mduserid='" + FRectMdUserID + "'"
		end if

		if FRectInitial<>"" then
			sqlStr = sqlStr + " and (c.userid like '" + FRectInitial + "%')"
		end if

        if (FRectDesignerID<>"") then
            sqlStr = sqlStr + " and c.userid='"&FRectDesignerID&"'"
        end if

		sqlStr = sqlStr + " order by c.userid asc"

        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.recordCount
		FTotalCount = FResultCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FPartnerList(FResultCount)
		i=0
		do until rsget.eof
			set FPartnerList(i) = new CPartnerUserItem
			FPartnerList(i).Fid    			= db2html(rsget("userid"))
			FPartnerList(i).Fcompany_name  	= db2html(rsget("company_name"))
			FPartnerList(i).Faddress        = db2html(rsget("address"))
			FPartnerList(i).Ftel            = rsget("tel")
			FPartnerList(i).Ffax            = rsget("fax")
			FPartnerList(i).Furl            = rsget("url")
			FPartnerList(i).Fmanager_name   = db2html(rsget("manager_name"))
			FPartnerList(i).Fmanager_address  = db2html(rsget("manager_address"))
			FPartnerList(i).Femail          = db2html(rsget("email"))

			FPartnerList(i).FVatinclude     = rsget("vatinclude")
			FPartnerList(i).Fmaeipdiv       = rsget("maeipdiv")
			FPartnerList(i).Fdefaultmargine = rsget("defaultmargine")
			FPartnerList(i).Fpid			= rsget("pid")
			'oneitem.Fisusing          = rsget("isusing")

			FPartnerList(i).Fcompany_no		= rsget("company_no")
			FPartnerList(i).Fzipcode		= rsget("zipcode")
			FPartnerList(i).Fceoname		= db2html(rsget("ceoname"))
			FPartnerList(i).Fmanager_phone	= rsget("manager_phone")
			FPartnerList(i).Fmanager_hp		= rsget("manager_hp")
			FPartnerList(i).Fdeliver_name	= rsget("deliver_name")
			FPartnerList(i).Fdeliver_phone	= rsget("deliver_phone")
			FPartnerList(i).Fdeliver_hp		= rsget("deliver_hp")
			FPartnerList(i).Fdeliver_email	= rsget("deliver_email")
			FPartnerList(i).Fjungsan_name	= db2html(rsget("jungsan_name"))
			FPartnerList(i).Fjungsan_phone	= rsget("jungsan_phone")
			FPartnerList(i).Fjungsan_hp		= rsget("jungsan_hp")
			FPartnerList(i).Fjungsan_email	= rsget("jungsan_email")
			FPartnerList(i).Fjungsan_gubun	= rsget("jungsan_gubun")
			FPartnerList(i).Fjungsan_bank	= rsget("jungsan_bank")
			FPartnerList(i).Fjungsan_date	= rsget("jungsan_date")
			FPartnerList(i).Fjungsan_acctname	= db2html(rsget("jungsan_acctname"))
			FPartnerList(i).Fjungsan_acctno		= rsget("jungsan_acctno")

			FPartnerList(i).Fcompany_upjong = rsget("company_upjong")
			FPartnerList(i).Fcompany_uptae  = rsget("company_uptae")

			FPartnerList(i).FGroupId  = rsget("groupid")
			FPartnerList(i).FSubId  = rsget("subid")
			FPartnerList(i).Fppass  = rsget("ppass")

			FPartnerList(i).Fsocname  = db2html(rsget("socname"))
			FPartnerList(i).Fsocname_kor  = db2html(rsget("socname_kor"))

			FPartnerList(i).Fisusing	 = rsget("isusing")
			FPartnerList(i).Fisextusing	 = rsget("isextusing")
			FPartnerList(i).Fspecialbrand	= rsget("specialbrand")
			FPartnerList(i).FPrtIdx	= rsget("prtidx")

			FPartnerList(i).Fstreetusing  = rsget("streetusing")
			FPartnerList(i).Fextstreetusing  = rsget("extstreetusing")
			FPartnerList(i).Fuserdiv = rsget("userdiv")
			FPartnerList(i).Fcatecode= rsget("catecode")

			FPartnerList(i).Fregdate = rsget("regdate")

            FPartnerList(i).Fpartnerusing = rsget("partnerusing")

            FPartnerList(i).Flec_yn         = rsget("lec_yn")
            FPartnerList(i).Fdiy_yn         = rsget("diy_yn")
            FPartnerList(i).Flec_margin     = rsget("lec_margin")
            FPartnerList(i).Fmat_margin			= rsget("mat_margin")
            FPartnerList(i).Fdiy_margin			= rsget("diy_margin")
			FPartnerList(i).FdefaultFreeBeasongLimit = rsget("defaultFreeBeasongLimit")
			FPartnerList(i).FdefaultDeliveryType = rsget("diy_dlv_gubun")
            FPartnerList(i).FdefaultDeliverPay	= rsget("defaultDeliveryPay")


			i=i+1
			rsget.movenext
		loop
		rsget.close

    end Sub

	public Sub GetPartnerQuickSearch()
		dim sqlStr


		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " c.userid, "
		sqlStr = sqlStr + " c.vatinclude, c.maeipdiv, c.defaultmargine, c.socname, c.socname_kor,"
		sqlStr = sqlStr + " c.isusing, c.isextusing, c.specialbrand, c.prtidx,c.streetusing,c.extstreetusing,c.userdiv,c.catecode,"
		sqlStr = sqlStr + " p.company_name,c.regdate,"
		sqlStr = sqlStr + " p.email, p.address, p.manager_address,"
		sqlStr = sqlStr + " p.tel, p.fax, p.url, p.manager_name, p.id as pid,"
		sqlStr = sqlStr + " p.isusing as partnerusing,"
		sqlStr = sqlStr + " p.company_no, p.zipcode, p.ceoname, p.manager_phone,"
		sqlStr = sqlStr + " p.manager_hp, p.deliver_name, p.deliver_phone, "
		sqlStr = sqlStr + " p.deliver_hp, p.deliver_email, p.jungsan_name, "
		sqlStr = sqlStr + " p.jungsan_phone, p.jungsan_hp, p.jungsan_email,"
		sqlStr = sqlStr + " p.jungsan_gubun, p.jungsan_bank, p.jungsan_date,"
		sqlStr = sqlStr + " p.jungsan_acctname, p.jungsan_acctno,"
		sqlStr = sqlStr + " p.company_upjong, p.company_uptae, IsNULL(p.groupid,'') as groupid, p.subid, p.password as ppass"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"
		sqlStr = sqlStr + " where c.userid<>''" + vbCrlf
		sqlStr = sqlStr + " and c.userdiv < 22" + vbCrlf

		if FRectMdUserID<>"" then
			sqlStr = sqlStr + " and c.mduserid='" + FRectMdUserID + "'"
		end if

		if FRectInitial<>"" then
			sqlStr = sqlStr + " and (c.userid like '" + FRectInitial + "%')"
		end if

		if FRectUserDiv<>"" then
			sqlStr = sqlStr + " and c.userdiv='" + FRectUserDiv + "'"
		end if

		sqlStr = sqlStr + " order by c.userid asc"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		FTotalCount = FResultCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FPartnerList(FResultCount)
		i=0
		do until rsget.eof
			set FPartnerList(i) = new CPartnerUserItem
			FPartnerList(i).Fid    			= db2html(rsget("userid"))
			FPartnerList(i).Fcompany_name  	= db2html(rsget("company_name"))
			FPartnerList(i).Faddress        = db2html(rsget("address"))
			FPartnerList(i).Ftel            = rsget("tel")
			FPartnerList(i).Ffax            = rsget("fax")
			FPartnerList(i).Furl            = rsget("url")
			FPartnerList(i).Fmanager_name   = db2html(rsget("manager_name"))
			FPartnerList(i).Fmanager_address  = db2html(rsget("manager_address"))
			FPartnerList(i).Femail          = db2html(rsget("email"))

			FPartnerList(i).FVatinclude     = rsget("vatinclude")
			FPartnerList(i).Fmaeipdiv       = rsget("maeipdiv")
			FPartnerList(i).Fdefaultmargine = rsget("defaultmargine")
			FPartnerList(i).Fpid			= rsget("pid")
			'oneitem.Fisusing          = rsget("isusing")

			FPartnerList(i).Fcompany_no		= rsget("company_no")
			FPartnerList(i).Fzipcode		= rsget("zipcode")
			FPartnerList(i).Fceoname		= db2html(rsget("ceoname"))
			FPartnerList(i).Fmanager_phone	= rsget("manager_phone")
			FPartnerList(i).Fmanager_hp		= rsget("manager_hp")
			FPartnerList(i).Fdeliver_name	= rsget("deliver_name")
			FPartnerList(i).Fdeliver_phone	= rsget("deliver_phone")
			FPartnerList(i).Fdeliver_hp		= rsget("deliver_hp")
			FPartnerList(i).Fdeliver_email	= rsget("deliver_email")
			FPartnerList(i).Fjungsan_name	= db2html(rsget("jungsan_name"))
			FPartnerList(i).Fjungsan_phone	= rsget("jungsan_phone")
			FPartnerList(i).Fjungsan_hp		= rsget("jungsan_hp")
			FPartnerList(i).Fjungsan_email	= rsget("jungsan_email")
			FPartnerList(i).Fjungsan_gubun	= rsget("jungsan_gubun")
			FPartnerList(i).Fjungsan_bank	= rsget("jungsan_bank")
			FPartnerList(i).Fjungsan_date	= rsget("jungsan_date")
			FPartnerList(i).Fjungsan_acctname	= db2html(rsget("jungsan_acctname"))
			FPartnerList(i).Fjungsan_acctno		= rsget("jungsan_acctno")

			FPartnerList(i).Fcompany_upjong = rsget("company_upjong")
			FPartnerList(i).Fcompany_uptae  = rsget("company_uptae")

			FPartnerList(i).FGroupId  = rsget("groupid")
			FPartnerList(i).FSubId  = rsget("subid")
			FPartnerList(i).Fppass  = rsget("ppass")

			FPartnerList(i).Fsocname  = db2html(rsget("socname"))
			FPartnerList(i).Fsocname_kor  = db2html(rsget("socname_kor"))

			FPartnerList(i).Fisusing	 = rsget("isusing")
			FPartnerList(i).Fisextusing	 = rsget("isextusing")
			FPartnerList(i).Fspecialbrand	= rsget("specialbrand")
			FPartnerList(i).FPrtIdx	= rsget("prtidx")

			FPartnerList(i).Fstreetusing  = rsget("streetusing")
			FPartnerList(i).Fextstreetusing  = rsget("extstreetusing")
			FPartnerList(i).Fuserdiv = rsget("userdiv")
			FPartnerList(i).Fcatecode= rsget("catecode")

			FPartnerList(i).Fregdate = rsget("regdate")

            FPartnerList(i).Fpartnerusing = rsget("partnerusing")
			i=i+1
			rsget.movenext
		loop
		rsget.close
	end sub


	public Sub GetOutBrandList()
		dim i,sqlstr
        dim BaseDT : BaseDT = dateadd("m",1,FRectYYYYMM+"-01")
		sqlStr = "select top "&FPageSize&" o.*, c.mduserid, c.regdate as brandregdate, c.maeipdiv, c.defaultmargine"
		sqlStr = sqlStr + " ,c.isusing, c.isextusing, c.streetusing, c.extstreetusing, c.specialbrand"
		sqlStr = sqlStr + " ,IsNULL(T.usingcnt,0) as currentusingitemcnt"
		sqlStr = sqlStr + " ,IsNULL(T.sellcount,0) as currentsellitemcnt"
		sqlStr = sqlStr + " ,IsNULL(T2.usingoffallcnt,0) as offcurrentusingitemcnt"
		sqlStr = sqlStr + " ,IsNULL(T2.usingoffallcnt,0) - IsNULL(T2.usingoffcnt,0) as etccurrentusingitemcnt"
		sqlStr = sqlStr + " ,p.isusing as partnerusing"
		sqlStr = sqlStr + " ,p.groupid"
		sqlStr = sqlStr + " ,g.company_name, g.company_no"
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_outbrand o" + vbCrlf
		sqlStr = sqlStr + "     Join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + "     on o.makerid=c.userid"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on o.makerid=p.id"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner_group g on p.groupid=g.groupid"
		sqlStr = sqlStr + " left join ( "
		sqlStr = sqlStr + " 	select makerid, count(*) as usingcnt"
		sqlStr = sqlStr + " 	,sum(case when sellyn='Y' THEN 1 ELSE 0 END) as sellcount"
		sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item"
		sqlStr = sqlStr + " 	where isusing='Y'"
		sqlStr = sqlStr + "		group by makerid"
		sqlStr = sqlStr + " ) as T on T.makerid=o.makerid"
		sqlStr = sqlStr + " left join ( "
		sqlStr = sqlStr + " 	select makerid, count(*) as usingoffallcnt"
		sqlStr = sqlStr + " 	,sum(CASE WHEN itemgubun in ('10','90') then  1 ELSE 0 END) as usingoffcnt"
		sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_shop_item "
		sqlStr = sqlStr + " 	where isusing='Y'"
		sqlStr = sqlStr + "		group by makerid"
		sqlStr = sqlStr + " ) as T2 on T2.makerid=o.makerid"

		sqlStr = sqlStr + " where yyyymm='" + FRectYYYYMM + "'"

		if (FRectDesignerID<>"") then
		    sqlStr = sqlStr + " and o.makerid='"&FRectDesignerID&"'"
		end if

		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing='"&FRectIsUsing&"'"
		end if

        if FRectPartnerIsusing<>"" then
			sqlStr = sqlStr + " and p.isusing='"&FRectPartnerIsusing&"'"
		end if

		if (FRectnewbrandgbn<>"") then
		    if (FRectnewbrandgbn="N") then ''신규브랜드 (6개월 일경우 7로 세팅)
		        sqlStr = sqlStr + " and datediff(m,p.regdate,'"&BaseDT&"')<7"
		    else
		        sqlStr = sqlStr + " and datediff(m,p.regdate,'"&BaseDT&"')>=7"
		    end if
		end if

		if (FRectGroupid<>"") then
		    sqlStr = sqlStr + " and p.groupid='"&FRectGroupid&"'"
		end if

		if (FRectOutReqBrand="Y") then
		    sqlStr = sqlStr + " and isNULL(o.lastsellDateON,'2001-01-01')<'"&dateAdd("yyyy",-1,BaseDT)&"'"
		    sqlStr = sqlStr + " and isNULL(o.lastsellDateOF,'2001-01-01')<'"&dateAdd("yyyy",-1,BaseDT)&"'"
		    sqlStr = sqlStr + " and o.newitemcount<1" ''신상품 0
		    sqlStr = sqlStr + " and IsNULL(T.sellcount,0)<1" ''ON 판매상품수
		    ''sqlStr = sqlStr + " and IsNULL(T2.usingoffallcnt,0)<1" ''OFF 판매상품수

		    ''lastgrouplogindate
		end if

'		if FRectCatecode<>"" then
'			sqlStr = sqlStr + " and c.catecode='" + FRectCatecode + "'"
'		end if
'
'		if FRectMdUserID<>"" then
'			sqlStr = sqlStr + " and c.mduserid='" + FRectMdUserID + "'"
'		end if
'
'		if FRectmakerlevel<>"" then
'			sqlStr = sqlStr + " and o.makerlevel=" + FRectmakerlevel + ""
'		end if
		''sqlStr = sqlStr + " order by o.makerlevel desc, o.lastonjungsansum "
		sqlStr = sqlStr + " order by o.makerid"

'response.write sqlStr
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount

		redim preserve FPartnerList(FResultCount)
		i=0
		do until rsget.eof
			set FPartnerList(i) = new COutBrandItem
			FPartnerList(i).Fyyyymm         = rsget("yyyymm")
			FPartnerList(i).Fmakerid        = rsget("makerid")
			FPartnerList(i).Fmakername      = db2html(rsget("makername"))
			FPartnerList(i).Fmakerlevel     = rsget("makerlevel")
			FPartnerList(i).Fnewitemcount   = rsget("newitemcount")
			FPartnerList(i).Flastonjungsansum = rsget("lastonjungsansum")

            FPartnerList(i).Flastoffjungsansum = rsget("lastoffjungsansum")
   			FPartnerList(i).Flastminuscnt = rsget("lastminuscnt")
   			FPartnerList(i).Flastminussum = rsget("lastminussum")

			FPartnerList(i).Fusingitemcount = rsget("usingitemcount")
			FPartnerList(i).Fregdate        = rsget("regdate")

			FPartnerList(i).Fbrandregdate	= rsget("brandregdate")
			FPartnerList(i).Fmaeipdiv       = rsget("maeipdiv")
			FPartnerList(i).Fdefaultmargine = rsget("defaultmargine")

			'FPartnerList(i).Femail		= rsget("email")
			FPartnerList(i).Fisusing	= rsget("isusing")
			FPartnerList(i).Fisextusing	= rsget("isextusing")

			FPartnerList(i).Fstreetusing 	= rsget("streetusing")
			FPartnerList(i).Fextstreetusing = rsget("extstreetusing")
			FPartnerList(i).Fspecialbrand	= rsget("specialbrand")

			FPartnerList(i).Fcurrentusingitemcnt = rsget("currentusingitemcnt")
			FPartnerList(i).Fcurrentsellitemcnt  = rsget("currentsellitemcnt")
            FPartnerList(i).Foffcurrentusingitemcnt = rsget("offcurrentusingitemcnt")
            FPartnerList(i).Fetccurrentusingitemcnt = rsget("etccurrentusingitemcnt")
			FPartnerList(i).Fmduserid       = rsget("mduserid")
            FPartnerList(i).Fpartnerusing   = rsget("partnerusing")


            FPartnerList(i).FlastsellDateON = rsget("lastsellDateON")
            FPartnerList(i).FlastsellDateOF = rsget("lastsellDateOF")
            FPartnerList(i).Flastgrouplogindate = rsget("lastgrouplogindate")
            FPartnerList(i).Fgroupid = rsget("groupid")
            FPartnerList(i).Fcompany_name   = rsget("company_name")
            FPartnerList(i).Fcompany_no     = rsget("company_no")
            FPartnerList(i).FLastPartnerLogindate = rsget("LastPartnerLogindate")
			i=i+1
			rsget.movenext
		loop
		rsget.close
	end sub

	public Sub GetOnePartnerNUser()
		dim sqlStr
        dim islecturer : islecturer=FALSE


		sqlStr = "select top 1 c.userid,c.userdiv "
		sqlStr = sqlStr + " , c.vatinclude, c.maeipdiv, c.defaultmargine, c.defaultFreeBeasongLimit, c.defaultDeliverPay, c.defaultDeliveryType"
		sqlStr = sqlStr + " , c.socname, c.socname_kor"
		sqlStr = sqlStr + " ,c.isusing, c.isextusing, c.specialbrand, c.prtidx,c.streetusing,c.extstreetusing,c.catecode"
		sqlStr = sqlStr + " ,c.socicon, c.soclogo, c.titleimgurl,c.dgncomment,c.samebrand, c.mduserid, c.regdate, c.onlyflg, c.artistflg, c.kdesignflg, isNull(C.standardCateCode,'') as standardCateCode "
		sqlStr = sqlStr + " ,IsNull(p.M_margin,0) as M_margin, IsNull(p.W_margin,0) as W_margin, IsNull(p.U_margin,0) as U_margin "
		sqlStr = sqlStr + " ,c.socicon, c.soclogo, c.titleimgurl,c.dgncomment, c.mduserid, c.regdate "
		sqlStr = sqlStr + " ,p.company_name,"
		sqlStr = sqlStr + " p.email, p.address, p.manager_address,"
		sqlStr = sqlStr + " p.tel, p.fax, p.url, p.manager_name, p.id as pid,"
		sqlStr = sqlStr + " p.company_no, p.zipcode, p.ceoname, p.manager_phone,"
		sqlStr = sqlStr + " p.manager_hp, p.deliver_name, p.deliver_phone, "
		sqlStr = sqlStr + " p.deliver_hp, p.deliver_email, p.jungsan_name, "
		sqlStr = sqlStr + " p.jungsan_phone, p.jungsan_hp, p.jungsan_email,"
		sqlStr = sqlStr + " p.jungsan_gubun, p.jungsan_bank, p.jungsan_date,p.jungsan_date_off, p.jungsan_date_frn,"
		sqlStr = sqlStr + " p.jungsan_acctname, p.jungsan_acctno,"
		sqlStr = sqlStr + " p.company_upjong, p.company_uptae, IsNULL(p.groupid,'') as groupid, p.subid, p.password as ppass, p.isusing as partnerusing, p.purchaseType, isNull(p.offcatecode,'') as offcatecode, isNull(p.offmduserid,'') as offmduserid, "
		sqlStr = sqlStr + " p.return_zipcode, p.return_address, p.return_address2,isNull(p.userdiv,'') as puserdiv,"  ''추가
		sqlStr = sqlStr + " p.sellType,p.sellBizCd,p.commission,p.bigo,p.taxevaltype,"                     ''추가
		sqlStr = sqlStr + " IsNULL(T.cnt,0) as ttlitemcnt, p.defaultsongjangdiv,"
		sqlStr = sqlStr + " IsNULL(s.divname,'') as takbae_name, IsNULL(s.tel,'') as takbae_tel"
		sqlStr = sqlStr + " ,f.padminUrl,f.padminId,f.padminPwd,f.pmallSellType,f.pcomType,p.etcjungsantype,p.tplcompanyid, p.lastInfoChgDT"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner_addInfo f"
		sqlStr = sqlStr + " on p.id=F.partnerid"
		sqlStr = sqlStr + " left join ( "
		sqlStr = sqlStr + "     select makerid, count(itemid) as cnt from [db_item].[dbo].tbl_item where makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + "     group by makerid "
		sqlStr = sqlStr + " ) as T on c.userid=T.makerid"
		''택배사 명,전화 추가.
		sqlStr = sqlStr + " left join [db_order].[dbo].tbl_songjang_div s"
		sqlStr = sqlStr + "     on p.defaultsongjangdiv=s.divcd"
		sqlStr = sqlStr + " where c.userid='" + FRectDesignerID + "'"

		' response.write sqlstr &"<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount

		if Not rsget.Eof then
			set FOneItem = new CPartnerUserItem
			FOneItem.Fid    			= db2html(rsget("userid"))
			FOneItem.Fcompany_name  	= db2html(rsget("company_name"))
			FOneItem.Faddress        	= db2html(rsget("address"))
			FOneItem.Ftel            	= rsget("tel")
			FOneItem.Ffax            	= rsget("fax")
			FOneItem.Furl            	= rsget("url")
			FOneItem.Fmanager_name   	= db2html(rsget("manager_name"))
			FOneItem.Fmanager_address  	= db2html(rsget("manager_address"))
			FOneItem.Femail          	= db2html(rsget("email"))

			FOneItem.FVatinclude     	= rsget("vatinclude")
			FOneItem.Fmaeipdiv       	= rsget("maeipdiv")
			FOneItem.Fdefaultmargine 	= rsget("defaultmargine")
			FOneItem.FM_margin 		 	= rsget("M_margin")
			FOneItem.FW_margin       	= rsget("W_margin")
			FOneItem.FU_margin       	= rsget("U_margin")

			FOneItem.Fpid				= rsget("pid")
			'oneitem.Fisusing          	= rsget("isusing")

			FOneItem.Fcompany_no		= rsget("company_no")
			FOneItem.Fzipcode			= rsget("zipcode")
			FOneItem.Fceoname			= db2html(rsget("ceoname"))
			FOneItem.Fmanager_phone		= rsget("manager_phone")
			FOneItem.Fmanager_hp		= rsget("manager_hp")
			FOneItem.Fdeliver_name		= rsget("deliver_name")
			FOneItem.Fdeliver_phone		= rsget("deliver_phone")
			FOneItem.Fdeliver_hp		= rsget("deliver_hp")
			FOneItem.Fdeliver_email		= rsget("deliver_email")
			FOneItem.Fjungsan_name		= db2html(rsget("jungsan_name"))
			FOneItem.Fjungsan_phone		= rsget("jungsan_phone")
			FOneItem.Fjungsan_hp		= rsget("jungsan_hp")
			FOneItem.Fjungsan_email		= rsget("jungsan_email")
			FOneItem.Fjungsan_gubun		= rsget("jungsan_gubun")
			FOneItem.Fjungsan_bank		= rsget("jungsan_bank")
			FOneItem.Fjungsan_date		= rsget("jungsan_date")
			FOneItem.Fjungsan_date_off	= rsget("jungsan_date_off")
			FOneItem.Fjungsan_date_frn	= rsget("jungsan_date_frn")

			FOneItem.Fjungsan_acctname	= db2html(rsget("jungsan_acctname"))
			FOneItem.Fjungsan_acctno	= rsget("jungsan_acctno")

			FOneItem.Fcompany_upjong 	= db2html(rsget("company_upjong"))
			FOneItem.Fcompany_uptae  	= db2html(rsget("company_uptae"))

			FOneItem.FGroupId  			= rsget("groupid")
			FOneItem.FSubId  			= rsget("subid")
			FOneItem.Fppass  			= rsget("ppass")

			FOneItem.Fsocname  			= db2html(rsget("socname"))
			FOneItem.Fsocname_kor  		= db2html(rsget("socname_kor"))

			FOneItem.Fisusing	 		= rsget("isusing")
			FOneItem.Fisextusing	 	= rsget("isextusing")
			FOneItem.Fspecialbrand		= rsget("specialbrand")
			FOneItem.FPrtIdx			= rsget("prtidx")

			FOneItem.Fstreetusing  		= rsget("streetusing")
			FOneItem.Fextstreetusing  	= rsget("extstreetusing")
			FOneItem.Fuserdiv 			= rsget("userdiv")
			FOneItem.Fcatecode			= rsget("catecode")

			FOneItem.FTotalitemcount 	= rsget("ttlitemcnt")
			FOneItem.Fsocicon 			= db2html(rsget("socicon"))
			FOneItem.Fsoclog 			= db2html(rsget("soclogo"))
			FOneItem.Ftitleimgurl 		= db2html(rsget("titleimgurl"))
			FOneItem.Fdgncomment 		= db2html(rsget("dgncomment"))
			FOneItem.Fsamebrand 		= db2html(rsget("samebrand"))

			FOneItem.Fmduserid 			= rsget("mduserid")
			FOneItem.Fregdate 			= rsget("regdate")

			FOneItem.Fonlyflg			= rsget("onlyflg")
			FOneItem.Fartistflg			= rsget("artistflg")
			FOneItem.Fkdesignflg		= rsget("kdesignflg")
			FOneItem.FstandardCateCode	= rsget("standardCateCode")

			FOneItem.Fpartnerusing 		= rsget("partnerusing")
			FOneItem.FpurchaseType		= rsget("purchaseType")

			FOneItem.Foffcatecode		= rsget("offcatecode")
			FOneItem.Foffmduserid 		= rsget("offmduserid")

			FOneItem.FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
			FOneItem.FdefaultDeliverPay         = rsget("defaultDeliverPay")
			FOneItem.FdefaultDeliveryType       = rsget("defaultDeliveryType")

			FOneItem.Fdefaultsongjangdiv 		= rsget("defaultsongjangdiv")
			FOneItem.Ftakbae_name 				= db2html(rsget("takbae_name"))
			FOneItem.Ftakbae_tel  				= rsget("takbae_tel")

            FOneItem.Freturn_zipcode        = rsget("return_zipcode")
            FOneItem.Freturn_address        = rsget("return_address")
            FOneItem.Freturn_address2       = rsget("return_address2")

            FOneItem.FpcUserDiv  = rsget("puserdiv") &"_" &FOneItem.Fuserdiv
            FOneItem.FsellType   = rsget("sellType")
            FOneItem.FsellBizCd  = rsget("sellBizCd")
            FOneItem.Fcommission = rsget("commission")
            FOneItem.Fbigo       = rsget("bigo")

            FOneItem.FpadminUrl     = rsget("padminUrl")
            FOneItem.FpadminId      = rsget("padminId")
            FOneItem.FpadminPwd     = rsget("padminPwd")
            FOneItem.FpmallSellType = rsget("pmallSellType")
            FOneItem.FpcomType      = rsget("pcomType")
            FOneItem.Ftaxevaltype   = rsget("taxevaltype")
            FOneItem.Fetcjungsantype= rsget("etcjungsantype")
            FOneItem.Ftplcompanyid  = rsget("tplcompanyid")
            islecturer = (CStr(FOneItem.Fuserdiv) = "14")

			FOneItem.FlastInfoChgDT  = rsget("lastInfoChgDT")

		end if
		rsget.close

        IF (islecturer) THEN
            sqlStr = " select U.lec_yn, U.diy_yn, U.lec_margin, U.mat_margin, U.diy_margin, U.diy_dlv_gubun"
            sqlStr = sqlStr + " , U.defaultFreeBeasongLimit as defaultFreeBeasongLimitAcademy, U.defaultDeliveryPay as defaultDeliveryPayAcademy "
    		sqlStr = sqlStr + " from [ACADEMYDB].[db_academy].[dbo].tbl_lec_user U where U.lecturer_id='"&FRectDesignerID&"'"

			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    		if Not rsget.Eof then
                FOneItem.Flec_yn         			= rsget("lec_yn")
                FOneItem.Fdiy_yn         			= rsget("diy_yn")
                FOneItem.Flec_margin     			= rsget("lec_margin")
                FOneItem.Fmat_margin				= rsget("mat_margin")
                FOneItem.Fdiy_margin				= rsget("diy_margin")

            	'강사
    			FOneItem.FdefaultFreeBeasongLimit 	= rsget("defaultFreeBeasongLimitAcademy")
    			FOneItem.FdefaultDeliveryType 		= rsget("diy_dlv_gubun")
                FOneItem.FdefaultDeliverPay			= rsget("defaultDeliveryPayAcademy")
            end if
            rsget.close
        ENd IF
	end sub

    '//신규버전 20120821	'//admin/member/designerinfolist1.asp
	public Sub GetPartnerSearch()
		dim sqlStr, sqlAdd

		''사용여부 확인..(partner or user_c)
		IF (FRectPCuserDiv<>"") then
		    sqlAdd = sqlAdd + " and p.userdiv='"&splitValue(FRectPCuserDiv,"_",0)&"'"
		    sqlAdd = sqlAdd + " and c.userdiv='"&splitValue(FRectPCuserDiv,"_",1)&"'"
		ELSE
		    sqlAdd = sqlAdd + " and p.userdiv>='500'"
		    sqlAdd = sqlAdd + " and c.userdiv is Not NULL"
		END IF

        if FrectIsUsing="on" then
		    sqlAdd = sqlAdd + " and c.isusing='Y'"
	    end if

	    if FRectInitial<>"" then
			sqlAdd = sqlAdd + " and (p.id like '" + FRectInitial + "%')"
		end if

	    if FRectCompanyNo<>"" then
	        sqlAdd = sqlAdd + " and replace(p.company_no,'-','') = '" + Replace(FRectCompanyNo,"-","") + "'"
	    end if

	    if FRectGroupid<>"" then
	        sqlAdd = sqlAdd + " and p.groupid='" + FRectGroupid + "'"
	    end if

	    if FRectCompanyName<>"" then
	        sqlAdd = sqlAdd + " and p.company_name like '%" + FRectCompanyName + "%'"
	    end if

	    if FRectManagerName<>"" then
	        sqlAdd = sqlAdd + " and p.manager_name like '%" + FRectManagerName + "%'"
	    end if

	    if FRectSOCName<>"" then
	        sqlAdd = sqlAdd + " and c.socname_kor like '%" + FRectSOCName + "%'"
	    end if

	    if FRectDesignerDiv<>"" then
			sqlAdd = sqlAdd + " and c.userdiv='" + FRectDesignerDiv + "'"
		end if

        if FRectDesignerID<>"" then
			sqlAdd = sqlAdd + " and c.userid='" + FRectDesignerID + "'"
		end if

		if FRectMdUserID<>"" then
			sqlAdd = sqlAdd + " and c.mduserid='" + FRectMdUserID + "'"
		end if

		if FRectCatecode<>"" then
			sqlAdd = sqlAdd + " and c.catecode='" + FRectCatecode + "'"
		end if

	    if FRectoffcatecode <> "" then
	    	sqlAdd = sqlAdd + " and p.offcatecode = '"&FRectoffcatecode&"'"
	    end if

	    if FRectoffmduserid <> "" then
	    	sqlAdd = sqlAdd + " and p.offmduserid = '"&offmduserid&"'"
	    end if

	    if FRectManageremail <> "" then
	    	sqlAdd = sqlAdd + " and p.email = '"&FRectManageremail&"'"
	    end if

	    if FRectManagerhp <> "" then
	    	sqlAdd = sqlAdd + " and (p.Manager_phone = '"&FRectManagerhp&"' or p.Manager_hp = '"&FRectManagerhp&"')"
	    end if

		sqlStr = "select Count(c.userid) as cnt"
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p"
		sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " 	on c.userid=p.id"

		If (Fitemid <> "") Then
			sqlStr = sqlStr + " inner join [db_item].[dbo].tbl_item i on p.id=i.makerid"
			sqlStr = sqlStr + " and i.itemid in(" & Fitemid & ")"
		End IF

		sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + sqlAdd

		''response.write sqlStr &"<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " p.id , "
		sqlStr = sqlStr + " c.vatinclude, c.maeipdiv, c.defaultmargine, c.socname, c.socname_kor,"
		sqlStr = sqlStr + " IsNULL(c.defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(c.defaultDeliverPay,0) as defaultDeliverPay, IsNULL(c.defaultDeliveryType,'') as defaultDeliveryType, "
		sqlStr = sqlStr + " c.isusing, c.isextusing, c.specialbrand, c.prtidx,c.streetusing,c.extstreetusing,c.userdiv,c.catecode,"
		sqlStr = sqlStr + " p.company_name,c.regdate,"
		sqlStr = sqlStr + " p.email, p.address, p.manager_address,"
		sqlStr = sqlStr + " p.tel, p.fax, p.url, p.manager_name, p.id as pid,"
		sqlStr = sqlStr + " p.company_no, p.zipcode, p.ceoname, p.manager_phone,"
		sqlStr = sqlStr + " p.manager_hp, p.deliver_name, p.deliver_phone, "
		sqlStr = sqlStr + " p.deliver_hp, p.deliver_email, p.jungsan_name, "
		sqlStr = sqlStr + " p.jungsan_phone, p.jungsan_hp, p.jungsan_email,"
		sqlStr = sqlStr + " p.jungsan_gubun, p.jungsan_bank, p.jungsan_date,"
		sqlStr = sqlStr + " p.jungsan_acctname, p.jungsan_acctno,"
		sqlStr = sqlStr + " p.company_upjong, p.company_uptae, IsNULL(p.groupid,'') as groupid, p.subid, p.password as ppass, p.isusing as partnerusing"
		sqlStr = sqlStr + " ,p.userdiv as puserdiv"
		sqlStr = sqlStr + " , p.sellBizCd, p.selltype, p.purchasetype, p.commission, p.taxevaltype, b.BIZSECTION_NM as sellBizNm "
		sqlStr = sqlStr + " ,f.pmallSellType,f.pcomType"
		sqlStr = sqlStr + " ,a.pcomm_name as selltypeNm, j.pcomm_name as purchasetypeNm, t.pcomm_name as taxevaltypeNm, cc.pcomm_name as pmallSellTypeNm, l.pcomm_name as pcomTypeNm "
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p"
		sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " 	on c.userid=p.id"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner_addInfo f"
		sqlStr = sqlStr + " 	on p.id=F.partnerid"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_TMS_BA_BIZSECTION b"
		sqlStr = sqlStr + " 	on p.sellBizCd=b.BIZSECTION_CD"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner_comm_code a "
		sqlStr = sqlStr + " 	on p.selltype=a.pcomm_cd and a.pcomm_group = 'sellacccd' "
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner_comm_code j "
		sqlStr = sqlStr + " 	on p.purchasetype=j.pcomm_cd and j.pcomm_group = 'selljungsantype' "
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner_comm_code t "
		sqlStr = sqlStr + " 	on p.taxevaltype=t.pcomm_cd and t.pcomm_group = 'taxevaltype' "
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner_comm_code cc "
		sqlStr = sqlStr + " 	on f.pmallSellType=cc.pcomm_cd and cc.pcomm_group = 'mallSellType' "
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner_comm_code l "
		sqlStr = sqlStr + " 	on f.pcomType=l.pcomm_cd and l.pcomm_group = 'pcomType' "

		If (Fitemid <> "") Then
			sqlStr = sqlStr + " inner join [db_item].[dbo].tbl_item i on p.id=i.makerid"
			sqlStr = sqlStr + " and i.itemid in(" & Fitemid & ")"
		End IF

		sqlStr = sqlStr + " where 1=1" + vbCrlf
        sqlStr = sqlStr + sqlAdd

		if FRectOrder="group" then
			sqlStr = sqlStr + " order by p.groupid, p.subid, c.userid "
		elseif FRectOrder="acct" then
			sqlStr = sqlStr + " order by p.jungsan_acctno, p.groupid, p.subid, c.userid "
		else
			sqlStr = sqlStr + " order by c.userid asc"		'브랜드ID순
			'sqlStr = sqlStr + " order by c.regdate desc"		'신규순
		end if

		'response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FPartnerList(FResultCount)
		i=0

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof

				set FPartnerList(i) = new CPartnerUserItem

				FPartnerList(i).Fid    			= db2html(rsget("id"))
				FPartnerList(i).Fcompany_name  	= db2html(rsget("company_name"))
				FPartnerList(i).Faddress        = db2html(rsget("address"))
				FPartnerList(i).Ftel            = rsget("tel")
				FPartnerList(i).Ffax            = rsget("fax")
				FPartnerList(i).Furl            = rsget("url")
				FPartnerList(i).Fmanager_name   = db2html(rsget("manager_name"))
				FPartnerList(i).Fmanager_address  = db2html(rsget("manager_address"))
				FPartnerList(i).Femail          = db2html(rsget("email"))
				FPartnerList(i).FVatinclude     = rsget("vatinclude")
				FPartnerList(i).Fmaeipdiv       = rsget("maeipdiv")
				FPartnerList(i).Fdefaultmargine = rsget("defaultmargine")
			    FPartnerList(i).FdefaultFreeBeasongLimit	= rsget("defaultFreeBeasongLimit")
			    FPartnerList(i).FdefaultDeliverPay			= rsget("defaultDeliverPay")
			    FPartnerList(i).FdefaultDeliveryType		= rsget("defaultDeliveryType")
				FPartnerList(i).Fpid			= rsget("pid")
				'oneitem.Fisusing          = rsget("isusing")
				FPartnerList(i).Fcompany_no		= rsget("company_no")
				FPartnerList(i).Fzipcode		= rsget("zipcode")
				FPartnerList(i).Fceoname		= db2html(rsget("ceoname"))
				FPartnerList(i).Fmanager_phone	= rsget("manager_phone")
				FPartnerList(i).Fmanager_hp		= rsget("manager_hp")
				FPartnerList(i).Fdeliver_name	= rsget("deliver_name")
				FPartnerList(i).Fdeliver_phone	= rsget("deliver_phone")
				FPartnerList(i).Fdeliver_hp		= rsget("deliver_hp")
				FPartnerList(i).Fdeliver_email	= rsget("deliver_email")
				FPartnerList(i).Fjungsan_name	= db2html(rsget("jungsan_name"))
				FPartnerList(i).Fjungsan_phone	= rsget("jungsan_phone")
				FPartnerList(i).Fjungsan_hp		= rsget("jungsan_hp")
				FPartnerList(i).Fjungsan_email	= rsget("jungsan_email")
				FPartnerList(i).Fjungsan_gubun	= rsget("jungsan_gubun")
				FPartnerList(i).Fjungsan_bank	= rsget("jungsan_bank")
				FPartnerList(i).Fjungsan_date	= rsget("jungsan_date")
				FPartnerList(i).Fjungsan_acctname	= db2html(rsget("jungsan_acctname"))
				FPartnerList(i).Fjungsan_acctno		= rsget("jungsan_acctno")
				FPartnerList(i).Fcompany_upjong = rsget("company_upjong")
				FPartnerList(i).Fcompany_uptae  = rsget("company_uptae")
				FPartnerList(i).FGroupId  = rsget("groupid")
				FPartnerList(i).FSubId  = rsget("subid")
				FPartnerList(i).Fppass  = rsget("ppass")
				FPartnerList(i).Fsocname  = db2html(rsget("socname"))
				FPartnerList(i).Fsocname_kor  = db2html(rsget("socname_kor"))
				FPartnerList(i).Fisusing	 = rsget("isusing")
				FPartnerList(i).Fisextusing	 = rsget("isextusing")
				FPartnerList(i).Fspecialbrand	= rsget("specialbrand")
				FPartnerList(i).FPrtIdx	= rsget("prtidx")
				FPartnerList(i).Fstreetusing  = rsget("streetusing")
				FPartnerList(i).Fextstreetusing  = rsget("extstreetusing")
				FPartnerList(i).Fuserdiv = rsget("userdiv")
				FPartnerList(i).Fcatecode= rsget("catecode")
				FPartnerList(i).Fregdate = rsget("regdate")
				FPartnerList(i).Fpartnerusing = rsget("partnerusing")

                FPartnerList(i).FpcUserDiv = rsget("puserdiv")&"_"&rsget("userdiv")

                FPartnerList(i).FsellBizCd 			= rsget("sellBizCd")
                FPartnerList(i).FsellBizNm 			= rsget("sellBizNm")

                FPartnerList(i).Fselltype 			= rsget("selltype")
                FPartnerList(i).FselltypeNm 		= rsget("selltypeNm")
                FPartnerList(i).Fpurchasetype 		= rsget("purchasetype")
                FPartnerList(i).FpurchasetypeNm 	= rsget("purchasetypeNm")
                FPartnerList(i).Fcommission 		= rsget("commission")
                FPartnerList(i).Ftaxevaltype 		= rsget("taxevaltype")
                FPartnerList(i).FtaxevaltypeNm 		= rsget("taxevaltypeNm")
                FPartnerList(i).FpmallSellType 		= rsget("pmallSellType")
                FPartnerList(i).FpmallSellTypeNm 	= rsget("pmallSellTypeNm")
                FPartnerList(i).FpcomType 			= rsget("pcomType")
                FPartnerList(i).FpcomTypeNm			= rsget("pcomTypeNm")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
	end sub

    '//common/offshop/beasong/popupchejumunsms_off.asp		'//admin/offshop/popupchejumunsms_off.asp
	public Sub GetPartnerNUserCList()
		dim sqlStr
		''#################################################
		''총 갯수.
		''#################################################
		sqlStr = "select Count(c.userid) as cnt"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c" + vbCrlf
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"

		If Fitemid <> "" Then
			sqlStr = sqlStr + " inner join [db_item].[dbo].tbl_item i on c.userid=i.makerid"
		End IF

		sqlStr = sqlStr + " where 1 = 1 "

		if (FRectUserDiv = "all") then
			''
		else
			sqlStr = sqlStr + " and c.userdiv < 22 "
		end if

		If Fitemid <> "" Then
			sqlStr = sqlStr + " and i.itemid in(" & Fitemid & ")"
		End IF

		if FRectUserDivUnder<>"" then
			sqlStr = sqlStr + " and c.userdiv <" + CStr(FRectUserDivUnder) + "" + vbCrlf
		else
			sqlStr = sqlStr + " and c.userdiv < 22 "
		end if

        if FrectIsUsing="on" then
		    sqlStr = sqlStr + " and c.isusing='Y'"
	    end if

	    if FRectDesignerDiv<>"" then
			sqlStr = sqlStr + " and c.userdiv='" + FRectDesignerDiv + "'"
		end if


		if FRectInitial="etc" then
			sqlStr = sqlStr + " and ((Left(c.userid,1)<'a') or (Left(c.userid,1)>'Z'))"
		elseif FRectInitial<>"" then
			sqlStr = sqlStr + " and (c.userid like '" + FRectInitial + "%')"
	    elseif FRectCompanyName<>"" then
	        sqlStr = sqlStr + " and p.company_name like '%" + FRectCompanyName + "%'"
	    elseif FRectCompanyNo<>"" then
	        sqlStr = sqlStr + " and replace(p.company_no,'-','') = '" + Replace(FRectCompanyNo,"-","") + "'"
	    elseif FRectManagerName<>"" then
	        sqlStr = sqlStr + " and p.manager_name like '%" + FRectManagerName + "%'"
	    elseif FRectSOCName<>"" then
	        sqlStr = sqlStr + " and c.socname_kor like '%" + FRectSOCName + "%'"
		else
			if FrectIsUsing="off_new" then
				sqlStr = sqlStr + " and c.isusing='N'"
				sqlStr = sqlStr + " and datediff(d,c.regdate,getdate())<31"
			elseif FrectIsUsing="off_old" then
				sqlStr = sqlStr + " and c.isusing='N'"
				sqlStr = sqlStr + " and datediff(d,c.regdate,getdate())>92"
			elseif FrectIsUsing="outbrand" then
				sqlStr = sqlStr + " and c.isusing='N'"
				sqlStr = sqlStr + " and p.isusing='Y'"
				sqlStr = sqlStr + " and datediff(d,c.regdate,getdate())>92"
			end if

			if FRectDesignerID<>"" then
				sqlStr = sqlStr + " and c.userid='" + FRectDesignerID + "'"
			end if

			if FRectMdUserID<>"" then
				sqlStr = sqlStr + " and c.mduserid='" + FRectMdUserID + "'"
			end if

			if FRectCatecode<>"" then
				sqlStr = sqlStr + " and c.catecode='" + FRectCatecode + "'"
			end if
		end if
		'response.write sqlStr &"<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " c.userid, "
		sqlStr = sqlStr + " c.vatinclude, c.maeipdiv, c.defaultmargine, c.socname, c.socname_kor,"
		sqlStr = sqlStr + " IsNULL(c.defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(c.defaultDeliverPay,0) as defaultDeliverPay, IsNULL(c.defaultDeliveryType,'') as defaultDeliveryType, "
		sqlStr = sqlStr + " c.isusing, c.isextusing, c.specialbrand, c.prtidx,c.streetusing,c.extstreetusing,c.userdiv,c.catecode,"
		sqlStr = sqlStr + " p.company_name,c.regdate,"
		sqlStr = sqlStr + " p.email, p.address, p.manager_address,"
		sqlStr = sqlStr + " p.tel, p.fax, p.url, p.manager_name, p.id as pid,"
		sqlStr = sqlStr + " p.company_no, p.zipcode, p.ceoname, p.manager_phone,"
		sqlStr = sqlStr + " p.manager_hp, p.deliver_name, p.deliver_phone, "
		sqlStr = sqlStr + " p.deliver_hp, p.deliver_email, p.jungsan_name, "
		sqlStr = sqlStr + " p.jungsan_phone, p.jungsan_hp, p.jungsan_email,"
		sqlStr = sqlStr + " p.jungsan_gubun, p.jungsan_bank, p.jungsan_date,"
		sqlStr = sqlStr + " p.jungsan_acctname, p.jungsan_acctno,"
		sqlStr = sqlStr + " p.company_upjong, p.company_uptae, IsNULL(p.groupid,'') as groupid, p.subid, p.password as ppass, p.isusing as partnerusing"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"

		If Fitemid <> "" Then
			sqlStr = sqlStr + " inner join [db_item].[dbo].tbl_item i on c.userid=i.makerid"
		End IF

		sqlStr = sqlStr + " where 1 = 1 "

		if (FRectUserDiv = "all") then
			''
		else
			sqlStr = sqlStr + " and c.userdiv < 22 "
		end if

		If Fitemid <> "" Then
			sqlStr = sqlStr + " and i.itemid in(" & Fitemid & ")"
		End IF

		if FRectUserDivUnder<>"" then
			sqlStr = sqlStr + " and c.userdiv <" + CStr(FRectUserDivUnder) + "" + vbCrlf
		end if

        if FrectIsUsing="on" then
		    sqlStr = sqlStr + " and c.isusing='Y'"
	    end if

	    if FRectDesignerDiv<>"" then
			sqlStr = sqlStr + " and c.userdiv='" + FRectDesignerDiv + "'"
		end if

		if FRectInitial="etc" then
			sqlStr = sqlStr + " and ((Left(c.userid,1)<'a') or (Left(c.userid,1)>'Z'))"
		elseif FRectInitial<>"" then
			sqlStr = sqlStr + " and (c.userid like '" + FRectInitial + "%')"
		elseif FRectCompanyName<>"" then
	        sqlStr = sqlStr + " and p.company_name like '%" + FRectCompanyName + "%'"
	    elseif FRectCompanyNo<>"" then
	        sqlStr = sqlStr + " and replace(p.company_no,'-','') = '" + Replace(FRectCompanyNo,"-","") + "'"
	    elseif FRectManagerName<>"" then
	        sqlStr = sqlStr + " and p.manager_name like '%" + FRectManagerName + "%'"
	     elseif FRectSOCName<>"" then
	        sqlStr = sqlStr + " and c.socname_kor like '%" + FRectSOCName + "%'"
		else
			if FrectIsUsing="off_new" then
				sqlStr = sqlStr + " and c.isusing='N'"
				sqlStr = sqlStr + " and datediff(d,c.regdate,getdate())<31"
			elseif FrectIsUsing="off_old" then
				sqlStr = sqlStr + " and c.isusing='N'"
				sqlStr = sqlStr + " and datediff(d,c.regdate,getdate())>92"
			elseif FrectIsUsing="outbrand" then
				sqlStr = sqlStr + " and c.isusing='N'"
				sqlStr = sqlStr + " and p.isusing='Y'"
				sqlStr = sqlStr + " and datediff(d,c.regdate,getdate())>92"
			end if

			if FRectDesignerID<>"" then
				sqlStr = sqlStr + " and c.userid='" + FRectDesignerID + "'"
			end if

			if FRectMdUserID<>"" then
				sqlStr = sqlStr + " and c.mduserid='" + FRectMdUserID + "'"
			end if

			if FRectCatecode<>"" then
				sqlStr = sqlStr + " and c.catecode='" + FRectCatecode + "'"
			end if
		end if

		if FRectOrder="group" then
			sqlStr = sqlStr + " order by p.groupid, p.subid, c.userid "
		elseif FRectOrder="acct" then
			sqlStr = sqlStr + " order by p.jungsan_acctno, p.groupid, p.subid, c.userid "
		else
			sqlStr = sqlStr + " order by c.userid asc"
		end if

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FPartnerList(FResultCount)
		i=0

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof

				set FPartnerList(i) = new CPartnerUserItem
				FPartnerList(i).Fid    			= db2html(rsget("userid"))
				FPartnerList(i).Fcompany_name  	= db2html(rsget("company_name"))
				FPartnerList(i).Faddress        = db2html(rsget("address"))
				FPartnerList(i).Ftel            = rsget("tel")
				FPartnerList(i).Ffax            = rsget("fax")
				FPartnerList(i).Furl            = rsget("url")
				FPartnerList(i).Fmanager_name   = db2html(rsget("manager_name"))
				FPartnerList(i).Fmanager_address  = db2html(rsget("manager_address"))
				FPartnerList(i).Femail          = db2html(rsget("email"))

				FPartnerList(i).FVatinclude     = rsget("vatinclude")
				FPartnerList(i).Fmaeipdiv       = rsget("maeipdiv")
				FPartnerList(i).Fdefaultmargine = rsget("defaultmargine")
			    FPartnerList(i).FdefaultFreeBeasongLimit	= rsget("defaultFreeBeasongLimit")
			    FPartnerList(i).FdefaultDeliverPay			= rsget("defaultDeliverPay")
			    FPartnerList(i).FdefaultDeliveryType		= rsget("defaultDeliveryType")

				FPartnerList(i).Fpid			= rsget("pid")
				'oneitem.Fisusing          = rsget("isusing")

				FPartnerList(i).Fcompany_no		= rsget("company_no")
				FPartnerList(i).Fzipcode		= rsget("zipcode")
				FPartnerList(i).Fceoname		= db2html(rsget("ceoname"))
				FPartnerList(i).Fmanager_phone	= rsget("manager_phone")
				FPartnerList(i).Fmanager_hp		= rsget("manager_hp")
				FPartnerList(i).Fdeliver_name	= rsget("deliver_name")
				FPartnerList(i).Fdeliver_phone	= rsget("deliver_phone")
				FPartnerList(i).Fdeliver_hp		= rsget("deliver_hp")
				FPartnerList(i).Fdeliver_email	= rsget("deliver_email")
				FPartnerList(i).Fjungsan_name	= db2html(rsget("jungsan_name"))
				FPartnerList(i).Fjungsan_phone	= rsget("jungsan_phone")
				FPartnerList(i).Fjungsan_hp		= rsget("jungsan_hp")
				FPartnerList(i).Fjungsan_email	= rsget("jungsan_email")
				FPartnerList(i).Fjungsan_gubun	= rsget("jungsan_gubun")
				FPartnerList(i).Fjungsan_bank	= rsget("jungsan_bank")
				FPartnerList(i).Fjungsan_date	= rsget("jungsan_date")
				FPartnerList(i).Fjungsan_acctname	= db2html(rsget("jungsan_acctname"))
				FPartnerList(i).Fjungsan_acctno		= rsget("jungsan_acctno")

				FPartnerList(i).Fcompany_upjong = rsget("company_upjong")
				FPartnerList(i).Fcompany_uptae  = rsget("company_uptae")

				FPartnerList(i).FGroupId  = rsget("groupid")
				FPartnerList(i).FSubId  = rsget("subid")
				FPartnerList(i).Fppass  = rsget("ppass")

				FPartnerList(i).Fsocname  = db2html(rsget("socname"))
				FPartnerList(i).Fsocname_kor  = db2html(rsget("socname_kor"))

				FPartnerList(i).Fisusing	 = rsget("isusing")
				FPartnerList(i).Fisextusing	 = rsget("isextusing")
				FPartnerList(i).Fspecialbrand	= rsget("specialbrand")
				FPartnerList(i).FPrtIdx	= rsget("prtidx")

				FPartnerList(i).Fstreetusing  = rsget("streetusing")
				FPartnerList(i).Fextstreetusing  = rsget("extstreetusing")
				FPartnerList(i).Fuserdiv = rsget("userdiv")
				FPartnerList(i).Fcatecode= rsget("catecode")
				FPartnerList(i).Fregdate = rsget("regdate")

				FPartnerList(i).Fpartnerusing = rsget("partnerusing")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
	end sub

	public function duplicateUserID(byval userid)
		dim sqlStr, CNT
		sqlStr = "select count(id) as cnt from [db_partner].[dbo].tbl_partner"
		sqlStr = sqlStr + " where id='" + CStr(userid) + "'"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		    CNT = rsget("cnt")
		rsget.close

		if (CNT<1) then
		    sqlStr = "select count(*) as cnt from [db_user].[dbo].tbl_logindata where userid='" & userid & "'"
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    	        CNT = rsget("cnt")
    		rsget.close
		end if

		duplicateUserID = (CNT>0)
	end function

	public Sub addNewPartner(byval userid,userpass,username,usermail,userdiv, discountrate,commission,bigo)
	    dim sqlStr
        dim C_userdiv : C_userdiv = "50"
        dim Enc_userpass : Enc_userpass= MD5(userpass)
        dim Enc_userpass64 : Enc_userpass64= SHA256(MD5(userpass))
        dim manager_name : manager_name=""
        dim catecode : catecode ="999"
        dim isusing : isusing="N"
        dim isextusing : isextusing="N"
        dim isb2b : isb2b="Y"
        dim maeipdiv : maeipdiv = "M"
        dim defaultmargine : defaultmargine = commission*100
        dim socname_kor : socname_kor = username
        dim company_name : company_name = username
        dim streetusing : streetusing = "N"
        dim extstreetusing : extstreetusing ="N"
        dim specialbrand : specialbrand="N"
        dim mduserid : mduserid = ""
        dim vDefaultFreeBeasongLimit : vDefaultFreeBeasongLimit="null"
        dim vDefaultDeliverPay : vDefaultDeliverPay= "null"

        ''userdiv in tbl_user_c / tbl_logindata Char(2)
         '02' >매입처(일반) :: 6565
      	 '14' >아카데미     :: 214
      	 '19'	            :: 2  streetshop_01, streetshop_02  ==>정리요망
      	 '20' >가맹점매입처 :: 24  menu002 외/ 기존 waple 등 ==>정리요망
      	 '21' >출고처       :: 172
      	 '50' >제휴몰               (신규)
      	 '95' >사용안함     :: 60
      	 '99'	            :: 1   Gift_Manager     ==>정리요망



        ''userdiv in tbl_partner  (INT)
         '9999' >매입처(업체)   :: 6881
         '9000' >강사(?)
         '999'  >제휴사         :: 89
         '900'  >기타출고처     :: 신규

         ''''' 매장관련
         '500' >매장공통
         '501' >직영매장        :: 18
         '502' >수수료매장      :: 1
         '503' >대리점          :: 44
         '509'                  :: 1

         '101' >오프샾          :: 54
         '111' >오프샾점장      :: 10
         '112' >오프샾부점장
         '509' >오프매출조회

         '''기타 매장.
         '201' >Zoom            :: 2        ==>정리요망
         '301' >College         :: 22       ==>정리요망

         '' 사원
         '9' >관리자            :: 16
         '7' >마스타            :: 4
         '5' >LV4               :: 8
         '4' >LV3               :: 148
         '2' >LV2               :: 124
         '1' >LV1               :: 8

    On Error Resume Next
    dbget.beginTrans
        sqlStr = "insert into [db_user].[dbo].tbl_logindata"
    	sqlStr = sqlStr + "(userid,userpass,userdiv,lastlogin,Enc_userpass,Enc_userpass64,counter)" + vbCrlf
    	sqlStr = sqlStr + " Values("
    	sqlStr = sqlStr + " '" + (userid) + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + (userpass) + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + (C_userdiv) + "'" + vbCrlf
    	sqlStr = sqlStr + ",getdate()" + vbCrlf
    	sqlStr = sqlStr + ",'" + (Enc_userpass) + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + (Enc_userpass64) + "'" + vbCrlf
    	sqlStr = sqlStr + ",0" & ")"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	    ''insert tbl_user_c
    	sqlStr = "insert into [db_user].[dbo].tbl_user_c" & vbCrlf
    	sqlStr = sqlStr + "(userid,socno,socname,birthday,socmail,socurl,ceoname," + vbCrlf
    	sqlStr = sqlStr + "prcname," + vbCrlf
    	sqlStr = sqlStr + "regdate,mileage,userdiv,catecode," + vbCrlf
    	sqlStr = sqlStr + "isusing, isb2b, isextusing, vatinclude, maeipdiv," + vbCrlf
    	sqlStr = sqlStr + "defaultmargine, socname_kor," & vbCrlf
    	sqlStr = sqlStr + "coname,streetusing,extstreetusing,specialbrand,mduserid" + vbCrlf
    	sqlStr = sqlStr + ",defaultDeliveryType" + vbCrlf
    	sqlStr = sqlStr + ",defaultFreeBeasongLimit,defaultDeliverPay" + vbCrlf

    	sqlStr = sqlStr + " )Values("
    	sqlStr = sqlStr + "'" + userid + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + "" + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + username + "'" + vbCrlf
    	sqlStr = sqlStr + ",convert(varchar(10),getdate(),20)" + vbCrlf
    	sqlStr = sqlStr + ",'" + usermail + "'" + vbCrlf
    	sqlStr = sqlStr + ",''" + vbCrlf
    	sqlStr = sqlStr + ",'" + "" + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + manager_name + "'" + vbCrlf
    	sqlStr = sqlStr + ", getDate()"  + vbCrlf
    	sqlStr = sqlStr + ",0" + vbCrlf
    	sqlStr = sqlStr + ",'" + C_userdiv + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + catecode + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + isusing + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + isb2b + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + isextusing + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + "Y" + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + maeipdiv + "'" + vbCrlf
    	sqlStr = sqlStr + "," + CStr(defaultmargine) + vbCrlf
    	sqlStr = sqlStr + ",'" + socname_kor + "'" + vbCrlf

    	sqlStr = sqlStr + ",'" + company_name + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + streetusing + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + extstreetusing + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + specialbrand + "'" + vbCrlf
    	sqlStr = sqlStr + ",'" + mduserid + "'" + vbCrlf

    	sqlStr = sqlStr + ",null" + vbCrlf
    	sqlStr = sqlStr + ",null" + vbCrlf
    	sqlStr = sqlStr + ",null" + vbCrlf

    	sqlStr = sqlStr +  ")"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		sqlStr = "insert into [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + "(id,password,company_name,email,userdiv,discountrate,commission,bigo)" + vbCrlf
		sqlStr = sqlStr + " values('" + userid + "'," + vbCrlf
		sqlStr = sqlStr + " '" + userpass + "'," + vbCrlf
		sqlStr = sqlStr + " '" + username + "'," + vbCrlf
		sqlStr = sqlStr + " '" + usermail + "'," + vbCrlf
		sqlStr = sqlStr + " " + userdiv + "," + vbCrlf
		sqlStr = sqlStr + " " + discountrate + "," + vbCrlf
		sqlStr = sqlStr + " " + commission + "," + vbCrlf
		sqlStr = sqlStr + " '" + bigo + "'" + vbCrlf
		sqlStr = sqlStr + ")"

		dbget.execute sqlStr
    If Err.Number = 0 Then
	        dbget.CommitTrans
	Else
	        dbget.RollBackTrans
	        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.\n입력한 값들이 너무 길지 않는지 확인바랍니다.')</script>"
	        ''response.write "<script>history.back()</script>"
	        dbget.close()
	        response.end
	End If

	on error Goto 0
   end sub

   public Sub editPartner(byval userid,userpass,username,usermail,userdiv, isusing, discountrate,commission,bigo)
	    dim sqlStr
	    dim Enc_userpass : Enc_userpass= MD5(userpass)

		sqlStr = "update [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + " set password='" + userpass + "'," + vbCrlf
		sqlStr = sqlStr + " company_name='" + username + "'," + vbCrlf
		sqlStr = sqlStr + " email='" + usermail + "'," + vbCrlf
		sqlStr = sqlStr + " userdiv=" + CStr(userdiv) + "," + vbCrlf
		sqlStr = sqlStr + " isusing='" + isusing + "'," + vbCrlf
		sqlStr = sqlStr + " discountrate=" + discountrate + "," + vbCrlf
		sqlStr = sqlStr + " commission=" + commission + "," + vbCrlf
		sqlStr = sqlStr + " bigo='" + bigo + "'" + vbCrlf
		sqlStr = sqlStr + " where id='" + CStr(userid) + "'"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

''		'' tbl_user_n 에 값이 없는경우만 && tbl_user_c 에 값이 있는경우만 // 20120813 서동석
''        sqlStr = "IF Not Exists(select userid from [db_user].[dbo].tbl_user_n where userid='"+userid+"')" + VbCrlf
''        sqlStr = sqlStr + " BEGIN" + VbCrlf
''		sqlStr = sqlStr + "     update L" + VbCrlf
''		sqlStr = sqlStr + "     set userpass='" + userpass + "'" + VbCrlf
''		sqlStr = sqlStr + "     , Enc_userpass='" + Enc_userpass + "'" + VbCrlf
''		sqlStr = sqlStr + "     from [db_user].[dbo].tbl_logindata L" + VbCrlf
''		sqlStr = sqlStr + "         Join [db_user].[dbo].tbl_user_c C" + VbCrlf
''		sqlStr = sqlStr + "         on L.userid=C.userid" + VbCrlf
''		sqlStr = sqlStr + "     where L.userid='" + userid + "'"+ VbCrlf
''        sqlStr = sqlStr + " END"
''
''		dbget.Execute sqlStr


    end sub

	public Sub addNewEmploy(byval userid,userpass,username,usermail,userdiv,bigo)
	    response.write "사용중지메뉴-관리자문의요망(addNewEmploy)"
	    response.end
		dim sqlStr
		sqlStr = "insert into [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + "(id,password,company_name,email,bigo,userdiv)" + vbCrlf
		sqlStr = sqlStr + " values('" + userid + "'," + vbCrlf
		sqlStr = sqlStr + " '" + userpass + "'," + vbCrlf
		sqlStr = sqlStr + " '" + username + "'," + vbCrlf
		sqlStr = sqlStr + " '" + usermail + "'," + vbCrlf
		sqlStr = sqlStr + " '" + bigo + "'," + vbCrlf
		sqlStr = sqlStr + " " + userdiv + "" + vbCrlf
		sqlStr = sqlStr + ")"

		''response.write sqlStr
		dbget.execute sqlStr
	end sub

	public Sub editEmploy(byval userid,userpass,username,usermail,userdiv,bigo,isusing)
	    response.write "사용중지메뉴-관리자문의요망(editEmploy)"
	    response.end
		dim sqlStr
		sqlStr = "update [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + " set password='" + userpass + "'," + vbCrlf
		sqlStr = sqlStr + " company_name='" + username + "'," + vbCrlf
		sqlStr = sqlStr + " email='" + usermail + "'," + vbCrlf
		sqlStr = sqlStr + " userdiv=" + CStr(userdiv) + "," + vbCrlf
		sqlStr = sqlStr + " isusing='" + isusing + "'," + vbCrlf
		sqlStr = sqlStr + " bigo='" + bigo + "'" + vbCrlf
		sqlStr = sqlStr + " where id='" + CStr(userid) + "'"

		''response.write sqlStr
		dbget.execute sqlStr
	end sub

	public Sub GetOnePartner(byval userid)
		dim sqlStr
		dim oneitem

		sqlStr = "select top 1 * from [db_partner].[dbo].tbl_partner"
		sqlStr = sqlStr + " where id='" + CStr(userid) + "'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount
		if Not rsget.Eof then
			set oneitem = new CPartnerUserItem
			oneitem.Fid               = rsget("id")
			oneitem.Fpassword         = rsget("password")
			oneitem.Fdiscountrate     = rsget("discountrate")
			oneitem.Fcompany_name     = rsget("company_name")
			oneitem.Faddress          = rsget("address")
			oneitem.Ftel              = rsget("tel")
			oneitem.Fmanager_hp    = rsget("manager_hp")
			oneitem.Ffax              = rsget("fax")
			oneitem.Fbigo             = rsget("bigo")
			oneitem.Furl              = rsget("url")
			oneitem.Fmanager_name     = rsget("manager_name")
			oneitem.Fmanager_address  = rsget("manager_address")
			oneitem.Fcommission       = rsget("commission")
			oneitem.Femail            = rsget("email")
			oneitem.Fbirthday     = rsget("birthday")
			oneitem.Fmsn     = rsget("msn")
			oneitem.Fzipcode     = rsget("zipcode")
			oneitem.Fbuseo     = rsget("buseo")
			oneitem.Fpart    = rsget("part")
			oneitem.Fcposition     = rsget("cposition")
			oneitem.Fintro     = rsget("intro")
			oneitem.Fuserimg          = rsget("userimg")
			oneitem.Fuserdiv          = rsget("userdiv")
			oneitem.Fisusing          = rsget("isusing")
			set FPartnerList(0) = oneitem
		end if
		rsget.close
	end Sub

	public Function UpdateDesignerSet(byval idesignerid, isusing, isextusing, isb2b)
		dim sqlStr
		sqlStr = "update [db_user].[dbo].tbl_user_c" + vbCrlf
		sqlStr = sqlStr + " set isusing='" + isusing + "'," + vbCrlf
		sqlStr = sqlStr + " isextusing='" + isextusing + "'," + vbCrlf
		sqlStr = sqlStr + " isb2b='" + isb2b + "'" + vbCrlf
		sqlStr = sqlStr + " where userid='" + idesignerid + "'"

		'response.write sqlStr
		dbget.execute sqlStr
	end function

	public Sub editPartnerDesigner2(byval userid)
	    dim sqlStr
		sqlStr = "update [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + " set " + vbCrlf
		if FPass_yn = "" then
		else
		sqlStr = sqlStr + " password='" + Fpassword + "'," + vbCrlf
		end if
		sqlStr = sqlStr + " company_name='" + Fcompany_name + "'," + vbCrlf
		sqlStr = sqlStr + " zipcode='" + Fzipcode + "'," + vbCrlf
		sqlStr = sqlStr + " address='" + Faddress + "'," + vbCrlf
		sqlStr = sqlStr + " manager_address='" + Fmanager_address + "'," + vbCrlf
		sqlStr = sqlStr + " ceoname='" + Fceoname + "'," + vbCrlf
		sqlStr = sqlStr + " company_upjong ='" + Fcompany_upjong + "'," + vbCrlf
		sqlStr = sqlStr + " company_uptae='" + Fcompany_uptae + "'," + vbCrlf
		sqlStr = sqlStr + " company_no='" + Fcompany_no + "'," + vbCrlf
		sqlStr = sqlStr + " url='" + Furl + "'," + vbCrlf
		sqlStr = sqlStr + " tel='" + Ftel + "'," + vbCrlf
		sqlStr = sqlStr + " fax='" + Ffax + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_bank='" + Fjungsan_bank + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_acctno='" + Fjungsan_acctno + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_acctname='" + Fjungsan_acctname + "'," + vbCrlf
		sqlStr = sqlStr + " manager_name='" + Fmanager_name + "'," + vbCrlf
		sqlStr = sqlStr + " manager_phone='" + Fmanager_phone + "'," + vbCrlf
		sqlStr = sqlStr + " email='" + Femail + "'," + vbCrlf
		sqlStr = sqlStr + " manager_hp ='" + Fmanager_hp + "'," + vbCrlf
		sqlStr = sqlStr + " deliver_name ='" + Fdeliver_name + "'," + vbCrlf
		sqlStr = sqlStr + " deliver_phone='" + Fdeliver_phone + "'," + vbCrlf
		sqlStr = sqlStr + " deliver_email='" + Fdeliver_email + "'," + vbCrlf
		sqlStr = sqlStr + " deliver_hp   ='" + Fdeliver_hp + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_name ='" + Fjungsan_name + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_phone='" + Fjungsan_phone + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_email='" + Fjungsan_email + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_hp	 ='" + Fjungsan_hp + "'" + vbCrlf

		sqlStr = sqlStr + " where id='" + CStr(userid) + "'"

		''response.write sqlStr
		dbget.execute sqlStr
    end sub

'==========================================태훈 추가
        public Sub editPartnerDesigner(byval userid)
	        dim sqlStr
		sqlStr = "update [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + " set " + vbCrlf
		if FPass_yn = "" then
		else
		sqlStr = sqlStr + " password='" + Fpassword + "'," + vbCrlf
		end if
		sqlStr = sqlStr + " company_name='" + Fcompany_name + "'," + vbCrlf
		sqlStr = sqlStr + " address='" + Faddress + "'," + vbCrlf
		sqlStr = sqlStr + " tel='" + Ftel + "'," + vbCrlf
		sqlStr = sqlStr + " fax='" + Ffax + "'," + vbCrlf
		sqlStr = sqlStr + " url='" + Furl + "'," + vbCrlf
		sqlStr = sqlStr + " manager_name='" + Fmanager_name + "'," + vbCrlf
		sqlStr = sqlStr + " manager_address='" + Fmanager_address + "'," + vbCrlf
		sqlStr = sqlStr + " email='" + Femail + "'" + vbCrlf

		sqlStr = sqlStr + " where id='" + CStr(userid) + "'"

		''response.write sqlStr
		dbget.execute sqlStr
        end sub
'====================================================

	public Sub GetDesignerList()
		dim i,sqlStr,wheredetail
		dim oneitem

		if FRectDesignerID<>"" then
			wheredetail =" and c.userid='" + FRectDesignerID + "'"
		end if

		if FRectDesignerName<>"" then
			wheredetail =" and c.socname like '%" + FRectDesignerName + "'%"
		end if

		if FRectDesignerDiv<>"" then
			wheredetail =" and c.userdiv = '" + FRectDesignerDiv + "'"
		end if

		if FRectIsUsing<>"" then
			wheredetail =" and IsNull(c.isusing,'N') = '" + FRectIsUsing + "'"
		end if

		if FRectIsB2BUsing<>"" then
			wheredetail =" and IsNull(c.isb2b,'N') = '" + FRectIsB2BUsing + "'"
		end if

		if FRectIsExtUsing<>"" then
			wheredetail =" and IsNull(c.isextusing,'N') = '" + FRectIsExtUsing + "'"
		end if

		''#################################################
		''총 갯수.
		''#################################################
		sqlStr = "select Count(userid) as cnt from [db_user].[dbo].tbl_user_c c, [db_user].[dbo].tbl_user_div d" + vbCrlf
		sqlStr = sqlStr + " where c.userdiv=d.divcode" + vbCrlf
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr + wheredetail, dbget, adOpenForwardOnly, adLockReadOnly


		FTotalCount = rsget("cnt")
		rsget.Close

		''#################################################
		''현재 페이지 리스트.
		''#################################################
		sqlStr = "select top " + CStr(FPageSize) + " c.*, d.divename from [db_user].[dbo].tbl_user_c c, [db_user].[dbo].tbl_user_div d" + vbCrlf
		sqlStr = sqlStr + " where c.userdiv=d.divcode" + vbCrlf
		sqlStr = sqlStr + " and c.userid not in (" + vbCrlf
		sqlStr = sqlStr + " select top " + CStr((FCurrPage-1)*FPageSize)  + " c.userid from [db_user].[dbo].tbl_user_c c, [db_user].[dbo].tbl_user_div d" + vbCrlf
		sqlStr = sqlStr + " where c.userdiv=d.divcode" + vbCrlf
		sqlStr = sqlStr + wheredetail + vbCrlf
		sqlStr = sqlStr + " )" + vbCrlf
		sqlStr = sqlStr + wheredetail + vbCrlf

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FPartnerList(FResultCount)
		i=0
		do until rsget.eof
			set oneitem = new CDesignerUserItem
			oneitem.FUserID     = rsget("userid")
			oneitem.FSocNo      = rsget("socno")
			oneitem.FSocName    = rsget("socname")
			oneitem.FSocMail    = rsget("socmail")
			oneitem.FSocUrl     = rsget("socurl")
			oneitem.FSocPhone   = rsget("socphone")
			oneitem.FSocFax     = rsget("socfax")
			oneitem.FIsUsing    = rsget("isusing")
			oneitem.FIsB2B      = rsget("isb2b")
			oneitem.FUserDiv    = rsget("userdiv")
			oneitem.FIsExtUsing = rsget("isextusing")
			oneitem.FUserDivName= rsget("divename")

			set FPartnerList(i) = oneitem
			i=i+1
			rsget.movenext
		loop
		rsget.close
	end Sub

	public Sub GetPartnerList(byval ix)
		dim sqlStr, wheredetail
		dim oneitem,i

		if ix=1 then
			'' 파트너만.
			wheredetail = " and userdiv=999"
		elseif ix=2 then
			'' 직원만.
			wheredetail = " and userdiv<999"
		elseif ix=3 then
			'' 디자이너
			wheredetail = " and userdiv=9999"
			if FRectInitial="etc" then
				wheredetail = wheredetail + " and ((Left(id,1)<'a') or (Left(id,1)>'Z'))"
			elseif FRectInitial<>"" then
				wheredetail = wheredetail + " and (id like '" + FRectInitial + "%')"
			end if
		else
			wheredetail = ""
		end if

		''#################################################
		''총 갯수.
		''#################################################
		sqlStr = "select Count(id) as cnt from [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + " where id<>''" + vbCrlf
		if FRectUserDiv<>"" then
			sqlStr = sqlStr + " and userdiv='" + FRectUserDiv + "'"
		end if
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr + wheredetail, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget("cnt")
		rsget.Close

		''#################################################
		''현재 페이지 리스트.
		''#################################################
		sqlStr = "select top " + CStr(FCurrPage*FPageSize) + " * from [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + " where id<>''" + vbCrlf
		if FRectUserDiv<>"" then
			sqlStr = sqlStr + " and userdiv='" + FRectUserDiv + "'"
		end if
		sqlStr = sqlStr + wheredetail
'response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FPartnerList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set oneitem = new CPartnerUserItem
				oneitem.Fid               = rsget("id")
				oneitem.Fpassword         = rsget("password")
				oneitem.Fdiscountrate     = rsget("discountrate")
				oneitem.Fcompany_name     = db2html(rsget("company_name"))
				oneitem.Faddress          = db2html(rsget("address"))
				oneitem.Ftel              = rsget("tel")
				oneitem.Ffax              = rsget("fax")
				oneitem.Fbigo             = rsget("bigo")
				oneitem.Furl              = rsget("url")
				oneitem.Fmanager_name     = db2html(rsget("manager_name"))
				oneitem.Fmanager_address  = db2html(rsget("manager_address"))
				oneitem.Fcommission       = rsget("commission")
				oneitem.Femail            = db2html(rsget("email"))
				oneitem.Fuserdiv          = rsget("userdiv")
				oneitem.Fisusing          = rsget("isusing")
				set FPartnerList(i) = oneitem
				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
	end Sub


	public Function fnGetDispCateList
		Dim strSql
		strSql = "SELECT pd.catecode FROM [db_partner].[dbo].[tbl_partner_dispcate] AS pd " & _
				 "WHERE pd.makerid = '" & FRectDesignerID & "'"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		'response.write strSql
		IF not rsget.EOF THEN
			fnGetDispCateList = rsget.getRows()
		END IF
		rsget.close
	End Function

	public Function fnUserC_GetDispCateList
		Dim strSql
		strSql = ""
		strSql = strSql & " SELECT pd.catecode, C.userid, C.standardCateCode "
		strSql = strSql & " FROM db_user.dbo.tbl_user_c AS C "
		strSql = strSql & " JOIN [db_partner].[dbo].[tbl_partner_dispcate] as pd on c.userid = pd.makerid "
		strSql = strSql & " WHERE C.userid = '" & FRectDesignerID & "' AND isnull(pd.isdefault,'') <> 'Y'  "
		strSql = strSql & " ORDER BY pd.catecode ASC "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		IF not rsget.EOF THEN
			fnUserC_GetDispCateList = rsget.getRows()
		END IF
		rsget.close
	End Function

'브랜드별 기본마진 정보 가져오기 2014.01.15 정윤정 추가
	public Function fnGetDefaultMargine
	Dim strSql
		strSql =  " SELECT  defaultmargine "
		strSql = strSql & " FROM db_user.dbo.tbl_user_c  "
		strSql = strSql & " WHERE  userid = '" & FRectDesignerID & "'"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		IF not rsget.EOF THEN
			fnGetDefaultMargine = rsget(0)
		END IF
		rsget.close
	End Function

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end class
%>
