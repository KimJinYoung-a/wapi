<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 브랜드 계약 관리
' Hieditor : 2009.04.07 서동석 생성
'			 	 2010.05.26 한용민 수정
' 			2017.06.23 정윤정 전자계약 추가
'###########################################################
%>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" --> 
<!-- #include virtual="/lib/ecContractApi_function.asp"-->
<% 
'--===================================================
'-- gs 와 상태값 매칭 
    function GetContractEcState(ContractStateName)
        dim buf
        Select Case ContractStateName
            Case "미전송"
                : buf = "0"
            Case "검토대기"
                : buf =  "1"
             Case "검토반려"
                : buf =  "2"    
            Case "결재완료"
                : buf = "3"
            Case "전자서명진행"  
                : buf =   "6"
            Case "계약완료"
                : buf = "7"
            Case "계약파기요청"
                : buf = "8"
            Case "계약파기"
                : buf = "9"   
            Case "계약종료" 
                : buf =    "9"
            Case "삭제"
                : buf = -1
            Case "-1"
                : buf = -1
            Case else
                : buf = "-2"
        end Select

        GetContractEcState = buf
    end function
 '--===================================================  
dim sqlStr
dim oneContract,acctoken,reftoken,ecCtrState
dim  arrList, intLoop


 		sqlStr = " select  top 5 m.ctrKey, ecctrseq, g.company_no, ecBUser , m.ctrstate "
 		sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master as m  "
 		sqlStr = sqlStr & "	inner join db_partner.dbo.tbl_partner_group as g on m.groupid = g.groupid "
 		sqlStr = sqlStr & "	where CtrState > 0 and CtrState not in (7,9) "
 		sqlStr = sqlStr & "	 and ecCtrseq > 0  	"
 		sqlStr = sqlStr & " order by m.ecupdate asc , m.ctrkey asc "

 	    rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
 	     if not rsget.eof Then
 	     	 arrList = rsget.getrows()
 	    end if
 	    rsget.close
 	    
 	    if isArray(arrList) Then
 	    		
		'token 가져오기(db에서) 
			sqlStr = "select top 1 access_token, refresh_token from db_partner.dbo.tbl_partner_ctrLg_token order by tidx desc "
			rsget.CursorLocation = adUseClient
            rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			if not rsget.eof then
				acctoken = rsget("access_token")
				reftoken = rsget("refresh_token")
			end if
			rsget.close
		 
  		
  		'token이 없을때 token 발행
 				if not isNull(acctoken) then  
 	 	
 	 				for intLoop = 0 To uBound(arrList,2)
 	 					ecCtrState =  fnViewEcCont(arrList(1,intLoop),replace(arrList(2,intLoop),"-",""),arrList(3,intLoop),acctoken)
 	 				    
 	 				    response.write "ctrKey:"&arrList(0,intLoop)&"-"&"ecCtrState:"&ecCtrState&"-"&"chkerror:"&Fchkerror&"<br>"
 	 				    
 	 					if Fchkerror ="invalid_token" then
				 				call sbGetRefToken(reftoken)
				 				acctoken = Faccess_token
				 				ecCtrState =  fnViewEcCont(arrList(1,intLoop),replace(arrList(2,intLoop),"-",""),arrList(3,intLoop),acctoken)
				 		end if
				 		
				 		if ecCtrState <> "" then				 	 				 	 
    				 		sqlStr = "update db_partner.dbo.tbl_partner_ctr_master set ctrstate = "&GetContractEcState(ecCtrState)&", lastupdate =getdate(), ecupdate = getdate() "
    			 			sqlstr = sqlstr & " where ctrKey="&arrList(0,intLoop)  
    			 			dbget.Execute  sqlstr, 1	
    			 		else
    			 		    sqlStr = "update db_partner.dbo.tbl_partner_ctr_master set ecupdate = getdate() "
    			 			sqlstr = sqlstr & " where ctrKey="&arrList(0,intLoop)  
    			 			dbget.Execute  sqlstr, 1
			 			end if
 	 				next
 	 			    
 	 			end if	
 	 		response.flush
 		end if
 	  
	 
%>		
  <script type="text/javascript">
	alert("완료되었습니다.");
	 
</script>	
<!-- #include virtual="/lib/db/dbclose.asp" -->				