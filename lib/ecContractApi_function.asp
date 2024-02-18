
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%

'###########################################################
' Description : 전자계약 연동 api
' Hieditor : 2017.07.05 정윤정
'###########################################################
 
public  FecURL, FecID,FecAUser
public Fec_defctrtype , Fec_defctrtype_M,Fec_addctrtype,Fec_addctrtype_M

 if (application("Svr_Info")	= "Dev") then
	FecURL= "http://webtax-test.uplus.co.kr" 
	FecID = "CTENTEN"
	FecAUser = "CTENTEN"
	Fec_defctrtype ="429"
	Fec_defctrtype_M ="432"
	Fec_addctrtype = "435"
	Fec_addctrtype_M = "436"
else
	FecURL = "https://edocu.uplus.co.kr" 
	FecID = "CUBE1010"
	FecAUser = "CTENTEN"
	Fec_defctrtype ="2270"
	Fec_defctrtype_M ="2269"
	Fec_addctrtype = "2267"
	Fec_addctrtype_M = "2268"
end if
 
'########################### Docu용 api ####################################################   
 public  FecDocuURL, adminApiKey
 if (application("Svr_Info")	= "Dev") then
	FecDocuURL= "https://testadminapi.10x10.co.kr" 
else
	FecDocuURL = "https://adminapi.10x10.co.kr:31443"
	adminApiKey = "71ec2fbe40cfbcc43882e1141f662fk9e291ede5bdf9ed6a4defc28"
end if
 
 public Faccess_token 
 public Ftoken_type
 public Frefresh_token
 public Fchkerror
 public FErrMsg
 
'########################### 토큰 발행 #################################################### 
 	public Sub sbGetNewToken(ecId,ecPwd)
 	dim APIpath,strParam
 	dim objXML, iRbody,jsResult
 	dim sqlStr
 	 
			APIpath = FecURL&"/oauth/token"
			strParam = "?grant_type=password&client_id=edocuopenapi&username="&ecId&"&password="&ecPwd
		 	'On Error Resume Next
				Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
				    objXML.Open "GET", APIpath&strParam , False
					objXML.setRequestHeader "Content-Type", "text/html"
					objXML.Send()
					iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
			'		 response.write iRbody
					If objXML.Status = "200" Then
						Set jsResult = JSON.parse(iRbody)
							Faccess_token	= jsResult.access_token
							Ftoken_type		= jsResult.token_type
							Frefresh_token	= jsResult.refresh_token						
						Set jsResult = Nothing
					End If
				Set objXML = Nothing
		 	'On Error Goto 0 
		 	
		 	if Faccess_token <> "" Then
				sqlStr = "insert into db_partner.dbo. tbl_partner_ctrLg_token(access_token, refresh_token)"	
				sqlStr = sqlStr & "values('"&Faccess_token&"','"&Frefresh_token&"')"
				dbget.Execute sqlStr
			end if	
	End Sub		

 
'########################### 토큰 갱신 #################################################### 
 	public Sub sbGetRefToken(reftoken)
 	dim APIpath,strParam
 	dim objXML, iRbody,jsResult
 	dim sqlStr
			APIpath = FecURL&"/oauth/token"
			strParam = "?grant_type=refresh_token&client_id=edocuopenapi&refresh_token="&reftoken
		 	On Error Resume Next
				Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
				    objXML.Open "GET", APIpath&strParam , False
					objXML.setRequestHeader "Content-Type", "text/html"
					objXML.Send()
					iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
					 
					If objXML.Status = "200" Then
						Set jsResult = JSON.parse(iRbody)
							Faccess_token	= jsResult.access_token
							Ftoken_type		= jsResult.token_type
							Frefresh_token	= jsResult.refresh_token						
						Set jsResult = Nothing
					End If
				Set objXML = Nothing
		 	On Error Goto 0 
		 	
		 	if Faccess_token <> "" Then
				sqlStr = "insert into db_partner.dbo. tbl_partner_ctrLg_token(access_token, refresh_token)"	
				sqlStr = sqlStr & "values('"&Faccess_token&"','"&Frefresh_token&"')"
				dbget.Execute sqlStr
			end if	
	End Sub		
	
'########################### 회원확인 #################################################### 	
	public Function  fnCheckUser(bcompno,access_token)
	 	dim APIpath,strParam
 		dim objXML, iRbody,jsResult
 		dim userStatus 
				APIpath = FecURL&"/api/checkuser"
				strParam = "?corp_id="&bcompno
					 
				On Error Resume Next
					Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
					    objXML.Open "GET", APIpath&strParam , False
						objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
						objXML.SetRequestHeader "Authorization", "Bearer " & access_token
						objXML.Send()
						iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")				
						iRbody= replace(iRbody,"tmpCallBack({","{")
						iRbody = replace(iRbody,"})","}") 
						If objXML.Status = "200" Then 
							Set jsResult = JSON.parse(iRbody)
								userStatus	= jsResult.status_cd
								Fchkerror      = jsResult.error
							Set jsResult = Nothing			
						end if				
					Set objXML = Nothing
				On Error Goto 0
				
				fnCheckUser = userStatus
	End Function 
	
	
	'########################### 계약오픈  #################################################### 	
	public Function fnCheckCont(ecCtrSeq,B_COMPANY_NO,ecBUser, access_token)
		dim APIpath,strParam
 		dim objXML, iRbody,jsResult
 		
		APIpath =FecURL&"/api/checkCont" 
		strParam = "?seq="&ecCtrSeq&"&venderno="&B_COMPANY_NO&"&users="&server.URLEncode(ecBUser)
		
			On Error Resume Next
		
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			  objXML.Open "GET", APIpath&strParam , False
				objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
				objXML.SetRequestHeader "Authorization", "Bearer " & access_token
				objXML.Send()				
				iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
			
				iRbody= replace(iRbody,"tmpCallBack({","{")
			 	iRbody = replace(iRbody,"})","}") 
		
				If objXML.Status = "200" Then
					Set jsResult = JSON.parse(iRbody)
						con_status	= jsResult.status
						con_info= jsResult.info 
						Fchkerror = jsResult.error
						if con_status ="fail" Then 
								 if (con_info="001") then
								 		FErrMsg= "계약서번호 값 없음"
								 elseif (con_info="002")then
								  	FErrMsg= "사업자번호(venderno) 값 없음,venderno 에 사용자 존재하지 않음"
								 elseif (con_info="003")then
								  	FErrMsg= "해당 정보에 대한 문서가 존재하지않음" 
								 elseif (con_info="004")then
									  FErrMsg= "해당문서가 이미 검토요청 상태" 														  
								 end if 
								  
						end if
						
				Set jsResult = Nothing
				End If
			Set objXML = Nothing
			
				On Error Goto 0 
				fnCheckCont = con_status
	end function		
	
	'########################### 정보확인  #################################################### 	
 	public Function fnViewEcCont(ecCtrseq, bcompno, ecBuser,access_token)			 
 	dim APIpath,strParam
 		dim objXML, iRbody,jsResult
 		dim con_status,con_info,con_error,ecCtrState
	APIpath =FecURL&"/api/viewCont" 
	strParam = "?seq="&ecCtrseq&"&venderno="&bcompno&"&users=" &server.URLEncode(ecBUser)
	
 	On Error Resume Next 
 Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		  objXML.Open "GET", APIpath&strParam , False
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
			objXML.SetRequestHeader "Authorization", "Bearer " & access_token
			objXML.Send()				
			iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
		
			iRbody= replace(iRbody,"tmpCallBack({","{")
		 	iRbody = replace(iRbody,"})","}") 
  
			If objXML.Status = "200" Then
				Set jsResult = JSON.parse(iRbody)
					con_status	= jsResult.status
					con_info= jsResult.info 
				'	con_error = jsResult.error
					if con_status ="fail" Then 
							 if (con_info="001") then
							 		FErrMsg= "계약서번호 값 없음"
							 elseif (con_info="002")then
							  	FErrMsg= "사업자번호(venderno) 값 없음,venderno 에 사용자 존재하지 않음"
							 elseif (con_info="003")then
							  	FErrMsg= "해당 정보에 대한 문서가 존재하지않음"  
							 elseif (con_info="004")then
							  	FErrMsg= "해당문서가 삭제되었습니다."   	
							  	ecCtrState = "-1"
							 end if 
					else
						ecCtrState = jsResult.nowstat_vl			  
					end if
					
			Set jsResult = Nothing
			End If
		Set objXML = Nothing		 
	On Error Goto 0 
	fnViewEcCont = ecCtrState
	
	End Function
%>

	 