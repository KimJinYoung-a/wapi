<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : DocuSign 계약 완료 여부 체크
' Hieditor : 2022.02.08 원승현 생성
'###########################################################
%>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" --> 
<!-- #include virtual="/lib/ecContractApi_function.asp"-->
<% 

dim sqlStr
dim oneContract,acctoken,reftoken,ecCtrState
dim  arrList, intLoop
dim  docuStatusAdminCodeConversion, docuErrorStatusValue
dim docuSignEnvelopeId, docuSignStatus, docuSignStatusDateTime, docuSignUri, objXML, iRbody, jsResult
docuErrorStatusValue = ""


 		sqlStr = " select  top 5 m.ctrKey, m.docuSignId, m.ctrstate "
 		sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master as m  "
 		sqlStr = sqlStr & "	inner join db_partner.dbo.tbl_partner_group as g on m.groupid = g.groupid "
 		sqlStr = sqlStr & "	where CtrState > 0 and CtrState not in (7,9) "
        sqlStr = sqlStr & " AND ISNULL(m.docuSignId,'') <> '' "
        sqlStr = sqlStr & " AND signType='D' "
 		sqlStr = sqlStr & " order by m.ctrkey asc "

 	    rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
 	     if not rsget.eof Then
 	     	 arrList = rsget.getrows()
 	    end if
 	    rsget.close
 	    
 	    if isArray(arrList) Then
            for intLoop = 0 To uBound(arrList,2)
                Session.CodePage = 65001
                'Set objXML = CreateObject("Msxml2.ServerXMLHTTP")
                'objXML.SetTimeouts 40000, 40000, 40000, 40000
                Set objXML= CreateObject("MSXML2.XMLHTTP.3.0")
                objXML.Open "GET", FecDocuURL&"/api/contract/v1/docusign/envelope/"&arrList(1,intLoop), False
                objXML.setRequestHeader "Content-Type", "application/json"
                if (application("Svr_Info")	<> "Dev") then
                    objXML.SetRequestHeader "api-key-v1", ""+CStr(adminApiKey)+""
                End If                
                objXML.Send
                Session.CodePage = 949
                iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")

                If objXML.Status = "200" Then
                    Set jsResult = JSON.parse(iRbody)
                    docuSignEnvelopeId = jsResult.envelopeId
                    docuSignStatus = jsResult.status
                    docuSignStatusDateTime = jsResult.statusDateTime
                    docuSignUri = jsResult.uri
                    Set jsResult = Nothing

                    Select Case Trim(docuSignStatus)
                        case "created"
                            docuStatusAdminCodeConversion = 1		
                        case "sent"
                            docuStatusAdminCodeConversion = 1
                        case "delivered"
                            docuStatusAdminCodeConversion = 1
                        case "signed"
                            docuStatusAdminCodeConversion = 6
                        case "declined"
                            docuStatusAdminCodeConversion = 2
                        case "completed"
                            docuStatusAdminCodeConversion = 7
                        'case "faxpending" '' 텐바이텐에선 사용안함
                        'case "autoresponded" '' 텐바이텐에선 사용안함
                        Case Else
                            docuStatusAdminCodeConversion = arrList(2,intLoop)
                    End Select					

                    sqlStr = "update db_partner.dbo.tbl_partner_ctr_master set ctrstate = "&docuStatusAdminCodeConversion&", lastupdate =getdate()"
                    sqlstr = sqlstr & " where ctrstate<>-1 and ISNULL(docuSignId,'') <> '' and signType='D' AND DocuSignId='"&arrList(1,intLoop)&"' AND ctrkey='"&arrList(0,intLoop)&"' "
                    dbget.Execute  sqlstr, 1						
                Else
                    docuErrorStatusValue = docuErrorStatusValue &","& objXML.Status
                    'response.write "<script>alert('DocuSign 통신중 오류가 발생하였습니다.\nErrorCode("&objXML.Status&")');</script>"
                    'response.write "<script>location.replace('" & refer & "');</script>"
                    'dbget.close() : response.End
                End If
                Set objXML = Nothing
            next
 	 		response.flush
 		end if
 	  
	 
%>		
  <script type="text/javascript">
	alert("완료되었습니다.");
	 
</script>	
<!-- #include virtual="/lib/db/dbclose.asp" -->				